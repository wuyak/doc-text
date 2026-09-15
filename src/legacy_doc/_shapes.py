"""Bounded association of main-document anchors with textbox ranges.

This module deliberately deals with the *association* records only.  Text in a
textbox is stored in the textbox subdocument and is read by the text reader;
the result here is a mapping from an anchor CP in the main document to the
absolute CP range of the associated textbox text.

The records used here are the normative Word 97--2003 records:

* ``FibRgFcLcb97`` pair 40, ``PlcfSpaMom`` (26-byte ``SPA`` records),
* ``FibRgFcLcb97`` pair 56, ``PlcftxbxTxt`` (22-byte ``FTXBXS`` records), and
  pair 75, ``PlcfTxbxBkd`` (6-byte ``Tbkd`` records for linked ranges), and
* pair 50, ``OfficeArtContent`` (OfficeArt ``FSP``, ``ClientAnchor`` and
  ``ClientTextbox`` records).

OfficeArt geometry text (``gtextUNICODE_complex``) has no unique DOCX node or
property correspondence.  It is therefore exposed as an explicitly
unsupported scope by :func:`unsupported_geometry_text_shapes`; when a main
anchor selects such a shape, :func:`body_textboxes` raises instead of
silently dropping its text.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass
from typing import TYPE_CHECKING, Iterable

if TYPE_CHECKING:
    from legacy_doc._binary import BinaryDocument, FibInfo

from legacy_doc.exceptions import LegacyDocError


# FibRgFcLcb97 is an array of (fc, lcb) pairs.  These indices are fixed by the
# Word 97--2003 FIB layout and are also the indices used by the MS-DOC spec.
_PLC_SPA_MOM_PAIR = 40
_PLC_TXBX_TXT_PAIR = 56
_DGG_INFO_PAIR = 50
_PLC_TXBX_BKD_PAIR = 75

_SPA_SIZE = 26
_FTXBXS_SIZE = 22

# OfficeArt record types needed for the association path.
_DGG_CONTAINER = 0xF000
_DG_CONTAINER = 0xF002
_SPGR_CONTAINER = 0xF003
_SP_CONTAINER = 0xF004
_FSP = 0xF00A
_FOPT = 0xF00B
_CLIENT_TEXTBOX = 0xF00D
_CLIENT_ANCHOR = 0xF010
_SECONDARY_FOPT = 0xF121
_TERTIARY_FOPT = 0xF122

# The gtextUNICODE property identifier in the OfficeArt property table.
_GTEXT_UNICODE = 0x00C0

# A selected geometry-text shape fails explicitly rather than losing text.
UNSUPPORTED_GEOMETRY_TEXT_SCOPE = (
    "OfficeArt gtextUNICODE_complex geometry text is unsupported: "
    "it has no unique DOCX node/property correspondence"
)

_MAX_PLC_ENTRIES = 1_000_000
_MAX_OFFICEART_RECORDS = 1_000_000
_MAX_OFFICEART_DEPTH = 64


@dataclass(frozen=True, slots=True)
class TextboxRange:
    """An absolute CP range belonging to a main-document textbox.

    ``start`` and ``end`` are half-open, absolute CPs in the document-wide
    piece-table coordinate space.  The optional fields retain the association
    evidence for callers that need to inspect it; text extraction only needs
    ``start`` and ``end``.
    """

    start: int
    end: int
    shape_id: int | None = None
    textbox_index: int | None = None
    chain_index: int = 0


@dataclass(frozen=True, slots=True)
class GeometryTextScope:
    """An explicitly unsupported OfficeArt geometry-text scope."""

    shape_id: int
    kind: str = "gtextUNICODE_complex"
    reason: str = UNSUPPORTED_GEOMETRY_TEXT_SCOPE


@dataclass(frozen=True, slots=True)
class _Anchor:
    cp: int
    shape_id: int
    index: int


@dataclass(frozen=True, slots=True)
class _TextboxRecord:
    index: int
    start: int
    end: int
    shape_id: int
    chain_count: int


@dataclass(frozen=True, slots=True)
class _TextboxBreak:
    """One text-range component from ``PlcfTxbxBkd``."""

    start: int
    end: int


@dataclass(frozen=True, slots=True)
class _OfficeArtShape:
    shape_id: int
    anchor_index: int | None
    text_id: int | None
    text_chain_index: int | None
    geometry_text: bool


@dataclass(frozen=True, slots=True)
class _OfficeArtIndex:
    shapes: dict[int, _OfficeArtShape]
    geometry_text_shapes: tuple[int, ...]


@dataclass(frozen=True, slots=True)
class _Record:
    start: int
    payload_start: int
    end: int
    version: int
    instance: int
    record_type: int

    @property
    def is_container(self) -> bool:
        return self.version == 0xF


class _OfficeArtWalker:
    """State shared by one bounded OfficeArt traversal."""

    def __init__(self, *, max_records: int = _MAX_OFFICEART_RECORDS) -> None:
        self.max_records = max_records
        self.records_seen = 0

    def record(self) -> None:
        self.records_seen += 1
        if self.records_seen > self.max_records:
            raise LegacyDocError("OfficeArt record count exceeds parser limit")


def body_textboxes(doc: BinaryDocument) -> dict[int, TextboxRange]:
    """Return main-document textbox ranges keyed by their anchor CP.

    ``doc`` is the bounded ``BinaryDocument`` supplied by the Word reader.  It
    must expose ``table``, ``fib.pair(index)``, ``fib.ccp_text``,
    ``fib.ccp_txbx`` and (when the piece table is not empty) ``cp_limit``.
    The function does not read textbox characters; the returned absolute CPs
    are consumed by the text reader.

    Shapes which are not textboxes (for example pictures) are ignored.  A
    textbox that is referenced by a main-document anchor must have a complete
    SPA/FTXBXS/OfficeArt association; a broken required association raises
    :class:`LegacyDocError` instead of returning a partial mapping.
    """

    fib = doc.fib
    main_cp_limit = fib.ccp_text
    document_cp_limit = doc.cp_limit

    spa_data = _read_fib_region(
        doc,
        _PLC_SPA_MOM_PAIR,
        "PlcfSpaMom",
        allow_empty=True,
    )
    anchors = _parse_spa(spa_data, main_cp_limit)

    # No main anchors means there cannot be a body textbox insertion.  We do
    # not require OfficeArtContent for an otherwise valid document with no
    # selected anchors; this is important for body-only extraction.
    if not anchors:
        return {}

    textbox_cp_limit = fib.ccp_txbx
    txbx_data = _read_fib_region(
        doc,
        _PLC_TXBX_TXT_PAIR,
        "PlcftxbxTxt",
        allow_empty=True,
    )
    textbox_records = _parse_txbx(txbx_data, textbox_cp_limit)
    textbox_records_by_index = {record.index: record for record in textbox_records}
    textbox_record_count = _plc_record_count(txbx_data, _FTXBXS_SIZE, "PlcftxbxTxt")

    actual_by_shape: dict[int, list[_TextboxRecord]] = {}
    for record in textbox_records:
        actual_by_shape.setdefault(record.shape_id, []).append(record)

    anchor_shape_ids = {anchor.shape_id for anchor in anchors}
    candidate_shape_ids = anchor_shape_ids & actual_by_shape.keys()
    office_art_present = _region_present(doc, _DGG_INFO_PAIR, "OfficeArtContent")
    if not candidate_shape_ids and not office_art_present:
        # The main document may contain only pictures or other non-textbox
        # shapes.  There is no textbox association to resolve, and no
        # OfficeArt bytes from which to identify an unsupported geometry-text
        # shape.
        return {}

    art = _parse_office_art(doc)
    result: dict[int, TextboxRange] = {}
    textbox_base = _textbox_base(fib)
    bkd_by_index: dict[int, list[_TextboxBreak]] | None = None
    selected_chain_shapes: dict[int, dict[int, int]] = {}

    def load_bkd() -> dict[int, list[_TextboxBreak]]:
        nonlocal bkd_by_index
        if bkd_by_index is None:
            bkd_data = _read_fib_region(
                doc,
                _PLC_TXBX_BKD_PAIR,
                "PlcfTxbxBkd",
                allow_empty=True,
            )
            bkd_by_index = _parse_txbx_bkd(
                bkd_data,
                textbox_cp_limit,
                textbox_record_count,
            )
        return bkd_by_index

    for anchor in anchors:
        candidates = actual_by_shape.get(anchor.shape_id)
        shape = art.shapes.get(anchor.shape_id)

        if shape is not None and shape.geometry_text:
            # A main-document anchor selects this shape even when it has no
            # FTXBXS textbox story.  Geometry text cannot be projected to the
            # DOCX text-node scope, so make the unsupported boundary explicit
            # instead of returning a successful result that silently drops it.
            raise LegacyDocError(UNSUPPORTED_GEOMETRY_TEXT_SCOPE)

        if shape is None:
            if candidates:
                raise LegacyDocError(
                    "Body textbox shape is missing from OfficeArtContent"
                )
            # A non-textbox anchor does not need a textbox association.
            continue

        if shape.text_id is None:
            if candidates:
                if shape.anchor_index is None:
                    raise LegacyDocError(
                        "Body textbox shape is missing OfficeArtClientAnchor"
                    )
                raise LegacyDocError(
                    "Body textbox shape is missing OfficeArtClientTextbox"
                )
            # A non-textbox shape at this anchor is outside this function's
            # scope.  It must not cause a false failure.
            continue

        if shape.anchor_index is None:
            raise LegacyDocError(
                "Body textbox shape is missing OfficeArtClientAnchor"
            )
        if shape.anchor_index != anchor.index:
            raise LegacyDocError(
                "OfficeArtClientAnchor does not point to the SPA anchor"
            )

        record = textbox_records_by_index.get(shape.text_id - 1)
        if record is None:
            # A ClientTextbox record makes this an explicitly selected
            # textbox shape.  Returning success here would silently lose the
            # shape when its lTxid points at an absent, reusable, or otherwise
            # non-addressable FTXBXS entry.
            raise LegacyDocError(
                "OfficeArtClientTextbox lTxid does not reference a "
                "non-reusable FTXBXS record"
            )

        chain_index = shape.text_chain_index
        if chain_index is None or chain_index >= record.chain_count:
            raise LegacyDocError(
                "OfficeArtClientTextbox lTxid chain index exceeds cTxbx"
            )
        if chain_index == 0:
            if record.shape_id != shape.shape_id:
                raise LegacyDocError(
                    "FTXBXS.lid does not match the first shape in its chain"
                )
            if len(candidates) != 1 or candidates[0].index != record.index:
                raise LegacyDocError(
                    "PlcftxbxTxt has multiple textbox ranges for one shape"
                )
        elif candidates:
            raise LegacyDocError(
                "FTXBXS.lid conflicts with a linked textbox shape"
            )

        range_start = record.start
        range_end = record.end
        if record.chain_count > 1:
            breaks = load_bkd().get(record.index, [])
            if len(breaks) != record.chain_count:
                raise LegacyDocError(
                    "Unsupported PlcfTxbxBkd partial coverage of linked textbox ranges"
                )
            # The FTXBXS range is the complete linked textbox story.  Its
            # Tbkd entries partition that range in chain order; accepting a
            # correct count with a missing edge or an interior gap would make
            # the text reader silently lose part of the selected textbox.
            if breaks[0].start != record.start or breaks[-1].end != record.end:
                raise LegacyDocError(
                    "Unsupported PlcfTxbxBkd partial coverage of its FTXBXS textbox range"
                )
            for previous, current in zip(breaks, breaks[1:]):
                if current.start != previous.end:
                    raise LegacyDocError(
                        "PlcfTxbxBkd linked textbox ranges are not contiguous"
                    )
            selected_break = breaks[chain_index]
            if (
                selected_break.start < record.start
                or selected_break.end > record.end
            ):
                raise LegacyDocError(
                    "PlcfTxbxBkd range exceeds its FTXBXS textbox range"
                )
            range_start, range_end = selected_break.start, selected_break.end

            chain_shapes = selected_chain_shapes.setdefault(record.index, {})
            if chain_index in chain_shapes:
                raise LegacyDocError(
                    "OfficeArt linked textbox chain contains a duplicate component"
                )
            chain_shapes[chain_index] = shape.shape_id

        absolute_start = _checked_add(
            textbox_base,
            range_start,
            "textbox start CP",
        )
        absolute_end = _checked_add(
            textbox_base,
            range_end,
            "textbox end CP",
        )
        if absolute_end <= absolute_start:
            raise LegacyDocError("Textbox CP range is empty or reversed")
        if document_cp_limit is not None and absolute_end > document_cp_limit:
            raise LegacyDocError("Textbox CP range exceeds document CP limit")

        result[anchor.cp] = TextboxRange(
            start=absolute_start,
            end=absolute_end,
            shape_id=shape.shape_id,
            textbox_index=record.index,
            chain_index=chain_index,
        )

    for record_index, chain_shapes in selected_chain_shapes.items():
        record = textbox_records_by_index[record_index]
        # cTxbx is an untrusted signed 32-bit field.  A chain cannot have
        # more components than the bounded active-shape index, and this guard
        # also keeps the completeness check below from allocating a set
        # proportional to a corruptly huge count.
        if record.chain_count > len(art.shapes):
            raise LegacyDocError(
                "OfficeArt linked textbox chain does not select every component"
            )
        expected = set(range(record.chain_count))
        if set(chain_shapes) != expected:
            raise LegacyDocError(
                "OfficeArt linked textbox chain does not select every component"
            )

    # Preserve the PLC's document order even when a caller supplied a mapping
    # implementation with unusual insertion behavior.
    return {cp: result[cp] for cp in sorted(result)}


def unsupported_geometry_text_shapes(doc: BinaryDocument) -> tuple[GeometryTextScope, ...]:
    """Return bounded reports for ``gtextUNICODE_complex`` shape properties.

    Geometry text is deliberately not merged into :func:`body_textboxes`.
    Callers that expose extraction warnings can use this helper to report the
    unsupported scope explicitly.  It returns only active main-document
    drawing shapes; header drawings are excluded.
    """

    if not _region_present(doc, _DGG_INFO_PAIR, "OfficeArtContent"):
        return ()
    art = _parse_office_art(doc)
    return tuple(
        GeometryTextScope(shape_id=shape_id)
        for shape_id in art.geometry_text_shapes
    )


def _read_fib_region(
    doc: BinaryDocument,
    pair_index: int,
    name: str,
    *,
    allow_empty: bool,
) -> bytes:
    fib = doc.fib
    fc, lcb = fib.pair(pair_index)
    if lcb == 0:
        if not allow_empty and fc != 0:
            raise LegacyDocError(f"{name} has a nonzero offset with zero length")
        return b""

    table = doc.table
    table_length = len(table)
    if fc > table_length or lcb > table_length - fc:
        raise LegacyDocError(f"{name} lies outside the Table stream")
    return bytes(table[fc : fc + lcb])


def _region_present(doc: BinaryDocument, pair_index: int, name: str) -> bool:
    """Return whether a FIB stream region has bytes to parse."""

    _, lcb = doc.fib.pair(pair_index)
    return lcb != 0


def _parse_plc(data: bytes, element_size: int, name: str) -> tuple[list[int], list[bytes]]:
    if not data:
        return [], []
    if len(data) < 4:
        raise LegacyDocError(f"{name} is shorter than its PLC header")
    remainder = len(data) - 4
    denominator = 4 + element_size
    if remainder % denominator:
        raise LegacyDocError(f"{name} has a non-integral PLC record count")
    count = remainder // denominator
    if count > _MAX_PLC_ENTRIES:
        raise LegacyDocError(f"{name} record count exceeds parser limit")

    cp_bytes = data[: 4 * (count + 1)]
    cps = [struct.unpack_from("<I", cp_bytes, offset)[0] for offset in range(0, len(cp_bytes), 4)]
    for left, right in zip(cps, cps[1:]):
        if right <= left:
            raise LegacyDocError(f"{name} CPs are not strictly increasing")

    records_start = len(cp_bytes)
    records = [
        data[records_start + index * element_size : records_start + (index + 1) * element_size]
        for index in range(count)
    ]
    return cps, records


def _plc_record_count(data: bytes, element_size: int, name: str) -> int:
    """Return a validated PLC's number of data records."""

    _, records = _parse_plc(data, element_size, name)
    return len(records)


def _parse_spa(data: bytes, cp_limit: int) -> list[_Anchor]:
    cps, records = _parse_plc(data, _SPA_SIZE, "PlcfSpaMom")
    if not records:
        return []
    # FibRgFcLcb97.fcPlcSpaMom: the final CP is undefined and ignored,
    # except for the strict ordering already checked by _parse_plc.
    if cps[-2] > cp_limit:
        raise LegacyDocError("PlcfSpaMom CP exceeds main-document limit")

    anchors: list[_Anchor] = []
    for index, raw in enumerate(records):
        cp = cps[index]
        shape_id = struct.unpack_from("<I", raw, 0)[0]
        anchors.append(_Anchor(cp=cp, shape_id=shape_id, index=index))
    return anchors


def _parse_txbx(data: bytes, cp_limit: int) -> list[_TextboxRecord]:
    cps, records = _parse_plc(data, _FTXBXS_SIZE, "PlcftxbxTxt")
    if not records:
        return []

    actual: list[_TextboxRecord] = []
    last_index = len(records) - 1
    for index, raw in enumerate(records):
        # The union is eight bytes.  fReusable follows it at offset 8, lid at
        # offset 14.  The final record is reusable even when its flag is zero.
        f_reusable = struct.unpack_from("<H", raw, 8)[0]
        if index == last_index or (f_reusable & 0x0001):
            continue

        start = cps[index]
        end = cps[index + 1]
        # PlcftxbxTxt's final range is ignored. Only actual textbox ranges
        # address characters in this document part.
        if end > cp_limit:
            raise LegacyDocError("PlcftxbxTxt CP exceeds textbox-document limit")
        # A non-reusable range in a valid PLC includes at least its terminating
        # paragraph mark, hence its bounding CPs are more than one apart.
        if end - start <= 1:
            raise LegacyDocError("FTXBXS textbox range is too short")

        chain_count = struct.unpack_from("<i", raw, 0)[0]
        if chain_count <= 0:
            raise LegacyDocError("FTXBXNonReusable cTxbx is not positive")
        chain_edit = struct.unpack_from("<i", raw, 4)[0]
        if chain_edit != 0:
            raise LegacyDocError("FTXBXNonReusable cTxbxEdit is not zero")
        shape_id = struct.unpack_from("<I", raw, 14)[0]
        if shape_id == 0:
            raise LegacyDocError("FTXBXNonReusable lid is zero")
        actual.append(
            _TextboxRecord(
                index=index,
                start=start,
                end=end,
                shape_id=shape_id,
                chain_count=chain_count,
            )
        )
    return actual


def _parse_txbx_bkd(
    data: bytes,
    cp_limit: int,
    textbox_record_count: int,
) -> dict[int, list[_TextboxBreak]]:
    """Parse selected ``PlcfTxbxBkd`` ranges grouped by FTXBXS index."""

    cps, records = _parse_plc(data, 6, "PlcfTxbxBkd")
    if not records:
        return {}

    result: dict[int, list[_TextboxBreak]] = {}
    for index, raw in enumerate(records[:-1]):
        # The final Tbkd has no associated textbox; its end CP is unused.
        if cps[index + 1] > cp_limit:
            raise LegacyDocError("PlcfTxbxBkd CP exceeds textbox-document limit")
        itxbxs = struct.unpack_from("<h", raw, 0)[0]
        if itxbxs < 0 or itxbxs >= textbox_record_count:
            raise LegacyDocError("PlcfTxbxBkd references an invalid FTXBXS index")
        result.setdefault(itxbxs, []).append(
            _TextboxBreak(start=cps[index], end=cps[index + 1])
        )
    return result


def _textbox_base(fib: FibInfo) -> int:
    # Story order in the document-wide CP space is Main, Footnotes, Headers,
    # Comments, Endnotes, Textboxes, Header Textboxes.  A textbox PLC stores
    # local CPs, so its absolute base is the sum of the preceding five parts.
    values = [
        fib.ccp_text,
        fib.ccp_ftn,
        fib.ccp_hdd,
        fib.ccp_atn,
        fib.ccp_edn,
    ]
    return _checked_sum(values, "textbox CP base")


def _parse_office_art(doc: BinaryDocument) -> _OfficeArtIndex:
    data = _read_fib_region(doc, _DGG_INFO_PAIR, "OfficeArtContent", allow_empty=True)
    if not data:
        raise LegacyDocError("OfficeArtContent is required for a body textbox")
    if len(data) < 8:
        raise LegacyDocError("OfficeArtContent is shorter than its DGG header")

    walker = _OfficeArtWalker(max_records=_MAX_OFFICEART_RECORDS)
    dgg = _read_record(data, 0, len(data), walker)
    if dgg.record_type != _DGG_CONTAINER or not dgg.is_container:
        raise LegacyDocError("OfficeArtContent does not begin with a DGG container")

    shapes: dict[int, _OfficeArtShape] = {}
    geometry_text_ids: list[int] = []
    pos = dgg.end
    labels_seen: set[int] = set()
    while pos < len(data):
        if len(data) - pos < 1:
            raise LegacyDocError("OfficeArtWordDrawing is truncated")
        dgglbl = data[pos]
        pos += 1
        if dgglbl not in (0, 1):
            raise LegacyDocError("OfficeArtWordDrawing has an invalid dgglbl")
        if dgglbl in labels_seen:
            raise LegacyDocError("OfficeArtContent contains duplicate drawing labels")
        labels_seen.add(dgglbl)

        drawing = _read_record(data, pos, len(data), walker)
        if drawing.record_type != _DG_CONTAINER or not drawing.is_container:
            raise LegacyDocError("OfficeArtWordDrawing does not contain a DG container")
        if dgglbl == 0:
            _walk_drawing(
                data,
                drawing.payload_start,
                drawing.end,
                walker,
                shapes,
                geometry_text_ids,
                depth=0,
            )
        # Header drawings are intentionally bounded at the container header
        # and skipped, since header textboxes are outside the body scope.
        pos = drawing.end

    return _OfficeArtIndex(
        shapes=shapes,
        geometry_text_shapes=tuple(geometry_text_ids),
    )


def _walk_drawing(
    data: bytes,
    start: int,
    end: int,
    walker: _OfficeArtWalker,
    shapes: dict[int, _OfficeArtShape],
    geometry_text_ids: list[int],
    *,
    depth: int,
) -> None:
    if depth > _MAX_OFFICEART_DEPTH:
        raise LegacyDocError("OfficeArt container nesting exceeds parser limit")
    pos = start
    while pos < end:
        record = _read_record(data, pos, end, walker)
        if record.record_type == _SP_CONTAINER:
            shape = _parse_shape_container(data, record, walker, depth=depth + 1)
            if shape is None:
                pos = record.end
                continue
            if shape.shape_id in shapes:
                raise LegacyDocError("OfficeArt contains duplicate active shape IDs")
            shapes[shape.shape_id] = shape
            if shape.geometry_text:
                geometry_text_ids.append(shape.shape_id)
        elif record.is_container:
            _walk_drawing(
                data,
                record.payload_start,
                record.end,
                walker,
                shapes,
                geometry_text_ids,
                depth=depth + 1,
            )
        pos = record.end
    if pos != end:
        raise LegacyDocError("OfficeArt drawing records do not fill their container")


def _parse_shape_container(
    data: bytes,
    container: _Record,
    walker: _OfficeArtWalker,
    *,
    depth: int,
) -> _OfficeArtShape | None:
    if depth > _MAX_OFFICEART_DEPTH:
        raise LegacyDocError("OfficeArt shape nesting exceeds parser limit")

    shape_id: int | None = None
    fsp_flags = 0
    anchor_index: int | None = None
    anchor_seen = False
    text_id: int | None = None
    text_chain_index: int | None = None
    geometry_text = False

    pos = container.payload_start
    while pos < container.end:
        record = _read_record(data, pos, container.end, walker)
        payload = data[record.payload_start : record.end]
        if record.record_type == _FSP:
            if shape_id is not None:
                raise LegacyDocError("OfficeArt shape has duplicate FSP records")
            if record.version != 2 or record.end - record.payload_start != 8:
                raise LegacyDocError("OfficeArtFSP has an invalid size or version")
            shape_id = struct.unpack_from("<I", payload, 0)[0]
            fsp_flags = struct.unpack_from("<I", payload, 4)[0]
        elif record.record_type == _CLIENT_ANCHOR:
            if anchor_seen:
                raise LegacyDocError(
                    "OfficeArt shape has duplicate ClientAnchor records"
                )
            if (
                record.version != 0
                or record.instance != 0
                or record.end - record.payload_start != 4
            ):
                raise LegacyDocError("OfficeArtClientAnchor has an invalid size")
            candidate = struct.unpack_from("<i", payload, 0)[0]
            anchor_seen = True
            if candidate == -1:
                anchor_index = None
            elif candidate < 0:
                raise LegacyDocError("OfficeArtClientAnchor has an invalid index")
            else:
                anchor_index = candidate
        elif record.record_type == _CLIENT_TEXTBOX:
            if text_id is not None:
                raise LegacyDocError(
                    "OfficeArt shape has duplicate ClientTextbox records"
                )
            if (
                record.version != 0
                or record.instance != 0
                or record.end - record.payload_start != 4
            ):
                raise LegacyDocError("OfficeArtClientTextbox has an invalid size")
            ltxid = struct.unpack_from("<I", payload, 0)[0]
            text_id = ltxid >> 16
            text_chain_index = ltxid & 0xFFFF
            if text_id == 0:
                raise LegacyDocError("OfficeArtClientTextbox has a zero lTxid")
        elif record.record_type in {_FOPT, _SECONDARY_FOPT, _TERTIARY_FOPT}:
            geometry_text = geometry_text or _has_geometry_text(record, payload)
        elif record.is_container:
            # Nested group records are valid in a shape tree.  Their shape IDs
            # are collected by the parent drawing walker; here we only need to
            # advance over them without assigning their fields to this shape.
            _validate_container_shape_children(
                data, record.payload_start, record.end, walker, depth=depth + 1
            )
        pos = record.end

    if shape_id is None:
        raise LegacyDocError("OfficeArtSpContainer is missing OfficeArtFSP")

    # fDeleted is bit 3 of the FSP flags.  Deleted shapes are not live body
    # objects and must not collide with an active shape ID.
    if fsp_flags & 0x00000008:
        return None
    return _OfficeArtShape(
        shape_id=shape_id,
        anchor_index=anchor_index,
        text_id=text_id,
        text_chain_index=text_chain_index,
        geometry_text=geometry_text,
    )


def _validate_container_shape_children(
    data: bytes,
    start: int,
    end: int,
    walker: _OfficeArtWalker,
    *,
    depth: int,
) -> None:
    if depth > _MAX_OFFICEART_DEPTH:
        raise LegacyDocError("OfficeArt container nesting exceeds parser limit")
    pos = start
    while pos < end:
        record = _read_record(data, pos, end, walker)
        if record.is_container:
            _validate_container_shape_children(
                data, record.payload_start, record.end, walker, depth=depth + 1
            )
        pos = record.end
    if pos != end:
        raise LegacyDocError("OfficeArt nested container is malformed")


def _has_geometry_text(record: _Record, payload: bytes) -> bool:
    # FOPT records contain recInstance six-byte FOPTE entries followed by
    # complex property data in entry order.  Only the property ID and its
    # bounded length are needed to establish the unsupported scope.
    count = record.instance
    fixed_size = count * 6
    if fixed_size > len(payload):
        raise LegacyDocError("OfficeArtFOPT property table is truncated")

    complex_pos = fixed_size
    found = False
    for index in range(count):
        opid, value = struct.unpack_from("<HI", payload, index * 6)
        is_complex = bool(opid & 0x8000)
        property_id = opid & 0x3FFF
        if is_complex:
            if value > len(payload) - complex_pos:
                raise LegacyDocError("OfficeArt complex property is truncated")
            if property_id == _GTEXT_UNICODE:
                found = True
            complex_pos += value
    return found


def _read_record(
    data: bytes,
    start: int,
    bound: int,
    walker: _OfficeArtWalker,
) -> _Record:
    if start < 0 or bound < start or bound > len(data) or bound - start < 8:
        raise LegacyDocError("OfficeArt record header is truncated")
    options, record_type, length = struct.unpack_from("<HHI", data, start)
    if not 0xF000 <= record_type <= 0xFFFF:
        raise LegacyDocError("OfficeArt record has an invalid type")
    end = start + 8 + length
    if end < start or end > bound:
        raise LegacyDocError("OfficeArt record exceeds its containing boundary")
    walker.record()
    return _Record(
        start=start,
        payload_start=start + 8,
        end=end,
        version=options & 0x000F,
        instance=(options >> 4) & 0x0FFF,
        record_type=record_type,
    )


def _checked_add(left: int, right: int, label: str) -> int:
    result = left + right
    if result < left or result > 0x7FFFFFFF:
        raise LegacyDocError(f"{label} exceeds CP bounds")
    return result


def _checked_sum(values: Iterable[int], label: str) -> int:
    result = 0
    for value in values:
        result = _checked_add(result, value, label)
    return result
