from __future__ import annotations

import struct
from dataclasses import dataclass, field
from types import SimpleNamespace

import pytest

from legacy_doc._shapes import (
    UNSUPPORTED_GEOMETRY_TEXT_SCOPE,
    TextboxRange,
    body_textboxes,
    _parse_spa,
    _parse_txbx,
    _parse_txbx_bkd,
    unsupported_geometry_text_shapes,
)
from legacy_doc.exceptions import LegacyDocError


@pytest.mark.parametrize('kind', ['anchor', 'textbox', 'linked-textbox'])
def test_unused_terminal_cp_does_not_limit_selected_text(kind):
    # MS-DOC: fcPlcSpaMom's final CP is undefined; the final FTXBXS
    # and Tbkd ranges are unused. Real Word fixtures use a document-wide
    # terminal CP that exceeds the selected story's length.
    if kind == 'anchor':
        assert _parse_spa(_plc([_spa(1025)], [0, 135]), 3)[0].cp == 0
    elif kind == 'textbox':
        records = [_ftxbxs(shape_id=1025), _ftxbxs(shape_id=0)]
        assert _parse_txbx(_plc(records, [0, 9, 135]), 10)[0].end == 9
    else:
        records = [_tbkd(), _tbkd()]
        assert _parse_txbx_bkd(_plc(records, [0, 9, 135]), 10, 2)[0][0].end == 9


@pytest.mark.parametrize('kind', ['anchor', 'textbox', 'linked-textbox'])
def test_used_cp_still_must_fit_its_document_part(kind):
    with pytest.raises(LegacyDocError, match='exceeds'):
        if kind == 'anchor':
            _parse_spa(_plc([_spa(1025)], [4, 135]), 3)
        elif kind == 'textbox':
            records = [_ftxbxs(shape_id=1025), _ftxbxs(shape_id=0)]
            _parse_txbx(_plc(records, [0, 11, 135]), 10)
        else:
            _parse_txbx_bkd(_plc([_tbkd(), _tbkd()], [0, 11, 135]), 10, 2)


@dataclass
class _Fib:
    pairs: dict[int, tuple[int, int]]
    ccp_text: int = 10
    ccp_ftn: int = 1
    ccp_hdd: int = 2
    ccp_atn: int = 3
    ccp_edn: int = 4
    ccp_txbx: int = 6

    def pair(self, index: int) -> tuple[int, int]:
        return self.pairs.get(index, (0, 0))


@dataclass
class _Document:
    table: bytes
    fib: _Fib
    cp_limit: int = 26
    options: object = field(default_factory=lambda: SimpleNamespace())


def _plc(records: list[bytes], cps: list[int]) -> bytes:
    assert len(cps) == len(records) + 1
    return struct.pack(f"<{len(cps)}I", *cps) + b"".join(records)


def _spa(shape_id: int) -> bytes:
    return struct.pack("<I", shape_id) + b"\0" * 22


def _ftxbxs(*, shape_id: int, reusable: bool = False, chain_count: int = 1) -> bytes:
    raw = bytearray(22)
    # The non-reusable union starts at offset zero.  The common fields follow
    # it: fReusable@8, itxbxsDest@10, lid@14, txidUndo@18.
    struct.pack_into("<i", raw, 0, chain_count)
    struct.pack_into("<H", raw, 8, 1 if reusable else 0)
    struct.pack_into("<I", raw, 14, shape_id)
    return bytes(raw)


def _tbkd(*, textbox_index: int = 0) -> bytes:
    raw = bytearray(6)
    struct.pack_into("<h", raw, 0, textbox_index)
    return bytes(raw)


def _art_record(
    record_type: int,
    payload: bytes,
    *,
    version: int = 0xF,
    instance: int = 0,
) -> bytes:
    options = (instance << 4) | version
    return struct.pack("<HHI", options, record_type, len(payload)) + payload


def _fsp(shape_id: int, *, deleted: bool = False) -> bytes:
    return _art_record(
        0xF00A,
        struct.pack("<II", shape_id, 0x8 if deleted else 0),
        version=2,
    )


def _sp(
    shape_id: int,
    *,
    anchor_index: int | None = 0,
    textbox_index: int | None = 1,
    chain_index: int = 0,
    geometry_text: bool = False,
) -> bytes:
    children = [_fsp(shape_id)]
    if anchor_index is not None:
        children.append(_art_record(0xF010, struct.pack("<i", anchor_index), version=0))
    if textbox_index is not None:
        ltxid = (textbox_index << 16) | chain_index
        children.append(_art_record(0xF00D, struct.pack("<I", ltxid), version=0))
    if geometry_text:
        # FOPT has one complex gtextUNICODE property (ID 0x00C0).  The
        # complex string itself is irrelevant to the association result; the
        # bounded parser only needs its declared length.
        value = "字".encode("utf-16le") + b"\0\0"
        opid = 0x8000 | 0x00C0
        payload = struct.pack("<HI", opid, len(value)) + value
        children.append(_art_record(0xF00B, payload, version=3, instance=1))
    return _art_record(0xF004, b"".join(children))


def _office_art(*shapes: bytes, header_shape: bytes | None = None) -> bytes:
    shape_group = _art_record(0xF003, b"".join(shapes))
    main_dg = _art_record(0xF002, shape_group)
    content = _art_record(0xF000, b"") + b"\0" + main_dg
    if header_shape is not None:
        header_group = _art_record(0xF003, header_shape)
        header_dg = _art_record(0xF002, header_group)
        content += b"\1" + header_dg
    return content


def _document(
    *,
    anchors: list[bytes],
    anchor_cps: list[int],
    textboxes: list[bytes],
    textbox_cps: list[int],
    art: bytes,
    bkd: bytes | None = None,
    ccp_text: int = 10,
    ccp_txbx: int = 6,
    counts: tuple[int, int, int, int] = (1, 2, 3, 4),
) -> _Document:
    table = bytearray()
    def put(index: int, value: bytes) -> None:
        offset = len(table)
        table.extend(value)
        pairs[index] = (offset, len(value))

    pairs: dict[int, tuple[int, int]] = {}
    put(40, _plc(anchors, anchor_cps))
    put(56, _plc(textboxes, textbox_cps))
    put(50, art)
    if bkd is not None:
        put(75, bkd)
    fib = _Fib(
        pairs=pairs,
        ccp_text=ccp_text,
        ccp_ftn=counts[0],
        ccp_hdd=counts[1],
        ccp_atn=counts[2],
        ccp_edn=counts[3],
        ccp_txbx=ccp_txbx,
    )
    base = sum((ccp_text, *counts))
    return _Document(table=bytes(table), fib=fib, cp_limit=base + ccp_txbx)


def test_body_textboxes_resolves_absolute_range_and_shape_evidence() -> None:
    doc = _document(
        anchors=[_spa(42)],
        anchor_cps=[3, 10],
        textboxes=[_ftxbxs(shape_id=42), _ftxbxs(shape_id=0, reusable=True)],
        textbox_cps=[0, 5, 6],
        art=_office_art(_sp(42)),
    )

    result = body_textboxes(doc)

    assert result == {
        3: TextboxRange(
            start=20,
            end=25,
            shape_id=42,
            textbox_index=0,
            chain_index=0,
        )
    }


def test_reusable_slots_and_last_slot_are_ignored() -> None:
    doc = _document(
        anchors=[_spa(42)],
        anchor_cps=[3, 10],
        textboxes=[
            _ftxbxs(shape_id=42, reusable=True),
            _ftxbxs(shape_id=42),
            _ftxbxs(shape_id=0, reusable=False),
        ],
        textbox_cps=[0, 1, 5, 6],
        art=_office_art(_sp(42, textbox_index=2)),
    )

    result = body_textboxes(doc)

    assert result[3].textbox_index == 1
    assert result[3].start == 21
    assert result[3].end == 25


def test_header_drawing_is_not_used_for_body_textbox() -> None:
    doc = _document(
        anchors=[_spa(42)],
        anchor_cps=[3, 10],
        textboxes=[_ftxbxs(shape_id=42), _ftxbxs(shape_id=0, reusable=True)],
        textbox_cps=[0, 5, 6],
        art=_office_art(
            _sp(42, anchor_index=None, textbox_index=None),
            header_shape=_sp(42),
        ),
    )

    with pytest.raises(LegacyDocError, match="ClientAnchor"):
        body_textboxes(doc)


def test_ltxid_must_point_to_one_based_ftxbxs_index() -> None:
    doc = _document(
        anchors=[_spa(42)],
        anchor_cps=[3, 10],
        textboxes=[_ftxbxs(shape_id=42), _ftxbxs(shape_id=0, reusable=True)],
        textbox_cps=[0, 5, 6],
        art=_office_art(_sp(42, textbox_index=2)),
    )

    with pytest.raises(LegacyDocError, match="lTxid"):
        body_textboxes(doc)


def test_non_textbox_shape_does_not_require_textbox_associations() -> None:
    doc = _document(
        anchors=[_spa(99)],
        anchor_cps=[3, 10],
        textboxes=[],
        textbox_cps=[0],
        art=b"",
    )

    assert body_textboxes(doc) == {}


def test_geometry_text_is_reported_as_unsupported_scope() -> None:
    doc = _document(
        anchors=[],
        anchor_cps=[10],
        textboxes=[],
        textbox_cps=[6],
        art=_office_art(_sp(42, geometry_text=True)),
    )

    scopes = unsupported_geometry_text_shapes(doc)

    assert len(scopes) == 1
    assert scopes[0].shape_id == 42
    assert scopes[0].kind == "gtextUNICODE_complex"
    assert scopes[0].reason == UNSUPPORTED_GEOMETRY_TEXT_SCOPE


def test_body_geometry_text_anchor_fails_explicitly() -> None:
    doc = _document(
        anchors=[_spa(42)],
        anchor_cps=[3, 10],
        textboxes=[],
        textbox_cps=[6],
        art=_office_art(_sp(42, geometry_text=True)),
    )

    with pytest.raises(LegacyDocError, match="gtextUNICODE_complex"):
        body_textboxes(doc)


def test_selected_client_textbox_without_ftxbxs_fails() -> None:
    doc = _document(
        anchors=[_spa(42)],
        anchor_cps=[3, 10],
        textboxes=[],
        textbox_cps=[6],
        art=_office_art(_sp(42)),
    )

    with pytest.raises(LegacyDocError, match="FTXBXS"):
        body_textboxes(doc)


def test_linked_textbox_components_use_txbx_bkd_ranges() -> None:
    doc = _document(
        anchors=[_spa(42), _spa(43)],
        anchor_cps=[3, 8, 10],
        textboxes=[_ftxbxs(shape_id=42, chain_count=2), _ftxbxs(shape_id=0, reusable=True)],
        textbox_cps=[0, 8, 9],
        art=_office_art(
            _sp(42, textbox_index=1, chain_index=0),
            _sp(43, anchor_index=1, textbox_index=1, chain_index=1),
        ),
        bkd=_plc(
            [_tbkd(textbox_index=0), _tbkd(textbox_index=0), _tbkd(textbox_index=99)],
            [0, 4, 8, 9],
        ),
        ccp_txbx=9,
    )

    result = body_textboxes(doc)

    assert result[3] == TextboxRange(
        start=20, end=24, shape_id=42, textbox_index=0, chain_index=0
    )
    assert result[8] == TextboxRange(
        start=24, end=28, shape_id=43, textbox_index=0, chain_index=1
    )


def test_linked_textbox_without_bkd_fails_explicitly() -> None:
    doc = _document(
        anchors=[_spa(42)],
        anchor_cps=[3, 10],
        textboxes=[_ftxbxs(shape_id=42, chain_count=2), _ftxbxs(shape_id=0, reusable=True)],
        textbox_cps=[0, 5, 6],
        art=_office_art(_sp(42, textbox_index=1)),
    )

    with pytest.raises(LegacyDocError, match="PlcfTxbxBkd"):
        body_textboxes(doc)


@pytest.mark.parametrize(
    ("bkd_records", "bkd_cps", "message"),
    [
        ([_tbkd(textbox_index=0), _tbkd(textbox_index=0), _tbkd(textbox_index=99)],
         [1, 4, 8, 9], "cover"),
        ([_tbkd(textbox_index=0), _tbkd(textbox_index=1), _tbkd(textbox_index=0),
         _tbkd(textbox_index=99)], [0, 4, 6, 8, 9], "contiguous"),
    ],
)
def test_linked_textbox_bkd_must_cover_one_contiguous_ftxbxs_range(
    bkd_records: list[bytes], bkd_cps: list[int], message: str
) -> None:
    doc = _document(
        anchors=[_spa(42), _spa(43)],
        anchor_cps=[3, 8, 10],
        textboxes=[_ftxbxs(shape_id=42, chain_count=2), _ftxbxs(shape_id=0, reusable=True)],
        textbox_cps=[0, 8, 9],
        art=_office_art(
            _sp(42, textbox_index=1, chain_index=0),
            _sp(43, anchor_index=1, textbox_index=1, chain_index=1),
        ),
        bkd=_plc(bkd_records, bkd_cps),
        ccp_txbx=9,
    )

    with pytest.raises(LegacyDocError, match=f"PlcfTxbxBkd.*{message}"):
        body_textboxes(doc)


def test_linked_textbox_requires_every_chain_component() -> None:
    doc = _document(
        anchors=[_spa(42)],
        anchor_cps=[3, 10],
        textboxes=[_ftxbxs(shape_id=42, chain_count=2), _ftxbxs(shape_id=0, reusable=True)],
        textbox_cps=[0, 8, 9],
        art=_office_art(_sp(42, textbox_index=1, chain_index=0)),
        bkd=_plc(
            [_tbkd(textbox_index=0), _tbkd(textbox_index=0), _tbkd(textbox_index=99)],
            [0, 4, 8, 9],
        ),
        ccp_txbx=9,
    )

    with pytest.raises(LegacyDocError, match="every component"):
        body_textboxes(doc)


def test_required_officeart_association_damage_fails() -> None:
    doc = _document(
        anchors=[_spa(42)],
        anchor_cps=[3, 10],
        textboxes=[_ftxbxs(shape_id=42), _ftxbxs(shape_id=0, reusable=True)],
        textbox_cps=[0, 5, 6],
        art=_office_art(_art_record(0xF004, _fsp(42))),
    )

    with pytest.raises(LegacyDocError, match="ClientAnchor"):
        body_textboxes(doc)
