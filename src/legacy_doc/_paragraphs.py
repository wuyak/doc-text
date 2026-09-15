"""Paragraph and table-boundary records for the legacy Word binary format.

The public text extractor deliberately keeps the document model small.  This
module is the corresponding small reader for the paragraph records that are
needed by the text assembler.  It does not try to implement Word layout; it
only follows the records which describe paragraph boundaries, table depth and
row/cell state.

The document reader is owned by :mod:`legacy_doc.word`; this module consumes
its narrow paragraph interface: ``word``, ``table``, ``fib``, text/FC lookups,
CLX PRC groups, and bounded ``read_data``.
"""

from __future__ import annotations

from bisect import bisect_right
from dataclasses import dataclass
import struct
from typing import Iterable, Iterator

from legacy_doc._binary import BinaryDocument
from legacy_doc.exceptions import LegacyDocError


# ---------------------------------------------------------------------------
# Records consumed by the table/text assembler


@dataclass(frozen=True)
class CellFormat:
    """The merge flags from one ``TC80.tcgrf`` record.

    ``horizontal_merge`` is the two-bit ``TCGRF.horzMerge`` value.  Values 2
    and 3 both mean that this cell starts a horizontal merge according to
    MS-DOC; preserving the value is useful because it is what was stored in
    the file.  ``vertical_merge`` is the raw two-bit ``TCGRF.vertMerge`` /
    ``VerticalMergeFlag`` value (0, 1, 2, or 3).
    """

    horizontal_merge: int = 0
    vertical_merge: int = 0


@dataclass(frozen=True)
class Paragraph:
    """One paragraph or table/cell terminating record.

    ``start`` and ``end`` are CP coordinates.  ``end`` is exclusive and
    includes the terminating character when one exists.  ``text`` excludes
    that terminator.  The row's ``cells`` are present only on a row-ending
    record when a valid ``TDefTable`` definition was available.
    """

    start: int
    end: int
    text: str
    in_table: bool
    depth: int
    cell_end: bool
    row_end: bool
    cells: tuple[CellFormat, ...] | None = None


# ---------------------------------------------------------------------------
# MS-DOC SPRM identifiers used by paragraph/table reconstruction


SPRM_PF_IN_TABLE = 0x2416
SPRM_PF_TTP = 0x2417
SPRM_P_ITAP = 0x6649
SPRM_P_DTAP = 0x664A
SPRM_PF_INNER_TABLE_CELL = 0x244B
SPRM_PF_INNER_TTP = 0x244C
SPRM_P_HUGE_PAPX = 0x6646
SPRM_P_TABLE_PROPS = 0x646B

SPRM_T_DEF_TABLE = 0xD608
SPRM_T_INSERT = 0x7621
SPRM_T_DELETE = 0x5622
SPRM_T_MERGE = 0x5624
SPRM_T_SPLIT = 0x5625
SPRM_T_VERT_MERGE = 0xD62B


# The fComplex=0 PRM table contains compact one-byte paragraph properties.
# Only paragraph properties are useful to this module.  Keeping the full
# mapping makes the handling of an unknown compact PRM explicit: it is
# ignored as a character property rather than being guessed as a paragraph
# property.
_PRM0_PARAGRAPH_SPRMS: dict[int, int] = {
    0x04: 0x2404,  # sprmPIncLvl
    0x05: 0x2405,  # sprmPJc
    0x07: 0x2407,  # sprmPFKeep
    0x08: 0x2408,  # sprmPFKeepFollow
    0x09: 0x2409,  # sprmPFPageBreakBefore
    0x0C: 0x240C,  # sprmPFNoLineNumb
    0x0D: 0x240D,  # sprmPIlvl
    0x0E: 0x240E,  # sprmPFMirrorIndents
    0x0F: 0x240F,  # sprmPTwo
    0x18: SPRM_PF_IN_TABLE,
    0x19: SPRM_PF_TTP,
    0x1D: 0x241D,  # sprmPPc
    0x25: 0x2425,  # sprmPWr
    0x2C: 0x242C,  # sprmPFNoAutoHyph
    0x32: 0x2432,  # sprmPFLocked
    0x33: 0x2433,  # sprmPFWidowControl
    0x35: 0x2435,  # sprmPFKinsoku
    0x36: 0x2436,  # sprmPFWordWrap
    0x37: 0x2437,  # sprmPFOverflowPunct
    0x38: 0x2438,  # sprmPFTopLinePunct
    0x39: 0x2439,  # sprmPFAutoSpaceDE
    0x3A: 0x243A,  # sprmPFAutoSpaceDN
    0x40: 0x2640,  # sprmPOutLvl
    0x7E: 0x247E,  # sprmPFNumRMIns
}

# Limits are deliberately independent of the OLE sector limit.  A malformed
# PAPX can otherwise amplify one valid page into an unbounded number of
# entries, and indirect PrcData records can point back to one another.
_FKP_SIZE = 512
_MAX_FKP_PAGES = 1 << 15
_MAX_FKP_ENTRIES = 1 << 20
_MAX_INDIRECT_DEPTH = 32
_MAX_PRC_BYTES = 0x3FA2
_MAX_ROW_CELLS = 63
_MAX_SECTIONS = 1 << 20
_MAX_SPRMS = 1 << 20


@dataclass(frozen=True)
class _Sprm:
    opcode: int
    operand: bytes


@dataclass
class _ParagraphProperties:
    in_table: bool = False
    ttp: bool = False
    inner_cell: bool = False
    inner_ttp: bool = False
    depth: int = 0
    cells: list[CellFormat] | None = None

    def clone(self) -> "_ParagraphProperties":
        return _ParagraphProperties(
            in_table=self.in_table,
            ttp=self.ttp,
            inner_cell=self.inner_cell,
            inner_ttp=self.inner_ttp,
            depth=self.depth,
            cells=None if self.cells is None else list(self.cells),
        )


@dataclass(frozen=True)
class _FkpRun:
    fc_start: int
    fc_end: int
    grpprl: bytes | None


def _raise(message: str) -> None:
    raise LegacyDocError(message)


def _u16(data: bytes | bytearray | memoryview, offset: int) -> int:
    if offset < 0 or offset + 2 > len(data):
        _raise("Legacy .doc paragraph property data is truncated")
    return struct.unpack_from("<H", data, offset)[0]


def _i16(data: bytes | bytearray | memoryview, offset: int) -> int:
    if offset < 0 or offset + 2 > len(data):
        _raise("Legacy .doc table definition is truncated")
    return struct.unpack_from("<h", data, offset)[0]


def _u32(data: bytes | bytearray | memoryview, offset: int) -> int:
    if offset < 0 or offset + 4 > len(data):
        _raise("Legacy .doc paragraph property data is truncated")
    return struct.unpack_from("<I", data, offset)[0]


def _i32(data: bytes | bytearray | memoryview, offset: int) -> int:
    if offset < 0 or offset + 4 > len(data):
        _raise("Legacy .doc paragraph property data is truncated")
    return struct.unpack_from("<i", data, offset)[0]


# ---------------------------------------------------------------------------
# SPRM decoding


def _pchg_tabs_operand_len(data: bytes, offset: int) -> int:
    """Return the complete PChgTabs operand length, including its ``cb`` byte.

    ``cb=255`` is the extended form.  Its actual length is derived from the
    two record counts; treating it as the ordinary one-byte length is the
    classic failure mode that causes the next SPRM to be read in the middle of
    a tab operation.
    """

    if offset >= len(data):
        _raise("Legacy .doc PChgTabs operand is truncated")
    cb = data[offset]
    if cb < 2:
        _raise("Legacy .doc PChgTabs operand has an invalid length")
    del_count_offset = offset + 1
    if del_count_offset >= len(data):
        _raise("Legacy .doc PChgTabs operand is truncated")
    del_count = data[del_count_offset]
    if del_count > 64:
        _raise("Legacy .doc PChgTabs operand has too many deletions")
    # PChgTabsDelClose stores both the deleted position and its close
    # position, each as a four-byte signed integer.
    add_count_offset = del_count_offset + 1 + 8 * del_count
    if add_count_offset >= len(data):
        _raise("Legacy .doc PChgTabs operand is truncated")
    add_count = data[add_count_offset]
    if add_count > 64:
        _raise("Legacy .doc PChgTabs operand has too many additions")
    # Each added tab carries a four-byte position and one-byte descriptor.
    remainder_end = add_count_offset + 1 + 5 * add_count
    # ``remainder_len`` includes the two count bytes, but excludes cb.
    remainder_len = remainder_end - del_count_offset
    if cb != 0xFF and cb != remainder_len:
        _raise("Legacy .doc PChgTabs operand length does not match its records")
    return remainder_len + 1


def _operand_len(opcode: int, data: bytes, offset: int) -> int:
    """Decode the operand width from the SPRM's ``spra`` bits.

    The three high bits of a Word SPRM encode ``spra``.  In particular, the
    variable case is not always a one-byte count: TDefTable has a two-byte
    count whose value is the remaining byte count plus one, and PChgTabs has
    the extended 0xFF form handled above.
    """

    spra = (opcode >> 13) & 0x7
    if spra in (0, 1):
        return 1
    if spra in (2, 4, 5):
        return 2
    if spra == 3:
        return 4
    if spra == 7:
        return 3
    if spra != 6:
        _raise("Legacy .doc SPRM has an invalid operand class")

    if opcode == SPRM_T_DEF_TABLE:
        cb = _u16(data, offset)
        if cb == 0:
            _raise("Legacy .doc TDefTable operand has an invalid length")
        return cb + 1
    if opcode == 0xC615:  # sprmPChgTabs
        return _pchg_tabs_operand_len(data, offset)
    if offset >= len(data):
        _raise("Legacy .doc variable SPRM is truncated")
    return 1 + data[offset]


def _parse_sprms(grpprl: bytes | bytearray | memoryview) -> tuple[_Sprm, ...]:
    """Parse a complete ``GrpPrl`` without losing ordered duplicate SPRMs."""

    data = bytes(grpprl)
    result: list[_Sprm] = []
    pos = 0
    while pos < len(data):
        if len(result) >= _MAX_SPRMS:
            _raise("Legacy .doc paragraph properties exceed the SPRM limit")
        if len(data) - pos < 2:
            _raise("Legacy .doc grpprl has an orphaned SPRM byte")
        opcode = _u16(data, pos)
        operand_start = pos + 2
        width = _operand_len(opcode, data, operand_start)
        operand_end = operand_start + width
        if operand_end > len(data):
            _raise("Legacy .doc SPRM operand is truncated")
        result.append(_Sprm(opcode, data[operand_start:operand_end]))
        pos = operand_end
    return tuple(result)


# A public-ish alias is useful to focused tests and keeps the parser's
# boundary logic independently testable without exposing internal state.
parse_sprms = _parse_sprms


def _prm_value(value: int) -> int:
    if not isinstance(value, int) or isinstance(value, bool):
        _raise("Legacy .doc PCD PRM is not an integer")
    return value & 0xFFFF


def _prm_sprms(doc: BinaryDocument, prm_value: int) -> tuple[_Sprm, ...]:
    """Resolve a PCD PRM, retaining PRM0/PRM1 order after PAPX."""

    prm = _prm_value(prm_value)
    if prm & 1 == 0:  # Prm0: one compact SPRM and one-byte operand.
        isprm = (prm >> 1) & 0x7F
        operand = bytes([(prm >> 8) & 0xFF])
        if isprm == 0 and operand[0] == 0:
            return ()
        opcode = _PRM0_PARAGRAPH_SPRMS.get(isprm)
        if opcode is None:
            return ()
        return (_Sprm(opcode, operand),)

    # Prm1: the remaining 15 bits index Clx.RgPrc.  BinaryDocument exposes
    # the extracted grpprl list, so this path never reparses CLX itself.
    index = prm >> 1
    prcs = doc.prcs
    try:
        prc = prcs[index]
    except (IndexError, KeyError, TypeError):
        _raise("Legacy .doc Prm1 refers outside the CLX property groups")
    grpprl = _grpprl_from_prc(prc)
    return _parse_sprms(grpprl)


def _grpprl_from_prc(prc: bytes) -> bytes:
    """BinaryDocument stores each CLX Prc as its raw GrpPrl bytes."""

    return bytes(prc)


# ---------------------------------------------------------------------------
# PAPX FKP and row-definition parsing


def _parse_tdef_table(operand: bytes) -> list[CellFormat]:
    """Parse a TDefTable operand into the row's physical cell definitions."""

    if len(operand) < 3:
        _raise("Legacy .doc TDefTable operand is truncated")
    declared = _u16(operand, 0)
    if declared == 0 or declared + 1 != len(operand):
        _raise("Legacy .doc TDefTable length is invalid")
    column_count = operand[2]
    if column_count > _MAX_ROW_CELLS:
        _raise("Legacy .doc TDefTable has too many cells")
    centers_start = 3
    centers_end = centers_start + 2 * (column_count + 1)
    if centers_end > len(operand):
        _raise("Legacy .doc TDefTable cell boundaries are truncated")
    previous: int | None = None
    for pos in range(centers_start, centers_end, 2):
        center = _i16(operand, pos)
        if previous is not None and center < previous:
            _raise("Legacy .doc TDefTable cell boundaries are not ordered")
        previous = center

    tc_bytes = len(operand) - centers_end
    if tc_bytes % 20:
        _raise("Legacy .doc TDefTable has a partial TC80 record")
    tc_count = min(column_count, tc_bytes // 20)
    cells = [CellFormat() for _ in range(column_count)]
    for index in range(tc_count):
        tcgrf = _u16(operand, centers_end + index * 20)
        cells[index] = CellFormat(
            horizontal_merge=tcgrf & 0x03,
            vertical_merge=(tcgrf >> 5) & 0x03,
        )
    # TC80 records beyond NumberOfColumns are ignored by MS-DOC, but their
    # bytes must still have been a complete sequence (checked above).
    return cells


def _itc_first_lim(operand: bytes, cells: list[CellFormat]) -> tuple[int, int]:
    if len(operand) != 2:
        _raise("Legacy .doc table cell range operand is invalid")
    first, limit = operand
    if first > limit or limit > len(cells):
        _raise("Legacy .doc table cell range is outside the row")
    return first, limit


def _apply_row_sprm(state: _ParagraphProperties, item: _Sprm) -> bool:
    """Apply a table row modifier; return whether the item was consumed."""

    opcode, operand = item.opcode, item.operand
    cells = state.cells
    if opcode == SPRM_T_DEF_TABLE:
        state.cells = _parse_tdef_table(operand)
        return True
    if opcode == SPRM_T_INSERT:
        if len(operand) != 4:
            _raise("Legacy .doc TInsert operand is invalid")
        if state.cells is None:
            state.cells = []
        first, count = operand[0], operand[1]
        width = _u16(operand, 2)
        if count == 0 or first > len(state.cells) or len(state.cells) + count > _MAX_ROW_CELLS:
            _raise("Legacy .doc TInsert range is invalid")
        # Geometry (dxaCol) is intentionally outside this record's contract;
        # each inserted cell receives the default TCGRF merge state.
        del width
        state.cells[first:first] = [CellFormat() for _ in range(count)]
        return True
    if opcode == SPRM_T_DELETE:
        if cells is None:
            state.cells = []
            cells = state.cells
        first, limit = _itc_first_lim(operand, cells)
        if first == limit or limit - first >= len(cells):
            _raise("Legacy .doc TDelete must leave at least one cell")
        del cells[first:limit]
        return True
    if opcode in (SPRM_T_MERGE, SPRM_T_SPLIT):
        if cells is None:
            state.cells = []
            cells = state.cells
        first, limit = _itc_first_lim(operand, cells)
        if opcode == SPRM_T_SPLIT:
            for index in range(first, limit):
                cells[index] = CellFormat(
                    horizontal_merge=0,
                    vertical_merge=cells[index].vertical_merge,
                )
        elif limit - first >= 2:
            # TCGRF permits both 2 and 3 for a merge-start cell.  A direct
            # TMerge has no extra bit to preserve, so use the canonical 2.
            cells[first] = CellFormat(
                horizontal_merge=2,
                vertical_merge=cells[first].vertical_merge,
            )
            for index in range(first + 1, limit):
                cells[index] = CellFormat(
                    horizontal_merge=1,
                    vertical_merge=cells[index].vertical_merge,
                )
        return True
    if opcode == SPRM_T_VERT_MERGE:
        if len(operand) != 3 or operand[0] != 2:
            _raise("Legacy .doc TVertMerge operand is invalid")
        if cells is None:
            state.cells = []
            cells = state.cells
        index, flags = operand[1], operand[2]
        if index >= len(cells) or flags > 3:
            _raise("Legacy .doc TVertMerge cell is outside the row")
        cells[index] = CellFormat(
            horizontal_merge=cells[index].horizontal_merge,
            vertical_merge=flags,
        )
        return True
    return False


def _apply_regular_sprm(
    state: _ParagraphProperties,
    item: _Sprm,
    *,
    doc: BinaryDocument,
    indirect_depth: int,
    seen_indirect: set[int],
) -> tuple[bool, bool]:
    """Apply one SPRM.

    Returns ``(consumed, stop_group)``.  The latter encodes the MS-DOC rule
    that processing a PHugePapx/PTableProps PrcData stops the containing group
    at that property.
    """

    opcode, operand = item.opcode, item.operand
    if opcode == SPRM_PF_IN_TABLE:
        state.in_table = bool(operand[0])
        return True, False
    if opcode == SPRM_PF_TTP:
        state.ttp = bool(operand[0])
        return True, False
    if opcode == SPRM_PF_INNER_TABLE_CELL:
        state.inner_cell = bool(operand[0])
        return True, False
    if opcode == SPRM_PF_INNER_TTP:
        state.inner_ttp = bool(operand[0])
        return True, False
    if opcode == SPRM_P_ITAP:
        depth = _u32(operand, 0)
        state.depth = depth
        return True, False
    if opcode == SPRM_P_DTAP:
        delta = _i32(operand, 0)
        depth = state.depth + delta
        if depth < 0:
            _raise("Legacy .doc PDtap produces a negative table depth")
        state.depth = depth
        return True, False
    if opcode in (SPRM_P_HUGE_PAPX, SPRM_P_TABLE_PROPS):
        offset = _u32(operand, 0)
        _apply_indirect_group(
            state,
            doc,
            offset,
            indirect_depth=indirect_depth,
            seen_indirect=seen_indirect,
        )
        return True, True
    if _apply_row_sprm(state, item):
        return True, False
    return False, False


def _read_prc_data(doc: BinaryDocument, offset: int) -> bytes:
    reader = doc.read_data
    if offset < 0:
        _raise("Legacy .doc indirect paragraph property offset is invalid")
    try:
        header = reader(offset, 2)
    except LegacyDocError:
        raise
    except Exception as exc:
        _raise(f"Unable to read indirect paragraph properties: {exc}")
    if len(header) != 2:
        _raise("Legacy .doc indirect paragraph property header is truncated")
    cb = _u16(header, 0)
    if cb < 10 or cb > _MAX_PRC_BYTES:
        _raise("Legacy .doc indirect paragraph property length is invalid")
    try:
        grpprl = reader(offset + 2, cb)
    except LegacyDocError:
        raise
    except Exception as exc:
        _raise(f"Unable to read indirect paragraph properties: {exc}")
    if len(grpprl) != cb:
        _raise("Legacy .doc indirect paragraph property data is truncated")
    return grpprl


def _apply_indirect_group(
    state: _ParagraphProperties,
    doc: BinaryDocument,
    offset: int,
    *,
    indirect_depth: int,
    seen_indirect: set[int],
) -> None:
    if indirect_depth >= _MAX_INDIRECT_DEPTH:
        _raise("Legacy .doc indirect paragraph properties exceed the depth limit")
    if offset in seen_indirect:
        _raise("Legacy .doc indirect paragraph properties contain a cycle")
    seen_indirect.add(offset)
    try:
        grpprl = _read_prc_data(doc, offset)
        items = _parse_sprms(grpprl)
        _apply_group(
            state,
            items,
            doc=doc,
            indirect_depth=indirect_depth + 1,
            seen_indirect=seen_indirect,
        )
    finally:
        seen_indirect.remove(offset)


def _apply_group(
    state: _ParagraphProperties,
    items: Iterable[_Sprm],
    *,
    doc: BinaryDocument,
    indirect_depth: int = 0,
    seen_indirect: set[int] | None = None,
) -> None:
    if seen_indirect is None:
        seen_indirect = set()
    for index, item in enumerate(items):
        if item.opcode == SPRM_P_HUGE_PAPX and index != 0:
            # PHugePapx is meaningful only as the first item in its group.
            # It is still a fully-sized SPRM, so the parser has already
            # advanced over it correctly.  In particular, do not dereference
            # an ignored offset from the Data stream.
            continue
        consumed, stop = _apply_regular_sprm(
            state,
            item,
            doc=doc,
            indirect_depth=indirect_depth,
            seen_indirect=seen_indirect,
        )
        if consumed and stop:
            break


def _resolve_properties(doc: BinaryDocument, grpprl: bytes | None, prm: int = 0) -> _ParagraphProperties:
    state = _ParagraphProperties()
    if grpprl:
        _apply_group(state, _parse_sprms(grpprl), doc=doc)
    prm_items = _prm_sprms(doc, prm)
    if prm_items:
        _apply_group(state, prm_items, doc=doc)
    return state


def _parse_papx_in_fkp(page: bytes, offset: int, limit: int) -> bytes:
    if offset < 0 or offset >= limit:
        _raise("Legacy .doc PAPX offset is outside its FKP")
    cb = page[offset]
    if cb:
        width = 2 * cb - 1
        data_start = offset + 1
        data_end = data_start + width
    else:
        if offset + 1 >= limit:
            _raise("Legacy .doc PAPX extended length is truncated")
        cb_prime = page[offset + 1]
        if cb_prime == 0:
            _raise("Legacy .doc PAPX extended length is invalid")
        data_start = offset + 2
        data_end = data_start + 2 * cb_prime
    if data_end > limit or data_end < data_start:
        _raise("Legacy .doc PAPX extends beyond its FKP")
    if data_end - data_start < 2:
        _raise("Legacy .doc PAPX GrpPrlAndIstd is truncated")
    # The first two bytes are the style index (istd).  Paragraph structure is
    # carried by the subsequent grpprl and must remain in its stored order.
    return page[data_start + 2 : data_end]


def _parse_fkp(word: bytes, page_offset: int, cb_mac: int) -> tuple[_FkpRun, ...]:
    if page_offset < 0 or page_offset + _FKP_SIZE > len(word) or page_offset + _FKP_SIZE > cb_mac:
        _raise("Legacy .doc PAPX FKP is outside WordDocument cbMac")
    page = word[page_offset : page_offset + _FKP_SIZE]
    cpara = page[511]
    if cpara == 0 or cpara > 0x1D:
        _raise("Legacy .doc PAPX FKP has an invalid paragraph count")
    rgfc_end = 4 * (cpara + 1)
    bx_end = rgfc_end + 13 * cpara
    if bx_end > 511:
        _raise("Legacy .doc PAPX FKP fixed records overlap cpara")
    rgfc = [_u32(page, 4 * index) for index in range(cpara + 1)]
    if any(left >= right for left, right in zip(rgfc, rgfc[1:])):
        _raise("Legacy .doc PAPX FKP FC boundaries are not strictly ordered")
    runs: list[_FkpRun] = []
    for index in range(cpara):
        bx_offset = rgfc_end + 13 * index
        b_offset = page[bx_offset]
        if b_offset == 0:
            grpprl = None
        else:
            papx_offset = 2 * b_offset
            if papx_offset < bx_end or papx_offset >= 511:
                _raise("Legacy .doc PAPX bOffset is outside the FKP data area")
            grpprl = _parse_papx_in_fkp(page, papx_offset, 511)
        runs.append(_FkpRun(rgfc[index], rgfc[index + 1], grpprl))
    return tuple(runs)


class _SectionIndex:
    """Main-story section marks from the FIB-specified PlcfSed.

    U+000C is also used for a manual page break.  The paragraph-boundary
    algorithm treats only the character immediately before a non-final
    section boundary in PlcfSed as an end-of-section character.  Keeping this
    lookup separate means a textbox's local 0x0C is never classified with the
    main story's section table.
    """

    def __init__(self, doc: BinaryDocument):
        self._doc = doc
        self._marks: frozenset[int] | None = None

    def _load(self) -> frozenset[int]:
        if self._marks is not None:
            return self._marks
        fc, lcb = self._doc.fib.pair(6)
        if fc < 0 or lcb < 0:
            _raise("Legacy .doc PlcfSed location is invalid")
        if lcb == 0:
            self._marks = frozenset()
            return self._marks
        table = self._doc.table
        if fc > len(table) or lcb > len(table) - fc:
            _raise("Legacy .doc PlcfSed is outside the Table stream")
        if lcb < 4 or (lcb - 4) % 16:
            _raise("Legacy .doc PlcfSed has an invalid size")
        section_count = (lcb - 4) // 16
        if section_count > _MAX_SECTIONS:
            _raise("Legacy .doc PlcfSed exceeds the section limit")
        cp_values = tuple(
            _i32(table, fc + 4 * index) for index in range(section_count + 1)
        )
        if any(cp < 0 for cp in cp_values):
            _raise("Legacy .doc PlcfSed contains a negative CP")
        if any(left >= right for left, right in zip(cp_values, cp_values[1:])):
            _raise("Legacy .doc PlcfSed CP boundaries are not ordered")
        # The final aCP entry is a sentinel.  It can be beyond ccpText when
        # another story follows, so only non-final section boundaries become
        # marks here.
        self._marks = frozenset(
            cp_values[index] - 1
            for index in range(1, section_count)
            if cp_values[index] > 0
        )
        return self._marks

    def is_mark(self, cp: int) -> bool:
        return cp in self._load()


class _PapxIndex:
    """Lazy, bounded index over the document's PlcBtePapx pages."""

    def __init__(self, doc: BinaryDocument):
        self._doc = doc
        self._present = False
        self._fc_values: tuple[int, ...] = ()
        self._page_numbers: tuple[int, ...] = ()
        self._pages: dict[int, tuple[_FkpRun, ...]] = {}
        self._page_starts: dict[int, tuple[int, ...]] = {}
        self._cache: dict[tuple[tuple[int, int], int], _ParagraphProperties] = {}
        pair = doc.fib.pair(13)
        fc, lcb = pair
        if fc < 0 or lcb < 0:
            _raise("Legacy .doc PlcBtePapx location is invalid")
        if lcb == 0:
            return
        self._present = True
        table = doc.table
        if fc > len(table) or lcb > len(table) - fc:
            _raise("Legacy .doc PlcBtePapx is outside the Table stream")
        if lcb < 4 or (lcb - 4) % 8:
            _raise("Legacy .doc PlcBtePapx has an invalid size")
        page_count = (lcb - 4) // 8
        if page_count > _MAX_FKP_PAGES:
            _raise("Legacy .doc PlcBtePapx exceeds the page limit")
        fc_count = page_count + 1
        fc_values = tuple(_u32(table, fc + 4 * index) for index in range(fc_count))
        if any(left >= right for left, right in zip(fc_values, fc_values[1:])):
            _raise("Legacy .doc PlcBtePapx FC boundaries are not strictly ordered")
        pn_base = fc + 4 * fc_count
        page_numbers = tuple(
            _u32(table, pn_base + 4 * index) & 0x003F_FFFF
            for index in range(page_count)
        )
        # The BTE itself is shared by all stories, so its range structure is
        # validated eagerly.  FKP pages are selected by FC and loaded only on
        # demand: an unrelated header/textbox page may be corrupt without
        # invalidating the requested main-story range.
        self._fc_values = fc_values
        self._page_numbers = page_numbers

    def _load_page(self, page_number: int) -> tuple[_FkpRun, ...]:
        cached = self._pages.get(page_number)
        if cached is not None:
            return cached
        word = self._doc.word
        page_offset = page_number * _FKP_SIZE
        cb_mac = self._doc.fib.cb_mac
        page = _parse_fkp(word, page_offset, cb_mac)
        if len(page) > _MAX_FKP_ENTRIES:
            _raise("Legacy .doc PAPX entries exceed the processing limit")
        self._pages[page_number] = page
        self._page_starts[page_number] = tuple(run.fc_start for run in page)
        return page

    @property
    def present(self) -> bool:
        return self._present

    def _run_for_fc(self, fc: int) -> tuple[tuple[int, int], _FkpRun] | None:
        if not self._fc_values:
            return None
        page_index = bisect_right(self._fc_values, fc) - 1
        if page_index < 0 or page_index >= len(self._page_numbers):
            return None
        if fc >= self._fc_values[page_index + 1]:
            return None
        page_number = self._page_numbers[page_index]
        runs = self._load_page(page_number)
        starts = self._page_starts[page_number]
        run_index = bisect_right(starts, fc) - 1
        while run_index >= 0:
            run = runs[run_index]
            if run.fc_start <= fc < run.fc_end:
                if run.fc_start < self._fc_values[page_index] or run.fc_end > self._fc_values[page_index + 1]:
                    _raise("Legacy .doc PAPX FKP run exceeds its BTE range")
                return (page_index, run_index), run
            if run.fc_start < fc:
                break
            run_index -= 1
        return None

    def properties_for_fc(self, fc: int, prm: int = 0) -> _ParagraphProperties:
        found = self._run_for_fc(fc)
        if found is None:
            if self._present:
                _raise("Legacy .doc paragraph mark is outside PlcBtePapx")
            # A default PAPX still receives the PCD's PRM.  This matters for
            # non-complex paragraphs whose compact PRM carries PFInTable or a
            # table-row marker even though no direct PAPX record is stored.
            return _resolve_properties(self._doc, None, _prm_value(prm)).clone()
        index, run = found
        prm_int = _prm_value(prm)
        key = (index, prm_int)
        cached = self._cache.get(key)
        if cached is None:
            cached = _resolve_properties(self._doc, run.grpprl, prm_int)
            self._cache[key] = cached
        return cached.clone()


# ---------------------------------------------------------------------------
# CP/FC and piece helpers


def _read_text(doc: BinaryDocument, start: int, end: int) -> list[str]:
    reader = doc.read_text
    try:
        value = reader(start, end)
    except LegacyDocError:
        raise
    except Exception as exc:
        _raise(f"Unable to read legacy .doc text: {exc}")
    if not isinstance(value, str):
        _raise("BinaryDocument returned non-text CP data")
    count = end - start
    # Most readers preserve UTF-16 code-unit indexing by returning surrogate
    # code points for supplementary characters.  If it instead returns a
    # scalar emoji, expand it to one visible entry plus one empty continuation
    # entry so marker positions remain CP-aligned.
    expanded: list[str] = []
    index = 0
    while index < len(value):
        code = ord(value[index])
        if 0xD800 <= code <= 0xDBFF and index + 1 < len(value):
            low = ord(value[index + 1])
            if 0xDC00 <= low <= 0xDFFF:
                expanded.append(chr(0x10000 + ((code - 0xD800) << 10) + low - 0xDC00))
                expanded.append("")
                index += 2
                continue
        if code > 0xFFFF:
            expanded.append(value[index])
            expanded.append("")
        else:
            expanded.append(value[index])
        index += 1
    if len(expanded) == count:
        return expanded
    _raise("BinaryDocument CP text length does not match its CP range")
    return []  # unreachable


# ---------------------------------------------------------------------------
# Public iterator


def iter_paragraphs(doc: BinaryDocument, start: int, end: int) -> Iterator[Paragraph]:
    """Iterate paragraphs in CP range ``[start, end)``.

    The iterator reads all marker positions from the requested range before
    yielding records.  This is intentional: a terminating character's PAPX
    applies to the paragraph that ends there, and a paragraph may cross piece
    boundaries.  A final unterminated selection fragment is yielded as a
    normal paragraph, which lets callers request a sub-range without losing
    its text.
    """

    if not isinstance(start, int) or not isinstance(end, int) or start < 0 or end < start:
        _raise("Legacy .doc paragraph CP range is invalid")
    limit = doc.cp_limit
    if not isinstance(limit, int) or limit < 0:
        _raise("BinaryDocument has an invalid CP limit")
    if end > limit:
        _raise("Legacy .doc paragraph CP range exceeds the document")
    if start == end:
        return

    cp_chars = _read_text(doc, start, end)
    papx = _PapxIndex(doc)
    # PlcfSed is a main-story index.  A 0x0C in a textbox or another story is
    # a stored manual/page-break character and must remain inside its
    # paragraph.  Avoid even reading the shared section table unless the
    # selected main-story range contains such a candidate.
    section_index = None
    if end <= doc.fib.ccp_text and "\x0c" in cp_chars:
        section_index = _SectionIndex(doc)
    cursor = start
    for relative, char in enumerate(cp_chars):
        marker_cp = start + relative
        is_section_mark = (
            char == "\x0c" and section_index is not None and section_index.is_mark(marker_cp)
        )
        if char not in ("\r", "\x07") and not is_section_mark:
            continue
        marker_end = marker_cp + 1
        try:
            fc = doc.fc_for_cp(marker_cp)
            prm = doc.piece_at(marker_cp).prm
        except LegacyDocError:
            raise
        except Exception as exc:
            _raise(f"Unable to locate legacy .doc paragraph mark: {exc}")
        props = papx.properties_for_fc(fc, prm)
        text = "".join(cp_chars[cursor - start : relative])
        is_cell_mark = char == "\x07"
        # TTP is itself one of the paragraph terminators.  Real files usually
        # put it on a cell mark (0x07), while some writers use a CR/section
        # mark for an inner row.  The property is explicit structure, so it
        # takes precedence over the character value and must not manufacture
        # an empty cell before closing the row.
        row_end = props.ttp or props.inner_ttp
        cell_end = (is_cell_mark and props.in_table and not row_end) or (
            props.inner_cell and char in ("\r", "\x0c")
        )
        cells = tuple(props.cells) if row_end and props.cells is not None else None
        # Older Word writers set PFInTable but omit PItap for a first-level
        # table.  PItap's zero value is therefore ambiguous with an ordinary
        # body paragraph; explicit table membership resolves it to the
        # renderer's one-based frame depth.  The same applies to an explicit
        # TTP marker, which is structural evidence rather than a literal
        # 0x07 guess.
        in_table = bool(props.in_table or props.depth > 0 or props.ttp or props.inner_ttp)
        depth = props.depth
        if in_table and depth == 0:
            depth = 1
        yield Paragraph(
            start=cursor,
            end=marker_end,
            text=text,
            in_table=in_table,
            depth=depth,
            cell_end=cell_end,
            row_end=row_end,
            cells=cells,
        )
        cursor = marker_end

    if cursor < end:
        yield Paragraph(
            start=cursor,
            end=end,
            text="".join(cp_chars[cursor - start :]),
            in_table=False,
            depth=0,
            cell_end=False,
            row_end=False,
            cells=None,
        )


__all__ = [
    "CellFormat",
    "Paragraph",
    "iter_paragraphs",
    "parse_sprms",
]
