"""Focused tests for the DOC paragraph property reader.

These tests use a deliberately small BinaryDocument-shaped fake.  The OLE and
FIB readers have their own tests; keeping the FKP/PRM cases here independent
makes failures point at paragraph decoding rather than at a container fixture.
"""

from __future__ import annotations

from dataclasses import dataclass
import struct

import pytest

from legacy_doc._paragraphs import (
    SPRM_PF_IN_TABLE,
    SPRM_PF_TTP,
    SPRM_P_HUGE_PAPX,
    SPRM_P_ITAP,
    SPRM_P_TABLE_PROPS,
    SPRM_T_DEF_TABLE,
    SPRM_T_MERGE,
    SPRM_T_VERT_MERGE,
    CellFormat,
    iter_paragraphs,
    parse_sprms,
)
from legacy_doc.exceptions import LegacyDocError


_FC_TEXT = 0x400


@dataclass(frozen=True)
class _Piece:
    cp_start: int
    cp_end: int
    fc: int
    compressed: bool = False
    prm: int = 0


@dataclass(frozen=True)
class _Fib:
    pairs: dict[int, tuple[int, int]]
    cb_mac: int
    ccp_text: int = 0

    def pair(self, index: int) -> tuple[int, int]:
        return self.pairs.get(index, (0, 0))


def _cp_len(text: str) -> int:
    return len(text.encode("utf-16le", errors="surrogatepass")) // 2


def _encode_sprm(opcode: int, value: int | bytes) -> bytes:
    """Encode the subset used by these tests, as a GrpPrl would store it."""

    spra = (opcode >> 13) & 0x07
    sizes = {0: 1, 1: 1, 2: 2, 3: 4, 4: 2, 5: 2, 7: 3}
    if isinstance(value, int):
        value = value.to_bytes(sizes[spra], "little", signed=value < 0)
    if spra == 6:
        if opcode == SPRM_T_DEF_TABLE:
            prefix = struct.pack("<H", len(value) + 1)
        else:
            prefix = bytes([len(value)])
        return struct.pack("<H", opcode) + prefix + value
    return struct.pack("<H", opcode) + value


def _tdef(flags: list[int]) -> bytes:
    count = len(flags)
    centers = struct.pack("<" + "h" * (count + 1), *range(0, 100 * (count + 1), 100))
    tc80 = b"".join(struct.pack("<H", flag) + b"\0" * 18 for flag in flags)
    return _encode_sprm(SPRM_T_DEF_TABLE, bytes([count]) + centers + tc80)


def _papx_record(grpprl: bytes) -> bytes:
    # PapxInFkp begins with the two-byte istd.  The two encodings exercise
    # both cb != 0 and the cb == 0/cb' form from the specification.
    payload = b"\0\0" + grpprl
    if len(payload) & 1:
        return bytes([(len(payload) + 1) // 2]) + payload
    return b"\0" + bytes([len(payload) // 2]) + payload


class _FakeDocument:
    """Enough of BinaryDocument for iter_paragraphs to run."""

    def __init__(
        self,
        text: str,
        properties: dict[int, bytes] | None = None,
        *,
        pieces: tuple[_Piece, ...] | None = None,
        piece_prm: int = 0,
        prcs: tuple[bytes, ...] = (),
        data: bytes = b"",
        include_papx: bool = True,
        bad_unselected_page: bool = False,
        section_break_cps: set[int] | None = None,
    ) -> None:
        self.text = text
        self.cp_limit = _cp_len(text)
        encoded = text.encode("utf-16le", errors="surrogatepass")
        self.word = bytearray(_FC_TEXT) + bytearray(encoded)
        self.prcs = prcs
        self._data = data
        self._section_breaks = section_break_cps or set()
        self.options = type("Options", (), {"max_file_bytes": 1 << 20})()

        if pieces is None:
            pieces = (_Piece(0, self.cp_limit, _FC_TEXT, prm=piece_prm),)
        self.pieces = pieces
        self._piece_starts = tuple(piece.cp_start for piece in pieces)

        self.table = b""
        if include_papx:
            self._append_papx_pages(properties or {}, bad_unselected_page=bad_unselected_page)
        pairs = {13: (0, len(self.table))}
        if self._section_breaks:
            section_cps = [0, *(cp + 1 for cp in sorted(self._section_breaks)), self.cp_limit]
            sed = b"\0\0" + struct.pack("<I", 0xFFFFFFFF) + b"\0" * 6
            section_table = struct.pack("<" + "I" * len(section_cps), *section_cps)
            section_table += sed * (len(section_cps) - 1)
            pairs[6] = (len(self.table), len(section_table))
            self.table += section_table
        self.fib = _Fib(pairs, len(self.word), self.cp_limit)

    def _append_papx_pages(self, properties: dict[int, bytes], *, bad_unselected_page: bool) -> None:
        runs: list[tuple[int, int, bytes]] = []
        cp = 0
        run_start = _FC_TEXT
        for char in self.text:
            width = _cp_len(char)
            cp += width
            if char in "\r\x07" or (char == "\x0c" and cp - 1 in self._section_breaks):
                run_end = _FC_TEXT + cp * 2
                runs.append((run_start, run_end, properties.get(cp - 1, b"")))
                run_start = run_end
        if run_start < _FC_TEXT + len(self.text.encode("utf-16le", errors="surrogatepass")):
            runs.append((run_start, _FC_TEXT + len(self.text.encode("utf-16le", errors="surrogatepass")), b""))
        if not runs:
            return

        page_specs: list[list[tuple[int, int, bytes]]] = [runs[: min(8, len(runs))]]
        if len(runs) > 8:
            page_specs.append(runs[8:])

        page_numbers: list[int] = []
        page_fc_ranges: list[tuple[int, int]] = []
        for page_index, batch in enumerate(page_specs):
            self.word.extend(b"\0" * (-len(self.word) % 512))
            page_number = len(self.word) // 512
            page_numbers.append(page_number)
            page = bytearray(512)
            count = len(batch)
            for index, (left, _right, _grpprl) in enumerate(batch):
                struct.pack_into("<I", page, 4 * index, left)
            struct.pack_into("<I", page, 4 * count, batch[-1][1])
            bx_end = 4 * (count + 1) + 13 * count
            cursor = bx_end + (bx_end & 1)
            for index, (_left, _right, grpprl) in enumerate(batch):
                if not grpprl:
                    continue
                record = _papx_record(grpprl)
                if cursor + len(record) > 511:
                    raise AssertionError("test PAPX page is full")
                page[4 * (count + 1) + 13 * index] = cursor // 2
                page[cursor : cursor + len(record)] = record
                cursor += len(record) + (len(record) & 1)
            if bad_unselected_page and page_index:
                page[511] = 0xFF
            else:
                page[511] = count
            self.word.extend(page)
            page_fc_ranges.append((batch[0][0], batch[-1][1]))

        bte = b"".join(struct.pack("<I", left) for left, _right in page_fc_ranges)
        bte += struct.pack("<I", page_fc_ranges[-1][1])
        bte += b"".join(struct.pack("<I", page_number) for page_number in page_numbers)
        self.table = bte

    def read_text(self, start: int, end: int) -> str:
        return self.text[start:end]

    def piece_at(self, cp: int) -> _Piece:
        for piece in self.pieces:
            if piece.cp_start <= cp < piece.cp_end:
                return piece
        raise LegacyDocError("CP outside fake piece table")

    def fc_for_cp(self, cp: int) -> int:
        piece = self.piece_at(cp)
        return piece.fc + (cp - piece.cp_start) * (1 if piece.compressed else 2)

    def read_data(self, offset: int, size: int) -> bytes:
        return self._data[offset : offset + size]


def test_variable_sprm_lengths_preserve_following_opcode() -> None:
    # One deleted and one added tab stop: the extended form's remainder is
    # 1 + 4 + 1 + 3 = 9 bytes, followed by a normal paragraph SPRM.
    body = bytes([1]) + struct.pack("<hh", 10, 20) + bytes([1]) + struct.pack("<hB", 30, 2)
    grpprl = struct.pack("<H", 0xC615) + bytes([0xFF]) + body
    grpprl += _encode_sprm(SPRM_PF_IN_TABLE, 1)
    parsed = parse_sprms(grpprl)
    assert [item.opcode for item in parsed] == [0xC615, SPRM_PF_IN_TABLE]
    assert parsed[0].operand == bytes([0xFF]) + body


def test_fkp_records_and_structural_terminators_are_cp_aligned() -> None:
    table_group = _encode_sprm(SPRM_PF_IN_TABLE, 1) + _encode_sprm(SPRM_P_ITAP, 1)
    row_group = table_group + _encode_sprm(SPRM_PF_TTP, 1) + _tdef([0, 0])
    doc = _FakeDocument(
        "A\rB\x07\x07C\x0cD\r", {1: b"", 3: table_group, 4: row_group}, section_break_cps={6}
    )
    records = list(iter_paragraphs(doc, 0, doc.cp_limit))
    assert [(p.start, p.end, p.text) for p in records] == [
        (0, 2, "A"), (2, 4, "B"), (4, 5, ""), (5, 7, "C"), (7, 9, "D"),
    ]
    assert records[1].in_table and records[1].depth == 1 and records[1].cell_end
    assert records[2].row_end and not records[2].cell_end
    assert records[2].cells == (CellFormat(), CellFormat())
    assert not records[3].in_table


def test_section_table_distinguishes_section_marks_from_page_breaks() -> None:
    manual = _FakeDocument("A \x0c B\r")
    manual_records = list(iter_paragraphs(manual, 0, manual.cp_limit))
    assert [(record.text, record.end) for record in manual_records] == [("A \x0c B", 6)]

    section = _FakeDocument("A \x0c B\r", section_break_cps={2})
    section_records = list(iter_paragraphs(section, 0, section.cp_limit))
    assert [(record.text, record.end) for record in section_records] == [
        ("A ", 3), (" B", 6),
    ]


def test_prm0_and_prm1_apply_after_papx_and_without_a_papx() -> None:
    # Compact PRM0: isprm 0x18 is sprmPFInTable, with value 1 in the high byte.
    prm0 = (1 << 8) | (0x18 << 1)
    direct_clear = _encode_sprm(SPRM_PF_IN_TABLE, 0)
    no_papx = _FakeDocument("A\x07", include_papx=False, piece_prm=prm0)
    only = list(iter_paragraphs(no_papx, 0, no_papx.cp_limit))
    assert only[0].in_table and only[0].cell_end

    # PRM1 index 0 is a regular grpprl, and must override direct PAPX order.
    prc = _encode_sprm(SPRM_PF_IN_TABLE, 1) + _encode_sprm(SPRM_P_ITAP, 1)
    prm1 = (0 << 1) | 1
    with_prm1 = _FakeDocument("A\x07", {1: direct_clear}, piece_prm=prm1, prcs=(prc,))
    result = list(iter_paragraphs(with_prm1, 0, with_prm1.cp_limit))
    assert result[0].in_table and result[0].depth == 1 and result[0].cell_end


def test_horizontal_merge_survives_a_vertical_merge_modifier() -> None:
    cell_group = (
        _encode_sprm(SPRM_PF_IN_TABLE, 1)
        + _encode_sprm(SPRM_P_ITAP, 1)
        + _tdef([0, 0, 0])
        + _encode_sprm(SPRM_T_MERGE, bytes([0, 2]))
        + _encode_sprm(SPRM_T_VERT_MERGE, bytes([1, 2]))
    )
    row_group = cell_group + _encode_sprm(SPRM_PF_TTP, 1)
    doc = _FakeDocument("A\x07B\x07\x07", {1: cell_group, 3: cell_group, 4: row_group})
    records = list(iter_paragraphs(doc, 0, doc.cp_limit))
    assert records[-1].row_end
    assert records[-1].cells == (
        CellFormat(horizontal_merge=2),
        CellFormat(horizontal_merge=1),
        CellFormat(),
    )


@pytest.mark.parametrize("operand", [bytes([3, 2]), bytes([1, 4])])
def test_unused_vertical_merge_state_does_not_skip_record_validation(operand: bytes) -> None:
    group = _tdef([0, 0, 0]) + _encode_sprm(SPRM_T_VERT_MERGE, operand)
    doc = _FakeDocument("A\r", {1: group})
    with pytest.raises(LegacyDocError, match="TVertMerge"):
        list(iter_paragraphs(doc, 0, doc.cp_limit))


@pytest.mark.parametrize("indirect_opcode", [SPRM_P_HUGE_PAPX, SPRM_P_TABLE_PROPS])
def test_indirect_paragraph_properties_are_bounded_and_applied(indirect_opcode: int) -> None:
    indirect_group = _encode_sprm(SPRM_PF_IN_TABLE, 1) + _encode_sprm(SPRM_P_ITAP, 1) + _encode_sprm(SPRM_PF_TTP, 1)
    offset = 4
    data = b"\0" * offset + struct.pack("<H", len(indirect_group)) + indirect_group
    direct = _encode_sprm(indirect_opcode, offset) + _encode_sprm(SPRM_PF_IN_TABLE, 0)
    doc = _FakeDocument("A\x07", {1: direct}, data=data)
    record = list(iter_paragraphs(doc, 0, doc.cp_limit))[0]
    assert record.in_table and record.depth == 1 and record.row_end


def test_indirect_cycle_fails_instead_of_recursing() -> None:
    offset = 4
    indirect = _encode_sprm(SPRM_P_HUGE_PAPX, offset)
    # The header is large enough for the bounded PrcData rule and points back
    # to itself, so the failure is specifically cycle detection.
    group = indirect + _encode_sprm(SPRM_PF_IN_TABLE, 1) + _encode_sprm(SPRM_PF_TTP, 1)
    data = b"\0" * offset + struct.pack("<H", len(group)) + group
    doc = _FakeDocument("A\x07", {1: _encode_sprm(SPRM_P_HUGE_PAPX, offset)}, data=data)
    with pytest.raises(LegacyDocError, match="cycle|indirect"):
        list(iter_paragraphs(doc, 0, doc.cp_limit))


def test_unselected_fkp_page_is_not_loaded() -> None:
    text = "A\r" + "B\r" * 8
    props = {1: _encode_sprm(SPRM_PF_IN_TABLE, 1)}
    doc = _FakeDocument(text, props, bad_unselected_page=True)
    first = list(iter_paragraphs(doc, 0, 2))
    assert len(first) == 1 and first[0].text == "A"
