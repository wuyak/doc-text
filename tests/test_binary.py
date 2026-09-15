"""Focused tests for the FIB/CLX layer.

These tests use a tiny stream provider instead of the OLE fixture builder so
that malformed CLX and lazy Data reads can be exercised independently of the
other extraction stages.
"""

from __future__ import annotations

import struct

import pytest

from legacy_doc._binary import BinaryDocument
from legacy_doc.exceptions import LegacyDocError
from legacy_doc.types import ExtractionOptions


class MemoryOle:
    def __init__(self, streams: dict[str, bytes]) -> None:
        self.streams = streams
        self.options = ExtractionOptions()
        self.reads: list[str] = []

    def read_stream(self, name: str, max_size: int | None = None) -> bytes:
        self.reads.append(name)
        try:
            value = self.streams[name]
        except KeyError as exc:
            raise LegacyDocError(f"Required OLE stream '{name}' not found") from exc
        if max_size is not None and len(value) > max_size:
            raise LegacyDocError(f"OLE stream '{name}' exceeds parser limit")
        return value


def _document(
    pieces: list[tuple[int, bytes, bool]],
    *,
    stories: tuple[int, int, int, int, int, int] = (0, 0, 0, 0, 0, 0),
    ccp_text: int | None = None,
    prcs: tuple[bytes, ...] = (),
    prms: tuple[int, ...] | None = None,
    cb_mac: int | None = None,
    table_name: str = "0Table",
    extra_table_prefix: bytes = b"",
    data: bytes | None = None,
    n_fib: int = 0x00C1,
    n_fib_new: int | None = None,
    physical_fcs: tuple[int, ...] | None = None,
) -> tuple[MemoryOle, BinaryDocument]:
    """Build a valid FIB and a CLX with explicit physical pieces."""

    text_parts = []
    physical = 0x400
    pcds: list[bytes] = []
    cp_values = [0]
    for index, (cp_length, raw, compressed) in enumerate(pieces):
        width = 1 if compressed else 2
        if len(raw) != cp_length * width:
            raise AssertionError("piece bytes do not match CP length")
        fc = physical_fcs[index] if physical_fcs is not None else physical
        text_parts.append((fc, raw))
        encoded_fc = fc * 2 if compressed else fc
        if compressed:
            encoded_fc |= 0x40000000
        prm = (prms or ())[index] if prms is not None else 0
        pcds.append(struct.pack("<HIH", 0, encoded_fc, prm))
        cp_values.append(cp_values[-1] + cp_length)
        physical = max(physical, fc + len(raw))

    auxiliary_cp = sum(stories)
    if ccp_text is None:
        ccp_text = cp_values[-1] if not auxiliary_cp else pieces[0][0]
    expected_fib_cp = ccp_text + auxiliary_cp + (1 if auxiliary_cp else 0)
    if cp_values[-1] != expected_fib_cp:
        raise AssertionError(
            f"piece CP limit {cp_values[-1]} does not match FIB stories {expected_fib_cp}"
        )

    pair_counts = {0x00C1: 93, 0x00D9: 108, 0x0101: 136, 0x010C: 164, 0x0112: 183}
    effective_n_fib = n_fib_new or n_fib
    pair_count = pair_counts[effective_n_fib]
    csw_new = 0 if n_fib_new is None else (5 if n_fib_new == 0x0112 else 2)
    extension_size = csw_new * 2
    fib_end = 32 + 2 + 28 + 2 + 88 + 2 + pair_count * 8 + 2 + extension_size
    word = bytearray(max(physical, fib_end, 0x400))
    for offset, raw in text_parts:
        word[offset : offset + len(raw)] = raw
    struct.pack_into("<H", word, 0, 0xA5EC)
    struct.pack_into("<H", word, 2, n_fib)
    struct.pack_into("<H", word, 0x0A, 0x0200 if table_name == "1Table" else 0)
    struct.pack_into("<H", word, 32, 14)
    struct.pack_into("<H", word, 62, 22)
    struct.pack_into("<I", word, 64, max(physical, fib_end) if cb_mac is None else cb_mac)
    struct.pack_into("<i", word, 76, ccp_text)
    for offset, value in zip((80, 84, 92, 96, 100, 104), stories):
        struct.pack_into("<i", word, offset, value)
    struct.pack_into("<H", word, 152, pair_count)

    plc = struct.pack("<" + "I" * len(cp_values), *cp_values) + b"".join(pcds)
    clx = bytearray(extra_table_prefix)
    for grpprl in prcs:
        clx.extend(b"\x01" + struct.pack("<H", len(grpprl)) + grpprl)
    # fcClx/lcbClx covers the complete CLX, including every leading Prc.
    # The real FIB pointer is an offset in the selected Table stream, while
    # the test's optional prefix deliberately gives the CLX a non-zero start.
    fc_clx = 0
    clx.extend(b"\x02" + struct.pack("<I", len(plc)) + plc)
    struct.pack_into("<II", word, 0x1A2, fc_clx, len(clx) - fc_clx)
    csw_offset = 154 + pair_count * 8
    struct.pack_into("<H", word, csw_offset, csw_new)
    if n_fib_new is not None:
        struct.pack_into("<H", word, csw_offset + 2, n_fib_new)

    streams = {"WordDocument": bytes(word), table_name: bytes(clx)}
    if data is not None:
        streams["Data"] = data
    ole = MemoryOle(streams)
    return ole, BinaryDocument(ole, ExtractionOptions())


def test_fib_uses_selected_table_and_exposes_piece_mapping() -> None:
    ole, doc = _document([(3, "abc".encode("utf-16le"), False)], table_name="1Table")

    assert ole.reads == ["WordDocument", "1Table"]
    assert doc.fib.n_fib == 0x00C1
    assert doc.fib.cb_mac == len(doc.word)
    assert doc.fib.ccp_text == 3
    assert doc.fib.pair(33) == (0, len(doc.table))
    assert doc.cp_limit == 3
    assert doc.pieces[0].fc == 0x400
    assert doc.read_text(0, 3) == "abc"
    assert doc.fc_for_cp(2) == 0x404
    assert doc.piece_at(1) == doc.pieces[0]


def test_prm_and_prc_are_preserved() -> None:
    # A Prm1 value of 1 refers to the first Prc (fComplex in bit zero).
    _, doc = _document(
        [(1, "A".encode("utf-16le"), False)],
        prcs=(b"\x16\x24\x01",),
        prms=(1,),
    )
    assert doc.prcs == (b"\x16\x24\x01",)
    assert doc.pieces[0].prm == 1


def test_compressed_ms_doc_mapping_is_not_plain_cp1252() -> None:
    raw = bytes((0x80, 0x82, 0x85, 0x9F))
    _, doc = _document([(4, raw, True)])
    assert doc.read_text(0, 4) == "\x80\u201a\u2026\u0178"


def test_utf16_surrogate_pair_can_cross_piece_boundary() -> None:
    _, doc = _document([
        (2, "A\ud83d".encode("utf-16le", "surrogatepass"), False),
        (2, "\ude00Z".encode("utf-16le", "surrogatepass"), False),
    ])
    assert doc.read_text(0, 4) == "A😀Z"
    assert doc.read_text(1, 3) == "😀"


def test_selected_piece_bytes_are_checked_against_cbmac() -> None:
    ole, doc = _document([(3, "abc".encode("utf-16le"), False)], cb_mac=0x402)
    with pytest.raises(LegacyDocError, match="cbMac"):
        doc.read_text(0, 3)
    # The FIB/CLX can still be inspected before the selected range is read.
    assert ole.reads == ["WordDocument", "0Table"]


def test_bad_clx_does_not_scan_a_candidate_elsewhere() -> None:
    valid_ole, _ = _document([(1, b"A\x00", False)])
    table = bytearray(b"\x02\x04\x00\x00\x00" + valid_ole.streams["0Table"])
    # FIB points at the prefix, which is intentionally malformed.  A valid
    # CLX follows it, but the parser must fail at the specified location.
    word = bytearray(valid_ole.streams["WordDocument"])
    struct.pack_into("<II", word, 0x1A2, 0, 5)
    ole = MemoryOle({"WordDocument": bytes(word), "0Table": bytes(table)})
    with pytest.raises(LegacyDocError, match="CLX|Pcdt"):
        BinaryDocument(ole, ExtractionOptions())


def test_data_stream_is_loaded_lazily_and_bounded() -> None:
    ole, doc = _document([(1, b"A\x00", False)], data=b"xyz")
    assert ole.reads == ["WordDocument", "0Table"]
    assert doc.read_data(1, 2) == b"yz"
    assert ole.reads == ["WordDocument", "0Table", "Data"]
    with pytest.raises(LegacyDocError, match="Data range"):
        doc.read_data(2, 2)


@pytest.mark.parametrize("version", (0x00C1, 0x00D9, 0x0101, 0x010C, 0x0112))
def test_supported_effective_fib_versions(version: int) -> None:
    if version == 0x00C1:
        _, doc = _document([(1, b"A\x00", False)], n_fib=version)
    else:
        # A C1 FibBase with FibRgCswNew.nFibNew is the normal upgrade path;
        # the pair count must follow the effective version after the suffix.
        _, doc = _document(
            [(1, b"A\x00", False)], n_fib=0x00C1, n_fib_new=version
        )
    assert doc.fib.n_fib == version
    assert doc.fib.cb_rg_fc_lcb == {
        0x00C1: 93,
        0x00D9: 108,
        0x0101: 136,
        0x010C: 164,
        0x0112: 183,
    }[version]
    assert doc.read_text(0, 1) == "A"


def test_modern_base_fib_with_zero_cswnew_is_accepted() -> None:
    _, doc = _document([(1, b"A\x00", False)], n_fib=0x0101)
    assert doc.fib.n_fib == 0x0101
    assert doc.fib.csw_new == 0
    assert doc.read_text(0, 1) == "A"


def test_fib_reserved_offsets_do_not_bound_text() -> None:
    original, _ = _document([(1, b"A\x00", False)])
    word = bytearray(original.streams["WordDocument"])
    struct.pack_into("<II", word, 0x18, 0xDEADBEEF, 0xCAFEBABE)
    ole = MemoryOle({
        "WordDocument": bytes(word),
        "0Table": original.streams["0Table"],
    })
    doc = BinaryDocument(ole, ExtractionOptions())
    assert doc.read_text(0, 1) == "A"


def test_cp_order_is_independent_of_fc_order_and_selection_is_exact() -> None:
    _, doc = _document(
        [
            (2, b"A\x00B\x00", False),
            (2, b"CD", True),
        ],
        physical_fcs=(0x500, 0x400),
    )
    assert doc.read_text(0, 4) == "ABCD"
    assert doc.read_text(1, 3) == "BC"
    assert doc.fc_for_cp(0) == 0x500
    assert doc.fc_for_cp(2) == 0x400
    assert doc.fc_for_cp(3) == 0x401


def test_uncompressed_fc_is_an_absolute_byte_offset_without_alignment_guess() -> None:
    _, doc = _document(
        [(1, b"A\x00", False)],
        physical_fcs=(0x401,),
    )
    assert doc.fc_for_cp(0) == 0x401
    assert doc.read_text(0, 1) == "A"


def test_bad_selected_middle_piece_fails_before_returning_partial_text() -> None:
    _, doc = _document(
        [
            (1, b"A\x00", False),
            (1, b"B\x00", False),
            (1, b"C\x00", False),
        ],
        physical_fcs=(0x400, 0x3000, 0x402),
        cb_mac=0x1000,
    )
    with pytest.raises(LegacyDocError, match="cbMac|stream"):
        doc.read_text(0, 3)


def test_unselected_story_bytes_are_not_decoded() -> None:
    # The footnote piece contains an unpaired UTF-16 high surrogate.  Its
    # shared CP/FC record is valid, but it is outside the selected main story.
    _, doc = _document(
        [
            (1, b"A\x00", False),
            (1, b"\x00\xD8", False),
            (1, b"\r\x00", False),  # global auxiliary-story guard CP
        ],
        stories=(1, 0, 0, 0, 0, 0),
        ccp_text=1,
    )
    assert doc.read_text(0, 1) == "A"
    with pytest.raises(LegacyDocError, match="UTF-16"):
        doc.read_text(1, 2)


def test_invalid_utf16_does_not_use_replacement_characters() -> None:
    _, doc = _document([(1, b"\x00\xD8", False)])
    with pytest.raises(LegacyDocError, match="UTF-16"):
        doc.read_text(0, 1)
