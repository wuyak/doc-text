"""Strict low-level readers for the binary Word document streams.

The public extraction code in :mod:`legacy_doc.word` deliberately has a small
surface.  This module keeps the format-specific bookkeeping in one place so
that callers which need to inspect paragraph or table properties can use the
same CP to FC mapping as the text reader.

Only the part of the MS-DOC FIB needed to locate document content is decoded
here.  The FIB and the CLX are indexes, so their structure is checked eagerly;
references to WordDocument bytes are checked when the corresponding CP range
is requested.  That distinction matters for a document containing a broken
unselected story.
"""

from __future__ import annotations

from bisect import bisect_right
from dataclasses import dataclass
import struct

from legacy_doc.exceptions import LegacyDocError
from legacy_doc.ole import OleReader
from legacy_doc.types import ExtractionOptions


# MS-DOC names the following five FIB layouts.  The value is the number of
# 64-bit fc/lcb pairs in FibRgFcLcb (the pair count is called cbRgFcLcb in the
# file, despite its slightly confusing name).
FIB_PAIR_COUNTS: dict[int, int] = {
    0x00C1: 0x005D,  # Word 97
    0x00D9: 0x006C,  # Word 2000
    0x0101: 0x0088,  # Word 2002
    0x010C: 0x00A4,  # Word 2003
    0x0112: 0x00B7,  # Word 2007-compatible FIB
}

# cswNew is a count of 16-bit values.  Its first value is nFibNew; the
# remaining values are version-specific extension data.  The values below
# are the minimum known prefix sizes; a longer declared section is bounded by
# the WordDocument stream and its unknown tail is skipped.
FIB_CSW_NEW_COUNTS: dict[int, int] = {
    0x00C1: 0,
    0x00D9: 2,
    0x0101: 2,
    0x010C: 2,
    0x0112: 5,
}

FIB_RGFCLCB97_PAIR_COUNT = FIB_PAIR_COUNTS[0x00C1]
PAIR_FC_CLX = 33
PAIR_FC_PLCF_BTE_CHPX = 12
PAIR_FC_PLCF_BTE_PAPX = 13
PAIR_FC_PLCF_SPA_MOM = 40
PAIR_FC_PLCFTXBX_TXT = 56
PAIR_FC_PLCF_HDRTXBX_TXT = 58

# A PlcPcd's data area consists of n Pcd records and n + 1 CP values.  This
# upper bound keeps a corrupt but otherwise file-sized CLX from causing a
# disproportionate number of Python objects to be allocated.
MAX_PIECES = 1_000_000
MAX_PRCS = 1 << 15
MAX_GRPPRL_BYTES = 0x3FA2


@dataclass(frozen=True)
class Piece:
    """One CP range and its corresponding WordDocument byte range.

    ``fc`` is already converted to the actual byte offset.  For compressed
    pieces this means the encoded FcCompressed value has been divided by two.
    ``prm`` is kept as the raw 16-bit Prm value for the property readers.
    """

    cp_start: int
    cp_end: int
    fc: int
    compressed: bool
    prm: int
    f_no_para_last: bool = False

    @property
    def cp_length(self) -> int:
        return self.cp_end - self.cp_start

    @property
    def byte_width(self) -> int:
        return 1 if self.compressed else 2


@dataclass(frozen=True)
class FibInfo:
    """The validated FIB fields used by :class:`BinaryDocument`."""

    n_fib: int
    cb_mac: int
    ccp_text: int
    ccp_ftn: int
    ccp_hdd: int
    ccp_atn: int
    ccp_edn: int
    ccp_txbx: int
    ccp_hdr_txbx: int
    pairs: tuple[tuple[int, int], ...]
    table_stream_name: str
    flags: int = 0
    n_fib_base: int | None = None
    csw: int = 0
    cslw: int = 0
    cb_rg_fc_lcb: int = 0
    csw_new: int = 0
    n_fib_new: int | None = None
    fib_end: int = 0

    def pair(self, index: int) -> tuple[int, int]:
        """Return the fc/lcb pair at ``index``.

        The index is the pair index within FibRgFcLcb, rather than a byte
        offset.  Invalid indexes are format errors, not absent optional data:
        callers that need an optional record should first inspect its lcb.
        """

        if isinstance(index, bool) or not isinstance(index, int):
            raise LegacyDocError("FIB pair index must be an integer")
        if index < 0 or index >= len(self.pairs):
            raise LegacyDocError(f"FIB pair index {index} is outside the FIB")
        return self.pairs[index]

    @property
    def expected_cp_limit(self) -> int:
        """Return the global CP limit prescribed by FibRgLw97 story counts."""

        auxiliary = (
            self.ccp_ftn
            + self.ccp_hdd
            + self.ccp_atn
            + self.ccp_edn
            + self.ccp_txbx
            + self.ccp_hdr_txbx
        )
        total = self.ccp_text + auxiliary
        # MS-DOC reserves one guard CP in the global PlcPcd when any
        # specialized story exists.  It is part of the piece table's final
        # CP range even though it is not part of an individual ccp value.
        return total + 1 if auxiliary else total


class BinaryDocument:
    """Read the FIB, selected Table stream, and its CLX piece table.

    ``ole`` is expected to be an :class:`~legacy_doc.ole.OleReader`, but the
    implementation only uses its stream-reading methods.  ``options`` is
    retained as a public attribute because higher-level binary readers share
    the same resource budget.
    """

    def __init__(
        self,
        ole: OleReader,
        options: ExtractionOptions | None = None,
    ) -> None:
        if options is None:
            options = ole.options
        self.options = options
        self.ole = ole

        self.word = _read_stream(ole, "WordDocument", options)
        if not isinstance(self.word, bytes):
            self.word = bytes(self.word)

        self.fib = _parse_fib(self.word)
        # The selected stream is determined solely by fWhichTblStm.  In
        # particular, do not search the other Table stream for a candidate.
        self.table = _read_stream(ole, self.fib.table_stream_name, options)
        if not isinstance(self.table, bytes):
            self.table = bytes(self.table)

        fc_clx, lcb_clx = self.fib.pair(PAIR_FC_CLX)
        clx = _bounded_slice(
            self.table,
            fc_clx,
            lcb_clx,
            what="FIB-specified CLX",
            allow_empty=False,
        )
        self.pieces, self.prcs = _parse_clx(clx)
        self._piece_starts = tuple(piece.cp_start for piece in self.pieces)
        self.cp_limit = self.pieces[-1].cp_end if self.pieces else 0
        if self.cp_limit != self.fib.expected_cp_limit:
            raise LegacyDocError(
                "CLX CP limit does not match FibRgLw97 story counts "
                f"({self.cp_limit} != {self.fib.expected_cp_limit})"
            )

        # max_text_bytes is the final rendered-text budget.  A low value is
        # allowed to coexist with a valid story whose terminating CR (or
        # other structural controls) is consumed by the renderer.  Binary
        # decoding therefore uses an independent, file-sized work budget;
        # TextRenderer applies the user-visible UTF-8 limit after structure is
        # interpreted.
        self._binary_work_limit = max(
            1,
            max(options.max_file_bytes, 1) * 4,
        )

        # Data is deliberately not read in __init__.  Indirect PAPX records
        # are selected by the paragraph reader, and a corrupt unused Data
        # stream must not make a valid main story fail.
        self._data: bytes | None = None

    def read_text(self, start: int, end: int) -> str:
        """Read text in the strict CP half-open interval ``[start, end)``.

        Every selected byte reference must fit within both the physical stream
        and ``cbMac``.  UTF-16 fragments that meet at a piece boundary are
        decoded together, preserving valid surrogate pairs.
        """

        _validate_cp_range(start, end, self.cp_limit)
        if start == end:
            return ""

        parts: list[str] = []
        utf8_size = 0
        pending_utf16 = bytearray()
        pending_raw_size = 0

        def append_text(fragment: str) -> None:
            nonlocal utf8_size
            try:
                fragment_size = len(fragment.encode("utf-8"))
            except UnicodeEncodeError as exc:  # defensive; strict decode
                raise LegacyDocError("Word text contains an invalid Unicode value") from exc
            utf8_size += fragment_size
            if utf8_size > self._binary_work_limit:
                raise LegacyDocError("Extracted Word text exceeds parser work limit")
            parts.append(fragment)

        def flush_utf16() -> None:
            nonlocal pending_raw_size
            if not pending_utf16:
                return
            try:
                decoded = bytes(pending_utf16).decode("utf-16le", errors="strict")
            except UnicodeDecodeError as exc:
                raise LegacyDocError("WordDocument contains invalid UTF-16 text") from exc
            append_text(decoded)
            pending_utf16.clear()
            pending_raw_size = 0

        first_piece = max(0, bisect_right(self._piece_starts, start) - 1)
        for piece_index in range(first_piece, len(self.pieces)):
            piece = self.pieces[piece_index]
            if piece.cp_end <= start:
                continue
            if piece.cp_start >= end:
                break

            selected_start = max(start, piece.cp_start)
            selected_end = min(end, piece.cp_end)
            if selected_start >= selected_end:
                continue

            raw = self._read_piece_bytes(piece, selected_start, selected_end)
            if piece.compressed:
                # A compressed piece cannot carry a UTF-16 surrogate unit, so
                # a pending uncompressed fragment can be decoded before it.
                flush_utf16()
                append_text(_decode_compressed(raw))
                continue

            if len(raw) % 2:
                raise LegacyDocError("Uncompressed Word text has an odd byte length")
            pending_utf16.extend(raw)
            pending_raw_size += len(raw)
            # A valid UTF-16 string is at most 3 bytes per code point in UTF-8
            # for BMP text and 4 for a supplementary code point.  Keeping a
            # small raw-input bound prevents a huge malformed selection from
            # bypassing max_text_bytes while it waits for a piece boundary.
            if pending_raw_size > self._binary_work_limit + 2:
                raise LegacyDocError("Extracted Word text exceeds parser input limit")

        flush_utf16()
        return "".join(parts)

    def fc_for_cp(self, cp: int) -> int:
        """Return the physical WordDocument byte address for one CP."""

        piece = self.piece_at(cp)
        delta = cp - piece.cp_start
        offset = piece.fc + delta * piece.byte_width
        self._validate_word_range(offset, piece.byte_width)
        return offset

    def piece_at(self, cp: int) -> Piece:
        """Return the piece containing ``cp`` (CP intervals are half-open)."""

        if isinstance(cp, bool) or not isinstance(cp, int):
            raise LegacyDocError("CP must be an integer")
        if cp < 0 or cp >= self.cp_limit:
            raise LegacyDocError(f"CP {cp} is outside the valid range")
        index = bisect_right(self._piece_starts, cp) - 1
        if index < 0 or index >= len(self.pieces):
            raise LegacyDocError(f"No CLX piece contains CP {cp}")
        piece = self.pieces[index]
        if cp >= piece.cp_end:
            raise LegacyDocError(f"No CLX piece contains CP {cp}")
        return piece

    def read_data(self, offset: int, size: int) -> bytes:
        """Read a bounded slice from the Data stream on first use."""

        _validate_nonnegative_int(offset, "Data offset")
        _validate_nonnegative_int(size, "Data size")
        if size == 0:
            return b""
        end = _checked_add(offset, size, "Data range")
        if self._data is None:
            self._data = _read_stream(self.ole, "Data", self.options)
            if not isinstance(self._data, bytes):
                self._data = bytes(self._data)
        if end > len(self._data):
            raise LegacyDocError(
                f"Data range [{offset}, {end}) exceeds Data stream ({len(self._data)} bytes)"
            )
        return self._data[offset:end]

    def _read_piece_bytes(self, piece: Piece, start: int, end: int) -> bytes:
        delta_start = start - piece.cp_start
        count = end - start
        width = piece.byte_width
        offset = piece.fc + delta_start * width
        size = count * width
        self._validate_word_range(offset, size)
        return self.word[offset : offset + size]

    def _validate_word_range(self, offset: int, size: int) -> None:
        _validate_nonnegative_int(offset, "WordDocument offset")
        _validate_nonnegative_int(size, "WordDocument size")
        end = _checked_add(offset, size, "WordDocument range")
        if end > self.fib.cb_mac:
            raise LegacyDocError(
                "WordDocument text reference exceeds cbMac "
                f"({end} > {self.fib.cb_mac})"
            )
        if end > len(self.word):
            raise LegacyDocError(
                "WordDocument text reference exceeds stream "
                f"({end} > {len(self.word)})"
            )


def _parse_fib(word: bytes) -> FibInfo:
    if len(word) < 32:
        raise LegacyDocError("WordDocument stream is too small for a FIB")

    ident = _u16(word, 0, "FIB wIdent")
    if ident == 0xA5DC:
        raise LegacyDocError("Unsupported Word 6/95 FIB (wIdent 0xA5DC)")
    if ident != 0xA5EC:
        raise LegacyDocError("Invalid Word FIB wIdent (expected 0xA5EC)")

    n_fib_base = _u16(word, 2, "FIB nFib")
    if n_fib_base not in FIB_PAIR_COUNTS:
        raise LegacyDocError(f"Unsupported Word FIB nFib 0x{n_fib_base:04X}")

    flags = _u16(word, 0x0A, "FIB flags")
    if flags & 0x0100:
        raise LegacyDocError("Encrypted legacy .doc files are not supported")
    table_stream_name = "1Table" if flags & 0x0200 else "0Table"

    # FibBase (32 bytes), then the counted sections.  The supported modern
    # versions require the canonical counts; accepting larger unknown arrays
    # here would make all subsequent field indexes ambiguous.
    offset = 32
    csw = _u16(word, offset, "FIB csw")
    offset += 2
    if csw != 0x000E:
        raise LegacyDocError(f"Unsupported Word FIB csw {csw}; expected 14")
    _require_range(word, offset, csw * 2, "FibRgW")
    offset += csw * 2

    cslw = _u16(word, offset, "FIB cslw")
    offset += 2
    if cslw != 0x0016:
        raise LegacyDocError(f"Unsupported Word FIB cslw {cslw}; expected 22")
    rglw_start = offset
    _require_range(word, rglw_start, cslw * 4, "FibRgLw")
    lw = [_i32(word, rglw_start + index * 4, "FibRgLw field") for index in range(cslw)]
    for name, index in (
        ("ccpText", 3),
        ("ccpFtn", 4),
        ("ccpHdd", 5),
        ("ccpAtn", 7),
        ("ccpEdn", 8),
        ("ccpTxbx", 9),
        ("ccpHdrTxbx", 10),
    ):
        if lw[index] < 0:
            raise LegacyDocError(f"FIB {name} is negative")
    cb_mac = _u32(word, rglw_start, "FibRgLw.cbMac")
    offset = rglw_start + cslw * 4

    cb_rg_fc_lcb = _u16(word, offset, "FIB cbRgFcLcb")
    offset += 2
    # nFibNew, which determines the effective layout, appears after this
    # variable-length pair table.  Read the declared count only after
    # applying a hard supported-version bound; validate it against the
    # effective nFib once FibRgCswNew has been read below.
    supported_pair_counts = frozenset(FIB_PAIR_COUNTS.values())
    if cb_rg_fc_lcb not in supported_pair_counts:
        raise LegacyDocError(
            "Invalid FIB cbRgFcLcb "
            f"{cb_rg_fc_lcb}; supported counts are "
            f"{sorted(supported_pair_counts)}"
        )
    pairs_size = _checked_mul(cb_rg_fc_lcb, 8, "FIB pair table")
    _require_range(word, offset, pairs_size, "FibRgFcLcb")
    pairs = tuple(
        (
            _u32(word, offset + index * 8, "FIB fc"),
            _u32(word, offset + index * 8 + 4, "FIB lcb"),
        )
        for index in range(cb_rg_fc_lcb)
    )
    offset += pairs_size

    csw_new = _u16(word, offset, "FIB cswNew")
    offset += 2
    # Real DOC files can carry a larger known pair table while leaving
    # cswNew at zero.
    # Zero means that FibBase.nFib remains authoritative.  For a nonzero
    # count, the first extension value supplies nFibNew and therefore the
    # effective pair-table layout; this is why cbRgFcLcb cannot be checked
    # against the base nFib above.  MS-DOC 2.5.15 permits a declared extension
    # to be larger than the in-memory structure, so only the known prefix is
    # validated and the remaining words are skipped.
    new_data_size = _checked_mul(csw_new, 2, "FibRgCswNew")
    _require_range(word, offset, new_data_size, "FibRgCswNew")
    n_fib_new: int | None = None
    n_fib = n_fib_base
    if csw_new:
        n_fib_new = _u16(word, offset, "FibRgCswNew.nFibNew")
        if n_fib_new not in FIB_PAIR_COUNTS or n_fib_new == 0x00C1:
            raise LegacyDocError(f"Unsupported Word FIB nFibNew 0x{n_fib_new:04X}")
        n_fib = n_fib_new
        expected_csw_new = FIB_CSW_NEW_COUNTS[n_fib_new]
        if csw_new < expected_csw_new:
            raise LegacyDocError(
                "Invalid FIB cswNew "
                f"{csw_new}; expected at least {expected_csw_new} for nFibNew "
                f"0x{n_fib_new:04X}"
            )

    expected_pairs = FIB_PAIR_COUNTS[n_fib]
    if cb_rg_fc_lcb < expected_pairs:
        version_label = (
            f"nFibNew 0x{n_fib:04X}" if n_fib_new is not None
            else f"nFib 0x{n_fib_base:04X}"
        )
        raise LegacyDocError(
            "Invalid FIB cbRgFcLcb "
            f"{cb_rg_fc_lcb}; expected {expected_pairs} for {version_label}"
        )

    fib_end = offset + new_data_size
    if cb_mac > len(word):
        raise LegacyDocError(
            f"FIB cbMac {cb_mac} exceeds WordDocument stream ({len(word)} bytes)"
        )
    if cb_mac < fib_end:
        raise LegacyDocError(
            f"FIB cbMac {cb_mac} ends before the FIB ({fib_end} bytes)"
        )

    ccp_text, ccp_ftn, ccp_hdd = lw[3], lw[4], lw[5]
    ccp_atn, ccp_edn, ccp_txbx, ccp_hdr_txbx = lw[7], lw[8], lw[9], lw[10]
    total_cp = ccp_text + ccp_ftn + ccp_hdd + ccp_atn + ccp_edn + ccp_txbx + ccp_hdr_txbx
    if total_cp > 0x7FFFFFFF:
        raise LegacyDocError("FIB story CP total exceeds the supported CP range")
    if ccp_ftn or ccp_hdd or ccp_atn or ccp_edn or ccp_txbx or ccp_hdr_txbx:
        total_cp += 1
    if total_cp > 0x7FFFFFFF:
        raise LegacyDocError("FIB global CP limit exceeds the supported CP range")

    return FibInfo(
        n_fib=n_fib,
        cb_mac=cb_mac,
        ccp_text=ccp_text,
        ccp_ftn=ccp_ftn,
        ccp_hdd=ccp_hdd,
        ccp_atn=ccp_atn,
        ccp_edn=ccp_edn,
        ccp_txbx=ccp_txbx,
        ccp_hdr_txbx=ccp_hdr_txbx,
        pairs=pairs,
        table_stream_name=table_stream_name,
        flags=flags,
        n_fib_base=n_fib_base,
        csw=csw,
        cslw=cslw,
        cb_rg_fc_lcb=cb_rg_fc_lcb,
        csw_new=csw_new,
        n_fib_new=n_fib_new,
        fib_end=fib_end,
    )


def _parse_clx(clx: bytes) -> tuple[tuple[Piece, ...], tuple[bytes, ...]]:
    if not clx:
        raise LegacyDocError("Invalid Word CLX: empty CLX")

    position = 0
    prcs: list[bytes] = []
    while position < len(clx):
        tag = clx[position]
        if tag == 0x01:  # Prc
            if len(prcs) >= MAX_PRCS:
                raise LegacyDocError("Word CLX contains too many Prc records")
            if position + 3 > len(clx):
                raise LegacyDocError("Invalid Word CLX: truncated Prc header")
            cb_grpprl = _u16(clx, position + 1, "Prc.cbGrpprl")
            if cb_grpprl > MAX_GRPPRL_BYTES:
                raise LegacyDocError("Invalid Word CLX: Prc grpprl is too large")
            end = _checked_add(position + 3, cb_grpprl, "Prc range")
            if end > len(clx):
                raise LegacyDocError("Invalid Word CLX: truncated Prc grpprl")
            # Preserve exactly the GrpPrl bytes.  SPRM interpretation belongs
            # to the paragraph/character property reader and has special
            # variable-length exceptions that cannot be inferred here.
            prcs.append(clx[position + 3 : end])
            position = end
            continue

        if tag != 0x02:
            raise LegacyDocError(
                f"Invalid Word CLX: unexpected record type 0x{tag:02X}"
            )

        if position + 5 > len(clx):
            raise LegacyDocError("Invalid Word CLX: truncated Pcdt header")
        lcb_plc = _u32(clx, position + 1, "Pcdt.lcb")
        if lcb_plc < 4 or lcb_plc > len(clx) - (position + 5):
            raise LegacyDocError("Invalid Word CLX: Pcdt length is outside CLX")
        if (lcb_plc - 4) % 12:
            raise LegacyDocError("Invalid Word CLX: Pcdt is not a whole PlcPcd")

        piece_count = (lcb_plc - 4) // 12
        if piece_count > MAX_PIECES:
            raise LegacyDocError("Word CLX contains too many pieces")
        plc_start = position + 5
        plc_end = plc_start + lcb_plc
        if plc_end != len(clx):
            # Pcdt is final in a Clx.  Silently accepting trailing bytes would
            # reintroduce the old "find a candidate" behaviour.
            raise LegacyDocError("Invalid Word CLX: bytes follow the Pcdt")

        cp_count = piece_count + 1
        cps = tuple(
            _u32(clx, plc_start + index * 4, "PlcPcd CP") for index in range(cp_count)
        )
        if not cps or cps[0] != 0:
            raise LegacyDocError("Invalid Word CLX: PlcPcd must start at CP zero")
        for left, right in zip(cps, cps[1:]):
            if right <= left:
                raise LegacyDocError("Invalid Word CLX: CP values are not increasing")

        pcd_start = plc_start + cp_count * 4
        pieces: list[Piece] = []
        for index in range(piece_count):
            pcd_offset = pcd_start + index * 8
            flags = _u16(clx, pcd_offset, "Pcd flags")
            # Pcd.fDirty is bit 2.  fR1 and fR2 are undefined and must not be
            # used to infer ranges, but fDirty has a normative zero value.
            if flags & 0x0004:
                raise LegacyDocError("Invalid Word CLX: Pcd.fDirty is set")
            fc_encoded = _u32(clx, pcd_offset + 2, "Pcd.fc")
            if fc_encoded & 0x80000000:
                raise LegacyDocError("Invalid Word CLX: FcCompressed.r1 is set")
            compressed = bool(fc_encoded & 0x40000000)
            encoded_fc = fc_encoded & 0x3FFFFFFF
            if compressed:
                if encoded_fc & 1:
                    raise LegacyDocError(
                        "Invalid Word CLX: compressed FC is not divisible by two"
                    )
                fc = encoded_fc // 2
            else:
                fc = encoded_fc
            prm = _u16(clx, pcd_offset + 6, "Pcd.prm")
            if prm & 1:
                prc_index = prm >> 1
                if prc_index >= len(prcs):
                    raise LegacyDocError(
                        "Invalid Word CLX: Prm1 references a missing Prc"
                    )
            pieces.append(
                Piece(
                    cp_start=cps[index],
                    cp_end=cps[index + 1],
                    fc=fc,
                    compressed=compressed,
                    prm=prm,
                    f_no_para_last=bool(flags & 0x0001),
                )
            )

        return tuple(pieces), tuple(prcs)

    # The loop can only leave when no record was found, which is malformed.
    raise LegacyDocError("Invalid Word CLX: missing Pcdt")


def _decode_compressed(raw: bytes) -> str:
    # FcCompressed uses an ANSI byte for all values not listed in its special
    # table.  Those ordinary values are direct Unicode code points; this is
    # intentionally not Python's cp1252 codec, which rejects five undefined
    # C1 slots and applies additional mappings (notably 0x80).
    return "".join(chr(_COMPRESSED_MAP.get(byte, byte)) for byte in raw)


_COMPRESSED_MAP = {
    0x82: 0x201A,
    0x83: 0x0192,
    0x84: 0x201E,
    0x85: 0x2026,
    0x86: 0x2020,
    0x87: 0x2021,
    0x88: 0x02C6,
    0x89: 0x2030,
    0x8A: 0x0160,
    0x8B: 0x2039,
    0x8C: 0x0152,
    0x91: 0x2018,
    0x92: 0x2019,
    0x93: 0x201C,
    0x94: 0x201D,
    0x95: 0x2022,
    0x96: 0x2013,
    0x97: 0x2014,
    0x98: 0x02DC,
    0x99: 0x2122,
    0x9A: 0x0161,
    0x9B: 0x203A,
    0x9C: 0x0153,
    0x9F: 0x0178,
}


def _read_stream(ole: OleReader, name: str, options: ExtractionOptions) -> bytes:
    try:
        value = ole.read_stream(name, max_size=options.max_file_bytes)
    except LegacyDocError:
        raise
    except (IndexError, KeyError, ValueError) as exc:
        raise LegacyDocError(f"Could not read OLE stream '{name}'") from exc
    if value is None:
        raise LegacyDocError(f"Required OLE stream '{name}' is empty or unavailable")
    try:
        return bytes(value)
    except (TypeError, ValueError) as exc:
        raise LegacyDocError(f"OLE stream '{name}' did not return bytes") from exc


def _bounded_slice(
    data: bytes,
    offset: int,
    size: int,
    *,
    what: str,
    allow_empty: bool,
) -> bytes:
    _validate_nonnegative_int(offset, f"{what} offset")
    _validate_nonnegative_int(size, f"{what} size")
    if not allow_empty and size == 0:
        raise LegacyDocError(f"{what} has zero length")
    end = _checked_add(offset, size, what)
    if end > len(data):
        raise LegacyDocError(
            f"{what} range [{offset}, {end}) exceeds stream ({len(data)} bytes)"
        )
    return data[offset:end]


def _validate_cp_range(start: int, end: int, limit: int) -> None:
    _validate_nonnegative_int(start, "CP start")
    _validate_nonnegative_int(end, "CP end")
    if start > end:
        raise LegacyDocError("CP range start is greater than end")
    if end > limit:
        raise LegacyDocError(f"CP range [{start}, {end}) exceeds CP limit {limit}")


def _validate_nonnegative_int(value: int, label: str) -> None:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise LegacyDocError(f"{label} must be a non-negative integer")


def _checked_add(left: int, right: int, label: str) -> int:
    if left < 0 or right < 0:
        raise LegacyDocError(f"{label} contains a negative value")
    result = left + right
    # Stream offsets are represented by uint32 in the FIB.  Refusing values
    # beyond that range also prevents architecture-dependent integer surprises
    # in downstream slicing code.
    if result > 0xFFFFFFFF:
        raise LegacyDocError(f"{label} overflows a 32-bit stream offset")
    return result


def _checked_mul(left: int, right: int, label: str) -> int:
    if left < 0 or right < 0:
        raise LegacyDocError(f"{label} contains a negative value")
    result = left * right
    if result > 0xFFFFFFFF:
        raise LegacyDocError(f"{label} overflows a 32-bit size")
    return result


def _require_range(data: bytes, offset: int, size: int, label: str) -> None:
    _validate_nonnegative_int(offset, f"{label} offset")
    _validate_nonnegative_int(size, f"{label} size")
    end = _checked_add(offset, size, label)
    if end > len(data):
        raise LegacyDocError(
            f"Truncated {label}: range [{offset}, {end}) exceeds WordDocument stream"
        )


def _u16(data: bytes, offset: int, label: str) -> int:
    _require_range(data, offset, 2, label)
    return struct.unpack_from("<H", data, offset)[0]


def _u32(data: bytes, offset: int, label: str) -> int:
    _require_range(data, offset, 4, label)
    return struct.unpack_from("<I", data, offset)[0]


def _i32(data: bytes, offset: int, label: str) -> int:
    _require_range(data, offset, 4, label)
    return struct.unpack_from("<i", data, offset)[0]


__all__ = [
    "BinaryDocument",
    "FibInfo",
    "Piece",
    "FIB_PAIR_COUNTS",
    "FIB_CSW_NEW_COUNTS",
    "PAIR_FC_CLX",
    "PAIR_FC_PLCF_BTE_CHPX",
    "PAIR_FC_PLCF_BTE_PAPX",
    "PAIR_FC_PLCF_SPA_MOM",
    "PAIR_FC_PLCFTXBX_TXT",
    "PAIR_FC_PLCF_HDRTXBX_TXT",
]
