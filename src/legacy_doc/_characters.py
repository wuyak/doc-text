"""Read the character property needed to distinguish symbols from literal text."""
from __future__ import annotations

from bisect import bisect_right
from dataclasses import dataclass
import struct
from typing import TYPE_CHECKING

from legacy_doc._paragraphs import parse_sprms
from legacy_doc.exceptions import LegacyDocError

if TYPE_CHECKING:
    from legacy_doc._binary import BinaryDocument


@dataclass(frozen=True)
class CharacterProperties:
    special: bool | None = False
    ole: bool = False
    object: bool = False
    location: int | None = None


class CharacterIndex:
    """Lazy CHPX access; hidden/revision/font properties do not filter text."""

    def __init__(self, doc: BinaryDocument) -> None:
        self.doc = doc
        self._boundaries: tuple[int, ...] | None = None
        self._pages: tuple[int, ...] = ()
        self._cache: dict[int, tuple[bytes, tuple[int, ...]]] = {}
        self._values: dict[tuple[tuple[int, int], int], CharacterProperties] = {}

    def _load_index(self) -> None:
        offset, size = self.doc.fib.pair(12)  # PlcBteChpx
        if size == 0:
            self._boundaries = ()
            return
        if size < 4 or (size - 4) % 8 or offset + size > len(self.doc.table):
            raise LegacyDocError('Invalid PlcBteChpx range')
        count = (size - 4) // 8
        if count > 32768:
            raise LegacyDocError('CHPX index exceeds parser limit')
        boundaries = struct.unpack_from(f'<{count + 1}I', self.doc.table, offset)
        if any(a >= b for a, b in zip(boundaries, boundaries[1:])):
            raise LegacyDocError('PlcBteChpx FC boundaries are not ordered')
        self._boundaries = boundaries
        self._pages = tuple(value & 0x3FFFFF for value in struct.unpack_from(
            f'<{count}I', self.doc.table, offset + 4 * (count + 1)))

    def _grpprl(self, fc: int) -> tuple[tuple[int, int], bytes]:
        if self._boundaries is None:
            self._load_index()
        boundaries = self._boundaries
        if not boundaries:
            return (-1, -1), b''
        index = bisect_right(boundaries, fc) - 1
        if index < 0 or index >= len(self._pages) or fc >= boundaries[index + 1]:
            raise LegacyDocError('Selected character is outside PlcBteChpx')
        number = self._pages[index]
        cached = self._cache.get(number)
        if cached is None:
            offset = number * 512
            if offset + 512 > min(len(self.doc.word), self.doc.fib.cb_mac):
                raise LegacyDocError('Selected CHPX FKP is outside WordDocument')
            page = self.doc.word[offset:offset + 512]
            count = page[511]
            if not 1 <= count <= 101:
                raise LegacyDocError('Invalid CHPX FKP run count')
            fcs = struct.unpack_from(f'<{count + 1}I', page)
            if any(a >= b for a, b in zip(fcs, fcs[1:])):
                raise LegacyDocError('CHPX FKP FC boundaries are not ordered')
            cached = (page, fcs)
            self._cache[number] = cached
        page, fcs = cached
        run = bisect_right(fcs, fc) - 1
        if run < 0 or run >= len(fcs) - 1 or fc >= fcs[run + 1]:
            raise LegacyDocError('Selected character is outside CHPX FKP')
        if fcs[run] < boundaries[index] or fcs[run + 1] > boundaries[index + 1]:
            raise LegacyDocError('CHPX FKP run exceeds its BTE range')
        offset = 2 * page[4 * len(fcs) + run]
        if offset == 0:
            return (number, run), b''
        if offset < 4 * len(fcs) + len(fcs) - 1 or offset >= 511:
            raise LegacyDocError('CHPX offset overlaps FKP index')
        size = page[offset]
        if offset + 1 + size > 511:
            raise LegacyDocError('CHPX properties extend beyond FKP')
        return (number, run), page[offset + 1:offset + 1 + size]

    def properties(self, cp: int) -> CharacterProperties:
        piece = self.doc.piece_at(cp)
        run, grpprl = self._grpprl(self.doc.fc_for_cp(cp))
        key = (run, piece.prm)
        if key in self._values:
            return self._values[key]
        groups = [grpprl]
        if piece.prm & 1:
            groups.append(self.doc.prcs[piece.prm >> 1])
        elif (piece.prm >> 1) & 0x7F == 0x75:  # Prm0 -> sprmCFSpec
            groups.append(b'\x55\x08' + bytes([piece.prm >> 8]))
        special: bool | None = False
        ole = embedded = False
        location = None
        for group in groups:
            for item in parse_sprms(group):
                if item.opcode == 0x0855:  # sprmCFSpec
                    value = item.operand[0]
                    if value in (0, 1):
                        special = bool(value)
                    elif value in (0x80, 0x81):
                        special = None
                    else:
                        raise LegacyDocError('Invalid sprmCFSpec operand')
                elif item.opcode in (0x080A, 0x0856):  # CFOle2 / CFObj (Bool8)
                    value = item.operand[0]
                    if value not in (0, 1):
                        raise LegacyDocError('Invalid OLE character Bool8 operand')
                    if item.opcode == 0x080A:
                        ole = bool(value)
                    else:
                        embedded = bool(value)
                elif item.opcode == 0x6A03:  # CPicLocation is signed, not a CP
                    location = struct.unpack('<i', item.operand)[0]
        result = CharacterProperties(special, ole, embedded, location)
        self._values[key] = result
        return result

    def is_special(self, cp: int) -> bool:
        special = self.properties(cp).special
        if special is None:
            raise LegacyDocError('Unsupported style-dependent sprmCFSpec value')
        return special
