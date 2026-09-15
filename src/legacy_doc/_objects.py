"""Locate embedded fields in selected stories; never traverse their payloads.

MS-DOC 2.8.25 (Plcfld), 2.9.88 (Fld), 2.6.1 (CPicLocation),
and 2.1.4 (ObjectPool) supply the association. Ordinary fields are retained.
"""
from __future__ import annotations

from bisect import bisect_left, bisect_right
from dataclasses import dataclass
import struct
from typing import TYPE_CHECKING

from legacy_doc.exceptions import LegacyDocError

if TYPE_CHECKING:
    from legacy_doc._binary import BinaryDocument
    from legacy_doc._characters import CharacterIndex


@dataclass(frozen=True)
class EmbeddedField:
    start: int
    separator: int | None
    end: int  # exclusive
    zombie: bool


class EmbeddedObjects:
    def __init__(self, doc: BinaryDocument, characters: CharacterIndex) -> None:
        self.doc = doc
        self.characters = characters
        self._stories: dict[int, tuple[tuple[int, ...], list[EmbeddedField]]] = {}
        self._labels: dict[int, str] = {}
        self._markers: dict[int, tuple[tuple[int, ...], tuple[int, ...]]] = {}

    def _load(self, pair: int, base: int, length: int) -> tuple[tuple[int, ...], list[EmbeddedField]]:
        if pair in self._stories:
            return self._stories[pair]
        offset, size = self.doc.fib.pair(pair)
        fields: list[EmbeddedField] = []
        marker_cps: list[int] = []
        marker_chars: list[int] = []
        if size:
            if size < 4 or (size - 4) % 6 or offset + size > len(self.doc.table):
                raise LegacyDocError('Invalid embedded-field Plcfld range')
            count = (size - 4) // 6
            if count > 1_000_000:
                raise LegacyDocError('DOC field index exceeds parser limit')
            cps = struct.unpack_from(f'<{count + 1}I', self.doc.table, offset)
            if any(a >= b for a, b in zip(cps, cps[1:])):
                raise LegacyDocError('Plcfld CPs are not ordered')
            # The final PLC CP is unused, including for story bounds.
            if count and cps[count - 1] >= length:
                raise LegacyDocError('Plcfld character is outside its story')
            records = offset + 4 * (count + 1)
            stack: list[tuple[int, int, int | None]] = []
            for i, cp in enumerate(cps[:-1]):
                ch, flags = self.doc.table[records + 2*i:records + 2*i + 2]
                ch &= 0x1F  # fldch also contains non-character flags
                cp += base
                marker_cps.append(cp)
                marker_chars.append(ch)
                if ch == 0x13:
                    if len(stack) >= 64:
                        raise LegacyDocError('DOC field nesting exceeds parser limit')
                    stack.append((cp, flags, None))
                elif ch == 0x14 and stack:
                    begin, kind, separator = stack[-1]
                    if separator is not None:
                        raise LegacyDocError('DOC field has multiple separators')
                    stack[-1] = (begin, kind, cp)
                elif ch == 0x15 and stack:
                    begin, kind, separator = stack.pop()
                    if kind == 0x3A:  # EMBED; LINK and CONTROL are not files
                        fields.append(EmbeddedField(begin, separator, cp + 1, bool(flags & 2)))
                else:
                    raise LegacyDocError('Invalid DOC field boundary sequence')
            if stack:
                raise LegacyDocError('Unclosed DOC field in Plcfld')
        fields.sort(key=lambda field: field.start)
        # An embedded result is opaque. Flatten before bisecting so a
        # selection inside a nested field cannot hide its enclosing object.
        outer: list[EmbeddedField] = []
        for field in fields:
            if not outer or field.start >= outer[-1].end:
                outer.append(field)
        fields = outer
        result = (tuple(field.start for field in fields), fields)
        self._markers[pair] = (tuple(marker_cps), tuple(marker_chars))
        self._stories[pair] = result
        return result

    def selected(self, start: int, end: int) -> tuple[tuple[int, ...], list[EmbeddedField]]:
        fib = self.doc.fib
        if end <= fib.ccp_text:
            pair = 16
            starts, fields = self._load(pair, 0, fib.ccp_text)
        else:
            base = fib.ccp_text + fib.ccp_ftn + fib.ccp_hdd + fib.ccp_atn + fib.ccp_edn
            if not base <= start <= end <= base + fib.ccp_txbx:
                raise LegacyDocError('Embedded field selection is outside supported stories')
            pair = 57
            starts, fields = self._load(pair, base, fib.ccp_txbx)
        marker_cps, marker_chars = self._markers[pair]
        # Validate shared index markers only inside this selected range.
        # Unselected textbox content must not be decoded for this check.
        for i in range(bisect_left(marker_cps, start), bisect_left(marker_cps, end)):
            cp = marker_cps[i]
            if self.doc.read_text(cp, cp + 1) != chr(marker_chars[i]):
                raise LegacyDocError('Plcfld marker does not match selected text')
        first = max(0, bisect_right(starts, start) - 1)
        selected: list[EmbeddedField] = []
        for field in fields[first:]:
            if field.start >= end:
                break
            if field.end <= start:
                continue
            if field.start < start or field.end > end:
                raise LegacyDocError('Selected text splits an embedded field')
            selected.append(field)
        return tuple(field.start for field in selected), selected

    def label(self, field: EmbeddedField) -> str:
        if field.start in self._labels:
            return self._labels[field.start]
        for cp, char in ((field.start, '\x13'), (field.end - 1, '\x15')):
            if self.doc.read_text(cp, cp + 1) != char or not self.characters.is_special(cp):
                raise LegacyDocError('Embedded field does not reference a special field character')
        label = '[嵌入文件]'
        if field.separator is not None:
            cp = field.separator
            if self.doc.read_text(cp, cp + 1) != '\x14' or not self.characters.is_special(cp):
                raise LegacyDocError('Embedded field separator is not a special field character')
            props = self.characters.properties(cp)
            # Zombie fields explicitly forbid using CPicLocation as storage ID.
            if not field.zombie and props.ole and props.object and props.location is not None:
                from legacy_doc._object_metadata import object_label
                try:
                    storage = self.doc.ole.find_storage(('ObjectPool', f'_{props.location}'))
                except LegacyDocError:
                    storage = None  # Optional attachment metadata cannot erase the body.
                if storage is not None:
                    label = object_label(self.doc.ole, storage)
        self._labels[field.start] = label
        return label
