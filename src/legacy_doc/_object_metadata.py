"""Bounded object descriptions; attachment payloads stay opaque.

MS-OLEDS CompObjStream provides display types. Package filenames use the
structured Ole10Native header documented by Apache POI; MS-OLEDS itself
leaves NativeData application-defined. No encoding or suffix is guessed.
"""
from __future__ import annotations

import struct
from typing import TYPE_CHECKING

from legacy_doc.exceptions import LegacyDocError

if TYPE_CHECKING:
    from legacy_doc.ole import DirectoryEntry, OleReader

MAX_METADATA_BYTES = 4096
MAX_STRING_BYTES = 1024
MAX_LABEL_CHARS = 256


def object_label(ole: OleReader, storage: DirectoryEntry) -> str:
    raw = _prefix(ole, storage, '\x01CompObj')
    object_type = _comp_obj_type(raw) if raw is not None else None
    # NativeData has many application-specific representations. Only the
    # named Package forms have the filename header handled here.
    if object_type and object_type.casefold() in ('package', 'package2'):
        raw = _prefix(ole, storage, '\x01Ole10Native')
        filename = _package_filename(raw) if raw is not None else None
        if filename:
            return f'[嵌入文件：{filename}]'
    return f'[嵌入文件：{object_type}]' if object_type else '[嵌入文件]'


def _prefix(ole: OleReader, storage: DirectoryEntry, name: str) -> bytes | None:
    try:
        return ole.try_read_storage_stream_prefix(storage, name, max_size=MAX_METADATA_BYTES)
    except LegacyDocError:
        return None  # Optional description failure does not erase main text.


def _u32(data: bytes, offset: int) -> int | None:
    if offset < 0 or offset + 4 > len(data):
        return None
    return struct.unpack_from('<I', data, offset)[0]


def _lp(data: bytes, offset: int, *, unicode: bool = False) -> tuple[bytes, int] | None:
    length = _u32(data, offset)
    if length is None or length > MAX_STRING_BYTES:
        return None
    start = offset + 4
    end = start + length
    if end > len(data):
        return None
    if not length:
        return b'', end
    unit = 2 if unicode else 1
    if length % unit or data[end-unit:end] != b'\0' * unit:
        return None
    return data[start:end-unit], end


def _clip_end(data: bytes, offset: int) -> int | None:
    marker = _u32(data, offset)
    if marker is None:
        return None
    if marker == 0:
        return offset + 4
    if marker in (0xFFFFFFFF, 0xFFFFFFFE):
        return offset + 8 if offset + 8 <= len(data) else None
    if marker > 0x190:
        return None
    value = _lp(data, offset)
    return value[1] if value else None


def _description(raw: bytes, encoding: str) -> str | None:
    try:
        value = raw.decode(encoding, errors='strict').strip()
    except UnicodeDecodeError:
        return None
    if not value or len(value) > MAX_LABEL_CHARS:
        return None
    if any(ord(char) < 0x20 or 0x7F <= ord(char) < 0xA0 for char in value):
        return None
    return value


def _comp_obj_type(data: bytes) -> str | None:
    ansi = _lp(data, 28)  # CompObjHeader is 28 bytes (reserved fields ignored).
    if ansi is None:
        return None
    raw, cursor = ansi
    fallback = _description(raw, 'ascii')
    cursor = _clip_end(data, cursor)
    if cursor is None:
        return fallback
    length = _u32(data, cursor)
    # MS-OLEDS requires ignoring all following fields for missing/zero/large
    # Reserved1; a later UnicodeMarker cannot revive that ignored tail.
    if length is None or not 1 <= length <= 0x28:
        return fallback
    reserved = _lp(data, cursor)
    if reserved is None:
        return fallback
    cursor = reserved[1]
    if _u32(data, cursor) != 0x71B239F4:
        return fallback
    unicode = _lp(data, cursor + 4, unicode=True)
    if unicode is None:
        return fallback
    return _description(unicode[0], 'utf-16le') or fallback


def _cstring(data: bytes, cursor: int) -> tuple[bytes, int] | None:
    end = data.find(b'\0', cursor, min(len(data), cursor + MAX_STRING_BYTES))
    return (data[cursor:end], end + 1) if end >= 0 else None


def _package_filename(data: bytes) -> str | None:
    total = _u32(data, 0)
    if total is None or total < 2 or data[4:6] != b'\x02\0':
        return None
    # The four-byte total excludes itself. Restrict header reads to this
    # declared record as well as the small physical prefix already obtained.
    data = data[:min(len(data), total + 4)]
    label = _cstring(data, 6)
    if label is None:
        return None
    filename = _cstring(data, label[1])
    if filename is None:
        return None
    # flags2 and unknown1 are two bytes each; then command ASCIIZ and dataSize.
    command = _cstring(data, filename[1] + 4)
    if command is None:
        return None
    payload_size = _u32(data, command[1])
    if payload_size is None or command[1] + 4 + payload_size > total + 4:
        return None
    value = _description(filename[0], 'ascii')
    if not value:
        return None
    value = value.replace('\\', '/').rsplit('/', 1)[-1]
    if '.' not in value:
        return None
    stem, extension = value.rsplit('.', 1)
    if not stem or not extension or any(char.isspace() for char in extension):
        return None
    return value
