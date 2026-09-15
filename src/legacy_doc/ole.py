from __future__ import annotations

import struct
from collections.abc import Iterable
from dataclasses import dataclass

from legacy_doc.exceptions import LegacyDocError
from legacy_doc.types import ExtractionOptions

OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
END_OF_CHAIN = 0xFFFFFFFE
FREE_SECTOR = 0xFFFFFFFF
FAT_SECTOR = 0xFFFFFFFD
MINIFAT_SECTOR = 0xFFFFFFFC
DIRECTORY_NOSTREAM = FREE_SECTOR
MAX_DIRECTORY_DEPTH = 64
MAX_DIRECTORY_ENTRIES = 1_000_000


@dataclass(frozen=True)
class DirectoryEntry:
    name: str
    object_type: int
    start_sector: int
    size: int
    directory_id: int = 0
    left_sibling_id: int = DIRECTORY_NOSTREAM
    right_sibling_id: int = DIRECTORY_NOSTREAM
    child_id: int = DIRECTORY_NOSTREAM

    @property
    def is_storage(self) -> bool:
        return self.object_type in {1, 5}

    @property
    def is_stream(self) -> bool:
        return self.object_type == 2


class OleReader:
    def __init__(self, data: bytes, *, options: ExtractionOptions) -> None:
        self.options = options
        if len(data) > options.max_file_bytes:
            raise LegacyDocError(".doc file exceeds parser file-size limit")
        if len(data) < 512 or data[:8] != OLE_MAGIC:
            raise LegacyDocError("Only OLE Compound File .doc files are supported")

        self.data = data
        if self._u16(0x1C) != 0xFFFE:
            raise LegacyDocError("Unsupported OLE byte order")

        self.major_version = self._u16(0x1A)
        if self.major_version not in {3, 4}:
            raise LegacyDocError(
                f"Unsupported OLE major version {self.major_version}"
            )

        sector_shift = self._u16(0x1E)
        mini_sector_shift = self._u16(0x20)
        expected_sector_shift = 9 if self.major_version == 3 else 12
        if sector_shift != expected_sector_shift:
            raise LegacyDocError(
                "OLE major version and sector size do not match "
                f"(version {self.major_version}, shift {sector_shift})"
            )
        if mini_sector_shift != 6:
            raise LegacyDocError("Unsupported OLE mini-stream sector size")

        self.sector_size = 1 << sector_shift
        self.mini_sector_size = 1 << mini_sector_shift
        self.header_size = self.sector_size
        if len(data) < self.header_size:
            raise LegacyDocError("OLE file is truncated before its header sector")

        self.num_directory_sectors = self._u32(0x28)
        if self.major_version == 3 and self.num_directory_sectors != 0:
            raise LegacyDocError(
                "Version 3 OLE files must not declare directory sectors"
            )
        self.num_fat_sectors = self._u32(0x2C)
        if self.num_fat_sectors == 0:
            raise LegacyDocError("OLE file does not declare a FAT sector")

        self.mini_stream_cutoff = self._u32(0x38)
        self.first_dir_sector = self._u32(0x30)
        self.first_minifat_sector = self._u32(0x3C)
        self.num_minifat_sectors = self._u32(0x40)
        first_difat_sector = self._u32(0x44)
        num_difat_sectors = self._u32(0x48)

        difat = [
            self._u32(offset)
            for offset in range(0x4C, 0x4C + 109 * 4, 4)
            if self._u32(offset) not in {FREE_SECTOR, END_OF_CHAIN}
        ]
        difat.extend(self._read_difat_chain(first_difat_sector, num_difat_sectors))
        if len(difat) < self.num_fat_sectors:
            raise LegacyDocError(
                "OLE DIFAT contains fewer FAT sectors than the header declares"
            )
        difat = difat[: self.num_fat_sectors]

        self.fat = self._read_fat(difat)
        self.directory = self._read_directory()
        self._directory_by_id = {entry.directory_id: entry for entry in self.directory}
        self._children_cache: dict[int, tuple[DirectoryEntry, ...]] = {}
        self._directory_tree_parents: dict[int, int] = {}
        root = self._find_root_entry()
        self.root_entry = root
        if self.root_entry.size > self.options.max_file_bytes:
            raise LegacyDocError("OLE root storage exceeds parser file-size limit")
        self.mini_stream = self._read_regular_stream(
            self.root_entry.start_sector,
            self.root_entry.size,
        )
        self.minifat = self._read_minifat()

    def read_stream(self, name: str, max_size: int | None = None) -> bytes:
        entry = self._find_direct_child(self.root_entry, name, object_type=2)
        if entry is None:
            raise LegacyDocError(f"Required OLE stream '{name}' not found")
        return self._read_entry(entry, max_size=max_size)

    def has_stream(self, name: str) -> bool:
        return self._find_direct_child(self.root_entry, name, object_type=2) is not None

    def try_read_stream(self, name: str, max_size: int | None = None) -> bytes | None:
        entry = self._find_direct_child(self.root_entry, name, object_type=2)
        if entry is None:
            return None
        return self._read_entry(entry, max_size=max_size)

    def list_streams(self) -> tuple[str, ...]:
        """Return all stream names in directory order for inventory metadata.

        Stream lookup is intentionally root-scoped; this inventory method
        retains the historical all-stream behavior so callers can account for
        nested object streams without accidentally using them as DOC inputs.
        """

        return tuple(entry.name for entry in self.directory if entry.object_type == 2)

    def find_storage(self, path: Iterable[str]) -> DirectoryEntry | None:
        """Find a storage by a root-relative, case-insensitive path.

        Each path component must identify a direct storage child of the
        previous component.  The method never searches descendants by name,
        which prevents an embedded storage from satisfying a missing root
        stream or an unrelated path component.
        """

        if isinstance(path, str):
            components = (path,)
        else:
            try:
                collected: list[str] = []
                for component in path:
                    if len(collected) >= MAX_DIRECTORY_DEPTH:
                        raise LegacyDocError("OLE storage path exceeds parser limit")
                    collected.append(component)
                components = tuple(collected)
            except TypeError as exc:
                raise LegacyDocError("OLE storage path must be iterable") from exc

        storage = self.root_entry
        for component in components:
            if not isinstance(component, str):
                raise LegacyDocError("OLE storage path components must be strings")
            child = self._find_direct_child(storage, component, object_type=1)
            if child is None:
                return None
            storage = child
        return storage

    def storage_children(
        self,
        storage: DirectoryEntry | int,
    ) -> tuple[DirectoryEntry, ...]:
        """Return the bounded direct children of a storage.

        The returned entries are handles into this reader's directory table.
        Their ``directory_id`` and sibling/child IDs remain stable, so a
        caller can pass an entry back to the bounded prefix reader without
        exposing arbitrary directory traversal.
        """

        owner = self._resolve_storage(storage)
        cached = self._children_cache.get(owner.directory_id)
        if cached is not None:
            return cached
        children = tuple(self._iter_sibling_tree(owner.child_id))
        for child in children:
            if child.directory_id == owner.directory_id:
                raise LegacyDocError("Cyclic OLE directory storage tree")
            previous_owner = self._directory_tree_parents.get(child.directory_id)
            if previous_owner is not None and previous_owner != owner.directory_id:
                raise LegacyDocError("OLE directory entry has multiple parents")
            self._directory_tree_parents[child.directory_id] = owner.directory_id
        self._children_cache[owner.directory_id] = children
        return children

    def try_read_storage_stream_prefix(
        self,
        storage: DirectoryEntry | int,
        name: str,
        max_size: int | None = None,
    ) -> bytes | None:
        """Read a bounded direct child prefix, returning ``None`` if absent."""

        owner = self._resolve_storage(storage)
        entry = self._find_direct_child(owner, name, object_type=2)
        if entry is None:
            return None
        return self._read_entry_prefix(entry, max_size=max_size)

    def _read_difat_chain(self, first_sector: int, count: int) -> list[int]:
        if count == 0:
            if first_sector not in {FREE_SECTOR, END_OF_CHAIN}:
                raise LegacyDocError(
                    "OLE DIFAT has a starting sector but no declared DIFAT sectors"
                )
            return []
        if first_sector in {FREE_SECTOR, END_OF_CHAIN}:
            raise LegacyDocError("OLE DIFAT chain is truncated")
        result: list[int] = []
        sector = first_sector
        seen: set[int] = set()
        for index in range(count):
            if sector in seen:
                raise LegacyDocError("Cyclic OLE DIFAT chain")
            if len(seen) >= self.options.max_chain_sectors:
                raise LegacyDocError("OLE DIFAT chain exceeds parser limit")
            seen.add(sector)
            raw = self._sector(sector)
            entries_per_sector = self.sector_size // 4
            result.extend(
                value
                for value in struct.unpack_from(f"<{entries_per_sector - 1}I", raw, 0)
                if value not in {FREE_SECTOR, END_OF_CHAIN}
            )
            sector = struct.unpack_from("<I", raw, self.sector_size - 4)[0]
            if index == count - 1:
                if sector != END_OF_CHAIN:
                    raise LegacyDocError("OLE DIFAT chain exceeds declared length")
                break
            if sector in {FREE_SECTOR, END_OF_CHAIN}:
                raise LegacyDocError("OLE DIFAT chain is truncated")
        return result

    def _read_fat(self, fat_sectors: Iterable[int]) -> list[int]:
        fat: list[int] = []
        entries_per_sector = self.sector_size // 4
        for sector in fat_sectors:
            raw = self._sector(sector)
            fat.extend(struct.unpack_from(f"<{entries_per_sector}I", raw, 0))
        return fat

    def _read_minifat(self) -> list[int]:
        if self.num_minifat_sectors == 0 or self.first_minifat_sector in {
            FREE_SECTOR,
            END_OF_CHAIN,
        }:
            if self.num_minifat_sectors:
                raise LegacyDocError("OLE MiniFAT chain is truncated")
            return []
        size = self.num_minifat_sectors * self.sector_size
        if size > self.options.max_file_bytes:
            raise LegacyDocError("OLE MiniFAT exceeds parser file-size limit")
        raw = self._read_regular_stream(
            self.first_minifat_sector,
            size,
            expected_sector_count=self.num_minifat_sectors,
        )
        return list(struct.unpack_from(f"<{len(raw) // 4}I", raw, 0))

    def _read_directory(self) -> list[DirectoryEntry]:
        expected_sectors = (
            self.num_directory_sectors if self.major_version == 4 else None
        )
        raw = self._read_regular_stream(
            self.first_dir_sector,
            self.options.max_file_bytes,
            exact_chain=False,
            expected_sector_count=expected_sectors,
        )
        if len(raw) % 128:
            raise LegacyDocError("OLE directory stream is truncated")
        entries: list[DirectoryEntry] = []
        for directory_id, offset in enumerate(range(0, len(raw), 128)):
            chunk = raw[offset : offset + 128]
            if len(chunk) != 128:
                raise LegacyDocError("OLE directory entry is truncated")
            name_len = struct.unpack_from("<H", chunk, 64)[0]
            object_type = chunk[66]
            if object_type == 0 or name_len < 2:
                continue
            if name_len > 64 or name_len % 2:
                raise LegacyDocError("Invalid OLE directory entry name length")
            try:
                name = chunk[: name_len - 2].decode("utf-16le", errors="strict")
            except UnicodeDecodeError:
                raise LegacyDocError("Invalid OLE directory entry name")
            entries.append(
                DirectoryEntry(
                    name=name,
                    object_type=object_type,
                    start_sector=struct.unpack_from("<I", chunk, 116)[0],
                    size=self._directory_stream_size(chunk),
                    directory_id=directory_id,
                    left_sibling_id=struct.unpack_from("<I", chunk, 68)[0],
                    right_sibling_id=struct.unpack_from("<I", chunk, 72)[0],
                    child_id=struct.unpack_from("<I", chunk, 76)[0],
                )
            )
        if not entries:
            raise LegacyDocError("OLE directory is empty or unreadable")
        return entries

    def _read_regular_stream(
        self,
        start_sector: int,
        size: int,
        *,
        exact_chain: bool = True,
        expected_sector_count: int | None = None,
    ) -> bytes:
        if size < 0:
            raise LegacyDocError("OLE stream size cannot be negative")
        if expected_sector_count is not None and expected_sector_count < 0:
            raise LegacyDocError("OLE sector count cannot be negative")
        if size == 0:
            if expected_sector_count not in {None, 0}:
                raise LegacyDocError("OLE stream chain is truncated")
            return b""
        if start_sector in {FREE_SECTOR, END_OF_CHAIN}:
            raise LegacyDocError("OLE stream chain is truncated")
        chunks: list[bytes] = []
        sector = start_sector
        seen: set[int] = set()
        bytes_read = 0
        while sector != END_OF_CHAIN:
            if sector in seen:
                raise LegacyDocError("Cyclic OLE FAT chain")
            if len(seen) >= self.options.max_chain_sectors:
                raise LegacyDocError("OLE FAT chain exceeds parser limit")
            seen.add(sector)
            chunks.append(self._sector(sector))
            bytes_read += self.sector_size
            if sector >= len(self.fat):
                raise LegacyDocError("OLE FAT chain references an invalid sector")
            next_sector = self.fat[sector]
            if next_sector in {FREE_SECTOR, FAT_SECTOR, MINIFAT_SECTOR}:
                raise LegacyDocError("OLE FAT chain references a non-data sector")
            sector = next_sector
            if exact_chain and bytes_read >= size:
                break
            if not exact_chain and bytes_read > size:
                raise LegacyDocError("OLE stream exceeds parser limit")
        if exact_chain and bytes_read < size:
            raise LegacyDocError("OLE stream chain is truncated")
        if expected_sector_count is not None and len(chunks) != expected_sector_count:
            raise LegacyDocError("OLE stream chain length does not match header")
        return b"".join(chunks)[:size]

    def _read_mini_stream(self, start_sector: int, size: int) -> bytes:
        if size == 0:
            return b""
        if start_sector in {FREE_SECTOR, END_OF_CHAIN}:
            raise LegacyDocError("OLE MiniFAT chain is truncated")
        chunks: list[bytes] = []
        sector = start_sector
        seen: set[int] = set()
        bytes_read = 0
        while sector != END_OF_CHAIN:
            if sector in seen:
                raise LegacyDocError("Cyclic OLE MiniFAT chain")
            if len(seen) >= self.options.max_chain_sectors:
                raise LegacyDocError("OLE MiniFAT chain exceeds parser limit")
            seen.add(sector)
            offset = sector * self.mini_sector_size
            end = offset + self.mini_sector_size
            if end > len(self.mini_stream):
                raise LegacyDocError("MiniFAT chain references an invalid mini sector")
            chunks.append(self.mini_stream[offset:end])
            bytes_read += self.mini_sector_size
            if sector >= len(self.minifat):
                raise LegacyDocError("MiniFAT chain references an invalid sector")
            sector = self.minifat[sector]
            if sector in {FREE_SECTOR, FAT_SECTOR, MINIFAT_SECTOR}:
                raise LegacyDocError("MiniFAT chain references a non-data sector")
            if bytes_read >= size:
                break
        if bytes_read < size:
            raise LegacyDocError("OLE MiniFAT chain is truncated")
        return b"".join(chunks)[:size]

    def _read_entry(
        self,
        entry: DirectoryEntry,
        *,
        max_size: int | None,
    ) -> bytes:
        limit = self._stream_limit(max_size)
        if entry.size > limit:
            raise LegacyDocError(f"OLE stream '{entry.name}' exceeds parser limit")
        if entry.size < self.mini_stream_cutoff:
            return self._read_mini_stream(entry.start_sector, entry.size)
        return self._read_regular_stream(entry.start_sector, entry.size)

    def _read_entry_prefix(
        self,
        entry: DirectoryEntry,
        *,
        max_size: int | None,
    ) -> bytes:
        limit = self._stream_limit(max_size)
        if limit == 0 or entry.size == 0:
            return b""
        prefix_size = min(entry.size, limit)
        if entry.size < self.mini_stream_cutoff:
            return self._read_mini_stream(entry.start_sector, prefix_size)
        return self._read_regular_stream(entry.start_sector, prefix_size)

    def _stream_limit(self, max_size: int | None) -> int:
        if max_size is None:
            return self.options.max_file_bytes
        if isinstance(max_size, bool) or not isinstance(max_size, int) or max_size < 0:
            raise LegacyDocError("OLE stream limit must be a nonnegative integer")
        return min(max_size, self.options.max_file_bytes)

    def _find_direct_child(
        self,
        storage: DirectoryEntry,
        name: str,
        *,
        object_type: int | None = None,
    ) -> DirectoryEntry | None:
        if not isinstance(name, str):
            raise LegacyDocError("OLE directory name must be a string")
        wanted = name.casefold()
        matches: list[DirectoryEntry] = []
        for entry in self.storage_children(storage):
            if entry.name.casefold() != wanted:
                continue
            if object_type is None or entry.object_type == object_type:
                matches.append(entry)
        if len(matches) > 1:
            raise LegacyDocError(
                f"Ambiguous OLE directory name '{name}' in storage '{storage.name}'"
            )
        return matches[0] if matches else None

    def _iter_sibling_tree(self, root_id: int) -> Iterable[DirectoryEntry]:
        if root_id in {DIRECTORY_NOSTREAM, END_OF_CHAIN}:
            return ()
        result: list[DirectoryEntry] = []
        stack: list[int] = []
        current = root_id
        seen: set[int] = set()
        while stack or current not in {DIRECTORY_NOSTREAM, END_OF_CHAIN}:
            while current not in {DIRECTORY_NOSTREAM, END_OF_CHAIN}:
                if current in seen:
                    raise LegacyDocError("Cyclic OLE directory sibling tree")
                if len(seen) >= MAX_DIRECTORY_ENTRIES:
                    raise LegacyDocError("OLE directory tree exceeds parser limit")
                seen.add(current)
                entry = self._directory_by_id.get(current)
                if entry is None:
                    raise LegacyDocError(
                        "OLE directory tree references an invalid entry"
                    )
                if entry.object_type not in {1, 2, 5}:
                    raise LegacyDocError("OLE directory tree references an unallocated entry")
                stack.append(current)
                current = entry.left_sibling_id
            if not stack:
                break
            current = stack.pop()
            entry = self._directory_by_id[current]
            result.append(entry)
            current = entry.right_sibling_id
        return tuple(result)

    def _resolve_entry(self, entry: DirectoryEntry | int) -> DirectoryEntry:
        if isinstance(entry, bool):
            raise LegacyDocError("OLE directory entry ID must be an integer")
        if isinstance(entry, int):
            resolved = self._directory_by_id.get(entry)
        elif isinstance(entry, DirectoryEntry):
            resolved = self._directory_by_id.get(entry.directory_id)
            if resolved is not entry:
                raise LegacyDocError("OLE directory entry does not belong to this file")
        else:
            raise LegacyDocError("OLE directory entry must be an entry or ID")
        if resolved is None:
            raise LegacyDocError("OLE directory entry does not exist")
        return resolved

    def _resolve_storage(self, storage: DirectoryEntry | int) -> DirectoryEntry:
        entry = self._resolve_entry(storage)
        if not entry.is_storage:
            raise LegacyDocError("OLE directory entry is not a storage")
        return entry

    def _find_root_entry(self) -> DirectoryEntry:
        # MS-CFB DirectoryEntry.ObjectType 0x05 identifies the root storage;
        # its display name is implementation-defined.
        roots = [entry for entry in self.directory if entry.object_type == 5]
        if not roots:
            raise LegacyDocError("Required OLE root entry not found")
        if len(roots) != 1:
            raise LegacyDocError("OLE directory contains multiple root entries")
        return roots[0]

    def _directory_stream_size(self, chunk: bytes) -> int:
        size = struct.unpack_from("<Q", chunk, 120)[0]
        if self.major_version == 3:
            return size & 0xFFFFFFFF
        return size

    def _sector(self, sector: int) -> bytes:
        start = self.header_size + sector * self.sector_size
        end = start + self.sector_size
        if sector < 0 or end > len(self.data):
            raise LegacyDocError("OLE sector index is outside the file")
        return self.data[start:end]

    def _u16(self, offset: int) -> int:
        return struct.unpack_from("<H", self.data, offset)[0]

    def _u32(self, offset: int) -> int:
        return struct.unpack_from("<I", self.data, offset)[0]
