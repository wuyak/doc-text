"""Regression tests for CFB directory scope and version-specific stream sizes."""

from __future__ import annotations

import struct
from dataclasses import dataclass
from math import ceil

import pytest

from legacy_doc.exceptions import LegacyDocError
from legacy_doc.ole import END_OF_CHAIN, FAT_SECTOR, FREE_SECTOR, OleReader
from legacy_doc.types import ExtractionOptions

MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


@dataclass
class _Spec:
    name: str
    object_type: int
    parent_id: int | None
    payload: bytes = b""


def _make_cfb(
    *,
    version: int = 3,
    include_root_word: bool = True,
    size_overrides: dict[tuple[int, str], int] | None = None,
) -> bytes:
    """Build a small regular-sector CFB with nested storages.

    The directory order deliberately puts the nested ``WordDocument`` before
    the root stream.  All stream payloads are one full sector-sized stream so
    these tests exercise directory scope without requiring a mini FAT.
    """

    root_word = b"ROOT WORD".ljust(4096, b"\x00")
    nested_word = b"NESTED WORD".ljust(4096, b"\x00")
    contents = b"OBJECT CONTENTS".ljust(4096, b"\x00")
    specs = [
        _Spec("Root Entry", 5, None),
        _Spec("Nested", 1, 0),
        _Spec("WordDocument", 2, 1, nested_word),
        _Spec("\x05SummaryInformation", 2, 1, b"NESTED SUMMARY".ljust(4096, b"\x00")),
    ]
    if include_root_word:
        specs.append(_Spec("WordDocument", 2, 0, root_word))
    object_pool_id = len(specs)
    specs.append(_Spec("ObjectPool", 1, 0))
    object_id = len(specs)
    specs.append(_Spec("_42", 1, object_pool_id))
    specs.append(_Spec("Contents", 2, object_id, contents))
    specs.append(_Spec("RootMarker", 2, 0, b"ROOT MARKER".ljust(4096, b"\x00")))

    sector_size = 1 << (9 if version == 3 else 12)
    directory_sector_count = max(1, ceil(len(specs) * 128 / sector_size))
    sector_payloads: list[bytearray] = [
        bytearray(sector_size) for _ in range(directory_sector_count)
    ]
    fat: list[int] = [FREE_SECTOR] * directory_sector_count

    # Build each storage's red-black-tree-shaped sibling links as a balanced
    # binary search tree.  The color byte is irrelevant to this reader.
    left: dict[int, int] = {}
    right: dict[int, int] = {}
    child: dict[int, int] = {}

    def tree(ids: list[int]) -> int:
        if not ids:
            return FREE_SECTOR
        mid = len(ids) // 2
        current = ids[mid]
        left[current] = tree(ids[:mid])
        right[current] = tree(ids[mid + 1 :])
        return current

    for parent_id in [index for index, spec in enumerate(specs) if spec.object_type in {1, 5}]:
        members = [
            index
            for index, spec in enumerate(specs)
            if spec.parent_id == parent_id
        ]
        child[parent_id] = tree(
            sorted(members, key=lambda index: (specs[index].name.casefold(), index))
        )

    stream_locations: dict[int, tuple[int, int]] = {}
    for index, spec in enumerate(specs):
        if spec.object_type != 2:
            continue
        start_sector = len(sector_payloads)
        physical_size = max(len(spec.payload), 1)
        sector_count = max(1, ceil(physical_size / sector_size))
        payload = spec.payload.ljust(sector_count * sector_size, b"\x00")
        for sector_index in range(sector_count):
            sector_payloads.append(
                bytearray(
                    payload[sector_index * sector_size : (sector_index + 1) * sector_size]
                )
            )
            fat.append(
                END_OF_CHAIN
                if sector_index == sector_count - 1
                else start_sector + sector_index + 1
            )
        size = len(spec.payload)
        if size_overrides and (spec.parent_id, spec.name) in size_overrides:
            size = size_overrides[(spec.parent_id, spec.name)]
        stream_locations[index] = (start_sector, size)

    fat_sector = len(sector_payloads)
    sector_payloads.append(bytearray(sector_size))
    fat.append(FAT_SECTOR)
    if len(fat) > sector_size // 4:
        raise AssertionError("test CFB exceeds one FAT sector")

    for index in range(directory_sector_count):
        fat[index] = (
            END_OF_CHAIN
            if index == directory_sector_count - 1
            else index + 1
        )
    sector_payloads[fat_sector][: len(fat) * 4] = struct.pack(
        f"<{len(fat)}I", *fat
    )

    directory = bytearray(directory_sector_count * sector_size)
    for index, spec in enumerate(specs):
        offset = index * 128
        name = spec.name.encode("utf-16le") + b"\x00\x00"
        directory[offset : offset + len(name)] = name
        struct.pack_into("<H", directory, offset + 64, len(name))
        directory[offset + 66] = spec.object_type
        struct.pack_into(
            "<III",
            directory,
            offset + 68,
            left.get(index, FREE_SECTOR),
            right.get(index, FREE_SECTOR),
            child.get(index, FREE_SECTOR),
        )
        start_sector, size = stream_locations.get(
            index, (END_OF_CHAIN, 0)
        )
        struct.pack_into("<I", directory, offset + 116, start_sector)
        struct.pack_into("<Q", directory, offset + 120, size)
    for index in range(directory_sector_count):
        sector_payloads[index][:] = directory[
            index * sector_size : (index + 1) * sector_size
        ]

    header = bytearray(sector_size)
    header[:8] = MAGIC
    struct.pack_into("<H", header, 0x18, 0x003E)
    struct.pack_into("<H", header, 0x1A, version)
    struct.pack_into("<H", header, 0x1C, 0xFFFE)
    struct.pack_into("<H", header, 0x1E, 9 if version == 3 else 12)
    struct.pack_into("<H", header, 0x20, 6)
    struct.pack_into("<I", header, 0x28, 0 if version == 3 else directory_sector_count)
    struct.pack_into("<I", header, 0x2C, 1)
    struct.pack_into("<I", header, 0x30, 0)
    struct.pack_into("<I", header, 0x38, 4096)
    struct.pack_into("<I", header, 0x3C, END_OF_CHAIN)
    struct.pack_into("<I", header, 0x40, 0)
    struct.pack_into("<I", header, 0x44, END_OF_CHAIN)
    struct.pack_into("<I", header, 0x48, 0)
    struct.pack_into("<I", header, 0x4C, fat_sector)
    for offset in range(0x50, 0x4C + 109 * 4, 4):
        struct.pack_into("<I", header, offset, FREE_SECTOR)

    return bytes(header) + b"".join(bytes(sector) for sector in sector_payloads)


def _patch_header(document: bytes, offset: int, value: int, fmt: str = "<H") -> bytes:
    result = bytearray(document)
    struct.pack_into(fmt, result, offset, value)
    return bytes(result)


def test_root_stream_lookup_does_not_fall_back_to_nested_same_name() -> None:
    reader = OleReader(_make_cfb(), options=ExtractionOptions())

    assert reader.read_stream("WordDocument").startswith(b"ROOT WORD")
    assert reader.has_stream("WordDocument") is True
    assert reader.try_read_stream("WordDocument").startswith(b"ROOT WORD")
    assert reader.list_streams().index("WordDocument") == 0

    nested = reader.find_storage(("Nested",))
    assert nested is not None
    nested_word = reader.try_read_storage_stream_prefix(
        nested, "WordDocument", max_size=4096
    )
    assert nested_word is not None
    assert nested_word.startswith(b"NESTED WORD")


def test_directory_ids_and_sibling_child_links_are_preserved() -> None:
    reader = OleReader(_make_cfb(), options=ExtractionOptions())
    by_id = {entry.directory_id: entry for entry in reader.directory}

    assert reader.root_entry.directory_id == 0
    assert reader.root_entry.child_id != FREE_SECTOR
    root_children = reader.storage_children(reader.root_entry)
    assert {entry.name for entry in root_children} == {
        "Nested",
        "ObjectPool",
        "RootMarker",
        "WordDocument",
    }
    for entry in root_children:
        assert by_id[entry.directory_id] == entry
        assert entry.directory_id != reader.root_entry.directory_id

    object_pool = reader.find_storage(("ObjectPool",))
    object_storage = reader.find_storage(("ObjectPool", "_42"))
    assert object_pool is not None
    assert object_storage is not None
    assert object_storage.directory_id in by_id
    contents = reader.try_read_storage_stream_prefix(
        object_storage, "Contents", max_size=4096
    )
    assert contents is not None
    assert contents.startswith(b"OBJECT CONTENTS")


def test_missing_root_stream_is_not_satisfied_by_nested_stream() -> None:
    reader = OleReader(
        _make_cfb(include_root_word=False), options=ExtractionOptions()
    )

    assert "WordDocument" in reader.list_streams()
    assert reader.has_stream("WordDocument") is False
    assert reader.try_read_stream("WordDocument") is None
    with pytest.raises(LegacyDocError, match="Required OLE stream 'WordDocument'"):
        reader.read_stream("WordDocument")


def test_nested_metadata_stream_does_not_masquerade_as_root_metadata() -> None:
    reader = OleReader(_make_cfb(), options=ExtractionOptions())

    assert "\x05SummaryInformation" in reader.list_streams()
    assert reader.has_stream("\x05SummaryInformation") is False
    assert reader.try_read_stream("\x05SummaryInformation") is None


def test_v3_uses_low_32_bits_of_directory_stream_size() -> None:
    reader = OleReader(
        _make_cfb(
            version=3,
            size_overrides={(0, "WordDocument"): (0xDEADBEEF << 32) | 4096},
        ),
        options=ExtractionOptions(),
    )

    entry = next(entry for entry in reader.directory if entry.name == "WordDocument")
    assert entry.size == 4096
    assert reader.read_stream("WordDocument").startswith(b"ROOT WORD")


def test_v4_uses_header_sector_offset_and_full_64_bit_size() -> None:
    reader = OleReader(_make_cfb(version=4), options=ExtractionOptions())

    assert reader.sector_size == 4096
    assert reader.header_size == 4096
    assert reader.read_stream("WordDocument").startswith(b"ROOT WORD")

    huge = OleReader(
        _make_cfb(
            version=4,
            size_overrides={(0, "WordDocument"): (1 << 32) | 4096},
        ),
        options=ExtractionOptions(),
    )
    with pytest.raises(LegacyDocError, match="exceeds parser limit"):
        huge.read_stream("WordDocument")


def test_v4_genuine_high_size_still_rejects_a_truncated_chain() -> None:
    reader = OleReader(
        _make_cfb(
            version=4,
            size_overrides={(0, "WordDocument"): (1 << 32) | 1},
        ),
        options=ExtractionOptions(max_file_bytes=1 << 33),
    )

    with pytest.raises(LegacyDocError, match="chain is truncated"):
        reader.read_stream("WordDocument")


def test_storage_prefix_read_is_bounded_before_full_stream_validation() -> None:
    reader = OleReader(
        _make_cfb(
            version=4,
            size_overrides={(6, "Contents"): (1 << 32) | 4096},
        ),
        options=ExtractionOptions(),
    )
    storage = reader.find_storage(("ObjectPool", "_42"))
    assert storage is not None

    assert reader.try_read_storage_stream_prefix(storage, "Contents", max_size=7) == b"OBJECT "
    assert reader.try_read_storage_stream_prefix(storage, "Missing", max_size=7) is None


@pytest.mark.parametrize(
    ("version", "sector_shift"),
    ((3, 12), (4, 9)),
)
def test_header_rejects_mismatched_version_and_sector_size(
    version: int, sector_shift: int
) -> None:
    document = _patch_header(_make_cfb(version=version), 0x1E, sector_shift)

    with pytest.raises(LegacyDocError, match="version and sector size"):
        OleReader(document, options=ExtractionOptions())


def test_header_rejects_unknown_major_version() -> None:
    document = _patch_header(_make_cfb(version=3), 0x1A, 2)

    with pytest.raises(LegacyDocError, match="Unsupported OLE major version"):
        OleReader(document, options=ExtractionOptions())


@pytest.mark.parametrize('name', ['0Table', '1Table', 'Data'])
def test_other_required_root_streams_never_fall_back_to_nested(name):
    data = bytearray(_make_cfb(include_root_word=False))
    # Nested WordDocument is directory entry 2. Rename only that child.
    position = 512 + 2 * 128
    encoded = name.encode('utf-16le') + b'\0\0'
    data[position:position+64] = encoded.ljust(64, b'\0')
    struct.pack_into('<H', data, position + 64, len(encoded))
    reader = OleReader(bytes(data), options=ExtractionOptions())
    assert reader.try_read_stream(name) is None
    with pytest.raises(LegacyDocError, match='Required OLE stream'):
        reader.read_stream(name)


def test_unused_nested_directory_is_not_traversed_for_main_stream():
    data = bytearray(_make_cfb())
    # Nested storage is directory entry 1; its malformed child is unselected.
    struct.pack_into('<I', data, 512 + 128 + 76, 0x12345678)
    reader = OleReader(bytes(data), options=ExtractionOptions())
    assert reader.read_stream('WordDocument').startswith(b'ROOT WORD')
