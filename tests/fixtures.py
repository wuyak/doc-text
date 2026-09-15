from __future__ import annotations

import struct
from datetime import datetime, timezone
from math import ceil
from typing import Any

END_OF_CHAIN = 0xFFFFFFFE
FREE_SECTOR = 0xFFFFFFFF
FAT_SECTOR = 0xFFFFFFFD
SECTOR_SIZE = 512
SUMMARY_INFORMATION_STREAM = "\x05SummaryInformation"
DOCUMENT_SUMMARY_INFORMATION_STREAM = "\x05DocumentSummaryInformation"
VT_I2 = 0x0002
VT_I4 = 0x0003
VT_LPSTR = 0x001E
VT_FILETIME = 0x0040


def make_doc_bytes(
    text: str = "Hello legacy doc",
    *,
    compressed: bool = False,
    encrypted: bool = False,
    use_1table: bool = False,
    include_piece_table: bool = True,
    summary_properties: dict[int, Any] | None = None,
    document_summary_properties: dict[int, Any] | None = None,
    extra_streams: dict[str, bytes] | None = None,
    stories: dict[str, str] | None = None,
    paragraph_properties: dict[int, bytes] | None = None,
    section_break_cps: set[int] | None = None,
    character_properties: dict[int, bytes] | None = None,
) -> bytes:
    """Build a bounded DOC fixture with explicit FIB/CLX and paragraph records.

    paragraph_properties maps the CP of a paragraph's terminal character to
    its ordered grpprl. Text gets a final paragraph mark if one is absent.
    """
    streams = make_word_streams(
        text, compressed=compressed, encrypted=encrypted,
        use_1table=use_1table, include_piece_table=include_piece_table,
        stories=stories, paragraph_properties=paragraph_properties,
        section_break_cps=section_break_cps,
        character_properties=character_properties,
    )
    streams.update(extra_streams or {})
    if summary_properties is not None:
        streams[SUMMARY_INFORMATION_STREAM] = make_property_stream(summary_properties)
    if document_summary_properties is not None:
        streams[DOCUMENT_SUMMARY_INFORMATION_STREAM] = make_property_stream(
            document_summary_properties
        )
    return make_ole_file(streams)


def make_word_streams(
    text: str = "Hello legacy doc", *, compressed: bool = False,
    encrypted: bool = False, use_1table: bool = False,
    include_piece_table: bool = True, stories: dict[str, str] | None = None,
    paragraph_properties: dict[int, bytes] | None = None,
    section_break_cps: set[int] | None = None,
    character_properties: dict[int, bytes] | None = None,
) -> dict[str, bytes]:
    """Expose streams so malformed and multi-piece tests can alter exact records."""
    if not text.endswith("\r"):
        text += "\r"
    story_names = ("ftn", "hdd", "atn", "edn", "txbx", "hdr_txbx")
    tails = [(stories or {}).get(name, "") for name in story_names]
    raw_text = text + "".join(tails) + ("\r" if any(tails) else "")
    cp_count = lambda value: len(value.encode("utf-16le", errors="surrogatepass")) // 2
    encoded = raw_text.encode("cp1252") if compressed else raw_text.encode("utf-16le", errors="surrogatepass")
    step = 1 if compressed else 2
    fc = 0x400
    word = bytearray(fc) + encoded
    struct.pack_into("<HH", word, 0, 0xA5EC, 0x00C1)
    struct.pack_into("<H", word, 0x0A, (0x0100 if encrypted else 0) | (0x0200 if use_1table else 0))
    struct.pack_into("<H", word, 0x0C, 0x00BF)
    struct.pack_into("<H", word, 32, 14)
    struct.pack_into("<H", word, 62, 22)
    # FibRgLw97 starts at 64: cbMac then reserved words and story lengths.
    struct.pack_into("<i", word, 76, cp_count(text))
    for offset, value in zip((80, 84, 92, 96, 100, 104), tails):
        struct.pack_into("<i", word, offset, cp_count(value))
    struct.pack_into("<H", word, 152, 93)
    struct.pack_into("<H", word, 898, 0)  # cswNew
    table = bytearray()

    def put_pair(index: int, payload: bytes) -> None:
        struct.pack_into("<II", word, 154 + 8 * index, len(table), len(payload))
        table.extend(payload)

    if include_piece_table:
        encoded_fc = fc * 2 | 0x40000000 if compressed else fc
        plc = struct.pack("<IIHIH", 0, cp_count(raw_text), 0, encoded_fc, 0)
        put_pair(33, b"\x02" + struct.pack("<I", len(plc)) + plc)

    # Section marks are distinguished from manual page breaks by PlcfSed.
    section_breaks = sorted(section_break_cps or ())
    section_cps = [0, *(cp + 1 for cp in section_breaks), cp_count(text)]
    sed = b"\0\0" + struct.pack("<I", 0xFFFFFFFF) + b"\0" * 6
    put_pair(6, struct.pack('<' + 'I' * len(section_cps), *section_cps)
             + sed * (len(section_cps) - 1))

    # Construct PAPX FKPs, preserving real FC coordinates. A single paragraph
    # may contain supplementary Unicode characters, hence UTF-16 CP counting.
    runs = []
    cp = 0
    start = fc
    for char in raw_text:
        cp += cp_count(char)
        if char in "\r\x07" or (char == "\x0c" and cp - 1 in section_breaks):
            end = fc + cp * step
            runs.append((start, end, (paragraph_properties or {}).get(cp - 1, b"")))
            start = end
    if start < fc + len(encoded):
        runs.append((start, fc + len(encoded), b""))
    page_runs = []
    # Small batches leave enough space for independently specified properties.
    for begin in range(0, len(runs), 8):
        batch = runs[begin:begin + 8]
        word.extend(b"\0" * (-len(word) % 512))
        pn = len(word) // 512
        page = bytearray(512)
        count = len(batch)
        for index, (left, right, grpprl) in enumerate(batch):
            struct.pack_into("<I", page, index * 4, left)
        struct.pack_into("<I", page, count * 4, batch[-1][1])
        cursor = 4 * (count + 1) + 13 * count
        cursor += cursor % 2
        for index, (_, _, grpprl) in enumerate(batch):
            if not grpprl:
                continue  # zero bOffset means default paragraph properties
            payload = b"\0\0" + grpprl  # istd, then complete PRLs
            # Odd grpprl lengths use cb; even lengths use cb=0, cb'.
            record = (bytes([(len(payload) + 1) // 2]) + payload
                      if len(payload) % 2 else b"\0" + bytes([len(payload) // 2]) + payload)
            if cursor + len(record) > 511:
                raise ValueError("Fixture PAPX page property capacity exceeded")
            page[4 * (count + 1) + 13 * index] = cursor // 2
            page[cursor:cursor + len(record)] = record
            cursor += len(record)
            cursor += cursor % 2
        page[511] = count
        word.extend(page)
        page_runs.append((batch[0][0], batch[-1][1], pn))
    if page_runs:
        bte = b"".join(struct.pack("<I", item[0]) for item in page_runs)
        bte += struct.pack("<I", page_runs[-1][1])
        bte += b"".join(struct.pack("<I", item[2]) for item in page_runs)
        put_pair(13, bte)
    # Mark special object/field characters with CFSpec. Normal text remains
    # visible regardless of revision/hidden properties in this extraction policy.
    character_runs = []
    cp = 0
    for char in raw_text:
        next_cp = cp + cp_count(char)
        props = sprm(0x0855, 1) if char in "\x01\x02\x03\x04\x05\x08\x13\x14\x15" else b""
        props = (character_properties or {}).get(cp, props)
        left, right = fc + cp * step, fc + next_cp * step
        if character_runs and character_runs[-1][2] == props:
            character_runs[-1] = (character_runs[-1][0], right, props)
        else:
            character_runs.append((left, right, props))
        cp = next_cp
    if any(props for _, _, props in character_runs):
        chpx_pages = []
        for begin in range(0, len(character_runs), 40):
            batch = character_runs[begin:begin + 40]
            word.extend(b"\0" * (-len(word) % 512))
            pn = len(word) // 512
            page = bytearray(512)
            count = len(batch)
            for index, (left, _, _) in enumerate(batch):
                struct.pack_into("<I", page, index * 4, left)
            struct.pack_into("<I", page, count * 4, batch[-1][1])
            cursor = 4 * (count + 1) + count
            cursor += cursor % 2
            for index, (_, _, props) in enumerate(batch):
                if props:
                    page[4 * (count + 1) + index] = cursor // 2
                    page[cursor:cursor + len(props) + 1] = bytes([len(props)]) + props
                    cursor += len(props) + 1
                    cursor += cursor % 2
            page[511] = count
            word.extend(page)
            chpx_pages.append((batch[0][0], batch[-1][1], pn))
        bte = b"".join(struct.pack("<I", item[0]) for item in chpx_pages)
        bte += struct.pack("<I", chpx_pages[-1][1])
        bte += b"".join(struct.pack("<I", item[2]) for item in chpx_pages)
        put_pair(12, bte)
    struct.pack_into("<I", word, 64, len(word))
    return {"WordDocument": bytes(word), "1Table" if use_1table else "0Table": bytes(table)}


def sprm(opcode: int, value: int | bytes) -> bytes:
    """Encode a fixed-size SPRM for independently specified test properties."""
    sizes = {0: 1, 1: 1, 2: 2, 3: 4, 4: 2, 5: 2, 7: 3}
    if isinstance(value, int):
        value = value.to_bytes(sizes[opcode >> 13], "little", signed=value < 0)
    if opcode >> 13 == 6:
        length = struct.pack("<H", len(value) + 1) if opcode in {0xD608, 0xD606} else bytes([len(value)])
    else:
        length = b""
    return struct.pack("<H", opcode) + length + value


def make_property_stream(properties: dict[int, Any]) -> bytes:
    properties = {1: 1252, **properties}
    values = bytearray()
    entries: list[tuple[int, int]] = []
    value_start = 8 + len(properties) * 8

    for property_id, value in properties.items():
        entries.append((property_id, value_start + len(values)))
        values.extend(_property_value(value))
        while len(values) % 4:
            values.append(0)

    section_size = value_start + len(values)
    section = bytearray()
    section.extend(struct.pack("<II", section_size, len(properties)))
    for property_id, offset in entries:
        section.extend(struct.pack("<II", property_id, offset))
    section.extend(values)

    header = bytearray()
    header.extend(struct.pack("<HHI", 0xFFFE, 0, 0))
    header.extend(b"\x00" * 16)
    header.extend(struct.pack("<I", 1))
    header.extend(b"\x00" * 16)
    header.extend(struct.pack("<I", 48))
    return bytes(header) + bytes(section)


def make_ole_file(streams: dict[str, bytes]) -> bytes:
    sector_payloads: list[bytes] = [b"", b""]
    stream_locations: dict[str, tuple[int, int]] = {}
    fat_entries = [END_OF_CHAIN, FAT_SECTOR]

    for name, payload in streams.items():
        payload = payload.ljust(4096, b"\x00")
        start_sector = len(sector_payloads)
        sector_count = max(1, ceil(len(payload) / SECTOR_SIZE))
        stream_locations[name] = (start_sector, len(payload))
        padded = payload.ljust(sector_count * SECTOR_SIZE, b"\x00")
        for index in range(sector_count):
            sector_payloads.append(
                padded[index * SECTOR_SIZE : (index + 1) * SECTOR_SIZE]
            )
            current_sector = start_sector + index
            next_sector = (
                END_OF_CHAIN if index == sector_count - 1 else current_sector + 1
            )
            fat_entries.append(next_sector)

    entries = [bytearray(_directory_entry("Root Entry", 5, END_OF_CHAIN, 0))]
    entries.extend(bytearray(_directory_entry(name, 2, start_sector, size))
                   for name, (start_sector, size) in stream_locations.items())
    ordered_ids = sorted(range(1, len(entries)),
                         key=lambda i: (len(list(stream_locations)[i - 1]),
                                        list(stream_locations)[i - 1].upper()))
    deepest = len(ordered_ids).bit_length() - 1

    def directory_tree(ids: list[int], depth: int = 0) -> int:
        if not ids:
            return FREE_SECTOR
        mid = len(ids) // 2
        entry_id = ids[mid]
        left = directory_tree(ids[:mid], depth + 1)
        right = directory_tree(ids[mid + 1:], depth + 1)
        struct.pack_into("<II", entries[entry_id], 68, left, right)
        entries[entry_id][67] = 0 if depth == deepest and depth else 1
        return entry_id

    entries[0][67] = 1
    struct.pack_into("<I", entries[0], 76, directory_tree(ordered_ids))
    directory_payload = b"".join(entries)
    directory_sector_count = max(1, ceil(len(directory_payload) / SECTOR_SIZE))
    directory_payload = directory_payload.ljust(
        directory_sector_count * SECTOR_SIZE,
        b"\x00",
    )
    sector_payloads[0] = directory_payload[:SECTOR_SIZE]
    if directory_sector_count > 1:
        first_extra_directory_sector = len(sector_payloads)
        fat_entries[0] = first_extra_directory_sector
        for index in range(1, directory_sector_count):
            sector_payloads.append(
                directory_payload[index * SECTOR_SIZE : (index + 1) * SECTOR_SIZE]
            )
            current_sector = first_extra_directory_sector + index - 1
            next_sector = (
                END_OF_CHAIN
                if index == directory_sector_count - 1
                else current_sector + 1
            )
            fat_entries.append(next_sector)
    if len(fat_entries) > 128:
        raise ValueError("Fixture exceeds the single FAT-sector capacity")
    sector_payloads[1] = struct.pack(
        "<128I",
        *fat_entries[:128],
        *([FREE_SECTOR] * (128 - len(fat_entries[:128]))),
    )

    header = bytearray(SECTOR_SIZE)
    header[:8] = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
    struct.pack_into("<H", header, 0x18, 0x003E)
    struct.pack_into("<H", header, 0x1A, 3)
    struct.pack_into("<H", header, 0x1C, 0xFFFE)
    struct.pack_into("<H", header, 0x1E, 9)
    struct.pack_into("<H", header, 0x20, 6)
    struct.pack_into("<I", header, 0x2C, 1)
    struct.pack_into("<I", header, 0x30, 0)
    struct.pack_into("<I", header, 0x38, 4096)
    struct.pack_into("<I", header, 0x3C, END_OF_CHAIN)
    struct.pack_into("<I", header, 0x40, 0)
    struct.pack_into("<I", header, 0x44, END_OF_CHAIN)
    struct.pack_into("<I", header, 0x48, 0)
    struct.pack_into("<I", header, 0x4C, 1)
    for offset in range(0x50, 0x4C + 109 * 4, 4):
        struct.pack_into("<I", header, offset, FREE_SECTOR)

    return bytes(header) + b"".join(sector_payloads)


def corrupt_fat_entry(document: bytes, sector: int, value: int) -> bytes:
    data = bytearray(document)
    fat_offset = SECTOR_SIZE + SECTOR_SIZE + sector * 4
    struct.pack_into("<I", data, fat_offset, value)
    return bytes(data)


def _property_value(value: Any) -> bytes:
    if isinstance(value, datetime):
        epoch = datetime(1601, 1, 1, tzinfo=timezone.utc)
        filetime = int((value.astimezone(timezone.utc) - epoch).total_seconds() * 10**7)
        return struct.pack("<HHQ", VT_FILETIME, 0, filetime)
    if isinstance(value, str):
        raw = value.encode("cp1252") + b"\x00"
        return struct.pack("<HHI", VT_LPSTR, 0, len(raw)) + raw
    if isinstance(value, int):
        if -32768 <= value <= 32767:
            return struct.pack("<HHh", VT_I2, 0, value)
        return struct.pack("<HHi", VT_I4, 0, value)
    raise TypeError(f"Unsupported property fixture value: {value!r}")


def _directory_entry(name: str, object_type: int, start_sector: int, size: int) -> bytes:
    raw = bytearray(128)
    encoded_name = name.encode("utf-16le") + b"\x00\x00"
    raw[: len(encoded_name)] = encoded_name
    struct.pack_into("<H", raw, 64, len(encoded_name))
    raw[66] = object_type
    struct.pack_into("<III", raw, 68, FREE_SECTOR, FREE_SECTOR, FREE_SECTOR)
    struct.pack_into("<I", raw, 116, start_sector)
    struct.pack_into("<Q", raw, 120, size)
    return bytes(raw)
