"""Regression tests for the MS-DOC ``sprmPChgTabs`` operand boundary."""

from __future__ import annotations

from pathlib import Path
import struct

import pytest

from legacy_doc import extract_text
from legacy_doc._paragraphs import SPRM_PF_IN_TABLE, parse_sprms
from legacy_doc.exceptions import LegacyDocError


_PCHG_TABS = 0xC615
_CORPUS = Path(__file__).parent / "data" / "corpus" / "apache-poi"


def _pchg_tabs_body(
    deletions: tuple[tuple[int, int], ...] = (),
    additions: tuple[tuple[int, int], ...] = (),
) -> bytes:
    """Build the two count bytes and official 16-bit tab records."""

    # PChgTabsDelClose stores two arrays, rgdxaDel followed by rgdxaClose;
    # they are not an array of interleaved (position, close) pairs.
    deleted = bytes([len(deletions)])
    deleted += b"".join(struct.pack("<h", position) for position, _close in deletions)
    deleted += b"".join(struct.pack("<h", close) for _position, close in deletions)

    # PChgTabsAdd likewise stores rgdxaAdd as one array followed by the TBD
    # descriptor array, rather than interleaving each position and descriptor.
    added = bytes([len(additions)])
    added += b"".join(struct.pack("<h", position) for position, _descriptor in additions)
    added += bytes(descriptor for _position, descriptor in additions)
    return deleted + added


def _pchg_tabs(
    *,
    extended: bool,
    deletions: tuple[tuple[int, int], ...] = (),
    additions: tuple[tuple[int, int], ...] = (),
) -> bytes:
    body = _pchg_tabs_body(deletions, additions)
    cb = 0xFF if extended else len(body)
    return struct.pack("<H", _PCHG_TABS) + bytes([cb]) + body


@pytest.mark.parametrize("extended", [False, True], ids=["normal-cb", "extended-cb255"])
def test_pchg_tabs_width_preserves_following_sprm(extended: bool) -> None:
    deletions = ((100, 25),)
    additions = ((300, 2),)
    grpprl = _pchg_tabs(extended=extended, deletions=deletions, additions=additions)
    grpprl += struct.pack("<H", SPRM_PF_IN_TABLE) + b"\x01"

    parsed = parse_sprms(grpprl)

    assert [item.opcode for item in parsed] == [_PCHG_TABS, SPRM_PF_IN_TABLE]
    expected_operand = bytes([0xFF if extended else 9]) + _pchg_tabs_body(deletions, additions)
    assert parsed[0].operand == expected_operand


def test_extended_pchg_tabs_accepts_derived_width_above_byte_cb() -> None:
    deletions = tuple((position, 25) for position in range(64))
    additions = tuple((position + 100, 2) for position in range(64))
    grpprl = _pchg_tabs(extended=True, deletions=deletions, additions=additions)
    grpprl += struct.pack("<H", SPRM_PF_IN_TABLE) + b"\x01"

    parsed = parse_sprms(grpprl)

    assert [item.opcode for item in parsed] == [_PCHG_TABS, SPRM_PF_IN_TABLE]
    # The cb=255 extended form has a two-byte count prefix in its remainder;
    # including cb itself, the complete operand is 3 + 4*d + 3*a bytes.
    assert len(parsed[0].operand) == 3 + (4 * 64) + (3 * 64)


@pytest.mark.parametrize("extended", [False, True], ids=["normal-cb", "extended-cb255"])
def test_pchg_tabs_rejects_truncated_record(extended: bool) -> None:
    operation = _pchg_tabs(extended=extended, additions=((300, 2),))
    truncated = operation[:-1]

    with pytest.raises(LegacyDocError, match="PChgTabs operand is truncated"):
        parse_sprms(truncated)


@pytest.mark.parametrize(
    "body",
    [bytes([65]), bytes([0, 65])],
    ids=["too-many-deletions", "too-many-additions"],
)
def test_pchg_tabs_rejects_count_overrun(body: bytes) -> None:
    operation = struct.pack("<H", _PCHG_TABS) + bytes([0xFF]) + body

    with pytest.raises(LegacyDocError, match="too many"):
        parse_sprms(operation)


def test_pchg_tabs_rejects_mismatched_normal_cb() -> None:
    body = _pchg_tabs_body(deletions=((100, 25),))
    operation = struct.pack("<H", _PCHG_TABS) + bytes([5]) + body

    with pytest.raises(LegacyDocError, match="length does not match"):
        parse_sprms(operation)


@pytest.mark.parametrize("cb", [0, 1])
def test_pchg_tabs_rejects_cb_below_spec_minimum(cb: int) -> None:
    operation = struct.pack("<H", _PCHG_TABS) + bytes([cb, 0, 0])

    with pytest.raises(LegacyDocError, match="invalid length"):
        parse_sprms(operation)


def test_bug33519_keeps_meaningful_main_story_text() -> None:
    text = extract_text((_CORPUS / "Bug33519.doc").read_bytes()).text

    assert text.startswith("РУСЕНСКИ КЛУБ ЗА ПЪТЕШЕСТВИЯ БЯЛА ЗВЕЗДА")
    assert "Планински турове:" in text
    assert "Карпати, Румъния" in text
    assert text.endswith("Явор Асенов")
