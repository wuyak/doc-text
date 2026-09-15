"""Regression coverage for real-world FIB layouts in the Apache POI corpus."""

from __future__ import annotations

from pathlib import Path
import struct

import pytest

from legacy_doc import extract_text
from legacy_doc._binary import BinaryDocument, _parse_fib
from legacy_doc.exceptions import LegacyDocError
from legacy_doc.ole import OleReader
from legacy_doc.types import ExtractionOptions


CORPUS = Path(__file__).parent / "data" / "corpus" / "apache-poi"


@pytest.mark.parametrize('filename', ['52117.doc', 'Bug49933.doc'])
def test_old_word_signature_is_reported_as_unsupported(filename):
    # Upstream uses HWPFOldDocument/Word6Extractor for these samples.
    # Their names do not identify their format; the FIB signature does.
    from legacy_doc import extract_text

    with pytest.raises(LegacyDocError, match='Unsupported Word 6/95'):
        extract_text((CORPUS / filename).read_bytes())


def _document(filename: str) -> BinaryDocument:
    options = ExtractionOptions()
    data = (CORPUS / filename).read_bytes()
    ole = OleReader(data, options=options)
    return BinaryDocument(ole, options)


@pytest.mark.parametrize(
    ("filename", "expected"),
    (
        (
            "47304.doc",
            {
                "n_fib": 0x010C,
                "n_fib_base": 0x00C1,
                "cb_rg_fc_lcb": 183,
                "csw_new": 7,
                "n_fib_new": 0x010C,
                "ccp_text": 18,
            },
        ),
        (
            "testCroppedPictures.doc",
            {
                "n_fib": 0x010C,
                "n_fib_base": 0x00C1,
                "cb_rg_fc_lcb": 183,
                "csw_new": 7,
                "n_fib_new": 0x010C,
                "ccp_text": 23,
            },
        ),
        (
            "Bug48075.doc",
            {
                "n_fib": 0x00C1,
                "n_fib_base": 0x00C1,
                "cb_rg_fc_lcb": 183,
                "csw_new": 0,
                "n_fib_new": None,
                "ccp_text": 2377,
            },
        ),
    ),
)
def test_real_corpus_fib_fields_are_read_with_bounded_known_layout(
    filename: str, expected: dict[str, int | None]
) -> None:
    fib = _document(filename).fib

    for field, value in expected.items():
        assert getattr(fib, field) == value


def test_bug48075_main_story_starts_with_expected_cyrillic_heading() -> None:
    document = _document("Bug48075.doc")

    assert document.read_text(0, document.fib.ccp_text).startswith("Приложение")


@pytest.mark.parametrize(
    ("filename", "expected"),
    (("47304.doc", "Just  a “test”"), ("testCroppedPictures.doc", "Sunset:\n\nWinter:")),
)
def test_real_corpus_main_story_text_is_preserved(filename: str, expected: str) -> None:
    data = (CORPUS / filename).read_bytes()

    assert extract_text(data).text == expected


def _sample_word() -> bytearray:
    word = bytearray(
        OleReader(
            (CORPUS / "47304.doc").read_bytes(), options=ExtractionOptions()
        ).read_stream("WordDocument")
    )
    return word


def _sample_csw_new_offset(word: bytes) -> int:
    pair_count = struct.unpack_from("<H", word, 152)[0]
    return 154 + pair_count * 8


def test_known_new_version_accepts_a_longer_declared_extension_tail() -> None:
    word = _sample_word()
    csw_offset = _sample_csw_new_offset(word)
    struct.pack_into("<H", word, csw_offset, 6)

    fib = _parse_fib(bytes(word))
    assert fib.n_fib_new == 0x010C
    assert fib.csw_new == 6


@pytest.mark.parametrize("case", ("short", "truncated", "unknown"))
def test_csw_new_rejects_short_truncated_or_unknown_extensions(case: str) -> None:
    word = _sample_word()
    csw_offset = _sample_csw_new_offset(word)
    if case == "short":
        struct.pack_into("<H", word, csw_offset, 1)
    elif case == "unknown":
        struct.pack_into("<H", word, csw_offset, 2)
        struct.pack_into("<H", word, csw_offset + 2, 0x00C3)
    else:
        word = word[:csw_offset + 2 + 12]

    with pytest.raises(LegacyDocError, match="FIB|Truncated FibRgCswNew"):
        _parse_fib(bytes(word))


def test_unknown_fib_pair_count_is_rejected() -> None:
    word = _sample_word()
    struct.pack_into("<H", word, 152, 184)

    with pytest.raises(LegacyDocError, match="FIB cbRgFcLcb"):
        _parse_fib(bytes(word))
