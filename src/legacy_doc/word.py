from __future__ import annotations

from dataclasses import dataclass

from legacy_doc.exceptions import LegacyDocError
from legacy_doc._binary import BinaryDocument
from legacy_doc._shapes import body_textboxes
from legacy_doc._text import TextRenderer
from legacy_doc.ole import OleReader
from legacy_doc.properties import (
    DOCUMENT_SUMMARY_INFORMATION_STREAM,
    SUMMARY_INFORMATION_STREAM,
    parse_document_summary_information,
    parse_summary_information,
)
from legacy_doc.types import ExtractionOptions


@dataclass(frozen=True)
class WordExtraction:
    text: str
    metadata: dict[str, object]
    warnings: tuple[str, ...] = ()


def extract_word_document(
    document_bytes: bytes,
    *,
    options: ExtractionOptions,
) -> WordExtraction:
    ole = OleReader(document_bytes, options=options)
    if ole.has_stream("EncryptedPackage"):
        raise LegacyDocError("Encrypted legacy .doc files are not supported")

    metadata, warnings = _extract_metadata(ole)

    document = BinaryDocument(ole, options=options)
    text = TextRenderer(document, body_textboxes(document)).render()
    return WordExtraction(text=text, metadata=metadata, warnings=tuple(warnings))


def _extract_metadata(ole: OleReader) -> tuple[dict[str, object], list[str]]:
    stream_names = ole.list_streams()
    folded_stream_names = {name.casefold() for name in stream_names}
    metadata: dict[str, object] = {
        "has_summary_information": ole.has_stream(SUMMARY_INFORMATION_STREAM),
        "has_document_summary_information": ole.has_stream(
            DOCUMENT_SUMMARY_INFORMATION_STREAM
        ),
        "has_macros": _has_macros(folded_stream_names),
        "has_embedded_objects": (ole.find_storage(("ObjectPool",)) is not None
                                 or _has_embedded_objects(folded_stream_names)),
        "ole_stream_count": len(stream_names),
    }
    warnings: list[str] = []

    for stream_name, parser, warning in (
        (SUMMARY_INFORMATION_STREAM, parse_summary_information,
         "SummaryInformation stream could not be parsed"),
        (DOCUMENT_SUMMARY_INFORMATION_STREAM, parse_document_summary_information,
         "DocumentSummaryInformation stream could not be parsed"),
    ):
        try:
            stream = ole.try_read_stream(stream_name)
            if stream is not None:
                metadata.update(parser(stream))
        except LegacyDocError:
            warnings.append(warning)

    return metadata, warnings


def _has_macros(stream_names: set[str]) -> bool:
    macro_markers = {"vba", "_vba_project", "_vba_project_cur", "dir", "project"}
    return bool(stream_names & macro_markers)


def _has_embedded_objects(stream_names: set[str]) -> bool:
    return any(
        name == "objectpool" or name.startswith("ole") or name.startswith("package")
        for name in stream_names
    )
