from pathlib import Path

import pytest

from legacy_doc import ExtractionOptions, LegacyDocError, extract_text
from legacy_doc.ole import OleReader

from tests.fixtures import make_doc_bytes


CORPUS = Path(__file__).parent / "data" / "corpus"


@pytest.mark.parametrize(
    ("relative_path", "root_name"),
    [
        ("apache-poi/test-fields.doc", "Y:\\Desktop"),
        ("libreoffice/n652364.doc", "Y:\\tmp\\ree"),
        ("libreoffice/n750255.doc", "Y:\\tmp\\New"),
        ("libreoffice/n757118.doc", "Y:\\wrk\\lo\\"),
    ],
)
def test_corpus_root_storage_is_identified_by_type(
    relative_path: str, root_name: str
) -> None:
    reader = OleReader(
        (CORPUS / relative_path).read_bytes(), options=ExtractionOptions()
    )

    roots = [entry for entry in reader.directory if entry.object_type == 5]
    assert len(roots) == 1
    assert roots[0].name == root_name
    assert reader.read_stream("WordDocument").startswith(b"\xec\xa5")


@pytest.mark.parametrize(
    ("relative_path", "fragments"),
    [
        (
            "apache-poi/test-fields.doc",
            ("Field in text box:", "EDITTIME", "Here is a link to an endnote"),
        ),
        ("libreoffice/n652364.doc", ("TEXT1", "text1", "TEXT2", "text2")),
        ("libreoffice/n750255.doc", ("one", "two")),
    ],
)
def test_corpus_documents_extract_known_text_fragments(
    relative_path: str, fragments: tuple[str, ...]
) -> None:
    result = extract_text((CORPUS / relative_path).read_bytes())

    assert result.parser == "doc-text"
    for fragment in fragments:
        assert fragment in result.text


def test_corpus_document_with_header_field_is_accepted_by_public_extractor() -> None:
    result = extract_text((CORPUS / "libreoffice/n757118.doc").read_bytes())

    assert result.parser == "doc-text"
    # The source document's PAGE field is in the excluded header story; its
    # SummaryInformation character count still verifies the original content.
    assert result.text == ""
    assert result.metadata["character_count"] == 12


def test_missing_root_storage_is_still_rejected() -> None:
    document = bytearray(make_doc_bytes("No root storage"))
    # The fixture's first directory entry is the root entry.  Clearing its
    # object type makes it an empty entry while leaving all streams intact.
    document[512 + 66] = 0

    with pytest.raises(LegacyDocError, match="root entry"):
        extract_text(bytes(document))
