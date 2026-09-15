from legacy_doc.normalize import normalize_word_text


def test_keeps_field_instructions_and_results() -> None:
    text = "Before\rPAGE \\* MERGEFORMAT 1\rAfter"
    assert normalize_word_text(text) == "Before\nPAGE \\* MERGEFORMAT 1\nAfter"


def test_preserves_internal_spaces_blank_lines_and_tabs() -> None:
    assert normalize_word_text("  A  B\r\n\r\nC\tD  \n") == "A  B\n\nC\tD"


def test_keeps_private_use_and_bom_as_stored_text() -> None:
    assert normalize_word_text("A\ue123\ufeffB") == "A\ue123\ufeffB"


def test_boundary_tabs_are_not_stripped() -> None:
    assert normalize_word_text("\tA\t") == "\tA\t"
