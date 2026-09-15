"""Upstream DOC fixtures, classified and asserted under our text contract.

Source test names and adapted conditions are recorded in data/upstream/cases.json.
Tests read local bytes only; no Office installation or network is required.
"""

import hashlib
import json
from pathlib import Path

import pytest

from legacy_doc import LegacyDocError, extract_text


DATA = Path(__file__).parent / 'data' / 'upstream'
CASES = {case['id']: case for case in json.loads((DATA / 'cases.json').read_text())}


def document(case_id):
    case = CASES[case_id]
    data = (DATA / case['file']).read_bytes()
    assert hashlib.sha256(data).hexdigest() == case['sha256'], case_id
    return data


def test_body_paragraphs_keep_stored_text_and_blank_lines():
    # POI TestWordExtractor.p_text1 / testExtractFromParagraphs.
    paragraphs = [
        'This is a simple word document', '',
        'It has a number of paragraphs in it', '',
        'Some of them even feature bold, italic and underlined text', '', '',
        'This bit is in a different font and size', '', '',
        'This bit features some red text.', '', '',
        'It is otherwise very very boring.',
    ]
    assert extract_text(document('body-paragraphs')).text == '\n'.join(paragraphs)


def test_body_unicode_survives_while_headers_and_footers_are_excluded():
    text = extract_text(document('body-unicode')).text
    assert text.startswith('This is a fairly simple word document, over two pages,')
    assert 'GBP - £' in text
    assert 'EUR - €' in text
    assert 'Molière' in text
    assert "L'Avare ou l'École du mensonge" in text
    assert text.endswith('This is page two. Les Précieuses ridicules. The end.')
    # These are the header/footer strings asserted present by POI. Our scope
    # excludes their stories, while preserving identical words in the body.
    assert 'This is a simple header, with a € euro symbol in it.' not in text
    assert 'The footer, with Molière, has Unicode in it.' not in text


def test_body_excludes_footnotes_endnotes_and_comments():
    # POI separately checks TestFootnote, TestEndnote and TestComment in the
    # independent stories. The main document itself contains only Test text.
    assert extract_text(document('body-notes')).text == 'Test text'


def test_empty_document_returns_empty_text_successfully():
    result = extract_text(document('empty-body'))
    assert result.text == ''
    assert result.metadata['chars'] == result.metadata['bytes'] == 0
    assert result.warnings == ()


@pytest.mark.parametrize(('case_id', 'expected'), [
    ('table-simple',
     'This is a Word document that was created using Word 97 – SR2.  '
     'It contains a paragraph, a table consisting of 2 rows and 3 columns '
     'and a final paragraph.\n'
     'Cell 1,1\tCell 1,2\tCell 1,3\nCell 2,1\tCell 2,2\tCell 2,3\n'
     'This text is below the table.'),
    ('table-merges', 'A\tB\nC\tD\tE\tF\n\tG\tH\tI\nJ\nK'),
    ('table-nested', 'A\tB\tC\nD\tE\n1\n2\n3\n4\nF\tG\nH\tI\tJ'),
])
def test_tables_follow_our_text_order_and_separator_rules(case_id, expected):
    # Structural classes come from POI. Expected text was also checked
    # against the previously saved paired DOCX baseline, not generated here.
    assert extract_text(document(case_id)).text == expected


def test_old_word_format_remains_outside_supported_scope():
    with pytest.raises(LegacyDocError, match='Word FIB|Unsupported Word'):
        extract_text(document('unsupported-word95'))


def test_transparent_body_textbox_keeps_its_text():
    # Upstream checks fill transparency. Our condition uses the sample's
    # stored text: transparent formatting must not erase the body textbox.
    assert extract_text(document('textbox-transparent')).text == 'Testing'


def test_different_textbox_fills_do_not_change_text_inclusion():
    # The labels are literal document content, independently visible in its
    # WordDocument stream. Upstream checks fill styles, not text extraction.
    labels = [
        'Basic Color Fill', 'Basic Pattern Fill', 'Basic 1 Color Fill',
        'Basic picture Fill', 'Basic 2 Color Fill', 'Basic Texture Fill',
    ]
    text = extract_text(document('textbox-fills')).text
    for label in labels:
        assert text.count(label) == 1
        text = text.replace(label, '')
    assert not text.strip()
