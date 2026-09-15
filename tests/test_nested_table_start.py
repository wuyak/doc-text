"""MS-DOC CellN permits TableN+1 before any ParaN; terminators remain required."""
from pathlib import Path

import pytest

from legacy_doc import LegacyDocError, extract_text
from tests.test_extraction_contract import doc_from_paragraphs, properties


def nested_first(depth):
    paragraphs = [
        ('甲\r', properties(depth, cell=True)),
        ('乙\r', properties(depth, cell=True)),
        ('\r', properties(depth, row=True, cells=[0, 0])),
    ]
    if depth == 3:
        paragraphs += [
            ('中\r', properties(2, cell=True)),
            ('\r', properties(2, row=True, cells=[0])),
        ]
    return paragraphs + [
        ('尾\x07', properties()), ('右\x07', properties()),
        ('\x07', properties(row=True, cells=[0, 0])),
    ]


@pytest.mark.parametrize('depth', [2, 3])
@pytest.mark.parametrize('after_row', [False, True])
def test_cell_can_start_with_nested_table(depth, after_row):
    paragraphs = [('正文\r', b'')]
    if after_row:
        paragraphs += [
            ('上一行\x07', properties()),
            ('\x07', properties(row=True, cells=[0])),
        ]
    paragraphs += nested_first(depth) + [('结束\r', b'')]
    expected = '正文\n' + ('上一行\n' if after_row else '')
    expected += '甲\n乙\n' + ('中\n' if depth == 3 else '') + '尾\t右\n结束'
    assert extract_text(doc_from_paragraphs(paragraphs)).text == expected


@pytest.mark.parametrize('damage', [
    'outer-cell', 'outer-row', 'outer-row-at-eof', 'inner-definition', 'middle-cell',
])
def test_nested_first_still_requires_complete_structure(damage):
    paragraphs = nested_first(3)
    if damage == 'outer-cell':
        # Remove both outer cell endings; the row marker cannot close them.
        del paragraphs[5:7]
    elif damage in ('outer-row', 'outer-row-at-eof'):
        paragraphs.pop()
    elif damage == 'inner-definition':
        paragraphs[2] = ('\r', properties(3, row=True, cells=[0]))
    else:
        del paragraphs[3]  # Middle row now lacks its cell mark.
    if damage != 'outer-row-at-eof':
        paragraphs += [('结束\r', b'')]
    with pytest.raises(LegacyDocError, match='unfinished cell|cell or row terminator|definition'):
        extract_text(doc_from_paragraphs(paragraphs))


def test_nested_first_respects_depth_limit():
    with pytest.raises(LegacyDocError, match='nesting|depth'):
        extract_text(doc_from_paragraphs([('深\r', properties(65, cell=True))]))


def test_real_nested_first_document_has_exact_body_output():
    path = Path(__file__).parent / 'data/corpus/libreoffice/fdo53985.doc'
    # Derived from the fixture's paragraphs and cell/row marks: its inner
    # "2nd paragraph" precedes "3rd paragraph" in the first outer cell.
    assert extract_text(path.read_bytes()).text == (
        '1st paragraph\n2nd paragraph\n3rd paragraph\n4th paragraph\n\n'
        '5th paragraph\n6th paragraph.  FORMCHECKBOX  A\n\n'
        '7th paragraph:\n8th paragraph FORMTEXT \u2002\u2002\u2002\u2002\u2002\n9th paragraph\n\n'
        '10th paragraph\n\t FORMCHECKBOX Yes      FORMCHECKBOX No'
    )


def test_real_three_level_document_preserves_nested_content_order():
    path = Path(__file__).parent / 'data/corpus/apache-poi/Bug51890.doc'
    text = extract_text(path.read_bytes()).text
    sections = [
        'Qu’ils s’appellent', 'Pour accéder à un favori',
        'Ajouter une adresse dans les « Favoris »', 'Pour cela, cliquez',
        'Organisation des favoris', 'Pour gérer ces adresses', 'Tri des favoris',
        "Sauvegarde des Favoris d'Internet Explorer",
        "Retrouver la trace d'un site", '\xa0EXERCICE :',
    ]
    positions = [text.index(section) for section in sections]
    assert positions == sorted(positions)
    assert all(text.count(section) == 1 for section in sections)
    # Every top-level row has one cell; the inner columns flatten with LF.
    assert '\t' not in text
    assert text.count('INCLUDEPICTURE') == 5
    assert 'Ie_favo1.jpg' in text and 'icon_ie5.gif' in text
