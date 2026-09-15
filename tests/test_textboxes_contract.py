"""Complete OLE inputs exercising body-anchor placement and XML traversal policy."""
import struct

import pytest

from legacy_doc import LegacyDocError, extract_text
from tests.fixtures import make_ole_file, make_word_streams
from tests.test_extraction_contract import properties
from tests.test_shapes import _ftxbxs, _office_art, _plc, _sp, _spa


def textbox_document(body='前\x08后\r', box='框一\r框二\r', *, table=False, broken=False,
                     anchor_cp=1, character_properties=None):
    paragraph_properties = None
    if table:
        body = '前\x08后\x07右\x07\x07\r'
        paragraph_properties = {
            3: properties(), 5: properties(),
            6: properties(row=True, cells=[0, 0]),
        }
    # A last reusable slot is stored after the referenced textbox range.
    streams = make_word_streams(body, stories={'txbx': box + '\r'},
                                paragraph_properties=paragraph_properties,
                                character_properties=character_properties)
    word = bytearray(streams['WordDocument'])
    contents = bytearray(streams['0Table'])
    body_limit = len(body.encode('utf-16le')) // 2
    box_limit = len(box.encode('utf-16le')) // 2
    records = {
        40: _plc([_spa(42)], [anchor_cp, max(body_limit, anchor_cp + 1)]),
        56: _plc([_ftxbxs(shape_id=42), _ftxbxs(shape_id=0, reusable=True)],
                 [0, box_limit, box_limit + 1]),
        50: _office_art(_sp(42, textbox_index=2 if broken else 1)),
    }
    for index, data in records.items():
        struct.pack_into('<II', word, 154 + index * 8, len(contents), len(data))
        contents.extend(data)
    streams['WordDocument'] = bytes(word)
    streams['0Table'] = bytes(contents)
    return make_ole_file(streams)


def test_body_textbox_inserts_at_anchor_without_paragraph_delimiters():
    assert extract_text(textbox_document()).text == '前框一框二后'


def test_cell_textbox_repeats_descendant_paragraphs_like_reference_xml():
    assert extract_text(textbox_document(table=True)).text == '前框一框二后\n框一\n框二\t右'


def test_box_paragraph_whitespace_is_trimmed_only_at_parent_boundary():
    assert extract_text(textbox_document(box=' 框一 \r\r 框二 \r')).text == '前 框一  框二 后'


def test_body_textbox_cp_coordinates_follow_utf16_units():
    assert extract_text(textbox_document(box='😀𠀀\r尾\r')).text == '前😀𠀀尾后'


def test_broken_selected_textbox_is_failure_instead_of_partial_body():
    with pytest.raises(LegacyDocError, match='lTxid'):
        extract_text(textbox_document(broken=True))


@pytest.mark.parametrize(('body', 'anchor_cp', 'character_properties'), [
    ('前后\r', 3, None),  # The body end is outside the actual text traversal.
    ('前后\r', 2, None),  # A paragraph terminator is not an inline character.
    ('前后\r', 1, None),
    ('前\x08后\r', 1, {1: b'\x55\x08\x00'}),  # CFSpec is explicitly clear.
])
def test_selected_textbox_requires_a_consumable_special_anchor(body, anchor_cp, character_properties):
    with pytest.raises(LegacyDocError, match='Textbox anchor'):
        extract_text(textbox_document(body=body, anchor_cp=anchor_cp,
                                      character_properties=character_properties))


@pytest.mark.parametrize(('body', 'anchor_cp', 'expected'), [
    ('\x08后\r', 0, '框一框二后'),
    ('前\x08\r', 1, '前框一框二'),
])
def test_textbox_anchor_at_either_body_edge_still_extracts(body, anchor_cp, expected):
    assert extract_text(textbox_document(body=body, anchor_cp=anchor_cp)).text == expected


def test_linked_textbox_ranges_are_inserted_at_each_anchor_once():
    from tests.test_shapes import _tbkd

    streams = make_word_streams('前\x08中\x08后\r', stories={'txbx': '框一\r框二\r\r'})
    word = bytearray(streams['WordDocument'])
    table = bytearray(streams['0Table'])
    records = {
        40: _plc([_spa(42), _spa(43)], [1, 3, 6]),
        56: _plc([_ftxbxs(shape_id=42, chain_count=2), _ftxbxs(shape_id=0, reusable=True)],
                 [0, 6, 7]),
        50: _office_art(_sp(42), _sp(43, anchor_index=1, chain_index=1)),
        75: _plc([_tbkd(textbox_index=0), _tbkd(textbox_index=0), _tbkd(textbox_index=99)],
                 [0, 3, 6, 7]),
    }
    for index, data in records.items():
        struct.pack_into('<II', word, 154 + index * 8, len(table), len(data))
        table.extend(data)
    streams['WordDocument'] = bytes(word)
    streams['0Table'] = bytes(table)
    assert extract_text(make_ole_file(streams)).text == '前框一中框二后'
