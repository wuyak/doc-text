"""Symbol placeholders must not be mistaken for literal opening parentheses."""
import struct

import pytest

from legacy_doc import LegacyDocError, extract_text
from tests.fixtures import make_doc_bytes, make_ole_file, make_word_streams, sprm


def test_literal_parentheses_are_preserved():
    assert extract_text(make_doc_bytes('甲(乙)')).text == '甲(乙)'


def test_special_parenthesis_is_a_symbol_element_not_literal_text():
    data = make_doc_bytes('甲(乙)', character_properties={1: sprm(0x0855, 1)})
    assert extract_text(data).text == '甲乙)'


def test_normal_parenthesis_after_supplementary_unicode_uses_cp_coordinates():
    data = make_doc_bytes('😀(甲(乙)', character_properties={2: sprm(0x0855, 1)})
    assert extract_text(data).text == '😀甲(乙)'


def test_complex_prm_can_override_character_special_property():
    streams = make_word_streams('(字', character_properties={0: sprm(0x0855, 1)})
    word = bytearray(streams['WordDocument'])
    table = bytearray(streams['0Table'])
    fc, size = struct.unpack_from('<II', word, 154 + 33 * 8)
    clx = bytearray(table[fc:fc + size])
    # Last two bytes of the one-piece PCD are Prm1: first RgPrc entry.
    struct.pack_into('<H', clx, len(clx) - 2, 1)
    group = sprm(0x0855, 0)
    new_clx = b'\1' + struct.pack('<H', len(group)) + group + clx
    struct.pack_into('<II', word, 154 + 33 * 8, len(table), len(new_clx))
    table.extend(new_clx)
    streams.update(WordDocument=bytes(word), **{'0Table': bytes(table)})
    assert extract_text(make_ole_file(streams)).text == '(字'


def test_selected_character_property_damage_is_not_silently_ignored():
    streams = make_word_streams('(字', character_properties={0: sprm(0x0855, 1)})
    word = bytearray(streams['WordDocument'])
    fc, size = struct.unpack_from('<II', word, 154 + 12 * 8)
    table = bytearray(streams['0Table'])
    count = (size - 4) // 8
    struct.pack_into('<I', table, fc + (count + 1) * 4, 0x3FFFFF)
    streams['0Table'] = bytes(table)
    with pytest.raises(LegacyDocError, match='CHPX FKP'):
        extract_text(make_ole_file(streams))


def test_symbol_metadata_does_not_implicitly_override_cfspec():
    data = make_doc_bytes('甲(乙)', character_properties={1: sprm(0x6A09, b'\0\0\x00\xf0')})
    assert extract_text(data).text == '甲(乙)'


@pytest.mark.parametrize('value', [0x80, 0x81])
def test_style_dependent_symbol_property_is_explicitly_unsupported(value):
    data = make_doc_bytes('(字', character_properties={0: sprm(0x0855, value)})
    with pytest.raises(LegacyDocError, match='Unsupported style-dependent'):
        extract_text(data)


def test_prm0_cfspec_can_mark_a_symbol_without_chpx():
    streams = make_word_streams('(字')
    word = streams['WordDocument']
    table = bytearray(streams['0Table'])
    fc, size = struct.unpack_from('<II', word, 154 + 33 * 8)
    struct.pack_into('<H', table, fc + size - 2, (1 << 8) | (0x75 << 1))
    streams['0Table'] = bytes(table)
    assert extract_text(make_ole_file(streams)).text == '字'


def test_revision_deleted_and_hidden_character_flags_do_not_filter_text():
    data = make_doc_bytes('旧新隐藏', character_properties={
        0: sprm(0x0800, 1),  # sprmCFRMarkDel
        2: sprm(0x083C, 1),  # sprmCFVanish
        3: sprm(0x083C, 1),
    })
    assert extract_text(data).text == '旧新隐藏'
