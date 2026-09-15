"""Embedded objects replace only their indexed field, without reading attachments."""
from pathlib import Path
import struct

import pytest

from legacy_doc import ExtractionOptions, LegacyDocError, extract_text
from legacy_doc.ole import OleReader
from tests.fixtures import make_ole_file, make_word_streams, sprm
from tests.test_extraction_contract import properties


def add_fields(streams, records, *, pair=16, final_cp=None):
    word = bytearray(streams['WordDocument'])
    table = bytearray(streams['0Table'])
    cps = [cp for cp, _, _ in records]
    final_cp = final_cp if final_cp is not None else cps[-1] + 1
    payload = struct.pack('<' + 'I' * (len(cps) + 1), *cps, final_cp)
    payload += bytes(value for _, ch, flags in records for value in (ch, flags))
    struct.pack_into('<II', word, 154 + pair * 8, len(table), len(payload))
    table.extend(payload)
    return {**streams, 'WordDocument': bytes(word), '0Table': bytes(table)}


def indexed_doc(text, *, kind=0x3A, char_props=None, para_props=None, end_flags=0x80):
    cps = []
    cp = 0
    for char in text:
        if char in '\x13\x14\x15':
            cps.append((cp, ord(char), kind if char == '\x13' else end_flags if char == '\x15' else 255))
        cp += len(char.encode('utf-16le')) // 2
    return make_ole_file(add_fields(make_word_streams(
        text, character_properties=char_props, paragraph_properties=para_props), cps))


def test_missing_object_metadata_leaves_one_placeholder_in_position():
    result = extract_text(indexed_doc('前😀\x13 EMBED Unknown.Type \x14\x01\x15后\r'))
    assert result.text == '前😀[嵌入文件]后'
    assert result.warnings == ()


def test_embedded_field_in_cell_keeps_neighbor_and_row_position():
    text = '前\x13 EMBED Unknown \x14\x01\x15后\x07右\x07\x07\r'
    props = {i: properties(row=(text[i-1] == '\x07'), cells=[0, 0] if text[i-1] == '\x07' else None)
             for i, char in enumerate(text) if char == '\x07'}
    assert extract_text(indexed_doc(text, para_props=props)).text == '前[嵌入文件]后\t右'


@pytest.mark.parametrize('kind', [0x38, 0x57, 0x21])
def test_external_link_control_and_ordinary_field_are_not_embedded_files(kind):
    data = indexed_doc('前\x13 EMBED is literal here \x14结果\x15后\r', kind=kind)
    assert extract_text(data).text == '前 EMBED is literal here 结果后'


def test_unindexed_literal_embed_word_and_picture_do_not_create_placeholder():
    assert extract_text(make_ole_file(make_word_streams('EMBED Excel.Sheet.8 图片\x01'))).text == 'EMBED Excel.Sheet.8 图片'


def test_repeated_object_positions_are_not_deduplicated():
    data = indexed_doc('\x13 EMBED X\x14\x01\x15及\x13 EMBED X\x14\x01\x15')
    assert extract_text(data).text == '[嵌入文件]及[嵌入文件]'


def test_field_inside_opaque_embedded_result_is_not_rendered_twice():
    data = indexed_doc('前\x13 EMBED X\x14\x13 EMBED X\x14\x01\x15\x15后')
    assert extract_text(data).text == '前[嵌入文件]后'



def test_embed_nested_in_ordinary_field_preserves_surrounding_saved_text():
    text = '前\x13 IF condition \x14甲\x13 EMBED X\x14\x01\x15乙\x15后'
    records = []
    for cp, char in enumerate(text):
        if char in '\x13\x14\x15':
            kind = 7 if cp == 1 else 58
            records.append((cp, ord(char), kind if char == '\x13' else 128))
    data = make_ole_file(add_fields(make_word_streams(text), records))
    assert extract_text(data).text == '前 IF condition 甲[嵌入文件]乙后'


def test_outer_ordinary_field_marker_must_match_text_around_embed():
    text = '前\x13 IF condition \x14甲\x13 EMBED X\x14\x01\x15乙\x15后'
    records = []
    for cp, char in enumerate(text):
        if char in '\x13\x14\x15':
            kind = 7 if cp == 1 else 58
            records.append((cp, ord(char), kind if char == '\x13' else 128))
    cp, ch, flags = records[1]
    records[1] = (cp - 1, ch, flags)  # Points at a space, not the IF separator.
    with pytest.raises(LegacyDocError, match='Plcfld marker'):
        extract_text(make_ole_file(add_fields(make_word_streams(text), records)))

def test_field_type_without_storage_is_enough_for_generic_placeholder():
    assert extract_text(indexed_doc('前\x13 EMBED X\x15后')).text == '前[嵌入文件]后'


def test_zombie_object_does_not_dereference_location(monkeypatch):
    text = '\x13 EMBED X\x14\x01\x15'
    cp = text.index('\x14')
    data = indexed_doc(text, char_props={cp: sprm(0x0855, 1) + sprm(0x080A, 1)
                                       + sprm(0x0856, 1) + sprm(0x6A03, 123)}, end_flags=0x82)
    original = OleReader.find_storage
    def no_storage(self, path):
        assert len(path) == 1, 'Zombie location must be ignored'
        return original(self, path)
    monkeypatch.setattr(OleReader, 'find_storage', no_storage)
    assert extract_text(data).text == '[嵌入文件]'


@pytest.mark.parametrize('damage', ['marker', 'special', 'order', 'unclosed'])
def test_damaged_object_association_is_rejected(damage):
    streams = make_word_streams('\x13 EMBED X\x14\x01\x15')
    records = [(0, 19, 58), (9, 20, 255), (11, 21, 128)]
    if damage == 'marker':
        records[1] = (8, 20, 255)
    if damage == 'special':
        streams = make_word_streams('\x13 EMBED X\x14\x01\x15', character_properties={0: sprm(0x0855, 0)})
    if damage == 'order':
        records[1] = (0, 20, 255)
    if damage == 'unclosed':
        records.pop()
    with pytest.raises(LegacyDocError, match='field|Plcfld'):
        extract_text(make_ole_file(add_fields(streams, records)))


def test_placeholder_obeys_final_utf8_output_limit():
    with pytest.raises(LegacyDocError, match='extracted text exceeds parser limit'):
        extract_text(indexed_doc('\x13 EMBED X\x14\x01\x15'), options=ExtractionOptions(max_text_bytes=5))


def test_unused_textbox_fields_do_not_affect_body():
    streams = make_word_streams('正文', stories={'txbx': '\x13坏\r'})
    streams = add_fields(streams, [(0, 19, 58)], pair=57)
    assert extract_text(make_ole_file(streams)).text == '正文'


def test_selected_textbox_object_uses_story_relative_field_coordinates():
    from legacy_doc._binary import BinaryDocument
    from legacy_doc._shapes import TextboxRange
    from legacy_doc._text import TextRenderer
    body = '前\x08后\r'
    box = '\x13 EMBED X\x14\x01\x15\r'
    streams = make_word_streams(body, stories={'txbx': box})
    streams = add_fields(streams, [(0, 19, 58), (9, 20, 255), (11, 21, 128)], pair=57, final_cp=1000)
    doc = BinaryDocument(OleReader(make_ole_file(streams), options=ExtractionOptions()))
    # Association is already tested separately; render an explicit selected box.
    renderer = TextRenderer(doc, {1: TextboxRange(len(body), len(body) + len(box))})
    assert renderer.render() == '前[嵌入文件]后'


def test_selection_cannot_hide_enclosing_object_by_starting_in_nested_field():
    from legacy_doc._binary import BinaryDocument
    from legacy_doc._characters import CharacterIndex
    from legacy_doc._objects import EmbeddedObjects
    data = indexed_doc('\x13 EMBED X\x14\x13 EMBED X\x14\x01\x15\x15')
    doc = BinaryDocument(OleReader(data, options=ExtractionOptions()))
    objects = EmbeddedObjects(doc, CharacterIndex(doc))
    with pytest.raises(LegacyDocError, match='splits an embedded field'):
        objects.selected(12, doc.fib.ccp_text)


CORPUS = Path(__file__).parent / 'data' / 'corpus' / 'apache-poi'


def test_real_embedded_word_payloads_are_not_read(monkeypatch):
    original = OleReader._read_entry
    reads = []
    def guarded(self, entry, *args, **kwargs):
        reads.append(entry.name)
        if entry.name in ('WordDocument', '0Table', '1Table', 'Data'):
            assert entry in self.storage_children(self.root_entry), 'Attempted to read embedded document payload'
        return original(self, entry, *args, **kwargs)
    monkeypatch.setattr(OleReader, '_read_entry', guarded)
    result = extract_text((CORPUS / 'Bug47731.doc').read_bytes())
    assert result.text.count('[嵌入文件') == 4
    assert 'EMBED' not in result.text
    assert 'rice' not in result.text.lower()
    assert reads.count('WordDocument') == 1
    assert result.warnings == ()


def test_real_mixed_attachments_keep_main_text_and_order():
    result = extract_text((CORPUS / 'word_with_embeded.doc').read_bytes())
    assert result.text == '\n'.join([
        'I have lots of embedded files in me',
        '[嵌入文件：Microsoft Office Excel 2003 Worksheet]',
        '[嵌入文件：Microsoft Office Excel 2003 Worksheet]',
        '[嵌入文件：Microsoft Office PowerPoint 97-2003 Presentation]',
        '[嵌入文件：Microsoft Office Word 97-2003 Document]',
    ])
    assert result.metadata['has_embedded_objects'] is True
    assert result.warnings == ()


@pytest.mark.parametrize(('name', 'begin', 'end'), [
    ('53379.doc', '[嵌入文件：Microsoft Drawing]\nMEDICINES CONTROL AGENCY\nClinical Trials Unit',
     '[Provide a flowchart listing the safety parameters to be measured along with the timepoints at which these are considered.]'),
    ('Bug50936_3.doc', 'Développement d’un moteur de recherche d’informations sur Internet.',
     'code source (pour les projets d’informatique)'),
])
def test_real_root_table_selection_keeps_main_body(name, begin, end):
    text = extract_text((CORPUS / name).read_bytes()).text
    assert text.startswith(begin)
    assert text.endswith(end)
    assert 'EMBED ' not in text


def test_real_high_length_bits_do_not_hide_valid_body():
    text = extract_text((CORPUS / '53446.doc').read_bytes()).text
    assert text.startswith('DRAFT\n\nProposal\nto\nLakeland Electric\nby\nEnron North America Corp.')
    assert 'Scope of Agreement' in text
    assert text.endswith('Lakegas2326.doc\t09/1728/99')


def test_root_fix_does_not_mask_independent_textbox_damage():
    with pytest.raises(LegacyDocError, match='Body textbox shape is missing OfficeArtClientTextbox'):
        extract_text((CORPUS / 'Bug61268.doc').read_bytes())


def object_document(metadata_streams):
    """A real CFB: root DOC plus one ObjectPool child with opaque payload."""
    text = '前\x13 EMBED Package\x14\x01\x15后'
    sep = text.index('\x14')
    props = {sep: sprm(0x0855, 1) + sprm(0x080A, 1) + sprm(0x0856, 1) + sprm(0x6A03, 42)}
    streams = add_fields(make_word_streams(text, character_properties=props),
                         [(1, 19, 58), (sep, 20, 255), (sep + 2, 21, 128)])
    streams.update({'ObjectPool': b'', '_42': b'', 'Payload': b'NOT TO BE PARSED', **metadata_streams})
    data = bytearray(make_ole_file(streams))
    reader = OleReader(bytes(data), options=ExtractionOptions())
    directory_sectors = []
    sector = reader.first_dir_sector
    while sector != 0xFFFFFFFE:
        directory_sectors.append(sector)
        sector = reader.fat[sector]
    offsets = {entry.name: 512 + directory_sectors[entry.directory_id // 4] * 512
               + (entry.directory_id % 4) * 128 for entry in reader.directory}
    ids = {entry.name: entry.directory_id for entry in reader.directory}
    positions = {entry.directory_id: offsets[entry.name] for entry in reader.directory}
    for position in positions.values():
        struct.pack_into('<III', data, position + 68, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF)
    for name in ('ObjectPool', '_42'):
        data[offsets[name] + 66] = 1
        struct.pack_into('<IQ', data, offsets[name] + 116, 0xFFFFFFFE, 0)
    def children(owner, names):
        names = sorted(names, key=lambda name: (len(name), name.upper()))
        deepest = len(names).bit_length() - 1
        def tree(names, depth=0):
            if not names:
                return 0xFFFFFFFF
            mid = len(names) // 2
            name = names[mid]
            struct.pack_into('<II', data, offsets[name] + 68,
                             tree(names[:mid], depth + 1), tree(names[mid+1:], depth + 1))
            data[offsets[name] + 67] = 0 if depth == deepest and depth else 1
            return ids[name]
        struct.pack_into('<I', data, offsets[owner] + 76, tree(names))
    children('Root Entry', ['WordDocument', '0Table', 'ObjectPool'])
    children('ObjectPool', ['_42'])
    children('_42', ['Payload', *metadata_streams])
    return bytes(data)


def test_resolved_object_type_is_inserted_via_real_storage_path():
    from tests.test_object_metadata import _comp_obj
    result = extract_text(object_document({'\x01CompObj': _comp_obj(ansi_type='Example document')}))
    assert result.text == '前[嵌入文件：Example document]后'
    assert result.warnings == ()


def test_broken_optional_object_directory_keeps_generic_placeholder():
    data = bytearray(object_document({}))
    # The ObjectPool storage is directory entry 3 in this fixture.
    struct.pack_into('<I', data, 512 + 3 * 128 + 76, 0x12345678)
    result = extract_text(bytes(data))
    assert result.text == '前[嵌入文件]后'
    assert result.warnings == ()


def test_named_package_in_body_uses_metadata_prefix_only(monkeypatch):
    from tests.test_object_metadata import _comp_obj, _ole10_native
    data = object_document({'\x01CompObj': _comp_obj(ansi_type='Package'),
                            '\x01Ole10Native': _ole10_native(b'C:\\temp\\quote.xls', payload=b'x'*8192)})
    original = OleReader._read_entry
    def no_payload(self, entry, *args, **kwargs):
        assert entry.name not in ('Payload', '\x01Ole10Native')
        return original(self, entry, *args, **kwargs)
    monkeypatch.setattr(OleReader, '_read_entry', no_payload)
    result = extract_text(data)
    assert result.text == '前[嵌入文件：quote.xls]后'
    assert result.warnings == ()
