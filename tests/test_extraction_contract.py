"""Expected outputs are the agreed extraction contract, not implementation snapshots."""
import struct

import pytest

from legacy_doc import ExtractionOptions, LegacyDocError, extract_text
from tests.fixtures import make_doc_bytes, make_ole_file, make_word_streams, sprm


def properties(depth=1, *, row=False, cell=False, cells=None):
    result = sprm(0x2416, 1) + sprm(0x6649, depth)
    if depth == 1 and row:
        result += sprm(0x2417, 1)
    if depth > 1:
        if row:
            result += sprm(0x244C, 1)
        elif cell:
            result += sprm(0x244B, 1)
    if cells is not None:
        widths = struct.pack('<' + 'h' * (len(cells) + 1), *range(0, (len(cells) + 1) * 100, 100))
        tc80 = b''.join(struct.pack('<H', flag) + b'\0' * 18 for flag in cells)
        result += sprm(0xD608, bytes([len(cells)]) + widths + tc80)
    return result


def doc_from_paragraphs(paragraphs):
    """Each item is an explicit (paragraph with terminator, grpprl) pair."""
    text = ''
    props = {}
    cp = 0
    for paragraph, group in paragraphs:
        text += paragraph
        cp += len(paragraph.encode('utf-16le')) // 2
        props[cp - 1] = group
    return make_doc_bytes(text, paragraph_properties=props)


def test_default_body_excludes_independent_stories():
    data = make_doc_bytes('正文', stories={name: value + '\r' for name, value in {
        'ftn': '脚注', 'hdd': '页眉', 'atn': '批注', 'edn': '尾注'}.items()})
    assert extract_text(data).text == '正文'


def test_empty_body_is_a_success_with_metadata():
    result = extract_text(make_doc_bytes('', stories={'hdd': '页眉\r'}))
    assert result.text == ''
    assert result.metadata['chars'] == result.metadata['bytes'] == 0
    assert result.warnings == ()


def test_unicode_and_main_paragraph_whitespace_survive():
    assert extract_text(make_doc_bytes('  A😀  B\r\rC\tD\r')).text == 'A😀  B\n\nC\tD'


def test_field_controls_are_consumed_but_saved_text_is_kept():
    assert extract_text(make_doc_bytes('前\x13 PAGE \\* MERGEFORMAT \x143\x15后\r')).text == '前 PAGE \\* MERGEFORMAT 3后'


def test_two_rows_and_surrounding_body_keep_order():
    data = doc_from_paragraphs([
        ('名单\r', b''),
        ('姓名\x07', properties()), ('年龄\x07', properties()), ('\x07', properties(row=True, cells=[0, 0])),
        ('张三\x07', properties()), ('28\x07', properties()), ('\x07', properties(row=True, cells=[0, 0])),
        ('结束\r', b''),
    ])
    assert extract_text(data).text == '名单\n姓名\t年龄\n张三\t28\n结束'


def test_empty_cells_and_cell_paragraphs_match_docx_rule():
    data = doc_from_paragraphs([
        ('\x07', properties()),
        ('A\r', properties()), ('\r', properties()), ('B\x07', properties()),
        ('\x07', properties()), ('D\x07', properties()), ('\x07', properties()),
        ('\x07', properties(row=True, cells=[0] * 5)), ('\r', b''),
    ])
    assert extract_text(data).text == '\tA\nB\t\tD'


def test_nested_table_flattens_paragraphs_inside_parent_cell():
    data = doc_from_paragraphs([
        ('前\r', properties()),
        ('甲\r', properties(2, cell=True)), ('乙\r', properties(2, cell=True)),
        ('\r', properties(2, row=True, cells=[0, 0])),
        ('后\x07', properties()), ('右\x07', properties()),
        ('\x07', properties(row=True, cells=[0, 0])), ('\r', b''),
    ])
    assert extract_text(data).text == '前\n甲\n乙\n后\t右'


def test_horizontal_merge_projects_one_logical_cell():
    data = doc_from_paragraphs([
        ('A\x07', properties()), ('\x07', properties()), ('B\x07', properties()),
        ('\x07', properties(row=True, cells=[2, 1, 0])), ('\r', b''),
    ])
    assert extract_text(data).text == 'A\tB'


def test_row_definition_mismatch_is_not_silently_repaired():
    data = doc_from_paragraphs([
        ('A\x07', properties()), ('\x07', properties(row=True, cells=[0, 0])), ('\r', b''),
    ])
    with pytest.raises(LegacyDocError, match='cell|Cell|table'):
        extract_text(data)


def test_cell_mark_without_table_membership_is_not_a_plain_line_break():
    with pytest.raises(LegacyDocError, match='cell mark.*table'):
        extract_text(make_doc_bytes('A\x07B\r'))


def test_unselected_story_cell_marks_do_not_affect_body():
    data = make_doc_bytes('正文\r', stories={'hdd': 'A\x07B\r'})
    assert extract_text(data).text == '正文'


def test_literal_page_field_text_is_not_deleted():
    assert extract_text(make_doc_bytes('说明\rPAGE \\* MERGEFORMAT 3\r结束')).text == '说明\nPAGE \\* MERGEFORMAT 3\n结束'


def test_private_use_text_is_not_discarded():
    assert extract_text(make_doc_bytes('A\ue123\ufeffB')).text == 'A\ue123\ufeffB'


def test_reserved_fib_fields_do_not_control_text_range():
    streams = make_word_streams('准确正文')
    word = bytearray(streams['WordDocument'])
    struct.pack_into('<II', word, 0x18, 0xFFFFFFFF, 0)
    streams['WordDocument'] = bytes(word)
    assert extract_text(make_ole_file(streams)).text == '准确正文'


def test_utf8_output_budget_and_empty_zero_budget():
    assert extract_text(make_doc_bytes(''), options=ExtractionOptions(max_text_bytes=0)).text == ''
    with pytest.raises(LegacyDocError, match='extracted text exceeds'):
        extract_text(make_doc_bytes('中文'), options=ExtractionOptions(max_text_bytes=5))


def test_known_content_doc_export():
    from pathlib import Path

    data = (Path(__file__).parent / 'data' / 'empty-multipara-emoji.doc').read_bytes()
    assert extract_text(data).text == (
        '正文含 HEADER_ONLY 相同文字应保留 😀𠀀\n'
        '甲\t\t丙\n'
        '\t第二行\n单元格第二段\n\n'
        '正文结尾 BODY_END 😀'
    )


def test_three_table_levels_preserve_stored_paragraph_order():
    data = doc_from_paragraphs([
        ('一前\r', properties()), ('二前\r', properties(2)),
        ('甲\r', properties(3, cell=True)), ('乙\r', properties(3, cell=True)),
        ('\r', properties(3, row=True, cells=[0, 0])),
        ('二后\r', properties(2, cell=True)),
        ('\r', properties(2, row=True, cells=[0])),
        ('一后\x07', properties()), ('右\x07', properties()),
        ('\x07', properties(row=True, cells=[0, 0])), ('\r', b''),
    ])
    assert extract_text(data).text == '一前\n二前\n甲\n乙\n二后\n一后\t右'


def test_vertical_merge_continuation_does_not_repeat_previous_text():
    data = doc_from_paragraphs([
        ('上\x07', properties()), ('右上\x07', properties()),
        ('\x07', properties(row=True, cells=[3 << 5, 0])),
        ('\x07', properties()), ('右下\x07', properties()),
        ('\x07', properties(row=True, cells=[1 << 5, 0])), ('\r', b''),
    ])
    assert extract_text(data).text == '上\t右上\n\t右下'


def test_rows_may_have_different_cell_counts_and_tables_may_be_separate():
    data = doc_from_paragraphs([
        ('A\x07', properties()), ('B\x07', properties()),
        ('\x07', properties(row=True, cells=[0, 0])),
        ('C\x07', properties()), ('\x07', properties(row=True, cells=[0])),
        ('\r', b''),
        ('D\x07', properties()), ('\x07', properties(row=True, cells=[0])),
        ('\r', b''),
    ])
    assert extract_text(data).text == 'A\tB\nC\n\nD'


def test_manual_page_break_preserves_spaces_inside_paragraph():
    assert extract_text(make_doc_bytes('A \x0c B\r')).text == 'A \n B'


def test_section_break_ends_the_paragraph():
    assert extract_text(make_doc_bytes('A \x0c B\r', section_break_cps={2})).text == 'A\nB'


def test_output_limit_applies_after_table_row_trailing_whitespace_is_removed():
    blank = doc_from_paragraphs([
        ('\t\x07', properties()), ('\x07', properties(row=True, cells=[0])), ('\r', b''),
    ])
    assert extract_text(blank, options=ExtractionOptions(max_text_bytes=0)).text == ''
    text = doc_from_paragraphs([
        ('ABC\u00a0\x07', properties()), ('\x07', properties(row=True, cells=[0])), ('\r', b''),
    ])
    assert extract_text(text, options=ExtractionOptions(max_text_bytes=3)).text == 'ABC'


def test_encryption_is_reported_before_parsing_protected_fib_fields():
    streams = make_word_streams('protected', encrypted=True)
    word = bytearray(streams['WordDocument'])
    word[32:900] = b'\xff' * 868
    streams['WordDocument'] = bytes(word)
    with pytest.raises(LegacyDocError, match='Encrypted'):
        extract_text(make_ole_file(streams))
