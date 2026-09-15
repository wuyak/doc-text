"""Names require a known metadata representation; payloads remain opaque."""
import struct

import pytest

from legacy_doc._object_metadata import MAX_METADATA_BYTES, object_label
from legacy_doc.exceptions import LegacyDocError


class PrefixReader:
    def __init__(self, streams):
        self.streams = streams
        self.reads = []

    def try_read_storage_stream_prefix(self, storage, name, max_size):
        assert storage == 'selected object'
        self.reads.append((name, max_size))
        return self.streams[name][:max_size] if name in self.streams else None


def _lp(value):
    return struct.pack('<I', len(value) + 1) + value + b'\0'


def _comp_obj(*, ansi_type='', unicode_type='', reserved=True, clipboard=0):
    raw = b'\0'*28 + _lp(ansi_type.encode('ascii')) + struct.pack('<I', clipboard)
    if clipboard in (0xFFFFFFFF, 0xFFFFFFFE):
        raw += struct.pack('<I', 3)  # CF_METAFILEPICT
    raw += _lp(b'Ignored') if reserved else struct.pack('<I', 0)
    value = unicode_type.encode('utf-16le') + b'\0\0'
    return raw + struct.pack('<II', 0x71B239F4, len(value)) + value


def _ole10_native(filename, *, payload=b''):
    body = b'\x02\0' + b'label\0' + filename + b'\0' + struct.pack('<HH', 0, 3)
    body += b'command\0' + struct.pack('<I', len(payload)) + payload
    return struct.pack('<I', len(body)) + body


def label(streams):
    reader = PrefixReader(streams)
    return object_label(reader, 'selected object'), reader.reads


def test_filename_uses_package_header_without_path_or_payload():
    value, reads = label({'\x01CompObj': _comp_obj(ansi_type='Package'),
                          '\x01Ole10Native': _ole10_native(b'C:\\temp\\quote.xls', payload=b'x'*20000)})
    assert value == '[嵌入文件：quote.xls]'
    assert reads == [('\x01CompObj', MAX_METADATA_BYTES), ('\x01Ole10Native', MAX_METADATA_BYTES)]


@pytest.mark.parametrize('clipboard', [0, 0xFFFFFFFF, 0xFFFFFFFE])
def test_unicode_type_and_clipboard_layout(clipboard):
    value, _ = label({'\x01CompObj': _comp_obj(ansi_type='Fallback', unicode_type='中文对象类型', clipboard=clipboard)})
    assert value == '[嵌入文件：中文对象类型]'


def test_zero_reserved1_requires_ignoring_unicode_tail():
    value, _ = label({'\x01CompObj': _comp_obj(ansi_type='Fallback', unicode_type='Must be ignored', reserved=False)})
    assert value == '[嵌入文件：Fallback]'


@pytest.mark.parametrize('filename', [b'without-suffix', '报价表.xls'.encode(), b'bad\xff.xls', b'bad\nname.xls'])
def test_ambiguous_encoding_or_missing_suffix_does_not_invent_filename(filename):
    value, _ = label({'\x01CompObj': _comp_obj(ansi_type='Package'), '\x01Ole10Native': _ole10_native(filename)})
    assert value == '[嵌入文件：Package]'


def test_unknown_native_format_is_not_interpreted_as_package():
    value, reads = label({'\x01CompObj': _comp_obj(ansi_type='An application'),
                          '\x01Ole10Native': _ole10_native(b'not-a-real-name.xls')})
    assert value == '[嵌入文件：An application]'
    assert reads == [('\x01CompObj', MAX_METADATA_BYTES)]


@pytest.mark.parametrize('raw', [b'', b'\0'*28 + struct.pack('<I', 0xFFFFFFFF), _comp_obj(ansi_type='bad\ntype')])
def test_malformed_metadata_has_generic_fallback(raw):
    assert label({'\x01CompObj': raw})[0] == '[嵌入文件]'


def test_missing_metadata_has_generic_fallback():
    assert label({})[0] == '[嵌入文件]'


def test_truncated_native_header_does_not_yield_a_name():
    raw = _ole10_native(b'quote.xls')
    assert label({'\x01CompObj': _comp_obj(ansi_type='Package'), '\x01Ole10Native': raw[:-5]})[0] == '[嵌入文件：Package]'


def test_optional_stream_error_falls_back_but_programming_error_surfaces():
    class Broken(PrefixReader):
        def try_read_storage_stream_prefix(self, *args, **kwargs):
            raise LegacyDocError('bad optional stream')
    assert object_label(Broken({}), 'selected object') == '[嵌入文件]'
    class Bug(PrefixReader):
        def try_read_storage_stream_prefix(self, *args, **kwargs):
            raise TypeError('programming error')
    with pytest.raises(TypeError):
        object_label(Bug({}), 'selected object')
