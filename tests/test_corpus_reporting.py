"""Regression tests for the corpus runner's rejection categories."""

from __future__ import annotations

import pytest

import legacy_doc
from scripts.check_corpus import extract_one

OLE_PREFIX = bytes.fromhex('d0cf11e0a1b11ae1')


@pytest.mark.parametrize(
    'message',
    [
        '.doc extracted text exceeds parser limit',
        'OLE root storage exceeds parser file-size limit',
        'OLE MiniFAT exceeds parser file-size limit',
        'OLE storage path exceeds parser limit',
        'OLE DIFAT chain exceeds parser limit',
        'OLE FAT chain exceeds parser limit',
        'OLE MiniFAT chain exceeds parser limit',
        'OLE stream exceeds parser limit',
        "OLE stream 'WordDocument' exceeds parser limit",
        'OLE directory tree exceeds parser limit',
        'Extracted Word text exceeds parser work limit',
        'Extracted Word text exceeds parser input limit',
        'DOC text traversal exceeds parser work limit',
        'DOC field index exceeds parser limit',
        'DOC field nesting exceeds parser limit',
        'CHPX index exceeds parser limit',
        'Word CLX contains too many pieces',
        'OfficeArt record count exceeds parser limit',
        'OfficeArt container nesting exceeds parser limit',
        'OfficeArt shape nesting exceeds parser limit',
        'PlcfSpaMom record count exceeds parser limit',
        'PlcftxbxTxt record count exceeds parser limit',
        'PlcfTxbxBkd record count exceeds parser limit',
        'DOC table nesting exceeds parser limit',
        'Legacy .doc paragraph properties exceed the SPRM limit',
        'Legacy .doc indirect paragraph properties exceed the depth limit',
        'Legacy .doc PlcfSed exceeds the section limit',
        'Legacy .doc PlcBtePapx exceeds the page limit',
        'Legacy .doc PAPX entries exceed the processing limit',
    ],
)
def test_injected_resource_rejections_are_classified(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path,
    message: str,
) -> None:
    def reject(_raw: bytes) -> None:
        raise legacy_doc.LegacyDocError(message)

    monkeypatch.setattr(legacy_doc, 'extract_text', reject)
    path = tmp_path / 'resource-limit.doc'
    path.write_bytes(OLE_PREFIX)

    outcome = extract_one(path)

    assert outcome['status'] == 'resource-limit'
    assert outcome['error'] == message


@pytest.mark.parametrize(
    'message',
    [
        'CP range [0, 99) exceeds CP limit 10',
        'WordDocument text reference exceeds cbMac (100 > 20)',
        'CHPX FKP run exceeds its BTE range',
        'OLE stream chain is truncated',
        'Cyclic OLE FAT chain',
        'Invalid Word CLX: bytes follow the Pcdt',
        'Word CLX contains too many Prc records',
        'Invalid Word CLX: Prc grpprl is too large',
        'Textbox CP range exceeds document CP limit',
        'Cyclic or excessive DOC textbox nesting',
        'Legacy .doc table cell range is outside the row',
        'Legacy .doc PChgTabs operand has too many deletions',
        'Legacy .doc PChgTabs operand has too many additions',
        'Legacy .doc TDefTable has too many cells',
    ],
)
def test_bounds_and_structure_rejections_remain_needs_review(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path,
    message: str,
) -> None:
    def reject(_raw: bytes) -> None:
        raise legacy_doc.LegacyDocError(message)

    monkeypatch.setattr(legacy_doc, 'extract_text', reject)
    path = tmp_path / 'needs-review.doc'
    path.write_bytes(OLE_PREFIX)

    outcome = extract_one(path)

    assert outcome['status'] == 'rejected-needs-review'
    assert outcome['error'] == message


def test_extract_one_reports_a_real_default_file_size_rejection(tmp_path) -> None:
    path = tmp_path / 'over-size.doc'
    max_file_bytes = 25 * 1024 * 1024
    path.write_bytes(OLE_PREFIX + b'\0' * (max_file_bytes + 1 - len(OLE_PREFIX)))

    outcome = extract_one(path)

    assert outcome['status'] == 'resource-limit'
    assert outcome['error'] == '.doc file exceeds parser file-size limit'
