"""Run the local DOC corpus with per-file timeouts and explicit outcomes.

This checks parser behavior and records failures; it does not invent expected
text or treat an explicit rejection as proof that a document is malformed.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
import time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / 'tests' / 'data'

# Keep this list tied to the parser's explicit resource guards.  A generic
# ``limit`` match would mislabel CP/FC bounds, chain truncation, and malformed
# structures as capacity failures.  The two patterns below cover messages
# whose resource name is interpolated by the parser.
# Format-level cardinality checks (for example, Prm1 references, PChgTabs,
# and TDefTable) remain review cases even when their implementation also has a
# bounded field width.
_RESOURCE_LIMIT_MESSAGES = frozenset({
    '.doc file exceeds parser file-size limit',
    '.doc extracted text exceeds parser limit',
    'OLE root storage exceeds parser file-size limit',
    'OLE MiniFAT exceeds parser file-size limit',
    'OLE storage path exceeds parser limit',
    'OLE DIFAT chain exceeds parser limit',
    'OLE FAT chain exceeds parser limit',
    'OLE MiniFAT chain exceeds parser limit',
    'OLE stream exceeds parser limit',
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
    'DOC table nesting exceeds parser limit',
    'Legacy .doc paragraph properties exceed the SPRM limit',
    'Legacy .doc indirect paragraph properties exceed the depth limit',
    'Legacy .doc PlcfSed exceeds the section limit',
    'Legacy .doc PlcBtePapx exceeds the page limit',
    'Legacy .doc PAPX entries exceed the processing limit',
})
_RESOURCE_LIMIT_PATTERNS = (
    re.compile(r"^OLE stream '[^'\r\n]+' exceeds parser limit$"),
    re.compile(
        r'^(?:PlcfSpaMom|PlcftxbxTxt|PlcfTxbxBkd) '
        r'record count exceeds parser limit$'
    ),
)


def _is_resource_limit_error(message: str) -> bool:
    """Return whether *message* names a parser capacity guard."""

    return message in _RESOURCE_LIMIT_MESSAGES or any(
        pattern.fullmatch(message) for pattern in _RESOURCE_LIMIT_PATTERNS
    )


def extract_one(path: Path) -> dict:
    from legacy_doc import LegacyDocError, extract_text

    raw = path.read_bytes()
    start = time.monotonic()
    try:
        result = extract_text(raw)
    except LegacyDocError as error:
        message = str(error)
        if raw[:8] != bytes.fromhex('d0cf11e0a1b11ae1'):
            status = 'unsupported-container'
        elif 'Encrypted' in message or 'encrypted' in message:
            status = 'unsupported-encryption'
        elif _is_resource_limit_error(message):
            status = 'resource-limit'
        elif message.startswith('Unsupported Word'):
            status = 'unsupported-word-version'
        elif 'Unsupported' in message or 'not supported' in message:
            status = 'unsupported-feature'
        else:
            status = 'rejected-needs-review'
        return {'status': status, 'error': message, 'seconds': round(time.monotonic() - start, 4)}
    except Exception as error:
        return {'status': 'unexpected-error', 'error': f'{type(error).__name__}: {error}',
                'seconds': round(time.monotonic() - start, 4)}
    return {
        'status': 'extracted', 'chars': len(result.text),
        'text_sha256': hashlib.sha256(result.text.encode()).hexdigest(),
        'preview': result.text[:160], 'warnings': list(result.warnings),
        'seconds': round(time.monotonic() - start, 4),
    }


def check(case: dict, timeout: float) -> dict:
    path = DATA / case['file']
    result = {'file': case['file'], 'category': case.get('category', 'unclassified'),
              'source': case['source']}
    if hashlib.sha256(path.read_bytes()).hexdigest() != case['sha256']:
        return dict(result, status='fixture-hash-mismatch')
    env = dict(os.environ, PYTHONPATH=str(ROOT / 'src'))
    try:
        run = subprocess.run([sys.executable, __file__, '--one', str(path)],
                             capture_output=True, text=True, env=env, timeout=timeout)
    except subprocess.TimeoutExpired:
        return dict(result, status='timeout', timeout_seconds=timeout)
    if run.returncode:
        return dict(result, status='process-error', error=run.stderr[-1000:])
    outcome = json.loads(run.stdout)
    if case.get('expected_outcome') == 'reject':
        if outcome['status'] == 'rejected-needs-review':
            outcome['status'] = 'expected-rejection'
        elif outcome['status'] == 'extracted':
            outcome['status'] = 'unexpected-acceptance'
    return dict(result, **outcome)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--one', type=Path, help=argparse.SUPPRESS)
    parser.add_argument('--manifest', type=Path, default=DATA / 'corpus.json')
    parser.add_argument('--output', type=Path, default=ROOT / 'docs' / 'corpus-results.json')
    parser.add_argument('--workers', type=int, default=8)
    parser.add_argument('--timeout', type=float, default=5)
    args = parser.parse_args()
    if args.one:
        print(json.dumps(extract_one(args.one), ensure_ascii=False))
        return 0
    if args.workers < 1 or args.timeout <= 0:
        parser.error('workers and timeout must be positive')
    cases = json.loads(args.manifest.read_text())
    start = time.monotonic()
    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        results = list(pool.map(lambda case: check(case, args.timeout), cases))
    counts = dict(Counter(result['status'] for result in results))
    report = {'files': len(results), 'unique_files': len({case['sha256'] for case in cases}),
              'seconds': round(time.monotonic() - start, 3), 'outcomes': counts,
              'meaning': 'extracted means the parser returned text; only curated assertions establish expected text correctness.',
              'results': results}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(report, indent=2, ensure_ascii=False) + '\n')
    print(json.dumps({key: value for key, value in report.items() if key != 'results'}, ensure_ascii=False))
    return int(any(result['status'] in {'unexpected-error', 'unexpected-acceptance', 'process-error', 'timeout', 'fixture-hash-mismatch', 'rejected-needs-review'} for result in results))


if __name__ == '__main__':
    raise SystemExit(main())
