"""Run the local DOC corpus with per-file timeouts and explicit outcomes.

This checks parser behavior and records failures; it does not invent expected
text or treat an explicit rejection as proof that a document is malformed.
"""
from __future__ import annotations

import argparse
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import time

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / 'tests' / 'data'


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
        elif 'exceeds parser limit' in message:
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
