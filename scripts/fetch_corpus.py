#!/usr/bin/env python3
"""Fetch the pinned upstream DOC regression corpus and write its manifest.

The script deliberately keeps every upstream download in a cache.  Only the
first copy of a byte-identical file is copied into tests/data/corpus; the ten
fixtures already checked in under tests/data/upstream are referenced in place.
"""

from __future__ import annotations

import argparse
import concurrent.futures
import hashlib
import json
import shutil
import subprocess
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "tests" / "data"
CORPUS = DATA / "corpus"
MANIFEST = DATA / "corpus.json"
CACHE = ROOT / "tools" / "corpus-cache"

POI = {
    "id": "apache-poi",
    "repo": "apache/poi",
    "commit": "dca7cda1dc8a64197f7e7e8e7892ba9a0d79c860",
    "path": "test-data/document",
    "label": "Apache POI",
}
LO = {
    "id": "libreoffice",
    "repo": "LibreOffice/core",
    "commit": "23efc7f91ea1af24dafdbd1ca7d8b121cee96fec",
    "path": "sw/qa/extras/ww8export/data",
    "label": "LibreOffice core",
}
OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def _read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def list_source(source: dict[str, str]) -> list[dict[str, Any]]:
    """Read a pinned contents listing, caching the API response."""
    cache_path = CACHE / f"{source['id']}-directory.json"
    if cache_path.exists():
        listing = _read_json(cache_path)
    else:
        api_url = (
            f"repos/{source['repo']}/contents/{source['path']}"
            f"?ref={source['commit']}"
        )
        result = subprocess.run(
            ["gh", "api", api_url, "--paginate", "--slurp"],
            check=True,
            capture_output=True,
            text=True,
        )
        pages = json.loads(result.stdout)
        listing = [item for page in pages for item in page] if pages and isinstance(pages[0], list) else pages
        _write_json(cache_path, listing)
    return [
        item
        for item in listing
        if item.get("type") == "file" and item.get("name", "").endswith(".doc")
    ]


def source_url(source: dict[str, str], name: str) -> str:
    encoded = urllib.parse.quote(name, safe="")
    return f"https://github.com/{source['repo']}/blob/{source['commit']}/{source['path']}/{encoded}"


def raw_url(source: dict[str, str], name: str) -> str:
    encoded = urllib.parse.quote(name, safe="")
    return f"https://raw.githubusercontent.com/{source['repo']}/{source['commit']}/{source['path']}/{encoded}"


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def git_blob_sha(data: bytes) -> str:
    header = f"blob {len(data)}\0".encode("ascii")
    return hashlib.sha1(header + data).hexdigest()


def fetch_one(source: dict[str, str], item: dict[str, Any]) -> dict[str, Any]:
    """Fetch and validate one source file; failures are returned to the caller."""
    name = item["name"]
    target = CACHE / "raw" / source["id"] / name
    target.parent.mkdir(parents=True, exist_ok=True)
    try:
        if not target.exists():
            request = urllib.request.Request(raw_url(source, name), headers={"User-Agent": "legacy-doc-corpus"})
            with urllib.request.urlopen(request, timeout=90) as response:
                data = response.read()
            target.write_bytes(data)
        else:
            data = target.read_bytes()
        actual_blob = git_blob_sha(data)
        expected_blob = item.get("sha")
        if actual_blob != expected_blob:
            raise RuntimeError(f"Git blob SHA mismatch: expected {expected_blob}, got {actual_blob}")
        return {
            "source": source,
            "item": item,
            "name": name,
            "cache_path": target,
            "sha256": sha256_bytes(data),
            "blob_sha": actual_blob,
            "size": len(data),
            "is_ole": data.startswith(OLE_SIGNATURE),
        }
    except Exception as exc:  # noqa: BLE001 - collect every failed source item
        return {"source": source, "item": item, "name": name, "error": repr(exc)}


def fetch_many(source: dict[str, str], items: list[dict[str, Any]], workers: int) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(fetch_one, source, item) for item in items]
        for future in concurrent.futures.as_completed(futures):
            results.append(future.result())
    return sorted(results, key=lambda result: result["name"])


def existing_entries() -> tuple[list[dict[str, Any]], set[str]]:
    cases_path = DATA / "upstream" / "cases.json"
    cases = _read_json(cases_path)
    entries: list[dict[str, Any]] = []
    hashes: set[str] = set()
    for case in cases:
        entry = dict(case)
        relative = entry["file"]
        if not relative.startswith("upstream/"):
            relative = "upstream/" + relative
        entry["file"] = relative
        entry.setdefault("source_condition", entry.pop("upstream_condition", None))
        entry.setdefault("upstream_test", entry.pop("upstream_test", None))
        entry.setdefault("source_commit", None)
        path = DATA / relative
        data = path.read_bytes()
        actual_sha256 = sha256_bytes(data)
        if actual_sha256 != entry["sha256"]:
            raise RuntimeError(f"Existing fixture SHA-256 mismatch: {relative}")
        actual_blob = git_blob_sha(data)
        if actual_blob != entry["blob_sha"]:
            raise RuntimeError(f"Existing fixture blob SHA mismatch: {relative}")
        entry["size"] = len(data)
        entry["is_ole"] = data.startswith(OLE_SIGNATURE)
        if not entry["is_ole"]:
            entry["format_note"] = "non-OLE bytes with .doc extension"
        hashes.add(actual_sha256)
        entries.append(entry)
    return entries, hashes


def category_for(name: str) -> tuple[str, str]:
    """Return only categories supported by an unambiguous filename signal."""
    lower = name.lower()
    if "word95" in lower or "word6" in lower:
        return "unsupported-format", "Filename explicitly identifies an old Word format."
    if "password" in lower or "encrypt" in lower:
        return "unsupported-format", "Filename explicitly identifies password/encryption coverage."
    if "fuzz" in lower or "truncated" in lower:
        return "malformed-input", "Filename explicitly identifies fuzzed or truncated input."
    if "textbox" in lower:
        return "body-textboxes", "Filename explicitly identifies a textbox regression."
    if any(token in lower for token in ("table", "cell", "column")):
        return "tables", "Filename explicitly identifies table/cell/column coverage."
    if any(token in lower for token in ("header", "footer", "headfoot", "footnote", "endnote")):
        return "headers-footers", "Filename explicitly identifies header/footer/note coverage."
    if any(token in lower for token in ("list", "numbering", "number")):
        return "lists-numbering", "Filename explicitly identifies list/numbering coverage."
    if any(token in lower for token in ("field", "pageref", "mergefield")):
        return "fields", "Filename explicitly identifies field coverage."
    return "general-regression", "General parser regression; no specific text-content assertion has been verified."


def make_entry(result: dict[str, Any], relative: str) -> dict[str, Any]:
    source = result["source"]
    item = result["item"]
    category, basis = category_for(result["name"])
    entry: dict[str, Any] = {
        "file": relative,
        "sha256": result["sha256"],
        "size": result["size"],
        "blob_sha": result["blob_sha"],
        "source": source_url(source, result["name"]),
        "source_commit": source["commit"],
        "category": category,
        "category_basis": basis,
        "upstream_test": None,
        "source_condition": None,
        "is_ole": result["is_ole"],
        "format_note": "OLE Compound File" if result["is_ole"] else "non-OLE bytes with .doc extension",
        "licensing": (
            "Apache POI source fixture; repository LICENSE/NOTICE retained under "
            "tests/data/upstream/apache-poi/. No document-specific notice was found."
            if source["id"] == "apache-poi"
            else "LibreOffice core regression fixture; source readlicense_oo/license/NOTICE is retained in "
            "tests/data/corpus/LICENSES.md. No document-specific notice was found; local regression use."
        ),
    }
    # Keep the API's blob SHA as an extra checkable field even though it is
    # already represented by blob_sha in the manifest.
    if item.get("size") is not None and item["size"] != result["size"]:
        raise RuntimeError(f"Size mismatch for {source['id']}/{result['name']}")
    return entry


def copy_unique(result: dict[str, Any], relative_dir: str) -> str:
    destination = CORPUS / relative_dir / result["name"]
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(result["cache_path"], destination)
    return destination.relative_to(DATA).as_posix()


def apply_conditions(entries: list[dict[str, Any]]) -> None:
    """Keep reviewed source conditions as data, separate from file fetching."""
    conditions = _read_json(DATA / "corpus-conditions.json")
    aliases = {"encoding": "body", "pictures": "body", "headers-footers": "body-scope",
               "lists-numbering": "body", "unknown": "general-regression"}
    rules = {
        "body": "Extract stored body text; formatting and images do not create extra text.",
        "body-scope": "Extract body text and exclude independent header, footer and note stories.",
        "tables": "Preserve logical text order with the agreed TAB/LF table projection.",
        "fields": "Keep stored field instructions and results; consume field control markers.",
        "body-textboxes": "Insert supported textbox text at its body anchor.",
        "empty-body": "Return an empty string successfully for an empty selected body.",
        "unsupported-format": "Reject document formats or encryption outside the supported scope.",
        "malformed-input": "Do not crash or hang; source evidence determines whether rejection is expected.",
        "general-regression": "Run bounded extraction and record explicit outcomes; exact expected text is not yet asserted.",
    }
    for entry in entries:
        entry.update(conditions.get(entry["file"], {}))
        old = entry.get("category", "general-regression")
        entry["category"] = aliases.get(old, old)
        if old != entry["category"]:
            entry["upstream_category"] = old
        if old == "unknown":
            entry["category_basis"] = "General parser regression; no specific text-content assertion has been verified."
        entry.setdefault("classification_status", "curated" if entry["file"].startswith("upstream/") else "filename-only")
        entry.setdefault("our_condition", rules.get(entry["category"], rules["general-regression"]))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--target", type=int, default=200, help="minimum unique entries including existing fixtures")
    parser.add_argument("--workers", type=int, default=8)
    args = parser.parse_args()
    if args.workers < 1:
        parser.error("--workers must be positive")

    CACHE.mkdir(parents=True, exist_ok=True)
    CORPUS.mkdir(parents=True, exist_ok=True)
    entries, seen = existing_entries()

    poi_items = list_source(POI)
    poi_results = fetch_many(POI, poi_items, args.workers)
    failures = [result for result in poi_results if "error" in result]
    if failures:
        for result in failures:
            print(f"ERROR {POI['id']}/{result['name']}: {result['error']}")
        return 1

    for result in poi_results:
        if result["sha256"] in seen:
            continue
        relative = copy_unique(result, "apache-poi")
        entries.append(make_entry(result, relative))
        seen.add(result["sha256"])

    if len(seen) < args.target:
        lo_items = list_source(LO)
        lo_results = fetch_many(LO, lo_items, args.workers)
        failures = [result for result in lo_results if "error" in result]
        if failures:
            for result in failures:
                print(f"ERROR {LO['id']}/{result['name']}: {result['error']}")
            return 1
        for result in lo_results:
            if len(seen) >= args.target:
                break
            if result["sha256"] in seen:
                continue
            relative = copy_unique(result, "libreoffice")
            entries.append(make_entry(result, relative))
            seen.add(result["sha256"])

    if len(seen) < args.target:
        raise RuntimeError(f"Sources produced only {len(seen)} unique files; target is {args.target}")
    entries.sort(key=lambda entry: entry["file"])
    apply_conditions(entries)
    _write_json(MANIFEST, entries)
    print(f"Wrote {len(entries)} manifest entries ({len(seen)} unique SHA-256 files) to {MANIFEST}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
