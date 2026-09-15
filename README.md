# doc-text

`doc-text` is an independent, dependency-free Python project for extracting
text from classic Microsoft Word `.doc` files. It is maintained locally in
its own Git repository. The Python import remains `legacy_doc`.

The project derives from [legacy-doc](https://github.com/lcorrigan704/legacy-doc)
at commit `33aac7cace0dc30d95e671df5aaa825e4c2223a1`. Its original MIT copyright
and permission notice are preserved in `LICENSE`.

It is intentionally narrow:

- supports OLE Compound File Word 97-2003 `.doc` files
- locates text through the FIB-designated CLX and extracts the selected body content
- rejects encrypted, malformed, non-OLE, or unsafe files clearly
- does not execute macros, embedded objects, links, scripts, or external references
- does not use native binaries, LibreOffice, `antiword`, or third-party parsers

## Install locally

From this checkout:

```bash
python3 -m pip install .
```

The distribution is named `doc-text`, with its own version starting at `0.1.0`.
Install this checkout to use it; the upstream `legacy-doc` distribution is a
different project. Both expose the `legacy_doc` import, so use a separate
environment when comparing them.

## Usage

```python
from legacy_doc import extract_text

with open("document.doc", "rb") as file:
    result = extract_text(file.read())

print(result.text)
print(result.metadata)
```

`metadata` always includes extracted text size fields and lightweight OLE
inventory fields. When present in the document's saved Word metadata, it can
also include values such as title, author, company, page count, word count,
character count, creation time, and last-saved time.
Only returned properties and their codepage are decoded. Unreturned property
values are skipped; their index bounds are still checked.

## API

```python
extract_text(document_bytes: bytes, *, options: ExtractionOptions | None = None) -> DocExtractionResult
```

`DocExtractionResult` contains:

- `text`: body text with paragraph and table separators; an empty string for a valid document with no selected text
- `parser`: parser name
- `version`: parser version
- `metadata`: lightweight extraction metadata
- `warnings`: non-fatal parser warnings

Example metadata:

```python
{
    "chars": 1200,
    "bytes": 1225,
    "title": "Quarterly Notes",
    "author": "Liam Corrigan",
    "page_count": 4,
    "word_count": 210,
    "has_macros": False,
    "has_embedded_objects": False,
}
```

## Text extraction rules

This development branch changes the default output of `extract_text()`. It uses
one binary parsing path and does not provide the former whitespace-compression
or heuristic CLX-scanning modes.

- Extract main-body paragraphs and tables. Do not append independent headers,
  footers, footnotes, endnotes, comments, or header textboxes.
- Read supported body textboxes through their drawing anchors and text ranges.
- Replace indexed EMBED fields with an inline `[嵌入文件：name/type]` or
  `[嵌入文件]` placeholder. Attachment payloads are not parsed and these
  placeholders do not add warnings. Naming limits are documented in
  [the development plan](docs/development-plan.md#81-主文档数据选择与嵌入文件占位).
- Preserve other saved field instructions and results, revision-deleted text, hidden
  text, consecutive spaces, and Unicode private-use characters. The extractor
  does not calculate fields or apply revision visibility.
- Separate paragraphs and table rows with LF and cells with TAB. Keep empty
  paragraphs between body paragraphs. Inside a cell, omit empty paragraphs and
  join the remaining paragraphs with LF.
- Flatten nested-table paragraphs inside their containing cell. Remove trailing
  whitespace from each serialized table row, including trailing empty cells.
- Return `text=""` when valid selected content is empty. Required text or
  structural corruption raises `LegacyDocError`; unreadable optional metadata
  produces a warning.

For example, a two-row table containing `Name | Age` and `Alice | 28` produces
`"Name\tAge\nAlice\t28"`. Existing callers receive this output directly through
`result.text`; the call signature and result type are unchanged.

The text policy follows a fixed DOCX XML traversal: text, deleted-text and field
instruction nodes are retained, direct table cells are traversed once, and no
visual grid expansion is performed. `python-docx` is a source reference only;
it is not a dependency and its `row.cells` behavior is not the output policy.

## Limitations

This is a text extractor, not a Microsoft Word renderer. It does not calculate
layout or page breaks, perform OCR, execute macros, dereference external links,
or extract embedded documents. Word 6/95 and encrypted DOC files are unsupported.

Body shapes that store text only in OfficeArt geometry-text properties are
explicitly rejected because those strings do not have a unique correspondence
to the selected DOCX text nodes. They are not silently omitted. Style-dependent
`CFSpec` toggle values also raise an explicit unsupported error when needed to
classify a drawing anchor or symbol; this branch does not evaluate the full
Word style hierarchy.

Linked textboxes require complete range and anchor coverage. Files whose
textbox break tables provide only partial coverage are explicitly rejected.

Floating text follows logical anchor order, not visual position. Horizontal DOC
merge groups are projected to one logical output cell, preserving stored
paragraphs; vertical continuations do not repeat text from preceding rows.
Different DOCX representations of the same visible document can produce different
text. In particular, XML wrappers, textbox descendants and geometry text do not
have a universal one-to-one mapping to binary DOC structures. Supported structural
cases and explicit rejection paths are covered by tests. Body textbox coverage
includes constructed binary records and two unchanged LibreOffice regression
documents whose metadata identifies Microsoft Office Word. WPS and real linked
textbox files still need dedicated coverage. Arbitrary visual DOC/DOCX
equivalence is not promised. The borrowed fixtures and adapted text conditions
are listed in `tests/data/upstream/cases.json`.

`page_count` comes from saved document metadata and is not recalculated. It can
be stale if the document was saved without updating its statistics.

## Development

```bash
python3 -m pip install -e ".[dev]"
python3 -m pytest
python3 -m build
python3 -m twine check dist/*
```

The test suite includes generated FIB/CLX and property records, a known-content DOC export, and cases for parser limits, encryption, malformed references, Unicode and table boundaries. Tests do not require Word or a conversion program.

The development contract is in `docs/development-plan.md`. Run the larger local
DOC corpus separately with:

```bash
python3 scripts/check_corpus.py
```

The report in `docs/corpus-results.json` separates extracted files, explicit
unsupported cases, rejected files requiring investigation, and unexpected
failures. A returned string is a successful extraction attempt; exact text
correctness is asserted separately in the curated regression tests.

## Security

`doc-text` treats input files as untrusted binary data. It performs bounded parsing, rejects encrypted documents, and never executes macros or embedded content.
