# DOC fixtures

`empty-multipara-emoji.rtf` was authored for this development task. Its `.doc`
was exported locally with LibreOffice 26.2.2.2 (MS Word 97 filter). Both are
provided under the repository MIT license. It covers body/header/footer
separation, empty cells, a multi-paragraph cell and supplementary Unicode.
The test expectation is the authored content under the selected text rules.
This file does not establish Microsoft Word or WPS compatibility by itself.
Tests read the committed DOC directly and do not invoke a converter.

## Borrowed regression fixtures

`upstream/` contains 10 unchanged DOC files from pinned Apache POI and
LibreOffice revisions. `upstream/cases.json` maps each source test condition
to our categories and text assertions, with source links and SHA-256 hashes.
`tests/test_real_documents.py` runs them without network access or converters.

The categories are body text, body scope, empty body, tables, body textboxes,
and unsupported formats. Layout and fill-style tests are adapted to text
inclusion; their original visual assertions are not imported. Textbox labels
were checked directly in the fixture bytes. Table expectations also use the
previously saved paired DOCX text baseline.

These third-party files are not covered by the MIT statement for the authored
RTF/DOC above. Apache's original LICENSE and NOTICE are retained alongside its
samples. The catalogue records the upstream licensing information and any
missing document-specific notice; the files are kept for local regression use.


## Expanded corpus

`corpus.json` lists 200 unique DOC files (including the 10 curated upstream
fixtures), with pinned sources, hashes and our adapted conditions. Reviewed
annotations live in `corpus-conditions.json`; general regressions explicitly
lack exact expected text. Use `python3 scripts/check_corpus.py` for bounded
batch execution. The latest report and unresolved rejections are in
`docs/corpus-results.json`. Cached downloads under `tools/corpus-cache/` are
ignored by Git; the 200 referenced documents are part of the project.
