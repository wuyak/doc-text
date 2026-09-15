"""Project DOC paragraphs onto the selected DOCX text traversal semantics.

This is a logical text traversal, not a page layout or visual reading-order model.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from bisect import bisect_right
from typing import TYPE_CHECKING, Iterable

from legacy_doc._characters import CharacterIndex
from legacy_doc._objects import EmbeddedObjects, EmbeddedField
from legacy_doc.exceptions import LegacyDocError
from legacy_doc.normalize import normalize_word_text

if TYPE_CHECKING:
    from legacy_doc._binary import BinaryDocument
    from legacy_doc._paragraphs import Paragraph, CellFormat
    from legacy_doc._shapes import TextboxRange


@dataclass
class _Paragraph:
    text: str
    raw_text: str
    descendants: list[str] = field(default_factory=list)


@dataclass
class _Table:
    rows: list[list[list[_Paragraph | _Table]]] = field(default_factory=list)


@dataclass
class _Frame:
    table: _Table
    row: list[list[_Paragraph | _Table]] = field(default_factory=list)
    cell: list[_Paragraph | _Table] = field(default_factory=list)


class TextRenderer:
    def __init__(self, doc: BinaryDocument, textboxes: dict[int, TextboxRange]) -> None:
        self.doc = doc
        self.textboxes = textboxes
        self.characters = CharacterIndex(doc)
        # Validate selected anchors before traversal, including positions that
        # would otherwise be skipped at a story or paragraph boundary.
        for cp in textboxes:
            if not 0 <= cp < doc.fib.ccp_text:
                raise LegacyDocError("Textbox anchor is outside main-document text")
            if doc.read_text(cp, cp + 1) != "\x08" or not self.characters.is_special(cp):
                raise LegacyDocError("Textbox anchor does not reference a drawing character")
        self.objects = EmbeddedObjects(doc, self.characters)
        self.limit = doc.options.max_text_bytes
        self._active_textboxes: set[tuple[int, int]] = set()
        # Bound work even when repeated references produce little final output.
        self._work_left = max(doc.options.max_file_bytes, 1) * 8

    def _charge(self, count: int) -> None:
        self._work_left -= count
        if self._work_left < 0:
            raise LegacyDocError("DOC text traversal exceeds parser work limit")

    def _join(self, values: Iterable[str], separator: str = "", *, final: bool = True) -> str:
        parts: list[str] = []
        size = 0
        for value in values:
            self._charge(len(value) + 1)
            size += len(value.encode("utf-8")) + (len(separator) if parts else 0)
            if final and size > self.limit:
                raise LegacyDocError(".doc extracted text exceeds parser limit")
            parts.append(value)
        return separator.join(parts)

    def _inline(
        self, paragraph: Paragraph,
        embedded: tuple[tuple[int, ...], list[EmbeddedField]],
    ) -> _Paragraph:
        parts: list[str] = []
        descendants: list[str] = []
        cp = paragraph.start
        raw_bytes = 0
        leading_bytes = 0
        content_seen = False

        def append(value: str) -> None:
            # Check the stable output prefix before storing another fragment.
            # A table row may strip all trailing whitespace; once followed by
            # text, that whitespace is internal and must fit the output limit.
            nonlocal raw_bytes, leading_bytes, content_seen
            raw_bytes += len(value.encode("utf-8"))
            if not content_seen:
                without_leading = value.lstrip(" \n")
                leading_bytes += len(value) - len(without_leading)
                content_seen = bool(without_leading)
            if value.rstrip():
                trailing_bytes = len(value.encode("utf-8")) - len(value.rstrip().encode("utf-8"))
                if raw_bytes - leading_bytes - trailing_bytes > self.limit:
                    raise LegacyDocError(".doc extracted text exceeds parser limit")
            parts.append(value)
        starts, fields = embedded
        for char in paragraph.text:
            self._charge(1)
            index = bisect_right(starts, cp) - 1
            embedded_field = fields[index] if index >= 0 and cp < fields[index].end else None
            textbox = self.textboxes.get(cp)
            if embedded_field is not None:
                if cp == embedded_field.start:
                    append(self.objects.label(embedded_field))
            elif textbox is not None:
                key = (textbox.start, textbox.end)
                if key in self._active_textboxes or len(self._active_textboxes) >= 64:
                    raise LegacyDocError("Cyclic or excessive DOC textbox nesting")
                self._active_textboxes.add(key)
                try:
                    nodes = self._read_blocks(*key)
                    # The DOCX inline visitor sees text nodes inside every box
                    # paragraph without adding paragraph delimiters of its own.
                    box_paragraphs = list(self._paragraph_texts(nodes))
                    append(self._join(self._raw_paragraph_texts(nodes), final=False))
                    descendants.extend(box_paragraphs)
                finally:
                    self._active_textboxes.remove(key)
            elif char == "(" and self.characters.is_special(cp):
                pass  # DOC symbol placeholder corresponds to a w:sym attribute.
            elif char in "\x0b\x0c\r":
                append("\n")
            elif char == "\t" or ord(char) >= 0x20:
                append(char)
            # Other DOC control characters denote fields, objects, references,
            # optional/nonbreaking-hyphen elements, etc.; their saved text is
            # already traversed in order and must not be removed by regex.
            cp += 2 if ord(char) > 0xFFFF else 1
        # Trim at the paragraph boundary exactly as the XML reference does.
        raw_text = "".join(parts)
        value = normalize_word_text(raw_text)
        if len(value.rstrip().encode("utf-8")) > self.limit:
            raise LegacyDocError(".doc extracted text exceeds parser limit")
        return _Paragraph(value, raw_text, descendants)

    def _raw_paragraph_texts(self, nodes: Iterable[_Paragraph | _Table]) -> Iterable[str]:
        for node in nodes:
            if isinstance(node, _Paragraph):
                yield node.raw_text
            else:
                for row in node.rows:
                    for cell in row:
                        yield from self._raw_paragraph_texts(cell)

    def _paragraph_texts(self, nodes: Iterable[_Paragraph | _Table]) -> Iterable[str]:
        for node in nodes:
            if isinstance(node, _Paragraph):
                yield node.text
                yield from node.descendants
            else:
                for row in node.rows:
                    for cell in row:
                        yield from self._paragraph_texts(cell)

    @staticmethod
    def _close(frame: _Frame) -> None:
        if frame.row or frame.cell:
            raise LegacyDocError("DOC table ended before a cell or row terminator")
        if not frame.table.rows:
            raise LegacyDocError("DOC table contains no complete rows")

    def _read_blocks(self, start: int, end: int) -> list[_Paragraph | _Table]:
        from legacy_doc._paragraphs import iter_paragraphs

        embedded = self.objects.selected(start, end)
        blocks: list[_Paragraph | _Table] = []
        frames: list[_Frame] = []
        for paragraph in iter_paragraphs(self.doc, start, end):
            self._charge(1)
            depth = paragraph.depth if paragraph.in_table else 0
            if depth < 0 or depth > 64:
                raise LegacyDocError("DOC table nesting exceeds parser limit")
            while len(frames) > depth:
                self._close(frames.pop())
            # MS-DOC 2.4.3: CellN may begin with TableN+1, before any ParaN.
            # Create each containing cell now; its actual cell/row marks must
            # still close it below. No text or terminators are synthesized.
            while depth > len(frames):
                self._charge(1)
                table = _Table()
                (frames[-1].cell if frames else blocks).append(table)
                frames.append(_Frame(table))
            node = self._inline(paragraph, embedded)
            if depth == 0:
                if paragraph.row_end or paragraph.cell_end:
                    raise LegacyDocError("Table terminator outside a DOC table")
                blocks.append(node)
                continue
            frame = frames[-1]
            if paragraph.row_end:
                if node.text or node.descendants or frame.cell:
                    raise LegacyDocError("DOC row terminator has an unfinished cell")
                if not frame.row or len(frame.row) > 63:
                    raise LegacyDocError("Invalid number of cells in DOC table row")
                cell_formats = paragraph.cells
                if cell_formats is not None and len(cell_formats) != len(frame.row):
                    raise LegacyDocError("DOC table definition does not match cell boundaries")
                frame.table.rows.append(self._merged_row(frame.row, cell_formats))
                frame.row = []
            else:
                frame.cell.append(node)
                if paragraph.cell_end:
                    frame.row.append(frame.cell)
                    frame.cell = []
        while frames:
            self._close(frames.pop())
        return blocks

    @staticmethod
    def _merged_row(
        cells: list[list[_Paragraph | _Table]], formats: tuple[CellFormat, ...] | None,
    ) -> list[list[_Paragraph | _Table]]:
        if formats is None:
            return cells
        # Project a DOC horizontal merged group to one logical output cell,
        # matching a DOCX tc with gridSpan. Preserve all stored paragraphs:
        # visibility alone is not a reason to discard text in this extractor.
        result: list[list[_Paragraph | _Table]] = []
        merging = False
        for cell, fmt in zip(cells, formats):
            if fmt.horizontal_merge == 1:
                if not merging:
                    raise LegacyDocError("DOC horizontal merge has no primary cell")
                result[-1].extend(cell)
            else:
                merging = fmt.horizontal_merge in (2, 3)
                result.append(list(cell))
        return result

    def render(self) -> str:
        chunks: list[str] = []
        for node in self._read_blocks(0, self.doc.fib.ccp_text):
            if isinstance(node, _Paragraph):
                chunks.append(node.text)
            else:
                for row in node.rows:
                    cells = [self._join((p for p in self._paragraph_texts(cell) if p), "\n", final=False)
                             for cell in row]
                    # rstrip is part of the selected reference's row semantics.
                    while cells and not cells[-1].rstrip():
                        cells.pop()
                    if cells:
                        cells[-1] = cells[-1].rstrip()
                    chunks.append(self._join(cells, "\t"))
        # Outer clean_text strips only spaces and LF, not leading/trailing TAB.
        first = 0
        last = len(chunks)
        while first < last and not chunks[first]:
            first += 1
        while last > first and not chunks[last - 1]:
            last -= 1
        return normalize_word_text(self._join(chunks[first:last], "\n"))
