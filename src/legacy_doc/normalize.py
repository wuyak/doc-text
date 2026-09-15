"""Whitespace rule shared with the selected DOCX text extraction contract."""
from __future__ import annotations


def normalize_word_text(text: str) -> str:
    """Normalize line endings and trim boundary spaces/LF, preserving content.

    DOC structural controls are consumed by the paragraph/text traversal before
    this function is called. Field instructions and private-use text are content.
    """
    return text.replace("\r\n", "\n").replace("\r", "\n").strip(" \n")
