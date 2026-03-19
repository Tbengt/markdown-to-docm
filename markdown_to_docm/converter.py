from __future__ import annotations

from pathlib import Path

import mistune
from docx import Document
from docx.oxml.ns import qn

from .md_renderer import DocxRenderer
from .plantuml_client import fetch_diagram


def _clear_body(doc: Document) -> None:
    """Remove all body content from the document, preserving section properties."""
    body = doc.element.body
    sect_pr_tag = qn("w:sectPr")
    for child in list(body):
        if child.tag != sect_pr_tag:
            body.remove(child)


def convert(
    input_path: Path,
    template_path: Path,
    output_path: Path,
    plantuml_server: str,
) -> None:
    text = input_path.read_text(encoding="utf-8")

    doc = Document(str(template_path))
    _clear_body(doc)

    md = mistune.create_markdown(renderer="ast", plugins=["table"])
    tokens = md(text)

    def fetcher(uml_text: str) -> bytes:
        return fetch_diagram(uml_text, plantuml_server)

    renderer = DocxRenderer(doc, fetcher)
    renderer.render(tokens)

    doc.save(str(output_path))
