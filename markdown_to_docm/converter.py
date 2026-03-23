from __future__ import annotations

from pathlib import Path

import mistune
from docx import Document

from .md_renderer import DocxRenderer
from .plantuml_client import fetch_diagram


def convert(
    input_path: Path,
    document_path: Path,
    output_path: Path,
    plantuml_server: str,
) -> None:
    text = input_path.read_text(encoding="utf-8")

    doc = Document(str(document_path))
    doc.add_page_break()

    md = mistune.create_markdown(renderer="ast", plugins=["table"])
    tokens = md(text)

    def fetcher(uml_text: str) -> bytes:
        return fetch_diagram(uml_text, plantuml_server)

    renderer = DocxRenderer(doc, fetcher)
    renderer.render(tokens)

    doc.save(str(output_path))
