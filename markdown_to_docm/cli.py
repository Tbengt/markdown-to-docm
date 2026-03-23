from __future__ import annotations

from pathlib import Path

import click

from .converter import convert


@click.command()
@click.argument("input_file", metavar="INPUT", type=click.Path(exists=True, dir_okay=False))
@click.option(
    "--document", "-d",
    required=True,
    type=click.Path(exists=True, dir_okay=False),
    help="Path to the .docx/.docm file to append content to.",
)
@click.option(
    "--output", "-o",
    required=True,
    type=click.Path(dir_okay=False),
    help="Output file path.",
)
@click.option(
    "--plantuml-server", "-s",
    default="http://www.plantuml.com/plantuml",
    show_default=True,
    help="Base URL of the PlantUML server.",
)
def main(input_file: str, document: str, output: str, plantuml_server: str) -> None:
    """Append rendered Markdown with PlantUML diagrams to an existing .docx/.docm file."""
    convert(
        input_path=Path(input_file),
        document_path=Path(document),
        output_path=Path(output),
        plantuml_server=plantuml_server.rstrip("/"),
    )
    click.echo(f"Saved: {output}")
