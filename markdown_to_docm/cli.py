from __future__ import annotations

from pathlib import Path

import click

from .converter import convert


@click.command()
@click.argument("input_file", metavar="INPUT", type=click.Path(exists=True, dir_okay=False))
@click.option(
    "--template", "-t",
    required=True,
    type=click.Path(exists=True, dir_okay=False),
    help="Path to the .docm template file.",
)
@click.option(
    "--output", "-o",
    required=True,
    type=click.Path(dir_okay=False),
    help="Output .docm file path.",
)
@click.option(
    "--plantuml-server", "-s",
    default="http://www.plantuml.com/plantuml",
    show_default=True,
    help="Base URL of the PlantUML server.",
)
def main(input_file: str, template: str, output: str, plantuml_server: str) -> None:
    """Convert a Markdown file with PlantUML diagrams to a .docm Word document."""
    convert(
        input_path=Path(input_file),
        template_path=Path(template),
        output_path=Path(output),
        plantuml_server=plantuml_server.rstrip("/"),
    )
    click.echo(f"Saved: {output}")
