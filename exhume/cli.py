"""Click CLI entry point for Exhume."""

from __future__ import annotations

from pathlib import Path

import click

from exhume.workbook import Workbook
from exhume.converter import convert_workbook
from exhume.output import format_json, format_text


@click.group()
@click.version_option(version="0.1.0", prog_name="exhume")
def main():
    """Exhume -- Excel introspection and embedded object extraction."""


def _output_option(f):
    f = click.option("--output", "out_fmt", type=click.Choice(["json", "text"]), default="json",
                     help="Output format (default: json)")(f)
    f = click.option("--pretty", is_flag=True, default=False, help="Pretty-print JSON output")(f)
    f = click.option("--quiet", is_flag=True, default=False, help="Suppress progress messages")(f)
    return f


def _format_result(data, out_fmt: str, pretty: bool) -> str:
    if out_fmt == "text":
        return format_text(data)
    return format_json(data, pretty=pretty)


@main.command()
@click.argument("file", type=click.Path(exists=True))
@_output_option
def info(file: str, out_fmt: str, pretty: bool, quiet: bool):
    """Show workbook structure: sheets, dimensions, merged cells, named ranges."""
    wb = Workbook(file)
    result = wb.info()
    click.echo(_format_result(result, out_fmt, pretty))


@main.command()
@click.argument("file", type=click.Path(exists=True))
@click.option("--sheet", default=None, help="Sheet name or index (default: all sheets)")
@_output_option
def cells(file: str, sheet: str | None, out_fmt: str, pretty: bool, quiet: bool):
    """Show cell data with types, formulas, and styles."""
    wb = Workbook(file)
    if sheet is not None:
        # Try to parse as integer index
        try:
            sheet_ref = int(sheet)
        except ValueError:
            sheet_ref = sheet
        result = wb.cells(sheet_ref)
        click.echo(_format_result(result, out_fmt, pretty))
    else:
        # All sheets
        if out_fmt == "text":
            for name in wb.sheet_names():
                click.echo(f"\n--- {name} ---")
                sheet_cells = wb.cells(name)
                click.echo(format_text(sheet_cells) if sheet_cells else "  (empty)")
        else:
            all_cells = {}
            for name in wb.sheet_names():
                all_cells[name] = [c.to_dict() for c in wb.cells(name)]
            click.echo(format_json(all_cells, pretty=pretty))


@main.command()
@click.argument("file", type=click.Path(exists=True))
@click.option("--sheet", default=None, help="Filter by sheet name")
@_output_option
def objects(file: str, sheet: str | None, out_fmt: str, pretty: bool, quiet: bool):
    """List all embedded objects with metadata and neighbor cell values."""
    wb = Workbook(file)
    objs = wb.list_objects(sheet_name=sheet)
    result = {
        "file": Path(file).name,
        "totalObjects": len(objs),
        "objects": objs,
    }
    click.echo(_format_result(result, out_fmt, pretty))


@main.command()
@click.argument("file", type=click.Path(exists=True))
@click.option("--out", default="./extracted", help="Output directory (default: ./extracted)")
@click.option("--flat/--by-row", default=True, help="Flat output or organized by row")
@click.option("--no-clean", is_flag=True, default=False, help="Keep binary path prefix in content")
@_output_option
def extract(file: str, out: str, flat: bool, no_clean: bool, out_fmt: str, pretty: bool, quiet: bool):
    """Extract all embedded objects to disk."""
    wb = Workbook(file)
    objs = wb.extract_objects(out, by_row=not flat, clean=not no_clean)

    result = {
        "outputDir": out,
        "totalExtracted": len(objs),
        "objects": objs,
    }
    if not quiet:
        click.echo(_format_result(result, out_fmt, pretty))


@main.command()
@click.argument("file", type=click.Path(exists=True))
@click.option("--format", "fmt", type=click.Choice(["json", "csv"]), required=True, help="Output data format")
@click.option("--out", default="./converted", help="Output directory (default: ./converted)")
@click.option("--quiet", is_flag=True, default=False, help="Suppress progress messages")
def convert(file: str, fmt: str, out: str, quiet: bool):
    """Full export: cell data + metadata + embedded objects."""
    convert_workbook(file, out, fmt=fmt)
    if not quiet:
        click.echo(format_json({"outputDir": out, "format": fmt, "status": "complete"}))


if __name__ == "__main__":
    main()
