# Released under GPL3 License.
# Copyright (c) 2026 Ladislav Bartos and Robert Vacha Lab

import argparse
import sys
from pathlib import Path

from .core import Summary, convert, format_table


def gather(inputs: list[Path]) -> list[Path]:
    """
    Expand the given paths into the XML files to read.

    A directory contributes the .xml files directly inside it, sorted by name.
    """
    sources: list[Path] = []
    for path in inputs:
        sources.extend(sorted(path.glob("*.xml")) if path.is_dir() else [path])

    return sources


def main(argv: list[str] | None = None) -> int:
    """Parse arguments, convert, report. Return a process exit code."""
    parser = argparse.ArgumentParser(
        description="Convert an XML file exported from a microplate reader "
        "to an XLSX file where each plate is a sheet."
    )
    parser.add_argument(
        "--input",
        "-i",
        dest="inputs",
        type=Path,
        nargs="+",
        action="extend",
        required=True,
        metavar="PATH",
        help="XML export, or a directory of them; repeatable, order is kept",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=Path,
        required=True,
        metavar="PATH",
        help="XLSX file to write",
    )
    arguments = parser.parse_args(argv)

    sources = gather(arguments.inputs)
    if not sources:
        print("error: no XML files among the given inputs", file=sys.stderr)
        return 1

    try:
        report = convert(sources, arguments.output)
    except OSError as exc:
        print(f"error: could not write {arguments.output}: {exc}", file=sys.stderr)
        return 1

    # every plate gets a line, written or not, so a missing plate can be
    # traced to the reason it was dropped
    for line in format_table(report):
        print(line)

    summary = Summary.of(report)
    print(f"\n{summary.text(arguments.output)}")

    return 1 if summary.errors or not summary.written else 0


if __name__ == "__main__":
    sys.exit(main())
