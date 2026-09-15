# xml2xlsx

Converts a microplate reader's XML export into an XLSX workbook, one sheet per plate.

The input file is SpreadsheetML 2003: XML that describes a spreadsheet, one row and cell at a time. It can hold several plates below each other on one worksheet, each starting with a "Plate" header row. The script finds them and copies each one to its own sheet of a new .xlsx file, keeping the cells and their formatting as they were.

## Scope

- This is a small tool for our own plate data. It knows the layout our reader produces and it relies on that layout.
- It is **not** a general XML to XLSX converter. If you give it a file from somewhere else, it will most likely find no plates and write nothing.
- It does not interpret the data in any way.
- Linux x86_64 only. There are no Windows or macOS builds.

## Installing

The script is provided in CLI form and a GUI form. Download it, make it executable, and put it somewhere in your `PATH`:

```bash
mkdir -p ~/.local/bin

# GUI
curl -Lo ~/.local/bin/xml2xlsx-gui https://github.com/VachaLab/xml2xlsx/releases/latest/download/xml2xlsx-gui-linux-x86_64

# CLI
curl -Lo ~/.local/bin/xml2xlsx https://github.com/VachaLab/xml2xlsx/releases/latest/download/xml2xlsx-linux-x86_64

chmod u+x ~/.local/bin/xml2xlsx-gui ~/.local/bin/xml2xlsx
```

That URL always points at the newest release, so running it again updates. If `xml2xlsx-gui` is not found afterwards, `~/.local/bin` is not in your `PATH`.

Both files are also on the [releases page](https://github.com/VachaLab/xml2xlsx/releases/latest) if you would rather click.

## GUI

```bash
xml2xlsx-gui
```

Drop XML files onto it, or click to pick them. Press Convert and choose where to save. Every plate gets a line in the table saying what happened to it.

## CLI

```bash
xml2xlsx -i plates.xml -o plates.xlsx
```

`-i` takes one or more XML files, or a folder, in which case all .xml files _directly_ inside it are read. Repeat it to merge several inputs into one workbook, in the order given:

```bash
xml2xlsx -i plates1.xml plates2.xml -i rerun.xml -o merged.xlsx
```

It prints one line per plate saying what happened to it, and exits non-zero if nothing was written or a file could not be read.

## What happens to the plates

A plate with no readings is skipped, as is one whose name has already been used, since those are typically the same plate exported twice; if the values disagree, you get a warning and the first one is kept. A plate with no name becomes "Unknown plate #1" and so on. An existing output file is overwritten without asking.

## Using from source

Needs [uv](https://docs.astral.sh/uv/getting-started/installation/).

```bash
uv run xml2xlsx -i plates.xml -o plates.xlsx
uv run xml2xlsx-gui
```

## Disclaimer

The GUI was built by Claude Opus 5.
