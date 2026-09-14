## xml2xlsx

Converts a microplate reader's XML export into an XLSX workbook, one sheet per plate.

The input file is SpreadsheetML 2003: XML that describes a spreadsheet, one row and cell at a time. It can hold several plates below each other on one worksheet, each starting with a "Plate" header row. The script finds them and copies each one to its own sheet of a new .xlsx file, keeping the cells and their formatting as they were.

### Scope

- This is a small tool for our own plate data. It knows the layout our reader produces and it relies on that layout.
- It is **not** a general XML to XLSX converter. If you give it a file from somewhere else, it will most likely find no plates and write nothing.
- It does not interpret the data in any way.
- It is intentionally provided as a standalone script, not a python package.

### How to run it

Needs [uv](https://docs.astral.sh/uv/getting-started/installation/).

Once, per machine:

```bash
chmod u+x xml2xlsx.py
```

Then:

```bash
./xml2xlsx.py -i plates.xml -o plates.xlsx
```

`-i` takes one or more XML files, or a folder, in which case all .xml files _directly_ inside it are read. Repeat it to merge several inputs into one workbook, in the order given:

```bash
./xml2xlsx.py -i plates1.xml plates2.xml -i rerun.xml -o merged.xlsx
```

The script prints one line per plate saying what happened to it. A plate with no readings is skipped, as is one whose name has already been used, since those are typically the same plate exported twice; if the values disagree, the script warns. An existing output file is overwritten without asking.
