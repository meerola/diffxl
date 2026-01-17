# diffxl

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.13+](https://img.shields.io/badge/python-3.13+-blue.svg)](https://www.python.org/downloads/)
[![Code Style: Ruff](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json)](https://github.com/astral-sh/ruff)

This is a CLI tool for comparing two excel tables based on a unique identifier column. The differences between the old and new file are output into a report in xlsx and html format.

A typical use case could be comparing two revisions of the same document to answer the question: "What has been added/removed/changed in the table, and how have the cell values changes?". 

For example, in plant engineering projects, diffxl can show exactly what has changed between two revisions of a valve list, giving a portable report in xlsx/html format that show exactly what details have changed since the list was last submitted.

diffxl is meant to be able to quickly do a comparison, and will try to do the comparison even if no sheet name or UID column name is specified. If no sheet name is specified, diffxl will use the active sheet, find the table and use the left-most column as the UID. 

## Features

* **Detailed Comparison on cell and column level:**
  * Identifies **Added** rows and columns.
  * Identifies **Removed** rows and columns.
  * Identifies **Changed** cell values.
* **Smart Table Detection:** Looks through the sheets for the UID column. Automatically finds the table header and ignores data above/below the table.
* **Web-based Diff Report:** Generates an interactive HTML diff report with filtering, highlighting and inline comparison.
* **Excel Diff Report:** Same principle as for HTML report, but in excel format. Additions, removals and changes are shown in separate sheets.
* **Failure Analysis:** If diffxl fails (e.g., missing key), the tool analyzes the files and can suggest candidate columns. Using the --diagnostic flag will generate a HTML diagnostic report. 
* **NaN-value normalization**: by default (e.g., treating `NaN` as equal to `None` or `""`). Using the --raw flag will disable this behavior.
* **Diff with duplicate UIDs**: The --dedup flag allows diffxl to proceed by ignoring duplicate UID rows. The ignored rows will be shown in the diff report so the user knows what was removed.
* **Format Support:** Supports `.xlsx`, `.xlsm`, `.xls`, and `.csv`.é

## Installation

```bash
pip install diffxl
```

## Usage

```bash
# Simple usage (uses leftmost column as key, generates Excel + HTML)
diffxl <old_file> <new_file>

# Specify a key column
diffxl <old_file> <new_file> --key "Tag"
```

### Arguments

* `old_file`: Path to the original file.
* `new_file`: Path to the new file.
* `--key`, `-k`: The column name to use as the unique identifier (default: **column** **furthest to the left in the identified table**).
* `--sheet`, `-s`: (Optional) Specific sheet name to compare.
* `--output`, `-o`: Output filename (default: `diff_report.xlsx`).
* `--prefix`, `-p`: Add a prefix to output filenames (e.g., `ABC_diff_report.xlsx`) to keep track of multiple runs.
* `--raw`: Perform exact string comparison (disable smart normalization like treating `NaN` as equal to `None`).
* `--no-web`: Disable HTML report generation.
* `--diagnostic`, `-d`: Generate a detailed HTML diagnostic report if validation fails.
* `--dedup`: Remove duplicate rows based on Key column (keeps first occurrence).

### Example

```bash
diffxl samples/valvelist_v1.xlsx samples/valvelist_v2.xlsx
```

## Output

By default, the tool generates:

1. **Excel Report (`diff_report.xlsx`)**: Contains Added, Removed, Changed, and Complete diff sheets.
2. **Web Report (`diff_report.html`)**: An interactive comparison view.

## Smart Diagnostics

If the tool cannot find your specified key column, it will automatically analyze the file and suggest alternative columns that look like unique identifiers (UIDs), sorted by confidence.

To get a full visual analysis when something goes wrong, run with:

```bash
diffxl <old_file> <new_file> --key "WrongKey" --diagnostic
```

## Development

This project uses [uv](https://github.com/astral-sh/uv) for dependency management.

```bash
# Install dependencies
uv sync
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
