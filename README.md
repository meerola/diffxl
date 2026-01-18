# diffxl

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.13+](https://img.shields.io/badge/python-3.13+-blue.svg)](https://www.python.org/downloads/)
[![Code Style: Ruff](https://img.shields.io/endpoint?url=https://raw.githubusercontent.com/astral-sh/ruff/main/assets/badge/v2.json)](https://github.com/astral-sh/ruff)

**diffxl** is a robust CLI tool designed to compare two Excel tables based on a Unique Identifier (UID). It identifies exactly what has been added, removed, or changed, and generates portable reports in both XLSX and HTML formats.

### The Use Case
A typical use case involves comparing two revisions of the same document to answer: *"What has been added, removed, or changed in the table, and how have the cell values shifted?"*

> **Example:** In plant engineering projects, diffxl can compare two revisions of a valve list. It produces a clear audit trail showing exactly which technical details (pressure ratings, materials, tag numbers) have changed since the last submission.

### Quick Start Logic
**diffxl** is built for quickly getting a comparison done. It attempts to perform a comparison even with minimal configuration:
* **No Key Column?** It defaults to the active sheet, and automatically detects the table header and uses the left-most column as the UID.
* **No Sheet Name?** It scans through the file to find which sheet contains the table with the UID column.
* The html report is opened automatically after a successful diff.

---

## Features

### Deep Comparison
* **Row-Level Diff:** Identifies **Added** and **Removed** rows based on the UID.
* **Cell-Level Diff:** Highlights exactly which cell values have **Changed** between revisions.
* **Format Support:** Compatible with `.xlsx`, `.xlsm`, `.xls`, and `.csv`.

### Reporting
* **Interactive Web Report:** Generates an HTML report with filtering, highlighting, and inline value comparisons.
* **Excel Diff Report:** Generates a multi-sheet Excel workbook separating Additions, Removals, and Changes per column for easy filtering. Also includes a color-coded complete table similar to the html report.

### Smart Processing
* **Auto-Detection:** Automatically locates the table header, ignoring metadata or titles above/below the actual data.
* **NaN Normalization:** By default, treats `NaN`, `None`, and `""` as equal to reduce noise. (Disable with `--raw`).
* **Failure Analysis:** If a comparison fails (e.g., a missing key), `diffxl` analyzes the files and suggests the most likely UID columns.
* **Duplicate Handling:** The `--dedup` flag allows processing of files with duplicate UIDs by keeping the first occurrence and logging the ignored rows.

---

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

Alternatively, run directly with uvx:
```bash
uvx diffxl <old_file> <new_file>
```

### Arguments

* `old_file`: Path to the original file.
* `new_file`: Path to the new file.
* `--key`, `-k`: The column name to use as the unique identifier (default: **column** **furthest to the left in the identified table**).
* `--sheet`, `-s`: (Optional) Specific sheet name to compare.
* `--output`, `-o`: Output filename or path for the xlsx report(default: `diff_report.xlsx`). html report will be saved in the same directory with the same name.
* `--prefix`, `-p`: Add a prefix to output filenames (e.g., `ABC_diff_report.xlsx`) to keep track of multiple runs.
* `--raw`: Perform exact string comparison (disable smart normalization like treating `NaN` as equal to `None`).
* `--no-web`: Disable HTML report generation.
* `--diagnostic`, `-d`: Generate a detailed HTML diagnostic report if validation fails.
* `--dedup`: Remove duplicate rows based on Key column (keeps first occurrence).

### Example

```bash
diffxl samples/valvelist_v1.xlsx samples/valvelist_v2.xlsx --prefix "valves_v1_v2_"
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
