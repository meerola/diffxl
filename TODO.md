# diffxl Roadmap & Tasks

## Tasks

- [ ] Confirm the tests do what they should.
- [ ] Clean up the CLI messages. Too much happening by default. Add --verbose flag that could show the stats in the terminal, and not open the html report for example.
- [ ] html report: Add possibility to only see rows that have changes in a specific column.
- [ ] Add support for "Column Mapping" (if column names changed). UID column can be named differently between the files.
- [ ] Improve the xlsx report. E.g. there's no way to see the full diff of a row with new and old values side by side.
- [ ] Implement column header cleanup before the diff is computed. E.g. remove leading/trailing whitespace, convert to lowercase, remove excel line breaks '\n' that might cause a column mismatch.
- [ ] Add support for "Threshold" comparison for numeric values (ignore small floating point diffs).

## Future

- [ ] diffxl could accept config files to define what files to compare with what arguments.
- [ ] Add some folder watching functionality that can run diffxl automatically based on file changes. --> Always have an updated comparison available automagically.
- [ ] Improve the "diffability" of files without having to modify the source files. Foe example, diffxl could combine two columns to create a UID column (Think line list with same tag spanning two drawings: tag is not unique, but tag+dwg must be unique)
