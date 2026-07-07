# Changelog

All notable changes to ModernJsonInVBA are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project uses
[semantic versioning](https://semver.org/spec/v2.0.0.html).

## [3.5.0] - 2026-07-07

### Changed

- The three ListObject upsert entry points return the table they created or
  updated: `Excel_UpsertListObjectOnSheet`,
  `Excel_UpsertListObjectFromJsonAtRoot`, and `Excel_UpsertListObjectFromSource`
  are now `Function ... As ListObject` instead of `Sub`. The returned reference
  lets a caller style, read, or extend the table without a second
  `ws.ListObjects(name)` lookup, matching how `Excel_EnsureListObject` and
  Excel's own `ListObjects.Add` return what they create. Backward compatible: a
  `Function` can be called as a statement, so existing calls that ignore the
  return keep compiling, and the error numbers are unchanged.

## [3.4.0] - 2026-07-07

### Added

- `Json_TryParse`: non-raising parse. Returns `True` and fills the output value
  on success; returns `False` on malformed input, with the output value set to
  `Null` and an optional `outError` holding the position-aware reason (the same
  message `Json_Parse` raises). For JSON of uncertain provenance (an API
  response that might be an error page, pasted input, an untrusted file), the
  caller branches on the return value instead of installing an error handler.
  The raising `Json_Parse` and `Json_ParseInto` are unchanged. Host-agnostic,
  so it is in both single-file builds.

## [3.3.0] - 2026-07-07

### Added

- `Json_StringifyPretty`: serialize the model to indented ("pretty") JSON text,
  one member or element per line, with empty objects and arrays kept inline as
  `{}` and `[]`. The indent unit defaults to two spaces and can be any string
  (pass `vbTab` for tabs). Escaping, number formatting, and error numbers match
  `Json_Stringify`, so the output parses back to an identical model. The compact
  `Json_Stringify` is unchanged. Host-agnostic, so it is in both single-file
  builds.

## [3.2.0] - 2026-07-06

### Added

- `NdjsonToJson` and `NdjsonFileToJson`: convert NDJSON (newline-delimited
  JSON, also called JSON Lines) into the library's JSON array. Each line
  becomes one record, so the result feeds straight into `Json_Parse` or the
  table upsert (one line to a row). Line endings `\n`, `\r\n`, and `\r` are
  accepted and blank lines are skipped. Host-agnostic, so it is in both
  single-file builds.

## [3.1.0] - 2026-07-06

### Added

- `Excel_RangeToJson`: convert a worksheet range to a JSON array-of-objects,
  the same way `Excel_ListObjectToJson` converts a table. The first row
  supplies the property names unless `hasHeaderRow` is False, in which case
  columns are named Column1, Column2, and so on. It shares the export engine
  with `Excel_ListObjectToJson`, so a range holding the same data produces
  identical JSON. Available in the workbook and the Excel single-file build.

## [3.0.0] - 2026-07-05

This release reorganizes the library into eleven modules split by concern and
rewrites the parser and Excel ingestion path for speed. The public API is
unchanged, so existing calling code runs without edits. The module split is
the only breaking change: you import eleven files instead of pasting one.

### Breaking

- The library ships as eleven `.bas` modules (in `vba_source/`) instead of the
  single `zz_ModernJsonInVBA` module. Remove the old module and import all
  eleven. Function names, arguments, return shapes, and error numbers are
  unchanged, so calling code does not change.

### Added

- `json_payloads/`: a seeded payload generator (`generate_payloads.py`) and a
  workbook macro suite (`Run_JsonPerfSuite`, `Run_JsonPerfMatrix`) that print
  Markdown timing reports to the Immediate window.
- `PERFORMANCE.md`: a measured timing table across the JSON parser and Excel
  ListObject surface, regenerable from the workbook.
- `vba_source/`: the eleven modules as individual files for import and version
  control.

### Changed

- Module layout by concern: `Json_Common` (shared plumbing), `Json_Parser`,
  `Json_Serializer`, `Json_Model`, `Json_Transforms`, `Json_Tables`,
  `Json_Coalesce`, `Json_Csv`, `Json_Xml`, `Json_Excel` (table ingestion), and
  `Json_Excel_Export` (table and range export). The README lists what each
  module owns.
- Rewrote the module and README comments and removed dead procedures left over
  from the single-module layout.

### Fixed

- The 32-bit FNV hash used `LongLong`, which does not compile on 32-bit Office.
  It is replaced with a Long-only rolling hash, so every module now compiles on
  both 32-bit and 64-bit Office.
- CSV conversion escapes control characters, so `CsvTextToJson` output always
  parses back through `Json_Parse`. Previously a tab or other control byte in a
  field produced JSON that the parser rejected.

### Performance

Measured against the previous single-module release on a Ryzen 7 9800X3D with
64-bit Excel. Times are hardware dependent; see `PERFORMANCE.md` for the full
table and method.

- Table ingestion streams JSON text directly into a 2D array for the common
  `tableRoot = "$"` case and does not build the intermediate object model. A
  per-row key cache reuses the column layout from one row to the next.
- Large table writes go to the sheet in one block. The former 50,000-cell
  chunking is removed; a block-write fallback remains for memory-constrained
  hosts.
- Character scanning reads UTF-16 code units from a byte snapshot of the input
  rather than allocating a one-character string per position.
- JSON parsing is about 4x faster. Table-row extraction and CSV conversion,
  which previously scaled quadratically on large inputs, are about an order of
  magnitude faster.
- A 500,000-row, 110 MB document loads into a ListObject in about 18 seconds on
  the benchmark machine.

[3.5.0]: https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/tag/v3.5.0
[3.4.0]: https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/tag/v3.4.0
[3.3.0]: https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/tag/v3.3.0
[3.2.0]: https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/tag/v3.2.0
[3.1.0]: https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/tag/v3.1.0
[3.0.0]: https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/tag/v3.0.0
