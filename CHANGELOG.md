# Changelog

All notable changes to ModernJsonInVBA are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project uses
[semantic versioning](https://semver.org/spec/v2.0.0.html).

## [3.8.2] - 2026-07-09

### Changed

- Every library module header and both single-file builds now carry the
  release version and date. `build_dist.py` stamps them from the top entry
  of this changelog, so an imported module identifies which release it came
  from and the stamps cannot drift from the release notes.

## [3.8.1] - 2026-07-08

### Fixed

- An empty result (an empty JSON array, or a null table root) no longer
  disturbs an existing table's layout. Previously the internal `["value"]`
  placeholder schema leaked into reconciliation, so a default refresh
  appended a spurious `value` column to the table (and an empty append did
  the same instead of being a no-op). Now: refresh clears the rows and
  leaves the schema untouched, append changes nothing, and the
  `removeMissingColumns` combination keeps its schema-preserving clear.
- The `value` placeholder no longer sticks to a table once real data
  arrives. The placeholder appears only when an empty result must create a
  brand-new table (a ListObject cannot have zero columns); the first result
  with real headers now replaces that zero-row placeholder schema instead
  of merging with it. A single-column `value` table that holds rows is
  treated as user data and merges normally.

## [3.8.0] - 2026-07-07

### Added

- Path segments in `tableRoot`, `Json_TryResolvePath`, and the coalesce
  paths accept the column-header escape convention: `\.` addresses a key
  containing a literal dot and `\\` a key containing a literal backslash,
  so `$.a\.b.items` reaches rows under the key `a.b`. The streamed import
  and the model path resolve escapes identically; escape plus a bracket
  index (`$.a\.b[0]`) resolves through the model path. This closes the
  limitation documented in 3.7.0: previously a dotted key was unreachable
  by any path.

  Behavior note: a path containing the literal character sequences `\.` or
  `\\` previously matched keys containing those raw characters; those
  sequences now mean the escape. Paths without backslashes are unchanged.

## [3.7.0] - 2026-07-07

### Added

- `Json_ReadTextFile`: read a text file into a VBA string with encoding
  detection. UTF-16 LE/BE are recognized by BOM; everything else goes through
  a strict pure-VBA UTF-8 decoder (BOM stripped if present) and falls back to
  the system ANSI codepage when the bytes are not valid UTF-8, which keeps
  legacy ANSI exports readable. No Declare statements, so it works on every
  VBA host, Mac included. Host-agnostic, so it is in both single-file builds.
- `CONFORMANCE.md`: results of running the parser against JSONTestSuite, the
  standard RFC 8259 conformance corpus. All 95 valid documents accepted, all
  176 invalid documents rejected, no crashes; deep-nesting attacks reject
  through a trappable error. Implementation-defined choices are documented.

### Fixed

- `CsvFileToJson`, `XmlFileToJson`, and `NdjsonFileToJson` read files through
  `Json_ReadTextFile` instead of the ANSI-only `Input$` path. UTF-8 files with
  accented characters, CJK text, or emoji previously decoded as mojibake;
  UTF-16 files decoded as garbage. Pure-ASCII files (and legacy ANSI files)
  read exactly as before.

## [3.6.1] - 2026-07-07

### Performance

- Table and range export is 10 to 19 percent faster (measured on 22 MB mixed
  and 27 MB numeric exports). Two changes: integral cell numbers, which Excel
  hands over as Double, print through the integer formatter instead of the
  floating-point formatter, and the member separator is folded into each
  column's cached key prefix, halving builder calls per cell. The integer
  fast path also applies to `Json_Stringify` on integral Doubles. Output is
  byte-identical: both benchmark exports were compared byte for byte before
  and after.
- `PERFORMANCE.md` and the README timing table are regenerated from the
  current build, so the published numbers now include the row-schema cache,
  nested-root streaming, and this release's export changes.

## [3.6.0] - 2026-07-07

### Changed

- `Excel_UpsertListObjectFromJsonAtRoot`: `tableRoot` is now optional and
  defaults to `"$"` (the document root array-of-objects), matching
  `Excel_UpsertListObjectFromSource`. The common call drops from five required
  arguments to four:
  `Excel_UpsertListObjectFromJsonAtRoot(ws, name, cell, json)`. Backward
  compatible: callers that pass `tableRoot` explicitly are unchanged.

### Performance

- Table imports with a nested `tableRoot` (for example `"$.data.items"`, the
  usual shape of an API response) now use the streaming path that previously
  applied only to `"$"`: the reader descends to the table root and skips
  sibling members with validating scanners instead of building the whole
  document into the object model. A 150,000-row, 13.8 MB nested import drops
  from 4.1 s to 2.4 s on the benchmark machine. Skipped text is validated
  with the same rules and error numbers as before, so malformed documents,
  missing roots, and non-array roots raise exactly as they did; bracket-index
  paths (`"$.arr[0]"`) keep resolving through the model path.

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

[3.8.2]: https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/tag/v3.8.2
[3.8.1]: https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/tag/v3.8.1
[3.8.0]: https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/tag/v3.8.0
[3.7.0]: https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/tag/v3.7.0
[3.6.1]: https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/tag/v3.6.1
[3.6.0]: https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/tag/v3.6.0
[3.5.0]: https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/tag/v3.5.0
[3.4.0]: https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/tag/v3.4.0
[3.3.0]: https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/tag/v3.3.0
[3.2.0]: https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/tag/v3.2.0
[3.1.0]: https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/tag/v3.1.0
[3.0.0]: https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/tag/v3.0.0
