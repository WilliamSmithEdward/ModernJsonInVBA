# Getting Started with ModernJsonInVBA

Ten short examples, from a two-row table to nested API envelopes. Each one
shows a JSON payload, the one VBA call that loads it, and the Excel table
that call produces, drawn as a table so you can check your result against
the page. Read it top to bottom in about ten minutes, or jump to the shape
that matches your data.

## Contents

- [Before you start](#before-you-start)
- **Part 1: Your first tables**
  - [1. A flat array becomes a table](#1-a-flat-array-becomes-a-table)
  - [2. Rows with different keys](#2-rows-with-different-keys)
  - [3. What each JSON type becomes in a cell](#3-what-each-json-type-becomes-in-a-cell)
- **Part 2: Real-world payloads**
  - [4. Nested objects become dotted columns](#4-nested-objects-become-dotted-columns)
  - [5. The table is inside an API envelope](#5-the-table-is-inside-an-api-envelope)
  - [6. Arrays inside rows](#6-arrays-inside-rows)
- **Part 3: Keeping a table current**
  - [7. Refresh vs append](#7-refresh-vs-append)
- **Part 4: Excel back to JSON**
  - [8. Export a table to JSON](#8-export-a-table-to-json)
  - [9. Export a plain range](#9-export-a-plain-range)
- **Part 5: Beyond JSON**
  - [10. CSV, NDJSON, and files](#10-csv-ndjson-and-files)
- [Cheat sheet](#cheat-sheet)
- [Where next](#where-next)

---

## Before you start

**Install (30 seconds).** Download `ModernJsonInVBA_Excel.bas` from the
[latest release](https://github.com/WilliamSmithEdward/ModernJsonInVBA/releases/latest),
then in Excel press `ALT+F11`, right-click your VBA project, choose
`Import File...`, and pick the file. One module, no references, no setup.

**Every example assumes a worksheet variable:**

```vba
Dim ws As Worksheet
Set ws = ThisWorkbook.Worksheets("Sheet1")
```

**About JSON in VBA strings.** VBA doubles quotes inside string literals, so
`{"id":1}` is written `"{""id"":1}"`. Example 1 shows the literal once; after
that, examples show the JSON in its own block for readability. Loading the
JSON from a file or an HTTP response avoids the escaping entirely:

```vba
json = Json_ReadTextFile("C:\data\payload.json")   ' any encoding, BOM handled
```

---

## Part 1: Your first tables

### 1. A flat array becomes a table

A JSON array of objects: each object becomes one row.

The JSON:

```json
[
  { "id": 1, "name": "Alice" },
  { "id": 2, "name": "Bob" }
]
```

The call (shown once with VBA's doubled quotes):

```vba
Dim json As String
json = "[{""id"":1,""name"":""Alice""},{""id"":2,""name"":""Bob""}]"

Excel_UpsertListObjectFromJsonAtRoot ws, "People", ws.Range("A1"), json
```

The result, a real ListObject named `People` anchored at A1:

| id | name |
|---:|---|
| 1 | Alice |
| 2 | Bob |

Keys become column headers, in the order they first appear. Run the call
again and the table refreshes in place instead of duplicating.

### 2. Rows with different keys

Rows do not need identical keys. The columns are the union of every key
seen, still in first-seen order, and a row without a key gets a blank
cell.

The JSON:

```json
[
  { "id": 1, "name": "Alice" },
  { "id": 2, "email": "bob@example.com" }
]
```

```vba
Excel_UpsertListObjectFromJsonAtRoot ws, "People", ws.Range("A1"), json
```

The result:

| id | name | email |
|---:|---|---|
| 1 | Alice |  |
| 2 |  | bob@example.com |

`name` was seen first (row 1), so it comes before `email` (first seen in
row 2). The order is deterministic: the same JSON always produces the same
columns.

### 3. What each JSON type becomes in a cell

The JSON:

```json
[
  { "item": "bolt", "qty": 40, "price": 0.25, "inStock": true, "note": null }
]
```

```vba
Excel_UpsertListObjectFromJsonAtRoot ws, "Inventory", ws.Range("A1"), json
```

The result:

| item | qty | price | inStock | note |
|---|---:|---:|---|---|
| bolt | 40 | 0.25 | TRUE |  |

- Strings stay text, numbers stay numeric (real Excel numbers you can sum).
- `true` / `false` become Excel `TRUE` / `FALSE`.
- `null` becomes a blank cell, same as a missing key. Example 8 shows how
  the export side lets you choose between omitting blanks and emitting
  `null`.

---

## Part 2: Real-world payloads

### 4. Nested objects become dotted columns

A nested object does not get crammed into one cell. Each leaf becomes its
own column, named by its path.

The JSON:

```json
[
  { "id": 1, "customer": { "name": "Ada",  "city": "Austin" } },
  { "id": 2, "customer": { "name": "Grace", "city": "Boston" } }
]
```

```vba
Excel_UpsertListObjectFromJsonAtRoot ws, "Orders", ws.Range("A1"), json
```

The result:

| id | customer.name | customer.city |
|---:|---|---|
| 1 | Ada | Austin |
| 2 | Grace | Boston |

Nesting can go as deep as your payload does; a
`{"a":{"b":{"c":1}}}` leaf becomes column `a.b.c`. The dotted names matter
on the way back out: example 8 rebuilds the nesting from them.

A key that itself contains a dot stays one column: `{"a.b": 1}` gets the
header `a\.b` (the dot is escaped), so it never collides with real nesting,
and it exports back as the original `"a.b"` key.

### 5. The table is inside an API envelope

When the rows sit under a key, with metadata alongside, point `tableRoot`
at the array instead of reshaping the JSON yourself.

The JSON:

```json
{
  "meta": { "page": 1, "perPage": 50 },
  "data": {
    "items": [
      { "id": 101, "status": "shipped" },
      { "id": 102, "status": "pending" }
    ]
  },
  "ok": true
}
```

```vba
Excel_UpsertListObjectFromJsonAtRoot ws, "Orders", ws.Range("A1"), json, "$.data.items"
```

The result:

| id | status |
|---:|---|
| 101 | shipped |
| 102 | pending |

`tableRoot` is a simple path from the document root: `$` is the root itself
(and the default when you omit the argument), `$.data.items` descends two
keys. A key containing a literal dot is escaped the same way as in column
headers: `$.a\.b.items` walks the key `a.b`. The siblings (`meta`, `ok`)
are validated and skipped, and the import streams at the same speed as a
bare array.

### 6. Arrays inside rows

A scalar list inside a row (tags, ids) has no single tabular answer, so you
choose. By default it is left out:

The JSON:

```json
[
  { "id": 1, "name": "Alice", "tags": ["admin", "ops"] },
  { "id": 2, "name": "Bob",   "tags": ["dev"] }
]
```

```vba
Excel_UpsertListObjectFromJsonAtRoot ws, "People", ws.Range("A1"), json
```

| id | name |
|---:|---|
| 1 | Alice |
| 2 | Bob |

Or keep each array as JSON text in its cell:

```vba
Excel_UpsertListObjectFromJsonAtRoot ws, "People", ws.Range("A1"), json, _
    nonTableArraysAsJson:=True
```

| id | name | tags |
|---:|---|---|
| 1 | Alice | ["admin","ops"] |
| 2 | Bob | ["dev"] |

The cell text is valid JSON, so it round-trips: on export (example 8),
`parseJsonInCells:=True` turns it back into a real array.

---

## Part 3: Keeping a table current

### 7. Refresh vs append

The default is a refresh: rows are replaced, the schema is reconciled, and
formula columns you added by hand survive. Pass `clearExisting:=False` to
append instead.

Start with example 1's table, then run:

```json
[
  { "id": 3, "name": "Cy" }
]
```

```vba
Excel_UpsertListObjectFromJsonAtRoot ws, "People", ws.Range("A1"), json, _
    clearExisting:=False
```

The result:

| id | name |
|---:|---|
| 1 | Alice |
| 2 | Bob |
| 3 | Cy |

New keys in appended rows add columns automatically
(`addMissingColumns:=True` is the default); the README's
[Schema Control](README.md#schema-control) section covers the strict-schema
switches when you want a feed to fail loudly instead of drifting.

---

## Part 4: Excel back to JSON

### 8. Export a table to JSON

`Excel_ListObjectToJson` walks the table and rebuilds JSON, including the
nesting encoded in dotted headers. Exporting example 4's table:

| id | customer.name | customer.city |
|---:|---|---|
| 1 | Ada | Austin |
| 2 | Grace | Boston |

```vba
Dim js As String
js = Excel_ListObjectToJson(ws.ListObjects("Orders"))
```

produces:

```json
[{"id":1,"customer":{"name":"Ada","city":"Austin"}},{"id":2,"customer":{"name":"Grace","city":"Boston"}}]
```

Blank cells are omitted from the row's object by default; pass
`includeBlanksAsNull:=True` to emit `"key": null` instead. For output meant
for human eyes, wrap it:

```vba
Debug.Print Json_StringifyPretty(Json_Parse(js))
```

```json
[
  {
    "id": 1,
    "customer": {
      "name": "Ada",
      "city": "Austin"
    }
  }
]
```

### 9. Export a plain range

No table required: any rectangular range with a header row exports the same
way, through the same engine.

The cells (`A1:C3`):

| region | units | rep |
|---|---:|---|
| East | 120 | Ada |
| West | 95 | Bob |

```vba
js = Excel_RangeToJson(ws.Range("A1:C3"))
```

```json
[{"region":"East","units":120,"rep":"Ada"},{"region":"West","units":95,"rep":"Bob"}]
```

Pass `hasHeaderRow:=False` when there is no header row; columns are then
named `Column1`, `Column2`, and so on.

---

## Part 5: Beyond JSON

### 10. CSV, NDJSON, and files

**CSV** routes through the same pipeline. The first record supplies the
headers. CSV carries no types, so the converted JSON holds every field as a
string; when the values land on the sheet, Excel applies its usual entry
conversion, so numeric-looking fields such as `id` become real numbers:

```vba
Dim csv As String
csv = "id,name" & vbCrLf & "1,Alice" & vbCrLf & "2,Bob"

Excel_UpsertListObjectFromSource ws, "FromCsv", ws.Range("A1"), csv, ExcelSourceFormat_CSV
```

| id | name |
|---:|---|
| 1 | Alice |
| 2 | Bob |

**NDJSON** (one JSON value per line) converts to a JSON array with one
call, so it feeds anything above:

```vba
Excel_UpsertListObjectFromJsonAtRoot ws, "Events", ws.Range("A1"), NdjsonToJson(text)
```

**Files** of any of these formats load with one call each: `CsvFileToJson`,
`XmlFileToJson`, `NdjsonFileToJson`, or `Json_ReadTextFile` for raw text.
Encoding is detected automatically (UTF-8 with or without BOM, UTF-16,
legacy ANSI), so accented characters and emoji arrive intact.

**HTTP** responses drop straight in; the README's
[HTTP Helper](README.md#http-helper-windows) section has a copy-paste
`HttpGetText` function (Windows and Mac variants):

```vba
Excel_UpsertListObjectFromJsonAtRoot ws, "Api", ws.Range("A1"), HttpGetText(url), "$.data.items"
```

---

## Cheat sheet

| You have | You want | Call |
|---|---|---|
| JSON array text | an Excel table | `Excel_UpsertListObjectFromJsonAtRoot ws, name, cell, json` |
| JSON with the array under a key | an Excel table | same call with `"$.data.items"` as `tableRoot` |
| a JSON file | an Excel table | `... jsonText:=Json_ReadTextFile(path)` |
| CSV or XML text | an Excel table | `Excel_UpsertListObjectFromSource ... ExcelSourceFormat_CSV / _XML` |
| NDJSON text or file | a JSON array | `NdjsonToJson(text)` / `NdjsonFileToJson(path)` |
| a ListObject | JSON | `Excel_ListObjectToJson(lo)` |
| a plain range | JSON | `Excel_RangeToJson(rng)` |
| JSON of unknown validity | a safe parse | `If Json_TryParse(text, model, why) Then ...` |
| a parsed model | one value by path | `Json_TryResolvePath model, "$.a.b[0]", out` |
| a parsed model | readable JSON text | `Json_StringifyPretty(model)` |

Every upsert returns the `ListObject` it created or updated, so you can
style it without a second lookup:

```vba
With Excel_UpsertListObjectFromJsonAtRoot(ws, "People", ws.Range("A1"), json)
    .TableStyle = "TableStyleMedium2"
    .Range.Columns.AutoFit
End With
```

## Where next

- [README](README.md): the full function reference, `tableRoot` semantics,
  schema control, and the determinism guarantees.
- [PERFORMANCE.md](PERFORMANCE.md): measured timings up to a 500,000-row,
  110 MB document, regenerable with one macro.
- [CONFORMANCE.md](CONFORMANCE.md): the JSONTestSuite results behind the
  badge.
