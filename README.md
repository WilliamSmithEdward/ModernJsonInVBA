# ModernJsonInVBA

[![GitHub stars](https://img.shields.io/github/stars/WilliamSmithEdward/ModernJsonInVBA)](https://github.com/WilliamSmithEdward/ModernJsonInVBA/stargazers)
[![Last commit](https://img.shields.io/github/last-commit/WilliamSmithEdward/ModernJsonInVBA)](https://github.com/WilliamSmithEdward/ModernJsonInVBA/commits/main)
[![MIT License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

**Deterministic JSON (and CSV / XML) → Excel Tables → JSON Roundtrip**  
\
Pure VBA. No dependencies. No silent schema drift.  
\
Take nested or complex API payloads and  convert them into normalized Excel tables with parent keys preserved. Supports Excel formula and scalar value injection.

------------------------------------------------------------------------

## Contents

- [Introduction](#introduction)
- [What This Solves](#what-this-solves)
- [Core Capabilities](#core-capabilities)
- [Performance](#performance)
  - [Loading JSON into Excel tables](#loading-json-into-excel-tables)
  - [Core JSON engine](#core-json-engine)
  - [Reproducing these numbers](#reproducing-these-numbers)
- [Installation](#installation)
  - [Option 1 — Import Into Your Workbook](#option-1--import-into-your-workbook)
  - [Option 2 — Use the Provided Workbook](#option-2--use-the-provided-workbook)
- [Requirements](#requirements)
- [Basic API Example](#basic-api-example)
  - [Refresh Mode](#refresh-mode)
  - [Append Mode](#append-mode)
  - [Strict Schema Mode](#strict-schema-mode)
  - [API Example With Nested Objects](#api-example-with-nested-objects)
- [Accessing Json Elements (Directly in VBA)](#accessing-json-elements-directly-in-vba)
- [Excel to JSON (Reverse Materialization)](#excel-to-json-reverse-materialization)
- [Understanding tableRoot](#understanding-tableroot)
- [HTTP Helper (Windows)](#http-helper-windows)
- [Schema Control](#schema-control)
- [Deterministic Errors](#deterministic-errors)
- [Function & Subroutine Reference](#function--subroutine-reference)

------------------------------------------------------------------------

## Introduction

**ModernJsonInVBA** is a pure-VBA JSON engine for structured Excel
workflows, organized as a small set of focused modules (parser,
serializer, transforms, tables, CSV/XML conversion, Excel integration).

It materializes JSON text into Excel tables (ListObjects) with
predictable, repeatable behavior.

The focus is control:

-   Explicit structural rules
-   Intentional schema changes
-   Deterministic failures
-   No hidden behavior

Excel becomes a controlled data surface rather than a loosely
interpreted one.

See the .XLSM file for:

-   Unit Tests
-   Examples
-   Performance Testing
-   A Quick Start Module
-   CSV / XML → Json Pipeline

------------------------------------------------------------------------

## What This Solves

Working with JSON in Excel commonly leads to:

-   Columns appearing in different orders
-   Tables silently changing shape
-   Layout drift over time
-   Hidden external dependencies
-   Fragile refresh logic

ModernJsonInVBA eliminates those risks through:

-   Stable column discovery order
-   Strict structural validation
-   Explicit schema controls
-   Deterministic error behavior

When JSON structure does not match the declared `tableRoot`, execution
stops with a clear, stable error.

No guessing.
No fallback tables.

------------------------------------------------------------------------

## Core Capabilities

-   Parse JSON into VBA Variants (objects, arrays, primitives)
-   Convert VBA structures back into JSON
-   Flatten and rebuild object graphs
-   Discover array-of-object roots
-   Convert JSON tables into 2D arrays
-   Upsert Excel ListObjects deterministically
-   Enforce strict schema contracts when required
-   Round trip json → list object → json
-   Emoji-Ready: Full support for non-BMP Unicode characters via surrogate pair parsing.
-   Memory Efficient: Linear-time string processing designed for high-volume data.
-   State-Machine Parsing: Handles nested arrays and objects to any depth without breaking.

All implemented in pure VBA.

-   No `Scripting.Dictionary`
-   No COM references
-   No external libraries

Zero Dependencies: No need for Scripting.Dictionary or external DLLs. It’s pure, portable VBA.

<pre>
JSON Text
   ↓
Parser
   ↓
Tagged Object Model
   ↓
Array-of-Objects Root
   ↓
2D Array Materialization
   ↓
Excel ListObject Upsert
</pre>
------------------------------------------------------------------------

## Performance

All figures below come from the reproducible benchmark suite in
[`json_payloads/`](json_payloads) — seven deterministic, generated payloads
totalling ~200 MB, driven by the `Run_JsonPerfSuite` macro in the workbook.

**Benchmark environment**

- AMD Ryzen 7 9800X3D, 64 GB RAM
- Excel 16.0, 64-bit VBA
- Payloads read from local disk; times are wall-clock via `Timer`

Everything is pure VBA — no `Scripting.Dictionary`, no COM references, no
external libraries.

### Loading JSON into Excel tables

Full path from a JSON file on disk to a populated, refreshable Excel
ListObject (`Excel_UpsertListObjectFromJsonAtRoot`). The **Create** column
is the one-shot load of a fresh table; **Refresh** re-loads the same table
in place (clear + rewrite). Rows/sec and cells/sec are for Create.

| Payload | Rows × Cols | File | Create | Refresh | Rows/sec | Cells/sec |
|---|---|---:|---:|---:|---:|---:|
| Flat, small | 10,000 × 10 | 2.1 MB | 0.44 s | 0.50 s | 22,700 | 227,000 |
| Flat, medium | 100,000 × 10 | 21.6 MB | 4.53 s | 5.06 s | 22,100 | 221,000 |
| **Flat, large** | **500,000 × 10** | **109.5 MB** | **23.5 s** | **26.1 s** | **21,200** | **212,000** |
| Numeric-heavy | 200,000 × 8 | 25.9 MB | 6.57 s | 7.59 s | 30,500 | 244,000 |
| Escape-dense | 50,000 × 6 | 14.8 MB | 2.43 s | 2.53 s | 20,600 | 124,000 |
| Nested objects | 50,000 × 9 | 9.2 MB | 2.92 s | 3.00 s | 17,100 | 154,000 |
| Wide schema | 5,000 × 200 | 16.0 MB | 4.04 s | 4.67 s | 1,240 | 248,000 |

> A **half-million-row, 110 MB** JSON document becomes a fully materialized
> Excel table in **~24 seconds** — the whole document is written to the
> sheet in a single block, with no chunking.

Cell throughput holds around **210,000–250,000 cells/sec** regardless of row
count; it dips only when a column is unusually escape-heavy (the escape
payload is deliberately hostile — every string carries quotes, backslashes,
newlines, and emoji). Nested payloads reconstruct dotted column paths
(`customer.address.city`) and still clear 150,000 cells/sec.

### Core JSON engine

Standalone engine operations (no worksheet involved), measured on the same
payloads:

| Operation | Throughput | Notes |
|---|---|---|
| `Json_Parse` → model | 6–11 MB/s | full recursive-descent parse into the tagged-Collection model |
| `Json_Stringify` ← model | 5–9 MB/s | model back to compact JSON text |
| `Excel_ListObjectToJson` | 3–13 MB/s | table → JSON array-of-objects |

For example, `Json_Parse` turns the 21.6 MB / 100,000-row document into the
in-memory model in **2.4 s**, and `Json_Stringify` serializes it back in
**2.8 s**. The table upsert path is faster than a raw parse because it
streams JSON text straight into a 2-D array for the common root table,
bypassing model construction entirely.

### Reproducing these numbers

```text
1. python json_payloads/generate_payloads.py     ' generate the payloads
2. Open ModernJsonInVBA.xlsm, press ALT+F11, then Ctrl+G (Immediate window)
3. Run_JsonPerfSuite
```

The macro prints a Markdown report (like the tables above) to the Immediate
window, ready to paste into an issue or PR. Drop your own `*.json` files
into `json_payloads/` and they are picked up automatically.

## Installation

ModernJsonInVBA supports two usage models:

-   Import the modules into your own workbook (recommended)
-   Use the provided `.xlsm` file directly

### Option 1 — Import Into Your Workbook

1.  Download the `.bas` files from the `vba_source/` folder
2.  Open your Excel workbook
3.  Press `ALT + F11`
4.  File → Import File… and import **all eleven** modules
5.  Save as `.xlsm`

The library is split by concern; all eleven modules are required:

| Module              | Responsibility                                    |
|---------------------|---------------------------------------------------|
| `Json_Common`       | Shared plumbing: tag constant, path utilities, hash index, string builder |
| `Json_Parser`       | JSON text → in-memory model (`Json_Parse`)        |
| `Json_Serializer`   | Model → JSON text (`Json_Stringify`)              |
| `Json_Model`        | Type checks, object get/set, path resolution      |
| `Json_Transforms`   | Flatten / unflatten / array-root discovery        |
| `Json_Tables`       | Table extraction, 2D conversion                   |
| `Json_Coalesce`     | Merging JSON arrays (child hoisting, concatenation) |
| `Json_Csv`          | CSV → JSON conversion                             |
| `Json_Xml`          | XML → JSON conversion                             |
| `Json_Excel`        | ListObject ingestion (upsert pipeline)            |
| `Json_Excel_Export` | ListObject / range → JSON export                  |

The public API (every function documented below) is unchanged from the
earlier single-module releases — existing code keeps working as-is.

------------------------------------------------------------------------

### Option 2 — Use the Provided Workbook

1.  Download `ModernJsonInVBA.xlsm`
2.  Open the file
3.  Enable macros

You may export/import the modules into another workbook if needed.

------------------------------------------------------------------------

## Requirements

-   Excel with VBA support (Windows and macOS)
-   Macros enabled
-   No external references required

------------------------------------------------------------------------

## Basic API Example

Endpoint used:

https://jsonplaceholder.typicode.com/users

This endpoint returns a root array. `tableRoot` is `$`.

### Refresh Mode

``` vba
Public Sub Example_Api_Refresh()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Quick Start")

    Dim jsonText As String
    jsonText = HttpGetText("https://jsonplaceholder.typicode.com/users")

    'Clear existing values, add missing columns from JSON, preserve columns not found in JSON, preserve existing formulas
    Excel_UpsertListObjectFromJsonAtRoot _
        ws, "tUsers", ws.Range("A1"), _
        jsonText, "$", _
        True, True, False, True
    
    'Preserve existing values, append new values from JSON, ignore missing columns from JSON, preserve columns not found in JSON, preserve existing formulas, don't add formulas to newly appended data
    Excel_UpsertListObjectFromJsonAtRoot _
        ws, "tUsers2", ws.Range("A15"), _
        jsonText, "$", _
        False, True, False, True, False

End Sub
```

### Append Mode

``` vba
Excel_UpsertListObjectFromJsonAtRoot _
    ws, "tUsersLog", ws.Range("A1"), _
    jsonText, "$", _
    False, True, False
```

### Strict Schema Mode

``` vba
Excel_UpsertListObjectFromJsonAtRoot _
    ws, "tUsersStrict", ws.Range("A1"), _
    jsonText, "$", _
    True, False, True
```

------------------------------------------------------------------------

### API Example With Nested Objects

Also see: [API Product Intelligence Demo Case Study](https://github.com/WilliamSmithEdward/APIProductIntelligenceDemo)

``` vba

Public Sub Example_Api_Refresh()

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Quick Start")

    ' Fetch API
    Dim jsonText As String
    jsonText = HttpGetText("https://dummyjson.com/products")

    ' Products table
    Excel_UpsertListObjectFromJsonAtRoot _
        ws, "tProducts", ws.Range("A1"), _
        jsonText, "$.products", _
        True, True, False, True, True, True

    ' ---------------------------------------------
    ' Build key map for parent injections
    ' ---------------------------------------------
    Dim keyMap As New Collection
    keyMap.Add Array("id", "productId")

    ' Extract + merge reviews with parent key injected
    Dim reviewsJson As String
    reviewsJson = Json_CoalesceChildArrays( _
        jsonText, _
        "$.products", _
        "reviews", _
        True, _
        keyMap)

    ' Reviews table
    Excel_UpsertListObjectFromJsonAtRoot _
        ws, "tReviews", ws.Range("A35"), _
        reviewsJson, "$", _
        True, True, False, True, True, True

CleanExit:

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

CleanFail:
    Resume CleanExit

End Sub
```

------------------------------------------------------------------------

### HTTP Helper (Windows)

``` vba
Private Function HttpGetText(ByVal url As String) As String

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/json"
    http.send

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + 1500, "HttpGetText", _
            "HTTP " & http.Status & " " & http.statusText & " | " & url
    End If

    HttpGetText = CStr(http.responseText)

End Function
```

### HTTP Helper (Mac)

```vba
Private Function HttpGetTextMac(ByVal url As String) As String

    Dim cmd As String
    Dim result As String
    
    ' Use curl (installed on all modern macOS systems)
    cmd = "curl -s -L -H ""Accept: application/json"" """ & url & """"
    
    result = MacScript("do shell script " & Chr(34) & cmd & Chr(34))
    
    If Len(result) = 0 Then
        Err.Raise vbObjectError + 1500, "HttpGetTextMac", _
            "HTTP request returned empty response | " & url
    End If
    
    HttpGetTextMac = result

End Function
```

------------------------------------------------------------------------

## Accessing Json Elements (Directly in VBA)

```vba
' =============================================================================
' Example_ReadValuesFromJson
'
' Purpose
'   Demonstrates how to extract values from JSON using the ModernJsonInVBA library.
'
' What this example shows
'   1) Parse JSON text into an in-memory structure
'   2) Access top-level values
'   3) Access nested values
'   4) Iterate arrays
'   5) Read properties from objects
'
' Important idea
'   JSON contains three main structures:
'
'       Object  -> { key : value }
'       Array   -> [ value, value, value ]
'       Value   -> text, number, true/false, or null
'
'   In this library:
'
'       JSON Object  -> VBA Collection (tagged internally)
'       JSON Array   -> VBA Collection
'       Values       -> normal VBA types (String, Double, Boolean, etc.)
'
' =============================================================================
Public Sub Example_ReadValuesFromJson()

    ' ------------------------------------------------------------
    ' Step 1 — Create some example JSON
    ' (Normally this would come from an API response)
    ' ------------------------------------------------------------
    
    Dim jsonText As String
    
    jsonText = "{"
    jsonText = jsonText & """orders"":["
    
    jsonText = jsonText & "{"
    jsonText = jsonText & """orderId"":""A100"","
    jsonText = jsonText & """customer"":{""id"":""C01"",""name"":""Ada""},"
    jsonText = jsonText & """status"":""open"","
    jsonText = jsonText & """items"":["
    
    jsonText = jsonText & "{""sku"":""SKU-1"",""qty"":2,""price"":9.99,""promos"":[""P10"",""P20""]},"
    jsonText = jsonText & "{""sku"":""SKU-2"",""qty"":1,""price"":19.5,""promos"":[]}"
    
    jsonText = jsonText & "]"
    jsonText = jsonText & "},"
    
    jsonText = jsonText & "{"
    jsonText = jsonText & """orderId"":""A101"","
    jsonText = jsonText & """customer"":{""id"":""C02"",""name"":""Grace""},"
    jsonText = jsonText & """status"":""shipped"","
    jsonText = jsonText & """items"":["
    
    jsonText = jsonText & "{""sku"":""SKU-3"",""qty"":4,""price"":2.5,""promos"":[""P5""]}"
    
    jsonText = jsonText & "]"
    jsonText = jsonText & "}"
    
    jsonText = jsonText & "]"
    jsonText = jsonText & "}"
    
    
    ' ------------------------------------------------------------
    ' Step 2 — Parse the JSON text
    '
    ' Json_ParseInto converts the JSON string into an in-memory
    ' structure that VBA can navigate.
    ' ------------------------------------------------------------
    
    Dim root As Variant
    Json_ParseInto jsonText, root
    
    
    ' ------------------------------------------------------------
    ' Step 3 — Access a value using a JSON path
    '
    ' "$.orders[0].orderId"
    ' means:
    '
    ' root
    '   -> orders array
    '       -> first element
    '           -> orderId property
    ' ------------------------------------------------------------
    
    Dim orderIdV As Variant
    
    If Json_TryResolvePath(root, "$.orders[0].orderId", orderIdV) Then
        Debug.Print "First orderId:", orderIdV
    End If
    
    
    ' ------------------------------------------------------------
    ' Step 4 — Access nested values
    ' ------------------------------------------------------------
    
    Dim custName As Variant
    
    Json_TryResolvePath root, "$.orders[1].customer.name", custName
    
    Debug.Print "Second order customer:", custName
    
    
    ' ------------------------------------------------------------
    ' Step 5 — Extract the orders array
    '
    ' Arrays are represented as VBA Collections
    ' ------------------------------------------------------------
    
    Dim ordersV As Variant
    
    Json_TryResolvePath root, "$.orders", ordersV
    
    Dim orders As Collection
    Set orders = ordersV
    
    
    ' ------------------------------------------------------------
    ' Step 6 — Loop through each order
    ' ------------------------------------------------------------
    
    Dim orderObj As Collection
    
    For Each orderObj In orders
        
        Dim idV As Variant
        Json_TryObjGet orderObj, "orderId", idV
        
        Debug.Print "Order:", idV
        
        
        ' --------------------------------------------------------
        ' Step 7 — Access the items array inside each order
        ' --------------------------------------------------------
        
        Dim itemsV As Variant
        
        If Json_TryObjGet(orderObj, "items", itemsV) Then
            
            Dim items As Collection
            Set items = itemsV
            
            Dim itemObj As Collection
            
            For Each itemObj In items
                
                Dim skuV As Variant
                Json_TryObjGet itemObj, "sku", skuV
                
                Debug.Print "  SKU:", skuV
                
                
                ' ------------------------------------------------
                ' Step 8 — Access the promo codes array
                ' ------------------------------------------------
                
                Dim promosV As Variant
                
                If Json_TryObjGet(itemObj, "promos", promosV) Then
                    
                    Dim promos As Collection
                    Set promos = promosV
                    
                    Dim promo As Variant
                    
                    For Each promo In promos
                        Debug.Print "    Promo:", promo
                    Next promo
                    
                End If
                
            Next itemObj
            
        End If
        
    Next orderObj
    
    
End Sub
```

------------------------------------------------------------------------

## Excel to JSON (Reverse Materialization)

ModernJsonInVBA is not only JSON to Excel.

It also supports deterministic **Excel Table to JSON** conversion.

This enables:

- Exporting structured tables to APIs
- Serializing curated Excel datasets
- Creating reproducible JSON snapshots
- Round-trip validation workflows

### Function

```vb
Excel_ListObjectToJson(lo As ListObject, Optional includeBlanksAsNull As Boolean = False) As String
```

### Behavior

- Each table row becomes a JSON object
- Column order is preserved
- Row order is preserved
- Headers become property names
- Nested paths supported via dot notation (`a.b.c`)
- Literal dots supported via escape (`a\.b`)
- Array index paths (`[0]`) are intentionally rejected (error 905)
- Blank cells:
  - Skipped by default (key omitted)
  - Optional `includeBlanksAsNull=True` to emit explicit `null`

### Example

Given a table:

| id | name    | active |
|---:|---------|:------:|
| 1  | Alice   | TRUE   |
| 2  | Bob     | FALSE  |
| 3  | Charlie | TRUE   |

```vb
Dim jsonText As String
jsonText = Excel_ListObjectToJson(lo)
```

Produces:

```json
[
  {"id":1,"name":"Alice","active":true},
  {"id":2,"name":"Bob","active":false},
  {"id":3,"name":"Charlie","active":true}
]
```

### Determinism Guarantees

- No silent type coercion
- No hidden schema mutation
- Excel errors (`#N/A`, etc.) trigger stable error 1170
- Duplicate headers trigger 1121
- Blank headers trigger 1120

### Why This Matters

Most VBA JSON libraries only parse JSON.

ModernJsonInVBA supports **bidirectional structured transformation**:

```text
JSON Text
   ↓
Excel Table
   ↓
JSON Text
```

Excel becomes a structured JSON surface, not just a spreadsheet.

## Understanding `tableRoot`

`tableRoot` defines which portion of the JSON becomes the Excel table.

It must resolve to:

-   An array of objects
-   Or `null` (treated as zero rows)

Anything else triggers a deterministic error.

Supported path patterns:

-   `$`
-   `$.property`
-   `$.property.child`
-   `$.array[0].items`

Zero-based indexing inside brackets.

------------------------------------------------------------------------

## Schema Control

Three switches govern update behavior:

-   `clearExisting`
-   `addMissingColumns`
-   `removeMissingColumns`

### Recommended Default

    True, True, False

Rows refresh.
New columns allowed.
Columns never disappear.

------------------------------------------------------------------------

## Deterministic Errors

The engine stops execution on structural violations.

Common cases:

-   `tableRoot` not found → 1160
-   `tableRoot` not array-of-objects → 1162 / 1163
-   Duplicate headers → 1121
-   Blank headers → 1120
-   Invalid flag combination → 1101

Errors protect against:

-   Silent schema drift
-   Column collapse
-   Partial table corruption
-   Ambiguous data states

------------------------------------------------------------------------

## Function / Subroutine Reference

### Enums

- `Public Enum ExcelSourceFormat`  
  Source format selector for `Excel_UpsertListObjectFromSource`.

  - `ExcelSourceFormat_JSON = 1`  
    Treat input as JSON text.
  - `ExcelSourceFormat_CSV = 2`  
    Treat input as CSV text.
  - `ExcelSourceFormat_XML = 3`  
    Treat input as XML text.

### JSON Parse / Serialize

- `Public Function Json_Parse(ByVal jsonText As String) As Variant`  
  Parses JSON text into the module’s in-memory model. Returns a tagged `Collection` for objects, an untagged `Collection` for arrays, or a primitive `Variant` for scalar values.

- `Public Sub Json_ParseInto(ByVal jsonText As String, ByRef outValue As Variant)`  
  Parses JSON text into an output `Variant` passed by reference.

- `Public Function Json_Stringify(ByVal v As Variant) As String`  
  Serializes the in-memory JSON model back into deterministic JSON text.

## JSON Type / Object Helpers

- `Public Function Json_IsObject(ByVal v As Variant) As Boolean`  
  Returns `True` when `v` is a tagged JSON object in this library’s model.

- `Public Function Json_IsArray(ByVal v As Variant) As Boolean`  
  Returns `True` when `v` is an untagged JSON array in this library’s model.

- `Public Sub Json_ObjSet(ByVal obj As Collection, ByVal key As String, ByVal value As Variant)`  
  Sets or overwrites a property on a tagged JSON object while preserving deterministic order.

- `Public Function Json_ObjGet(ByVal obj As Collection, ByVal key As String) As Variant`  
  Gets a property value from a tagged JSON object. Raises an error if the key is missing.

- `Public Function Json_TryObjGet(ByVal obj As Collection, ByVal key As String, ByRef outValue As Variant) As Boolean`  
  Attempts to get a property value from a tagged JSON object without raising when the key is missing.

### Flatten / Unflatten / Path Utilities

- `Public Function Json_Flatten(ByVal parsedJson As Variant, Optional ByVal maxDepth As Long = 12, Optional ByVal tableRootToExpand As String = vbNullString, Optional ByVal arrayMode As Long = 0) As Collection`  
  Flattens nested JSON into a tagged object of `[path, value]` pairs for inspection, extraction, and table shaping.

- `Public Function Json_FlatGet(ByVal flatObj As Collection, ByVal path As String) As Variant`  
  Returns the primitive value stored at an exact flattened path.

- `Public Function Json_FlatContains(ByVal flatObj As Collection, ByVal path As String) As Boolean`  
  Returns `True` if the flattened object contains an exact path.

- `Public Function Json_Unflatten(ByVal flatObj As Collection) As Collection`  
  Reconstructs a nested tagged object from flattened `[path, value]` pairs. Array index paths are not supported.

- `Public Function Json_TryResolvePath(ByVal root As Variant, ByVal path As String, ByRef outValue As Variant) As Boolean`  
  Resolves a minimal JSONPath-like expression such as `$.orders[0].id`.

- `Public Function Json_TryReadBracketIndex(ByVal path As String, ByRef i As Long, ByRef outIndex As Long) As Boolean`  
  Low-level helper that parses a bracket index like `[3]` from a path string.

- `Public Function Json_ResolveArrayPath(ByVal root As Variant, ByVal path As String) As Collection`  
  Resolves a simple path like `$.products` directly to a JSON array collection.

### Array-of-Object Discovery / Table Extraction

- `Public Function Json_FindArrayObjectRoots(ByVal flatObj As Collection, Optional ByVal stopAfterFirst As Boolean = False) As Collection`  
  Scans flattened paths and returns candidate roots that look like array-of-object tables.

- `Public Function Json_ExtractTableRows(ByVal flatObj As Collection, ByVal tableRoot As String) As Collection`  
  Extracts tagged row objects from flattened JSON for a specific table root.

- `Public Function Json_TableTo2D(ByVal rows As Collection, ByRef headers As Variant) As Variant`  
  Converts a collection of tagged row objects into a 2D Excel-ready array and outputs headers in first-seen order.

### Excel ListObject Helpers / Upsert Pipeline

- `Public Function Excel_GetListObject(ByVal ws As Worksheet, ByVal tableName As String) As ListObject`  
  Finds a `ListObject` by name on a worksheet.

- `Public Function Excel_EnsureListObject(ByVal ws As Worksheet, ByVal tableName As String, ByVal topLeft As Range, ByVal headers As Variant) As ListObject`  
  Ensures a table exists. Creates it if missing using the supplied header list and top-left anchor.

- `Public Sub Excel_UpsertListObjectOnSheet(ByVal ws As Worksheet, ByVal tableName As String, ByVal topLeft As Range, ByVal headers As Variant, ByVal data2D As Variant, Optional ByVal clearExisting As Boolean = True, Optional ByVal addMissingColumns As Boolean = True, Optional ByVal removeMissingColumns As Boolean = False, Optional ByVal preserveFormulaColumns As Boolean = True, Optional ByVal fillFormulasOnAppend As Boolean = True)`  
  Core Excel upsert entry point for writing headers and a 2D array into a `ListObject`, with schema evolution and formula-preservation options.

- `Public Sub Excel_ResizeTableToRowCol(ByVal lo As ListObject, ByVal finalHeaders As Variant, ByVal bodyRowCount As Long)`  
  Resizes a `ListObject` to the requested header/body shape while handling Excel table materialization edge cases.

- `Public Sub Excel_UpsertListObjectFromJsonAtRoot(ByVal ws As Worksheet, ByVal tableName As String, ByVal topLeft As Range, ByVal jsonText As String, ByVal tableRoot As String, Optional ByVal clearExisting As Boolean = True, Optional ByVal addMissingColumns As Boolean = True, Optional ByVal removeMissingColumns As Boolean = False, Optional ByVal preserveFormulaColumns As Boolean = True, Optional ByVal fillFormulasOnAppend As Boolean = True, Optional ByVal nonTableArraysAsJson As Boolean = False)`  
  High-level JSON-to-table ingestion entry point. Parses JSON, resolves a root array-of-objects, shapes rows and headers, then upserts into Excel.

- `Public Function Excel_ListObjectToJson(ByVal lo As ListObject, Optional ByVal includeBlanksAsNull As Boolean = False, Optional ByVal parseJsonInCells As Boolean = False, Optional ByVal parseArraysOnly As Boolean = False, Optional ByVal preserveFormulas As Boolean = False) As String`  
  Converts an Excel table into a JSON array-of-objects. Supports nested dot-path headers and optional parsing of JSON text stored inside cells.

- `Public Sub Excel_UpsertListObjectFromSource(ByVal ws As Worksheet, ByVal tableName As String, ByVal topLeft As Range, ByVal sourceText As String, ByVal format As ExcelSourceFormat, Optional ByVal tableRoot As String = "$", Optional ByVal clearExisting As Boolean = True, Optional ByVal addMissingColumns As Boolean = True, Optional ByVal removeMissingColumns As Boolean = False, Optional ByVal preserveFormulaColumns As Boolean = True, Optional ByVal fillFormulasOnAppend As Boolean = True, Optional ByVal nonTableArraysAsJson As Boolean = False)`  
  Unified ingestion entry point for JSON, CSV, or XML source text. Converts the source into JSON and routes through the deterministic Excel upsert pipeline.

### CSV / XML Adapters

- `Public Function CsvFileToJson(ByVal filePath As String) As String`  
  Reads a CSV file from disk and converts it into a JSON array-of-objects.

- `Public Function CsvTextToJson(ByVal txt As String) As String`  
  Converts raw CSV text into a JSON array-of-objects using the module’s built-in CSV parser.

- `Public Function XmlFileToJson(ByVal filePath As String) As String`  
  Reads an XML file from disk and converts it into JSON.

- `Public Function XmlTextToJson(ByVal txt As String) As String`  
  Converts raw XML text into JSON using the module’s lightweight pure-VBA XML parser.

### JSON Coalescing Utilities

- `Public Function Json_CoalesceChildArrays(ByVal parentJson As String, ByVal parentRoot As String, ByVal childProperty As String, Optional ByVal strictMode As Boolean = False, Optional ByVal parentKeyMap As Collection = Nothing) As String`  
  Pulls child arrays out of parent rows and merges them into one JSON array, optionally injecting parent fields, literals, or formulas.

- `Public Function Json_CoalesceArraysFromStrings(ByVal jsonStrings As Collection, Optional ByVal strictMode As Boolean = False) As String`  
  Merges multiple JSON array strings into a single JSON array string, with optional strict shape validation.

- `Public Function Json_CoalesceArraysFromRange(ByVal rng As Range, Optional ByVal strictMode As Boolean = False) As String`  
  Reads JSON array strings from an Excel range and merges them into one JSON array string.

### Excel Range Utility

- `Public Function Excel_RangeToJsonStrings(ByVal rng As Range) As Collection`  
  Reads non-empty cell values from a range and returns them as a collection of trimmed JSON strings.

---

## Support Open Source

XLIDE is open-source software. If it saves you time or helps your team keep VBA
workbooks maintainable, support helps keep the project moving.

- [GitHub Sponsors](https://github.com/sponsors/WilliamSmithEdward)
- [PayPal](https://www.paypal.com/donate/?business=ML855BRLNR838&no_recurring=0&item_name=VBA+has+always+treated+me+well.+It+was+how+I+first+grew+professional+as+a+programmer%2C+I%27m+happy+to+show+it+some+love+%E2%9D%A4%EF%B8%8F&currency_code=USD)
- [Cash App](https://cash.app/$williamesmithjcil)
