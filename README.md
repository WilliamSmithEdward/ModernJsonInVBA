# ModernJsonInVBA

Deterministic JSON → Excel Tables → JSON  
Pure VBA. No dependencies. No silent schema drift.

------------------------------------------------------------------------

## Contents

- [Introduction](#introduction)
- [What This Solves](#what-this-solves)
- [Core Capabilities](#core-capabilities)
- [Real-World Performance](#real-world-performance---vba-json--excel-upsert-benchmark)
- [Installation](#installation)
  - [Option 1 — Copy Into Your Workbook](#option-1--copy-into-your-workbook)
  - [Option 2 — Use the Provided Workbook](#option-2--use-the-provided-workbook)
- [Requirements](#requirements)
- [Basic API Example](#basic-api-example)
  - [Refresh Mode](#refresh-mode)
  - [Append Mode](#append-mode)
  - [Strict Schema Mode](#strict-schema-mode)
- [Excel to JSON (Reverse Materialization)](#excel-to-json-reverse-materialization)
- [Understanding tableRoot](#understanding-tableroot)
- [HTTP Helper (Windows)](#http-helper-windows)
- [Schema Control](#schema-control)
- [Deterministic Errors](#deterministic-errors)

------------------------------------------------------------------------

## Introduction

**ModernJsonInVBA** is a single-module JSON engine for structured Excel
workflows.

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

## Real-World Performance - VBA JSON → Excel Upsert Benchmark

| Stage   | Seconds  |
|---------|----------|
| HTTP    | 0.019531 |
| Parse   | 0.011719 |
| Write   | 0.000000 |
| Upsert  | 0.015625 |
| **Total** | **0.046875** |

**Payload:** 55,040 bytes  
**Rows:** 100  
**Columns:** 15  
**Throughput:** **2133.33 rows/sec**

## Installation

ModernJsonInVBA supports two usage models:

-   Copy the module into your own workbook (recommended)
-   Use the provided `.xlsm` file directly

### Option 1 — Copy Into Your Workbook

1.  Download `ModernJsonInVBA.vba`
2.  Open the file in a text editor
3.  Select all → Copy

Then:

4.  Open your Excel workbook
5.  Press `ALT + F11`
6.  Insert → Module
7.  Paste the code
8.  Save as `.xlsm`

Module name:

    zz_ModernJsonInVba

------------------------------------------------------------------------

### Option 2 — Use the Provided Workbook

1.  Download `ModernVBAJson_1.0.0.xlsm`
2.  Open the file
3.  Enable macros

You may copy the module into another workbook if needed.

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

``` vb
Public Sub Example_Api_Refresh()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

    Dim jsonText As String
    jsonText = HttpGetText("https://jsonplaceholder.typicode.com/users")

    Excel_UpsertListObjectFromJsonAtRoot _
        ws, "tUsers", ws.Range("A1"), _
        jsonText, "$", _
        True, True, False

End Sub
```

### Append Mode

``` vb
Excel_UpsertListObjectFromJsonAtRoot _
    ws, "tUsersLog", ws.Range("A1"), _
    jsonText, "$", _
    False, True, False
```

### Strict Schema Mode

``` vb
Excel_UpsertListObjectFromJsonAtRoot _
    ws, "tUsersStrict", ws.Range("A1"), _
    jsonText, "$", _
    True, False, True
```

------------------------------------------------------------------------

### HTTP Helper (Windows)

``` vb
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
