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
  - [API Example With Nested Objects](#api-example-with-nested-objects)
- [Accessing Json Elements (Directly in VBA)](#accessing-json-elements-directly-in-vba)
- [Excel_UpsertListObjectFromJsonAtRoot](#excel_upsertlistobjectfromjsonatroot)
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
**Columns:** 4  
**Throughput:** **7314.28 cells/sec**

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
``` vba
Public Sub Example_Api_Refresh()

    '--------------------------------------------------------------------------
    ' Example_Api_Refresh
    '
    ' Demonstrates a full JSON ? Excel relational materialization pipeline.
    '
    ' Steps:
    '   1. Fetch JSON from remote API
    '   2. Materialize primary table (products) into tProducts
    '   3. Iterate rows to extract nested "reviews" arrays
    '   4. Inject foreign key (parentId) into each review object
    '   5. Materialize child table (tReviews)
    '
    ' Notes:
    '   - Uses deterministic JSON parser + Excel upsert engine
    '   - Avoids schema drift and maintains stable column ordering
    '--------------------------------------------------------------------------

    On Error GoTo CleanFail

    ' Improve performance during bulk operations
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' Target worksheet
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("Quick Start")

    '--------------------------------------------------------------------------
    ' Step 1: Retrieve JSON payload from API
    '--------------------------------------------------------------------------
    Dim jsonText As String
    jsonText = HttpGetText("https://dummyjson.com/products")

    '--------------------------------------------------------------------------
    ' Step 2: Materialize primary table (products)
    '
    ' Root: $.products
    ' Destination: ListObject "tProducts"
    '--------------------------------------------------------------------------
    Excel_UpsertListObjectFromJsonAtRoot _
        ws, "tProducts", ws.Range("A1"), _
        jsonText, "$.products", _
        True, True, False, True, True, True

    ' Reference the newly populated table
    Dim lo As ListObject: Set lo = ws.ListObjects("tProducts")

    '--------------------------------------------------------------------------
    ' Step 3: Prepare reviews table (child table)
    '
    ' If it already exists, clear the body before repopulating
    '--------------------------------------------------------------------------
    Dim loReviews As ListObject

    On Error Resume Next
    Set loReviews = ws.ListObjects("tReviews")
    On Error GoTo 0

    If Not loReviews Is Nothing Then
        If Not loReviews.DataBodyRange Is Nothing Then
            loReviews.DataBodyRange.Delete
        End If
    End If

    '--------------------------------------------------------------------------
    ' Step 4: Iterate parent rows to extract nested review arrays
    '--------------------------------------------------------------------------
    Dim rw As ListRow
    Dim reviewsColl As Collection
    Dim reviewObj As Object

    For Each rw In lo.ListRows

        ' Parent product ID (used as foreign key)
        Dim parentId As Variant
        parentId = rw.Range.Columns(lo.ListColumns("id").Index).value

        ' Nested reviews JSON stored as string
        Dim reviewsJson As String
        reviewsJson = rw.Range.Columns(lo.ListColumns("reviews").Index).value

        ' Skip empty values
        If Len(reviewsJson) > 0 Then

            ' Parse JSON array into collection
            Set reviewsColl = Nothing
            Json_ParseInto reviewsJson, reviewsColl

            ' Ensure parsed array contains elements
            If Not reviewsColl Is Nothing Then
                If reviewsColl.count > 0 Then

                    '----------------------------------------------------------
                    ' Inject parentId into each review object (relational foreign key)
                    ' Inject ISO date conversion formula
                    '----------------------------------------------------------
                    For Each reviewObj In reviewsColl
                        Json_ObjSet reviewObj, "parentId", parentId
                        Json_ObjSet reviewObj, "date (Pacific)", _
                            "=LET(utc,--SUBSTITUTE(LEFT([@date],19),""T"","" "")," & _
                            "y,YEAR(utc)," & _
                            "dstStartUTC,DATE(y,3,14)-WEEKDAY(DATE(y,3,14)-1)+10/24," & _
                            "dstEndUTC,DATE(y,11,7)-WEEKDAY(DATE(y,11,7)-1)+9/24," & _
                            "utc+IF((utc>=dstStartUTC)*(utc<dstEndUTC),-7/24,-8/24))"
                    Next

                    ' Convert enriched collection back to JSON
                    reviewsJson = Json_Stringify(reviewsColl)

                    '----------------------------------------------------------
                    ' Step 5: Materialize child table (reviews)
                    '
                    ' Root: $
                    ' Destination: ListObject "tReviews"
                    '----------------------------------------------------------
                    Excel_UpsertListObjectFromJsonAtRoot _
                        ws, "tReviews", ws.Range("A35"), _
                        reviewsJson, "$", _
                        False, True, False, True, True, True

                End If
            End If

        End If

    Next

CleanExit:

    ' Restore Excel environment
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

CleanFail:

    ' Ensure environment is restored even if an error occurs
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

## Excel_UpsertListObjectFromJsonAtRoot

### Overview

`Excel_UpsertListObjectFromJsonAtRoot` parses JSON, extracts a specific **array-of-objects** using a JSON path (`tableRoot`), converts the objects into rows, and **upserts** the result into an Excel table (`ListObject`).

The function handles nested objects automatically, preserves schema deterministically, and optionally stores nested arrays as JSON text in cells.

---

### Parameters

| Parameter | Description |
|---|---|
`ws` | Worksheet where the table exists or will be created |
`tableName` | Name of the Excel table (`ListObject`) |
`topLeft` | Cell where the table should be created if it does not already exist |
`jsonText` | The JSON text to parse |
`tableRoot` | JSON path that resolves to the **array-of-objects** that should become table rows |

---

### Behavior Flags

| Flag | Description |
|---|---|
`clearExisting` | If `True`, existing rows are cleared before writing new data (refresh mode). If `False`, rows are appended. |
`addMissingColumns` | If `True`, new columns discovered in the JSON will be added to the table. |
`removeMissingColumns` | If `True`, columns not present in the JSON will be removed from the table. |
`preserveFormulaColumns` | If `True`, existing formula columns are preserved during updates. |
`fillFormulasOnAppend` | If `True`, formulas automatically fill newly appended rows. |
`nonTableArraysAsJson` | If `True`, nested arrays that are not part of `tableRoot` are stored in cells as JSON text. If `False`, those arrays are excluded from the table extraction. |

---

### Notes

- `tableRoot` must resolve to an **array of JSON objects** (or `null`).
- Nested objects are flattened into dot-notation columns (`customer.id`, `customer.name`, etc.).
- Nested arrays not part of the selected table root can optionally be preserved as JSON text inside cells.

For more advanced usage (including nested table extraction and round-trip workflows), see the **complex nesting example in the provided `.xlsm` workbook**.

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
