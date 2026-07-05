Attribute VB_Name = "Json_Excel"
Option Explicit

' =============================================================================
' Module:      Json_Excel
' Project:     ModernJsonInVBA
'
' Excel ListObject ingestion: deterministic loading of JSON/CSV/XML into
' tables. The reverse direction (tables/ranges back to JSON) lives in
' Json_Excel_Export.
'
'   Excel_UpsertListObjectFromSource     unified entry point (JSON/CSV/XML)
'   Excel_UpsertListObjectFromJsonAtRoot JSON -> table at a JSONPath root
'   Excel_UpsertListObjectOnSheet        headers + 2D data -> table
'   Excel_GetListObject / Excel_EnsureListObject / Excel_ResizeTableToRowCol
'
' Ingestion guarantees:
'   - Header discovery preserves first-seen order; column order is stable
'     across runs.
'   - Schema evolution is explicit: addMissingColumns / removeMissingColumns.
'   - Formula columns can be preserved across refreshes and filled on append.
'   - The table body is written with a single Range assignment (verified
'     stable to millions of cells); a block-write fallback engages only if
'     that write raises on a memory-constrained host.
'   - Application state (calculation, events, screen updating, status bar)
'     is saved and restored even when an error is raised.
'
' Error numbers (all vbObjectError + n):
'   1101 removeMissingColumns requires clearExisting
'   1102 target exceeds worksheet bounds
'   1120 blank header            1121 duplicate header
'   1130 JSON root not object/array
'   1140 ListObject has no header row
'   1160 tableRoot not found     1162 tableRoot not array-of-objects
'   1163 array element is not an object
'   1400 unsupported source format
' =============================================================================

' Block size for the degraded write path only. The normal path writes the
' whole body in one Range assignment; see Excel_WriteBody.
Private Const EXCEL_FALLBACK_BLOCK_CELLS As Long = 1000000

' Source formats accepted by Excel_UpsertListObjectFromSource.
Public Enum ExcelSourceFormat
    ExcelSourceFormat_JSON = 1
    ExcelSourceFormat_CSV = 2
    ExcelSourceFormat_XML = 3
End Enum

' =============================================================================
' Public API: ListObject lookup and creation
' =============================================================================

' Find a ListObject on ws by name (case-insensitive). Returns Nothing when
' absent. Only ws is searched.
Public Function Excel_GetListObject(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If StrComp(lo.name, tableName, vbTextCompare) = 0 Then
            Set Excel_GetListObject = lo
            Exit Function
        End If
    Next lo
    Set Excel_GetListObject = Nothing
End Function

' Return the named ListObject, creating it at topLeft with the given headers
' (validated for blanks/duplicates) when it does not exist. A new table is
' created with a header row only; the body is empty.
Public Function Excel_EnsureListObject( _
    ByVal ws As Worksheet, _
    ByVal tableName As String, _
    ByVal topLeft As Range, _
    ByVal headers As Variant _
) As ListObject

    Dim lo As ListObject
    Set lo = Excel_GetListObject(ws, tableName)

    If lo Is Nothing Then
        Excel_ValidateHeaders headers, "Excel_EnsureListObject"

        Dim colCount As Long
        colCount = UBound(headers) - LBound(headers) + 1

        Dim headerRange As Range
        Set headerRange = ws.Range(topLeft, topLeft.Offset(0, colCount - 1))

        headerRange.Value2 = Excel_HeadersTo2D(headers)

        Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=headerRange, XlListObjectHasHeaders:=xlYes)
        lo.name = tableName
    End If

    Set Excel_EnsureListObject = lo
End Function

' =============================================================================
' Public API: upsert from headers + 2D data
' =============================================================================

' Create-or-update the named table with the given headers and 2D data.
' See Excel_ListObjectUpsertData for the schema-evolution semantics.
Public Sub Excel_UpsertListObjectOnSheet( _
    ByVal ws As Worksheet, _
    ByVal tableName As String, _
    ByVal topLeft As Range, _
    ByVal headers As Variant, _
    ByVal data2D As Variant, _
    Optional ByVal clearExisting As Boolean = True, _
    Optional ByVal addMissingColumns As Boolean = True, _
    Optional ByVal removeMissingColumns As Boolean = False, _
    Optional ByVal preserveFormulaColumns As Boolean = True, _
    Optional ByVal fillFormulasOnAppend As Boolean = True _
)
    Dim lo As ListObject
    Set lo = Excel_GetListObject(ws, tableName)

    If lo Is Nothing Then
        Set lo = Excel_EnsureListObject(ws, tableName, topLeft, headers)
    End If

    Excel_ListObjectUpsertData lo, headers, data2D, _
        clearExisting, addMissingColumns, removeMissingColumns, _
        preserveFormulaColumns, fillFormulasOnAppend
End Sub

' Core upsert. Resolves the final schema from the incoming headers and the
' existing table:
'
'   removeMissingColumns=True  incoming schema wins verbatim
'   addMissingColumns=True     union: existing columns first, new appended
'   both False                 existing schema wins; incoming data reshaped
'
' clearExisting picks replace vs append. Data is written in chunks; formula
' column templates are captured up-front and reapplied afterwards.
Private Sub Excel_ListObjectUpsertData( _
    ByVal lo As ListObject, _
    ByVal headers As Variant, _
    ByVal data2D As Variant, _
    Optional ByVal clearExisting As Boolean = True, _
    Optional ByVal addMissingColumns As Boolean = True, _
    Optional ByVal removeMissingColumns As Boolean = False, _
    Optional ByVal preserveFormulaColumns As Boolean = True, _
    Optional ByVal fillFormulasOnAppend As Boolean = True _
)
    Const ERR_SRC As String = "Excel_ListObjectUpsertData"
    Const ERR_SHEET_BOUNDS As Long = vbObjectError + 1102

    If removeMissingColumns And (Not clearExisting) Then
        Err.Raise vbObjectError + 1101, ERR_SRC, _
            "removeMissingColumns=True requires clearExisting=True (schema shrink would corrupt existing rows)."
    End If

    ' Save and suspend application state for bulk writes.
    Dim calcOld As XlCalculation
    Dim eventsOld As Boolean
    Dim updatingOld As Boolean
    Dim statusBarOld As Variant

    calcOld = Application.Calculation
    eventsOld = Application.EnableEvents
    updatingOld = Application.ScreenUpdating
    statusBarOld = Application.StatusBar

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = False

    ' Capture formula templates before any structural change.
    Dim fHdrs() As String
    Dim fFmls() As String
    Dim fCount As Long
    If preserveFormulaColumns Then
        Excel_CaptureFormulaTemplates lo, fHdrs, fFmls, fCount
    Else
        fCount = 0
    End If

    Dim oldCols As Long
    oldCols = lo.ListColumns.count

    Dim oldBodyRows As Long
    Dim oldBody As Range
    If lo.DataBodyRange Is Nothing Then
        oldBodyRows = 0
    Else
        Set oldBody = lo.DataBodyRange
        oldBodyRows = oldBody.rows.count
    End If

    Dim existingHeaders As Variant
    existingHeaders = Excel_ListObjectHeadersTo1D(lo)

    ' ---- Resolve final schema ----
    Dim finalHeaders As Variant
    Dim finalData As Variant

    Dim incomingCount As Long
    incomingCount = UBound(headers) - LBound(headers) + 1

    Dim existingCount As Long
    existingCount = UBound(existingHeaders) - LBound(existingHeaders) + 1

    Dim sameSchema As Boolean
    sameSchema = False

    If incomingCount = existingCount Then
        sameSchema = True

        Dim hs As Long
        For hs = 1 To existingCount
            If StrComp( _
                Trim$(CStr(headers(LBound(headers) + hs - 1))), _
                Trim$(CStr(existingHeaders(LBound(existingHeaders) + hs - 1))), _
                vbTextCompare _
            ) <> 0 Then
                sameSchema = False
                Exit For
            End If
        Next hs
    End If

    If removeMissingColumns Then
        finalHeaders = headers
        finalData = data2D
    ElseIf addMissingColumns Then
        If sameSchema Then
            finalHeaders = existingHeaders
            finalData = data2D
        Else
            finalHeaders = Excel_UnionHeadersFromListObject(lo, headers)

            ' Reshape only when the union actually differs from the incoming
            ' layout; a straight match can use the data as-is.
            Dim finalCount1 As Long
            finalCount1 = UBound(finalHeaders) - LBound(finalHeaders) + 1

            If finalCount1 = incomingCount Then
                Dim unionMatchesIncoming As Boolean
                unionMatchesIncoming = True

                Dim hu As Long
                For hu = 1 To incomingCount
                    If StrComp( _
                        Trim$(CStr(headers(LBound(headers) + hu - 1))), _
                        Trim$(CStr(finalHeaders(LBound(finalHeaders) + hu - 1))), _
                        vbTextCompare _
                    ) <> 0 Then
                        unionMatchesIncoming = False
                        Exit For
                    End If
                Next hu

                If unionMatchesIncoming Then
                    finalData = data2D
                Else
                    finalData = Excel_ReshapeDataToHeaders(headers, finalHeaders, data2D)
                End If
            Else
                finalData = Excel_ReshapeDataToHeaders(headers, finalHeaders, data2D)
            End If
        End If
    Else
        finalHeaders = existingHeaders

        If sameSchema Then
            finalData = data2D
        Else
            finalData = Excel_ReshapeDataToHeaders(headers, finalHeaders, data2D)
        End If
    End If

    Excel_ValidateHeaders finalHeaders, ERR_SRC

    Dim newBodyRows As Long
    newBodyRows = Excel_RowCount2D(finalData)

    ' Zero-row schema-replace against the default ["value"] header keeps the
    ' existing schema instead of wiping it (an empty result should clear
    ' data, not destroy a real table layout).
    If removeMissingColumns Then
        If newBodyRows = 0 Then
            If Excel_IsDefaultValueOnlyHeaders(headers) Then
                finalHeaders = existingHeaders
                finalData = Empty
                newBodyRows = 0
                removeMissingColumns = False
                addMissingColumns = False
                clearExisting = True
            Else
                finalHeaders = headers
                finalData = Empty
                newBodyRows = 0
            End If
        End If
    End If

    Dim newCols As Long
    newCols = UBound(finalHeaders) - LBound(finalHeaders) + 1

    Dim targetBodyRows As Long
    If clearExisting Then
        targetBodyRows = newBodyRows
    Else
        targetBodyRows = oldBodyRows + newBodyRows
    End If

    Dim headerRow As Long
    headerRow = lo.HeaderRowRange.row

    If newCols > lo.parent.Columns.count Then
        Err.Raise ERR_SHEET_BOUNDS, ERR_SRC, _
            "Target column count exceeds worksheet limit. cols=" & newCols & ", max=" & lo.parent.Columns.count
    End If

    If headerRow + targetBodyRows > lo.parent.rows.count Then
        Err.Raise ERR_SHEET_BOUNDS, ERR_SRC, _
            "Target row count exceeds worksheet limit. required_last_row=" & (headerRow + targetBodyRows) & _
            ", max=" & lo.parent.rows.count
    End If

    If newCols < oldCols Then
        Excel_ClearOrphanedColumns lo, newCols, oldCols, oldBodyRows
    End If

    Dim writeHeaders As Boolean
    writeHeaders = clearExisting Or (newCols <> oldCols)

    Dim header2D As Variant
    If writeHeaders Then
        header2D = Excel_HeadersTo2D(finalHeaders)
    End If

    ' ---- Write ----
    If clearExisting Then
        If oldBodyRows > 0 Then oldBody.ClearContents

        Excel_ResizeTableToRowCol lo, finalHeaders, newBodyRows

        If writeHeaders Then
            lo.HeaderRowRange.Value2 = header2D
        End If

        If newBodyRows > 0 Then
            Excel_WriteBody lo.DataBodyRange, finalData, 0, newBodyRows, newCols
        End If

        If IsArray(finalData) Then Erase finalData

        If preserveFormulaColumns And fCount > 0 Then
            Excel_ApplyFormulasToBody lo, finalHeaders, newBodyRows, fHdrs, fFmls, fCount
        End If
    Else
        Dim startRow As Long
        startRow = oldBodyRows

        Excel_ResizeTableToRowCol lo, finalHeaders, targetBodyRows

        If writeHeaders Then
            lo.HeaderRowRange.Value2 = header2D
        End If

        If newBodyRows > 0 Then
            Excel_WriteBody lo.DataBodyRange, finalData, startRow, newBodyRows, newCols
        End If

        If IsArray(finalData) Then Erase finalData

        If preserveFormulaColumns And fillFormulasOnAppend And fCount > 0 Then
            Excel_ApplyFormulasToAppendedRows lo, finalHeaders, startRow, newBodyRows, fHdrs, fFmls, fCount
        End If
    End If

    If newCols < oldCols Then
        Excel_ClearOrphanedHeaderOnly lo, newCols, oldCols
    End If

CleanExit:
    Application.StatusBar = statusBarOld
    Application.Calculation = calcOld
    Application.EnableEvents = eventsOld
    Application.ScreenUpdating = updatingOld
    Exit Sub

CleanFail:
    Application.StatusBar = statusBarOld
    Application.Calculation = calcOld
    Application.EnableEvents = eventsOld
    Application.ScreenUpdating = updatingOld
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Write srcData into body starting after rowOffset.
'
' The whole payload is written with ONE Range assignment: modern Excel
' accepts multi-million-cell writes from in-process VBA without trouble
' (verified to 8M cells), and per-write overhead dwarfs everything else.
' If the single write raises anyway (out of memory on constrained hosts),
' the fallback retries in large blocks before giving up.
Private Sub Excel_WriteBody( _
    ByVal body As Range, _
    ByRef srcData As Variant, _
    ByVal rowOffset As Long, _
    ByVal rowCount As Long, _
    ByVal colCount As Long _
)
    On Error GoTo BlockFallback
    body.Cells(rowOffset + 1, 1).Resize(rowCount, colCount).Value2 = srcData
    Exit Sub

BlockFallback:
    Err.Clear
    On Error GoTo 0

    Dim rowsPerBlock As Long
    rowsPerBlock = EXCEL_FALLBACK_BLOCK_CELLS \ colCount
    If rowsPerBlock < 1 Then rowsPerBlock = 1

    ' The payload fit no better as one write than it will as one block.
    If rowsPerBlock >= rowCount Then rowsPerBlock = (rowCount + 1) \ 2
    If rowsPerBlock < 1 Then rowsPerBlock = 1

    Dim srcRowLb As Long
    Dim srcColLb As Long
    srcRowLb = LBound(srcData, 1)
    srcColLb = LBound(srcData, 2)

    Dim writeStart As Long
    writeStart = 1

    Do While writeStart <= rowCount
        Dim takeRows As Long
        takeRows = rowsPerBlock
        If writeStart + takeRows - 1 > rowCount Then
            takeRows = rowCount - writeStart + 1
        End If

        Dim blockData As Variant
        ReDim blockData(1 To takeRows, 1 To colCount)

        Dim rr As Long
        Dim cc As Long
        For rr = 1 To takeRows
            For cc = 1 To colCount
                blockData(rr, cc) = srcData(srcRowLb + writeStart + rr - 2, srcColLb + cc - 1)
            Next cc
        Next rr

        body.Cells(rowOffset + writeStart, 1).Resize(takeRows, colCount).Value2 = blockData
        Erase blockData

        writeStart = writeStart + takeRows
    Loop
End Sub

' Resize a ListObject to the requested header/body shape.
'
' Excel does not always materialize body rows after Resize; missing rows are
' added explicitly. For bodyRowCount = 0 a temporary body row is used during
' the resize and then deleted, which is the only reliable way to shrink a
' table to header-only.
Public Sub Excel_ResizeTableToRowCol( _
    ByVal lo As ListObject, _
    ByVal finalHeaders As Variant, _
    ByVal bodyRowCount As Long _
)
    If Not lo.ShowHeaders Then lo.ShowHeaders = True
    If lo.HeaderRowRange Is Nothing Then
        Err.Raise vbObjectError + 1140, "Excel_ResizeTableToRowCol", _
            "ListObject has no HeaderRowRange (headers hidden or table corrupted): " & lo.name
    End If

    Dim headerTopLeft As Range
    Set headerTopLeft = lo.HeaderRowRange.Cells(1, 1)

    Dim colCount As Long
    colCount = UBound(finalHeaders) - LBound(finalHeaders) + 1

    Dim resizeBodyRows As Long
    resizeBodyRows = bodyRowCount
    If resizeBodyRows < 1 Then resizeBodyRows = 1

    lo.Resize headerTopLeft.Resize(1 + resizeBodyRows, colCount)

    If bodyRowCount <= 0 Then
        If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
        lo.HeaderRowRange.Value2 = Excel_HeadersTo2D(finalHeaders)
        Exit Sub
    End If

    Dim haveRows As Long
    If lo.DataBodyRange Is Nothing Then
        haveRows = 0
    Else
        haveRows = lo.DataBodyRange.rows.count
    End If

    If haveRows < bodyRowCount Then
        Dim needRows As Long
        needRows = bodyRowCount - haveRows

        Dim i As Long
        For i = 1 To needRows
            lo.ListRows.Add
        Next i
    End If
End Sub

' =============================================================================
' Public API: upsert from JSON at a table root
' =============================================================================

' Parse jsonText, resolve tableRoot to an array-of-objects (or null), and
' upsert the rows into the named table. Rows are filled into a single 2D
' array directly from the parsed model (no intermediate flatten of the
' document) and written in one pipeline pass.
'
' nonTableArraysAsJson:
'   False => nested arrays inside rows are excluded (prevents explosion)
'   True  => nested arrays are stored as JSON text in their cell
Public Sub Excel_UpsertListObjectFromJsonAtRoot( _
    ByVal ws As Worksheet, _
    ByVal tableName As String, _
    ByVal topLeft As Range, _
    ByVal jsonText As String, _
    ByVal tableRoot As String, _
    Optional ByVal clearExisting As Boolean = True, _
    Optional ByVal addMissingColumns As Boolean = True, _
    Optional ByVal removeMissingColumns As Boolean = False, _
    Optional ByVal preserveFormulaColumns As Boolean = True, _
    Optional ByVal fillFormulasOnAppend As Boolean = True, _
    Optional ByVal nonTableArraysAsJson As Boolean = False _
)
    Const SRC As String = "Excel_UpsertListObjectFromJsonAtRoot"

    On Error GoTo Fail

    Dim parsed As Variant
    Json_ParseInto jsonText, parsed

    If (Not IsObject(parsed)) Or (TypeName(parsed) <> "Collection") Then
        Err.Raise vbObjectError + 1130, SRC, _
            "JSON root must be an object or array (Collection). Primitive root is not supported for table upsert."
    End If

    Dim resolved As Variant
    If Not Json_TryResolvePath(parsed, tableRoot, resolved) Then
        Err.Raise vbObjectError + 1160, SRC, "tableRoot not found: " & tableRoot
    End If

    If Not IsNull(resolved) Then
        If (Not IsObject(resolved)) _
            Or (TypeName(resolved) <> "Collection") _
            Or Json_IsObject(resolved) Then
            Err.Raise vbObjectError + 1162, SRC, _
                "tableRoot must resolve to an array-of-objects (or null): " & tableRoot
        End If
    End If

    Dim arr As Collection
    If IsNull(resolved) Then
        Set arr = New Collection
    Else
        Set arr = resolved
    End If

    Dim rowCount As Long
    rowCount = arr.count

    ' SINGLE pass over the rows: validate each element, then fill its cells
    ' into one 2D array while headers register on the fly (the array's
    ' column dimension grows as new paths appear; see Json_RowObjectFillRow).
    ' Nothing touches the sheet until the one upsert call at the end, so a
    ' validation failure still leaves the table untouched - the same
    ' guarantee the old validate-then-collect-then-fill triple sweep gave,
    ' at a third of the traversal cost. For Each throughout: indexed arr(i)
    ' access walks the Collection's linked list and is quadratic.
    Dim headerIdx As JsonStringIndex
    JsonIdx_Init headerIdx, 64

    Dim data As Variant
    Dim rowVar As Variant
    Dim rowObj As Collection
    Dim rowIndex As Long

    If rowCount > 0 Then
        ReDim data(1 To rowCount, 1 To 8)   ' columns grow on demand

        rowIndex = 0
        For Each rowVar In arr
            rowIndex = rowIndex + 1

            If (Not IsObject(rowVar)) _
                Or (TypeName(rowVar) <> "Collection") _
                Or (Not Json_IsObject(rowVar)) Then
                Err.Raise vbObjectError + 1163, SRC, _
                    "Array element at index " & (rowIndex - 1) & " is not an object for root: " & tableRoot
            End If

            Set rowObj = rowVar
            Json_RowObjectFillRow rowObj, vbNullString, nonTableArraysAsJson, headerIdx, data, rowIndex
        Next rowVar
    End If

    ' Resolve the final schema and trim the array's spare column capacity.
    Dim headersOut As Variant
    If headerIdx.count = 0 Then
        ReDim headersOut(1 To 1) As Variant
        headersOut(1) = "value"

        If rowCount > 0 Then
            ReDim data(1 To rowCount, 1 To 1)   ' all Empty cells
        End If
    Else
        ReDim headersOut(1 To headerIdx.count) As Variant

        Dim hc As Long
        For hc = 1 To headerIdx.count
            headersOut(hc) = headerIdx.keys(hc)
        Next hc

        If rowCount > 0 Then
            If UBound(data, 2) > headerIdx.count Then
                ReDim Preserve data(1 To rowCount, 1 To headerIdx.count)
            End If
        End If
    End If

    ' Empty result + removeMissingColumns: keep the existing schema and just
    ' clear the data instead of collapsing the table to ["value"].
    If removeMissingColumns And rowCount = 0 Then
        Dim loExisting As ListObject
        Set loExisting = Excel_GetListObject(ws, tableName)

        If Not loExisting Is Nothing Then
            headersOut = Excel_ListObjectHeadersTo1D(loExisting)
            addMissingColumns = False
            removeMissingColumns = False
            clearExisting = True
        End If
    End If

    If rowCount = 0 Then
        Dim emptyData As Variant
        emptyData = Empty

        Excel_UpsertListObjectOnSheet ws, tableName, topLeft, _
            headersOut, emptyData, _
            clearExisting, addMissingColumns, removeMissingColumns, _
            preserveFormulaColumns, fillFormulasOnAppend

        Exit Sub
    End If

    ' One pipeline pass: one schema resolution, one resize, one body write.
    Excel_UpsertListObjectOnSheet ws, tableName, topLeft, _
        headersOut, data, _
        clearExisting, addMissingColumns, removeMissingColumns, _
        preserveFormulaColumns, fillFormulasOnAppend

    Erase data
    Exit Sub

Fail:
    ' Re-raise from this source, keeping the inner source in the message so
    ' the failure layer stays diagnosable.
    Dim n As Long
    Dim d As String
    Dim s As String

    n = Err.Number
    d = Err.Description
    s = Err.Source

    Err.Clear
    If Len(s) > 0 And StrComp(s, SRC, vbBinaryCompare) <> 0 Then
        d = d & " | inner_source=" & s
    End If

    Err.Raise n, SRC, d
End Sub

' =============================================================================
' Public API: unified source ingestion (JSON / CSV / XML)
' =============================================================================

' Convert sourceText to JSON when needed, then delegate to
' Excel_UpsertListObjectFromJsonAtRoot:
'
'   ExcelSourceFormat_JSON  used as-is; caller supplies tableRoot
'   ExcelSourceFormat_CSV   CsvTextToJson; table root is always "$"
'   ExcelSourceFormat_XML   XmlTextToJson; caller supplies tableRoot
'                           (commonly "$.item")
'
' All determinism and schema-evolution behavior comes from the JSON
' pipeline; this function is routing only.
Public Sub Excel_UpsertListObjectFromSource( _
    ByVal ws As Worksheet, _
    ByVal tableName As String, _
    ByVal topLeft As Range, _
    ByVal sourceText As String, _
    ByVal format As ExcelSourceFormat, _
    Optional ByVal tableRoot As String = "$", _
    Optional ByVal clearExisting As Boolean = True, _
    Optional ByVal addMissingColumns As Boolean = True, _
    Optional ByVal removeMissingColumns As Boolean = False, _
    Optional ByVal preserveFormulaColumns As Boolean = True, _
    Optional ByVal fillFormulasOnAppend As Boolean = True, _
    Optional ByVal nonTableArraysAsJson As Boolean = False _
)
    Const ERR_SRC As String = "Excel_UpsertListObjectFromSource"

    Dim jsonText As String
    Dim resolvedRoot As String

    Select Case format
        Case ExcelSourceFormat_JSON
            jsonText = sourceText
            resolvedRoot = tableRoot

        Case ExcelSourceFormat_CSV
            jsonText = CsvTextToJson(sourceText)
            resolvedRoot = "$"

        Case ExcelSourceFormat_XML
            jsonText = XmlTextToJson(sourceText)
            resolvedRoot = tableRoot

        Case Else
            Err.Raise vbObjectError + 1400, ERR_SRC, "Unsupported source format."
    End Select

    Excel_UpsertListObjectFromJsonAtRoot _
        ws, tableName, topLeft, jsonText, resolvedRoot, _
        clearExisting, addMissingColumns, removeMissingColumns, _
        preserveFormulaColumns, fillFormulasOnAppend, nonTableArraysAsJson
End Sub

' =============================================================================
' Formula preservation
'
' A "formula column" is any column whose body contains at least one formula;
' the template is the first formula found top-down (R1C1, so it is position
' independent). After a clear+write the template is reapplied down the whole
' column; after an append it is applied to the new rows only. Incoming data
' for a formula column is overwritten by the formula.
' =============================================================================

Private Sub Excel_CaptureFormulaTemplates( _
    ByVal lo As ListObject, _
    ByRef outHdrs() As String, _
    ByRef outFmlR1C1() As String, _
    ByRef outCount As Long _
)
    outCount = 0
    Erase outHdrs
    Erase outFmlR1C1

    If lo Is Nothing Then Exit Sub
    If lo.ListColumns.count = 0 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim c As Long
    For c = 1 To lo.ListColumns.count
        Dim f As String
        If Excel_TryFindFirstFormulaR1C1(lo.DataBodyRange.Columns(c), f) Then
            outCount = outCount + 1
            If outCount = 1 Then
                ReDim outHdrs(1 To 8) As String
                ReDim outFmlR1C1(1 To 8) As String
            ElseIf outCount > UBound(outHdrs) Then
                ReDim Preserve outHdrs(1 To UBound(outHdrs) * 2) As String
                ReDim Preserve outFmlR1C1(1 To UBound(outFmlR1C1) * 2) As String
            End If

            outHdrs(outCount) = CStr(lo.ListColumns(c).name)
            outFmlR1C1(outCount) = f
        End If
    Next c
End Sub

' Find the first formula in a column without touching cells one by one:
' Range.HasFormula answers True/False for the whole column in one COM call,
' and a mixed column is scanned via one bulk Formula read.
Private Function Excel_TryFindFirstFormulaR1C1(ByVal colRng As Range, ByRef outFormulaR1C1 As String) As Boolean
    Excel_TryFindFirstFormulaR1C1 = False
    outFormulaR1C1 = vbNullString

    If colRng Is Nothing Then Exit Function

    Dim hf As Variant
    hf = colRng.HasFormula

    If VarType(hf) = vbBoolean Then
        If hf = False Then Exit Function

        ' Entire column is formulas: take the first cell's template.
        outFormulaR1C1 = colRng.Cells(1, 1).FormulaR1C1
        Excel_TryFindFirstFormulaR1C1 = (Len(outFormulaR1C1) > 0)
        Exit Function
    End If

    ' Mixed column (HasFormula = Null): bulk-read formulas and find the
    ' first entry that is one.
    Dim formulas As Variant
    formulas = colRng.Formula

    If Not IsArray(formulas) Then
        If VarType(formulas) = vbString Then
            If Left$(CStr(formulas), 1) = "=" Then
                outFormulaR1C1 = colRng.Cells(1, 1).FormulaR1C1
                Excel_TryFindFirstFormulaR1C1 = (Len(outFormulaR1C1) > 0)
            End If
        End If
        Exit Function
    End If

    Dim r As Long
    For r = LBound(formulas, 1) To UBound(formulas, 1)
        Dim f As Variant
        f = formulas(r, LBound(formulas, 2))

        If VarType(f) = vbString Then
            If Left$(CStr(f), 1) = "=" Then
                outFormulaR1C1 = colRng.Cells(r - LBound(formulas, 1) + 1, 1).FormulaR1C1
                Excel_TryFindFirstFormulaR1C1 = (Len(outFormulaR1C1) > 0)
                Exit Function
            End If
        End If
    Next r
End Function

Private Function Excel_TryGetFormulaForHeader( _
    ByRef fHdrs() As String, _
    ByRef fFmls() As String, _
    ByVal fCount As Long, _
    ByVal headerName As String, _
    ByRef outFormulaR1C1 As String _
) As Boolean
    Excel_TryGetFormulaForHeader = False
    outFormulaR1C1 = vbNullString

    If fCount <= 0 Then Exit Function

    Dim i As Long
    For i = 1 To fCount
        If StrComp(fHdrs(i), headerName, vbTextCompare) = 0 Then
            outFormulaR1C1 = fFmls(i)
            Excel_TryGetFormulaForHeader = (Len(outFormulaR1C1) > 0)
            Exit Function
        End If
    Next i
End Function

Private Sub Excel_ApplyFormulasToBody( _
    ByVal lo As ListObject, _
    ByRef finalHeaders As Variant, _
    ByVal bodyRowCount As Long, _
    ByRef fHdrs() As String, _
    ByRef fFmls() As String, _
    ByVal fCount As Long _
)
    If lo Is Nothing Then Exit Sub
    If bodyRowCount <= 0 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim newCols As Long
    newCols = (UBound(finalHeaders) - LBound(finalHeaders) + 1)

    Dim c As Long
    For c = 1 To newCols
        Dim h As String
        h = CStr(finalHeaders(LBound(finalHeaders) + c - 1))

        Dim f As String
        If Excel_TryGetFormulaForHeader(fHdrs, fFmls, fCount, h, f) Then
            lo.DataBodyRange.Columns(c).FormulaR1C1 = f
        End If
    Next c
End Sub

Private Sub Excel_ApplyFormulasToAppendedRows( _
    ByVal lo As ListObject, _
    ByRef finalHeaders As Variant, _
    ByVal startRowZeroBased As Long, _
    ByVal appendedRowCount As Long, _
    ByRef fHdrs() As String, _
    ByRef fFmls() As String, _
    ByVal fCount As Long _
)
    If lo Is Nothing Then Exit Sub
    If appendedRowCount <= 0 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim newCols As Long
    newCols = (UBound(finalHeaders) - LBound(finalHeaders) + 1)

    Dim c As Long
    For c = 1 To newCols
        Dim h As String
        h = CStr(finalHeaders(LBound(finalHeaders) + c - 1))

        Dim f As String
        If Excel_TryGetFormulaForHeader(fHdrs, fFmls, fCount, h, f) Then
            lo.DataBodyRange.Cells(startRowZeroBased + 1, c).Resize(appendedRowCount, 1).FormulaR1C1 = f
        End If
    Next c
End Sub

' =============================================================================
' Header helpers
' =============================================================================

' Trim all headers in place, then validate: no blanks (1120), no
' case-insensitive duplicates (1121).
Private Sub Excel_ValidateHeaders(ByRef headers As Variant, ByVal sourceName As String)
    Dim lb As Long
    lb = LBound(headers)

    Dim i As Long
    For i = lb To UBound(headers)
        Dim hi As String
        hi = Trim$(CStr(headers(i)))

        If Len(hi) = 0 Then
            Err.Raise vbObjectError + 1120, sourceName, "Header at index " & i & " is blank."
        End If

        headers(i) = hi
    Next i

    Dim dupIdx As JsonStringIndex
    JsonIdx_Init dupIdx, (UBound(headers) - lb + 1) * 2, True

    For i = lb To UBound(headers)
        Dim seenAt As Long
        seenAt = JsonIdx_Find(dupIdx, CStr(headers(i)))

        If seenAt > 0 Then
            Err.Raise vbObjectError + 1121, sourceName, _
                "Duplicate header (case-insensitive): '" & CStr(headers(i)) & _
                "' at indices " & (lb + seenAt - 1) & " and " & i & "."
        End If

        JsonIdx_Ensure dupIdx, CStr(headers(i))
    Next i
End Sub

Private Function Excel_ListObjectHeadersTo1D(ByVal lo As ListObject) As Variant
    Dim n As Long
    n = lo.ListColumns.count

    Dim arr As Variant
    ReDim arr(1 To n)

    Dim i As Long
    For i = 1 To n
        arr(i) = lo.ListColumns(i).name
    Next i

    Excel_ListObjectHeadersTo1D = arr
End Function

' Union of the table's headers and the incoming headers: existing columns
' first (trimmed), then new incoming columns in their own order.
' Case-insensitive de-duplication.
Private Function Excel_UnionHeadersFromListObject(ByVal lo As ListObject, ByVal incomingHeaders As Variant) As Variant
    Dim unionIdx As JsonStringIndex
    JsonIdx_Init unionIdx, 64, True

    Dim existing As Variant
    existing = Excel_ListObjectHeadersTo1D(lo)

    Dim i As Long
    For i = 1 To UBound(existing)
        JsonIdx_Ensure unionIdx, Trim$(CStr(existing(i)))
    Next i

    For i = LBound(incomingHeaders) To UBound(incomingHeaders)
        JsonIdx_Ensure unionIdx, Trim$(CStr(incomingHeaders(i)))
    Next i

    Dim arr As Variant
    ReDim arr(1 To unionIdx.count)

    For i = 1 To unionIdx.count
        arr(i) = unionIdx.keys(i)
    Next i

    Excel_UnionHeadersFromListObject = arr
End Function

' Rearrange inData columns (labeled by inHeaders) into the outHeaders layout.
' Matching is trimmed and case-insensitive; columns with no source stay
' Empty. Row copying is per matched column.
Private Function Excel_ReshapeDataToHeaders( _
    ByVal inHeaders As Variant, _
    ByVal outHeaders As Variant, _
    ByVal inData As Variant _
) As Variant
    If IsEmpty(inData) Then
        Excel_ReshapeDataToHeaders = Empty
        Exit Function
    End If

    Dim inRows As Long
    Dim inCols As Long
    Dim outCols As Long

    inRows = Excel_RowCount2D(inData)
    inCols = Excel_ColCount2D(inData)
    outCols = (UBound(outHeaders) - LBound(outHeaders) + 1)

    Dim outArr As Variant
    ReDim outArr(1 To inRows, 1 To outCols)

    ' Index the incoming headers once.
    Dim inIdx As JsonStringIndex
    JsonIdx_Init inIdx, inCols * 2, True

    Dim i As Long
    For i = LBound(inHeaders) To UBound(inHeaders)
        JsonIdx_Ensure inIdx, Trim$(CStr(inHeaders(i)))
    Next i

    Dim oc As Long
    For oc = 1 To outCols
        Dim foundIdx As Long
        foundIdx = JsonIdx_Find(inIdx, Trim$(CStr(outHeaders(LBound(outHeaders) + oc - 1))))

        If foundIdx > 0 And foundIdx <= inCols Then
            Dim srcCol As Long
            srcCol = LBound(inData, 2) + foundIdx - 1

            Dim r As Long
            For r = 1 To inRows
                outArr(r, oc) = inData(LBound(inData, 1) + r - 1, srcCol)
            Next r
        End If
    Next oc

    Excel_ReshapeDataToHeaders = outArr
End Function

Private Function Excel_HeadersTo2D(ByVal headers As Variant) As Variant
    Dim lb As Long
    Dim ub As Long
    lb = LBound(headers)
    ub = UBound(headers)

    Dim outArr As Variant
    ReDim outArr(1 To 1, 1 To (ub - lb + 1))

    Dim i As Long
    For i = lb To ub
        outArr(1, i - lb + 1) = CStr(headers(i))
    Next i

    Excel_HeadersTo2D = outArr
End Function

' True for the schema the pipeline generates when a result has zero rows
' and no discovered columns: exactly one header named "value".
Private Function Excel_IsDefaultValueOnlyHeaders(ByVal headers As Variant) As Boolean
    On Error GoTo Nope

    Dim lb As Long
    Dim ub As Long
    lb = LBound(headers)
    ub = UBound(headers)

    If (ub - lb + 1) <> 1 Then GoTo Nope

    Excel_IsDefaultValueOnlyHeaders = (LCase$(Trim$(CStr(headers(lb)))) = "value")
    Exit Function

Nope:
    Excel_IsDefaultValueOnlyHeaders = False
End Function

' =============================================================================
' Range hygiene and sizing
' =============================================================================

' When the schema shrinks, clear the cells that used to belong to the table
' (header + body) so stale values do not linger next to it.
Private Sub Excel_ClearOrphanedColumns( _
    ByVal lo As ListObject, _
    ByVal newColCount As Long, _
    ByVal oldColCount As Long, _
    ByVal oldBodyRows As Long _
)
    Dim tl As Range
    Set tl = lo.Range.Cells(1, 1)

    tl.Offset(0, newColCount).Resize(1, oldColCount - newColCount).ClearContents

    If oldBodyRows > 0 Then
        tl.Offset(1, newColCount).Resize(oldBodyRows, oldColCount - newColCount).ClearContents
    End If
End Sub

' Second pass after the resize: clear header cells orphaned by the final
' table shape (the body was already handled).
Private Sub Excel_ClearOrphanedHeaderOnly( _
    ByVal lo As ListObject, _
    ByVal newColCount As Long, _
    ByVal oldColCount As Long _
)
    If newColCount >= oldColCount Then Exit Sub

    Dim tl As Range
    Set tl = lo.Range.Cells(1, 1)

    tl.Offset(0, newColCount).Resize(1, oldColCount - newColCount).ClearContents
End Sub

Private Function Excel_RowCount2D(ByVal data2D As Variant) As Long
    If IsEmpty(data2D) Then
        Excel_RowCount2D = 0
    Else
        Excel_RowCount2D = (UBound(data2D, 1) - LBound(data2D, 1) + 1)
    End If
End Function

Private Function Excel_ColCount2D(ByVal data2D As Variant) As Long
    If IsEmpty(data2D) Then
        Excel_ColCount2D = 0
    Else
        Excel_ColCount2D = (UBound(data2D, 2) - LBound(data2D, 2) + 1)
    End If
End Function

