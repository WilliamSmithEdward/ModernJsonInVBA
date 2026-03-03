Option Explicit

' =============================================================================
' JSON -> Excel ListObject Integration Tests (with asserts)
'
' These tests validate:
'   - Table creation / reuse (ListObject exists, anchored at A1)
'   - Correct row count after refresh
'   - Header behavior (union vs force schema)
'   - Data written to expected cells (spot checks)
'
' Assumptions:
'   - Json_Parse / Json_Flatten / Json_ExtractTableRows / Json_TableTo2D exist
'   - Excel_UpsertListObjectOnSheet / Excel_UpsertListObjectFromJsonAtRoot exist
'   - Excel_GetListObject exists
'
' Notes:
'   - Do NOT ws.Cells.Clear in append tests (that destroys the table).
' =============================================================================

' =============================================================================
' Constants
' =============================================================================

Private Const TAG_OBJECT As String = "__OBJ__"

' =============================================================================
' Run all JSON->Excel integration tests
' - Stops on first failure (preferred while iterating)
' - Raises the underlying assert error with context
' =============================================================================
Public Sub RunAll_JsonExcelTests_StopOnFail()

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo Fail

    ' Delete all test sheets (backwards loop to avoid enumerator skip)
    Dim i As Long
    For i = wb.Worksheets.count To 1 Step -1
        If wb.Worksheets(i).Name <> "Modern Json in VBA" And wb.Worksheets(i).Name <> "Performance" And wb.Worksheets(i).Name <> "Quick Start" And wb.Worksheets(i).Name <> "ListObject to JSON" Then
            wb.Worksheets(i).Delete
        End If
    Next i

    ' Start Tests
    Test_JSON_ToTable_RootArray_WithAsserts
    Test_JSON_ToTable_NestedArray_WithAsserts
    Test_JSON_AddMissingColumns_TwoPass_WithAsserts
    Test_JSON_ForceSchemaReplace_TwoPass_WithAsserts
    Test_Upsert_FromJson_TableRoot_NotJsonRoot_WithAsserts
    Test_Append_RootArray_TwoPass_WithAsserts
    Test_API_Poke_151_ToTable_WithAsserts
    Test_ForceSchema_EmptySecondPass_PreservesSchema_WithAsserts
    Test_FlushToZero_ThenWrite_No91_WithAsserts
    Test_FlushToZero_ThenAppend_No91_WithAsserts
    Test_ShrinkSchema_ThenWrite_No91_WithAsserts
    Test_FirstWrite_MaterializesDataBodyRange_WithAsserts
    Test_WriteZeroRows_DoesNotTouchDataBodyRange_No91_WithAsserts
    Test_ShrinkSchema_ToOneCol_ZeroRows_ThenAppend_No91_WithAsserts
    Test_DuplicateHeaders_Throws_1121
    Test_Append_ZeroRows_SchemaExpansion_SameCall_WithAsserts
    Test_Append_NarrowIncoming_DoesNotShrinkSchema_WithAsserts
    Test_Append_WiderIncoming_GrowsSchema_WithAsserts
    Test_Append_NoSchemaGrow_IgnoresExtraFields_WithAsserts
    Test_Append_GrowSchema_PreservesOrder_WithAsserts
    Test_ExistingTable_NotAtTopLeft_DoesNotMove_WithAsserts
    Test_TableNameCollisionAcrossSheets_Throws
    Test_Header_LeadingTrailingSpaces_AreTrimmed_WithAsserts
    Test_UnionHeaders_DoesNotCreateTrimDuplicate
    Test_Append_TrimVariantHeader_MapsToExistingColumn
    Test_NestedObject_FlattensToDottedColumns_WithAsserts
    Test_NestedObject_MissingAcrossRows_SparseValues_WithAsserts
    Test_NestedObject_SchemaUnionAcrossPasses_WithAsserts
    Test_NestedObject_ForceSchema_RemovesNestedColumns_WithAsserts
    Test_NestedObject_DottedHeader_TrimDup_Throws_1121
    Test_NestedArrayProperty_IsIgnoredInRootTable_WithAsserts
    Test_NestedObject_EmptyObject_DoesNotCreateValueFallback_WithAsserts
    Test_TableToJson_FlatBasic_WithAsserts
    Test_TaggedObject_TagConstant_And_Untagged_Throws_1134
    Test_TableToJson_Nested_EscapedDots_And_BlanksPolicy
    Test_TableToJson_RejectsArrayIndexHeaders_Throws_905
    Test_TableToJson_NestedPaths_And_EscapedDotKey
    Test_DeterministicHeaders_SparseObjects
    Test_Refresh_PreservesFormulaColumns_WithAsserts
    Test_Append_PreservesFormulaAndFillsDown_WithAsserts
    Test_Refresh_PreservesFormulaColumn_AndFillsDown_WithAsserts
    Test_Append_FillsFormulaDown_ForNewRows_WithAsserts
    Test_Refresh_PreservesFormulaColumns_AndFillsDown_WithAsserts
    Test_Append_AutoFillFormulaColumns_ForNewRows_WithAsserts
    Test_Refresh_PreservesFormulaColumns_And_FillsDown_WithAsserts
    ' End Tests

    ' Cleanup (same safe backwards loop)
    For i = wb.Worksheets.count To 1 Step -1
        If wb.Worksheets(i).Name <> "Modern Json in VBA" And wb.Worksheets(i).Name <> "Performance" And wb.Worksheets(i).Name <> "Quick Start" And wb.Worksheets(i).Name <> "ListObject to JSON" Then
            wb.Worksheets(i).Delete
        End If
    Next i

    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts

    MsgBox "All JSON to ListObject tests passed.", vbInformation
    Exit Sub

Fail:
    Dim msg As String
    msg = "Test run failed." & vbCrLf & _
          "Err " & CStr(Err.Number) & " (" & Err.Source & "): " & Err.Description

    ' Always restore Excel state before surfacing error
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts

    Err.Clear
    Err.Raise vbObjectError + 614, "mJsonExcelTests", msg

End Sub

' =============================================================================
' ASSERTS
' =============================================================================
Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 610, "mJsonExcelTests", "ASSERT FAIL: " & message
End Sub

Private Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal message As String)
    If IsNull(expected) And IsNull(actual) Then Exit Sub
    If expected <> actual Then
        Err.Raise vbObjectError + 611, "mJsonExcelTests", _
            "ASSERT FAIL: " & message & " | expected=" & CStr(expected) & " actual=" & CStr(actual)
    End If
End Sub

Private Sub AssertNotNothing(ByVal obj As Object, ByVal message As String)
    If obj Is Nothing Then Err.Raise vbObjectError + 612, "mJsonExcelTests", "ASSERT FAIL: " & message
End Sub

Private Sub AssertHeaderEquals(ByVal lo As ListObject, ByVal expectedHeaders As Variant, ByVal message As String)
    Dim i As Long, n As Long
    n = lo.ListColumns.count
    AssertEquals UBound(expectedHeaders) - LBound(expectedHeaders) + 1, n, message & " (col count)"

    For i = 1 To n
        AssertEquals CStr(expectedHeaders(LBound(expectedHeaders) + i - 1)), lo.ListColumns(i).Name, _
            message & " (header #" & i & ")"
    Next i
End Sub

Private Sub AssertBodyCellEquals(ByVal lo As ListObject, ByVal row1Based As Long, ByVal colName As String, ByVal expected As Variant, ByVal message As String)
    Dim colIdx As Long
    colIdx = lo.ListColumns(colName).Index

    AssertTrue Not lo.DataBodyRange Is Nothing, message & " (no DataBodyRange)"
    AssertTrue row1Based >= 1 And row1Based <= lo.DataBodyRange.rows.count, message & " (row out of range)"

    Dim actual As Variant
    actual = lo.DataBodyRange.Cells(row1Based, colIdx).Value2

    AssertEquals expected, actual, message & " (" & colName & ", r=" & row1Based & ")"
End Sub

Private Sub AssertRowCount(ByVal lo As ListObject, ByVal expectedRows As Long, ByVal message As String)
    Dim actualRows As Long
    If lo.DataBodyRange Is Nothing Then
        actualRows = 0
    Else
        actualRows = lo.DataBodyRange.rows.count
    End If
    AssertEquals expectedRows, actualRows, message
End Sub

Private Sub AssertBodyCellHasFormula( _
    ByVal lo As ListObject, _
    ByVal row1Based As Long, _
    ByVal colName As String, _
    ByVal message As String _
)
    Dim colIdx As Long
    colIdx = lo.ListColumns(colName).Index

    AssertTrue Not lo.DataBodyRange Is Nothing, message & " (no DataBodyRange)"
    AssertTrue row1Based >= 1 And row1Based <= lo.DataBodyRange.rows.count, message & " (row out of range)"

    Dim c As Range
    Set c = lo.DataBodyRange.Cells(row1Based, colIdx)

    AssertTrue c.HasFormula, message & " (expected formula, got none)"
End Sub

Private Sub AssertBodyCellFormulaR1C1Equals( _
    ByVal lo As ListObject, _
    ByVal row1Based As Long, _
    ByVal colName As String, _
    ByVal expectedR1C1 As String, _
    ByVal message As String _
)
    Dim colIdx As Long
    colIdx = lo.ListColumns(colName).Index

    AssertTrue Not lo.DataBodyRange Is Nothing, message & " (no DataBodyRange)"
    AssertTrue row1Based >= 1 And row1Based <= lo.DataBodyRange.rows.count, message & " (row out of range)"

    Dim c As Range
    Set c = lo.DataBodyRange.Cells(row1Based, colIdx)

    AssertTrue c.HasFormula, message & " (expected formula, got none)"
    AssertEquals expectedR1C1, c.FormulaR1C1, message & " (FormulaR1C1 mismatch)"
End Sub

' =============================================================================
' Helper: detect if Obj_Get exists (compile-safe)
' =============================================================================
Private Function HasPublicObjGet() As Boolean
    ' We can't reflect function existence safely in VBA without references.
    ' So this is a pragmatic switch:
    '   - If your module has Obj_Get, set this to True.
    '   - If not, leave False and we use TaggedObject_FindValue.
    HasPublicObjGet = False
End Function

' =============================================================================
' Helper: find a key in the tagged-object internal representation
'
' Assumes object Collection layout:
'   obj(1) = TAG_OBJECT
'   obj(2..) = entries where each entry is either:
'       - 2-element array: [key, value]
'       - 2-item Collection: (1)=key, (2)=value
'
' =============================================================================
Private Function TaggedObject_FindValue(ByVal obj As Collection, ByVal key As String) As Variant
    Dim i As Long
    For i = 2 To obj.count
        Dim entry As Variant
        entry = obj(i)

        Dim k As String
        Dim v As Variant

        If IsArray(entry) Then
            Dim lb As Long
            lb = LBound(entry)
            If (UBound(entry) - lb + 1) >= 2 Then
                k = CStr(entry(lb))
                v = entry(lb + 1)
            End If

        ElseIf IsObject(entry) Then
            If TypeName(entry) = "Collection" Then
                If entry.count >= 2 Then
                    k = CStr(entry(1))
                    v = entry(2)
                End If
            End If
        End If

        If StrComp(k, key, vbTextCompare) = 0 Then
            TaggedObject_FindValue = v
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 642, "TaggedObject_FindValue", "Key not found: " & key
End Function

Public Function Obj_Get( _
    ByVal obj As Collection, _
    ByVal key As String _
) As Variant

    Const ERR_SRC As String = "Obj_Get"

    If obj Is Nothing Then
        Err.Raise vbObjectError + 1200, ERR_SRC, "Object is Nothing."
    End If

    If obj.count < 1 Or CStr(obj(1)) <> TAG_OBJECT Then
        Err.Raise vbObjectError + 1201, ERR_SRC, _
            "Collection is not a tagged object."
    End If

    Dim i As Long
    For i = 2 To obj.count

        Dim entry As Variant
        entry = obj(i)

        Dim k As String
        Dim v As Variant

        If IsArray(entry) Then
            If UBound(entry) - LBound(entry) + 1 >= 2 Then
                k = CStr(entry(LBound(entry)))
                v = entry(LBound(entry) + 1)
            End If

        ElseIf IsObject(entry) Then
            If TypeName(entry) = "Collection" Then
                If entry.count >= 2 Then
                    k = CStr(entry(1))
                    v = entry(2)
                End If
            End If
        End If

        If StrComp(k, key, vbTextCompare) = 0 Then
            Obj_Get = v
            Exit Function
        End If

    Next i

    Err.Raise vbObjectError + 1202, ERR_SRC, _
        "Key not found: " & key

End Function

Public Function Obj_TryGet( _
    ByVal obj As Collection, _
    ByVal key As String, _
    ByRef valueOut As Variant _
) As Boolean

    On Error GoTo Fail

    valueOut = Obj_Get(obj, key)
    Obj_TryGet = True
    Exit Function

Fail:
    Obj_TryGet = False
End Function

' =============================================================================
' SHEET + TABLE HELPERS
' =============================================================================
Private Function EnsureTestSheet(ByVal sheetName As String) As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim safeName As String
    safeName = Excel_SheetNameMakeValid(sheetName)

    ' If ANY sheet exists with this name (worksheet OR chart), return it if it's a worksheet,
    ' otherwise generate a unique worksheet name.
    Dim sh As Object
    On Error Resume Next
    Set sh = wb.Sheets(safeName)
    On Error GoTo 0

    If Not sh Is Nothing Then
        If TypeName(sh) = "Worksheet" Then
            Set EnsureTestSheet = sh
            Exit Function
        Else
            ' Name is taken by a Chart (or other sheet type). Must pick a new worksheet name.
            safeName = Excel_SheetNameMakeUnique(wb, safeName)
        End If
    End If

    Dim ws As Worksheet
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    ws.Name = safeName

    Set EnsureTestSheet = ws
End Function

Private Function Excel_SheetNameMakeValid(ByVal proposed As String) As String
    Dim s As String
    s = Trim$(proposed)

    If Len(s) = 0 Then s = "Sheet"

    ' Replace invalid characters: : \ / ? * [ ]
    s = Replace$(s, ":", "_")
    s = Replace$(s, "\", "_")
    s = Replace$(s, "/", "_")
    s = Replace$(s, "?", "_")
    s = Replace$(s, "*", "_")
    s = Replace$(s, "[", "_")
    s = Replace$(s, "]", "_")

    ' Excel also dislikes leading/trailing apostrophes in some contexts
    Do While Left$(s, 1) = "'"
        s = Mid$(s, 2)
    Loop
    Do While Right$(s, 1) = "'"
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) = 0 Then s = "Sheet"

    ' Enforce 31 char max
    If Len(s) > 31 Then s = Left$(s, 31)

    Excel_SheetNameMakeValid = s
End Function

Private Function Excel_SheetNameMakeUnique(ByVal wb As Workbook, ByVal baseName As String) As String
    Dim base As String
    base = Excel_SheetNameMakeValid(baseName)

    ' If available, use it
    On Error Resume Next
    Dim tmp As Object
    Set tmp = wb.Sheets(base)
    On Error GoTo 0
    If tmp Is Nothing Then
        Excel_SheetNameMakeUnique = base
        Exit Function
    End If

    Dim i As Long
    For i = 1 To 9999
        Dim suffix As String
        suffix = "_" & CStr(i)

        Dim candidate As String
        If Len(base) + Len(suffix) <= 31 Then
            candidate = base & suffix
        Else
            candidate = Left$(base, 31 - Len(suffix)) & suffix
        End If

        On Error Resume Next
        Set tmp = wb.Sheets(candidate)
        On Error GoTo 0

        If tmp Is Nothing Then
            Excel_SheetNameMakeUnique = candidate
            Exit Function
        End If
    Next i

    ' Should never happen
    Excel_SheetNameMakeUnique = Left$(base, 27) & "_9999"
End Function

Private Sub ResetSheetButKeepTableTestSafe(ByVal ws As Worksheet)
    ' Clears only values + formats; does not delete the sheet object.
    ' NOTE: This WILL destroy any existing ListObjects, because clearing the sheet
    ' destroys their range. Use ONLY in tests where you expect recreate/refresh.
    ws.Cells.Clear
End Sub

Private Function GetTable(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    Set GetTable = Excel_GetListObject(ws, tableName)
End Function


' =============================================================================
' INTERNAL: JSON -> headersOut + data2D pipeline (no Excel write)
' =============================================================================
Private Sub Build2DFromJsonRoot( _
    ByVal jsonText As String, _
    ByVal tableRoot As String, _
    ByRef headersOut As Variant, _
    ByRef data2D As Variant _
)
    Dim parsed As Variant
    Json_ParseInto jsonText, parsed

    Dim flat As Collection
    Set flat = Json_Flatten(parsed)

    Dim rows As Collection
    Set rows = Json_ExtractTableRows(flat, tableRoot)

    data2D = Json_TableTo2D(rows, headersOut)
End Sub


' =============================================================================
' TEST 1: Root JSON Array -> Excel ListObject (replace)
' =============================================================================
Public Sub Test_JSON_ToTable_RootArray_WithAsserts()
    ' Expected:
    '   - Sheet: zTest_JSON_RootArray
    '   - Table: tRootOrders at A1
    '   - Headers: id, name, active (order as first-seen)
    '   - Rows: 2
    '   - Values spot check:
    '       row1 id=1 name=Alpha active=True
    '       row2 id=2 name=Beta  active=False

    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_JSON_RootArray")
    ResetSheetButKeepTableTestSafe ws

    Dim jsonText As String
    jsonText = "[" & _
        "{""id"":1,""name"":""Alpha"",""active"":true}," & _
        "{""id"":2,""name"":""Beta"",""active"":false}" & _
    "]"

    Dim headersOut As Variant, data2D As Variant
    Build2DFromJsonRoot jsonText, "$", headersOut, data2D

    Excel_UpsertListObjectOnSheet ws, "tRootOrders", ws.Range("A1"), headersOut, data2D, True, True, False

    Dim lo As ListObject
    Set lo = GetTable(ws, "tRootOrders")
    AssertNotNothing lo, "tRootOrders should exist"

    AssertRowCount lo, 2, "tRootOrders row count"
    AssertHeaderEquals lo, Array("id", "name", "active"), "tRootOrders headers"

    AssertBodyCellEquals lo, 1, "id", 1, "tRootOrders row1"
    AssertBodyCellEquals lo, 1, "name", "Alpha", "tRootOrders row1"
    AssertBodyCellEquals lo, 1, "active", True, "tRootOrders row1"

    AssertBodyCellEquals lo, 2, "id", 2, "tRootOrders row2"
    AssertBodyCellEquals lo, 2, "name", "Beta", "tRootOrders row2"
    AssertBodyCellEquals lo, 2, "active", False, "tRootOrders row2"
End Sub


' =============================================================================
' TEST 2: Nested Array-of-Objects -> Excel ListObject (replace)
' =============================================================================
Public Sub Test_JSON_ToTable_NestedArray_WithAsserts()
    ' JSON shape:
    '   { "orders": [ {...}, {...} ] }
    ' tableRoot:
    '   "$.orders"

    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_JSON_NestedOrders")
    ResetSheetButKeepTableTestSafe ws

    Dim jsonText As String
    jsonText = "{" & _
        """orders"":[" & _
            "{""id"":100,""customer"":""Alice"",""total"":25.5}," & _
            "{""id"":200,""customer"":""Bob"",""total"":99.9}" & _
        "]" & _
    "}"

    Dim headersOut As Variant, data2D As Variant
    Build2DFromJsonRoot jsonText, "$.orders", headersOut, data2D

    Excel_UpsertListObjectOnSheet ws, "tNestedOrders", ws.Range("A1"), headersOut, data2D, True, True, False

    Dim lo As ListObject
    Set lo = GetTable(ws, "tNestedOrders")
    AssertNotNothing lo, "tNestedOrders should exist"

    AssertRowCount lo, 2, "tNestedOrders row count"
    AssertHeaderEquals lo, Array("id", "customer", "total"), "tNestedOrders headers"

    AssertBodyCellEquals lo, 1, "id", 100, "tNestedOrders row1"
    AssertBodyCellEquals lo, 1, "customer", "Alice", "tNestedOrders row1"
    AssertBodyCellEquals lo, 1, "total", 25.5, "tNestedOrders row1"

    AssertBodyCellEquals lo, 2, "id", 200, "tNestedOrders row2"
    AssertBodyCellEquals lo, 2, "customer", "Bob", "tNestedOrders row2"
    AssertBodyCellEquals lo, 2, "total", 99.9, "tNestedOrders row2"
End Sub


' =============================================================================
' TEST 3: Schema Union (addMissingColumns=True) across two passes
' =============================================================================
Public Sub Test_JSON_AddMissingColumns_TwoPass_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_JSON_AddCols")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' PASS A
    Dim jsonA As String
    jsonA = "[" & _
        "{""id"":1,""name"":""Alpha"",""active"":true}," & _
        "{""id"":2,""name"":""Beta"",""active"":false}" & _
    "]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA

    Excel_UpsertListObjectOnSheet ws, "tAddCols", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tAddCols")
    AssertNotNothing lo, "tAddCols should exist after PASS A"
    AssertRowCount lo, 2, "tAddCols PASS A row count"
    AssertHeaderEquals lo, Array("id", "name", "active"), "tAddCols PASS A headers"

    ' PASS B (expands schema)
    Dim jsonB As String
    jsonB = "[" & _
        "{""id"":10,""name"":""Gamma"",""active"":true,""created_at"":""2026-02-28T12:00:00Z""}," & _
        "{""id"":20,""name"":""Delta"",""active"":false,""priority"":3}" & _
    "]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonB, "$", headersB, dataB

    ' union: addMissingColumns=True keeps existing order, appends new fields
    Excel_UpsertListObjectOnSheet ws, "tAddCols", ws.Range("A1"), headersB, dataB, True, True, False

    Set lo = GetTable(ws, "tAddCols")
    AssertRowCount lo, 2, "tAddCols PASS B row count"
    AssertHeaderEquals lo, Array("id", "name", "active", "created_at", "priority"), "tAddCols PASS B headers"

    AssertBodyCellEquals lo, 1, "id", 10, "tAddCols PASS B row1"
    AssertBodyCellEquals lo, 1, "name", "Gamma", "tAddCols PASS B row1"
    AssertBodyCellEquals lo, 1, "active", True, "tAddCols PASS B row1"
    AssertBodyCellEquals lo, 1, "created_at", "2026-02-28T12:00:00Z", "tAddCols PASS B row1"
    AssertBodyCellEquals lo, 1, "priority", Empty, "tAddCols PASS B row1 priority missing => Empty"

    AssertBodyCellEquals lo, 2, "id", 20, "tAddCols PASS B row2"
    AssertBodyCellEquals lo, 2, "created_at", Empty, "tAddCols PASS B row2 created_at missing => Empty"
    AssertBodyCellEquals lo, 2, "priority", 3, "tAddCols PASS B row2"
End Sub


' =============================================================================
' TEST 4: Force schema replace (removeMissingColumns=True) across two passes
' =============================================================================
Public Sub Test_JSON_ForceSchemaReplace_TwoPass_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_JSON_ForceSchema")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' PASS A
    Dim jsonA As String
    jsonA = "[" & _
        "{""id"":1,""name"":""Alpha"",""active"":true}," & _
        "{""id"":2,""name"":""Beta"",""active"":false}" & _
    "]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA

    Excel_UpsertListObjectOnSheet ws, "tForceSchema", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tForceSchema")
    AssertNotNothing lo, "tForceSchema should exist after PASS A"
    AssertHeaderEquals lo, Array("id", "name", "active"), "tForceSchema PASS A headers"
    AssertRowCount lo, 2, "tForceSchema PASS A rows"

    ' PASS B: schema becomes EXACTLY sku, price
    Dim jsonB As String
    jsonB = "[" & _
        "{""sku"":""ABC"",""price"":9.99}," & _
        "{""sku"":""XYZ"",""price"":19.99}" & _
    "]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonB, "$", headersB, dataB

    ' Force contract:
    '   clearExisting := True
    '   addMissingColumns := False
    '   removeMissingColumns := True
    Excel_UpsertListObjectOnSheet ws, "tForceSchema", ws.Range("A1"), headersB, dataB, True, False, True

    Set lo = GetTable(ws, "tForceSchema")
    AssertHeaderEquals lo, Array("sku", "price"), "tForceSchema PASS B headers forced"
    AssertRowCount lo, 2, "tForceSchema PASS B rows"

    AssertBodyCellEquals lo, 1, "sku", "ABC", "tForceSchema PASS B row1"
    AssertBodyCellEquals lo, 1, "price", 9.99, "tForceSchema PASS B row1"
    AssertBodyCellEquals lo, 2, "sku", "XYZ", "tForceSchema PASS B row2"
    AssertBodyCellEquals lo, 2, "price", 19.99, "tForceSchema PASS B row2"
End Sub


' =============================================================================
' TEST 5: Nested tableRoot not equal to JSON root (uses Excel_UpsertListObjectFromJsonAtRoot)
' =============================================================================
Public Sub Test_Upsert_FromJson_TableRoot_NotJsonRoot_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_JsonRootNotTableRoot")
    ResetSheetButKeepTableTestSafe ws

    Dim jsonText As String
    jsonText = "{" & _
        """meta"":{""request_id"":""abc-123"",""generated_at"":""2026-02-28T12:00:00Z""}," & _
        """data"":{" & _
            """customers"":[" & _
                "{""id"":101,""name"":""Alice"",""active"":true,""tier"":""gold""}," & _
                "{""id"":102,""name"":""Bob"",""active"":false}" & _
            "]" & _
        "}" & _
    "}"

    Excel_UpsertListObjectFromJsonAtRoot _
        ws, _
        "tCustomers", _
        ws.Range("A1"), _
        jsonText, _
        "$.data.customers", _
        True, True, False

    Dim lo As ListObject
    Set lo = GetTable(ws, "tCustomers")
    AssertNotNothing lo, "tCustomers should exist"

    AssertRowCount lo, 2, "tCustomers rows"
    AssertHeaderEquals lo, Array("id", "name", "active", "tier"), "tCustomers headers"

    AssertBodyCellEquals lo, 1, "id", 101, "tCustomers row1"
    AssertBodyCellEquals lo, 1, "name", "Alice", "tCustomers row1"
    AssertBodyCellEquals lo, 1, "active", True, "tCustomers row1"
    AssertBodyCellEquals lo, 1, "tier", "gold", "tCustomers row1"

    AssertBodyCellEquals lo, 2, "id", 102, "tCustomers row2"
    AssertBodyCellEquals lo, 2, "name", "Bob", "tCustomers row2"
    AssertBodyCellEquals lo, 2, "active", False, "tCustomers row2"
    AssertBodyCellEquals lo, 2, "tier", Empty, "tCustomers row2 missing tier => Empty"
End Sub


' =============================================================================
' TEST 6: Append behavior (clearExisting=False) on existing table
' IMPORTANT: Do NOT clear the sheet between passes.
' =============================================================================
Public Sub Test_Append_RootArray_TwoPass_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_JSON_Append")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' PASS A create
    Dim jsonA As String
    jsonA = "[" & _
        "{""id"":1,""name"":""Alpha""}," & _
        "{""id"":2,""name"":""Beta""}" & _
    "]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA
    Excel_UpsertListObjectOnSheet ws, "tAppend", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tAppend")
    AssertRowCount lo, 2, "tAppend PASS A rows"
    AssertHeaderEquals lo, Array("id", "name"), "tAppend PASS A headers"

    ' PASS B append
    Dim jsonB As String
    jsonB = "[" & _
        "{""id"":3,""name"":""Gamma""}," & _
        "{""id"":4,""name"":""Delta""}" & _
    "]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonB, "$", headersB, dataB

    ' clearExisting := False => append
    Excel_UpsertListObjectOnSheet ws, "tAppend", ws.Range("A1"), headersB, dataB, False, True, False

    Set lo = GetTable(ws, "tAppend")
    AssertRowCount lo, 4, "tAppend PASS B appended rows"

    AssertBodyCellEquals lo, 3, "id", 3, "tAppend row3"
    AssertBodyCellEquals lo, 3, "name", "Gamma", "tAppend row3"
    AssertBodyCellEquals lo, 4, "id", 4, "tAppend row4"
    AssertBodyCellEquals lo, 4, "name", "Delta", "tAppend row4"
End Sub


' =============================================================================
' HTTP GET (late-bound, no references)
' =============================================================================
Private Function HttpGet(ByVal url As String) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    http.Open "GET", url, False
    http.setRequestHeader "Accept", "application/json"
    http.send

    If http.Status = 200 Then
        HttpGet = http.responseText
    Else
        Err.Raise vbObjectError + 4000, "HttpGet", _
            "HTTP Error " & http.Status & " - " & http.statusText
    End If
End Function


' =============================================================================
' TEST 7: PokeAPI -> $.results table (integration smoke test)
' =============================================================================
Public Sub Test_API_Poke_151_ToTable_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_API_Poke_151")
    ResetSheetButKeepTableTestSafe ws

    Dim url As String
    url = "https://pokeapi.co/api/v2/pokemon?limit=151&offset=0"

    Dim jsonText As String
    jsonText = HttpGet(url)

    Excel_UpsertListObjectFromJsonAtRoot _
        ws, _
        "tPoke151", _
        ws.Range("A1"), _
        jsonText, _
        "$.results", _
        True, True, False

    Dim lo As ListObject
    Set lo = GetTable(ws, "tPoke151")
    AssertNotNothing lo, "tPoke151 should exist"

    ' Robust assertions (API might add fields; don't demand exact schema)
    AssertRowCount lo, 151, "tPoke151 row count"
    AssertTrue lo.ListColumns.count >= 2, "tPoke151 should have >=2 columns"
    AssertEquals True, (lo.ListColumns(1).Name <> vbNullString), "tPoke151 header 1 nonblank"
End Sub


' =============================================================================
' TEST 8: Force schema + empty second pass
'   Multi-phase:
'       PASS A: build schema from data
'       PASS B: JSON returns empty array
'
' Expected:
'   - Body flushed to 0 rows
'   - Schema preserved (NOT collapsed to "value")
' =============================================================================
Public Sub Test_ForceSchema_EmptySecondPass_PreservesSchema_WithAsserts()

    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_JSON_ForceSchema_Empty")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' -----------------------------
    ' PASS A: seed schema
    ' -----------------------------
    Dim jsonA As String
    jsonA = "[" & _
        "{""id"":1,""name"":""Alpha""}," & _
        "{""id"":2,""name"":""Beta""}" & _
    "]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA

    Excel_UpsertListObjectOnSheet _
        ws, "tForceEmpty", ws.Range("A1"), _
        headersA, dataA, _
        True, False, True   ' clearExisting, addMissing=False, removeMissing=True

    Set lo = GetTable(ws, "tForceEmpty")
    AssertNotNothing lo, "tForceEmpty should exist after PASS A"
    AssertHeaderEquals lo, Array("id", "name"), "PASS A headers"
    AssertRowCount lo, 2, "PASS A row count"

    ' -----------------------------
    ' PASS B: empty JSON
    ' -----------------------------
    Dim jsonEmpty As String
    jsonEmpty = "[]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonEmpty, "$", headersB, dataB

    ' Still forcing schema contract
    Excel_UpsertListObjectOnSheet _
        ws, "tForceEmpty", ws.Range("A1"), _
        headersB, dataB, _
        True, False, True   ' removeMissingColumns=True

    Set lo = GetTable(ws, "tForceEmpty")
    AssertNotNothing lo, "tForceEmpty should still exist after PASS B"

    ' -----------------------------
    ' Assertions
    ' -----------------------------

    ' Body must be flushed
    AssertRowCount lo, 0, "PASS B row count should be 0"

    ' Schema must be preserved (NOT replaced with ["value"])
    AssertHeaderEquals lo, Array("id", "name"), "PASS B headers preserved"

End Sub


' =============================================================================
' TEST 9: Flush body to 0 rows, then write rows again (regression for error 91)
' =============================================================================
Public Sub Test_FlushToZero_ThenWrite_No91_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_ZeroRows_ThenWrite")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' PASS A: seed schema and rows
    Dim jsonA As String
    jsonA = "[{""id"":1,""name"":""A""},{""id"":2,""name"":""B""}]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA

    Excel_UpsertListObjectOnSheet ws, "tZeroWrite", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tZeroWrite")
    AssertNotNothing lo, "tZeroWrite should exist after PASS A"
    AssertHeaderEquals lo, Array("id", "name"), "PASS A headers"
    AssertRowCount lo, 2, "PASS A rows"

    ' PASS B: flush to 0 rows but preserve schema (your guard path)
    Dim jsonEmpty As String
    jsonEmpty = "[]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonEmpty, "$", headersB, dataB

    Excel_UpsertListObjectOnSheet ws, "tZeroWrite", ws.Range("A1"), headersB, dataB, True, False, True

    Set lo = GetTable(ws, "tZeroWrite")
    AssertRowCount lo, 0, "PASS B rows should be 0"
    AssertHeaderEquals lo, Array("id", "name"), "PASS B headers preserved"

    ' PASS C: write 1 row again (this used to be a common DataBodyRange=Nothing scenario)
    Dim jsonC As String
    jsonC = "[{""id"":100,""name"":""Zed""}]"

    Dim headersC As Variant, dataC As Variant
    Build2DFromJsonRoot jsonC, "$", headersC, dataC

    Excel_UpsertListObjectOnSheet ws, "tZeroWrite", ws.Range("A1"), headersC, dataC, True, True, False

    Set lo = GetTable(ws, "tZeroWrite")
    AssertRowCount lo, 1, "PASS C rows should be 1"
    AssertTrue Not lo.DataBodyRange Is Nothing, "PASS C DataBodyRange must exist"
    AssertBodyCellEquals lo, 1, "id", 100, "PASS C row1"
    AssertBodyCellEquals lo, 1, "name", "Zed", "PASS C row1"
End Sub


' =============================================================================
' TEST 10: Flush to 0 rows, then append (clearExisting=False)
' =============================================================================
Public Sub Test_FlushToZero_ThenAppend_No91_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_ZeroRows_ThenAppend")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' PASS A: seed
    Dim jsonA As String
    jsonA = "[{""id"":1,""name"":""A""}]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA
    Excel_UpsertListObjectOnSheet ws, "tZeroAppend", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tZeroAppend")
    AssertRowCount lo, 1, "PASS A rows"

    ' PASS B: flush to 0, preserve schema
    Dim jsonEmpty As String
    jsonEmpty = "[]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonEmpty, "$", headersB, dataB
    Excel_UpsertListObjectOnSheet ws, "tZeroAppend", ws.Range("A1"), headersB, dataB, True, False, True

    Set lo = GetTable(ws, "tZeroAppend")
    AssertRowCount lo, 0, "PASS B rows should be 0"
    AssertHeaderEquals lo, Array("id", "name"), "PASS B headers preserved"

    ' PASS C: append 2 rows (clearExisting=False)
    Dim jsonC As String
    jsonC = "[{""id"":10,""name"":""X""},{""id"":20,""name"":""Y""}]"

    Dim headersC As Variant, dataC As Variant
    Build2DFromJsonRoot jsonC, "$", headersC, dataC

    Excel_UpsertListObjectOnSheet ws, "tZeroAppend", ws.Range("A1"), headersC, dataC, False, True, False

    Set lo = GetTable(ws, "tZeroAppend")
    AssertRowCount lo, 2, "PASS C rows should be 2"
    AssertTrue Not lo.DataBodyRange Is Nothing, "PASS C DataBodyRange must exist"

    AssertBodyCellEquals lo, 1, "id", 10, "PASS C row1"
    AssertBodyCellEquals lo, 2, "id", 20, "PASS C row2"
End Sub


' =============================================================================
' TEST 11: Force shrink schema to single column, then write rows (regression)
' =============================================================================
Public Sub Test_ShrinkSchema_ThenWrite_No91_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_Shrink_ThenWrite")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' PASS A: wide schema
    Dim jsonA As String
    jsonA = "[{""id"":1,""a"":10,""b"":20},{""id"":2,""a"":11,""b"":21}]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA
    Excel_UpsertListObjectOnSheet ws, "tShrink", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tShrink")
    AssertHeaderEquals lo, Array("id", "a", "b"), "PASS A headers"
    AssertRowCount lo, 2, "PASS A rows"

    ' PASS B: shrink to only id (removeMissingColumns=True, clearExisting=True)
    Dim jsonB As String
    jsonB = "[{""id"":100}]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonB, "$", headersB, dataB

    Excel_UpsertListObjectOnSheet ws, "tShrink", ws.Range("A1"), headersB, dataB, True, False, True

    Set lo = GetTable(ws, "tShrink")
    AssertHeaderEquals lo, Array("id"), "PASS B headers shrunk"
    AssertRowCount lo, 1, "PASS B rows"
    AssertBodyCellEquals lo, 1, "id", 100, "PASS B row1"
End Sub


' =============================================================================
' TEST 12: First write creates DataBodyRange (sanity guard)
' =============================================================================
Public Sub Test_FirstWrite_MaterializesDataBodyRange_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_FirstWrite_DataBodyRange")
    ResetSheetButKeepTableTestSafe ws

    Dim jsonText As String
    jsonText = "[{""id"":1,""name"":""Alpha""}]"

    Dim headersOut As Variant, data2D As Variant
    Build2DFromJsonRoot jsonText, "$", headersOut, data2D

    Excel_UpsertListObjectOnSheet ws, "tFirstWrite", ws.Range("A1"), headersOut, data2D, True, True, False

    Dim lo As ListObject
    Set lo = GetTable(ws, "tFirstWrite")
    AssertNotNothing lo, "tFirstWrite should exist"
    AssertRowCount lo, 1, "tFirstWrite rows"
    AssertTrue Not lo.DataBodyRange Is Nothing, "tFirstWrite DataBodyRange should exist after writing rows"
    AssertBodyCellEquals lo, 1, "id", 1, "tFirstWrite row1"
    AssertBodyCellEquals lo, 1, "name", "Alpha", "tFirstWrite row1"
End Sub


' =============================================================================
' TEST 13: Upsert should NOT touch DataBodyRange when writing 0 rows (no 91)
'   Scenario:
'     - Create table with rows
'     - ClearExisting=True with EMPTY data (0 rows)
'   Expected:
'     - 0 rows
'     - headers preserved (forced)
'     - no runtime error 91 during the call
' =============================================================================
Public Sub Test_WriteZeroRows_DoesNotTouchDataBodyRange_No91_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_WriteZeroRows_No91")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' PASS A: seed
    Dim jsonA As String
    jsonA = "[{""id"":1,""name"":""A""},{""id"":2,""name"":""B""}]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA
    Excel_UpsertListObjectOnSheet ws, "tZeroNo91", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tZeroNo91")
    AssertNotNothing lo, "tZeroNo91 exists after PASS A"
    AssertRowCount lo, 2, "PASS A rows"

    ' PASS B: write 0 rows, force schema (removeMissing=True)
    Dim jsonEmpty As String
    jsonEmpty = "[]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonEmpty, "$", headersB, dataB

    ' This call is the regression point: newBodyRows=0 and lo.DataBodyRange is Nothing
    Excel_UpsertListObjectOnSheet ws, "tZeroNo91", ws.Range("A1"), headersB, dataB, True, False, True

    Set lo = GetTable(ws, "tZeroNo91")
    AssertHeaderEquals lo, Array("id", "name"), "PASS B headers preserved"
    AssertRowCount lo, 0, "PASS B rows should be 0"
    ' Key: DataBodyRange may legitimately be Nothing at 0 rows; don't assert it exists here.
End Sub


' =============================================================================
' TEST 14: Shrink schema to 1 col WITH 0 rows, then append rows (no 91)
'   Scenario:
'     - Create wide table
'     - Force schema shrink to only [id] AND 0 rows
'     - Append rows (clearExisting=False)
'   Expected:
'     - After shrink pass: headers=[id], rows=0
'     - After append pass: rows=2, DataBodyRange exists, values correct
' =============================================================================
Public Sub Test_ShrinkSchema_ToOneCol_ZeroRows_ThenAppend_No91_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_ShrinkOneCol_ZeroThenAppend")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' PASS A: wide seed
    Dim jsonA As String
    jsonA = "[{""id"":1,""a"":10,""b"":20},{""id"":2,""a"":11,""b"":21}]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA
    Excel_UpsertListObjectOnSheet ws, "tShrinkZero", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tShrinkZero")
    AssertHeaderEquals lo, Array("id", "a", "b"), "PASS A headers"
    AssertRowCount lo, 2, "PASS A rows"

    ' PASS B: shrink schema to [id] but 0 rows
    Dim jsonB As String
    jsonB = "[]"   ' 0 rows is the important part here

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonB, "$", headersB, dataB

    ' Force schema shrink by providing the contract headers explicitly:
    '   We can't rely on headersB from empty JSON.
    '   So we pass headersOut ourselves.
    Dim contractHeaders As Variant
    contractHeaders = Array("id")

    ' Write 0 rows + removeMissingColumns=True => table becomes [id], 0 rows.
    Excel_UpsertListObjectOnSheet ws, "tShrinkZero", ws.Range("A1"), contractHeaders, Empty, True, False, True

    Set lo = GetTable(ws, "tShrinkZero")
    AssertHeaderEquals lo, Array("id"), "PASS B headers shrunk to id"
    AssertRowCount lo, 0, "PASS B rows should be 0"

    ' PASS C: append 2 rows into the shrunk 1-col table (this used to be 91-prone)
    Dim jsonC As String
    jsonC = "[{""id"":10},{""id"":20}]"

    Dim headersC As Variant, dataC As Variant
    Build2DFromJsonRoot jsonC, "$", headersC, dataC

    Excel_UpsertListObjectOnSheet ws, "tShrinkZero", ws.Range("A1"), headersC, dataC, False, True, False

    Set lo = GetTable(ws, "tShrinkZero")
    AssertHeaderEquals lo, Array("id"), "PASS C headers remain id"
    AssertRowCount lo, 2, "PASS C rows should be 2"
    AssertTrue Not lo.DataBodyRange Is Nothing, "PASS C DataBodyRange must exist"
    AssertBodyCellEquals lo, 1, "id", 10, "PASS C row1"
    AssertBodyCellEquals lo, 2, "id", 20, "PASS C row2"
End Sub


' =============================================================================
' TEST 15: Duplicate headers should throw 1121 (case-insensitive)
'   Must force schema = incoming headers (removeMissingColumns=True),
'   otherwise union logic will de-dupe and no error will occur.
' =============================================================================
Public Sub Test_DuplicateHeaders_Throws_1121()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_DuplicateHeaders_1121")
    ResetSheetButKeepTableTestSafe ws

    ' Seed a valid table first (covers existing-table path too)
    Dim jsonA As String
    jsonA = "[{""id"":1,""name"":""A""}]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA
    Excel_UpsertListObjectOnSheet ws, "tDupHdr", ws.Range("A1"), headersA, dataA, True, True, False

    ' Now attempt an illegal header contract (case-insensitive dup)
    Dim dupHeaders As Variant
    dupHeaders = Array("id", "ID") ' duplicate by vbTextCompare

    Dim threw As Boolean
    threw = False

    On Error Resume Next

    ' Force schema = dupHeaders => must throw 1121 at Excel_ValidateHeaders
    Excel_UpsertListObjectOnSheet ws, "tDupHdr", ws.Range("A1"), dupHeaders, dataA, True, False, True

    If Err.Number <> 0 Then
        threw = True
        AssertEquals vbObjectError + 1121, Err.Number, "Duplicate headers should throw 1121"
        Err.Clear
    End If

    On Error GoTo 0

    AssertTrue threw, "Expected duplicate header contract to throw, but it did not"
End Sub

' =============================================================================
' TEST 16: Append + schema expansion in same call
'   Scenario:
'     - PASS A: create 1-col table [id], then flush to 0 rows (keep schema)
'     - PASS B: append rows whose incoming headers are [id,a] with addMissingColumns=True
'
' Expected:
'   - After PASS A: headers=[id], rows=0
'   - After PASS B: headers=[id,a], rows=2, values correct
' =============================================================================
Public Sub Test_Append_ZeroRows_SchemaExpansion_SameCall_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_Append_ZeroRows_SchemaExpand")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' -----------------------------
    ' PASS A: seed 1-col table and then flush to 0 rows (keep schema)
    ' -----------------------------
    Dim jsonSeed As String
    jsonSeed = "[{""id"":1},{""id"":2}]"

    Dim headersSeed As Variant, dataSeed As Variant
    Build2DFromJsonRoot jsonSeed, "$", headersSeed, dataSeed

    Excel_UpsertListObjectOnSheet ws, "tAppendExpand", ws.Range("A1"), headersSeed, dataSeed, True, True, False

    Set lo = GetTable(ws, "tAppendExpand")
    AssertNotNothing lo, "tAppendExpand should exist after seed"
    AssertHeaderEquals lo, Array("id"), "PASS A seed headers"
    AssertRowCount lo, 2, "PASS A seed rows"

    ' Flush to zero, preserve schema: removeMissingColumns=True with empty JSON
    Dim jsonEmpty As String
    jsonEmpty = "[]"

    Dim headersEmpty As Variant, dataEmpty As Variant
    Build2DFromJsonRoot jsonEmpty, "$", headersEmpty, dataEmpty

    Excel_UpsertListObjectOnSheet ws, "tAppendExpand", ws.Range("A1"), headersEmpty, dataEmpty, True, False, True

    Set lo = GetTable(ws, "tAppendExpand")
    AssertHeaderEquals lo, Array("id"), "PASS A flush headers preserved"
    AssertRowCount lo, 0, "PASS A flush rows should be 0"

    ' -----------------------------
    ' PASS B: append rows that introduce a new column "a"
    ' -----------------------------
    Dim jsonB As String
    jsonB = "[{""id"":10,""a"":100},{""id"":20,""a"":200}]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonB, "$", headersB, dataB

    ' clearExisting=False => append
    ' addMissingColumns=True => schema expands to include "a"
    Excel_UpsertListObjectOnSheet ws, "tAppendExpand", ws.Range("A1"), headersB, dataB, False, True, False

    Set lo = GetTable(ws, "tAppendExpand")
    AssertHeaderEquals lo, Array("id", "a"), "PASS B headers should expand to [id,a]"
    AssertRowCount lo, 2, "PASS B rows should be 2"
    AssertTrue Not lo.DataBodyRange Is Nothing, "PASS B DataBodyRange must exist"

    AssertBodyCellEquals lo, 1, "id", 10, "PASS B row1"
    AssertBodyCellEquals lo, 1, "a", 100, "PASS B row1"
    AssertBodyCellEquals lo, 2, "id", 20, "PASS B row2"
    AssertBodyCellEquals lo, 2, "a", 200, "PASS B row2"
End Sub


' =============================================================================
' TEST 17: Append with narrower incoming headers must NOT shrink schema
'   Scenario:
'     - PASS A: create wide table [id,a,b] with 2 rows
'     - PASS B: append rows where incoming headers are ONLY [id]
'              (addMissingColumns=True, removeMissingColumns=False)
'
' Expected:
'   - After PASS B: headers still [id,a,b]
'   - Rows become 4
'   - New appended rows have id populated; a/b are Empty
' =============================================================================
Public Sub Test_Append_NarrowIncoming_DoesNotShrinkSchema_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_Append_Narrow_NoShrink")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' -----------------------------
    ' PASS A: seed wide
    ' -----------------------------
    Dim jsonA As String
    jsonA = "[{""id"":1,""a"":10,""b"":20},{""id"":2,""a"":11,""b"":21}]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA

    Excel_UpsertListObjectOnSheet ws, "tAppendNoShrink", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tAppendNoShrink")
    AssertNotNothing lo, "tAppendNoShrink should exist after PASS A"
    AssertHeaderEquals lo, Array("id", "a", "b"), "PASS A headers"
    AssertRowCount lo, 2, "PASS A rows"

    ' -----------------------------
    ' PASS B: append narrow [id] only
    ' -----------------------------
    Dim jsonB As String
    jsonB = "[{""id"":10},{""id"":20}]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonB, "$", headersB, dataB

    Excel_UpsertListObjectOnSheet ws, "tAppendNoShrink", ws.Range("A1"), headersB, dataB, False, True, False

    Set lo = GetTable(ws, "tAppendNoShrink")
    AssertHeaderEquals lo, Array("id", "a", "b"), "PASS B headers should remain wide"
    AssertRowCount lo, 4, "PASS B rows should be 4"

    ' Existing rows unchanged
    AssertBodyCellEquals lo, 1, "id", 1, "PASS B row1"
    AssertBodyCellEquals lo, 1, "a", 10, "PASS B row1"
    AssertBodyCellEquals lo, 1, "b", 20, "PASS B row1"

    ' Appended rows: id set, a/b Empty
    AssertBodyCellEquals lo, 3, "id", 10, "PASS B row3"
    AssertBodyCellEquals lo, 3, "a", Empty, "PASS B row3 a should be Empty"
    AssertBodyCellEquals lo, 3, "b", Empty, "PASS B row3 b should be Empty"

    AssertBodyCellEquals lo, 4, "id", 20, "PASS B row4"
    AssertBodyCellEquals lo, 4, "a", Empty, "PASS B row4 a should be Empty"
    AssertBodyCellEquals lo, 4, "b", Empty, "PASS B row4 b should be Empty"
End Sub


' =============================================================================
' TEST 18: Append with wider incoming headers grows schema and appends rows
'   Scenario:
'     - PASS A: create table [id,name] with 2 rows
'     - PASS B: append rows with headers [id,name,extra] (addMissingColumns=True)
'
' Expected:
'   - After PASS B: headers become [id,name,extra]
'   - Rows become 4
'   - Existing rows: extra is Empty
'   - Appended rows: extra populated where present
' =============================================================================
Public Sub Test_Append_WiderIncoming_GrowsSchema_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_Append_Wider_GrowSchema")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' -----------------------------
    ' PASS A: seed [id,name]
    ' -----------------------------
    Dim jsonA As String
    jsonA = "[{""id"":1,""name"":""A""},{""id"":2,""name"":""B""}]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA

    Excel_UpsertListObjectOnSheet ws, "tAppendGrow", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tAppendGrow")
    AssertNotNothing lo, "tAppendGrow should exist after PASS A"
    AssertHeaderEquals lo, Array("id", "name"), "PASS A headers"
    AssertRowCount lo, 2, "PASS A rows"

    ' -----------------------------
    ' PASS B: append [id,name,extra]
    ' -----------------------------
    Dim jsonB As String
    jsonB = "[" & _
        "{""id"":10,""name"":""X"",""extra"":""E1""}," & _
        "{""id"":20,""name"":""Y""}" & _
    "]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonB, "$", headersB, dataB

    Excel_UpsertListObjectOnSheet ws, "tAppendGrow", ws.Range("A1"), headersB, dataB, False, True, False

    Set lo = GetTable(ws, "tAppendGrow")
    AssertHeaderEquals lo, Array("id", "name", "extra"), "PASS B headers grew"
    AssertRowCount lo, 4, "PASS B rows should be 4"

    ' Existing rows: extra should be Empty
    AssertBodyCellEquals lo, 1, "id", 1, "row1"
    AssertBodyCellEquals lo, 1, "extra", Empty, "row1 extra Empty"
    AssertBodyCellEquals lo, 2, "id", 2, "row2"
    AssertBodyCellEquals lo, 2, "extra", Empty, "row2 extra Empty"

    ' Appended rows
    AssertBodyCellEquals lo, 3, "id", 10, "row3"
    AssertBodyCellEquals lo, 3, "name", "X", "row3"
    AssertBodyCellEquals lo, 3, "extra", "E1", "row3"

    AssertBodyCellEquals lo, 4, "id", 20, "row4"
    AssertBodyCellEquals lo, 4, "name", "Y", "row4"
    AssertBodyCellEquals lo, 4, "extra", Empty, "row4 extra Empty"
End Sub


' =============================================================================
' TEST 19: Append with addMissingColumns=False does NOT grow schema
'   Scenario:
'     - PASS A: create table [id,name] with 2 rows
'     - PASS B: append rows whose JSON includes extra fields, but we pass
'               addMissingColumns=False so schema must remain [id,name]
'
' Expected:
'   - Headers remain exactly [id,name]
'   - Rows become 4
'   - Data for id/name correct in appended rows
'   - No "extra" column created
' =============================================================================
Public Sub Test_Append_NoSchemaGrow_IgnoresExtraFields_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_Append_NoGrow")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' -----------------------------
    ' PASS A: seed [id,name]
    ' -----------------------------
    Dim jsonA As String
    jsonA = "[{""id"":1,""name"":""A""},{""id"":2,""name"":""B""}]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA

    Excel_UpsertListObjectOnSheet ws, "tAppendNoGrow", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tAppendNoGrow")
    AssertNotNothing lo, "tAppendNoGrow should exist after PASS A"
    AssertHeaderEquals lo, Array("id", "name"), "PASS A headers"
    AssertRowCount lo, 2, "PASS A rows"

    ' -----------------------------
    ' PASS B: append JSON includes extra fields
    ' -----------------------------
    Dim jsonB As String
    jsonB = "[" & _
        "{""id"":10,""name"":""X"",""extra"":""E1""}," & _
        "{""id"":20,""name"":""Y"",""extra"":""E2""}" & _
    "]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonB, "$", headersB, dataB

    ' clearExisting := False (append)
    ' addMissingColumns := False (do NOT expand schema)
    ' removeMissingColumns := False
    Excel_UpsertListObjectOnSheet ws, "tAppendNoGrow", ws.Range("A1"), headersB, dataB, False, False, False

    Set lo = GetTable(ws, "tAppendNoGrow")
    AssertHeaderEquals lo, Array("id", "name"), "PASS B headers unchanged"
    AssertRowCount lo, 4, "PASS B rows should be 4"

    ' Appended rows should have id/name only; extra ignored
    AssertBodyCellEquals lo, 3, "id", 10, "row3"
    AssertBodyCellEquals lo, 3, "name", "X", "row3"
    AssertBodyCellEquals lo, 4, "id", 20, "row4"
    AssertBodyCellEquals lo, 4, "name", "Y", "row4"

    ' Schema should not contain 'extra'
    Dim hasExtra As Boolean
    hasExtra = True
    On Error Resume Next
    Dim tmp As ListColumn
    Set tmp = lo.ListColumns("extra")
    If Err.Number <> 0 Then hasExtra = False
    Err.Clear
    On Error GoTo 0

    AssertEquals False, hasExtra, "extra column should NOT exist"
End Sub


' =============================================================================
' TEST 20: Append with addMissingColumns=True grows schema at the end
'   Scenario:
'     - PASS A: create table [id,name] with 2 rows
'     - PASS B: append rows with new fields [created_at, priority]
'              using clearExisting=False and addMissingColumns=True
'
' Expected:
'   - Headers become [id,name,created_at,priority] (existing order preserved)
'   - Rows become 4
'   - Existing rows retain their original id/name values
'   - New columns for existing rows are Empty
'   - Appended rows have their new column values where present
' =============================================================================
Public Sub Test_Append_GrowSchema_PreservesOrder_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_Append_GrowSchema")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' -----------------------------
    ' PASS A: seed [id,name]
    ' -----------------------------
    Dim jsonA As String
    jsonA = "[{""id"":1,""name"":""A""},{""id"":2,""name"":""B""}]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA

    Excel_UpsertListObjectOnSheet ws, "tAppendGrow", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tAppendGrow")
    AssertNotNothing lo, "tAppendGrow should exist after PASS A"
    AssertHeaderEquals lo, Array("id", "name"), "PASS A headers"
    AssertRowCount lo, 2, "PASS A rows"
    AssertBodyCellEquals lo, 1, "id", 1, "PASS A row1"
    AssertBodyCellEquals lo, 2, "id", 2, "PASS A row2"

    ' -----------------------------
    ' PASS B: append rows with new fields
    ' -----------------------------
    Dim jsonB As String
    jsonB = "[" & _
        "{""id"":10,""name"":""X"",""created_at"":""2026-02-28T12:00:00Z""}," & _
        "{""id"":20,""name"":""Y"",""priority"":3}" & _
    "]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonB, "$", headersB, dataB

    ' clearExisting := False (append)
    ' addMissingColumns := True (grow schema)
    ' removeMissingColumns := False
    Excel_UpsertListObjectOnSheet ws, "tAppendGrow", ws.Range("A1"), headersB, dataB, False, True, False

    Set lo = GetTable(ws, "tAppendGrow")
    AssertHeaderEquals lo, Array("id", "name", "created_at", "priority"), "PASS B headers grown"
    AssertRowCount lo, 4, "PASS B rows should be 4"

    ' Existing rows preserved; new cols should be Empty
    AssertBodyCellEquals lo, 1, "id", 1, "existing row1 id"
    AssertBodyCellEquals lo, 1, "name", "A", "existing row1 name"
    AssertBodyCellEquals lo, 1, "created_at", Empty, "existing row1 created_at empty"
    AssertBodyCellEquals lo, 1, "priority", Empty, "existing row1 priority empty"

    AssertBodyCellEquals lo, 2, "id", 2, "existing row2 id"
    AssertBodyCellEquals lo, 2, "name", "B", "existing row2 name"
    AssertBodyCellEquals lo, 2, "created_at", Empty, "existing row2 created_at empty"
    AssertBodyCellEquals lo, 2, "priority", Empty, "existing row2 priority empty"

    ' Appended rows
    AssertBodyCellEquals lo, 3, "id", 10, "appended row3 id"
    AssertBodyCellEquals lo, 3, "name", "X", "appended row3 name"
    AssertBodyCellEquals lo, 3, "created_at", "2026-02-28T12:00:00Z", "appended row3 created_at"
    AssertBodyCellEquals lo, 3, "priority", Empty, "appended row3 priority empty"

    AssertBodyCellEquals lo, 4, "id", 20, "appended row4 id"
    AssertBodyCellEquals lo, 4, "name", "Y", "appended row4 name"
    AssertBodyCellEquals lo, 4, "created_at", Empty, "appended row4 created_at empty"
    AssertBodyCellEquals lo, 4, "priority", 3, "appended row4 priority"
End Sub


' =============================================================================
' TEST 21: Existing table is NOT at topLeft; upsert must NOT move it
'   Scenario:
'     - Create table at D5 (NOT A1)
'     - Call Excel_UpsertListObjectOnSheet with topLeft=A1
'   Expected:
'     - Table still anchored at D5
'     - Data updates correctly
' =============================================================================
Public Sub Test_ExistingTable_NotAtTopLeft_DoesNotMove_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_Table_NotAtTopLeft")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' -----------------------------
    ' PASS A: create at D5
    ' -----------------------------
    Dim jsonA As String
    jsonA = "[{""id"":1,""name"":""A""},{""id"":2,""name"":""B""}]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA

    Excel_UpsertListObjectOnSheet ws, "tNotA1", ws.Range("D5"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tNotA1")
    AssertNotNothing lo, "tNotA1 should exist after PASS A"

    ' Anchor check (top-left cell of entire table range)
    AssertEquals "D5", lo.Range.Cells(1, 1).Address(False, False), "PASS A table anchor should be D5"

    AssertHeaderEquals lo, Array("id", "name"), "PASS A headers"
    AssertRowCount lo, 2, "PASS A rows"

    ' -----------------------------
    ' PASS B: update, but caller passes A1 as topLeft (should be ignored because table exists)
    ' -----------------------------
    Dim jsonB As String
    jsonB = "[{""id"":10,""name"":""X""}]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonB, "$", headersB, dataB

    Excel_UpsertListObjectOnSheet ws, "tNotA1", ws.Range("A1"), headersB, dataB, True, True, False

    Set lo = GetTable(ws, "tNotA1")
    AssertNotNothing lo, "tNotA1 should still exist after PASS B"

    ' Must NOT move
    AssertEquals "D5", lo.Range.Cells(1, 1).Address(False, False), "PASS B table anchor should remain D5"

    ' Data updated in place
    AssertRowCount lo, 1, "PASS B rows"
    AssertBodyCellEquals lo, 1, "id", 10, "PASS B row1"
    AssertBodyCellEquals lo, 1, "name", "X", "PASS B row1"
End Sub


' =============================================================================
' TEST 22: Same tableName on another sheet should fail (workbook-unique constraint)
'   Scenario:
'     - Create tNameCollision on Sheet A
'     - Try to create tNameCollision on Sheet B
'   Expected:
'     - Excel raises an error (ListObject name must be unique in workbook)
' =============================================================================
Public Sub Test_TableNameCollisionAcrossSheets_Throws()
    Dim wsA As Worksheet
    Dim wsB As Worksheet

    Set wsA = EnsureTestSheet("zTest_NameCollision_A")
    Set wsB = EnsureTestSheet("zTest_NameCollision_B")

    ResetSheetButKeepTableTestSafe wsA
    ResetSheetButKeepTableTestSafe wsB

    Dim jsonA As String
    jsonA = "[{""id"":1,""name"":""A""}]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA

    ' Create on Sheet A
    Excel_UpsertListObjectOnSheet wsA, "tNameCollision", wsA.Range("A1"), headersA, dataA, True, True, False

    ' Attempt on Sheet B should throw
    On Error GoTo ExpectedFail

    Excel_UpsertListObjectOnSheet wsB, "tNameCollision", wsB.Range("A1"), headersA, dataA, True, True, False

    ' If we get here, that's a fail (it should not succeed without renaming policy)
    Err.Raise vbObjectError + 613, "mJsonExcelTests", "ASSERT FAIL: expected name collision error, but call succeeded."

ExpectedFail:
    ' Any error is acceptable for now; we just validate that it DID fail.
    ' Clear and exit cleanly.
    Err.Clear
End Sub


' =============================================================================
' TEST 23: Header trimming is part of the contract (leading/trailing spaces)
'
' NOTE:
'   This test isolates the EXCEL contract only.
'   It does NOT rely on Json_TableTo2D / flattening to produce headers, because
'   that pipeline can introduce its own header behaviors (and should be tested
'   separately).
' =============================================================================
Public Sub Test_Header_LeadingTrailingSpaces_AreTrimmed_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_Header_Spaces_Trimmed")
    ResetSheetButKeepTableTestSafe ws

    ' Deliberately include leading/trailing spaces in header contract
    Dim headersOut As Variant
    headersOut = Array(" id ", "name")

    ' 1 row x 2 cols
    Dim data2D As Variant
    ReDim data2D(1 To 1, 1 To 2)
    data2D(1, 1) = 1
    data2D(1, 2) = "A"

    Excel_UpsertListObjectOnSheet ws, "tHdrSpaces", ws.Range("A1"), headersOut, data2D, True, True, False

    Dim lo As ListObject
    Set lo = GetTable(ws, "tHdrSpaces")
    AssertNotNothing lo, "tHdrSpaces should exist"

    ' Contract: headers are trimmed before table is materialized
    AssertHeaderEquals lo, Array("id", "name"), "Headers should be trimmed"
    AssertRowCount lo, 1, "Row count"
    AssertBodyCellEquals lo, 1, "id", 1, "Row1 id"
    AssertBodyCellEquals lo, 1, "name", "A", "Row1 name"
End Sub


' =============================================================================
' TEST 24: Union must not create post-trim duplicates
'   Existing table has "id"
'   Incoming headers include " id " (same after Trim$)
'   addMissingColumns=True should NOT append a new column, and must not throw.
' =============================================================================
Public Sub Test_UnionHeaders_DoesNotCreateTrimDuplicate()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_Union_NoTrimDup")
    ResetSheetButKeepTableTestSafe ws

    ' -------------------------
    ' Create initial table: ["id","name"]
    ' -------------------------
    Dim headers1 As Variant
    headers1 = Array("id", "name")

    Dim data1 As Variant
    ReDim data1(1 To 1, 1 To 2)
    data1(1, 1) = 1
    data1(1, 2) = "A"

    Excel_UpsertListObjectOnSheet ws, "tUnion", ws.Range("A1"), headers1, data1, True, True, False

    Dim lo As ListObject
    Set lo = GetTable(ws, "tUnion")
    AssertNotNothing lo, "tUnion should exist"

    AssertHeaderEquals lo, Array("id", "name"), "Initial schema"
    AssertRowCount lo, 1, "Initial row count"

    ' -------------------------
    ' Upsert with incoming headers that would duplicate after trim: [" id ","name"]
    ' Should NOT throw, should NOT add a 3rd column.
    ' -------------------------
    Dim headers2 As Variant
    headers2 = Array(" id ", "name")

    Dim data2 As Variant
    ReDim data2(1 To 1, 1 To 2)
    data2(1, 1) = 2
    data2(1, 2) = "B"

    Dim threw As Boolean
    threw = False

    On Error Resume Next
    Excel_UpsertListObjectOnSheet ws, "tUnion", ws.Range("A1"), headers2, data2, True, True, False
    If Err.Number <> 0 Then
        threw = True
        Err.Clear
    End If
    On Error GoTo 0

    AssertTrue (Not threw), "Union should not throw when incoming header differs only by trim"

    ' Re-get table in case Excel refreshed object refs
    Set lo = GetTable(ws, "tUnion")
    AssertNotNothing lo, "tUnion should still exist"

    ' Schema should remain 2 columns, canonicalized
    AssertHeaderEquals lo, Array("id", "name"), "Schema should remain canonical with no extra column"
    AssertRowCount lo, 1, "Row count after clearExisting=True"

    ' Data should be written into canonical "id" / "name"
    AssertBodyCellEquals lo, 1, "id", 2, "Row1 id after upsert"
    AssertBodyCellEquals lo, 1, "name", "B", "Row1 name after upsert"
End Sub

' =============================================================================
' TEST 25: Append with whitespace-variant header maps to canonical column
'
' Purpose:
'   - Verifies header normalization (Trim) during append.
'   - Ensures a header like " id " maps to existing column "id".
'   - Confirms no duplicate column is created due to leading/trailing whitespace.
'
' Scenario:
'   1) Create table with headers: id, name
'   2) Append using headers: " id ", name
'
' Expected:
'   - Final schema remains exactly: id, name
'   - Row count increments correctly
'   - Appended values land in existing canonical column
'
' Contract Locked:
'   - Header comparison is Trim + case-insensitive.
'   - Schema integrity is preserved during append.
' =============================================================================
Public Sub Test_Append_TrimVariantHeader_MapsToExistingColumn()

    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_Append_TrimVariant")
    ResetSheetButKeepTableTestSafe ws

    ' Initial table
    Dim headers1 As Variant
    headers1 = Array("id", "name")

    Dim data1 As Variant
    ReDim data1(1 To 1, 1 To 2)
    data1(1, 1) = 1
    data1(1, 2) = "A"

    Excel_UpsertListObjectOnSheet ws, "tAppend", ws.Range("A1"), headers1, data1, True, True, False

    ' Append with whitespace-variant header
    Dim headers2 As Variant
    headers2 = Array(" id ", "name")

    Dim data2 As Variant
    ReDim data2(1 To 1, 1 To 2)
    data2(1, 1) = 2
    data2(1, 2) = "B"

    Excel_UpsertListObjectOnSheet ws, "tAppend", ws.Range("A1"), headers2, data2, False, True, False

    Dim lo As ListObject
    Set lo = GetTable(ws, "tAppend")

    AssertHeaderEquals lo, Array("id", "name"), "Schema remains canonical"
    AssertRowCount lo, 2, "Row count should be 2 after append"

    AssertBodyCellEquals lo, 1, "id", 1, "Row1 id"
    AssertBodyCellEquals lo, 2, "id", 2, "Row2 id"
    AssertBodyCellEquals lo, 2, "name", "B", "Row2 name"

End Sub


' =============================================================================
' TEST 26: Nested object flattens to dotted columns (basic)
'   Scenario:
'     - Root array contains nested object "customer"
'   Expected:
'     - Headers: id, customer.name, customer.vip
'     - Values written correctly
' =============================================================================
Public Sub Test_NestedObject_FlattensToDottedColumns_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_NestedObj_DottedCols")
    ResetSheetButKeepTableTestSafe ws

    Dim jsonText As String
    jsonText = "[" & _
        "{""id"":1,""customer"":{""name"":""Alice"",""vip"":true}}," & _
        "{""id"":2,""customer"":{""name"":""Bob"",""vip"":false}}" & _
    "]"

    Dim headersOut As Variant, data2D As Variant
    Build2DFromJsonRoot jsonText, "$", headersOut, data2D

    Excel_UpsertListObjectOnSheet ws, "tNestedObj", ws.Range("A1"), headersOut, data2D, True, True, False

    Dim lo As ListObject
    Set lo = GetTable(ws, "tNestedObj")
    AssertNotNothing lo, "tNestedObj should exist"

    AssertRowCount lo, 2, "tNestedObj row count"
    AssertHeaderEquals lo, Array("id", "customer.name", "customer.vip"), "tNestedObj headers"

    AssertBodyCellEquals lo, 1, "id", 1, "row1"
    AssertBodyCellEquals lo, 1, "customer.name", "Alice", "row1"
    AssertBodyCellEquals lo, 1, "customer.vip", True, "row1"

    AssertBodyCellEquals lo, 2, "id", 2, "row2"
    AssertBodyCellEquals lo, 2, "customer.name", "Bob", "row2"
    AssertBodyCellEquals lo, 2, "customer.vip", False, "row2"
End Sub


' =============================================================================
' TEST 27: Missing nested object across rows stays sparse (no schema collapse)
'   Scenario:
'     - Row1 has customer.name
'     - Row2 has no customer object at all
'   Expected:
'     - Header includes customer.name
'     - Row2 customer.name is Empty
' =============================================================================
Public Sub Test_NestedObject_MissingAcrossRows_SparseValues_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_NestedObj_MissingSparse")
    ResetSheetButKeepTableTestSafe ws

    Dim jsonText As String
    jsonText = "[" & _
        "{""id"":1,""customer"":{""name"":""Alice""}}," & _
        "{""id"":2}" & _
    "]"

    Dim headersOut As Variant, data2D As Variant
    Build2DFromJsonRoot jsonText, "$", headersOut, data2D

    Excel_UpsertListObjectOnSheet ws, "tNestedSparse", ws.Range("A1"), headersOut, data2D, True, True, False

    Dim lo As ListObject
    Set lo = GetTable(ws, "tNestedSparse")
    AssertNotNothing lo, "tNestedSparse should exist"

    AssertRowCount lo, 2, "tNestedSparse row count"
    AssertHeaderEquals lo, Array("id", "customer.name"), "tNestedSparse headers"

    AssertBodyCellEquals lo, 1, "id", 1, "row1"
    AssertBodyCellEquals lo, 1, "customer.name", "Alice", "row1"

    AssertBodyCellEquals lo, 2, "id", 2, "row2"
    AssertBodyCellEquals lo, 2, "customer.name", Empty, "row2 missing customer => Empty"
End Sub


' =============================================================================
' TEST 28: Nested object changes shape across passes with schema union
'   Scenario:
'     - PASS A: customer.name only
'     - PASS B: customer.name + customer.vip
'   Expected:
'     - PASS B headers include both dotted fields, preserving earlier order
'     - PASS B values correct and sparsity correct
' =============================================================================
Public Sub Test_NestedObject_SchemaUnionAcrossPasses_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_NestedObj_UnionTwoPass")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' PASS A
    Dim jsonA As String
    jsonA = "[" & _
        "{""id"":1,""customer"":{""name"":""Alice""}}," & _
        "{""id"":2,""customer"":{""name"":""Bob""}}" & _
    "]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA

    Excel_UpsertListObjectOnSheet ws, "tNestedUnion", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tNestedUnion")
    AssertNotNothing lo, "tNestedUnion exists after PASS A"
    AssertHeaderEquals lo, Array("id", "customer.name"), "PASS A headers"
    AssertRowCount lo, 2, "PASS A rows"

    ' PASS B (introduce new nested field)
    Dim jsonB As String
    jsonB = "[" & _
        "{""id"":10,""customer"":{""name"":""Gamma"",""vip"":true}}," & _
        "{""id"":20,""customer"":{""name"":""Delta""}}" & _
    "]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonB, "$", headersB, dataB

    Excel_UpsertListObjectOnSheet ws, "tNestedUnion", ws.Range("A1"), headersB, dataB, True, True, False

    Set lo = GetTable(ws, "tNestedUnion")
    AssertHeaderEquals lo, Array("id", "customer.name", "customer.vip"), "PASS B headers union"
    AssertRowCount lo, 2, "PASS B rows"

    AssertBodyCellEquals lo, 1, "id", 10, "PASS B row1"
    AssertBodyCellEquals lo, 1, "customer.name", "Gamma", "PASS B row1"
    AssertBodyCellEquals lo, 1, "customer.vip", True, "PASS B row1"

    AssertBodyCellEquals lo, 2, "id", 20, "PASS B row2"
    AssertBodyCellEquals lo, 2, "customer.name", "Delta", "PASS B row2"
    AssertBodyCellEquals lo, 2, "customer.vip", Empty, "PASS B row2 missing vip => Empty"
End Sub


' =============================================================================
' TEST 29: Force schema removes nested dotted columns not in contract
'   Scenario:
'     - PASS A: build [id, customer.name, customer.vip]
'     - PASS B: force schema to [id, customer.name] only
'   Expected:
'     - customer.vip column removed
' =============================================================================
Public Sub Test_NestedObject_ForceSchema_RemovesNestedColumns_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_NestedObj_ForceRemove")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' PASS A
    Dim jsonA As String
    jsonA = "[" & _
        "{""id"":1,""customer"":{""name"":""Alice"",""vip"":true}}," & _
        "{""id"":2,""customer"":{""name"":""Bob"",""vip"":false}}" & _
    "]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA

    Excel_UpsertListObjectOnSheet ws, "tNestedForce", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tNestedForce")
    AssertHeaderEquals lo, Array("id", "customer.name", "customer.vip"), "PASS A headers"
    AssertRowCount lo, 2, "PASS A rows"

    ' PASS B: force schema contract
    Dim contractHeaders As Variant
    contractHeaders = Array("id", "customer.name")

    Dim dataB As Variant
    ReDim dataB(1 To 1, 1 To 2)
    dataB(1, 1) = 10
    dataB(1, 2) = "Gamma"

    Excel_UpsertListObjectOnSheet ws, "tNestedForce", ws.Range("A1"), contractHeaders, dataB, True, False, True

    Set lo = GetTable(ws, "tNestedForce")
    AssertHeaderEquals lo, Array("id", "customer.name"), "PASS B headers forced"
    AssertRowCount lo, 1, "PASS B rows"

    AssertBodyCellEquals lo, 1, "id", 10, "PASS B row1"
    AssertBodyCellEquals lo, 1, "customer.name", "Gamma", "PASS B row1"

    ' Ensure removed col does not exist
    Dim hasVip As Boolean
    hasVip = True
    On Error Resume Next
    Dim tmp As ListColumn
    Set tmp = lo.ListColumns("customer.vip")
    If Err.Number <> 0 Then hasVip = False
    Err.Clear
    On Error GoTo 0
    AssertEquals False, hasVip, "customer.vip column should NOT exist after force schema"
End Sub


' =============================================================================
' TEST 30: Header contract rejects post-trim duplicates for dotted fields
'   Scenario:
'     - Force schema with ["customer.name", " customer.name "]
'   Expected:
'     - Throws 1121
' =============================================================================
Public Sub Test_NestedObject_DottedHeader_TrimDup_Throws_1121()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_NestedObj_DottedTrimDup_1121")
    ResetSheetButKeepTableTestSafe ws

    ' Seed valid table first
    Dim jsonA As String
    jsonA = "[{""id"":1,""customer"":{""name"":""A""}}]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA
    Excel_UpsertListObjectOnSheet ws, "tDotDup", ws.Range("A1"), headersA, dataA, True, True, False

    ' Now force illegal contract
    Dim dupHeaders As Variant
    dupHeaders = Array("customer.name", " customer.name ")

    Dim threw As Boolean
    threw = False

    On Error Resume Next
    Excel_UpsertListObjectOnSheet ws, "tDotDup", ws.Range("A1"), dupHeaders, dataA, True, False, True
    If Err.Number <> 0 Then
        threw = True
        AssertEquals vbObjectError + 1121, Err.Number, "Dotted post-trim duplicates should throw 1121"
        Err.Clear
    End If
    On Error GoTo 0

    AssertTrue threw, "Expected dotted post-trim duplicate headers to throw, but it did not"
End Sub


' =============================================================================
' TEST 31: Nested arrays should NOT appear as "[object]" and should be stable
'   Scenario:
'     - Root array objects include nested array "tags"
'     - This is a flatten-policy test: expect NO "tags" column (or "tags.value"),
'       depending on your current flatten rules.
'
' IMPORTANT:
'   Pick ONE expected behavior and enforce it.
'   Below assumes: arrays are ignored unless explicitly extracted as table rows.
' =============================================================================
Public Sub Test_NestedArrayProperty_IsIgnoredInRootTable_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_NestedArray_Ignored")
    ResetSheetButKeepTableTestSafe ws

    Dim jsonText As String
    jsonText = "[" & _
        "{""id"":1,""name"":""A"",""tags"":[""x"",""y""]}," & _
        "{""id"":2,""name"":""B"",""tags"":[]}" & _
    "]"

    Dim headersOut As Variant, data2D As Variant
    Build2DFromJsonRoot jsonText, "$", headersOut, data2D

    Excel_UpsertListObjectOnSheet ws, "tIgnoreTags", ws.Range("A1"), headersOut, data2D, True, True, False

    Dim lo As ListObject
    Set lo = GetTable(ws, "tIgnoreTags")
    AssertNotNothing lo, "tIgnoreTags should exist"

    AssertRowCount lo, 2, "tIgnoreTags rows"

    ' Expect: only id + name
    AssertHeaderEquals lo, Array("id", "name"), "tIgnoreTags headers (tags ignored)"

    AssertBodyCellEquals lo, 1, "id", 1, "row1"
    AssertBodyCellEquals lo, 1, "name", "A", "row1"
    AssertBodyCellEquals lo, 2, "id", 2, "row2"
    AssertBodyCellEquals lo, 2, "name", "B", "row2"
End Sub


' =============================================================================
' TEST 32: Nested objects with empty object values do not create "value" fallback
'   Scenario:
'     - customer is {} in one row
'   Expected:
'     - customer.* columns exist only if other rows introduce them
'     - empty object yields Empty in those columns
' =============================================================================
Public Sub Test_NestedObject_EmptyObject_DoesNotCreateValueFallback_WithAsserts()
    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_NestedObj_EmptyObj_NoValueFallback")
    ResetSheetButKeepTableTestSafe ws

    Dim jsonText As String
    jsonText = "[" & _
        "{""id"":1,""customer"":{""name"":""Alice""}}," & _
        "{""id"":2,""customer"":{}}" & _
    "]"

    Dim headersOut As Variant, data2D As Variant
    Build2DFromJsonRoot jsonText, "$", headersOut, data2D

    Excel_UpsertListObjectOnSheet ws, "tEmptyObj", ws.Range("A1"), headersOut, data2D, True, True, False

    Dim lo As ListObject
    Set lo = GetTable(ws, "tEmptyObj")
    AssertNotNothing lo, "tEmptyObj should exist"

    AssertRowCount lo, 2, "tEmptyObj rows"
    AssertHeaderEquals lo, Array("id", "customer.name"), "tEmptyObj headers"

    AssertBodyCellEquals lo, 1, "id", 1, "row1"
    AssertBodyCellEquals lo, 1, "customer.name", "Alice", "row1"

    AssertBodyCellEquals lo, 2, "id", 2, "row2"
    AssertBodyCellEquals lo, 2, "customer.name", Empty, "row2 customer empty obj => Empty"
End Sub


' =============================================================================
' TEST 33: Flat table -> JSON array-of-objects (basic)
'
' Scenario:
'   - Table headers: id, name, active
'   - 2 rows
'
' Expected:
'   - Excel_ListObjectToJson returns JSON array-of-objects
'   - Parsing + Flatten + ExtractTableRows + TableTo2D reproduces the same schema + values
'
' Notes:
'   - Self-contained: no EnsureTestSheet / Reset helpers / GetTable / Assert helpers.
'   - Uses only the production APIs in zz_ModernJsonInVba plus Excel object model.
' =============================================================================
Public Sub Test_TableToJson_FlatBasic_WithAsserts()

    Const SRC As String = "Test_TableToJson_FlatBasic_WithAsserts"

    On Error GoTo Fail

    ' -----------------------------
    ' Create / reset a dedicated sheet
    ' -----------------------------
    Dim ws As Worksheet
    Set ws = Nothing

    Dim shName As String
    shName = "zTest_TableToJson_FlatBasic"

    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If StrComp(sh.Name, shName, vbTextCompare) = 0 Then
            Set ws = sh
            Exit For
        End If
    Next sh

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = shName
    Else
        ws.Cells.Clear
        ' Remove any tables (deterministic clean slate)
        Dim iLo As Long
        For iLo = ws.ListObjects.count To 1 Step -1
            ws.ListObjects(iLo).Delete
        Next iLo
    End If

    ' -----------------------------
    ' Build table inputs
    ' -----------------------------
    Dim headers As Variant
    headers = Array("id", "name", "active")

    Dim data2D As Variant
    ReDim data2D(1 To 2, 1 To 3)
    data2D(1, 1) = 1: data2D(1, 2) = "Alpha": data2D(1, 3) = True
    data2D(2, 1) = 2: data2D(2, 2) = "Beta":  data2D(2, 3) = False

    ' -----------------------------
    ' Upsert into ListObject
    ' -----------------------------
    Excel_UpsertListObjectOnSheet ws, "tToJsonFlat", ws.Range("A1"), headers, data2D, True, True, False

    ' Locate table (self-contained: no GetTable helper)
    Dim lo As ListObject
    Set lo = Nothing

    Dim t As ListObject
    For Each t In ws.ListObjects
        If StrComp(t.Name, "tToJsonFlat", vbTextCompare) = 0 Then
            Set lo = t
            Exit For
        End If
    Next t

    If lo Is Nothing Then
        Err.Raise vbObjectError + 620, SRC, "Expected ListObject 'tToJsonFlat' was not created."
    End If

    ' Basic table assertions (schema + row count)
    If lo.ListColumns.count <> 3 Then
        Err.Raise vbObjectError + 621, SRC, "Expected 3 columns; got " & CStr(lo.ListColumns.count) & "."
    End If

    If StrComp(lo.ListColumns(1).Name, "id", vbTextCompare) <> 0 _
        Or StrComp(lo.ListColumns(2).Name, "name", vbTextCompare) <> 0 _
        Or StrComp(lo.ListColumns(3).Name, "active", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 622, SRC, "Header mismatch. Expected: id,name,active."
    End If

    Dim bodyRows As Long
    If lo.DataBodyRange Is Nothing Then
        bodyRows = 0
    Else
        bodyRows = lo.DataBodyRange.rows.count
    End If

    If bodyRows <> 2 Then
        Err.Raise vbObjectError + 623, SRC, "Expected 2 rows; got " & CStr(bodyRows) & "."
    End If

    ' -----------------------------
    ' Convert to JSON and round-trip through existing JSON pipeline
    ' -----------------------------
    Dim jsonOut As String
    jsonOut = Excel_ListObjectToJson(lo, False)

    Dim parsed As Variant
    Json_ParseInto jsonOut, parsed

    Dim flat As Collection
    Set flat = Json_Flatten(parsed)

    Dim rows As Collection
    Set rows = Json_ExtractTableRows(flat, "$")

    Dim headersBack As Variant
    Dim dataBack As Variant
    dataBack = Json_TableTo2D(rows, headersBack)

    ' -----------------------------
    ' Assertions: headersBack
    ' Json_TableTo2D returns 1-based headers array.
    ' -----------------------------
    If (UBound(headersBack) - LBound(headersBack) + 1) <> 3 Then
        Err.Raise vbObjectError + 624, SRC, "Round-trip header count mismatch."
    End If

    If StrComp(CStr(headersBack(1)), "id", vbTextCompare) <> 0 _
        Or StrComp(CStr(headersBack(2)), "name", vbTextCompare) <> 0 _
        Or StrComp(CStr(headersBack(3)), "active", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 625, SRC, "Round-trip headers mismatch. Expected: id,name,active."
    End If

    ' -----------------------------
    ' Assertions: dataBack (2D)
    ' -----------------------------
    If IsEmpty(dataBack) Then
        Err.Raise vbObjectError + 626, SRC, "Round-trip data unexpectedly Empty."
    End If

    Dim rb As Long, cb As Long
    rb = UBound(dataBack, 1) - LBound(dataBack, 1) + 1
    cb = UBound(dataBack, 2) - LBound(dataBack, 2) + 1

    If rb <> 2 Or cb <> 3 Then
        Err.Raise vbObjectError + 627, SRC, "Round-trip data shape mismatch. Expected 2x3."
    End If

    ' Compare values (normalize indices to 1-based in our local expected array)
    Dim r As Long, c As Long
    For r = 1 To 2
        For c = 1 To 3
            Dim expV As Variant, actV As Variant
            expV = data2D(r, c)
            actV = dataBack(LBound(dataBack, 1) + r - 1, LBound(dataBack, 2) + c - 1)

            ' Treat Null=Null as equal; otherwise use <> for primitives.
            If IsNull(expV) And IsNull(actV) Then
                ' ok
            ElseIf VarType(expV) = vbBoolean Or VarType(actV) = vbBoolean Then
                If CBool(expV) <> CBool(actV) Then
                    Err.Raise vbObjectError + 628, SRC, _
                        "Value mismatch at (r=" & r & ",c=" & c & "). Expected=" & CStr(expV) & " Actual=" & CStr(actV)
                End If
            Else
                If expV <> actV Then
                    Err.Raise vbObjectError + 628, SRC, _
                        "Value mismatch at (r=" & r & ",c=" & c & "). Expected=" & CStr(expV) & " Actual=" & CStr(actV)
                End If
            End If
        Next c
    Next r

    Exit Sub

Fail:
    ' Re-raise with stable source boundary for the test.
    Dim n As Long: n = Err.Number
    Dim d As String: d = Err.Description
    Err.Clear
    Err.Raise n, SRC, d
End Sub


' =============================================================================
' TEST 34: Tagged object contract (TAG_OBJECT) + deterministic failure when untagged
'
' Updated for new Json_Stringify behavior:
'   - Json_Stringify only throws vbObjectError+1134 when an *untagged* Collection
'     "looks like an object" (contains key/value pair entries), not for any
'     arbitrary untagged Collection.
'
' Purpose:
'   - Asserts Json_Stringify throws vbObjectError+1134 for an untagged, object-shaped Collection.
'   - Asserts Json_Stringify succeeds for a properly-tagged object and round-trips.
'
' Self-contained:
'   - No helper functions.
'   - Uses only production APIs + internal pair scanning:
'       Json_ObjSet, Json_ParseInto, Json_Stringify
' =============================================================================
Public Sub Test_TaggedObject_TagConstant_And_Untagged_Throws_1134()

    Const SRC As String = "Test_TaggedObject_TagConstant_And_Untagged_Throws_1134"

    On Error GoTo Fail

    ' -----------------------------
    ' Create / reset a dedicated sheet
    ' -----------------------------
    Dim ws As Worksheet
    Dim shName As String
    shName = "zTest_33_TaggedObject"

    Set ws = Nothing

    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If StrComp(sh.Name, shName, vbTextCompare) = 0 Then
            Set ws = sh
            Exit For
        End If
    Next sh

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = shName
    Else
        ws.Cells.Clear
    End If

    ws.Range("A1").Value2 = "Test 33: Tagged object contract"
    ws.Range("A2").Value2 = "Status"
    ws.Range("B2").Value2 = "RUNNING..."

    ' -----------------------------
    ' 1) Untagged, object-shaped Collection must throw 1134
    '    (new behavior: only object-shaped untagged triggers 1134)
    ' -----------------------------
    Dim untagged As New Collection
    untagged.Add Array("id", 123)   ' <- looks like an object pair, but collection is NOT tagged

    Dim threw As Boolean
    Dim gotErr As Long
    threw = False
    gotErr = 0

    On Error Resume Next
    Dim sBad As String
    sBad = Json_Stringify(untagged)  ' must raise 1134
    gotErr = Err.Number
    If gotErr <> 0 Then threw = True
    Err.Clear
    On Error GoTo 0

    If Not threw Then
        Err.Raise vbObjectError + 633, SRC, "Expected Json_Stringify(untagged object-shaped Collection) to throw, but it did not."
    End If

    If gotErr <> vbObjectError + 1134 Then
        Err.Raise vbObjectError + 634, SRC, _
            "Expected error " & CStr(vbObjectError + 1134) & " but got " & CStr(gotErr) & "."
    End If

    ws.Range("A4").Value2 = "Untagged object-shaped stringify throws 1134"
    ws.Range("B4").Value2 = "PASS"

    ' -----------------------------
    ' 2) Properly tagged object stringifies and round-trips
    '    Note: TAG_OBJECT is Private in the engine module, so we tag using the literal "__OBJ__".
    '    This still verifies the engine is checking the correct tag value internally.
    ' -----------------------------
    Dim obj As New Collection
    obj.Add "__OBJ__"               ' must be slot(1)

    Json_ObjSet obj, "id", 7
    Json_ObjSet obj, "name", "Alpha"
    Json_ObjSet obj, "active", True

    Dim jsonOut As String
    jsonOut = Json_Stringify(obj)

    If Len(jsonOut) = 0 Then
        Err.Raise vbObjectError + 635, SRC, "Json_Stringify(tagged object) returned empty string."
    End If

    ws.Range("A6").Value2 = "Tagged stringify produced JSON"
    ws.Range("B6").Value2 = "PASS"
    ws.Range("A7").Value2 = "JSON"
    ws.Range("B7").Value2 = jsonOut

    ' Round-trip: parse back and validate it's an object and tagged
    Dim parsed As Variant
    Json_ParseInto jsonOut, parsed

    If Not IsObject(parsed) Then
        Err.Raise vbObjectError + 636, SRC, "Round-trip parse did not return an object."
    End If

    If TypeName(parsed) <> "Collection" Then
        Err.Raise vbObjectError + 637, SRC, "Round-trip parse returned unexpected type: " & TypeName(parsed)
    End If

    Dim pobj As Collection
    Set pobj = parsed

    If pobj.count < 1 Or CStr(pobj(1)) <> "__OBJ__" Then
        Err.Raise vbObjectError + 638, SRC, "Round-trip object missing tag '__OBJ__' at index 1."
    End If

    ' -----------------------------
    ' Spot-check values (no Obj_Get helpers):
    '   scan pobj(2..) entries which are expected to be Array(key,value)
    ' -----------------------------
    Dim haveId As Boolean, haveName As Boolean, haveActive As Boolean
    Dim vId As Variant, vName As Variant, vActive As Variant
    haveId = False: haveName = False: haveActive = False

    Dim i As Long
    For i = 2 To pobj.count
        Dim entry As Variant
        entry = pobj(i)

        If IsArray(entry) Then
            Dim lb As Long
            lb = LBound(entry)

            If (UBound(entry) - lb + 1) >= 2 Then
                Dim k As String
                k = CStr(entry(lb))

                Select Case LCase$(k)
                    Case "id"
                        vId = entry(lb + 1)
                        haveId = True
                    Case "name"
                        vName = entry(lb + 1)
                        haveName = True
                    Case "active"
                        vActive = entry(lb + 1)
                        haveActive = True
                End Select
            End If
        End If
    Next i

    If Not haveId Then Err.Raise vbObjectError + 642, SRC, "Key not found after round-trip: id"
    If Not haveName Then Err.Raise vbObjectError + 642, SRC, "Key not found after round-trip: name"
    If Not haveActive Then Err.Raise vbObjectError + 642, SRC, "Key not found after round-trip: active"

    If CLng(vId) <> 7 Then
        Err.Raise vbObjectError + 639, SRC, "Round-trip id mismatch. Expected 7; got " & CStr(vId) & "."
    End If

    If CStr(vName) <> "Alpha" Then
        Err.Raise vbObjectError + 640, SRC, "Round-trip name mismatch. Expected Alpha; got " & CStr(vName) & "."
    End If

    If CBool(vActive) <> True Then
        Err.Raise vbObjectError + 641, SRC, "Round-trip active mismatch. Expected True; got " & CStr(vActive) & "."
    End If

    ws.Range("A9").Value2 = "Tagged round-trip fields match"
    ws.Range("B9").Value2 = "PASS"

    ws.Range("B2").Value2 = "PASS"
    Exit Sub

Fail:
    ws.Range("B2").Value2 = "FAIL"
    ws.Range("A12").Value2 = "Err"
    ws.Range("B12").Value2 = Err.Number
    ws.Range("A13").Value2 = "Source"
    ws.Range("B13").Value2 = SRC
    ws.Range("A14").Value2 = "Description"
    ws.Range("B14").Value2 = Err.Description

    Dim n As Long: n = Err.Number
    Dim d As String: d = Err.Description
    Err.Clear
    Err.Raise n, SRC, d
End Sub


' =============================================================================
' TEST 35: ListObject -> JSON (nested paths, escaped dots, blanks policy)
'
' Purpose:
'   - Verifies header path unflatten:
'       "person.name" => nested object {"person":{"name":...}}
'   - Verifies escaped dot header:
'       "meta\.version" => single key "meta.version" (NOT nested)
'   - Verifies blanks policy:
'       includeBlanksAsNull=False => blank cells omit key entirely
'       includeBlanksAsNull=True  => blank cells emit key with JSON null
'
' Uses:
'   Excel_UpsertListObjectOnSheet
'   Excel_ListObjectToJson
'   Json_ParseInto
'   Json_Flatten
'   Json_FlatGet
'   Json_FlatContains
'
' Expected:
'   - For includeBlanksAsNull=False:
'       "$[0].note" NOT present when blank
'   - For includeBlanksAsNull=True:
'       "$[0].note" present and is Null when blank
' =============================================================================
Public Sub Test_TableToJson_Nested_EscapedDots_And_BlanksPolicy()

    Const SRC As String = "Test_TableToJson_Nested_EscapedDots_And_BlanksPolicy"

    On Error GoTo Fail

    ' -----------------------------
    ' Create / reset a dedicated sheet
    ' -----------------------------
    Dim ws As Worksheet
    Dim shName As String
    shName = "zTest_34_TableToJson_Nested"

    Set ws = Nothing

    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If StrComp(sh.Name, shName, vbTextCompare) = 0 Then
            Set ws = sh
            Exit For
        End If
    Next sh

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = shName
    Else
        ws.Cells.Clear
        Dim iLo As Long
        For iLo = ws.ListObjects.count To 1 Step -1
            ws.ListObjects(iLo).Delete
        Next iLo
    End If

    ws.Range("A1").Value2 = "Test 34: ListObject -> JSON (nested paths, escaped dots, blanks policy)"
    ws.Range("A2").Value2 = "Status"
    ws.Range("B2").Value2 = "RUNNING..."

    ' -----------------------------
    ' Build table inputs
    ' -----------------------------
    Dim headers As Variant
    headers = Array( _
        "id", _
        "person.name", _
        "person.active", _
        "meta\.version", _
        "note" _
    )

    Dim data2D As Variant
    ReDim data2D(1 To 2, 1 To 5)

    ' Row 1: note blank
    data2D(1, 1) = 1
    data2D(1, 2) = "Ada"
    data2D(1, 3) = True
    data2D(1, 4) = 1
    data2D(1, 5) = ""          ' blank note

    ' Row 2: note populated
    data2D(2, 1) = 2
    data2D(2, 2) = "Bob"
    data2D(2, 3) = False
    data2D(2, 4) = 2
    data2D(2, 5) = "has-note"

    ' -----------------------------
    ' Upsert into ListObject
    ' -----------------------------
    Excel_UpsertListObjectOnSheet ws, "tToJsonNested", ws.Range("A4"), headers, data2D, True, True, False

    Dim lo As ListObject
    Set lo = Nothing

    Dim t As ListObject
    For Each t In ws.ListObjects
        If StrComp(t.Name, "tToJsonNested", vbTextCompare) = 0 Then
            Set lo = t
            Exit For
        End If
    Next t

    If lo Is Nothing Then
        Err.Raise vbObjectError + 650, SRC, "Expected ListObject 'tToJsonNested' was not created."
    End If

    ' -----------------------------
    ' Case A: includeBlanksAsNull = False (blank note omitted)
    ' -----------------------------
    Dim jsonA As String
    jsonA = Excel_ListObjectToJson(lo, False)

    ws.Range("A6").Value2 = "JSON (includeBlanksAsNull=False)"
    ws.Range("B6").Value2 = jsonA

    Dim parsedA As Variant
    Json_ParseInto jsonA, parsedA

    Dim flatA As Collection
    Set flatA = Json_Flatten(parsedA)

    ' Nested path assertions
    If CLng(Json_FlatGet(flatA, "$[0].id")) <> 1 Then
        Err.Raise vbObjectError + 651, SRC, "Case A: $[0].id mismatch."
    End If

    If CStr(Json_FlatGet(flatA, "$[0].person.name")) <> "Ada" Then
        Err.Raise vbObjectError + 652, SRC, "Case A: $[0].person.name mismatch."
    End If

    If CBool(Json_FlatGet(flatA, "$[0].person.active")) <> True Then
        Err.Raise vbObjectError + 653, SRC, "Case A: $[0].person.active mismatch."
    End If

    ' Escaped dot key assertion:
    ' Header "meta\.version" becomes key "meta.version"
    ' Flatten path re-escapes dot => "$[0].meta\.version"
    If CLng(Json_FlatGet(flatA, "$[0].meta\.version")) <> 1 Then
        Err.Raise vbObjectError + 654, SRC, "Case A: $[0].meta\.version mismatch."
    End If

    ' Blank policy: note is blank in row 1 => OMIT key when includeBlanksAsNull=False
    If Json_FlatContains(flatA, "$[0].note") Then
        Err.Raise vbObjectError + 655, SRC, "Case A: expected $[0].note to be absent, but it exists."
    End If

    ' Row 2 note exists
    If CStr(Json_FlatGet(flatA, "$[1].note")) <> "has-note" Then
        Err.Raise vbObjectError + 656, SRC, "Case A: $[1].note mismatch."
    End If

    ws.Range("A8").Value2 = "Case A (blank omitted) assertions"
    ws.Range("B8").Value2 = "PASS"

    ' -----------------------------
    ' Case B: includeBlanksAsNull = True (blank note becomes null)
    ' -----------------------------
    Dim jsonB As String
    jsonB = Excel_ListObjectToJson(lo, True)

    ws.Range("A10").Value2 = "JSON (includeBlanksAsNull=True)"
    ws.Range("B10").Value2 = jsonB

    Dim parsedB As Variant
    Json_ParseInto jsonB, parsedB

    Dim flatB As Collection
    Set flatB = Json_Flatten(parsedB)

    ' Same nested/escaped-dot checks for row 1
    If CLng(Json_FlatGet(flatB, "$[0].id")) <> 1 Then
        Err.Raise vbObjectError + 657, SRC, "Case B: $[0].id mismatch."
    End If

    If CStr(Json_FlatGet(flatB, "$[0].person.name")) <> "Ada" Then
        Err.Raise vbObjectError + 658, SRC, "Case B: $[0].person.name mismatch."
    End If

    If CBool(Json_FlatGet(flatB, "$[0].person.active")) <> True Then
        Err.Raise vbObjectError + 659, SRC, "Case B: $[0].person.active mismatch."
    End If

    If CLng(Json_FlatGet(flatB, "$[0].meta\.version")) <> 1 Then
        Err.Raise vbObjectError + 660, SRC, "Case B: $[0].meta\.version mismatch."
    End If

    ' Blank policy: note is blank in row 1 => PRESENT and Null when includeBlanksAsNull=True
    If Not Json_FlatContains(flatB, "$[0].note") Then
        Err.Raise vbObjectError + 661, SRC, "Case B: expected $[0].note to exist, but it is absent."
    End If

    Dim vNote0 As Variant
    vNote0 = Json_FlatGet(flatB, "$[0].note")
    If Not IsNull(vNote0) Then
        Err.Raise vbObjectError + 662, SRC, "Case B: expected $[0].note to be Null."
    End If

    ' Row 2 note exists
    If CStr(Json_FlatGet(flatB, "$[1].note")) <> "has-note" Then
        Err.Raise vbObjectError + 663, SRC, "Case B: $[1].note mismatch."
    End If

    ws.Range("A12").Value2 = "Case B (blank => null) assertions"
    ws.Range("B12").Value2 = "PASS"

    ws.Range("B2").Value2 = "PASS"
    Exit Sub

Fail:
    ws.Range("B2").Value2 = "FAIL"
    ws.Range("A14").Value2 = "Err"
    ws.Range("B14").Value2 = Err.Number
    ws.Range("A15").Value2 = "Source"
    ws.Range("B15").Value2 = SRC
    ws.Range("A16").Value2 = "Description"
    ws.Range("B16").Value2 = Err.Description

    Dim n As Long: n = Err.Number
    Dim d As String: d = Err.Description
    Err.Clear
    Err.Raise n, SRC, d
End Sub


' =============================================================================
' TEST 36: ListObject -> JSON rejects array-index header paths ([ ] ) with +905
'
' Purpose:
'   - Verifies Excel_ListObjectToJson enforces the "no array index path" contract.
'   - Any header containing "[" or "]" must raise vbObjectError+905.
'
' Uses:
'   Excel_UpsertListObjectOnSheet
'   Excel_ListObjectToJson
'
' Expected:
'   - Excel_ListObjectToJson throws vbObjectError + 905
' =============================================================================
Public Sub Test_TableToJson_RejectsArrayIndexHeaders_Throws_905()

    Const SRC As String = "Test_TableToJson_RejectsArrayIndexHeaders_Throws_905"

    On Error GoTo Fail

    ' -----------------------------
    ' Create / reset a dedicated sheet
    ' -----------------------------
    Dim ws As Worksheet
    Dim shName As String
    shName = "zTest_35_TableToJson_BadHeader"

    Set ws = Nothing

    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If StrComp(sh.Name, shName, vbTextCompare) = 0 Then
            Set ws = sh
            Exit For
        End If
    Next sh

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = shName
    Else
        ws.Cells.Clear
        Dim iLo As Long
        For iLo = ws.ListObjects.count To 1 Step -1
            ws.ListObjects(iLo).Delete
        Next iLo
    End If

    ws.Range("A1").Value2 = "Test 35: ListObject -> JSON rejects array-index headers"
    ws.Range("A2").Value2 = "Status"
    ws.Range("B2").Value2 = "RUNNING..."

    ' -----------------------------
    ' Build table inputs (intentionally invalid header path)
    ' -----------------------------
    Dim headers As Variant
    headers = Array("id", "tags[0]")   ' <-- must be rejected by Excel_ListObjectToJson

    Dim data2D As Variant
    ReDim data2D(1 To 1, 1 To 2)
    data2D(1, 1) = 1
    data2D(1, 2) = "x"

    ' -----------------------------
    ' Create table
    ' -----------------------------
    Excel_UpsertListObjectOnSheet ws, "tBadHeader", ws.Range("A4"), headers, data2D, True, True, False

    Dim lo As ListObject
    Set lo = Nothing

    Dim t As ListObject
    For Each t In ws.ListObjects
        If StrComp(t.Name, "tBadHeader", vbTextCompare) = 0 Then
            Set lo = t
            Exit For
        End If
    Next t

    If lo Is Nothing Then
        Err.Raise vbObjectError + 670, SRC, "Expected ListObject 'tBadHeader' was not created."
    End If

    ' -----------------------------
    ' Attempt stringify -> must throw +905
    ' -----------------------------
    Dim threw As Boolean
    Dim gotErr As Long
    Dim gotDesc As String
    threw = False
    gotErr = 0
    gotDesc = vbNullString

    On Error Resume Next
    Dim jsonOut As String
    jsonOut = Excel_ListObjectToJson(lo, False)
    gotErr = Err.Number
    gotDesc = Err.Description
    If gotErr <> 0 Then threw = True
    Err.Clear
    On Error GoTo 0

    If Not threw Then
        Err.Raise vbObjectError + 671, SRC, "Expected Excel_ListObjectToJson to throw for header containing [ ], but it did not."
    End If

    If gotErr <> vbObjectError + 905 Then
        Err.Raise vbObjectError + 672, SRC, _
            "Expected error " & CStr(vbObjectError + 905) & " but got " & CStr(gotErr) & ". Desc=" & gotDesc
    End If

    ws.Range("A6").Value2 = "Excel_ListObjectToJson throws +905 on '[' or ']' headers"
    ws.Range("B6").Value2 = "PASS"

    ws.Range("B2").Value2 = "PASS"
    Exit Sub

Fail:
    ws.Range("B2").Value2 = "FAIL"
    ws.Range("A8").Value2 = "Err"
    ws.Range("B8").Value2 = Err.Number
    ws.Range("A9").Value2 = "Source"
    ws.Range("B9").Value2 = SRC
    ws.Range("A10").Value2 = "Description"
    ws.Range("B10").Value2 = Err.Description

    Dim n As Long: n = Err.Number
    Dim d As String: d = Err.Description
    Err.Clear
    Err.Raise n, SRC, d
End Sub


' =============================================================================
' TEST 37: Header path unflatten + escaped-dot key round-trip
'
' Purpose:
'   Highest-value contract test for Excel_ListObjectToJson:
'     - "a.b.c" headers create nested objects
'     - "\." in a header segment means a literal dot in the JSON key
'     - Round-trip through Parse -> Flatten -> ExtractTableRows("$") -> TableTo2D
'       preserves schema + values deterministically
'
' Scenario:
'   Headers:
'     id
'     profile.name
'     profile.meta\.version     ' literal key "meta.version" under profile
'
'   Row:
'     7, "Alpha", 2
'
' Expected:
'   - JSON parses
'   - Flattened table headers back are exactly:
'       id, profile.name, profile.meta\.version
'   - Values match: 7, "Alpha", 2
'
' Notes:
'   - Self-contained: no helper modules, no Assert helpers.
'   - Uses only production APIs + Excel object model.
' =============================================================================
Public Sub Test_TableToJson_NestedPaths_And_EscapedDotKey()

    Const SRC As String = "Test_TableToJson_NestedPaths_And_EscapedDotKey"

    On Error GoTo Fail

    ' -----------------------------
    ' Create / reset a dedicated sheet
    ' -----------------------------
    Dim ws As Worksheet
    Set ws = Nothing

    Dim shName As String
    shName = "zTest_36_NestedEscDot"   ' <= 31 chars (fixes Excel error 1004)

    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If StrComp(sh.Name, shName, vbTextCompare) = 0 Then
            Set ws = sh
            Exit For
        End If
    Next sh

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = shName
    Else
        ws.Cells.Clear
        Dim iLo As Long
        For iLo = ws.ListObjects.count To 1 Step -1
            ws.ListObjects(iLo).Delete
        Next iLo
    End If

    ' -----------------------------
    ' Build table inputs
    ' -----------------------------
    Dim headers As Variant
    headers = Array("id", "profile.name", "profile.meta\.version")

    Dim data2D As Variant
    ReDim data2D(1 To 1, 1 To 3)
    data2D(1, 1) = 7
    data2D(1, 2) = "Alpha"
    data2D(1, 3) = 2

    ' -----------------------------
    ' Upsert into ListObject
    ' -----------------------------
    Excel_UpsertListObjectOnSheet ws, "tToJsonNested", ws.Range("A1"), headers, data2D, True, True, False

    Dim lo As ListObject
    Set lo = Nothing

    Dim t As ListObject
    For Each t In ws.ListObjects
        If StrComp(t.Name, "tToJsonNested", vbTextCompare) = 0 Then
            Set lo = t
            Exit For
        End If
    Next t

    If lo Is Nothing Then
        Err.Raise vbObjectError + 650, SRC, "Expected ListObject 'tToJsonNested' was not created."
    End If

    ' -----------------------------
    ' Convert to JSON
    ' -----------------------------
    Dim jsonOut As String
    jsonOut = Excel_ListObjectToJson(lo, False)

    If Len(jsonOut) = 0 Then
        Err.Raise vbObjectError + 651, SRC, "Excel_ListObjectToJson returned empty JSON."
    End If

    If InStr(1, jsonOut, """profile""", vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 652, SRC, "Expected JSON to contain 'profile' object."
    End If

    ' -----------------------------
    ' Round-trip through existing JSON pipeline
    ' -----------------------------
    Dim parsed As Variant
    Json_ParseInto jsonOut, parsed

    Dim flat As Collection
    Set flat = Json_Flatten(parsed)

    Dim rows As Collection
    Set rows = Json_ExtractTableRows(flat, "$")

    Dim headersBack As Variant
    Dim dataBack As Variant
    dataBack = Json_TableTo2D(rows, headersBack)

    ' -----------------------------
    ' Assertions: row count and schema
    ' -----------------------------
    If rows.count <> 1 Then
        Err.Raise vbObjectError + 653, SRC, "Expected 1 row extracted; got " & CStr(rows.count) & "."
    End If

    If (UBound(headersBack) - LBound(headersBack) + 1) <> 3 Then
        Err.Raise vbObjectError + 654, SRC, "Expected 3 headers back; got " & CStr(UBound(headersBack) - LBound(headersBack) + 1) & "."
    End If

    If StrComp(CStr(headersBack(1)), "id", vbTextCompare) <> 0 _
        Or StrComp(CStr(headersBack(2)), "profile.name", vbTextCompare) <> 0 _
        Or StrComp(CStr(headersBack(3)), "profile.meta\.version", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 655, SRC, _
            "Headers mismatch. Expected: id, profile.name, profile.meta\.version. " & _
            "Got: " & CStr(headersBack(1)) & ", " & CStr(headersBack(2)) & ", " & CStr(headersBack(3)) & "."
    End If

    If IsEmpty(dataBack) Then
        Err.Raise vbObjectError + 656, SRC, "dataBack unexpectedly Empty."
    End If

    Dim rb As Long, cb As Long
    rb = UBound(dataBack, 1) - LBound(dataBack, 1) + 1
    cb = UBound(dataBack, 2) - LBound(dataBack, 2) + 1

    If rb <> 1 Or cb <> 3 Then
        Err.Raise vbObjectError + 657, SRC, "Data shape mismatch. Expected 1x3; got " & CStr(rb) & "x" & CStr(cb) & "."
    End If

    ' -----------------------------
    ' Assertions: values
    ' -----------------------------
    Dim vId As Variant, vName As Variant, vVer As Variant
    vId = dataBack(LBound(dataBack, 1), LBound(dataBack, 2) + 0)
    vName = dataBack(LBound(dataBack, 1), LBound(dataBack, 2) + 1)
    vVer = dataBack(LBound(dataBack, 1), LBound(dataBack, 2) + 2)

    If CLng(vId) <> 7 Then
        Err.Raise vbObjectError + 658, SRC, "id mismatch. Expected 7; got " & CStr(vId) & "."
    End If

    If CStr(vName) <> "Alpha" Then
        Err.Raise vbObjectError + 659, SRC, "profile.name mismatch. Expected Alpha; got " & CStr(vName) & "."
    End If

    If CLng(vVer) <> 2 Then
        Err.Raise vbObjectError + 660, SRC, "profile.meta\.version mismatch. Expected 2; got " & CStr(vVer) & "."
    End If

    Exit Sub

Fail:
    Dim n As Long: n = Err.Number
    Dim d As String: d = Err.Description
    Err.Clear
    Err.Raise n, SRC, d
End Sub


' =============================================================================
' TEST 38: Deterministic header discovery with sparse objects (missing keys)
'
' Purpose:
'   - Validates deterministic "first-seen" header discovery order when rows have
'     different sets of keys (sparse objects).
'   - Validates missing keys materialize as Empty in the 2D table output.
'
' Scenario JSON (array-of-objects):
'   [
'     {"b":2,"a":1},
'     {"c":3},
'     {"a":10,"c":30}
'   ]
'
' Expected headersBack order (first-seen across scan):
'   b, a, c
'
' Expected dataBack shape:
'   3 rows x 3 cols
'
' Expected values:
'   row1: b=2,  a=1,  c=Empty
'   row2: b=Empty, a=Empty, c=3
'   row3: b=Empty, a=10, c=30
'
' Notes:
'   - Self-contained: no helper modules, no Assert helpers.
'   - Uses only: Json_ParseInto, Json_Flatten, Json_ExtractTableRows, Json_TableTo2D.
' =============================================================================
Public Sub Test_DeterministicHeaders_SparseObjects()

    Const SRC As String = "Test_DeterministicHeaders_SparseObjects"
    On Error GoTo Fail

    Dim jsonText As String
    jsonText = "[" & _
        "{""b"":2,""a"":1}," & _
        "{""c"":3}," & _
        "{""a"":10,""c"":30}" & _
    "]"

    Dim parsed As Variant
    Json_ParseInto jsonText, parsed

    Dim flat As Collection
    Set flat = Json_Flatten(parsed)

    Dim rows As Collection
    Set rows = Json_ExtractTableRows(flat, "$")

    Dim headersBack As Variant
    Dim dataBack As Variant
    dataBack = Json_TableTo2D(rows, headersBack)

    ' ---- basic shape
    If rows.count <> 3 Then
        Err.Raise vbObjectError + 670, SRC, "Expected 3 rows; got " & CStr(rows.count) & "."
    End If

    If (UBound(headersBack) - LBound(headersBack) + 1) <> 3 Then
        Err.Raise vbObjectError + 671, SRC, "Expected 3 headers; got " & CStr(UBound(headersBack) - LBound(headersBack) + 1) & "."
    End If

    Dim rb As Long, cb As Long
    rb = UBound(dataBack, 1) - LBound(dataBack, 1) + 1
    cb = UBound(dataBack, 2) - LBound(dataBack, 2) + 1

    If rb <> 3 Or cb <> 3 Then
        Err.Raise vbObjectError + 672, SRC, "Expected data shape 3x3; got " & CStr(rb) & "x" & CStr(cb) & "."
    End If

    ' ---- deterministic header order: b, a, c
    If StrComp(CStr(headersBack(1)), "b", vbTextCompare) <> 0 _
        Or StrComp(CStr(headersBack(2)), "a", vbTextCompare) <> 0 _
        Or StrComp(CStr(headersBack(3)), "c", vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 673, SRC, _
            "Header order mismatch. Expected: b,a,c. Got: " & _
            CStr(headersBack(1)) & "," & CStr(headersBack(2)) & "," & CStr(headersBack(3)) & "."
    End If

    ' convenience: base indices
    Dim r0 As Long, c0 As Long
    r0 = LBound(dataBack, 1)
    c0 = LBound(dataBack, 2)

    ' ---- row 1: b=2, a=1, c=Empty
    If CLng(dataBack(r0 + 0, c0 + 0)) <> 2 Then
        Err.Raise vbObjectError + 674, SRC, "Row1 col 'b' expected 2."
    End If
    If CLng(dataBack(r0 + 0, c0 + 1)) <> 1 Then
        Err.Raise vbObjectError + 675, SRC, "Row1 col 'a' expected 1."
    End If
    If Not IsEmpty(dataBack(r0 + 0, c0 + 2)) Then
        Err.Raise vbObjectError + 676, SRC, "Row1 col 'c' expected Empty."
    End If

    ' ---- row 2: b=Empty, a=Empty, c=3
    If Not IsEmpty(dataBack(r0 + 1, c0 + 0)) Then
        Err.Raise vbObjectError + 677, SRC, "Row2 col 'b' expected Empty."
    End If
    If Not IsEmpty(dataBack(r0 + 1, c0 + 1)) Then
        Err.Raise vbObjectError + 678, SRC, "Row2 col 'a' expected Empty."
    End If
    If CLng(dataBack(r0 + 1, c0 + 2)) <> 3 Then
        Err.Raise vbObjectError + 679, SRC, "Row2 col 'c' expected 3."
    End If

    ' ---- row 3: b=Empty, a=10, c=30
    If Not IsEmpty(dataBack(r0 + 2, c0 + 0)) Then
        Err.Raise vbObjectError + 680, SRC, "Row3 col 'b' expected Empty."
    End If
    If CLng(dataBack(r0 + 2, c0 + 1)) <> 10 Then
        Err.Raise vbObjectError + 681, SRC, "Row3 col 'a' expected 10."
    End If
    If CLng(dataBack(r0 + 2, c0 + 2)) <> 30 Then
        Err.Raise vbObjectError + 682, SRC, "Row3 col 'c' expected 30."
    End If

    Exit Sub

Fail:
    Dim n As Long: n = Err.Number
    Dim d As String: d = Err.Description
    Err.Clear
    Err.Raise n, SRC, d
End Sub


' =============================================================================
' TEST 39: Refresh (clearExisting=True) preserves formula columns
'
' Contract this test enforces:
'   - If a column already contains formulas, a refresh write must not wipe them.
'   - Data columns update, formula columns remain formulas after refresh.
'
' Notes:
'   - Uses FormulaR1C1 with computed offsets for stability.
' =============================================================================
Public Sub Test_Refresh_PreservesFormulaColumns_WithAsserts()

    Const SRC As String = "Test_Refresh_PreservesFormulaColumns_WithAsserts"

    On Error GoTo Fail

    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_Refresh_PreserveFormulas")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' -----------------------------
    ' PASS A: create base table with data cols + formula col
    ' -----------------------------
    Dim headersA As Variant
    headersA = Array("qty", "unit price", "total")  ' note the space

    Dim dataA As Variant
    ReDim dataA(1 To 2, 1 To 3)
    dataA(1, 1) = 2: dataA(1, 2) = 10: dataA(1, 3) = Empty
    dataA(2, 1) = 3: dataA(2, 2) = 20: dataA(2, 3) = Empty

    Excel_UpsertListObjectOnSheet ws, "tFormulas", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tFormulas")
    AssertNotNothing lo, "tFormulas should exist after PASS A"
    AssertRowCount lo, 2, "PASS A row count"
    AssertHeaderEquals lo, Array("qty", "unit price", "total"), "PASS A headers"

    ' Set formula in "total" using stable R1C1 offsets
    Dim qtyIdx As Long, unitIdx As Long, totalIdx As Long
    qtyIdx = lo.ListColumns("qty").Index
    unitIdx = lo.ListColumns("unit price").Index
    totalIdx = lo.ListColumns("total").Index

    Dim offQty As Long, offUnit As Long
    offQty = qtyIdx - totalIdx
    offUnit = unitIdx - totalIdx

    Dim fR1C1 As String
    fR1C1 = "=RC[" & CStr(offQty) & "]*RC[" & CStr(offUnit) & "]"

    lo.ListColumns("total").DataBodyRange.FormulaR1C1 = fR1C1

    ' Sanity: formulas exist and compute
    AssertBodyCellHasFormula lo, 1, "total", "PASS A total row1 has formula"
    AssertBodyCellHasFormula lo, 2, "total", "PASS A total row2 has formula"
    AssertEquals 20, lo.DataBodyRange.Cells(1, totalIdx).Value2, "PASS A total row1 value"
    AssertEquals 60, lo.DataBodyRange.Cells(2, totalIdx).Value2, "PASS A total row2 value"

    ' -----------------------------
    ' PASS B: refresh (clearExisting=True) with new qty/unit values
    ' -----------------------------
    Dim headersB As Variant
    headersB = Array("qty", "unit price")   ' incoming does not include "total"

    Dim dataB As Variant
    ReDim dataB(1 To 2, 1 To 2)
    dataB(1, 1) = 5: dataB(1, 2) = 7
    dataB(2, 1) = 4: dataB(2, 2) = 9

    ' clearExisting=True, addMissingColumns=True keeps existing schema, should not destroy formulas
    Excel_UpsertListObjectOnSheet ws, "tFormulas", ws.Range("A1"), headersB, dataB, True, True, False

    Set lo = GetTable(ws, "tFormulas")
    AssertNotNothing lo, "tFormulas should still exist after PASS B"
    AssertHeaderEquals lo, Array("qty", "unit price", "total"), "PASS B headers preserved"
    AssertRowCount lo, 2, "PASS B row count"

    ' Formula must still exist after refresh
    AssertBodyCellFormulaR1C1Equals lo, 1, "total", fR1C1, "PASS B total row1 formula preserved"
    AssertBodyCellFormulaR1C1Equals lo, 2, "total", fR1C1, "PASS B total row2 formula preserved"

    ' And it must recalc correctly with new inputs
    totalIdx = lo.ListColumns("total").Index
    AssertEquals 35, lo.DataBodyRange.Cells(1, totalIdx).Value2, "PASS B total row1 value"
    AssertEquals 36, lo.DataBodyRange.Cells(2, totalIdx).Value2, "PASS B total row2 value"

    Exit Sub

Fail:
    Dim n As Long: n = Err.Number
    Dim d As String: d = Err.Description
    Err.Clear
    Err.Raise n, SRC, d

End Sub


' =============================================================================
' TEST 40: Append mode preserves existing formula columns and auto-fills formulas
'
' Goal:
'   - Create a table with a formula column ("total") based on data columns.
'   - Append new JSON rows (clearExisting=False).
'   - Verify:
'       1) Existing formula cells remain formulas (not overwritten).
'       2) New appended rows get the formula filled down automatically.
'       3) Computed values are correct for both existing and appended rows.
'
' Notes:
'   - We avoid setting Formula on a single DataBodyRange cell (can throw 1004).
'   - We set the column's DataBodyRange formula in one shot AFTER seed rows exist.
'   - We assert formulas via .FormulaR1C1 to avoid localized structured-ref quirks.
' =============================================================================
Public Sub Test_Append_PreservesFormulaAndFillsDown_WithAsserts()

    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_Append_PreserveFormulas")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' -----------------------------
    ' PASS A: create base rows with [qty, unit_price]
    ' -----------------------------
    Dim jsonA As String
    jsonA = "[" & _
        "{""qty"":2,""unit_price"":5}," & _
        "{""qty"":3,""unit_price"":10}" & _
    "]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA

    ' Create table
    Excel_UpsertListObjectOnSheet ws, "tFormAppend", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tFormAppend")
    AssertNotNothing lo, "tFormAppend should exist after PASS A"
    AssertRowCount lo, 2, "PASS A rows"
    AssertHeaderEquals lo, Array("qty", "unit_price"), "PASS A headers"

    ' -----------------------------
    ' Add a formula column: "total" = qty * unit_price
    ' We add a column at the end and fill formulas for existing rows.
    ' -----------------------------
    lo.ListColumns.Add
    lo.ListColumns(lo.ListColumns.count).Name = "total"

    ' Fill the whole DataBodyRange of that column in one shot (safe)
    Dim colTotal As Long
    colTotal = lo.ListColumns("total").Index

    ' R1C1: total = RC[qty] * RC[unit_price]
    ' qty is 2 cols left of total, unit_price is 1 col left of total
    ' (because total was added at the end)
    lo.ListColumns("total").DataBodyRange.FormulaR1C1 = "=RC[-2]*RC[-1]"

    ' Assert computed values for existing rows
    AssertBodyCellEquals lo, 1, "total", 10, "PASS A row1 total"
    AssertBodyCellEquals lo, 2, "total", 30, "PASS A row2 total"

    ' Also assert formula exists for existing rows (R1C1)
    Dim f1 As String, f2 As String
    f1 = lo.DataBodyRange.Cells(1, colTotal).FormulaR1C1
    f2 = lo.DataBodyRange.Cells(2, colTotal).FormulaR1C1
    AssertEquals "=RC[-2]*RC[-1]", f1, "PASS A row1 total formula"
    AssertEquals "=RC[-2]*RC[-1]", f2, "PASS A row2 total formula"

    ' -----------------------------
    ' PASS B: append 2 more rows via JSON (clearExisting=False)
    ' -----------------------------
    Dim jsonB As String
    jsonB = "[" & _
        "{""qty"":4,""unit_price"":7}," & _
        "{""qty"":1,""unit_price"":9}" & _
    "]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonB, "$", headersB, dataB

    ' Append rows; schema union is fine (doesn't need to add columns here)
    Excel_UpsertListObjectOnSheet ws, "tFormAppend", ws.Range("A1"), headersB, dataB, False, True, False

    Set lo = GetTable(ws, "tFormAppend")
    AssertRowCount lo, 4, "PASS B rows should be 4 after append"
    AssertHeaderEquals lo, Array("qty", "unit_price", "total"), "PASS B headers preserved"

    ' Existing rows should still have formulas (not overwritten)
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(1, colTotal).FormulaR1C1, "PASS B row1 total formula preserved"
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(2, colTotal).FormulaR1C1, "PASS B row2 total formula preserved"

    ' New appended rows should have formula filled down automatically
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(3, colTotal).FormulaR1C1, "PASS B row3 total formula filled"
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(4, colTotal).FormulaR1C1, "PASS B row4 total formula filled"

    ' Assert computed totals for appended rows
    AssertBodyCellEquals lo, 3, "total", 28, "PASS B row3 total (4*7)"
    AssertBodyCellEquals lo, 4, "total", 9, "PASS B row4 total (1*9)"

End Sub


' =============================================================================
' TEST 41: Refresh (clearExisting=True) preserves formula columns and refills down
'
' Goal:
'   - Create a table with data columns + a formula column ("total").
'   - Refresh/replace data using clearExisting=True (like a "ListObject refresh").
'   - Verify:
'       1) The formula column still exists after refresh.
'       2) Formulas are present for all new rows (filled down).
'       3) Computed values match the refreshed data.
'
' Notes:
'   - We intentionally keep formula column OUT of incoming headers.
'   - This test asserts the engine does NOT delete or overwrite formula columns
'     when removeMissingColumns=False (default/typical refresh).
' =============================================================================
Public Sub Test_Refresh_PreservesFormulaColumn_AndFillsDown_WithAsserts()

    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_Refresh_PreserveFormulas")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' -----------------------------
    ' PASS A: seed with 2 rows
    ' -----------------------------
    Dim jsonA As String
    jsonA = "[" & _
        "{""qty"":2,""unit_price"":5}," & _
        "{""qty"":3,""unit_price"":10}" & _
    "]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA

    Excel_UpsertListObjectOnSheet ws, "tFormRefresh", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tFormRefresh")
    AssertNotNothing lo, "tFormRefresh should exist after PASS A"
    AssertRowCount lo, 2, "PASS A rows"
    AssertHeaderEquals lo, Array("qty", "unit_price"), "PASS A headers"

    ' -----------------------------
    ' Add formula column "total" and fill it for existing rows
    ' -----------------------------
    lo.ListColumns.Add
    lo.ListColumns(lo.ListColumns.count).Name = "total"

    Dim colTotal As Long
    colTotal = lo.ListColumns("total").Index

    lo.ListColumns("total").DataBodyRange.FormulaR1C1 = "=RC[-2]*RC[-1]"

    AssertBodyCellEquals lo, 1, "total", 10, "PASS A row1 total"
    AssertBodyCellEquals lo, 2, "total", 30, "PASS A row2 total"
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(1, colTotal).FormulaR1C1, "PASS A row1 total formula"
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(2, colTotal).FormulaR1C1, "PASS A row2 total formula"

    ' -----------------------------
    ' PASS B: refresh/replace (clearExisting=True) with NEW rows
    ' Incoming headers are still only [qty, unit_price]
    ' removeMissingColumns=False => formula column must remain
    ' -----------------------------
    Dim jsonB As String
    jsonB = "[" & _
        "{""qty"":4,""unit_price"":7}," & _
        "{""qty"":1,""unit_price"":9}," & _
        "{""qty"":6,""unit_price"":2}" & _
    "]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonB, "$", headersB, dataB

    Excel_UpsertListObjectOnSheet ws, "tFormRefresh", ws.Range("A1"), headersB, dataB, True, True, False

    Set lo = GetTable(ws, "tFormRefresh")
    AssertNotNothing lo, "tFormRefresh should still exist after PASS B"
    AssertRowCount lo, 3, "PASS B rows should be 3 after refresh"
    AssertHeaderEquals lo, Array("qty", "unit_price", "total"), "PASS B headers should preserve formula column"

    ' Re-acquire total column index (safe)
    colTotal = lo.ListColumns("total").Index

    ' Formulas must be filled down for all rows
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(1, colTotal).FormulaR1C1, "PASS B row1 total formula"
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(2, colTotal).FormulaR1C1, "PASS B row2 total formula"
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(3, colTotal).FormulaR1C1, "PASS B row3 total formula"

    ' Computed totals
    AssertBodyCellEquals lo, 1, "total", 28, "PASS B row1 total (4*7)"
    AssertBodyCellEquals lo, 2, "total", 9, "PASS B row2 total (1*9)"
    AssertBodyCellEquals lo, 3, "total", 12, "PASS B row3 total (6*2)"

End Sub


' =============================================================================
' TEST 42: Append mode fills formula down ONLY for newly appended rows
'
' Goal:
'   - Create table with data columns + formula column ("total").
'   - Append new rows using clearExisting=False.
'   - Verify:
'       1) Existing rows keep their formulas/values.
'       2) Newly appended rows receive the formula (filled down).
'       3) Computed values match expected results.
'
' Notes:
'   - Incoming headers exclude the formula column.
'   - This is the highest-value “append mode autofill” contract.
' =============================================================================
Public Sub Test_Append_FillsFormulaDown_ForNewRows_WithAsserts()

    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_Append_FillFormulas")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' -----------------------------
    ' PASS A: seed with 2 rows
    ' -----------------------------
    Dim jsonA As String
    jsonA = "[" & _
        "{""qty"":2,""unit_price"":5}," & _
        "{""qty"":3,""unit_price"":10}" & _
    "]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA

    Excel_UpsertListObjectOnSheet ws, "tFormAppend", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tFormAppend")
    AssertNotNothing lo, "tFormAppend should exist after PASS A"
    AssertRowCount lo, 2, "PASS A rows"
    AssertHeaderEquals lo, Array("qty", "unit_price"), "PASS A headers"

    ' Add formula column "total"
    lo.ListColumns.Add
    lo.ListColumns(lo.ListColumns.count).Name = "total"

    Dim colTotal As Long
    colTotal = lo.ListColumns("total").Index

    lo.ListColumns("total").DataBodyRange.FormulaR1C1 = "=RC[-2]*RC[-1]"

    AssertBodyCellEquals lo, 1, "total", 10, "PASS A row1 total"
    AssertBodyCellEquals lo, 2, "total", 30, "PASS A row2 total"
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(1, colTotal).FormulaR1C1, "PASS A row1 total formula"
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(2, colTotal).FormulaR1C1, "PASS A row2 total formula"

    ' -----------------------------
    ' PASS B: append 2 more rows
    ' -----------------------------
    Dim jsonB As String
    jsonB = "[" & _
        "{""qty"":4,""unit_price"":7}," & _
        "{""qty"":1,""unit_price"":9}" & _
    "]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonB, "$", headersB, dataB

    ' clearExisting := False => append
    Excel_UpsertListObjectOnSheet ws, "tFormAppend", ws.Range("A1"), headersB, dataB, False, True, False

    Set lo = GetTable(ws, "tFormAppend")
    AssertNotNothing lo, "tFormAppend should exist after PASS B"
    AssertRowCount lo, 4, "PASS B rows should be 4 after append"
    AssertHeaderEquals lo, Array("qty", "unit_price", "total"), "PASS B headers preserve formula column"

    colTotal = lo.ListColumns("total").Index

    ' Existing rows should remain correct
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(1, colTotal).FormulaR1C1, "PASS B row1 formula unchanged"
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(2, colTotal).FormulaR1C1, "PASS B row2 formula unchanged"
    AssertBodyCellEquals lo, 1, "total", 10, "PASS B row1 total unchanged"
    AssertBodyCellEquals lo, 2, "total", 30, "PASS B row2 total unchanged"

    ' New rows must have formula filled down
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(3, colTotal).FormulaR1C1, "PASS B row3 total formula"
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(4, colTotal).FormulaR1C1, "PASS B row4 total formula"

    ' Computed totals
    AssertBodyCellEquals lo, 3, "total", 28, "PASS B row3 total (4*7)"
    AssertBodyCellEquals lo, 4, "total", 9, "PASS B row4 total (1*9)"

End Sub


' =============================================================================
' TEST 43: Refresh (clearExisting=True) preserves formulas in existing formula columns
'
' Goal:
'   - Create table with data columns + formula column ("total").
'   - Refresh (clearExisting=True) with new data rows.
'   - Verify:
'       1) Formula column remains present (schema union behavior).
'       2) Formulas are present for ALL refreshed rows (filled down).
'       3) Computed values match expected results.
'
' Notes:
'   - Incoming headers exclude the formula column.
'   - This validates the "preserve formulas on refresh" contract.
' =============================================================================
Public Sub Test_Refresh_PreservesFormulaColumns_AndFillsDown_WithAsserts()

    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_Refresh_PreserveFormulas")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' -----------------------------
    ' PASS A: seed with 2 rows
    ' -----------------------------
    Dim jsonA As String
    jsonA = "[" & _
        "{""qty"":2,""unit_price"":5}," & _
        "{""qty"":3,""unit_price"":10}" & _
    "]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA

    Excel_UpsertListObjectOnSheet ws, "tFormRefresh", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tFormRefresh")
    AssertNotNothing lo, "tFormRefresh should exist after PASS A"
    AssertRowCount lo, 2, "PASS A rows"
    AssertHeaderEquals lo, Array("qty", "unit_price"), "PASS A headers"

    ' Add formula column "total"
    lo.ListColumns.Add
    lo.ListColumns(lo.ListColumns.count).Name = "total"

    Dim colTotal As Long
    colTotal = lo.ListColumns("total").Index
    lo.ListColumns("total").DataBodyRange.FormulaR1C1 = "=RC[-2]*RC[-1]"

    AssertBodyCellEquals lo, 1, "total", 10, "PASS A row1 total"
    AssertBodyCellEquals lo, 2, "total", 30, "PASS A row2 total"

    ' -----------------------------
    ' PASS B: refresh with 3 new rows (clearExisting=True)
    ' -----------------------------
    Dim jsonB As String
    jsonB = "[" & _
        "{""qty"":4,""unit_price"":7}," & _
        "{""qty"":1,""unit_price"":9}," & _
        "{""qty"":6,""unit_price"":2}" & _
    "]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonB, "$", headersB, dataB

    ' clearExisting := True => refresh/replace body
    ' addMissingColumns := True => keep existing columns (including formula col)
    Excel_UpsertListObjectOnSheet ws, "tFormRefresh", ws.Range("A1"), headersB, dataB, True, True, False

    Set lo = GetTable(ws, "tFormRefresh")
    AssertNotNothing lo, "tFormRefresh should exist after PASS B"

    ' Schema should preserve formula column
    AssertHeaderEquals lo, Array("qty", "unit_price", "total"), "PASS B headers preserve formula column"
    AssertRowCount lo, 3, "PASS B rows should be 3"

    colTotal = lo.ListColumns("total").Index

    ' Formula should exist for all refreshed rows
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(1, colTotal).FormulaR1C1, "PASS B row1 total formula"
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(2, colTotal).FormulaR1C1, "PASS B row2 total formula"
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(3, colTotal).FormulaR1C1, "PASS B row3 total formula"

    ' Computed totals
    AssertBodyCellEquals lo, 1, "total", 28, "PASS B row1 total (4*7)"
    AssertBodyCellEquals lo, 2, "total", 9, "PASS B row2 total (1*9)"
    AssertBodyCellEquals lo, 3, "total", 12, "PASS B row3 total (6*2)"

End Sub


' =============================================================================
' TEST 44: Append mode auto-fills formulas down for newly appended rows
'
' Goal:
'   - Create table with data columns + formula column ("total").
'   - Append new rows (clearExisting=False).
'   - Verify:
'       1) Formula exists for appended rows (filled down).
'       2) Computed values are correct for appended rows.
'
' Notes:
'   - Incoming headers exclude the formula column.
'   - This validates the "append autofill formulas" contract.
' =============================================================================
Public Sub Test_Append_AutoFillFormulaColumns_ForNewRows_WithAsserts()

    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_Append_AutoFillFormulas")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' -----------------------------
    ' PASS A: seed with 2 rows
    ' -----------------------------
    Dim jsonA As String
    jsonA = "[" & _
        "{""qty"":2,""unit_price"":5}," & _
        "{""qty"":3,""unit_price"":10}" & _
    "]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA

    Excel_UpsertListObjectOnSheet ws, "tFormAppend", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tFormAppend")
    AssertNotNothing lo, "tFormAppend should exist after PASS A"
    AssertRowCount lo, 2, "PASS A rows"
    AssertHeaderEquals lo, Array("qty", "unit_price"), "PASS A headers"

    ' Add formula column "total"
    lo.ListColumns.Add
    lo.ListColumns(lo.ListColumns.count).Name = "total"

    Dim colTotal As Long
    colTotal = lo.ListColumns("total").Index

    ' Seed formula for existing rows
    lo.ListColumns("total").DataBodyRange.FormulaR1C1 = "=RC[-2]*RC[-1]"

    AssertBodyCellEquals lo, 1, "total", 10, "PASS A row1 total"
    AssertBodyCellEquals lo, 2, "total", 30, "PASS A row2 total"

    ' -----------------------------
    ' PASS B: append 2 rows
    ' -----------------------------
    Dim jsonB As String
    jsonB = "[" & _
        "{""qty"":4,""unit_price"":7}," & _
        "{""qty"":1,""unit_price"":9}" & _
    "]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonB, "$", headersB, dataB

    ' clearExisting := False => append
    ' addMissingColumns := True => schema can grow (not needed here), but preserves formula col
    Excel_UpsertListObjectOnSheet ws, "tFormAppend", ws.Range("A1"), headersB, dataB, False, True, False

    Set lo = GetTable(ws, "tFormAppend")
    AssertNotNothing lo, "tFormAppend should exist after PASS B"
    AssertHeaderEquals lo, Array("qty", "unit_price", "total"), "PASS B headers"
    AssertRowCount lo, 4, "PASS B total rows should be 4"

    colTotal = lo.ListColumns("total").Index

    ' Formulas should exist for appended rows (rows 3 and 4)
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(3, colTotal).FormulaR1C1, "PASS B row3 total formula"
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(4, colTotal).FormulaR1C1, "PASS B row4 total formula"

    ' Computed totals for appended rows
    AssertBodyCellEquals lo, 3, "total", 28, "PASS B row3 total (4*7)"
    AssertBodyCellEquals lo, 4, "total", 9, "PASS B row4 total (1*9)"

End Sub


' =============================================================================
' TEST 45: Refresh (clearExisting=True) preserves formula columns + refills down
'
' Goal:
'   - Create table with data columns + formula column.
'   - Refresh (clearExisting=True) with a DIFFERENT row count.
'   - Verify:
'       1) Formula column is preserved (not overwritten by incoming headers).
'       2) Formula is filled down for ALL refreshed rows.
'       3) Computed values are correct after refresh.
'
' Notes:
'   - Incoming headers exclude the formula column.
'   - This validates the "refresh preserves formulas" contract.
' =============================================================================
Public Sub Test_Refresh_PreservesFormulaColumns_And_FillsDown_WithAsserts()

    Dim ws As Worksheet
    Set ws = EnsureTestSheet("zTest_Refresh_PreserveFormulas")
    ResetSheetButKeepTableTestSafe ws

    Dim lo As ListObject

    ' -----------------------------
    ' PASS A: seed with 2 rows
    ' -----------------------------
    Dim jsonA As String
    jsonA = "[" & _
        "{""qty"":2,""unit_price"":5}," & _
        "{""qty"":3,""unit_price"":10}" & _
    "]"

    Dim headersA As Variant, dataA As Variant
    Build2DFromJsonRoot jsonA, "$", headersA, dataA

    Excel_UpsertListObjectOnSheet ws, "tFormRefresh", ws.Range("A1"), headersA, dataA, True, True, False

    Set lo = GetTable(ws, "tFormRefresh")
    AssertNotNothing lo, "tFormRefresh should exist after PASS A"
    AssertRowCount lo, 2, "PASS A rows"
    AssertHeaderEquals lo, Array("qty", "unit_price"), "PASS A headers"

    ' Add formula column "total"
    lo.ListColumns.Add
    lo.ListColumns(lo.ListColumns.count).Name = "total"

    Dim colTotal As Long
    colTotal = lo.ListColumns("total").Index

    ' Seed formula for existing rows
    lo.ListColumns("total").DataBodyRange.FormulaR1C1 = "=RC[-2]*RC[-1]"

    AssertBodyCellEquals lo, 1, "total", 10, "PASS A row1 total"
    AssertBodyCellEquals lo, 2, "total", 30, "PASS A row2 total"

    ' -----------------------------
    ' PASS B: refresh with 3 rows (clearExisting=True)
    ' -----------------------------
    Dim jsonB As String
    jsonB = "[" & _
        "{""qty"":4,""unit_price"":7}," & _
        "{""qty"":1,""unit_price"":9}," & _
        "{""qty"":6,""unit_price"":2}" & _
    "]"

    Dim headersB As Variant, dataB As Variant
    Build2DFromJsonRoot jsonB, "$", headersB, dataB

    ' clearExisting := True => replace body
    ' addMissingColumns := True => keep existing schema (including formula col)
    ' removeMissingColumns := False
    Excel_UpsertListObjectOnSheet ws, "tFormRefresh", ws.Range("A1"), headersB, dataB, True, True, False

    Set lo = GetTable(ws, "tFormRefresh")
    AssertNotNothing lo, "tFormRefresh should exist after PASS B"
    AssertHeaderEquals lo, Array("qty", "unit_price", "total"), "PASS B headers preserve formula column"
    AssertRowCount lo, 3, "PASS B rows should be 3"

    colTotal = lo.ListColumns("total").Index

    ' Formula should be present for all refreshed rows
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(1, colTotal).FormulaR1C1, "PASS B row1 total formula"
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(2, colTotal).FormulaR1C1, "PASS B row2 total formula"
    AssertEquals "=RC[-2]*RC[-1]", lo.DataBodyRange.Cells(3, colTotal).FormulaR1C1, "PASS B row3 total formula"

    ' Computed totals after refresh
    AssertBodyCellEquals lo, 1, "total", 28, "PASS B row1 total (4*7)"
    AssertBodyCellEquals lo, 2, "total", 9, "PASS B row2 total (1*9)"
    AssertBodyCellEquals lo, 3, "total", 12, "PASS B row3 total (6*2)"

End Sub
