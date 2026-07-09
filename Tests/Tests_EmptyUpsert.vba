Attribute VB_Name = "Tests_EmptyUpsert"
Option Explicit

' =============================================================================
' Empty-result upsert tests
'
' An empty JSON array (or a null table root) must not disturb an existing
' table's layout:
'   - refresh (default): rows cleared, schema untouched, no "value" column
'   - append (clearExisting:=False): complete no-op
'   - removeMissingColumns:=True: schema kept, rows cleared (historical)
'
' The single-column "value" placeholder appears only when there is no prior
' table (a ListObject cannot have zero columns), and the first real result
' replaces the placeholder schema instead of merging with it. A "value"
' table that HOLDS rows is user data, not a placeholder, and merges
' normally.
'
' Patterns follow Tests_UpsertReturns: fresh sheet per test, cleanup that
' re-raises so a failed assert surfaces as an error, not a modal dialog.
' =============================================================================

Public Sub RunAll_EmptyUpsertTests_StopOnFail()
    On Error GoTo Fail

    Test_Empty_NewTable_CreatesPlaceholder
    Test_Empty_ExistingRefresh_KeepsSchema
    Test_Empty_ExistingAppend_NoOp
    Test_Empty_RemoveMissing_KeepsSchema
    Test_Empty_NullRoot_KeepsSchema
    Test_Placeholder_ReplacedByRealData
    Test_ValueTableWithRows_MergesNormally

    MsgBox "All empty-upsert tests passed.", vbInformation
    Exit Sub

Fail:
    Err.Raise vbObjectError + 840, "mEmptyUpsertTests", _
        "Empty-upsert test run failed. Err " & Err.Number & ": " & Err.Description
End Sub


' =============================================================================
' ASSERTS AND HELPERS
' =============================================================================

Private Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal message As String)
    If expected <> actual Then
        Err.Raise vbObjectError + 841, "mEmptyUpsertTests", _
            message & " expected=" & CStr(expected) & " actual=" & CStr(actual)
    End If
End Sub

' Assert the table's columns are exactly the listed names, in order.
Private Sub AssertColumns(ByVal lo As ListObject, ByVal names As Variant, ByVal message As String)
    AssertEquals UBound(names) - LBound(names) + 1, lo.ListColumns.count, message & " (column count)"

    Dim i As Long
    For i = LBound(names) To UBound(names)
        AssertEquals names(i), lo.ListColumns(i - LBound(names) + 1).name, _
            message & " (column " & (i - LBound(names) + 1) & ")"
    Next i
End Sub

Private Function FreshSheet() As Worksheet
    Set FreshSheet = ThisWorkbook.Worksheets.Add
End Function

Private Sub DropSheet(ByVal ws As Worksheet)
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
End Sub

Private Sub CleanupAndReraise(ByVal ws As Worksheet)
    Dim en As Long, ed As String, es As String
    en = Err.Number: ed = Err.Description: es = Err.Source
    On Error Resume Next
    DropSheet ws
    On Error GoTo 0
    Err.Raise en, es, ed
End Sub

' Two rows, columns id + name.
Private Function SeedTable(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    Set SeedTable = Excel_UpsertListObjectFromJsonAtRoot(ws, tableName, ws.Range("A1"), _
        "[{""id"":1,""name"":""Alice""},{""id"":2,""name"":""Bob""}]")
End Function


' =============================================================================
' TESTS
' =============================================================================

Private Sub Test_Empty_NewTable_CreatesPlaceholder()
    ' No prior table: an empty array still returns a table, and a ListObject
    ' cannot have zero columns, so it gets the single "value" placeholder.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_EU1", ws.Range("A1"), "[]")

    AssertColumns lo, Array("value"), "placeholder table"
    AssertEquals 0, lo.ListRows.count, "placeholder has no rows"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Empty_ExistingRefresh_KeepsSchema()
    ' The reported bug: refreshing an existing table with an empty array
    ' added a "value" column. It must clear the rows and leave the schema
    ' alone.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    SeedTable ws, "T_EU2"

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_EU2", ws.Range("A1"), "[]")

    AssertColumns lo, Array("id", "name"), "schema untouched by empty refresh"
    AssertEquals 0, lo.ListRows.count, "rows cleared by empty refresh"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Empty_ExistingAppend_NoOp()
    ' Appending an empty array is a complete no-op: schema and rows both
    ' stay.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    SeedTable ws, "T_EU3"

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_EU3", ws.Range("A1"), "[]", _
        clearExisting:=False)

    AssertColumns lo, Array("id", "name"), "schema untouched by empty append"
    AssertEquals 2, lo.ListRows.count, "rows untouched by empty append"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Empty_RemoveMissing_KeepsSchema()
    ' The historical special case: empty result with removeMissingColumns
    ' keeps the schema and clears the rows.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    SeedTable ws, "T_EU4"

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_EU4", ws.Range("A1"), "[]", _
        removeMissingColumns:=True)

    AssertColumns lo, Array("id", "name"), "schema kept with removeMissingColumns"
    AssertEquals 0, lo.ListRows.count, "rows cleared with removeMissingColumns"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Empty_NullRoot_KeepsSchema()
    ' A null table root resolves like an empty result and must respect the
    ' same rule against an existing table.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    SeedTable ws, "T_EU5"

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_EU5", ws.Range("A1"), _
        "{""data"":{""items"":null}}", "$.data.items")

    AssertColumns lo, Array("id", "name"), "schema untouched by null root"
    AssertEquals 0, lo.ListRows.count, "rows cleared by null root"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Placeholder_ReplacedByRealData()
    ' First sync returns nothing (placeholder table), second sync returns
    ' rows: the placeholder schema is replaced, not merged, so no "value"
    ' column lingers.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Excel_UpsertListObjectFromJsonAtRoot ws, "T_EU6", ws.Range("A1"), "[]"

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_EU6", ws.Range("A1"), _
        "[{""id"":1,""name"":""Alice""}]")

    AssertColumns lo, Array("id", "name"), "placeholder replaced by real schema"
    AssertEquals 1, lo.ListRows.count, "real row landed"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_ValueTableWithRows_MergesNormally()
    ' A "value" column holding rows is user data, not a placeholder: new
    ' headers merge alongside it under the default add-only rules.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Excel_UpsertListObjectFromJsonAtRoot ws, "T_EU7", ws.Range("A1"), "[{""value"":10}]"

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_EU7", ws.Range("A1"), "[{""id"":9}]")

    AssertColumns lo, Array("value", "id"), "value column with rows is kept"
    AssertEquals 1, lo.ListRows.count, "refresh replaced the rows"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub
