Attribute VB_Name = "Tests_UpsertReturns"
Option Explicit

' =============================================================================
' Upsert-returns-ListObject tests
'
' The three upsert entry points return the ListObject they created or updated:
'   Excel_UpsertListObjectOnSheet
'   Excel_UpsertListObjectFromJsonAtRoot
'   Excel_UpsertListObjectFromSource
'
' Validates that the returned reference is the same table that lives on the
' sheet, that it is live and usable, that the empty-array path still returns a
' table, and that statement-style calls (discarding the return) still work.
'
' "Same table" is checked by range address, not object identity: Excel can hand
' back distinct COM wrappers for one table, so `lo Is ws.ListObjects(name)` is
' not reliable. Each test uses a distinct table name (ListObject names are
' workbook-global), its own fresh worksheet, and a handler that drops the sheet
' and re-raises so a failure reports cleanly instead of leaking a sheet.
' =============================================================================

Public Sub RunAll_UpsertReturnsTests_StopOnFail()
    On Error GoTo Fail

    Test_OnSheet_ReturnsTable
    Test_FromJsonAtRoot_ReturnsTable
    Test_FromSource_ReturnsTable
    Test_ReturnedReferenceIsUsable
    Test_EmptyArray_ReturnsTable
    Test_StatementCallStillWorks

    MsgBox "All upsert-returns-ListObject tests passed.", vbInformation
    Exit Sub

Fail:
    Err.Raise vbObjectError + 800, "mUpsertReturnsTests", _
        "Upsert-returns test run failed. Err " & Err.Number & ": " & Err.Description
End Sub


' =============================================================================
' ASSERTS AND HELPERS
' =============================================================================

Private Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal message As String)
    If expected <> actual Then
        Err.Raise vbObjectError + 801, "mUpsertReturnsTests", _
            message & " expected=" & CStr(expected) & " actual=" & CStr(actual)
    End If
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 802, "mUpsertReturnsTests", message
End Sub

' The returned reference points at the same table as a fresh lookup. Compared by
' range address because Excel can return distinct COM wrappers for one table.
Private Sub AssertSameTable(ByVal ws As Worksheet, ByVal tableName As String, ByVal lo As ListObject, ByVal label As String)
    AssertTrue Not lo Is Nothing, label & ": returned a ListObject"
    AssertEquals tableName, lo.name, label & ": returned table name"
    AssertEquals ws.ListObjects(tableName).Range.Address, lo.Range.Address, _
        label & ": returned reference is the table on the sheet"
End Sub

Private Function FreshSheet() As Worksheet
    Set FreshSheet = ThisWorkbook.Worksheets.Add
End Function

Private Sub DropSheet(ByVal ws As Worksheet)
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
End Sub

' Drop the sheet (best effort) and re-raise the current error so the run fails
' cleanly rather than leaving an untrapped error to become a modal dialog.
Private Sub CleanupAndReraise(ByVal ws As Worksheet)
    Dim en As Long, ed As String, es As String
    en = Err.Number: ed = Err.Description: es = Err.Source
    On Error Resume Next
    DropSheet ws
    On Error GoTo 0
    Err.Raise en, es, ed
End Sub


' =============================================================================
' TESTS
' =============================================================================

Private Sub Test_OnSheet_ReturnsTable()
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim d(1 To 2, 1 To 2) As Variant
    d(1, 1) = 1: d(1, 2) = "Alice"
    d(2, 1) = 2: d(2, 2) = "Bob"

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectOnSheet(ws, "T_OnSheet", ws.Range("A1"), Array("id", "name"), d)

    AssertSameTable ws, "T_OnSheet", lo, "OnSheet"
    AssertEquals 2, lo.ListRows.count, "returned row count"
    AssertEquals 2, lo.ListColumns.count, "returned column count"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_FromJsonAtRoot_ReturnsTable()
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_Json", ws.Range("A1"), _
        "[{""id"":1,""name"":""Alice""},{""id"":2,""name"":""Bob""}]", "$")

    AssertSameTable ws, "T_Json", lo, "FromJsonAtRoot"
    AssertEquals 2, lo.ListRows.count, "returned row count"
    AssertEquals 2, lo.ListColumns.count, "returned column count"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_FromSource_ReturnsTable()
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromSource(ws, "T_Source", ws.Range("A1"), _
        "[{""id"":1,""name"":""X""}]", ExcelSourceFormat_JSON)

    AssertSameTable ws, "T_Source", lo, "FromSource"
    AssertEquals 1, lo.ListRows.count, "returned row count"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_ReturnedReferenceIsUsable()
    ' The point of the change: keep working with the table without a second
    ' ws.ListObjects(name) lookup.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_Use", ws.Range("A1"), _
        "[{""id"":1,""name"":""Alice""}]", "$")

    AssertEquals "id", lo.ListColumns(1).name, "first column name via the returned reference"
    AssertEquals "name", lo.ListColumns(2).name, "second column name via the returned reference"
    lo.TableStyle = "TableStyleMedium2"
    AssertEquals "TableStyleMedium2", lo.TableStyle, "style set through the returned reference sticks"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_EmptyArray_ReturnsTable()
    ' The rowCount = 0 branch also returns the ListObject.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_Empty", ws.Range("A1"), "[]", "$")

    AssertTrue Not lo Is Nothing, "empty array still returns a ListObject"
    AssertEquals "T_Empty", lo.name, "empty-array table name"
    AssertEquals 0, lo.ListRows.count, "empty-array table has no data rows"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_StatementCallStillWorks()
    ' Backward compatibility: a Function may be called as a statement, so
    ' existing code that ignored the (former Sub) return keeps working.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Excel_UpsertListObjectFromJsonAtRoot ws, "T_Stmt", ws.Range("A1"), "[{""id"":9}]", "$"

    Dim lo As ListObject
    Set lo = ws.ListObjects("T_Stmt")
    AssertEquals 1, lo.ListRows.count, "statement-style call created the table"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub
