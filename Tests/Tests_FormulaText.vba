Attribute VB_Name = "Tests_FormulaText"
Option Explicit

' =============================================================================
' formulaStringsAsText tests
'
' A Range assignment treats each string like typed input: "=..." becomes a
' formula and a leading apostrophe is consumed. By default the upsert
' functions keep that behavior (formula injection is a feature). With
' formulaStringsAsText:=True, payload values and headers are written as
' the exact text, so nothing from an untrusted payload is evaluated.
'
' Patterns follow Tests_EmptyUpsert: fresh sheet per test, cleanup that
' re-raises so a failed assert surfaces as an error, not a modal dialog.
' =============================================================================

Public Sub RunAll_FormulaTextTests_StopOnFail()
    On Error GoTo Fail

    Test_Default_EqualsStringIsFormula
    Test_TextMode_EqualsStringIsText
    Test_TextMode_ApostropheStringsKept
    Test_TextMode_OtherValuesUnchanged
    Test_TextMode_HeaderKeyIsText
    Test_TextMode_AppendIsText
    Test_TextMode_CsvSource
    Test_TextMode_OnSheet2D
    Test_TextMode_KeepsUserFormulaColumn
    Test_TextMode_RoundTrip

    MsgBox "All formula-text tests passed.", vbInformation
    Exit Sub

Fail:
    Err.Raise vbObjectError + 850, "mFormulaTextTests", _
        "Formula-text test run failed. Err " & Err.Number & ": " & Err.Description
End Sub


' =============================================================================
' ASSERTS AND HELPERS
' =============================================================================

Private Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal message As String)
    If VarType(expected) <> VarType(actual) Or expected <> actual Then
        Err.Raise vbObjectError + 851, "mFormulaTextTests", _
            message & " expected=" & CStr(expected) & " (" & TypeName(expected) & ")" & _
            " actual=" & CStr(actual) & " (" & TypeName(actual) & ")"
    End If
End Sub

' The cell holds exactly this text and no formula.
Private Sub AssertText(ByVal cell As Range, ByVal expected As String, ByVal message As String)
    AssertEquals False, CBool(cell.HasFormula), message & " (HasFormula)"
    AssertEquals expected, cell.Value2, message & " (Value2)"
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

' Upsert jsonText at "$" with formulaStringsAsText:=asText.
Private Function Upsert(ByVal ws As Worksheet, ByVal tableName As String, ByVal jsonText As String, _
    ByVal asText As Boolean, Optional ByVal clearExisting As Boolean = True) As ListObject
    Set Upsert = Excel_UpsertListObjectFromJsonAtRoot(ws, tableName, ws.Range("A1"), jsonText, _
        clearExisting:=clearExisting, formulaStringsAsText:=asText)
End Function


' =============================================================================
' TESTS
' =============================================================================

Private Sub Test_Default_EqualsStringIsFormula()
    ' Pins the default: formula injection from a payload still works.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Upsert(ws, "T_FT1", "[{""f"":""=1+1""}]", False)

    AssertEquals True, CBool(lo.DataBodyRange.Cells(1, 1).HasFormula), "default writes a formula"
    AssertEquals 2#, lo.DataBodyRange.Cells(1, 1).Value2, "default formula value"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_TextMode_EqualsStringIsText()
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Upsert(ws, "T_FT2", "[{""f"":""=1+1"",""w"":""=WEBSERVICE(A1)""}]", True)

    AssertText lo.DataBodyRange.Cells(1, 1), "=1+1", "=1+1 kept as text"
    AssertText lo.DataBodyRange.Cells(1, 2), "=WEBSERVICE(A1)", "WEBSERVICE kept as text"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_TextMode_ApostropheStringsKept()
    ' Without text mode a leading apostrophe is consumed as the text marker.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Upsert(ws, "T_FT3", "[{""a"":""'abc"",""b"":""''x"",""c"":""'""}]", True)

    AssertText lo.DataBodyRange.Cells(1, 1), "'abc", "one apostrophe kept"
    AssertText lo.DataBodyRange.Cells(1, 2), "''x", "two apostrophes kept"
    AssertText lo.DataBodyRange.Cells(1, 3), "'", "lone apostrophe kept"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_TextMode_OtherValuesUnchanged()
    ' Only strings that begin with "=" or "'" change. Numbers, booleans,
    ' empty strings, and strings with "=" later on are written as before.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Upsert(ws, "T_FT4", _
        "[{""n"":5,""b"":true,""s"":""a=b"",""p"":""+2+2"",""e"":"""",""u"":""x""}]", True)

    AssertEquals 5#, lo.DataBodyRange.Cells(1, 1).Value2, "number stays numeric"
    AssertEquals True, lo.DataBodyRange.Cells(1, 2).Value2, "boolean stays boolean"
    AssertText lo.DataBodyRange.Cells(1, 3), "a=b", "inner = untouched"
    AssertText lo.DataBodyRange.Cells(1, 4), "+2+2", "leading + untouched"
    AssertEquals "", lo.DataBodyRange.Cells(1, 5).Text, "empty string stays empty"
    AssertText lo.DataBodyRange.Cells(1, 6), "x", "plain text untouched"
    AssertEquals "", lo.DataBodyRange.Cells(1, 3).PrefixCharacter, "no prefix on plain text"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_TextMode_HeaderKeyIsText()
    ' A key beginning with "=" becomes a header; without text mode Excel
    ' evaluates it once and keeps the result as the header name.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Upsert(ws, "T_FT5", "[{""=1+1"":1,""'k"":2}]", True)

    AssertText lo.HeaderRowRange.Cells(1, 1), "=1+1", "= header kept as text"
    AssertText lo.HeaderRowRange.Cells(1, 2), "'k", "apostrophe header kept"

    ' A second refresh matches the existing headers instead of adding columns.
    Set lo = Upsert(ws, "T_FT5", "[{""=1+1"":3,""'k"":4}]", True)
    AssertEquals 2&, CLng(lo.ListColumns.count), "refresh matches escaped headers"
    AssertEquals 3#, lo.DataBodyRange.Cells(1, 1).Value2, "refresh value"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_TextMode_AppendIsText()
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Upsert(ws, "T_FT6", "[{""f"":""plain""}]", True)
    Set lo = Upsert(ws, "T_FT6", "[{""f"":""=1+1""},{""f"":""=2+2"",""g"":""=3+3""}]", True, False)

    AssertEquals 3&, CLng(lo.ListRows.count), "append row count"
    AssertText lo.DataBodyRange.Cells(2, 1), "=1+1", "appended row 1"
    AssertText lo.DataBodyRange.Cells(3, 1), "=2+2", "appended row 2"
    AssertText lo.DataBodyRange.Cells(3, 2), "=3+3", "appended new column"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_TextMode_CsvSource()
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromSource(ws, "T_FT7", ws.Range("A1"), _
        "a,b" & vbLf & "=1+1,ok", ExcelSourceFormat_CSV, formulaStringsAsText:=True)

    AssertText lo.DataBodyRange.Cells(1, 1), "=1+1", "CSV value kept as text"
    AssertText lo.DataBodyRange.Cells(1, 2), "ok", "CSV plain value"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_TextMode_OnSheet2D()
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim data(1 To 2, 1 To 1) As Variant
    data(1, 1) = "=1+1"
    data(2, 1) = "'q"

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectOnSheet(ws, "T_FT8", ws.Range("A1"), Array("=h"), data, _
        formulaStringsAsText:=True)

    AssertText lo.HeaderRowRange.Cells(1, 1), "=h", "2D header"
    AssertText lo.DataBodyRange.Cells(1, 1), "=1+1", "2D row 1"
    AssertText lo.DataBodyRange.Cells(2, 1), "'q", "2D row 2"
    AssertEquals "=1+1", data(1, 1), "caller's array not modified"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_TextMode_KeepsUserFormulaColumn()
    ' Text mode applies to payload values only. A formula column the user
    ' added to the table is still preserved and refilled.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Upsert(ws, "T_FT9", "[{""qty"":2,""total"":0}]", True)
    lo.ListColumns("total").DataBodyRange.FormulaR1C1 = "=RC[-1]*10"

    Set lo = Upsert(ws, "T_FT9", "[{""qty"":3,""total"":""=1+1""},{""qty"":4,""total"":0}]", True)

    AssertEquals True, CBool(lo.DataBodyRange.Cells(1, 2).HasFormula), "user formula kept row 1"
    AssertEquals 30#, lo.DataBodyRange.Cells(1, 2).Value2, "user formula value row 1"
    AssertEquals 40#, lo.DataBodyRange.Cells(2, 2).Value2, "user formula value row 2"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_TextMode_RoundTrip()
    ' Text written in text mode exports back to the original strings.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim src As String
    src = "[{""=k"":""=1+1"",""a"":""'abc"",""n"":5}]"

    Dim lo As ListObject
    Set lo = Upsert(ws, "T_FT10", src, True)

    AssertEquals src, Excel_ListObjectToJson(lo), "round trip"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub
