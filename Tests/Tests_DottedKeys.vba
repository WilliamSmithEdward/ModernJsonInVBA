Attribute VB_Name = "Tests_DottedKeys"
Option Explicit

' =============================================================================
' Dotted-key escaping tests
'
' A JSON key containing a literal dot must stay distinguishable from real
' nesting in both directions:
'   - import: {"a.b":1} gets header a\.b while {"a":{"b":2}} gets a.b
'   - export: a\.b unescapes back to the "a.b" key; a.b rebuilds nesting
'   - backslashes in keys escape the same way (c\d -> header c\\d)
'   - tableRoot and the path resolvers accept the same escapes: "$.a\.b"
'     addresses the literal key "a.b" (streamed and model paths), while an
'     unescaped dot keeps splitting levels (plain "$.a.b" raises 1160 when
'     only the literal key exists)
'
' Patterns follow Tests_UpsertReturns: fresh sheet per test, cleanup that
' re-raises so a failed assert surfaces as an error, not a modal dialog.
' =============================================================================

Public Sub RunAll_DottedKeyTests_StopOnFail()
    On Error GoTo Fail

    Test_Dotted_HeadersDistinct
    Test_Dotted_RoundTrip
    Test_Dotted_NestedDottedKey
    Test_Dotted_TableRootCannotAddress
    Test_Dotted_TableRootEscapedDot
    Test_Dotted_TableRootEscapeWithBracket
    Test_Dotted_ResolvePathEscaped

    MsgBox "All dotted-key tests passed.", vbInformation
    Exit Sub

Fail:
    Err.Raise vbObjectError + 830, "mDottedKeyTests", _
        "Dotted-key test run failed. Err " & Err.Number & ": " & Err.Description
End Sub


' =============================================================================
' ASSERTS AND HELPERS
' =============================================================================

Private Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal message As String)
    If expected <> actual Then
        Err.Raise vbObjectError + 831, "mDottedKeyTests", _
            message & " expected=" & CStr(expected) & " actual=" & CStr(actual)
    End If
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


' =============================================================================
' TESTS
' =============================================================================

Private Sub Test_Dotted_HeadersDistinct()
    ' One row holds the literal key "a.b", real nesting a:{b:}, and a key
    ' with a backslash. Three distinct columns must come out.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_DK1", ws.Range("A1"), _
        "[{""a.b"":1,""a"":{""b"":2},""c\\d"":3}]")

    AssertEquals 3, lo.ListColumns.count, "three distinct columns"
    AssertEquals "a\.b", lo.ListColumns(1).name, "literal dotted key escapes its dot"
    AssertEquals "a.b", lo.ListColumns(2).name, "real nesting uses the plain dot"
    AssertEquals "c\\d", lo.ListColumns(3).name, "backslash in a key doubles"
    AssertEquals 1, lo.DataBodyRange.Cells(1, 1).Value2, "literal-key cell"
    AssertEquals 2, lo.DataBodyRange.Cells(1, 2).Value2, "nested cell"
    AssertEquals 3, lo.DataBodyRange.Cells(1, 3).Value2, "backslash-key cell"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Dotted_RoundTrip()
    ' Export must rebuild the original document: the escaped header comes
    ' back as the literal "a.b" key, the plain dotted header as nesting.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim json As String
    json = "[{""a.b"":1,""a"":{""b"":2},""c\\d"":3}]"

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_DK2", ws.Range("A1"), json)

    AssertEquals Json_Stringify(Json_Parse(json)), Excel_ListObjectToJson(lo), _
        "dotted keys round-trip byte for byte"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Dotted_NestedDottedKey()
    ' A dotted key INSIDE a nested object: only its own dot is escaped, the
    ' level separator stays plain.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim json As String
    json = "[{""o"":{""x.y"":5}}]"

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_DK3", ws.Range("A1"), json)

    AssertEquals "o.x\.y", lo.ListColumns(1).name, "level dot plain, key dot escaped"
    AssertEquals Json_Stringify(Json_Parse(json)), Excel_ListObjectToJson(lo), _
        "nested dotted key round-trips"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Dotted_TableRootCannotAddress()
    ' An UNESCAPED dot always splits levels, so "$.a.b.items" cannot reach
    ' the literal key "a.b" and raises 1160. The escaped form is the next
    ' test; this one pins that plain dots keep their meaning.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim gotNumber As Long
    On Error Resume Next
    Excel_UpsertListObjectFromJsonAtRoot ws, "T_DK4", ws.Range("A1"), _
        "{""a.b"":{""items"":[{""x"":1}]}}", "$.a.b.items"
    gotNumber = Err.Number
    On Error GoTo Clean

    AssertEquals vbObjectError + 1160, gotNumber, "dotted key on tableRoot raises 1160"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Dotted_TableRootEscapedDot()
    ' The document holds BOTH the literal key "a.b" and real a -> b nesting,
    ' with different rows under each. The escaped path must pick the literal
    ' key (streamed), the plain path the nesting (also streamed).
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim json As String
    json = "{""a.b"":{""items"":[{""x"":1},{""x"":2}]}," & _
        """a"":{""b"":{""items"":[{""x"":99}]}}}"

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_DK5", ws.Range("A1"), json, "$.a\.b.items")
    AssertEquals "[{""x"":1},{""x"":2}]", Excel_ListObjectToJson(lo), _
        "escaped dot addresses the literal a.b key"

    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_DK5b", ws.Range("A10"), json, "$.a.b.items")
    AssertEquals "[{""x"":99}]", Excel_ListObjectToJson(lo), _
        "plain dots still walk the nesting"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Dotted_TableRootEscapeWithBracket()
    ' Escape plus a bracket index: the stream declines bracket paths, so
    ' this resolves through the model path, which must apply the same
    ' escape rules.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_DK6", ws.Range("A1"), _
        "{""a.b"":[[{""x"":7}]]}", "$.a\.b[0]")

    AssertEquals "[{""x"":7}]", Excel_ListObjectToJson(lo), _
        "escape plus bracket index via the model path"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Dotted_ResolvePathEscaped()
    ' Direct model resolution: "\." reaches a dotted key, "\\" a key with a
    ' literal backslash.
    Dim v As Variant
    Json_ParseInto "{""a.b"":{""c"":42},""p\\q"":7}", v

    Dim out As Variant
    AssertEquals True, Json_TryResolvePath(v, "$.a\.b.c", out), "escaped dot resolves"
    AssertEquals 42, out, "value under the dotted key"
    AssertEquals True, Json_TryResolvePath(v, "$.p\\q", out), "escaped backslash resolves"
    AssertEquals 7, out, "value under the backslash key"
    AssertEquals False, Json_TryResolvePath(v, "$.a.b.c", out), "plain dots do not reach the dotted key"
End Sub
