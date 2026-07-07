Attribute VB_Name = "Tests_NestedRoot"
Option Explicit

' =============================================================================
' Nested-tableRoot streaming tests
'
' The streaming table sink handles tableRoots such as "$.data.items" by
' descending past sibling members with validating skips. These tests check:
'   - nested imports match the equivalent root-level import exactly
'   - siblings before and after the table (including strings full of
'     braces, brackets, and escapes) are skipped without confusion
'   - duplicate keys resolve first-match, like the model path
'   - declined shapes (missing key, non-array root, bracket paths, null)
'     still work or raise through the model path with its error numbers
'   - malformed siblings and trailing text are rejected, so streaming keeps
'     whole-document validation
'
' Patterns follow Tests_UpsertReturns: fresh sheet per test, cleanup that
' re-raises (a failed assert must surface as an error, not a modal dialog),
' table sameness by range address.
' =============================================================================

Public Sub RunAll_NestedRootTests_StopOnFail()
    On Error GoTo Fail

    Test_Nested_Basic
    Test_Nested_ParityWithRoot
    Test_Nested_DeepPath
    Test_Nested_NastySiblings
    Test_Nested_FirstKeyWins
    Test_Nested_EmptyArray
    Test_Nested_MissingKey
    Test_Nested_NotAnArray
    Test_Nested_NullRoot
    Test_Nested_MalformedSibling
    Test_Nested_TrailingGarbage
    Test_Nested_BracketPathDeclines
    Test_Nested_Whitespace

    MsgBox "All nested-tableRoot streaming tests passed.", vbInformation
    Exit Sub

Fail:
    Err.Raise vbObjectError + 810, "mNestedRootTests", _
        "Nested-root test run failed. Err " & Err.Number & ": " & Err.Description
End Sub


' =============================================================================
' ASSERTS AND HELPERS
' =============================================================================

Private Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal message As String)
    If expected <> actual Then
        Err.Raise vbObjectError + 811, "mNestedRootTests", _
            message & " expected=" & CStr(expected) & " actual=" & CStr(actual)
    End If
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 812, "mNestedRootTests", message
End Sub

' Run an upsert expected to raise, and check the error number.
Private Sub AssertUpsertRaises( _
    ByVal expectedNumber As Long, _
    ByVal ws As Worksheet, _
    ByVal tableName As String, _
    ByVal json As String, _
    ByVal tableRoot As String, _
    ByVal message As String _
)
    Dim gotNumber As Long

    On Error Resume Next
    Excel_UpsertListObjectFromJsonAtRoot ws, tableName, ws.Range("A1"), json, tableRoot
    gotNumber = Err.Number
    On Error GoTo 0

    If gotNumber <> expectedNumber Then
        Err.Raise vbObjectError + 813, "mNestedRootTests", _
            message & " expected err=" & expectedNumber & " got=" & gotNumber
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

Private Sub Test_Nested_Basic()
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim json As String
    json = "{""meta"":{""page"":1,""per"":50},""data"":{""items"":" & _
        "[{""id"":1,""name"":""Alice""},{""id"":2,""name"":""Bob""}]},""ok"":true}"

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_NB", ws.Range("A1"), json, "$.data.items")

    AssertEquals 2, lo.ListRows.count, "row count"
    AssertEquals 2, lo.ListColumns.count, "column count"
    AssertEquals "[{""id"":1,""name"":""Alice""},{""id"":2,""name"":""Bob""}]", _
        Excel_ListObjectToJson(lo), "exported table"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Nested_ParityWithRoot()
    ' The same rows imported at "$" and nested under "$.data.items" must
    ' produce identical tables.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim rows As String
    rows = "[{""id"":1,""v"":{""a"":true,""b"":null}},{""id"":2,""v"":{""a"":false},""x"":1.5}]"

    Dim loRoot As ListObject
    Set loRoot = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_NP1", ws.Range("A1"), rows, "$")

    Dim loNested As ListObject
    Set loNested = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_NP2", ws.Range("A20"), _
        "{""before"":[9,8],""data"":{""items"":" & rows & "},""after"":{""z"":0}}", "$.data.items")

    AssertEquals Excel_ListObjectToJson(loRoot), Excel_ListObjectToJson(loNested), _
        "nested import equals root import"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Nested_DeepPath()
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_ND", ws.Range("A1"), _
        "{""a"":{""b"":{""c"":[{""n"":10},{""n"":20},{""n"":30}]}}}", "$.a.b.c")

    AssertEquals 3, lo.ListRows.count, "deep path row count"
    AssertEquals "[{""n"":10},{""n"":20},{""n"":30}]", Excel_ListObjectToJson(lo), "deep path export"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Nested_NastySiblings()
    ' Siblings on both sides of the table, at both levels, holding strings
    ' full of structural characters and escapes: the skip scanners must not
    ' mistake text content for structure.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim json As String
    json = "{""trap1"":""}]"",""deep"":{""trap2"":""a\""b{["",""nums"":[1,-2.5,3e2]," & _
        """items"":[{""id"":1}],""trap3"":{""inner"":[{""x"":""]}""}]}}," & _
        """trap4"":[[[]]],""trap5"":""\u00e9\\n"",""last"":null}"

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_NN", ws.Range("A1"), json, "$.deep.items")

    AssertEquals 1, lo.ListRows.count, "nasty siblings row count"
    AssertEquals "[{""id"":1}]", Excel_ListObjectToJson(lo), "nasty siblings export"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Nested_FirstKeyWins()
    ' Duplicate keys on the descent level: the first match is taken, the
    ' later duplicate is skipped. Same first-match rule as Json_TryObjGet.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_NF", ws.Range("A1"), _
        "{""d"":{""items"":[{""id"":1}]},""d"":{""items"":[{""id"":99},{""id"":98}]}}", "$.d.items")

    AssertEquals 1, lo.ListRows.count, "first duplicate wins"
    AssertEquals "[{""id"":1}]", Excel_ListObjectToJson(lo), "first duplicate export"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Nested_EmptyArray()
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_NE", ws.Range("A1"), _
        "{""data"":{""items"":[],""note"":""empty""}}", "$.data.items")

    AssertEquals 0, lo.ListRows.count, "empty nested array row count"
    AssertEquals 1, lo.ListColumns.count, "empty nested array column count"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Nested_MissingKey()
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    AssertUpsertRaises vbObjectError + 1160, ws, "T_NM", _
        "{""data"":{""rows"":[{""id"":1}]}}", "$.data.items", "missing key raises 1160"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Nested_NotAnArray()
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    AssertUpsertRaises vbObjectError + 1162, ws, "T_NA", _
        "{""data"":{""items"":{""id"":1}}}", "$.data.items", "object at root raises 1162"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Nested_NullRoot()
    ' items: null resolves like an empty table (model-path semantics; the
    ' stream declines on the non-array value).
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_NU", ws.Range("A1"), _
        "{""data"":{""items"":null}}", "$.data.items")

    AssertEquals 0, lo.ListRows.count, "null root row count"
    AssertEquals "value", lo.ListColumns(1).name, "null root placeholder column"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Nested_MalformedSibling()
    ' A defect in a SIBLING the descent skips must still be rejected:
    ' streaming keeps whole-document validation. "tru" raises 525.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    AssertUpsertRaises vbObjectError + 525, ws, "T_NX", _
        "{""bad"":tru,""data"":{""items"":[{""id"":1}]}}", "$.data.items", _
        "malformed sibling raises 525"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Nested_TrailingGarbage()
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    AssertUpsertRaises vbObjectError + 700, ws, "T_NT", _
        "{""data"":{""items"":[{""id"":1}]}} extra", "$.data.items", _
        "trailing text raises 700"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Nested_BracketPathDeclines()
    ' Bracket indices are outside the streamed shape: the stream declines
    ' and the model path resolves them, so the import still works.
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_NBK", ws.Range("A1"), _
        "{""wrap"":[[{""id"":7}]]}", "$.wrap[0]")

    AssertEquals 1, lo.ListRows.count, "bracket path row count"
    AssertEquals "[{""id"":7}]", Excel_ListObjectToJson(lo), "bracket path export"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub

Private Sub Test_Nested_Whitespace()
    Dim ws As Worksheet
    Set ws = FreshSheet()
    On Error GoTo Clean

    Dim json As String
    json = "  {  ""pre""  :  [ 1 , 2 ]  ,  ""data""  :  {  ""items""  :  " & vbCrLf & _
        "[ { ""id"" : 1 } , { ""id"" : 2 } ]  ,  ""post""  :  ""p""  }  ,  ""tail""  :  0  }  "

    Dim lo As ListObject
    Set lo = Excel_UpsertListObjectFromJsonAtRoot(ws, "T_NW", ws.Range("A1"), json, "$.data.items")

    AssertEquals 2, lo.ListRows.count, "whitespace row count"
    AssertEquals "[{""id"":1},{""id"":2}]", Excel_ListObjectToJson(lo), "whitespace export"

    DropSheet ws
    Exit Sub
Clean:
    CleanupAndReraise ws
End Sub
