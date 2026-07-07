Attribute VB_Name = "Tests_StringifyPretty"
Option Explicit

' =============================================================================
' Pretty (indented) serialization tests
'
' Validates Json_StringifyPretty:
'   - scalars, empty object, empty array
'   - one member, nested objects and arrays, array of objects
'   - the indentUnit option (two spaces by default, tabs on request)
'   - semantic equality with the compact Json_Stringify (round trip)
'
' Expected strings are built with vbLf so the tests do not depend on the
' editor's line endings. Models come from Json_Parse, never hand-built.
' =============================================================================

Public Sub RunAll_StringifyPrettyTests_StopOnFail()
    On Error GoTo Fail

    Test_Pretty_Scalars
    Test_Pretty_EmptyContainers
    Test_Pretty_SimpleObject
    Test_Pretty_Nested
    Test_Pretty_ArrayOfObjects
    Test_Pretty_NestedEmpties
    Test_Pretty_TabIndent
    Test_Pretty_ArrayRoot
    Test_Pretty_RoundTrip

    MsgBox "All pretty-print tests passed.", vbInformation
    Exit Sub

Fail:
    Err.Raise vbObjectError + 760, "mPrettyTests", _
        "Pretty-print test run failed. Err " & Err.Number & ": " & Err.Description
End Sub


' =============================================================================
' ASSERTS
' =============================================================================

Private Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal message As String)
    If expected <> actual Then
        Err.Raise vbObjectError + 761, "mPrettyTests", _
            message & vbLf & "expected:" & vbLf & CStr(expected) & vbLf & _
            "actual:" & vbLf & CStr(actual)
    End If
End Sub


' =============================================================================
' TESTS
' =============================================================================

Private Sub Test_Pretty_Scalars()
    ' A top-level scalar has no line structure; it matches the compact form.
    AssertEquals "42", Json_StringifyPretty(Json_Parse("42")), "integer scalar"
    AssertEquals "true", Json_StringifyPretty(Json_Parse("true")), "boolean scalar"
    AssertEquals "null", Json_StringifyPretty(Json_Parse("null")), "null scalar"
    AssertEquals """hi""", Json_StringifyPretty(Json_Parse("""hi""")), "string scalar"
End Sub

Private Sub Test_Pretty_EmptyContainers()
    AssertEquals "{}", Json_StringifyPretty(Json_Parse("{}")), "empty object stays inline"
    AssertEquals "[]", Json_StringifyPretty(Json_Parse("[]")), "empty array stays inline"
End Sub

Private Sub Test_Pretty_SimpleObject()
    Dim expected As String
    expected = "{" & vbLf & _
               "  ""a"": 1" & vbLf & _
               "}"

    AssertEquals expected, Json_StringifyPretty(Json_Parse("{""a"":1}")), "single member, two-space default"
End Sub

Private Sub Test_Pretty_Nested()
    Dim src As String
    src = "{""a"":1,""b"":[2,3],""c"":{""d"":true}}"

    Dim expected As String
    expected = "{" & vbLf & _
               "  ""a"": 1," & vbLf & _
               "  ""b"": [" & vbLf & _
               "    2," & vbLf & _
               "    3" & vbLf & _
               "  ]," & vbLf & _
               "  ""c"": {" & vbLf & _
               "    ""d"": true" & vbLf & _
               "  }" & vbLf & _
               "}"

    AssertEquals expected, Json_StringifyPretty(Json_Parse(src)), "nested object and array"
End Sub

Private Sub Test_Pretty_ArrayOfObjects()
    ' The common table-export shape.
    Dim src As String
    src = "[{""id"":1},{""id"":2}]"

    Dim expected As String
    expected = "[" & vbLf & _
               "  {" & vbLf & _
               "    ""id"": 1" & vbLf & _
               "  }," & vbLf & _
               "  {" & vbLf & _
               "    ""id"": 2" & vbLf & _
               "  }" & vbLf & _
               "]"

    AssertEquals expected, Json_StringifyPretty(Json_Parse(src)), "array of objects"
End Sub

Private Sub Test_Pretty_NestedEmpties()
    Dim src As String
    src = "{""a"":{},""b"":[]}"

    Dim expected As String
    expected = "{" & vbLf & _
               "  ""a"": {}," & vbLf & _
               "  ""b"": []" & vbLf & _
               "}"

    AssertEquals expected, Json_StringifyPretty(Json_Parse(src)), "empty containers as members"
End Sub

Private Sub Test_Pretty_TabIndent()
    Dim src As String
    src = "{""a"":[1]}"

    Dim expected As String
    expected = "{" & vbLf & _
               vbTab & """a"": [" & vbLf & _
               vbTab & vbTab & "1" & vbLf & _
               vbTab & "]" & vbLf & _
               "}"

    AssertEquals expected, Json_StringifyPretty(Json_Parse(src), vbTab), "tab indent unit"
End Sub

Private Sub Test_Pretty_ArrayRoot()
    Dim expected As String
    expected = "[" & vbLf & _
               "  1," & vbLf & _
               "  2" & vbLf & _
               "]"

    AssertEquals expected, Json_StringifyPretty(Json_Parse("[1,2]")), "array at the root"
End Sub

Private Sub Test_Pretty_RoundTrip()
    ' Pretty output must carry the same data as compact output: parse the
    ' pretty text, re-serialize it compactly, and compare to the compact
    ' serialization of the original model. Exercises escaping (quote, tab),
    ' reals, booleans, null, and nesting.
    Dim src As String
    src = "{""name"":""A\""B"",""vals"":[1,2.5,true,null],""nested"":{""x"":""y\ttab""}}"

    Dim model As Variant
    Json_ParseInto src, model

    Dim pretty As String
    pretty = Json_StringifyPretty(model)

    AssertEquals Json_Stringify(model), Json_Stringify(Json_Parse(pretty)), _
        "pretty round trips to the same model as compact"
End Sub
