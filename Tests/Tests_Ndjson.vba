Attribute VB_Name = "Tests_Ndjson"
Option Explicit

' =============================================================================
' NDJSON -> JSON tests
'
' Validates NdjsonToJson / NdjsonFileToJson:
'   - line assembly into a JSON array
'   - \n, \r\n, \r line endings
'   - blank and space-padded lines skipped
'   - single line, empty input, trailing newline
'   - the result parses to an array of the right length
' =============================================================================

Public Sub RunAll_NdjsonTests_StopOnFail()
    On Error GoTo Fail

    Test_Ndjson_Basic
    Test_Ndjson_BlankLines
    Test_Ndjson_CRLF
    Test_Ndjson_CR
    Test_Ndjson_SingleLine
    Test_Ndjson_Empty
    Test_Ndjson_TrailingNewline
    Test_Ndjson_LineWhitespace
    Test_Ndjson_ParsesToArray
    Test_Ndjson_File

    MsgBox "All NDJSON->JSON tests passed.", vbInformation
    Exit Sub

Fail:
    Err.Raise vbObjectError + 740, "mNdjsonTests", _
        "NDJSON test run failed. Err " & Err.Number & ": " & Err.Description
End Sub


' =============================================================================
' ASSERTS
' =============================================================================

Private Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal message As String)
    If expected <> actual Then
        Err.Raise vbObjectError + 741, "mNdjsonTests", _
            message & " expected=" & CStr(expected) & " actual=" & CStr(actual)
    End If
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 742, "mNdjsonTests", message
End Sub


' =============================================================================
' TESTS
' =============================================================================

Private Sub Test_Ndjson_Basic()
    AssertEquals "[{""id"":1},{""id"":2}]", _
        NdjsonToJson("{""id"":1}" & vbLf & "{""id"":2}"), "basic two records"
End Sub

Private Sub Test_Ndjson_BlankLines()
    AssertEquals "[{""a"":1},{""b"":2}]", _
        NdjsonToJson("{""a"":1}" & vbLf & vbLf & "{""b"":2}"), "interior blank line skipped"
End Sub

Private Sub Test_Ndjson_CRLF()
    AssertEquals "[{""a"":1},{""b"":2}]", _
        NdjsonToJson("{""a"":1}" & vbCrLf & "{""b"":2}"), "CRLF line endings"
End Sub

Private Sub Test_Ndjson_CR()
    AssertEquals "[{""a"":1},{""b"":2}]", _
        NdjsonToJson("{""a"":1}" & vbCr & "{""b"":2}"), "lone CR line ending"
End Sub

Private Sub Test_Ndjson_SingleLine()
    AssertEquals "[{""a"":1}]", NdjsonToJson("{""a"":1}"), "single line, no newline"
End Sub

Private Sub Test_Ndjson_Empty()
    AssertEquals "[]", NdjsonToJson(""), "empty input"
    AssertEquals "[]", NdjsonToJson(vbLf & vbLf), "all-blank input"
End Sub

Private Sub Test_Ndjson_TrailingNewline()
    AssertEquals "[{""a"":1}]", NdjsonToJson("{""a"":1}" & vbLf), "trailing newline"
End Sub

Private Sub Test_Ndjson_LineWhitespace()
    AssertEquals "[{""a"":1},{""b"":2}]", _
        NdjsonToJson("  {""a"":1}  " & vbLf & "{""b"":2}"), "line whitespace trimmed"
End Sub

Private Sub Test_Ndjson_ParsesToArray()
    Dim v As Variant
    Json_ParseInto NdjsonToJson("{""id"":1,""n"":""a""}" & vbLf & "{""id"":2,""n"":""b""}"), v

    AssertTrue Json_IsArray(v), "result parses to a JSON array"
    AssertEquals 2, v.count, "record count"
End Sub

Private Sub Test_Ndjson_File()
    Dim path As String
    path = Environ$("TEMP") & "\ndjson_test_" & Format(Now, "hhmmss") & ".ndjson"

    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, "{""x"":10}"
    Print #f, "{""x"":20}"
    Close #f

    AssertEquals "[{""x"":10},{""x"":20}]", NdjsonFileToJson(path), "file round trip"
End Sub
