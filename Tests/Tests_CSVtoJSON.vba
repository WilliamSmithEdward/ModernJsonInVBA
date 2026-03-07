Option Explicit

' =============================================================================
' CSV -> JSON Tests
'
' These tests validate:
'   - RFC-4180 parsing correctness
'   - JSON conversion correctness
'   - edge cases (quotes, commas, newlines)
'   - uneven rows
'   - empty fields
'
' Assumptions:
'   - CsvFileToJson(filePath) exists
' =============================================================================


' =============================================================================
' RUNNER
' =============================================================================
Public Sub RunAll_CsvJsonTests_StopOnFail()

    On Error GoTo Fail

    Test_Csv_Simple_WithAsserts
    Test_Csv_QuotedComma_WithAsserts
    Test_Csv_EscapedQuotes_WithAsserts
    Test_Csv_EmptyFields_WithAsserts
    Test_Csv_UnevenRows_WithAsserts
    Test_Csv_QuotedNewline_WithAsserts
    Test_Csv_LastRowWithoutNewline_WithAsserts
    Test_Csv_MultilineQuotedComma_WithAsserts
    Test_Csv_HeaderOnly_WithAsserts
    Test_Csv_SingleColumn_WithAsserts
    Test_Csv_AllEmptyFields_WithAsserts
    Test_Csv_EscapedQuotesMultiline_WithAsserts
    Test_Csv_CRLF_LineEndings_WithAsserts
    Test_Csv_WhitespacePreserved_WithAsserts

    MsgBox "All CSV->JSON tests passed.", vbInformation
    Exit Sub

Fail:
    Dim msg As String
    msg = "CSV test run failed." & vbCrLf & _
          "Err " & Err.Number & ": " & Err.Description

    Err.Clear
    Err.Raise vbObjectError + 720, "mCsvJsonTests", msg

End Sub


' =============================================================================
' ASSERTS
' =============================================================================

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 710, "mCsvJsonTests", message
End Sub

Private Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal message As String)
    If expected <> actual Then
        Err.Raise vbObjectError + 711, "mCsvJsonTests", _
            message & " expected=" & CStr(expected) & " actual=" & CStr(actual)
    End If
End Sub


' =============================================================================
' HELPERS
' =============================================================================

Private Function WriteTempCsv(ByVal text As String) As String

    Dim path As String
    path = Environ$("TEMP") & "\csv_test_" & Format(Now, "hhmmss") & ".csv"

    Dim f As Integer
    f = FreeFile

    Open path For Output As #f
    Print #f, text
    Close #f

    WriteTempCsv = path

End Function


' =============================================================================
' TESTS
' =============================================================================


' -----------------------------------------------------------------------------
' TEST 1: Simple CSV
' -----------------------------------------------------------------------------
Public Sub Test_Csv_Simple_WithAsserts()

    Dim csv As String
    csv = _
    "id,name" & vbLf & _
    "1,Alice" & vbLf & _
    "2,Bob"

    Dim path As String
    path = WriteTempCsv(csv)

    Dim json As String
    json = CsvFileToJson(path)

    AssertTrue InStr(json, """id"":""1""") > 0, "Missing id=1"
    AssertTrue InStr(json, """name"":""Alice""") > 0, "Missing Alice"

End Sub


' -----------------------------------------------------------------------------
' TEST 2: Quoted comma
' -----------------------------------------------------------------------------
Public Sub Test_Csv_QuotedComma_WithAsserts()

    Dim csv As String
    csv = _
    "name,city" & vbLf & _
    "Alice,""New York, NY"""

    Dim path As String
    path = WriteTempCsv(csv)

    Dim json As String
    json = CsvFileToJson(path)

    AssertTrue InStr(json, "New York, NY") > 0, "Quoted comma failed"

End Sub


' -----------------------------------------------------------------------------
' TEST 3: Escaped quotes
' -----------------------------------------------------------------------------
Public Sub Test_Csv_EscapedQuotes_WithAsserts()

    Dim csv As String
    csv = _
    "name,quote" & vbLf & _
    "Alice,""She said """"Hello"""""""

    Dim path As String
    path = WriteTempCsv(csv)

    Dim json As String
    json = CsvFileToJson(path)

    AssertTrue InStr(json, "Hello") > 0, "Escaped quotes failed"

End Sub


' -----------------------------------------------------------------------------
' TEST 4: Empty fields
' -----------------------------------------------------------------------------
Public Sub Test_Csv_EmptyFields_WithAsserts()

    Dim csv As String
    csv = _
    "a,b,c" & vbLf & _
    "1,,3"

    Dim path As String
    path = WriteTempCsv(csv)

    Dim json As String
    json = CsvFileToJson(path)

    AssertTrue InStr(json, """b"":""""") > 0, "Empty field not preserved"

End Sub


' -----------------------------------------------------------------------------
' TEST 5: Uneven rows
' -----------------------------------------------------------------------------
Public Sub Test_Csv_UnevenRows_WithAsserts()

    Dim csv As String
    csv = _
    "a,b,c" & vbLf & _
    "1,2"

    Dim path As String
    path = WriteTempCsv(csv)

    Dim json As String
    json = CsvFileToJson(path)

    AssertTrue InStr(json, """c"":""""") > 0, "Missing column not handled"

End Sub


' -----------------------------------------------------------------------------
' TEST 6: Quoted newline
' -----------------------------------------------------------------------------
Public Sub Test_Csv_QuotedNewline_WithAsserts()

    Dim csv As String
    csv = _
    "id,comment" & vbLf & _
    "1,""Hello" & vbLf & "World"""

    Dim path As String
    path = WriteTempCsv(csv)

    Dim json As String
    json = CsvFileToJson(path)

    AssertTrue InStr(json, "Hello") > 0, "Quoted newline failed"

End Sub

' -----------------------------------------------------------------------------
' TEST 7: Last row without newline
'
' Ensures the parser does not drop the final record when the CSV file
' does not end with a newline character.
'
' Input CSV:
'   id,name
'   1,Alice
'   2,Bob
'
' (No trailing newline after Bob)
'
' Expected:
'   Both rows appear in JSON output.
' -----------------------------------------------------------------------------
Public Sub Test_Csv_LastRowWithoutNewline_WithAsserts()

    Dim csv As String
    csv = _
    "id,name" & vbLf & _
    "1,Alice" & vbLf & _
    "2,Bob"

    Dim path As String
    path = WriteTempCsv(csv)

    Dim json As String
    json = CsvFileToJson(path)

    AssertTrue InStr(json, """id"":""1""") > 0, "Missing first row"
    AssertTrue InStr(json, """name"":""Alice""") > 0, "Missing Alice"

    AssertTrue InStr(json, """id"":""2""") > 0, "Missing second row"
    AssertTrue InStr(json, """name"":""Bob""") > 0, "Missing Bob"

End Sub


' -----------------------------------------------------------------------------
' TEST 8: Multiline quoted field containing commas
' -----------------------------------------------------------------------------
Public Sub Test_Csv_MultilineQuotedComma_WithAsserts()

    Dim csv As String
    csv = _
    "id,comment" & vbLf & _
    "1,""Hello," & vbLf & _
    "world"""

    Dim path As String
    path = WriteTempCsv(csv)

    Dim json As String
    json = CsvFileToJson(path)

    AssertTrue InStr(json, """id"":""1""") > 0, "Missing id"
    AssertTrue InStr(json, "Hello,") > 0, "Comma inside quoted field lost"
    AssertTrue InStr(json, "world") > 0, "Multiline value lost"

End Sub


' -----------------------------------------------------------------------------
' TEST 9: Header only (no rows)
' -----------------------------------------------------------------------------
Public Sub Test_Csv_HeaderOnly_WithAsserts()

    Dim csv As String
    csv = "id,name"

    Dim path As String
    path = WriteTempCsv(csv)

    Dim json As String
    json = CsvFileToJson(path)

    AssertEquals "[]", json, "Header-only CSV should produce empty JSON array"

End Sub


' -----------------------------------------------------------------------------
' TEST 10: Single column CSV
' -----------------------------------------------------------------------------
Public Sub Test_Csv_SingleColumn_WithAsserts()

    Dim csv As String
    csv = _
    "id" & vbLf & _
    "1" & vbLf & _
    "2"

    Dim path As String
    path = WriteTempCsv(csv)

    Dim json As String
    json = CsvFileToJson(path)

    AssertTrue InStr(json, """id"":""1""") > 0, "Row1 missing"
    AssertTrue InStr(json, """id"":""2""") > 0, "Row2 missing"

End Sub


' -----------------------------------------------------------------------------
' TEST 11: Entirely empty row
' -----------------------------------------------------------------------------
Public Sub Test_Csv_AllEmptyFields_WithAsserts()

    Dim csv As String
    csv = _
    "a,b,c" & vbLf & _
    ",,"

    Dim path As String
    path = WriteTempCsv(csv)

    Dim json As String
    json = CsvFileToJson(path)

    AssertTrue InStr(json, """a"":""""") > 0, "Column a missing"
    AssertTrue InStr(json, """b"":""""") > 0, "Column b missing"
    AssertTrue InStr(json, """c"":""""") > 0, "Column c missing"

End Sub


' -----------------------------------------------------------------------------
' TEST 12: Escaped quotes inside multiline field
' -----------------------------------------------------------------------------
Public Sub Test_Csv_EscapedQuotesMultiline_WithAsserts()

    Dim csv As String
    csv = _
    "id,text" & vbLf & _
    "1,""He said """"hello"""" " & vbLf & _
    "again"""

    Dim path As String
    path = WriteTempCsv(csv)

    Dim json As String
    json = CsvFileToJson(path)

    AssertTrue InStr(json, "hello") > 0, "Escaped quotes lost"
    AssertTrue InStr(json, "again") > 0, "Multiline continuation lost"

End Sub


' -----------------------------------------------------------------------------
' TEST 13: CRLF line endings
' -----------------------------------------------------------------------------
Public Sub Test_Csv_CRLF_LineEndings_WithAsserts()

    Dim csv As String
    csv = _
    "id,name" & vbCrLf & _
    "1,Alice" & vbCrLf & _
    "2,Bob"

    Dim path As String
    path = WriteTempCsv(csv)

    Dim json As String
    json = CsvFileToJson(path)

    AssertTrue InStr(json, """Alice""") > 0, "CRLF row1 missing"
    AssertTrue InStr(json, """Bob""") > 0, "CRLF row2 missing"

End Sub


' -----------------------------------------------------------------------------
' TEST 14: Whitespace preservation outside quotes
' -----------------------------------------------------------------------------
Public Sub Test_Csv_WhitespacePreserved_WithAsserts()

    Dim csv As String
    csv = _
    "a,b" & vbLf & _
    "1,  two"

    Dim path As String
    path = WriteTempCsv(csv)

    Dim json As String
    json = CsvFileToJson(path)

    AssertTrue InStr(json, """b"":""  two""") > 0, "Whitespace lost"

End Sub
