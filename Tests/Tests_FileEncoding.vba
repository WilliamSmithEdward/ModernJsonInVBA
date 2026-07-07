Attribute VB_Name = "Tests_FileEncoding"
Option Explicit

' =============================================================================
' Json_ReadTextFile encoding tests
'
' Files are written as raw bytes (so this source stays pure ASCII) and read
' back through Json_ReadTextFile or a FileToJson adapter. Covered:
'   - pure ASCII (the fast path)
'   - UTF-8 without BOM: 2-byte (e-acute) and 4-byte (emoji) sequences
'   - UTF-8 with BOM: BOM stripped
'   - UTF-16 LE and BE with BOMs
'   - invalid UTF-8 falls back to the ANSI codepage (legacy behavior)
'   - end to end: a UTF-8 NDJSON file through NdjsonFileToJson + Json_Parse
'
' Expected characters are asserted by code point (AscW), never by literal.
' =============================================================================

Public Sub RunAll_FileEncodingTests_StopOnFail()
    On Error GoTo Fail

    Test_Enc_Ascii
    Test_Enc_Utf8_TwoByte
    Test_Enc_Utf8_FourByte
    Test_Enc_Utf8_Bom
    Test_Enc_Utf16LE
    Test_Enc_Utf16BE
    Test_Enc_AnsiFallback
    Test_Enc_EmptyFile
    Test_Enc_EndToEnd_Ndjson

    MsgBox "All file-encoding tests passed.", vbInformation
    Exit Sub

Fail:
    Err.Raise vbObjectError + 820, "mFileEncodingTests", _
        "File-encoding test run failed. Err " & Err.Number & ": " & Err.Description
End Sub


' =============================================================================
' ASSERTS AND HELPERS
' =============================================================================

Private Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal message As String)
    If expected <> actual Then
        Err.Raise vbObjectError + 821, "mFileEncodingTests", _
            message & " expected=" & CStr(expected) & " actual=" & CStr(actual)
    End If
End Sub

' Write raw bytes to a fresh temp file and return its path.
Private Function WriteBytesFile(ByVal tag As String, ByRef b() As Byte) As String
    Dim path As String
    path = Environ$("TEMP") & "\jsonenc_" & tag & "_" & Format(Now, "hhmmss") & ".bin"

    Dim f As Integer
    f = FreeFile
    Open path For Binary Access Write As #f
    Put #f, 1, b
    Close #f

    WriteBytesFile = path
End Function

' Build a Byte array from ASCII text plus raw byte values: each character of
' asciiPart is appended, then each entry of extraBytes. Keeps test data
' readable while the interesting bytes stay explicit.
Private Function BuildBytes(ByVal asciiPrefix As String, ByRef middleBytes As Variant, ByVal asciiSuffix As String) As Byte()
    Dim total As Long
    total = Len(asciiPrefix) + (UBound(middleBytes) - LBound(middleBytes) + 1) + Len(asciiSuffix)

    Dim b() As Byte
    ReDim b(0 To total - 1) As Byte

    Dim p As Long
    Dim i As Long

    For i = 1 To Len(asciiPrefix)
        b(p) = Asc(Mid$(asciiPrefix, i, 1))
        p = p + 1
    Next i

    For i = LBound(middleBytes) To UBound(middleBytes)
        b(p) = CByte(middleBytes(i))
        p = p + 1
    Next i

    For i = 1 To Len(asciiSuffix)
        b(p) = Asc(Mid$(asciiSuffix, i, 1))
        p = p + 1
    Next i

    BuildBytes = b
End Function


' =============================================================================
' TESTS
' =============================================================================

Private Sub Test_Enc_Ascii()
    Dim b() As Byte
    b = BuildBytes("[1,2,3]", Array(), "")

    Dim path As String
    path = WriteBytesFile("ascii", b)

    AssertEquals "[1,2,3]", Json_ReadTextFile(path), "ascii file"
    Kill path
End Sub

Private Sub Test_Enc_Utf8_TwoByte()
    ' {"n":"e-acute"} with e-acute as UTF-8 C3 A9 -> one char U+00E9.
    Dim b() As Byte
    b = BuildBytes("{""n"":""", Array(&HC3, &HA9), """}")

    Dim path As String
    path = WriteBytesFile("utf8two", b)

    Dim txt As String
    txt = Json_ReadTextFile(path)
    Kill path

    ' {"n":"?"} is 9 characters; the accented char sits at position 7.
    AssertEquals 9, Len(txt), "two-byte length"
    AssertEquals 233, AscW(Mid$(txt, 7, 1)), "e-acute decoded to U+00E9"

    ' And it parses: the decoded value is the single accented character.
    Dim v As Variant
    Json_ParseInto txt, v
    AssertEquals ChrW$(233), Json_ObjGet(v, "n"), "parsed accented value"
End Sub

Private Sub Test_Enc_Utf8_FourByte()
    ' {"e":"emoji"} with U+1F600 as UTF-8 F0 9F 98 80 -> surrogate pair
    ' D83D DE00 (two UTF-16 units).
    Dim b() As Byte
    b = BuildBytes("{""e"":""", Array(&HF0, &H9F, &H98, &H80), """}")

    Dim path As String
    path = WriteBytesFile("utf8four", b)

    Dim txt As String
    txt = Json_ReadTextFile(path)
    Kill path

    ' {"e":"??"} is 10 UTF-16 units; the surrogate pair sits at 7 and 8.
    AssertEquals 10, Len(txt), "four-byte length (surrogate pair is 2 units)"
    AssertEquals &HD83D&, AscW(Mid$(txt, 7, 1)) And &HFFFF&, "high surrogate"
    AssertEquals &HDE00&, AscW(Mid$(txt, 8, 1)) And &HFFFF&, "low surrogate"
End Sub

Private Sub Test_Enc_Utf8_Bom()
    Dim b() As Byte
    b = BuildBytes("", Array(&HEF, &HBB, &HBF), "[true]")

    Dim path As String
    path = WriteBytesFile("utf8bom", b)

    Dim txt As String
    txt = Json_ReadTextFile(path)
    Kill path

    AssertEquals "[true]", txt, "BOM stripped"
    AssertEquals 91, AscW(Left$(txt, 1)), "first char is the bracket, not a BOM"
End Sub

Private Sub Test_Enc_Utf16LE()
    ' FF FE then "[1]" as UTF-16 LE pairs.
    Dim b() As Byte
    b = BuildBytes("", Array(&HFF, &HFE, 91, 0, 49, 0, 93, 0), "")

    Dim path As String
    path = WriteBytesFile("utf16le", b)

    AssertEquals "[1]", Json_ReadTextFile(path), "UTF-16 LE with BOM"
    Kill path
End Sub

Private Sub Test_Enc_Utf16BE()
    ' FE FF then "[1]" as UTF-16 BE pairs.
    Dim b() As Byte
    b = BuildBytes("", Array(&HFE, &HFF, 0, 91, 0, 49, 0, 93), "")

    Dim path As String
    path = WriteBytesFile("utf16be", b)

    AssertEquals "[1]", Json_ReadTextFile(path), "UTF-16 BE with BOM"
    Kill path
End Sub

Private Sub Test_Enc_AnsiFallback()
    ' A lone E9 byte is invalid UTF-8 (it expects continuation bytes), so the
    ' reader falls back to the ANSI codepage. On Western codepages E9 maps to
    ' U+00E9, which is also what the legacy reader produced.
    Dim b() As Byte
    b = BuildBytes("{""n"":""", Array(&HE9), """}")

    Dim path As String
    path = WriteBytesFile("ansi", b)

    Dim txt As String
    txt = Json_ReadTextFile(path)
    Kill path

    ' {"n":"?"} is 9 characters; the fallback char sits at position 7.
    AssertEquals 9, Len(txt), "ansi fallback length"
    AssertEquals 233, AscW(Mid$(txt, 7, 1)), "ansi E9 maps to U+00E9"
End Sub

Private Sub Test_Enc_EmptyFile()
    Dim b() As Byte
    ReDim b(0 To 0) As Byte
    b(0) = 32   ' single space: Put of a zero-length array is not possible

    Dim path As String
    path = WriteBytesFile("space", b)

    AssertEquals " ", Json_ReadTextFile(path), "single-space file"
    Kill path
End Sub

Private Sub Test_Enc_EndToEnd_Ndjson()
    ' Two NDJSON lines, UTF-8 with BOM, accented name in line one: the full
    ' path through NdjsonFileToJson -> Json_Parse must land the right chars.
    Dim b() As Byte
    b = BuildBytes("", Array(&HEF, &HBB, &HBF), "{""id"":1,""name"":""Ren")

    Dim b2() As Byte
    b2 = BuildBytes("", Array(&HC3, &HA9), """}" & Chr$(10) & "{""id"":2,""name"":""Ana""}")

    ' Concatenate the two parts.
    Dim whole() As Byte
    ReDim whole(0 To UBound(b) + UBound(b2) + 1) As Byte
    Dim i As Long
    For i = 0 To UBound(b)
        whole(i) = b(i)
    Next i
    For i = 0 To UBound(b2)
        whole(UBound(b) + 1 + i) = b2(i)
    Next i

    Dim path As String
    path = WriteBytesFile("ndjend", whole)

    Dim json As String
    json = NdjsonFileToJson(path)
    Kill path

    Dim v As Variant
    Json_ParseInto json, v
    AssertEquals 2, v.count, "two records"
    AssertEquals "Ren" & ChrW$(233), Json_ObjGet(v(1), "name"), "accented name decoded"
    AssertEquals "Ana", Json_ObjGet(v(2), "name"), "second record"
End Sub
