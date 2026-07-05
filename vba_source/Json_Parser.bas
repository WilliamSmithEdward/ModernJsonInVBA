Attribute VB_Name = "Json_Parser"
Option Explicit

' =============================================================================
' Module:      Json_Parser
' Project:     ModernJsonInVBA
'
' Recursive-descent JSON parser producing the library's in-memory model:
'
'   JSON Object => VBA Collection tagged with JSON_TAG_OBJECT in slot (1),
'                  then pairs as Array(key, value)
'   JSON Array  => VBA Collection (untagged)
'   Primitives  => Variant (Null, Boolean, Long/Double, String)
'
' Strictness follows RFC 8259: no leading zeros, no trailing commas, no
' unescaped control characters, surrogate pairs validated, no trailing text.
'
' Performance notes:
'   - Strings are read by scanning ahead with InStr and copying whole runs;
'     escape-free strings (the common case) are returned with a single Mid$.
'   - All character-level scanning reads UTF-16 code units from a byte-array
'     snapshot of the input (bytes = text is one native copy). Mid$-based
'     probing would allocate a fresh one-character string per position.
'
' Error numbers (all vbObjectError + n):
'   520 expected character
'   522 invalid escape          523 unterminated string
'   524 bad \uXXXX escape       525 bad literal (true/false/null)
'   526 unescaped control char  527 invalid surrogate pair
'   700 trailing characters     701 unexpected token
'   710 invalid number          711 invalid fraction   712 invalid exponent
'   730 expected ',' or ']'     760 expected ',' or '}'
' =============================================================================

Private Const ERR_SRC As String = "ModernJsonInVBA"

' Reader state: 1-based cursor over the input text. bytes holds the same
' text as UTF-16 code units (little-endian byte pairs) so scanning loops can
' read character codes without allocating.
Private Type JsonReader
    text As String
    bytes() As Byte
    pos As Long
End Type

' =============================================================================
' Public API
' =============================================================================

' Parse jsonText and return the model value:
'   - Object/Array: as Object (Collection)
'   - Primitive:    as Variant
'
' Raises vbObjectError + 700 when non-whitespace text follows the value.
Public Function Json_Parse(ByVal jsonText As String) As Variant
    Dim r As JsonReader
    JR_Init r, jsonText

    Dim tmp As Variant
    Json_ReadValue r, tmp

    JR_SkipWs r
    If Not JR_Eof(r) Then
        Err.Raise vbObjectError + 700, ERR_SRC, "Unexpected trailing characters at pos " & r.pos
    End If

    If IsObject(tmp) Then
        Dim o As Object
        Set o = tmp
        Set Json_Parse = o
    Else
        Json_Parse = tmp
    End If
End Function

' Parse jsonText into outValue. Unlike Json_Parse, the caller does not need
' to know in advance whether the root is an object or a primitive.
'
' Raises vbObjectError + 700 when non-whitespace text follows the value.
Public Sub Json_ParseInto(ByVal jsonText As String, ByRef outValue As Variant)
    Dim r As JsonReader
    JR_Init r, jsonText

    Json_ReadValue r, outValue

    JR_SkipWs r
    If Not JR_Eof(r) Then
        Err.Raise vbObjectError + 700, ERR_SRC, "Unexpected trailing characters at pos " & r.pos
    End If
End Sub

' =============================================================================
' Reader primitives
' =============================================================================

Private Sub JR_Init(ByRef r As JsonReader, ByVal jsonText As String)
    r.text = jsonText
    r.pos = 1
    If Len(jsonText) > 0 Then
        r.bytes = jsonText
    End If
End Sub

Private Function JR_Eof(ByRef r As JsonReader) As Boolean
    JR_Eof = (r.pos > Len(r.text))
End Function

' Character code (0..65535) at the cursor, or -1 at end of input. Unlike
' AscW this never goes negative for high code points, so range checks are
' straightforward.
Private Function JR_CodeAt(ByRef r As JsonReader) As Long
    If r.pos > Len(r.text) Then
        JR_CodeAt = -1
    Else
        Dim off As Long
        off = (r.pos - 1) * 2
        JR_CodeAt = r.bytes(off) + r.bytes(off + 1) * 256&
    End If
End Function

Private Sub JR_SkipWs(ByRef r As JsonReader)
    Dim L As Long
    L = Len(r.text)

    Do While r.pos <= L
        Dim off As Long
        off = (r.pos - 1) * 2

        Select Case r.bytes(off) + r.bytes(off + 1) * 256&
            Case 32, 9, 13, 10   ' space, tab, CR, LF
                r.pos = r.pos + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

Private Sub JR_ExpectChar(ByRef r As JsonReader, ByVal expectedCode As Long, ByRef expectedChar As String)
    JR_SkipWs r

    If JR_CodeAt(r) = expectedCode Then
        r.pos = r.pos + 1
        Exit Sub
    End If

    Dim ch As String
    If r.pos <= Len(r.text) Then
        ch = Mid$(r.text, r.pos, 1)
        r.pos = r.pos + 1
    Else
        ch = vbNullString
    End If

    Err.Raise vbObjectError + 520, ERR_SRC, _
        "Expected '" & expectedChar & "' at pos " & (r.pos - 1) & " but got '" & ch & "'"
End Sub

Private Sub JR_ExpectLiteral(ByRef r As JsonReader, ByVal lit As String)
    JR_SkipWs r

    If Mid$(r.text, r.pos, Len(lit)) <> lit Then
        Err.Raise vbObjectError + 525, ERR_SRC, _
            "Expected literal '" & lit & "' near pos " & r.pos
    End If

    r.pos = r.pos + Len(lit)
End Sub

' =============================================================================
' Value dispatch
' =============================================================================

Private Sub Json_ReadValue(ByRef r As JsonReader, ByRef outValue As Variant)
    JR_SkipWs r

    Select Case JR_CodeAt(r)
        Case 34                     ' quote
            outValue = JR_ReadJsonString(r)

        Case 116                    ' "t"
            JR_ExpectLiteral r, "true"
            outValue = True

        Case 102                    ' "f"
            JR_ExpectLiteral r, "false"
            outValue = False

        Case 110                    ' "n"
            JR_ExpectLiteral r, "null"
            outValue = Null

        Case 45, 48 To 57           ' "-", "0".."9"
            outValue = JR_ReadNumber(r)

        Case 91                     ' "["
            Dim arr As Collection
            Set arr = JR_ReadArray(r)
            Set outValue = arr

        Case 123                    ' "{"
            Dim obj As Collection
            Set obj = JR_ReadObject(r)
            Set outValue = obj

        Case Else
            Dim ch As String
            If r.pos <= Len(r.text) Then ch = Mid$(r.text, r.pos, 1) Else ch = vbNullString
            Err.Raise vbObjectError + 701, ERR_SRC, _
                "Unexpected token '" & ch & "' at pos " & r.pos
    End Select
End Sub

' =============================================================================
' Numbers
' =============================================================================

' Reads a strict JSON number. Integers that fit a Long come back as Long;
' everything else (fractions, exponents, overflow) comes back as Double.
Private Function JR_ReadNumber(ByRef r As JsonReader) As Variant
    JR_SkipWs r

    Dim startPos As Long
    startPos = r.pos

    Dim c As Long
    c = JR_CodeAt(r)
    If c = 45 Then          ' "-"
        r.pos = r.pos + 1
        c = JR_CodeAt(r)
    End If

    ' Integer part: "0" alone, or a nonzero digit followed by digits.
    If c = 48 Then          ' "0"
        r.pos = r.pos + 1
        c = JR_CodeAt(r)
    ElseIf c >= 49 And c <= 57 Then
        Do
            r.pos = r.pos + 1
            c = JR_CodeAt(r)
        Loop While c >= 48 And c <= 57
    Else
        Err.Raise vbObjectError + 710, ERR_SRC, "Invalid number at pos " & r.pos
    End If

    Dim isIntegral As Boolean
    isIntegral = True

    If c = 46 Then          ' "."
        isIntegral = False
        r.pos = r.pos + 1
        c = JR_CodeAt(r)
        If Not (c >= 48 And c <= 57) Then
            Err.Raise vbObjectError + 711, ERR_SRC, "Invalid fractional part"
        End If
        Do
            r.pos = r.pos + 1
            c = JR_CodeAt(r)
        Loop While c >= 48 And c <= 57
    End If

    If c = 101 Or c = 69 Then   ' "e" / "E"
        isIntegral = False
        r.pos = r.pos + 1
        c = JR_CodeAt(r)
        If c = 43 Or c = 45 Then    ' "+" / "-"
            r.pos = r.pos + 1
            c = JR_CodeAt(r)
        End If
        If Not (c >= 48 And c <= 57) Then
            Err.Raise vbObjectError + 712, ERR_SRC, "Invalid exponent"
        End If
        Do
            r.pos = r.pos + 1
            c = JR_CodeAt(r)
        Loop While c >= 48 And c <= 57
    End If

    Dim numText As String
    numText = Mid$(r.text, startPos, r.pos - startPos)

    If isIntegral Then
        On Error Resume Next
        Dim asLong As Long
        asLong = CLng(numText)
        If Err.Number = 0 Then
            JR_ReadNumber = asLong
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    End If

    JR_ReadNumber = CDbl(numText)
End Function

' =============================================================================
' Strings
' =============================================================================

' Reads a JSON string starting at the opening quote.
'
' Strategy: find the candidate closing quote with InStr, then search for
' escapes ONLY within that bounded chunk. Bounding matters: probing the whole
' document for a backslash would make escape-free documents (the common case)
' quadratic, because every string token would rescan to the end of the text.
' A string with no escapes is returned with a single Mid$ and never touches
' the text builder.
Private Function JR_ReadJsonString(ByRef r As JsonReader) As String
    JR_SkipWs r
    JR_ExpectChar r, 34, """"

    Dim qPos As Long
    qPos = InStr(r.pos, r.text, """", vbBinaryCompare)
    If qPos = 0 Then
        Err.Raise vbObjectError + 523, ERR_SRC, "Unterminated string"
    End If

    Dim chunkStart As Long
    chunkStart = r.pos

    Dim chunk As String
    chunk = Mid$(r.text, chunkStart, qPos - chunkStart)

    Dim relB As Long
    relB = InStr(1, chunk, "\", vbBinaryCompare)

    ' Fast path: no escapes before the closing quote.
    If relB = 0 Then
        JR_ValidateNoControlChars chunk, chunkStart
        r.pos = qPos + 1
        JR_ReadJsonString = chunk
        Exit Function
    End If

    Dim sb As JsonTextBuilder
    JsonSB_Init sb, Len(chunk) + 16

    ' cs is the cursor within chunk (1-based). The quote candidate can only
    ' move when it turns out to be escaped (\"), because chunk never contains
    ' an unescaped quote.
    Dim cs As Long
    cs = 1

    Do
        ' Clean run before the escape.
        If relB > cs Then
            Dim runText As String
            runText = Mid$(chunk, cs, relB - cs)
            JR_ValidateNoControlChars runText, chunkStart + cs - 1
            JsonSB_Append sb, runText
        End If

        If relB = Len(chunk) Then
            ' The escape character is the quote candidate itself: an escaped
            ' quote. Append it and extend the chunk past it.
            JsonSB_Append sb, """"

            chunkStart = qPos + 1
            qPos = InStr(chunkStart, r.text, """", vbBinaryCompare)
            If qPos = 0 Then
                Err.Raise vbObjectError + 523, ERR_SRC, "Unterminated string"
            End If

            chunk = Mid$(r.text, chunkStart, qPos - chunkStart)
            cs = 1
        Else
            Dim esc As String
            esc = Mid$(chunk, relB + 1, 1)

            Select Case esc
                Case "\":  JsonSB_Append sb, "\": cs = relB + 2
                Case "/":  JsonSB_Append sb, "/": cs = relB + 2
                Case "b":  JsonSB_Append sb, Chr$(8): cs = relB + 2
                Case "f":  JsonSB_Append sb, Chr$(12): cs = relB + 2
                Case "n":  JsonSB_Append sb, vbLf: cs = relB + 2
                Case "r":  JsonSB_Append sb, vbCr: cs = relB + 2
                Case "t":  JsonSB_Append sb, vbTab: cs = relB + 2
                Case "u"
                    ' Hex digits (and a possible low-surrogate "\uXXXX") can
                    ' never contain a quote, so they are inside chunk. Read
                    ' via the absolute cursor, then map back into the chunk.
                    r.pos = chunkStart + relB + 1
                    JsonSB_Append sb, JR_ReadUnicodeEscape(r)
                    cs = r.pos - chunkStart + 1
                Case Else
                    Err.Raise vbObjectError + 522, ERR_SRC, _
                        "Invalid escape '\" & esc & "' at pos " & (chunkStart + relB)
            End Select
        End If

        relB = InStr(cs, chunk, "\", vbBinaryCompare)

        If relB = 0 Then
            If cs <= Len(chunk) Then
                Dim tailText As String
                tailText = Mid$(chunk, cs)
                JR_ValidateNoControlChars tailText, chunkStart + cs - 1
                JsonSB_Append sb, tailText
            End If
            r.pos = qPos + 1
            JR_ReadJsonString = JsonSB_Text(sb)
            Exit Function
        End If
    Loop
End Function

' JSON forbids unescaped control characters (U+0000..U+001F) inside strings.
' The chunk is scanned as UTF-16 byte pairs: a control character is a low
' byte under 32 with a zero high byte.
Private Sub JR_ValidateNoControlChars(ByRef chunk As String, ByVal chunkStartPos As Long)
    Dim n As Long
    n = Len(chunk)
    If n = 0 Then Exit Sub

    Dim b() As Byte
    b = chunk

    Dim i As Long
    For i = 0 To 2 * n - 2 Step 2
        If b(i) < 32 Then
            If b(i + 1) = 0 Then
                Err.Raise vbObjectError + 526, ERR_SRC, _
                    "Unescaped control character in string at pos " & (chunkStartPos + (i \ 2))
            End If
        End If
    Next i
End Sub

' Reads the XXXX after "\u" (cursor already past the "u"), handling UTF-16
' surrogate pairs: a high surrogate must be followed by "\u" + low surrogate.
Private Function JR_ReadUnicodeEscape(ByRef r As JsonReader) As String
    Dim u1 As Long
    u1 = JR_ReadHex4(r)

    If u1 >= &HD800& And u1 <= &HDBFF& Then
        If Mid$(r.text, r.pos, 2) <> "\u" Then
            Err.Raise vbObjectError + 527, ERR_SRC, "Invalid surrogate pair (expected \u)"
        End If
        r.pos = r.pos + 2

        Dim u2 As Long
        u2 = JR_ReadHex4(r)

        If u2 < &HDC00& Or u2 > &HDFFF& Then
            Err.Raise vbObjectError + 527, ERR_SRC, "Invalid surrogate pair (low surrogate out of range)"
        End If

        JR_ReadUnicodeEscape = ChrW$(u1) & ChrW$(u2)
        Exit Function
    End If

    If u1 >= &HDC00& And u1 <= &HDFFF& Then
        Err.Raise vbObjectError + 527, ERR_SRC, "Invalid surrogate pair (unexpected low surrogate)"
    End If

    JR_ReadUnicodeEscape = ChrW$(u1)
End Function

' Reads exactly four hex digits and returns their value (0..65535).
Private Function JR_ReadHex4(ByRef r As JsonReader) As Long
    If r.pos + 3 > Len(r.text) Then
        Err.Raise vbObjectError + 524, ERR_SRC, "Incomplete \uXXXX escape"
    End If

    Dim v As Long
    Dim i As Long
    For i = 0 To 3
        Dim off As Long
        off = (r.pos + i - 1) * 2

        Dim c As Long
        c = r.bytes(off) + r.bytes(off + 1) * 256&

        Select Case c
            Case 48 To 57       ' 0-9
                v = v * 16 + (c - 48)
            Case 65 To 70       ' A-F
                v = v * 16 + (c - 55)
            Case 97 To 102      ' a-f
                v = v * 16 + (c - 87)
            Case Else
                Err.Raise vbObjectError + 524, ERR_SRC, "Invalid \uXXXX escape"
        End Select
    Next i

    r.pos = r.pos + 4
    JR_ReadHex4 = v
End Function

' =============================================================================
' Arrays and objects
' =============================================================================

Private Function JR_ReadArray(ByRef r As JsonReader) As Collection
    JR_SkipWs r
    JR_ExpectChar r, 91, "["

    Dim result As New Collection
    JR_SkipWs r

    If JR_CodeAt(r) = 93 Then   ' "]"
        r.pos = r.pos + 1
        Set JR_ReadArray = result
        Exit Function
    End If

    Do
        Dim value As Variant
        Json_ReadValue r, value
        result.Add value

        JR_SkipWs r

        Select Case JR_CodeAt(r)
            Case 44             ' ","
                r.pos = r.pos + 1
            Case 93             ' "]"
                r.pos = r.pos + 1
                Exit Do
            Case Else
                Err.Raise vbObjectError + 730, ERR_SRC, "Expected ',' or ']' at pos " & r.pos
        End Select
    Loop

    Set JR_ReadArray = result
End Function

Private Function JR_ReadObject(ByRef r As JsonReader) As Collection
    JR_SkipWs r
    JR_ExpectChar r, 123, "{"

    Dim obj As New Collection
    obj.Add JSON_TAG_OBJECT

    JR_SkipWs r

    If JR_CodeAt(r) = 125 Then  ' "}"
        r.pos = r.pos + 1
        Set JR_ReadObject = obj
        Exit Function
    End If

    Do
        Dim key As String
        key = JR_ReadJsonString(r)

        JR_SkipWs r
        JR_ExpectChar r, 58, ":"

        Dim value As Variant
        Json_ReadValue r, value

        Dim vv As Variant
        If IsObject(value) Then
            Set vv = value
        Else
            vv = value
        End If

        ' Pairs must be built with Array(...): a fixed-size local array would
        ' be reused across iterations and alias every pair to the same data.
        obj.Add Array(key, vv)

        JR_SkipWs r

        Select Case JR_CodeAt(r)
            Case 44             ' ","
                r.pos = r.pos + 1
            Case 125            ' "}"
                r.pos = r.pos + 1
                Exit Do
            Case Else
                Err.Raise vbObjectError + 760, ERR_SRC, "Expected ',' or '}' at pos " & r.pos
        End Select
    Loop

    Set JR_ReadObject = obj
End Function
