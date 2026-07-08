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
' Json_TryParseTableStream (internal) is an alternative sink over the same
' reader: it streams an array-of-objects, at the document root or nested
' under a simple tableRoot such as "$.data.items", directly into a 2D array
' plus header index for Excel ingestion, skipping model construction.
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
' read character codes without allocating. textLen caches Len(text), which
' the scanning loops consult constantly.
Private Type JsonReader
    text As String
    bytes() As Byte
    pos As Long
    textLen As Long
End Type

' Row-schema cache for the streaming table sink. Arrays-of-objects almost
' always repeat the same member sequence row after row, so each top-level
' key's raw bytes are recorded once (concatenated in keyData) and later rows
' match keys with a typed byte comparison: no string allocation, no path
' escaping, no hash lookup on the hit path. Entries store the escaped path
' and a lazily resolved column; the value's shape is dispatched per row, so
' a key whose value alternates between null and an object stays cached. A
' key MISMATCH truncates the sequence at that position and re-records
' (adaptive; mixed schemas degrade to the plain path, never below it).
' Only escape-free keys are recorded: for them, raw text equals the decoded
' key, and a byte-equal region can contain no quote or backslash, so the
' character after it is guaranteed to be the closing quote.
Private Type JsonRowSeq
    keyData() As Byte    ' concatenated raw key bytes for all entries
    keyOff() As Long     ' byte offset of each entry's key within keyData
    keyLens() As Long    ' character length of each raw key
    cols() As Long       ' resolved column for cell writes (0 = not yet)
    paths() As String    ' escaped path segment (also the descend prefix)
    keyDataUsed As Long
    count As Long
    pos As Long          ' entries matched so far in the current row
    allocated As Boolean ' arrays dimensioned (count alone can drop to 0)
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

' Non-raising parse. Returns True and fills outValue with the model on success
' (outError is set to ""); returns False on any malformed input, with outValue
' set to Null and outError holding the position-aware reason (the same message
' Json_Parse would have raised, e.g. "Invalid number at pos 12").
'
' This is the defensive companion to Json_ParseInto, for JSON of uncertain
' provenance: an API response that might be an error page, pasted user input, or
' a file that may not be JSON. The caller branches on the return value instead
' of installing an error handler.
Public Function Json_TryParse( _
    ByVal jsonText As String, _
    ByRef outValue As Variant, _
    Optional ByRef outError As String _
) As Boolean

    On Error GoTo Fail

    Json_ParseInto jsonText, outValue
    outError = vbNullString
    Json_TryParse = True
    Exit Function

Fail:
    ' Discard any partial value the parser assigned before it raised.
    outValue = Null
    outError = Err.Description
    Json_TryParse = False
End Function

' =============================================================================
' Reader primitives
' =============================================================================

Private Sub JR_Init(ByRef r As JsonReader, ByVal jsonText As String)
    r.text = jsonText
    r.pos = 1
    r.textLen = Len(jsonText)
    If r.textLen > 0 Then
        r.bytes = jsonText
    End If
End Sub

Private Function JR_Eof(ByRef r As JsonReader) As Boolean
    JR_Eof = (r.pos > r.textLen)
End Function

' Character code (0..65535) at the cursor, or -1 at end of input. Unlike
' AscW this never goes negative for high code points, so range checks are
' straightforward.
Private Function JR_CodeAt(ByRef r As JsonReader) As Long
    If r.pos > r.textLen Then
        JR_CodeAt = -1
    Else
        Dim off As Long
        off = (r.pos - 1) * 2
        JR_CodeAt = r.bytes(off) + r.bytes(off + 1) * 256&
    End If
End Function

Private Sub JR_SkipWs(ByRef r As JsonReader)
    Dim L As Long
    L = r.textLen

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
    If r.pos <= r.textLen Then
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
            If r.pos <= r.textLen Then ch = Mid$(r.text, r.pos, 1) Else ch = vbNullString
            Err.Raise vbObjectError + 701, ERR_SRC, _
                "Unexpected token '" & ch & "' at pos " & r.pos
    End Select
End Sub

' =============================================================================
' Numbers
' =============================================================================

' Reads a strict JSON number. Integers that fit a Long come back as Long;
' everything else (fractions, exponents, overflow) comes back as Double.
'
' Integers of up to nine digits (the overwhelming majority - ids, counts)
' are accumulated directly during the scan; only longer integers and
' non-integers pay for the substring + CLng/CDbl conversion.
Private Function JR_ReadNumber(ByRef r As JsonReader) As Variant
    JR_SkipWs r

    Dim startPos As Long
    startPos = r.pos

    Dim negative As Boolean

    Dim c As Long
    c = JR_CodeAt(r)
    If c = 45 Then          ' "-"
        negative = True
        r.pos = r.pos + 1
        c = JR_CodeAt(r)
    End If

    Dim acc As Long
    Dim digitCount As Long

    ' Integer part: "0" alone, or a nonzero digit followed by digits.
    If c = 48 Then          ' "0"
        digitCount = 1
        r.pos = r.pos + 1
        c = JR_CodeAt(r)
    ElseIf c >= 49 And c <= 57 Then
        Do
            ' Nine accumulated digits max 999,999,999: never overflows.
            If digitCount < 9 Then
                acc = acc * 10 + (c - 48)
            End If
            digitCount = digitCount + 1

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

    If isIntegral Then
        If digitCount <= 9 Then
            ' Exact value already accumulated: no substring, no CLng.
            If negative Then
                JR_ReadNumber = -acc
            Else
                JR_ReadNumber = acc
            End If
            Exit Function
        End If
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
        JR_ValidateRange r, chunkStart, qPos - 1
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
            JR_ValidateRange r, chunkStart + cs - 1, chunkStart + relB - 2
            JsonSB_Append sb, Mid$(chunk, cs, relB - cs)
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
                JR_ValidateRange r, chunkStart + cs - 1, qPos - 1
                JsonSB_Append sb, Mid$(chunk, cs)
            End If
            r.pos = qPos + 1
            JR_ReadJsonString = JsonSB_Text(sb)
            Exit Function
        End If
    Loop
End Function

' JSON forbids unescaped control characters (U+0000..U+001F) inside strings.
' Validates positions fromPos..toPos directly against the reader's byte
' snapshot (no substring copy): a control character is a low byte under 32
' with a zero high byte.
Private Sub JR_ValidateRange(ByRef r As JsonReader, ByVal fromPos As Long, ByVal toPos As Long)
    Dim i As Long
    For i = fromPos To toPos
        If r.bytes((i - 1) * 2) < 32 Then
            If r.bytes((i - 1) * 2 + 1) = 0 Then
                Err.Raise vbObjectError + 526, ERR_SRC, _
                    "Unescaped control character in string at pos " & i
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
    If r.pos + 3 > r.textLen Then
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

        ' Pairs must be built with Array(...): a fixed-size local array would
        ' be reused across iterations and alias every pair to the same data.
        ' Array() copies its Variant arguments itself, object references
        ' included, so no Set/Let staging is needed here.
        obj.Add Array(key, value)

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

' =============================================================================
' Streaming table sink (internal: used by Excel_UpsertListObjectFromJsonAtRoot)
' =============================================================================

' Stream a JSON array-of-objects directly into a 2D Variant array plus a
' header index, without building the Collection model. The array may sit at
' the document root (tableRoot "$") or nested under simple member steps
' (tableRoot "$.data.items"; a key containing a literal dot escapes it as
' "$.a\.b.items"): the reader descends to the table root, moving
' past sibling members with validating skip scanners that build nothing and,
' for the usual escape-free text, allocate nothing. Cell values, header
' discovery order, duplicate-key overwrites, and the validations the model
' path performs are reproduced exactly:
'
'   - nested objects contribute dotted, escaped column paths
'   - nested arrays are fully parsed through the model (so their content is
'     validated) and either stringified into the cell or discarded,
'     depending on nonTableArraysAsJson
'   - a non-object element raises 1163 with the caller-supplied source
'   - sibling members before and after the table array are validated, and
'     text after the document value raises 700, so a malformed document is
'     rejected exactly as if the model path had parsed all of it
'
' Returns False WITHOUT raising when the streamed shape does not apply: the
' value at the table root is not an array, a descent step hits a non-object
' or an absent key, or tableRoot uses features beyond simple member steps
' (bracket indices). The caller then falls back to the model path, which
' owns those error semantics (1130/1160/1162).
Public Function Json_TryParseTableStream( _
    ByVal jsonText As String, _
    ByVal nonTableArraysAsJson As Boolean, _
    ByRef headerIdx As JsonStringIndex, _
    ByRef outData As Variant, _
    ByRef outRowCount As Long, _
    ByVal rowErrorSource As String, _
    ByVal tableRoot As String _
) As Boolean

    Json_TryParseTableStream = False
    outRowCount = 0

    ' Path analysis: "$" streams at the document root; "$.a.b" descends one
    ' member step per segment. Segments are tokenized on unescaped dots and
    ' decoded, so "$.a\.b" addresses the literal key "a.b" (the column-header
    ' escape convention, resolved the same way by Json_TryResolvePath on the
    ' model path). Bracket paths and empty segments decline to the model
    ' path. Trimming matches Json_TryResolvePath, so both paths address the
    ' same root for the same input.
    Dim segs() As String
    Dim segCount As Long

    Dim root As String
    root = Trim$(tableRoot)

    If root <> "$" Then
        If Left$(root, 2) <> "$." Then Exit Function
        If InStr(1, root, "[", vbBinaryCompare) > 0 Then Exit Function

        Dim tokens As Collection
        Set tokens = Json_TokenizePath(Mid$(root, 3))
        segCount = tokens.count

        If segCount > 0 Then
            ReDim segs(0 To segCount - 1) As String

            Dim si As Long
            si = 0

            Dim tok As Variant
            For Each tok In tokens
                If Len(CStr(tok)) = 0 Then Exit Function
                segs(si) = Json_UnescapePathSegment(CStr(tok))
                si = si + 1
            Next tok
        End If
    End If

    Dim r As JsonReader
    JR_Init r, jsonText

    If segCount > 0 Then
        If Not JR_DescendToRoot(r, segs) Then Exit Function
    End If

    JR_SkipWs r
    If JR_CodeAt(r) <> 91 Then Exit Function    ' table root is not an array

    ' Committed to streaming. Pre-count the top-level elements so the row
    ' dimension (which ReDim Preserve cannot grow) is allocated exactly.
    Dim rowCap As Long
    rowCap = JR_CountTopLevelElements(r)

    r.pos = r.pos + 1
    JR_SkipWs r

    If JR_CodeAt(r) = 93 Then                   ' "]": empty array
        r.pos = r.pos + 1
        JR_FinishDocument r, segCount
        Json_TryParseTableStream = True
        Exit Function
    End If

    If rowCap < 1 Then rowCap = 1
    ReDim outData(1 To rowCap, 1 To 8)          ' columns grow on demand

    Dim seq As JsonRowSeq

    Do
        JR_SkipWs r

        If JR_CodeAt(r) <> 123 Then
            Err.Raise vbObjectError + 1163, rowErrorSource, _
                "Array element at index " & outRowCount & " is not an object for root: " & tableRoot
        End If

        outRowCount = outRowCount + 1
        If outRowCount > UBound(outData, 1) Then JR_GrowRows outData

        JR_StreamRowTop r, seq, nonTableArraysAsJson, headerIdx, outData, outRowCount

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

    JR_FinishDocument r, segCount
    Json_TryParseTableStream = True
End Function

' Stream one top-level row object using the row-schema cache. The cursor is
' at the row's "{" on entry. Nested objects descend through the uncached
' JR_StreamObjectRow with the cached path as prefix.
Private Sub JR_StreamRowTop( _
    ByRef r As JsonReader, _
    ByRef seq As JsonRowSeq, _
    ByVal nonTableArraysAsJson As Boolean, _
    ByRef headerIdx As JsonStringIndex, _
    ByRef outData As Variant, _
    ByVal rowNumber As Long _
)
    r.pos = r.pos + 1                           ' consume "{"
    JR_SkipWs r

    If JR_CodeAt(r) = 125 Then                  ' "}": empty object
        r.pos = r.pos + 1
        Exit Sub
    End If

    seq.pos = 0

    Do
        Dim matched As Boolean
        matched = False

        Dim idx As Long
        Dim c As Long

        ' ---- Cache-hit attempt: raw byte compare against the expected key ----
        If seq.pos < seq.count Then
            If JR_CodeAt(r) = 34 Then           ' at an opening quote
                idx = seq.pos + 1

                Dim keyStart As Long
                keyStart = r.pos + 1

                Dim klen As Long
                klen = seq.keyLens(idx)

                If keyStart + klen <= r.textLen Then
                    Dim off As Long
                    off = (keyStart - 1) * 2

                    ' Closing quote exactly where the cached key ends?
                    If r.bytes(off + klen * 2) = 34 Then
                        If r.bytes(off + klen * 2 + 1) = 0 Then
                            Dim ko As Long
                            ko = seq.keyOff(idx)

                            matched = True
                            Dim b As Long
                            For b = 0 To klen * 2 - 1
                                If r.bytes(off + b) <> seq.keyData(ko + b) Then
                                    matched = False
                                    Exit For
                                End If
                            Next b
                        End If
                    End If
                End If
            End If
        End If

        If matched Then
            r.pos = keyStart + klen + 1         ' past the closing quote
            JR_ExpectChar r, 58, ":"
            JR_SkipWs r

            ' Dispatch on the value's ACTUAL shape; the cached path and
            ' column apply regardless of what shape this key had before.
            c = JR_CodeAt(r)

            If c = 123 Then                     ' nested object
                JR_StreamObjectRow r, seq.paths(idx), nonTableArraysAsJson, headerIdx, outData, rowNumber

            ElseIf c = 91 Then                  ' nested array
                Dim av As Variant
                Json_ReadValue r, av

                If nonTableArraysAsJson Then
                    If seq.cols(idx) = 0 Then
                        seq.cols(idx) = JsonIdx_Ensure(headerIdx, seq.paths(idx))
                        Json_Grow2DCols outData, seq.cols(idx)
                    End If
                    outData(rowNumber, seq.cols(idx)) = Json_Stringify(av)
                End If

            Else                                ' primitive
                Dim v As Variant
                Json_ReadValue r, v

                If seq.cols(idx) = 0 Then
                    seq.cols(idx) = JsonIdx_Ensure(headerIdx, seq.paths(idx))
                    Json_Grow2DCols outData, seq.cols(idx)
                End If
                outData(rowNumber, seq.cols(idx)) = v
            End If

            seq.pos = seq.pos + 1
        Else
            ' ---- Plain path: decode, resolve, and (re)record the schema ----
            Dim posBeforeKey As Long
            posBeforeKey = r.pos

            Dim key As String
            key = JR_ReadJsonString(r)

            ' Raw length: cursor is one past the closing quote.
            Dim rawLen As Long
            rawLen = r.pos - posBeforeKey - 2

            JR_ExpectChar r, 58, ":"
            JR_SkipWs r

            Dim path As String
            path = Json_EscapePathSegment(key)

            ' A key that diverges from the recorded sequence invalidates the
            ' remainder; escape-free keys re-record at this position.
            seq.count = seq.pos

            Dim recorded As Long
            recorded = 0
            If rawLen = Len(key) And rawLen > 0 Then
                recorded = JR_SeqAppend(seq, r, posBeforeKey + 1, rawLen, path)
            End If

            c = JR_CodeAt(r)

            If c = 123 Then                     ' nested object
                JR_StreamObjectRow r, path, nonTableArraysAsJson, headerIdx, outData, rowNumber

            ElseIf c = 91 Then                  ' nested array
                Dim av2 As Variant
                Json_ReadValue r, av2

                If nonTableArraysAsJson Then
                    Dim colA As Long
                    colA = JsonIdx_Ensure(headerIdx, path)
                    Json_Grow2DCols outData, colA
                    outData(rowNumber, colA) = Json_Stringify(av2)
                    If recorded > 0 Then seq.cols(recorded) = colA
                End If

            Else                                ' primitive
                Dim v2 As Variant
                Json_ReadValue r, v2

                Dim colP As Long
                colP = JsonIdx_Ensure(headerIdx, path)
                Json_Grow2DCols outData, colP
                outData(rowNumber, colP) = v2
                If recorded > 0 Then seq.cols(recorded) = colP
            End If

            If recorded > 0 Then seq.pos = seq.count
        End If

        JR_SkipWs r

        Select Case JR_CodeAt(r)
            Case 44             ' ","
                r.pos = r.pos + 1
                JR_SkipWs r
            Case 125            ' "}"
                r.pos = r.pos + 1
                Exit Do
            Case Else
                Err.Raise vbObjectError + 760, ERR_SRC, "Expected ',' or '}' at pos " & r.pos
        End Select
    Loop
End Sub

' Append one entry to the row-schema cache, copying the raw key bytes into
' the flat buffer. Returns the new entry's 1-based index.
Private Function JR_SeqAppend( _
    ByRef seq As JsonRowSeq, _
    ByRef r As JsonReader, _
    ByVal keyStart As Long, _
    ByVal klen As Long, _
    ByRef path As String _
) As Long

    Dim n As Long
    n = seq.count + 1

    If Not seq.allocated Then
        ReDim seq.keyOff(1 To 8) As Long
        ReDim seq.keyLens(1 To 8) As Long
        ReDim seq.cols(1 To 8) As Long
        ReDim seq.paths(1 To 8) As String
        ReDim seq.keyData(0 To 255) As Byte
        seq.keyDataUsed = 0
        seq.allocated = True
    ElseIf n > UBound(seq.keyOff) Then
        ReDim Preserve seq.keyOff(1 To UBound(seq.keyOff) * 2) As Long
        ReDim Preserve seq.keyLens(1 To UBound(seq.keyLens) * 2) As Long
        ReDim Preserve seq.cols(1 To UBound(seq.cols) * 2) As Long
        ReDim Preserve seq.paths(1 To UBound(seq.paths) * 2) As String
    End If

    Dim needBytes As Long
    needBytes = klen * 2

    ' Truncation resets count but keyDataUsed keeps growing; compact only by
    ' appending (offsets of live entries never move).
    If seq.keyDataUsed + needBytes > UBound(seq.keyData) + 1 Then
        Dim newCap As Long
        newCap = (UBound(seq.keyData) + 1) * 2
        Do While seq.keyDataUsed + needBytes > newCap
            newCap = newCap * 2
        Loop
        ReDim Preserve seq.keyData(0 To newCap - 1) As Byte
    End If

    Dim srcOff As Long
    srcOff = (keyStart - 1) * 2

    Dim i As Long
    For i = 0 To needBytes - 1
        seq.keyData(seq.keyDataUsed + i) = r.bytes(srcOff + i)
    Next i

    seq.keyOff(n) = seq.keyDataUsed
    seq.keyLens(n) = klen
    seq.cols(n) = 0
    seq.paths(n) = path
    seq.keyDataUsed = seq.keyDataUsed + needBytes

    seq.count = n
    JR_SeqAppend = n
End Function

' Stream one object's members into row rowNumber, recursing into nested
' objects with dotted paths (mirrors Json_RowValueFill's rules).
Private Sub JR_StreamObjectRow( _
    ByRef r As JsonReader, _
    ByVal prefix As String, _
    ByVal nonTableArraysAsJson As Boolean, _
    ByRef headerIdx As JsonStringIndex, _
    ByRef outData As Variant, _
    ByVal rowNumber As Long _
)
    r.pos = r.pos + 1                           ' consume "{"
    JR_SkipWs r

    If JR_CodeAt(r) = 125 Then                  ' "}": empty object
        r.pos = r.pos + 1
        Exit Sub
    End If

    Do
        Dim key As String
        key = JR_ReadJsonString(r)

        JR_SkipWs r
        JR_ExpectChar r, 58, ":"

        Dim path As String
        If Len(prefix) = 0 Then
            path = Json_EscapePathSegment(key)
        Else
            path = prefix & "." & Json_EscapePathSegment(key)
        End If

        JR_SkipWs r

        Select Case JR_CodeAt(r)
            Case 123            ' nested object: dotted columns
                JR_StreamObjectRow r, path, nonTableArraysAsJson, headerIdx, outData, rowNumber

            Case 91             ' nested array: parse (validates) then keep or drop
                Dim av As Variant
                Json_ReadValue r, av

                If nonTableArraysAsJson Then
                    JR_WriteCell headerIdx, outData, rowNumber, path, Json_Stringify(av)
                End If

            Case Else           ' primitive (or a syntax error, raised here)
                Dim v As Variant
                Json_ReadValue r, v
                JR_WriteCell headerIdx, outData, rowNumber, path, v
        End Select

        JR_SkipWs r

        Select Case JR_CodeAt(r)
            Case 44             ' ","
                r.pos = r.pos + 1
                JR_SkipWs r
            Case 125            ' "}"
                r.pos = r.pos + 1
                Exit Do
            Case Else
                Err.Raise vbObjectError + 760, ERR_SRC, "Expected ',' or '}' at pos " & r.pos
        End Select
    Loop
End Sub

Private Sub JR_WriteCell( _
    ByRef headerIdx As JsonStringIndex, _
    ByRef outData As Variant, _
    ByVal rowNumber As Long, _
    ByRef path As String, _
    ByRef v As Variant _
)
    Dim col As Long
    col = JsonIdx_Ensure(headerIdx, path)

    Json_Grow2DCols outData, col
    outData(rowNumber, col) = v
End Sub

' Count elements of the array starting at the cursor's "[" (cursor is not
' moved): commas at depth 1 plus one, tracked with a string-aware scan.
' Exact for well-formed JSON; malformed input raises during the real parse,
' and JR_GrowRows guards the fill regardless.
Private Function JR_CountTopLevelElements(ByRef r As JsonReader) As Long
    Dim L As Long
    L = r.textLen

    Dim depth As Long
    Dim count As Long
    Dim inString As Boolean

    Dim i As Long
    i = r.pos

    Do While i <= L
        Dim c As Long
        c = r.bytes((i - 1) * 2) + r.bytes((i - 1) * 2 + 1) * 256&

        If inString Then
            If c = 92 Then          ' "\": skip the escaped character
                i = i + 1
            ElseIf c = 34 Then      ' closing quote
                inString = False
            End If
        Else
            Select Case c
                Case 34             ' opening quote
                    inString = True
                Case 91, 123        ' "[" / "{"
                    depth = depth + 1
                Case 93, 125        ' "]" / "}"
                    depth = depth - 1
                    If depth = 0 Then Exit Do
                Case 44             ' ","
                    If depth = 1 Then count = count + 1
            End Select
        End If

        i = i + 1
    Loop

    JR_CountTopLevelElements = count + 1
End Function

' Row-dimension growth safety net: ReDim Preserve cannot grow the first
' dimension, so grow by allocating double and copying. Unreachable for
' well-formed input (rows are pre-counted exactly).
Private Sub JR_GrowRows(ByRef outData As Variant)
    Dim oldRows As Long
    Dim cols As Long
    oldRows = UBound(outData, 1)
    cols = UBound(outData, 2)

    Dim newData As Variant
    ReDim newData(1 To oldRows * 2, 1 To cols)

    Dim rr As Long
    Dim cc As Long
    For rr = 1 To oldRows
        For cc = 1 To cols
            newData(rr, cc) = outData(rr, cc)
        Next cc
    Next rr

    outData = newData
End Sub

Private Sub JR_CheckTrailing(ByRef r As JsonReader)
    JR_SkipWs r
    If Not JR_Eof(r) Then
        Err.Raise vbObjectError + 700, ERR_SRC, "Unexpected trailing characters at pos " & r.pos
    End If
End Sub

' =============================================================================
' Nested-root descent and validating skips
'
' Support for streaming a table nested under a tableRoot such as
' "$.data.items". The descent walks member steps, skipping sibling values;
' after the table array closes, JR_FinishDocument consumes the rest of each
' enclosing object. Skipped text is validated with the same rules and error
' numbers as the reading path, so a malformed document is rejected whether
' its defect sits inside the table or beside it. Keys compare decoded and
' case-sensitive, first match wins: the same rules Json_TryObjGet applies
' when the model path resolves the identical tableRoot.
' =============================================================================

' Walk the reader down one object level per segment. On success the cursor
' rests at the matched member's value and True returns. Returns False (the
' stream then declines to the model path) when a step's container is not an
' object or the key is absent. Malformed text raises the reader's usual
' errors.
Private Function JR_DescendToRoot(ByRef r As JsonReader, ByRef segs() As String) As Boolean
    Dim si As Long
    For si = 0 To UBound(segs)
        JR_SkipWs r
        If JR_CodeAt(r) <> 123 Then Exit Function   ' not an object: decline
        r.pos = r.pos + 1
        JR_SkipWs r

        If JR_CodeAt(r) = 125 Then Exit Function    ' empty object: key absent

        Dim found As Boolean
        found = False

        Do
            Dim key As String
            key = JR_ReadJsonString(r)
            JR_ExpectChar r, 58, ":"

            If StrComp(key, segs(si), vbBinaryCompare) = 0 Then
                found = True
                Exit Do
            End If

            JR_SkipValue r
            JR_SkipWs r

            Select Case JR_CodeAt(r)
                Case 44             ' ","
                    r.pos = r.pos + 1
                Case 125            ' "}": key not present at this level
                    Exit Do
                Case Else
                    Err.Raise vbObjectError + 760, ERR_SRC, "Expected ',' or '}' at pos " & r.pos
            End Select
        Loop

        If Not found Then Exit Function
    Next si

    JR_DescendToRoot = True
End Function

' After the streamed array closes, consume the remainder of each enclosing
' object (validating any sibling members that follow the table), then
' require end of input. With levels = 0 this is exactly the old root-level
' trailing check.
Private Sub JR_FinishDocument(ByRef r As JsonReader, ByVal levels As Long)
    Dim lvl As Long
    For lvl = 1 To levels
        Do
            JR_SkipWs r

            Select Case JR_CodeAt(r)
                Case 44             ' ",": another sibling member
                    r.pos = r.pos + 1
                    JR_SkipJsonString r
                    JR_ExpectChar r, 58, ":"
                    JR_SkipValue r
                Case 125            ' "}"
                    r.pos = r.pos + 1
                    Exit Do
                Case Else
                    Err.Raise vbObjectError + 760, ERR_SRC, "Expected ',' or '}' at pos " & r.pos
            End Select
        Loop
    Next lvl

    JR_CheckTrailing r
End Sub

' Skip one JSON value, accepting and rejecting exactly what Json_ReadValue
' accepts and rejects, but building nothing. Numbers reuse JR_ReadNumber
' (allocation-free for short integers); containers recurse.
Private Sub JR_SkipValue(ByRef r As JsonReader)
    JR_SkipWs r

    Select Case JR_CodeAt(r)
        Case 34                     ' quote
            JR_SkipJsonString r

        Case 116                    ' "t"
            JR_ExpectLiteral r, "true"

        Case 102                    ' "f"
            JR_ExpectLiteral r, "false"

        Case 110                    ' "n"
            JR_ExpectLiteral r, "null"

        Case 45, 48 To 57           ' "-", "0".."9"
            JR_ReadNumber r

        Case 91                     ' "["
            JR_SkipArray r

        Case 123                    ' "{"
            JR_SkipObject r

        Case Else
            Dim ch As String
            If r.pos <= r.textLen Then ch = Mid$(r.text, r.pos, 1) Else ch = vbNullString
            Err.Raise vbObjectError + 701, ERR_SRC, _
                "Unexpected token '" & ch & "' at pos " & r.pos
    End Select
End Sub

' Skip a JSON string with the same validation as JR_ReadJsonString but no
' text materialization. One byte scan bounded to the string's own extent
' both rejects unescaped control characters and probes for a backslash; the
' escape-free case (the overwhelming majority) then just advances the
' cursor. A string containing escapes falls back to the full reader with
' the result discarded, so escape and surrogate validation stay in one
' place.
Private Sub JR_SkipJsonString(ByRef r As JsonReader)
    JR_SkipWs r

    If JR_CodeAt(r) <> 34 Then
        JR_ExpectChar r, 34, """"   ' raises 520 with the reader's message
    End If

    Dim contentStart As Long
    contentStart = r.pos + 1

    Dim qPos As Long
    qPos = InStr(contentStart, r.text, """", vbBinaryCompare)
    If qPos = 0 Then
        Err.Raise vbObjectError + 523, ERR_SRC, "Unterminated string"
    End If

    Dim i As Long
    For i = contentStart To qPos - 1
        If r.bytes((i - 1) * 2 + 1) = 0 Then
            Select Case r.bytes((i - 1) * 2)
                Case 92
                    ' Escapes present: the cursor still sits on the opening
                    ' quote, so the full reader takes over from scratch.
                    JR_ReadJsonString r
                    Exit Sub
                Case 0 To 31
                    Err.Raise vbObjectError + 526, ERR_SRC, _
                        "Unescaped control character in string at pos " & i
            End Select
        End If
    Next i

    r.pos = qPos + 1
End Sub

Private Sub JR_SkipArray(ByRef r As JsonReader)
    r.pos = r.pos + 1               ' consume "[" (dispatched on it)
    JR_SkipWs r

    If JR_CodeAt(r) = 93 Then       ' "]": empty array
        r.pos = r.pos + 1
        Exit Sub
    End If

    Do
        JR_SkipValue r
        JR_SkipWs r

        Select Case JR_CodeAt(r)
            Case 44                 ' ","
                r.pos = r.pos + 1
            Case 93                 ' "]"
                r.pos = r.pos + 1
                Exit Do
            Case Else
                Err.Raise vbObjectError + 730, ERR_SRC, "Expected ',' or ']' at pos " & r.pos
        End Select
    Loop
End Sub

Private Sub JR_SkipObject(ByRef r As JsonReader)
    r.pos = r.pos + 1               ' consume "{" (dispatched on it)
    JR_SkipWs r

    If JR_CodeAt(r) = 125 Then      ' "}": empty object
        r.pos = r.pos + 1
        Exit Sub
    End If

    Do
        JR_SkipJsonString r
        JR_ExpectChar r, 58, ":"
        JR_SkipValue r
        JR_SkipWs r

        Select Case JR_CodeAt(r)
            Case 44                 ' ","
                r.pos = r.pos + 1
            Case 125                ' "}"
                r.pos = r.pos + 1
                Exit Do
            Case Else
                Err.Raise vbObjectError + 760, ERR_SRC, "Expected ',' or '}' at pos " & r.pos
        End Select
    Loop
End Sub
