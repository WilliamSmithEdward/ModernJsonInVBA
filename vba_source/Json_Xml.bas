Attribute VB_Name = "Json_Xml"
Option Explicit

' =============================================================================
' Module:      Json_Xml
' Project:     ModernJsonInVBA
'
' XML -> JSON conversion using a lightweight pure-VBA parser (no MSXML, so
' the conversion works on every Office platform).
'
' Mapping rules:
'   - element with child elements  => JSON object
'   - repeated child element names => grouped into a JSON array placed at
'                                     the first occurrence, document order
'   - text-only element            => JSON string
'   - empty / self-closing element => null
'   - mixed content                => text stored under the key "value"
'   - attributes                   => skipped
'   - CDATA                        => raw text (no entity decoding)
'
' Supported: nested elements, text nodes, self-closing tags, CDATA,
' built-in and numeric entities.
' Not supported: XML declarations, processing instructions, comments,
' DTDs, namespaces (namespace prefixes are kept as part of the name).
'
' Performance notes:
'   - JSON assembly and entity decoding use the shared text builder; text
'     runs are copied in chunks located with InStr.
'   - Repeated-name grouping is a single pass over the children (previously
'     O(n^2) name counting per element).
'
' Error numbers (all vbObjectError + n, source "XmlTextToJson"):
'   821 parser stalled (malformed)   822 unexpected end / bad CDATA
'   823 invalid XML name             824 expected '>'
' =============================================================================

Private Const ERR_SRC As String = "XmlTextToJson"

' Read an XML file and convert it via XmlTextToJson.
Public Function XmlFileToJson(ByVal filePath As String) As String
    Dim f As Integer
    f = FreeFile

    Dim txt As String
    Open filePath For Input As #f
    txt = Input$(LOF(f), f)
    Close #f

    XmlFileToJson = XmlTextToJson(txt)
End Function

' Convert raw XML text into JSON text.
Public Function XmlTextToJson(ByVal txt As String) As String
    ' Strip a UTF-8/UTF-16 BOM if present.
    If Len(txt) > 0 Then
        If AscW(Left$(txt, 1)) = &HFEFF Then
            txt = Mid$(txt, 2)
        End If
    End If

    Dim pos As Long
    pos = 1

    Xml_SkipWhitespace txt, pos

    XmlTextToJson = Xml_ParseNode(txt, pos)
End Function

' =============================================================================
' Parser
' =============================================================================

' Parse one element starting at "<" and return its JSON representation.
' Advances pos past the element's closing tag.
Private Function Xml_ParseNode(ByVal txt As String, ByRef pos As Long) As String
    Dim L As Long
    L = Len(txt)

    pos = pos + 1                    ' consume "<"
    Xml_ReadName txt, pos

    Xml_SkipAttributes txt, pos

    If Mid$(txt, pos, 2) = "/>" Then
        pos = pos + 2
        Xml_ParseNode = "null"
        Exit Function
    End If

    If Mid$(txt, pos, 1) = ">" Then
        pos = pos + 1
    Else
        Err.Raise vbObjectError + 824, ERR_SRC, "Malformed XML: expected '>'."
    End If

    ' ---- Collect children and text ----
    ' Children accumulate in plain arrays (grow-doubling): Collections would
    ' pay a linked-list walk per indexed access during grouping below.
    Dim childNames() As String
    Dim childValues() As String
    Dim childCount As Long
    childCount = 0

    Dim textBuffer As String

    Do While pos <= L
        Dim startPos As Long
        startPos = pos

        Xml_SkipWhitespace txt, pos

        If Mid$(txt, pos, 2) = "</" Then
            pos = pos + 2
            Xml_ReadName txt, pos
            If Mid$(txt, pos, 1) = ">" Then pos = pos + 1
            Exit Do
        End If

        If Mid$(txt, pos, 9) = "<![CDATA[" Then
            Dim cEnd As Long
            cEnd = InStr(pos + 9, txt, "]]>")

            If cEnd = 0 Then
                Err.Raise vbObjectError + 822, ERR_SRC, "Unterminated CDATA section."
            End If

            textBuffer = textBuffer & Mid$(txt, pos + 9, cEnd - (pos + 9))
            pos = cEnd + 3

            GoTo ContinueLoop
        End If

        If Mid$(txt, pos, 1) = "<" Then
            ' Peek the child's name without consuming, then parse it.
            Dim tmp As Long
            tmp = pos + 1

            childCount = childCount + 1
            If childCount = 1 Then
                ReDim childNames(1 To 16) As String
                ReDim childValues(1 To 16) As String
            ElseIf childCount > UBound(childNames) Then
                ReDim Preserve childNames(1 To UBound(childNames) * 2) As String
                ReDim Preserve childValues(1 To UBound(childValues) * 2) As String
            End If

            childNames(childCount) = Xml_ReadName(txt, tmp)
            childValues(childCount) = Xml_ParseNode(txt, pos)
        Else
            Dim textVal As String
            textVal = Xml_ReadText(txt, pos)

            If Len(Trim$(textVal)) > 0 Then
                textBuffer = textBuffer & textVal
            End If
        End If

ContinueLoop:
        If pos = startPos Then
            Err.Raise vbObjectError + 821, ERR_SRC, "Malformed XML: parser stalled."
        End If
    Loop

    ' ---- Text-only node => primitive ----
    If childCount = 0 Then
        If Len(textBuffer) = 0 Then
            Xml_ParseNode = "null"
        Else
            Dim tsb As JsonTextBuilder
            JsonSB_Init tsb, Len(textBuffer) + 16
            JsonSB_Append tsb, """"
            Xml_AppendEscapedJson tsb, textBuffer
            JsonSB_Append tsb, """"
            Xml_ParseNode = JsonSB_Text(tsb)
        End If
        Exit Function
    End If

    ' ---- Object node ----
    ' Group children by name in one pass: names keep first-seen order;
    ' repeated names become arrays at their first position.
    Dim nameIdx As JsonStringIndex
    JsonIdx_Init nameIdx, 16

    Dim groupItems() As Collection   ' per distinct name: child indices

    Dim j As Long
    For j = 1 To childCount
        Dim before As Long
        before = nameIdx.count

        Dim g As Long
        g = JsonIdx_Ensure(nameIdx, childNames(j))

        If nameIdx.count > before Then
            If before = 0 Then
                ReDim groupItems(1 To 16) As Collection
            ElseIf g > UBound(groupItems) Then
                ReDim Preserve groupItems(1 To UBound(groupItems) * 2) As Collection
            End If
            Set groupItems(g) = New Collection
        End If

        groupItems(g).Add j
    Next j

    Dim sb As JsonTextBuilder
    JsonSB_Init sb, 256

    JsonSB_Append sb, "{"

    Dim first As Boolean
    first = True

    ' Mixed content keeps its text under "value".
    If Len(textBuffer) > 0 Then
        JsonSB_Append sb, """value"":"""
        Xml_AppendEscapedJson sb, textBuffer
        JsonSB_Append sb, """"
        first = False
    End If

    For g = 1 To nameIdx.count
        If Not first Then JsonSB_Append sb, ","
        first = False

        JsonSB_Append sb, """"
        Xml_AppendEscapedJson sb, nameIdx.keys(g)
        JsonSB_Append sb, """:"

        Dim members As Collection
        Set members = groupItems(g)

        If members.count = 1 Then
            JsonSB_Append sb, childValues(members(1))
        Else
            JsonSB_Append sb, "["

            Dim firstMember As Boolean
            firstMember = True

            Dim m As Variant
            For Each m In members
                If firstMember Then
                    firstMember = False
                Else
                    JsonSB_Append sb, ","
                End If
                JsonSB_Append sb, childValues(CLng(m))
            Next m

            JsonSB_Append sb, "]"
        End If
    Next g

    JsonSB_Append sb, "}"

    Xml_ParseNode = JsonSB_Text(sb)
End Function

' Read an XML name at pos ([A-Za-z_:] then [A-Za-z0-9_.:-]*).
' Character-code comparisons instead of Like: Like recompiles its pattern on
' every call, which dominated profile time on element-dense documents.
Private Function Xml_ReadName(ByVal txt As String, ByRef pos As Long) As String
    Dim startPos As Long
    startPos = pos

    If pos > Len(txt) Then
        Err.Raise vbObjectError + 822, ERR_SRC, "Unexpected end while reading tag name."
    End If

    Dim c As Long
    c = AscW(Mid$(txt, pos, 1))

    ' A-Z, a-z, "_", ":"
    Select Case c
        Case 65 To 90, 97 To 122, 95, 58
        Case Else
            Err.Raise vbObjectError + 823, ERR_SRC, "Invalid start of XML name."
    End Select

    pos = pos + 1

    Do While pos <= Len(txt)
        c = AscW(Mid$(txt, pos, 1))

        ' A-Z, a-z, 0-9, "_", ".", ":", "-"
        Select Case c
            Case 65 To 90, 97 To 122, 48 To 57, 95, 46, 58, 45
                pos = pos + 1
            Case Else
                Exit Do
        End Select
    Loop

    Xml_ReadName = Mid$(txt, startPos, pos - startPos)
End Function

' Read text content up to the next "<" (or end of input) and decode entities.
Private Function Xml_ReadText(ByVal txt As String, ByRef pos As Long) As String
    Dim startPos As Long
    startPos = pos

    Dim lt As Long
    lt = InStr(pos, txt, "<", vbBinaryCompare)

    If lt = 0 Then
        pos = Len(txt) + 1
    Else
        pos = lt
    End If

    Xml_ReadText = Xml_DecodeEntities(Mid$(txt, startPos, pos - startPos))
End Function

Private Sub Xml_SkipWhitespace(ByVal txt As String, ByRef pos As Long)
    Dim L As Long
    L = Len(txt)

    Do While pos <= L
        Select Case AscW(Mid$(txt, pos, 1))
            Case 32, 13, 10, 9
                pos = pos + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

' Skip everything up to the tag-closing ">" or "/", honoring quoted
' attribute values so a ">" inside quotes does not end the tag.
Private Sub Xml_SkipAttributes(ByVal txt As String, ByRef pos As Long)
    Do While pos <= Len(txt)
        Dim ch As String
        ch = Mid$(txt, pos, 1)

        Select Case ch
            Case """", "'"
                Dim closeQuote As Long
                closeQuote = InStr(pos + 1, txt, ch)
                If closeQuote = 0 Then
                    pos = Len(txt) + 1
                    Exit Sub
                End If
                pos = closeQuote

            Case ">", "/"
                Exit Sub
        End Select

        pos = pos + 1
    Loop
End Sub

' =============================================================================
' Text handling
' =============================================================================

' Append s to the builder with JSON escaping (quote, backslash, \b \t \n
' \f \r, \uXXXX for remaining control characters).
Private Sub Xml_AppendEscapedJson(ByRef sb As JsonTextBuilder, ByRef s As String)
    Dim L As Long
    L = Len(s)

    Dim runStart As Long
    runStart = 1

    Dim i As Long
    For i = 1 To L
        Dim c As Long
        c = AscW(Mid$(s, i, 1))

        Dim escText As String
        escText = vbNullString

        Select Case c
            Case 34: escText = "\"""
            Case 92: escText = "\\"
            Case 8:  escText = "\b"
            Case 9:  escText = "\t"
            Case 10: escText = "\n"
            Case 12: escText = "\f"
            Case 13: escText = "\r"
            Case 0 To 31
                escText = "\u" & Right$("0000" & Hex$(c), 4)
        End Select

        If Len(escText) > 0 Then
            If i > runStart Then
                JsonSB_Append sb, Mid$(s, runStart, i - runStart)
            End If
            JsonSB_Append sb, escText
            runStart = i + 1
        End If
    Next i

    If runStart = 1 Then
        JsonSB_Append sb, s
    ElseIf runStart <= L Then
        JsonSB_Append sb, Mid$(s, runStart)
    End If
End Sub

' Decode XML built-in entities (&lt; &gt; &amp; &apos; &quot;) and numeric
' entities (&#NNN; / &#xHH;). Unknown entities are left literal. Text with
' no "&" at all is returned unchanged without any copying.
Private Function Xml_DecodeEntities(ByVal s As String) As String
    Dim ampPos As Long
    ampPos = InStr(1, s, "&", vbBinaryCompare)

    If ampPos = 0 Then
        Xml_DecodeEntities = s
        Exit Function
    End If

    Dim n As Long
    n = Len(s)

    Dim sb As JsonTextBuilder
    JsonSB_Init sb, n + 16

    Dim i As Long
    i = 1

    Do While ampPos > 0
        ' Copy the clean run before the "&".
        If ampPos > i Then
            JsonSB_Append sb, Mid$(s, i, ampPos - i)
        End If

        Dim semi As Long
        semi = InStr(ampPos, s, ";")

        If semi = 0 Then
            ' No terminator: keep the "&" literal and continue after it.
            JsonSB_Append sb, "&"
            i = ampPos + 1
        Else
            Dim ent As String
            ent = Mid$(s, ampPos + 1, semi - ampPos - 1)

            Select Case ent
                Case "lt":   JsonSB_Append sb, "<"
                Case "gt":   JsonSB_Append sb, ">"
                Case "amp":  JsonSB_Append sb, "&"
                Case "apos": JsonSB_Append sb, "'"
                Case "quot": JsonSB_Append sb, """"
                Case Else
                    If Left$(ent, 1) = "#" Then
                        Dim code As Long
                        If LCase$(Mid$(ent, 2, 1)) = "x" Then
                            code = CLng("&H" & Mid$(ent, 3))
                        Else
                            code = CLng(Mid$(ent, 2))
                        End If
                        JsonSB_Append sb, ChrW$(code)
                    Else
                        ' Unknown entity: keep it literal.
                        JsonSB_Append sb, "&" & ent & ";"
                    End If
            End Select

            i = semi + 1
        End If

        If i > n Then Exit Do
        ampPos = InStr(i, s, "&", vbBinaryCompare)
    Loop

    If i <= n Then
        JsonSB_Append sb, Mid$(s, i)
    End If

    Xml_DecodeEntities = JsonSB_Text(sb)
End Function
