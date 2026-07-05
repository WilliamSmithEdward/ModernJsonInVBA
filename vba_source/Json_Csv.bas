Attribute VB_Name = "Json_Csv"
Option Explicit

' =============================================================================
' Module:      Json_Csv
' Project:     ModernJsonInVBA
'
' CSV -> JSON conversion.
'
' CsvTextToJson implements an RFC-4180 style parser:
'   - fields separated by commas, records by LF or CRLF
'   - quoted fields may contain commas, quotes ("" escape), and newlines
'   - whitespace outside quotes is preserved verbatim
'   - a final record without a trailing newline is kept
'
' The first record supplies the keys; every following record becomes one
' JSON object with all values as strings. Short records are padded with
' empty strings. Output is a JSON array-of-objects (tableRoot "$").
'
' Performance notes:
'   - Output is written into a single pre-allocated text builder.
'   - Field content is copied in runs (InStr / code scanning) instead of
'     one character at a time.
' =============================================================================

' Read a CSV file and convert it via CsvTextToJson.
Public Function CsvFileToJson(ByVal filePath As String) As String
    Dim f As Integer
    f = FreeFile

    Dim txt As String
    Open filePath For Input As #f
    txt = Input$(LOF(f), f)
    Close #f

    CsvFileToJson = CsvTextToJson(txt)
End Function

' Convert raw CSV text into a JSON array-of-objects.
Public Function CsvTextToJson(ByVal txt As String) As String

    ' ---- Parse into rows of fields ----
    Dim rows As New Collection
    Dim fields As New Collection

    Dim fieldSB As JsonTextBuilder
    JsonSB_Init fieldSB, 256

    Dim inQuotes As Boolean

    Dim L As Long
    L = Len(txt)

    ' UTF-16 snapshot: structural scanning reads byte pairs instead of
    ' allocating a one-character string per position.
    Dim tb() As Byte
    If L > 0 Then tb = txt

    Dim i As Long
    i = 1

    Do While i <= L

        If inQuotes Then
            ' Jump straight to the next quote; everything before it is data.
            Dim q As Long
            q = InStr(i, txt, """", vbBinaryCompare)

            If q = 0 Then
                ' Unterminated quoted field: keep the remainder as data.
                JsonSB_Append fieldSB, Mid$(txt, i)
                i = L + 1
            Else
                If q > i Then JsonSB_Append fieldSB, Mid$(txt, i, q - i)

                If q < L And Mid$(txt, q + 1, 1) = """" Then
                    ' Doubled quote = literal quote inside the field.
                    JsonSB_Append fieldSB, """"
                    i = q + 2
                Else
                    inQuotes = False
                    i = q + 1
                End If
            End If

        Else
            Select Case tb((i - 1) * 2) + tb((i - 1) * 2 + 1) * 256&

                Case 34         ' quote: enter quoted mode
                    inQuotes = True
                    i = i + 1

                Case 44         ' comma: field boundary
                    fields.Add JsonSB_Text(fieldSB)
                    fieldSB.used = 0
                    i = i + 1

                Case 13         ' CR: ignored (CRLF handled by the LF)
                    i = i + 1

                Case 10         ' LF: record boundary
                    fields.Add JsonSB_Text(fieldSB)
                    fieldSB.used = 0
                    Csv_FlushRow rows, fields
                    i = i + 1

                Case Else
                    ' Copy the run up to the next structural character.
                    Dim runStart As Long
                    runStart = i

                    Do While i <= L
                        Select Case tb((i - 1) * 2) + tb((i - 1) * 2 + 1) * 256&
                            Case 34, 44, 13, 10
                                Exit Do
                            Case Else
                                i = i + 1
                        End Select
                    Loop

                    JsonSB_Append fieldSB, Mid$(txt, runStart, i - runStart)
            End Select
        End If
    Loop

    ' Final record when the text does not end with a newline.
    If fieldSB.used > 0 Or fields.count > 0 Then
        fields.Add JsonSB_Text(fieldSB)
        fieldSB.used = 0
        Csv_FlushRow rows, fields
    End If

    If rows.count = 0 Then
        CsvTextToJson = "[]"
        Exit Function
    End If

    ' ---- Build JSON ----
    ' For Each over rows: indexed rows(r) access walks the Collection's
    ' linked list and is quadratic on files with many records.
    Dim headers() As String
    headers = rows(1)

    Dim out As JsonTextBuilder
    JsonSB_Init out, L + 64

    JsonSB_Append out, "["

    Dim isHeaderRow As Boolean
    isHeaderRow = True

    Dim firstRecord As Boolean
    firstRecord = True

    Dim rowVar As Variant
    For Each rowVar In rows
        If isHeaderRow Then
            isHeaderRow = False
        Else
            If Not firstRecord Then JsonSB_Append out, ","
            firstRecord = False

            JsonSB_Append out, "{"

            Dim cols() As String
            cols = rowVar

            Dim c As Long
            For c = LBound(headers) To UBound(headers)

                If c > LBound(headers) Then JsonSB_Append out, ","

                JsonSB_Append out, """"
                Csv_AppendEscaped out, headers(c)
                JsonSB_Append out, """:"""

                ' Short records read as empty strings for missing columns.
                If c <= UBound(cols) Then
                    Csv_AppendEscaped out, cols(c)
                End If

                JsonSB_Append out, """"
            Next c

            JsonSB_Append out, "}"
        End If
    Next rowVar

    JsonSB_Append out, "]"

    CsvTextToJson = JsonSB_Text(out)
End Function

' Move the accumulated fields into rows as a 0-based String array and start
' a fresh field list.
Private Sub Csv_FlushRow(ByVal rows As Collection, ByRef fields As Collection)
    Dim row() As String
    ReDim row(0 To fields.count - 1)

    Dim j As Long
    For j = 1 To fields.count
        row(j - 1) = fields(j)
    Next j

    rows.Add row
    Set fields = New Collection
End Sub

' JSON string escaping for CSV-sourced values: quote, backslash, \t, and
' newlines normalized to \n (CRLF collapses to a single \n, matching the
' historical behavior). Remaining control characters are emitted as \uXXXX
' so the output always survives a round trip through Json_Parse. The scan
' reads UTF-16 byte pairs from a one-time snapshot; only characters with a
' zero high byte can need escaping.
Private Sub Csv_AppendEscaped(ByRef sb As JsonTextBuilder, ByRef s As String)
    Dim L As Long
    L = Len(s)
    If L = 0 Then Exit Sub

    Dim b() As Byte
    b = s

    Dim runStart As Long
    runStart = 1

    Dim i As Long
    i = 1

    Do While i <= L
        If b((i - 1) * 2 + 1) = 0 Then
            Dim c As Long
            c = b((i - 1) * 2)

            Dim escText As String
            escText = vbNullString

            Select Case c
                Case 34: escText = "\"""
                Case 92: escText = "\\"
                Case 9:  escText = "\t"
                Case 10: escText = "\n"
                Case 13
                    escText = "\n"
                Case 0 To 8, 11, 12, 14 To 31
                    escText = "\u" & Right$("0000" & Hex$(c), 4)
            End Select

            If Len(escText) > 0 Then
                If i > runStart Then
                    JsonSB_Append sb, Mid$(s, runStart, i - runStart)
                End If
                JsonSB_Append sb, escText

                ' CRLF collapses to one \n.
                If c = 13 And i < L Then
                    If b(i * 2) = 10 And b(i * 2 + 1) = 0 Then i = i + 1
                End If

                runStart = i + 1
            End If
        End If

        i = i + 1
    Loop

    If runStart = 1 Then
        JsonSB_Append sb, s
    ElseIf runStart <= L Then
        JsonSB_Append sb, Mid$(s, runStart)
    End If
End Sub
