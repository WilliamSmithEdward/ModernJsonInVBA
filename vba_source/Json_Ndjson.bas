Attribute VB_Name = "Json_Ndjson"
Option Explicit

' =============================================================================
' Module:      Json_Ndjson
' Project:     ModernJsonInVBA
' Version:     3.8.2
' Released:    2026-07-09
'
' NDJSON (newline-delimited JSON, also called JSON Lines) to JSON conversion.
'
' NDJSON holds one JSON value per line, with no enclosing array and no commas
' between records:
'
'     {"id":1,"name":"Alice"}
'     {"id":2,"name":"Bob"}
'
' Each line is already a valid JSON value, so conversion wraps the non-blank
' lines in a JSON array. The result feeds straight into Json_Parse or the
' table upsert (each line becomes one row), the same way CsvTextToJson and
' XmlTextToJson feed the pipeline.
'
' Line endings \r\n, \r, and \n are all accepted; blank lines are skipped;
' a leading byte-order mark is stripped. Lines are not re-validated here: a
' malformed line surfaces when the result is parsed, which matches the other
' converters.
'
' This module has no Excel references, so it is part of the all-O365 build.
' =============================================================================

' Read an NDJSON file and convert it via NdjsonToJson. The file is read
' through Json_ReadTextFile, so UTF-8 (with or without BOM), UTF-16, and
' legacy ANSI files all decode correctly.
Public Function NdjsonFileToJson(ByVal filePath As String) As String
    NdjsonFileToJson = NdjsonToJson(Json_ReadTextFile(filePath))
End Function

' Convert NDJSON text into a JSON array-of-values string.
Public Function NdjsonToJson(ByVal text As String) As String
    ' Strip a leading BOM if present.
    If Len(text) > 0 Then
        If AscW(Left$(text, 1)) = &HFEFF Then text = Mid$(text, 2)
    End If

    ' Normalize line endings so \r\n, \r, and \n all split uniformly.
    text = Replace$(text, vbCrLf, vbLf)
    text = Replace$(text, vbCr, vbLf)

    Dim lines() As String
    lines = Split(text, vbLf)

    ' Assemble "[line1,line2,...]" through the shared text builder, which
    ' grows in O(n) rather than the O(n^2) of repeated string concatenation.
    Dim sb As JsonTextBuilder
    JsonSB_Init sb, Len(text) + 16
    JsonSB_Append sb, "["

    Dim first As Boolean
    first = True

    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Dim ln As String
        ln = Trim$(lines(i))

        If Len(ln) > 0 Then
            If Not first Then JsonSB_Append sb, ","
            first = False
            JsonSB_Append sb, ln
        End If
    Next i

    JsonSB_Append sb, "]"
    NdjsonToJson = JsonSB_Text(sb)
End Function
