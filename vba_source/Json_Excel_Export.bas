Attribute VB_Name = "Json_Excel_Export"
Option Explicit

' =============================================================================
' Module:      Json_Excel_Export
' Project:     ModernJsonInVBA
'
' Excel -> JSON export: the reverse direction of Json_Excel's ingestion.
'
'   Excel_ListObjectToJson       table -> JSON array-of-objects
'   Excel_RangeToJsonStrings     range -> Collection of JSON strings
'   Json_CoalesceArraysFromRange merge JSON arrays stored in a range
'
' Performance notes:
'   - Header paths are analyzed once per column (tokenize, bracket check,
'     unescape), not once per cell.
'   - Simple (non-dotted) columns append pairs directly; header validation
'     guarantees key uniqueness within a row, so the append is equivalent to
'     Json_ObjSet without its per-pair scan.
'
' Error numbers (all vbObjectError + n, source "Excel_ListObjectToJson"):
'   905  bracketed (array index) header path
'   1120 blank header            1121 duplicate header
'   1170 Excel error value       1171 body/header column mismatch
'   1172 TAG_OBJECT missing
' =============================================================================

' Convert a ListObject into a JSON array-of-objects. Each row becomes one
' object; headers define the property paths.
'
' Value precedence per cell:
'   1. JSON parsed from cell text  (parseJsonInCells=True and it parses)
'   2. Formula text                (preserveFormulas=True)
'   3. Raw Value2
'
' Dotted headers ("customer.name") nest; bracketed headers raise 905.
' Blank cells are omitted, or emitted as null when includeBlanksAsNull=True.
Public Function Excel_ListObjectToJson( _
    ByVal lo As ListObject, _
    Optional ByVal includeBlanksAsNull As Boolean = False, _
    Optional ByVal parseJsonInCells As Boolean = False, _
    Optional ByVal parseArraysOnly As Boolean = False, _
    Optional ByVal preserveFormulas As Boolean = False _
) As String

    Const SRC As String = "Excel_ListObjectToJson"

    If Len(JSON_TAG_OBJECT) = 0 Then
        Err.Raise vbObjectError + 1172, SRC, "TAG_OBJECT is blank or not initialized."
    End If

    ' ---- Headers ----
    Dim colCount As Long
    colCount = lo.ListColumns.count

    Dim headers() As String
    If colCount > 0 Then ReDim headers(1 To colCount)

    Dim c As Long
    For c = 1 To colCount
        headers(c) = Trim$(CStr(lo.ListColumns(c).name))
        If Len(headers(c)) = 0 Then
            Err.Raise vbObjectError + 1120, SRC, _
                "Header at index " & CStr(c) & " is blank."
        End If
    Next c

    Dim dupIdx As JsonStringIndex
    JsonIdx_Init dupIdx, colCount * 2, True

    For c = 1 To colCount
        Dim seenAt As Long
        seenAt = JsonIdx_Find(dupIdx, headers(c))
        If seenAt > 0 Then
            Err.Raise vbObjectError + 1121, SRC, _
                "Duplicate header: '" & headers(seenAt) & _
                "' at indices " & seenAt & " and " & c
        End If
        JsonIdx_Ensure dupIdx, headers(c)
    Next c

    ' ---- Per-column path analysis (once, not per cell) ----
    Dim colIsNested() As Boolean
    Dim colSimpleKey() As String
    Dim colTokens() As Collection

    If colCount > 0 Then
        ReDim colIsNested(1 To colCount)
        ReDim colSimpleKey(1 To colCount)
        ReDim colTokens(1 To colCount)
    End If

    For c = 1 To colCount
        Dim keyPath As String
        keyPath = headers(c)

        If InStr(keyPath, "[") > 0 Or InStr(keyPath, "]") > 0 Then
            Err.Raise vbObjectError + 905, SRC, _
                "Array index paths unsupported: " & keyPath
        End If

        Dim toks As Collection
        Set toks = Json_TokenizePath(keyPath)

        If toks.count > 1 Then
            colIsNested(c) = True
            ' Nested inserts strip a leading "$." exactly as
            ' Json_UnflattenInsert does, then reuse the token list per cell.
            Dim pathBody As String
            pathBody = keyPath
            If Left$(pathBody, 2) = "$." Then pathBody = Mid$(pathBody, 3)
            Set colTokens(c) = Json_TokenizePath(pathBody)
        Else
            colIsNested(c) = False
            colSimpleKey(c) = Json_UnescapePathSegment(CStr(toks(1)))
        End If
    Next c

    If lo.DataBodyRange Is Nothing Then
        Excel_ListObjectToJson = "[]"
        Exit Function
    End If

    ' ---- Bulk read ----
    Dim data As Variant
    data = lo.DataBodyRange.Value2

    Dim formulas As Variant
    If preserveFormulas Then
        formulas = lo.DataBodyRange.Formula
    End If

    ' A 1x1 body reads back as a scalar; normalize to a 2D array.
    If Not IsArray(data) Then
        Dim tmp(1 To 1, 1 To 1) As Variant
        tmp(1, 1) = data
        data = tmp
    End If

    If preserveFormulas Then
        If Not IsArray(formulas) Then
            Dim tmpf(1 To 1, 1 To 1) As Variant
            tmpf(1, 1) = formulas
            formulas = tmpf
        End If
    End If

    Dim rowCount As Long
    rowCount = UBound(data, 1) - LBound(data, 1) + 1

    Dim dataCols As Long
    dataCols = UBound(data, 2) - LBound(data, 2) + 1

    If dataCols <> colCount Then
        Err.Raise vbObjectError + 1171, SRC, _
            "DataBodyRange columns (" & dataCols & _
            ") do not match header count (" & colCount & ")."
    End If

    Dim anyNested As Boolean
    For c = 1 To colCount
        If colIsNested(c) Then anyNested = True
    Next c

    Dim r As Long
    Dim v As Variant
    Dim isBlank As Boolean

    ' ---- Streaming path (all headers simple, the common case) ----
    ' Serialize straight from the bulk-read array into one text builder: no
    ' per-row model Collections, and each key is quoted/escaped once per
    ' column instead of once per cell. Output is byte-identical to the model
    ' path because keys and values go through the same serializer writer.
    If Not anyNested Then
        Dim keyPrefix() As String
        ReDim keyPrefix(1 To colCount) As String

        Dim ksb As JsonTextBuilder
        Dim kv As Variant
        For c = 1 To colCount
            JsonSB_Init ksb, Len(colSimpleKey(c)) + 8
            kv = colSimpleKey(c)
            Json_StringifyInto ksb, kv
            keyPrefix(c) = JsonSB_Text(ksb) & ":"
        Next c

        Dim out As JsonTextBuilder
        JsonSB_Init out, rowCount * colCount * 12 + 16

        JsonSB_Append out, "["

        For r = 1 To rowCount
            If r > 1 Then JsonSB_Append out, ","
            JsonSB_Append out, "{"

            Dim firstMember As Boolean
            firstMember = True

            For c = 1 To colCount
                isBlank = Excel_ResolveCellValue(data, formulas, r, c, _
                    parseJsonInCells, parseArraysOnly, preserveFormulas, SRC, v)

                If isBlank Then
                    If Not includeBlanksAsNull Then GoTo NextStreamCell
                    v = Null
                End If

                If Not firstMember Then JsonSB_Append out, ","
                firstMember = False

                JsonSB_Append out, keyPrefix(c)
                Json_StringifyInto out, v

NextStreamCell:
            Next c

            JsonSB_Append out, "}"
        Next r

        JsonSB_Append out, "]"

        Excel_ListObjectToJson = JsonSB_Text(out)
        Exit Function
    End If

    ' ---- Model path (dotted headers present) ----
    ' Nested paths can interleave across columns, so rows build through the
    ' model where Json_UnflattenInsertTokens merges shared parents.
    Dim arr As New Collection

    For r = 1 To rowCount

        Dim rowObj As Collection
        Set rowObj = New Collection
        rowObj.Add JSON_TAG_OBJECT

        For c = 1 To colCount
            isBlank = Excel_ResolveCellValue(data, formulas, r, c, _
                parseJsonInCells, parseArraysOnly, preserveFormulas, SRC, v)

            If isBlank Then
                If Not includeBlanksAsNull Then GoTo NextCell
                v = Null
            End If

            If colIsNested(c) Then
                Json_UnflattenInsertTokens rowObj, colTokens(c), v
            Else
                ' Headers are validated unique, so within a row this key is
                ' always new; appending matches Json_ObjSet semantics.
                Dim vv As Variant
                Json_VarAssign vv, v
                rowObj.Add Array(colSimpleKey(c), vv)
            End If

NextCell:
        Next c

        arr.Add rowObj
    Next r

    Excel_ListObjectToJson = Json_Stringify(arr)
End Function

' Resolve one cell to its export value, applying the documented precedence:
' parsed JSON structure, then formula text, then the raw value. Returns True
' when the cell is blank. Raises 1170 on Excel error values.
Private Function Excel_ResolveCellValue( _
    ByRef data As Variant, _
    ByRef formulas As Variant, _
    ByVal r As Long, _
    ByVal c As Long, _
    ByVal parseJsonInCells As Boolean, _
    ByVal parseArraysOnly As Boolean, _
    ByVal preserveFormulas As Boolean, _
    ByVal errSource As String, _
    ByRef outV As Variant _
) As Boolean

    Dim v As Variant
    v = data(LBound(data, 1) + r - 1, LBound(data, 2) + c - 1)

    If IsError(v) Then
        Err.Raise vbObjectError + 1170, errSource, _
            "Excel error value at row " & r & ", col " & c
    End If

    ' 1) JSON structure parsed from cell text.
    Dim parsedJson As Boolean
    parsedJson = False

    If parseJsonInCells Then
        If VarType(v) = vbString Then
            Dim s As String
            s = Trim$(CStr(v))

            If Len(s) > 0 Then
                Dim firstCh As String
                firstCh = Left$(s, 1)

                Dim looksJson As Boolean
                If parseArraysOnly Then
                    looksJson = (firstCh = "[")
                Else
                    looksJson = (firstCh = "[" Or firstCh = "{")
                End If

                If looksJson Then
                    Dim parsedCell As Variant
                    If Excel_TryParseJsonCell(s, parsedCell) Then
                        If IsObject(parsedCell) Then
                            If TypeName(parsedCell) = "Collection" Then
                                If Json_IsObject(parsedCell) Or Json_IsArray(parsedCell) Then
                                    Json_VarAssign v, parsedCell
                                    parsedJson = True
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    ' 2) Formula text.
    If preserveFormulas And Not parsedJson Then
        Dim f As Variant
        f = formulas(LBound(formulas, 1) + r - 1, LBound(formulas, 2) + c - 1)

        If VarType(f) = vbString Then
            If Len(f) > 0 Then
                If Left$(f, 1) = "=" Then
                    v = f
                End If
            End If
        End If
    End If

    ' 3) Blank detection.
    If VarType(v) = vbString Then
        Excel_ResolveCellValue = (LenB(v) = 0)
    Else
        Excel_ResolveCellValue = IsEmpty(v)
    End If

    Json_VarAssign outV, v
End Function

' Parse cell text with the engine parser without ever raising; returns True
' only when the parse succeeded (failures mean "ordinary string cell").
Private Function Excel_TryParseJsonCell( _
    ByVal s As String, _
    ByRef outValue As Variant _
) As Boolean
    Excel_TryParseJsonCell = False
    Json_VarAssign outValue, Null

    On Error GoTo Fail

    Dim v As Variant
    Json_ParseInto s, v

    Json_VarAssign outValue, v
    Excel_TryParseJsonCell = True
    Exit Function

Fail:
    Err.Clear
End Function

' =============================================================================
' Ranges of JSON strings
' =============================================================================

' Merge JSON arrays stored in a range into a single array. Empty cells are
' ignored; strictMode validates object shape consistency.
Public Function Json_CoalesceArraysFromRange( _
    ByVal rng As Range, _
    Optional ByVal strictMode As Boolean = False _
) As String
    Json_CoalesceArraysFromRange = _
        Json_CoalesceArraysFromStrings(Excel_RangeToJsonStrings(rng), strictMode)
End Function

' Collect the non-empty, trimmed text values of a range (bulk Value2 read).
Public Function Excel_RangeToJsonStrings(ByVal rng As Range) As Collection
    Dim result As New Collection

    Dim data As Variant
    data = rng.Value2

    If IsArray(data) Then
        Dim r As Long
        Dim c As Long
        For r = LBound(data, 1) To UBound(data, 1)
            For c = LBound(data, 2) To UBound(data, 2)
                Dim txt As String
                txt = Trim$(CStr(data(r, c)))
                If Len(txt) > 0 Then result.Add txt
            Next c
        Next r
    Else
        Dim singleVal As String
        singleVal = Trim$(CStr(data))
        If Len(singleVal) > 0 Then result.Add singleVal
    End If

    Set Excel_RangeToJsonStrings = result
End Function
