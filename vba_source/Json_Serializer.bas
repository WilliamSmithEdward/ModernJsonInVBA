Attribute VB_Name = "Json_Serializer"
Option Explicit

' =============================================================================
' Module:      Json_Serializer
' Project:     ModernJsonInVBA
'
' Serializes the library's in-memory JSON model back to JSON text.
'
'   JSON Object => tagged Collection (JSON_TAG_OBJECT in slot 1) with pairs
'                  stored as Array(key, value), Collection(key, value), or
'                  alternating key/value entries
'   JSON Array  => untagged Collection
'   Primitives  => Variant (Null, Boolean, Number, String)
'
' Contracts:
'   - Untagged Collections are serialized as arrays UNLESS they contain
'     key/value-shaped entries, which indicates a missing tag and raises
'     vbObjectError + 1134. Tagging is the authoritative object signal.
'   - VBA arrays are rejected (vbObjectError + 1137): the model represents
'     JSON arrays as Collections; a Variant() at a value position is almost
'     always a caller mistake (e.g. Range.Value2 passed directly).
'
' Performance notes:
'   - The whole document is written into a single pre-allocated text builder;
'     no intermediate per-subtree strings are created.
'   - String escaping copies clean runs in chunks; strings that need no
'     escaping are appended with a single copy.
'
' Error numbers (all vbObjectError + n):
'   1134 untagged object-shaped Collection / invalid tagged object
'   1135 object entry has unrecognized shape
'   1136 object pair is malformed (missing key or value)
'   1137 VBA array encountered
' =============================================================================

Private Const ERR_SRC As String = "ModernJsonInVBA"

' =============================================================================
' Public API
' =============================================================================

' Serialize a model value to JSON text. See the module header for the model
' and error contract.
Public Function Json_Stringify(ByVal v As Variant) As String
    Dim sb As JsonTextBuilder
    JsonSB_Init sb, 512

    JsonW_WriteValue sb, v

    Json_Stringify = JsonSB_Text(sb)
End Function

' =============================================================================
' Writer internals
' =============================================================================

Private Sub JsonW_WriteValue(ByRef sb As JsonTextBuilder, ByRef v As Variant)

    ' Guard: this library's JSON arrays are Collections, not VBA arrays.
    If IsArray(v) Then
        Err.Raise vbObjectError + 1137, "Json_Stringify", _
            "VBA array encountered. This JSON engine represents arrays as Collection, not Variant(). " & _
            "You likely passed Range.Value2 or a key/value pair array as a top-level value."
    End If

    If IsObject(v) Then

        If Json_IsObject(v) Then
            JsonW_WriteObject sb, v
            Exit Sub
        End If

        If TypeName(v) = "Collection" Then
            Dim c As Collection
            Set c = v

            If Json_CollectionLooksLikeObject(c) Then
                Err.Raise vbObjectError + 1134, "Json_Stringify", _
                    "Collection appears to be an object but is not tagged with TAG_OBJECT."
            End If

            JsonW_WriteArray sb, c
            Exit Sub
        End If

        ' Non-Collection object: serialize its type name as a string.
        JsonSB_Append sb, """"
        JsonW_WriteEscaped sb, TypeName(v)
        JsonSB_Append sb, """"
        Exit Sub
    End If

    If IsNull(v) Then
        JsonSB_Append sb, "null"
    ElseIf VarType(v) = vbBoolean Then
        If v Then
            JsonSB_Append sb, "true"
        Else
            JsonSB_Append sb, "false"
        End If
    ElseIf VarType(v) = vbString Then
        JsonSB_Append sb, """"
        JsonW_WriteEscaped sb, CStr(v)
        JsonSB_Append sb, """"
    ElseIf IsNumeric(v) Then
        JsonSB_Append sb, Json_NumberToString(CDbl(v))
    Else
        JsonSB_Append sb, """"
        JsonW_WriteEscaped sb, CStr(v)
        JsonSB_Append sb, """"
    End If

End Sub

Private Sub JsonW_WriteArray(ByRef sb As JsonTextBuilder, ByVal c As Collection)
    JsonSB_Append sb, "["

    ' For Each: indexed c(i) access walks the Collection's linked list and
    ' is quadratic on large arrays.
    Dim first As Boolean
    first = True

    Dim item As Variant
    For Each item In c
        If first Then
            first = False
        Else
            JsonSB_Append sb, ","
        End If

        JsonW_WriteValue sb, item
    Next item

    JsonSB_Append sb, "]"
End Sub

' Writes a tagged object. Three pair shapes are accepted:
'   A) Array(key, value)            - the canonical shape the parser emits
'   B) Collection((1)=key,(2)=value)
'   C) alternating entries: key as String at i, value at i+1
Private Sub JsonW_WriteObject(ByRef sb As JsonTextBuilder, ByVal obj As Collection)

    If obj Is Nothing Then
        Err.Raise vbObjectError + 1134, ERR_SRC, _
            "Json_StringifyObject: object is Nothing."
    End If

    If obj.count < 1 Or CStr(obj(1)) <> JSON_TAG_OBJECT Then
        Err.Raise vbObjectError + 1134, ERR_SRC, _
            "Json_StringifyObject: collection is not a tagged object."
    End If

    JsonSB_Append sb, "{"

    ' For Each enumeration (indexed obj(i) is quadratic on wide objects).
    ' Shape C consumes two consecutive entries, so a string entry parks in
    ' pendingKey and pairs with the next enumerated entry. i tracks the
    ' 1-based collection position for error messages.
    Dim first As Boolean
    first = True

    Dim isTag As Boolean
    isTag = True

    Dim pendingKey As String
    Dim havePendingKey As Boolean
    havePendingKey = False

    Dim i As Long
    i = 1

    Dim entry As Variant
    For Each entry In obj
        If isTag Then
            isTag = False
        Else
            i = i + 1

            Dim keyStr As String
            Dim val As Variant
            Dim haveMember As Boolean
            haveMember = False

            If havePendingKey Then
                ' Second half of shape C: any entry is the value.
                keyStr = pendingKey
                Json_VarAssign val, entry
                havePendingKey = False
                haveMember = True

            ElseIf IsArray(entry) Then
                ' Shape A: Array(key, value)
                Dim lb As Long
                Dim ub As Long
                lb = LBound(entry)
                ub = UBound(entry)

                If (ub - lb + 1) < 2 Then
                    Err.Raise vbObjectError + 1136, ERR_SRC, _
                        "Json_StringifyObject: object pair at index " & CStr(i) & _
                        " must contain 2 elements (key,value)."
                End If

                keyStr = CStr(entry(lb))
                Json_VarAssign val, entry(lb + 1)
                haveMember = True

            ElseIf IsObject(entry) And TypeName(entry) = "Collection" Then
                ' Shape B: Collection((1)=key, (2)=value)
                If entry.count < 2 Then
                    Err.Raise vbObjectError + 1136, ERR_SRC, _
                        "Json_StringifyObject: object pair Collection at index " & CStr(i) & _
                        " must contain 2 elements (key,value)."
                End If

                keyStr = CStr(entry(1))
                Json_VarAssign val, entry(2)
                haveMember = True

            ElseIf VarType(entry) = vbString Then
                ' Shape C: alternating key/value entries; value follows.
                pendingKey = CStr(entry)
                havePendingKey = True

            Else
                Err.Raise vbObjectError + 1135, ERR_SRC, _
                    "Json_StringifyObject: object entry at index " & CStr(i) & _
                    " is not Array(key,value) or Collection(key,value) or String(key). Found type=" & TypeName(entry)
            End If

            If haveMember Then
                If Not first Then JsonSB_Append sb, ","
                first = False

                JsonSB_Append sb, """"
                JsonW_WriteEscaped sb, keyStr
                JsonSB_Append sb, """:"
                JsonW_WriteValue sb, val
            End If
        End If
    Next entry

    If havePendingKey Then
        Err.Raise vbObjectError + 1136, ERR_SRC, _
            "Json_StringifyObject: dangling key at final index " & CStr(i) & _
            " (missing value)."
    End If

    JsonSB_Append sb, "}"
End Sub

' True only if ANY element of the (untagged) Collection looks like a pair:
'   - an Array whose first element is a String key, or
'   - a 2-item Collection whose (1) is a String key.
' Tagged JSON objects nested inside are NOT pairs; skipping them prevents
' false positives on arrays-of-objects.
Private Function Json_CollectionLooksLikeObject(ByVal c As Collection) As Boolean
    Dim entry As Variant
    For Each entry In c

        If IsArray(entry) Then
            If (UBound(entry) - LBound(entry) + 1) >= 2 Then
                If VarType(entry(LBound(entry))) = vbString Then
                    Json_CollectionLooksLikeObject = True
                    Exit Function
                End If
            End If

        ElseIf IsObject(entry) Then
            If TypeName(entry) = "Collection" Then

                ' A tagged object is a value, not a pair.
                If entry.count >= 1 Then
                    If VarType(entry(1)) = vbString Then
                        If CStr(entry(1)) = JSON_TAG_OBJECT Then
                            GoTo NextItem
                        End If
                    End If
                End If

                If entry.count = 2 Then
                    If VarType(entry(1)) = vbString Then
                        Json_CollectionLooksLikeObject = True
                        Exit Function
                    End If
                End If
            End If
        End If

NextItem:
    Next entry

    Json_CollectionLooksLikeObject = False
End Function

' =============================================================================
' Escaping and numbers
' =============================================================================

' Appends s to the builder with JSON escaping applied. Clean runs between
' escapable characters are copied in one Mid$ each; a fully clean string is
' appended with a single copy.
'
' Escapes: quote, backslash, forward slash, \b \f \n \r \t, and \u00XX for
' remaining control characters. The scan reads UTF-16 byte pairs from a
' one-time snapshot (b = s) instead of allocating a one-character string
' per position; only characters with a zero high byte can need escaping.
Private Sub JsonW_WriteEscaped(ByRef sb As JsonTextBuilder, ByRef s As String)
    Dim L As Long
    L = Len(s)
    If L = 0 Then Exit Sub

    Dim b() As Byte
    b = s

    Dim runStart As Long
    runStart = 1

    Dim i As Long
    For i = 1 To L
        If b((i - 1) * 2 + 1) = 0 Then
            Dim escText As String
            escText = vbNullString

            Select Case b((i - 1) * 2)
                Case 34: escText = "\"""
                Case 92: escText = "\\"
                Case 47: escText = "\/"
                Case 8:  escText = "\b"
                Case 12: escText = "\f"
                Case 13: escText = "\r"
                Case 10: escText = "\n"
                Case 9:  escText = "\t"
                Case 0 To 31
                    escText = "\u" & Right$("0000" & Hex$(b((i - 1) * 2)), 4)
            End Select

            If Len(escText) > 0 Then
                If i > runStart Then
                    JsonSB_Append sb, Mid$(s, runStart, i - runStart)
                End If
                JsonSB_Append sb, escText
                runStart = i + 1
            End If
        End If
    Next i

    If runStart = 1 Then
        JsonSB_Append sb, s
    ElseIf runStart <= L Then
        JsonSB_Append sb, Mid$(s, runStart)
    End If
End Sub

' Culture-invariant number formatting: CStr uses the locale decimal
' separator, which JSON does not allow, so a non-dot separator is replaced.
Private Function Json_NumberToString(ByVal d As Double) As String
    Dim s As String
    s = CStr(d)

    Dim decSep As String
    decSep = Mid$(CStr(1.1), 2, 1)

    If decSep <> "." Then s = Replace$(s, decSep, ".")
    Json_NumberToString = s
End Function
