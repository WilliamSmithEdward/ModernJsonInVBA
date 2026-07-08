Attribute VB_Name = "Json_Model"
Option Explicit

' =============================================================================
' Module:      Json_Model
' Project:     ModernJsonInVBA
'
' The in-memory model contract and its accessors.
'
' Model:
'   JSON Object => VBA Collection tagged with JSON_TAG_OBJECT in slot (1),
'                  followed by pairs stored as Array(key, value)
'   JSON Array  => VBA Collection (untagged)
'   Primitives  => Variant (Null, Boolean, Long/Double, String)
'
' The tag is the authoritative signal distinguishing objects from arrays;
' shape is never inferred. Pair keys use binary (case-sensitive) comparison.
' Insertion order is preserved everywhere and is part of the library's
' determinism guarantees.
'
' This module also hosts the minimal JSONPath-style resolvers used by the
' table pipeline ("$", ".key", "[0]" - no wildcards, no filters).
' =============================================================================

Private Const ERR_SRC As String = "ModernJsonInVBA"

' =============================================================================
' Type checks
' =============================================================================

' True when v is a JSON object: a Collection whose slot (1) holds the tag.
' Pair shapes are not validated here; the tag alone decides.
Public Function Json_IsObject(ByVal v As Variant) As Boolean
    If Not IsObject(v) Then Exit Function
    If TypeName(v) <> "Collection" Then Exit Function

    Dim c As Collection
    Set c = v

    If c.count < 1 Then Exit Function
    If VarType(c(1)) = vbString Then
        Json_IsObject = (c(1) = JSON_TAG_OBJECT)
    End If
End Function

' True when v is a JSON array: any untagged Collection. Contents are not
' inspected; a Collection that "looks like" pairs but lacks the tag is still
' an array here (the serializer raises on that shape instead).
Public Function Json_IsArray(ByVal v As Variant) As Boolean
    If Not IsObject(v) Then Exit Function
    If TypeName(v) <> "Collection" Then Exit Function
    Json_IsArray = (Not Json_IsObject(v))
End Function

' =============================================================================
' Object member access
' =============================================================================

' Return the value stored under key, raising vbObjectError + 5320 when the
' key is absent. Keys compare case-sensitively.
Public Function Json_ObjGet(ByVal obj As Collection, ByVal key As String) As Variant
    Dim v As Variant

    If Json_TryObjGet(obj, key, v) Then
        Json_VarAssign Json_ObjGet, v
        Exit Function
    End If

    Err.Raise vbObjectError + 5320, "Json_ObjGet", _
        "JSON object key not found: '" & key & "'"
End Function

' Try-get variant of Json_ObjGet: returns False (and outValue = Null) when
' the key is absent. O(n) scan over the pairs.
Public Function Json_TryObjGet(ByVal obj As Collection, ByVal key As String, ByRef outValue As Variant) As Boolean
    Json_TryObjGet = False
    Json_VarAssign outValue, Null

    Dim isFirst As Boolean
    isFirst = True

    Dim pair As Variant
    For Each pair In obj
        If isFirst Then
            isFirst = False
        ElseIf StrComp(CStr(pair(0)), key, vbBinaryCompare) = 0 Then
            Json_VarAssign outValue, pair(1)
            Json_TryObjGet = True
            Exit Function
        End If
    Next pair
End Function

' Set key = value on a tagged object, overwriting in place when the key
' exists so the pair keeps its position (determinism contract). A new key
' is appended at the end.
Public Sub Json_ObjSet(ByVal obj As Collection, ByVal key As String, ByVal value As Variant)
    Dim vv As Variant
    Json_VarAssign vv, value

    ' Locate the key with the enumerator (indexed obj(i) is quadratic on
    ' wide objects), then mutate AFTER the enumeration has ended: removing
    ' items while a For Each is active is not safe on Collections.
    Dim foundAt As Long
    foundAt = 0

    Dim i As Long
    i = 0

    Dim entry As Variant
    For Each entry In obj
        i = i + 1
        If i >= 2 Then
            If IsArray(entry) Then
                If CStr(entry(LBound(entry))) = key Then
                    foundAt = i
                    Exit For
                End If
            End If
        End If
    Next entry

    If foundAt > 0 Then
        obj.Remove foundAt

        If foundAt - 1 >= 1 Then
            obj.Add Array(key, vv), , , foundAt - 1
        Else
            obj.Add Array(key, vv), , 2
        End If

        Exit Sub
    End If

    obj.Add Array(key, vv)
End Sub

' =============================================================================
' Path resolution
' =============================================================================

' Resolve a JSONPath-like path ("$", "$.a.b", "$.items[0].id") against a
' parsed model value. Returns False when any step cannot be resolved.
' No wildcards or filters; array indices are zero-based.
'
' Segments use the same escape convention as flattened paths and column
' headers: "\." is a literal dot inside a key and "\\" a literal backslash,
' so "$.a\.b.c" walks the key "a.b" then "c". Any other character after a
' backslash stays literal (Json_UnescapePathSegment's rules).
Public Function Json_TryResolvePath( _
    ByVal root As Variant, _
    ByVal path As String, _
    ByRef outValue As Variant _
) As Boolean

    Json_TryResolvePath = False
    Json_VarAssign outValue, Null

    path = Trim$(path)
    If Len(path) = 0 Then Exit Function

    If path = "$" Then
        Json_VarAssign outValue, root
        Json_TryResolvePath = True
        Exit Function
    End If

    If Left$(path, 2) <> "$." Then Exit Function
    If Not IsObject(root) Then Exit Function
    If TypeName(root) <> "Collection" Then Exit Function

    Dim cur As Variant
    Json_VarAssign cur, root

    Dim i As Long
    i = 3   ' after "$."

    Do While i <= Len(path)
        ' Read a member name up to the next unescaped "." or "[". A
        ' backslash keeps its following character inside the segment (raw;
        ' decoded just before the lookup), so "\." does not end the segment.
        Dim seg As String
        seg = vbNullString

        Do While i <= Len(path)
            Dim ch As String
            ch = Mid$(path, i, 1)
            If ch = "\" And i < Len(path) Then
                seg = seg & ch & Mid$(path, i + 1, 1)
                i = i + 2
            ElseIf ch = "." Or ch = "[" Then
                Exit Do
            Else
                seg = seg & ch
                i = i + 1
            End If
        Loop

        If Len(seg) > 0 Then
            If Not IsObject(cur) Then Exit Function
            If TypeName(cur) <> "Collection" Then Exit Function
            If Not Json_IsObject(cur) Then Exit Function

            Dim nextVal As Variant
            If Not Json_TryObjGet(cur, Json_UnescapePathSegment(seg), nextVal) Then Exit Function
            Json_VarAssign cur, nextVal
        End If

        ' Apply any number of "[n]" index steps.
        Do While i <= Len(path) And Mid$(path, i, 1) = "["
            Dim idx As Long
            If Not Json_TryReadBracketIndex(path, i, idx) Then Exit Function

            If Not IsObject(cur) Then Exit Function
            If TypeName(cur) <> "Collection" Then Exit Function
            If Json_IsObject(cur) Then Exit Function

            Dim arr As Collection
            Set arr = cur

            Dim oneBased As Long
            oneBased = idx + 1
            If oneBased < 1 Or oneBased > arr.count Then Exit Function

            Dim elem As Variant
            Json_VarAssign elem, arr(oneBased)
            Json_VarAssign cur, elem
        Loop

        If i <= Len(path) Then
            If Mid$(path, i, 1) = "." Then
                i = i + 1
            ElseIf Mid$(path, i, 1) <> "[" Then
                Exit Function
            End If
        End If
    Loop

    Json_VarAssign outValue, cur
    Json_TryResolvePath = True
End Function

' Parse "[n]" at position i of path. On success, advances i past the "]"
' and returns the zero-based index in outIndex.
Public Function Json_TryReadBracketIndex(ByVal path As String, ByRef i As Long, ByRef outIndex As Long) As Boolean
    Json_TryReadBracketIndex = False
    outIndex = 0

    If i > Len(path) Then Exit Function
    If Mid$(path, i, 1) <> "[" Then Exit Function

    Dim closePos As Long
    closePos = InStr(i + 1, path, "]")
    If closePos = 0 Then Exit Function

    Dim idxText As String
    idxText = Mid$(path, i + 1, closePos - i - 1)
    If Len(idxText) = 0 Or Not Json_IsAllDigits(idxText) Then Exit Function

    outIndex = CLng(idxText)
    i = closePos + 1
    Json_TryReadBracketIndex = True
End Function

' Resolve a dotted path ("$", "$.products") that must land on an array.
' Raising variant used by the coalesce pipeline; no index steps supported.
' Segments use the same "\." / "\\" escapes as Json_TryResolvePath.
'
' Errors (vbObjectError + n):
'   5310 empty path            5311 root is not an array
'   5312 path must start "$."  5313 traversal hit a non-object
'   5315 resolved value is not an array
Public Function Json_ResolveArrayPath( _
    ByVal root As Variant, _
    ByVal path As String _
) As Collection

    Const SRC As String = "Json_ResolveArrayPath"

    If Len(path) = 0 Then
        Err.Raise vbObjectError + 5310, SRC, "Path cannot be empty"
    End If

    If path = "$" Then
        If TypeOf root Is Collection Then
            Set Json_ResolveArrayPath = root
            Exit Function
        End If

        Err.Raise vbObjectError + 5311, SRC, _
            "Root is not an array"
    End If

    If Left$(path, 2) <> "$." Then
        Err.Raise vbObjectError + 5312, SRC, _
            "Path must begin with '$.'"
    End If

    Dim parts As Collection
    Set parts = Json_TokenizePath(Mid$(path, 3))

    Dim current As Object
    Set current = root

    Dim seg As Variant
    For Each seg In parts

        If Not TypeOf current Is Collection Then
            Err.Raise vbObjectError + 5313, SRC, _
                "Path traversal encountered non-object"
        End If

        Set current = Json_ObjGet(current, Json_UnescapePathSegment(CStr(seg)))
    Next seg

    If Not TypeOf current Is Collection Then
        Err.Raise vbObjectError + 5315, SRC, _
            "Resolved path is not an array"
    End If

    Set Json_ResolveArrayPath = current
End Function
