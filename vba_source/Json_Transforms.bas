Attribute VB_Name = "Json_Transforms"
Option Explicit

' =============================================================================
' Module:      Json_Transforms
' Project:     ModernJsonInVBA
'
' Structural transforms over the parsed model:
'
'   Json_Flatten              nested model -> tagged object of [path, value]
'   Json_FlatGet/FlatContains lookups against a flattened object
'   Json_Unflatten            [path, value] pairs -> nested tagged object
'   Json_FindArrayObjectRoots discover array-of-object table candidates
'
' Path format:
'   Root "$", object members "$.a.b", array elements "$.items[0].id".
'   Dots and backslashes inside keys are escaped ("\." / "\\") during
'   flattening so paths stay unambiguous.
'
' Determinism:
'   Emission order follows model insertion order; root discovery is
'   first-seen order. Both are stable across runs.
'
' Error numbers (all vbObjectError + n):
'   880/881/882 FlatGet (not tagged / path is object / path missing)
'   890  FlatContains expects tagged object
'   900  Unflatten expects tagged object
'   905  Unflatten does not support array index paths
'   907/908/909 Unflatten descend collisions
'   910  FindArrayObjectRoots expects tagged object
' =============================================================================

Private Const ERR_SRC As String = "ModernJsonInVBA"

' =============================================================================
' Public API: Flatten
' =============================================================================

' Flatten a parsed JSON value into a tagged object of [path, value] pairs:
'   (1) = JSON_TAG_OBJECT, (2..n) = Array(path As String, value As Variant)
'
' maxDepth:
'   Subtrees below maxDepth are stored as JSON text at their path.
'
' tableRootToExpand + arrayMode (table-aware flattening):
'   arrayMode = 0  expand every array (legacy behavior)
'   arrayMode = 1  expand ONLY the table root and its ancestors; other
'                  arrays are omitted entirely
'   arrayMode = 2  expand ONLY the table root and its ancestors; other
'                  arrays are stored as JSON text at their path
'   For modes 1/2 the root is normalized with Json_RemoveIndices so roots
'   like "$[0].items" still match their ancestors.
Public Function Json_Flatten( _
    ByVal parsedJson As Variant, _
    Optional ByVal maxDepth As Long = 12, _
    Optional ByVal tableRootToExpand As String = vbNullString, _
    Optional ByVal arrayMode As Long = 0 _
) As Collection

    Dim flat As New Collection
    flat.Add JSON_TAG_OBJECT

    Dim tableRootNorm As String
    tableRootNorm = Trim$(tableRootToExpand)

    If arrayMode <> 0 Then
        If Len(tableRootNorm) > 0 Then
            tableRootNorm = Json_RemoveIndices(tableRootNorm)
        End If
    End If

    If IsObject(parsedJson) Then
        If Json_IsObject(parsedJson) Or Json_IsArray(parsedJson) Then
            Json_FlattenInto flat, "$", parsedJson, 0, maxDepth, tableRootNorm, arrayMode
        Else
            Json_FlattenInto flat, vbNullString, parsedJson, 0, maxDepth, tableRootNorm, arrayMode
        End If
    Else
        Json_FlattenInto flat, vbNullString, parsedJson, 0, maxDepth, tableRootNorm, arrayMode
    End If

    Set Json_Flatten = flat
End Function

' Return the primitive value stored at an exact path (case-sensitive match).
' O(n) scan; use Json_FlatContains first if absence is expected.
Public Function Json_FlatGet(ByVal flatObj As Collection, ByVal path As String) As Variant
    If Not Json_IsObject(flatObj) Then
        Err.Raise vbObjectError + 880, ERR_SRC, "FlatGet expects tagged object"
    End If

    Dim isFirst As Boolean
    isFirst = True

    Dim pair As Variant
    For Each pair In flatObj
        If isFirst Then
            isFirst = False
        ElseIf CStr(pair(0)) = path Then
            If IsObject(pair(1)) Then
                Err.Raise vbObjectError + 881, ERR_SRC, "Path refers to object"
            End If
            Json_FlatGet = pair(1)
            Exit Function
        End If
    Next pair

    Err.Raise vbObjectError + 882, ERR_SRC, "Path not found: " & path
End Function

' True when the flattened object contains an exact path key. O(n) scan.
Public Function Json_FlatContains(ByVal flatObj As Collection, ByVal path As String) As Boolean
    If Not Json_IsObject(flatObj) Then
        Err.Raise vbObjectError + 890, ERR_SRC, "FlatContains expects tagged object"
    End If

    Dim isFirst As Boolean
    isFirst = True

    Dim pair As Variant
    For Each pair In flatObj
        If isFirst Then
            isFirst = False
        ElseIf CStr(pair(0)) = path Then
            Json_FlatContains = True
            Exit Function
        End If
    Next pair
End Function

' =============================================================================
' Public API: Unflatten
' =============================================================================

' Rebuild a nested tagged object from [path, value] pairs.
'
' A pair whose path is exactly "$" is stored under the key "$" in the result.
' Array index paths are NOT supported and raise vbObjectError + 905.
Public Function Json_Unflatten(ByVal flatObj As Collection) As Collection
    If Not Json_IsObject(flatObj) Then
        Err.Raise vbObjectError + 900, ERR_SRC, "Unflatten expects tagged object"
    End If

    Dim root As New Collection
    root.Add JSON_TAG_OBJECT

    Dim isFirst As Boolean
    isFirst = True

    Dim pair As Variant
    For Each pair In flatObj
        If isFirst Then
            isFirst = False
        Else
            Dim path As String
            path = CStr(pair(0))

            Dim value As Variant
            Json_VarAssign value, pair(1)

            If path = "$" Then
                Dim vv As Variant
                Json_VarAssign vv, value
                root.Add Array("$", vv)
            Else
                Json_UnflattenInsert root, path, value
            End If
        End If
    Next pair

    Set Json_Unflatten = root
End Function

' Insert value at a dotted path beneath root, creating intermediate tagged
' objects as needed. Internal: also used by Excel_ListObjectToJson to expand
' dotted column headers.
'
' Errors:
'   905 array index paths unsupported
'   907 existing value at a segment is a primitive
'   908 existing value is an object but not a Collection
'   909 existing Collection is not a tagged object
Public Sub Json_UnflattenInsert(ByVal root As Collection, ByVal path As String, ByVal value As Variant)
    If Left$(path, 2) = "$." Then
        path = Mid$(path, 3)
    End If

    If InStr(1, path, "[", vbBinaryCompare) > 0 Or InStr(1, path, "]", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 905, ERR_SRC, "Unflatten does not support array index paths: " & path
    End If

    Json_UnflattenInsertTokens root, Json_TokenizePath(path), value
End Sub

' Token-list variant of Json_UnflattenInsert. Internal: lets callers that
' insert repeatedly along the same path (per-row column writes) tokenize the
' path once instead of once per insert. Tokens are still escaped; each
' segment is unescaped here.
Public Sub Json_UnflattenInsertTokens(ByVal root As Collection, ByVal tokens As Collection, ByVal value As Variant)
    Dim current As Collection
    Set current = root

    Dim i As Long
    For i = 1 To tokens.count
        Dim key As String
        key = Json_UnescapePathSegment(CStr(tokens(i)))

        If i = tokens.count Then
            Json_ObjSet current, key, value
        Else
            Set current = Json_FindOrCreateChild(current, key)
        End If
    Next i
End Sub

' =============================================================================
' Public API: array-of-object root discovery
' =============================================================================

' Scan flattened paths and return candidate roots for array-of-object tables
' (paths shaped like root & "[n]." & column). Roots are unique and returned
' in first-seen order.
Public Function Json_FindArrayObjectRoots( _
    ByVal flatObj As Collection, _
    Optional ByVal stopAfterFirst As Boolean = False _
) As Collection

    If Not Json_IsObject(flatObj) Then
        Err.Raise vbObjectError + 910, ERR_SRC, "FindArrayObjectRoots expects tagged object"
    End If

    Dim roots As New Collection

    Dim seen As JsonStringIndex

    Dim isFirst As Boolean
    isFirst = True

    Dim pair As Variant
    For Each pair In flatObj
        If isFirst Then
            isFirst = False
        Else
            Json_CollectRootsFromPath roots, seen, CStr(pair(0)), stopAfterFirst

            If stopAfterFirst Then
                If roots.count > 0 Then Exit For
            End If
        End If
    Next pair

    Set Json_FindArrayObjectRoots = roots
End Function

' =============================================================================
' Flatten internals
' =============================================================================

Private Sub Json_FlattenInto( _
    ByVal flat As Collection, _
    ByVal prefix As String, _
    ByVal v As Variant, _
    ByVal depth As Long, _
    ByVal maxDepth As Long, _
    ByVal tableRootNorm As String, _
    ByVal arrayMode As Long _
)

    If depth > maxDepth Then
        AddFlat flat, IIf(Len(prefix) = 0, "$", prefix), Json_Stringify(v)
        Exit Sub
    End If

    If Not IsObject(v) Then
        AddFlat flat, IIf(Len(prefix) = 0, "$", prefix), v
        Exit Sub
    End If

    ' ---- Arrays ----
    If Json_IsArray(v) Then
        Dim arr As Collection
        Set arr = v

        Dim basePath As String
        basePath = IIf(Len(prefix) = 0, "$", prefix)

        Dim expandArray As Boolean

        If arrayMode = 0 Then
            expandArray = True
        Else
            ' Table-aware: expand only when this array IS the table root or
            ' an ancestor of it (compared with indices stripped).
            expandArray = False

            Dim baseNoIdx As String
            baseNoIdx = Json_RemoveIndices(basePath)

            If Len(tableRootNorm) > 0 Then
                If StrComp(baseNoIdx, tableRootNorm, vbBinaryCompare) = 0 Then
                    expandArray = True
                ElseIf Left$(tableRootNorm, Len(baseNoIdx) + 1) = (baseNoIdx & ".") Then
                    expandArray = True
                End If
            End If
        End If

        If expandArray Then
            ' For Each: indexed arr(i) access walks the Collection's linked
            ' list and is quadratic on large arrays.
            Dim i As Long
            i = 0

            Dim elem As Variant
            For Each elem In arr
                Dim idxPath As String
                idxPath = basePath & "[" & i & "]"
                i = i + 1

                If IsObject(elem) Then
                    If Json_IsObject(elem) Or Json_IsArray(elem) Then
                        Json_FlattenInto flat, idxPath, elem, depth + 1, maxDepth, tableRootNorm, arrayMode
                    Else
                        AddFlat flat, idxPath, Json_Stringify(elem)
                    End If
                Else
                    AddFlat flat, idxPath, elem
                End If
            Next elem
        Else
            ' Array outside the table root path: mode 2 stores it as JSON
            ' text; mode 1 omits it entirely.
            If arrayMode = 2 Then
                AddFlat flat, basePath, Json_Stringify(arr)
            End If
        End If

        Exit Sub
    End If

    ' ---- Objects ----
    If Json_IsObject(v) Then
        Dim obj As Collection
        Set obj = v

        Dim isFirst As Boolean
        isFirst = True

        Dim pair As Variant
        For Each pair In obj
            If isFirst Then
                isFirst = False
            Else
                Dim seg As String
                seg = Json_EscapePathSegment(CStr(pair(0)))

                Dim nextPrefix As String
                If Len(prefix) = 0 Then
                    nextPrefix = seg
                Else
                    nextPrefix = prefix & "." & seg
                End If

                Dim child As Variant
                Json_VarAssign child, pair(1)

                If IsObject(child) Then
                    If Json_IsObject(child) Or Json_IsArray(child) Then
                        Json_FlattenInto flat, nextPrefix, child, depth + 1, maxDepth, tableRootNorm, arrayMode
                    Else
                        AddFlat flat, nextPrefix, Json_Stringify(child)
                    End If
                Else
                    AddFlat flat, nextPrefix, child
                End If
            End If
        Next pair

        Exit Sub
    End If

    ' Non-Collection object: store its JSON representation.
    AddFlat flat, IIf(Len(prefix) = 0, "$", prefix), Json_Stringify(v)
End Sub

Private Sub AddFlat(ByVal flat As Collection, ByVal key As String, ByVal value As Variant)
    Dim vv As Variant
    Json_VarAssign vv, value

    ' Array(...) allocates a fresh pair per call; a reused local array would
    ' alias every pair to the same data.
    flat.Add Array(key, vv)
End Sub

' =============================================================================
' Unflatten internals
' =============================================================================

Private Function Json_FindOrCreateChild(ByVal parent As Collection, ByVal key As String) As Collection
    Dim i As Long
    For i = 2 To parent.count
        Dim pair As Variant
        pair = parent(i)

        If StrComp(CStr(pair(0)), key, vbBinaryCompare) = 0 Then
            If Not IsObject(pair(1)) Then
                Err.Raise vbObjectError + 907, ERR_SRC, _
                    "Unflatten collision at key '" & key & "': existing value is primitive, cannot descend."
            End If
            If TypeName(pair(1)) <> "Collection" Then
                Err.Raise vbObjectError + 908, ERR_SRC, _
                    "Unflatten collision at key '" & key & "': existing value is not a Collection."
            End If
            If Not Json_IsObject(pair(1)) Then
                Err.Raise vbObjectError + 909, ERR_SRC, _
                    "Unflatten collision at key '" & key & "': existing value is not a tagged object."
            End If

            Set Json_FindOrCreateChild = pair(1)
            Exit Function
        End If
    Next i

    Dim newObj As New Collection
    newObj.Add JSON_TAG_OBJECT

    parent.Add Array(key, newObj)

    Set Json_FindOrCreateChild = newObj
End Function

' =============================================================================
' Root discovery internals
' =============================================================================

' Extract array-of-object roots from one flattened path. A root is any
' prefix immediately followed by "[n]." in the path. The seen-index
' de-duplicates while roots preserves first-seen order.
Private Sub Json_CollectRootsFromPath( _
    ByVal roots As Collection, _
    ByRef seen As JsonStringIndex, _
    ByVal path As String, _
    ByVal stopAfterFirst As Boolean _
)
    ' Fast path for the most common shape: a root array of objects,
    ' producing paths like "$[0].id".
    If Len(path) >= 5 Then
        If Mid$(path, 1, 2) = "$[" Then
            If InStr(3, path, "].", vbBinaryCompare) > 0 Then
                Roots_AddIfMissing roots, seen, "$"
                Exit Sub
            End If
        End If
    End If

    Dim p As Long
    p = 1

    Do
        Dim openPos As Long
        openPos = InStr(p, path, "[")
        If openPos = 0 Then Exit Do

        Dim closePos As Long
        closePos = InStr(openPos + 1, path, "]")
        If closePos = 0 Then Exit Do

        If closePos < Len(path) Then
            If Mid$(path, closePos + 1, 1) = "." Then
                Dim rootPath As String
                rootPath = Left$(path, openPos - 1)

                If InStr(1, rootPath, "[", vbBinaryCompare) > 0 Then
                    rootPath = Json_RemoveIndices(rootPath)
                End If

                If Len(rootPath) > 0 Then
                    Roots_AddIfMissing roots, seen, rootPath
                    If stopAfterFirst Then Exit Sub
                End If
            End If
        End If

        p = closePos + 1
    Loop
End Sub

Private Sub Roots_AddIfMissing( _
    ByVal roots As Collection, _
    ByRef seen As JsonStringIndex, _
    ByVal s As String _
)
    Dim before As Long
    before = seen.count

    JsonIdx_Ensure seen, s

    If seen.count > before Then roots.Add s
End Sub
