Attribute VB_Name = "Json_Coalesce"
Option Explicit

' =============================================================================
' Module:      Json_Coalesce
' Project:     ModernJsonInVBA
' Version:     3.8.2
' Released:    2026-07-09
'
' Merges JSON arrays into a single array-of-objects:
'
'   Json_CoalesceChildArrays        hoist nested child arrays out of a parent
'                                   array, optionally injecting parent fields
'   Json_CoalesceArraysFromStrings  concatenate JSON array strings
'
' Both are host-agnostic (no Excel references); the Range-based variant
' Json_CoalesceArraysFromRange lives in Json_Excel_Export.
'
' strictMode on either function validates that every object shares the shape
' (ordered key list) of the first object seen, failing fast on drift.
'
' Error numbers (all vbObjectError + n):
'   5201 value is not a JSON array         (FromStrings)
'   5202/5203 strict mode: non-object      (FromStrings)
'   5204 strict mode: shape mismatch       (FromStrings)
'   5301 parent key not found              (ChildArrays)
'   5302 child property is not an array    (ChildArrays)
'   5303 strict mode: non-object           (ChildArrays)
'   5304 strict mode: shape mismatch       (ChildArrays)
' =============================================================================

' Extract nested child arrays from a parent JSON array and merge them into a
' single array-of-objects, optionally injecting parent fields into each row.
'
' parentKeyMap entries are Array(source, destination) where source is:
'   "orderId"    copy the parent's field
'   "'literal"   inject a literal constant (leading apostrophe)
Public Function Json_CoalesceChildArrays( _
    ByVal parentJson As String, _
    ByVal parentRoot As String, _
    ByVal childProperty As String, _
    Optional ByVal strictMode As Boolean = False, _
    Optional ByVal parentKeyMap As Collection = Nothing _
) As String

    Const SRC As String = "Json_CoalesceChildArrays"

    Dim parsed As Variant
    Json_ParseInto parentJson, parsed

    Dim parents As Collection
    Set parents = Json_ResolveArrayPath(parsed, parentRoot)

    Dim result As New Collection

    Dim firstShape As Collection
    Dim shapeCaptured As Boolean

    Dim parentObj As Variant
    Dim childArr As Variant
    Dim childObj As Variant

    For Each parentObj In parents

        Dim found As Boolean
        found = Json_TryObjGet(parentObj, childProperty, childArr)

        If Not found Then GoTo NextParent
        If IsNull(childArr) Then GoTo NextParent

        If (Not IsObject(childArr)) Or (TypeName(childArr) <> "Collection") Then
            Err.Raise vbObjectError + 5302, SRC, _
                "Child property is not an array: '" & childProperty & "'"
        End If

        For Each childObj In childArr

            If strictMode Then
                If childObj.count = 0 Or childObj(1) <> JSON_TAG_OBJECT Then
                    Err.Raise vbObjectError + 5303, SRC, _
                        "Strict mode requires arrays of objects"
                End If

                If Not shapeCaptured Then
                    Set firstShape = Json_ObjectShape(childObj)
                    shapeCaptured = True
                ElseIf Not Json_ObjectShapeMatches(childObj, firstShape) Then
                    Err.Raise vbObjectError + 5304, SRC, _
                        "Child array shapes are inconsistent"
                End If
            End If

            ' Shallow-clone the child so injected fields never mutate the
            ' parsed source document. For Each: indexed access into wide
            ' objects walks the Collection's linked list per hit.
            Dim row As Collection
            Set row = New Collection
            row.Add JSON_TAG_OBJECT

            Dim cloneFirst As Boolean
            cloneFirst = True

            Dim pair As Variant
            For Each pair In childObj
                If cloneFirst Then
                    cloneFirst = False
                Else
                    row.Add pair
                End If
            Next pair

            If Not parentKeyMap Is Nothing Then
                Dim keyPair As Variant
                For Each keyPair In parentKeyMap

                    Dim srcKey As String
                    srcKey = CStr(keyPair(0))

                    If Left$(srcKey, 1) = "'" Then
                        ' Literal injection: strip the marker apostrophe.
                        Json_ObjSet row, keyPair(1), Mid$(srcKey, 2)
                    Else
                        Dim parentVal As Variant
                        If Not Json_TryObjGet(parentObj, srcKey, parentVal) Then
                            Err.Raise vbObjectError + 5301, SRC, _
                                "Parent key not found: '" & srcKey & "'"
                        End If

                        Json_ObjSet row, keyPair(1), parentVal
                    End If
                Next keyPair
            End If

            result.Add row
        Next childObj

NextParent:
    Next parentObj

    Json_CoalesceChildArrays = Json_Stringify(result)
End Function

' Merge multiple JSON array strings into one array, preserving order.
Public Function Json_CoalesceArraysFromStrings( _
    ByVal jsonStrings As Collection, _
    Optional ByVal strictMode As Boolean = False _
) As String

    Const SRC As String = "Json_CoalesceArraysFromStrings"

    Dim result As New Collection

    Dim firstShape As Collection
    Dim shapeCaptured As Boolean

    Dim i As Long
    For i = 1 To jsonStrings.count

        Dim parsed As Variant
        Json_ParseInto CStr(jsonStrings(i)), parsed

        If Not TypeOf parsed Is Collection Then
            Err.Raise vbObjectError + 5201, SRC, _
                "Value is not a JSON array"
        End If

        Dim arr As Collection
        Set arr = parsed

        Dim obj As Variant
        For Each obj In arr

            If strictMode Then
                If Not TypeOf obj Is Collection Then
                    Err.Raise vbObjectError + 5202, SRC, _
                        "Strict mode requires arrays of objects"
                End If

                If obj.count = 0 Or obj(1) <> JSON_TAG_OBJECT Then
                    Err.Raise vbObjectError + 5203, SRC, _
                        "Strict mode requires arrays of objects"
                End If

                If Not shapeCaptured Then
                    Set firstShape = Json_ObjectShape(obj)
                    shapeCaptured = True
                ElseIf Not Json_ObjectShapeMatches(obj, firstShape) Then
                    Err.Raise vbObjectError + 5204, SRC, _
                        "Array object shapes are inconsistent"
                End If
            End If

            result.Add obj
        Next obj
    Next i

    Json_CoalesceArraysFromStrings = Json_Stringify(result)
End Function

' =============================================================================
' Shape validation internals
' =============================================================================

' Ordered key list of a tagged object (assumes Array(key, value) pairs).
Private Function Json_ObjectShape(ByVal obj As Collection) As Collection
    Dim shape As New Collection

    Dim isFirst As Boolean
    isFirst = True

    Dim pair As Variant
    For Each pair In obj
        If isFirst Then
            isFirst = False
        Else
            shape.Add pair(0)
        End If
    Next pair

    Set Json_ObjectShape = shape
End Function

' True when obj has exactly the keys of shape, in the same order.
Private Function Json_ObjectShapeMatches( _
    ByVal obj As Collection, _
    ByVal shape As Collection _
) As Boolean

    If obj.count - 1 <> shape.count Then Exit Function

    ' Snapshot obj's keys once, then compare while enumerating shape:
    ' element-by-element indexed access would walk both linked lists.
    Dim keyCount As Long
    keyCount = shape.count

    If keyCount = 0 Then
        Json_ObjectShapeMatches = True
        Exit Function
    End If

    Dim keys() As String
    ReDim keys(1 To keyCount) As String

    Dim n As Long
    n = 0

    Dim isFirst As Boolean
    isFirst = True

    Dim pair As Variant
    For Each pair In obj
        If isFirst Then
            isFirst = False
        Else
            n = n + 1
            keys(n) = CStr(pair(0))
        End If
    Next pair

    Dim i As Long
    i = 0

    Dim s As Variant
    For Each s In shape
        i = i + 1
        If keys(i) <> CStr(s) Then Exit Function
    Next s

    Json_ObjectShapeMatches = True
End Function
