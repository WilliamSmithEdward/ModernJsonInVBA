Option Explicit

' =============================================================================
' Module:      zz_ModernJsonInVba
' Project:     ModernJsonInVBA
'
' Author:      William Smith
' Created:     2026-02-28
'
' Summary
'   Deterministic JSON parsing, flattening, table extraction, and Excel ListObject
'   upsert utilities. No external dependencies. No Scripting.Dictionary.
'
' Model
'   - JSON Object  => VBA Collection tagged with TAG_OBJECT in slot(1),
'                    then pairs as Variant(0 To 1): [key, value]
'   - JSON Array   => VBA Collection (NOT tagged)
'   - Primitives   => Variant (Null, Boolean, Double/Long, String)
'
' Determinism Contracts
'   - Collection insertion order is preserved and used for stable results.
'   - Header discovery is first-seen order.
'   - Errors are raised with stable numbers and clear sources.
' =============================================================================

' =============================================================================
' Constants
' =============================================================================

Private Const TAG_OBJECT As String = "__OBJ__"
Private Const ERR_SRC As String = "zz_ModernJsonInVBA"
Private Const EXCEL_CHUNK_MAX_CELLS As Long = 50000

' =============================================================================
' Types
' =============================================================================

Private Type JsonReader
    text As String
    pos As Long ' 1-based, next char to read
End Type

Private Type RowKeyMap
    cap As Long
    slotHash() As Long
    slotIdx() As Long          ' stores 1-based index into rowKeys/rowObjs
    rowKeys() As String        ' 1-based
    rowObjs() As Collection    ' 1-based
    count As Long
End Type

' =============================================================================
' Enums
' =============================================================================

Public Enum ExcelSourceFormat
    ExcelSourceFormat_JSON = 1
    ExcelSourceFormat_CSV = 2
    ExcelSourceFormat_XML = 3
End Enum

' =============================================================================
' Public API: JSON Parse
' =============================================================================

' Parse jsonText and return:
'   - Object: as Object (Collection tagged with TAG_OBJECT)
'   - Array:  as Object (Collection)
'   - Primitive: as Variant
'
' Errors:
'   vbObjectError + 700  Trailing characters
'   plus parse-specific errors from lower layers
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

' Parse jsonText into outValue.
'
' Errors:
'   vbObjectError + 700  Trailing characters
'   plus parse-specific errors from lower layers
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
' Public API: JSON Type Helpers
' =============================================================================

' =============================================================================
' Json_IsObject
'
' Purpose:
'   Determine whether v is a JSON Object in this library's in-memory model.
'
' Model Contract:
'   - JSON Object is represented as a VBA Collection tagged with TAG_OBJECT at (1).
'   - Pairs follow at (2..n) as Array(key,value) (preferred) or other accepted shapes.
'
' Returns:
'   True  => v is a Collection and v(1)=TAG_OBJECT
'   False => otherwise
'
' Notes:
'   - This does NOT validate the object’s internal pair shapes; it only checks the tag.
' =============================================================================
Public Function Json_IsObject(ByVal v As Variant) As Boolean
    If Not IsObject(v) Then Exit Function
    If TypeName(v) <> "Collection" Then Exit Function

    Dim c As Collection
    Set c = v

    If c.count < 1 Then Exit Function
    If VarType(c(1)) = vbString Then
        Json_IsObject = (c(1) = TAG_OBJECT)
    End If
End Function

' =============================================================================
' Json_IsArray
'
' Purpose:
'   Determine whether v is a JSON Array in this library's in-memory model.
'
' Model Contract:
'   - JSON Array is represented as an UNTAGGED VBA Collection.
'   - A tagged Collection (TAG_OBJECT at (1)) is treated as an object, not an array.
'
' Returns:
'   True  => v is a Collection AND NOT a tagged object
'   False => otherwise
'
' Notes:
'   - This function treats any untagged Collection as an array, even if its contents
'     “look like” key/value pairs. Tagging is the authoritative object signal.
' =============================================================================
Public Function Json_IsArray(ByVal v As Variant) As Boolean
    If Not IsObject(v) Then Exit Function
    If TypeName(v) <> "Collection" Then Exit Function
    Json_IsArray = (Not Json_IsObject(v))
End Function

' =============================================================================
' Public API: JSON Stringify
' =============================================================================

' =============================================================================
' Json_Stringify
'
' Purpose:
'   Serialize the library's in-memory JSON model into JSON text.
'
' Model:
'   - JSON Object  => VBA Collection tagged with TAG_OBJECT in slot(1),
'                    followed by key/value entries (2-tuple arrays or 2-item Collections).
'   - JSON Array   => VBA Collection (NOT tagged)
'   - Primitives   => Variant (Null, Boolean, Number, String)
'
' Determinism / Contracts:
'   - Tagged objects are required for object serialization.
'   - Untagged Collections are treated as arrays UNLESS they are "object-shaped"
'     (i.e., contain key/value pair entries). Object-shaped but untagged is a
'     contract violation and MUST raise vbObjectError + 1134.
'
' Errors:
'   - vbObjectError + 1134 : Collection appears to be an object but is not tagged.
' =============================================================================
Public Function Json_Stringify(ByVal v As Variant) As String

    ' Guard: this library's JSON arrays are Collections, not VBA arrays.
    If IsArray(v) Then
        Err.Raise vbObjectError + 1137, "Json_Stringify", _
            "VBA array encountered. This JSON engine represents arrays as Collection, not Variant(). " & _
            "You likely passed Range.Value2 or a key/value pair array as a top-level value."
    End If

    If IsObject(v) Then

        If Json_IsObject(v) Then
            Json_Stringify = Json_StringifyObject(v)
            Exit Function
        End If

        If TypeName(v) = "Collection" Then
            Dim c As Collection
            Set c = v

            If Json_CollectionLooksLikeObject(c) Then
                Err.Raise vbObjectError + 1134, "Json_Stringify", _
                    "Collection appears to be an object but is not tagged with TAG_OBJECT."
            End If

            Json_Stringify = Json_StringifyArray(c)
            Exit Function
        End If

        Json_Stringify = """" & Json_EscapeString(TypeName(v)) & """"
        Exit Function
    End If

    If IsNull(v) Then
        Json_Stringify = "null"
    ElseIf VarType(v) = vbBoolean Then
        If v Then Json_Stringify = "true" Else Json_Stringify = "false"
    ElseIf VarType(v) = vbString Then
        Json_Stringify = """" & Json_EscapeString(CStr(v)) & """"
    ElseIf IsNumeric(v) Then
        Json_Stringify = Json_NumberToString(CDbl(v))
    Else
        Json_Stringify = """" & Json_EscapeString(CStr(v)) & """"
    End If

End Function

Private Function Json_CollectionLooksLikeObject(ByVal c As Collection) As Boolean
    ' Returns True only if ANY element looks like a *pair*:
    '   - 2-element Array(key, value) where key is String
    '   - 2-item Collection where (1)=key (String) and (2)=value
    '
    ' IMPORTANT:
    '   - Do NOT treat tagged JSON objects (Collection with TAG_OBJECT at (1))
    '     as "pair collections". This prevents false-positive on array-of-objects.
    '   - Must use Set when pulling object items out of a Collection.

    Dim i As Long
    For i = 1 To c.count

        Dim entry As Variant
        If IsObject(c(i)) Then
            Set entry = c(i)
        Else
            entry = c(i)
        End If

        ' Pair as 2-element array: Array(key, value)
        If IsArray(entry) Then
            If (UBound(entry) - LBound(entry) + 1) >= 2 Then
                If VarType(entry(LBound(entry))) = vbString Then
                    Json_CollectionLooksLikeObject = True
                    Exit Function
                End If
            End If

        ' Pair as 2-item Collection: (1)=key, (2)=value
        ElseIf IsObject(entry) Then
            If TypeName(entry) = "Collection" Then

                ' If it's a tagged object, it is NOT a pair.
                If entry.count >= 1 Then
                    If VarType(entry(1)) = vbString Then
                        If CStr(entry(1)) = TAG_OBJECT Then
                            GoTo NextItem
                        End If
                    End If
                End If

                ' Treat as pair only if it looks exactly like (key,value)
                If entry.count = 2 Then
                    If VarType(entry(1)) = vbString Then
                        Json_CollectionLooksLikeObject = True
                        Exit Function
                    End If
                End If
            End If
        End If

NextItem:
    Next i

    Json_CollectionLooksLikeObject = False
End Function

' =============================================================================
' Json_Flatten
'
' Purpose:
'   Flatten a parsed JSON value into a tagged object of [path,value] pairs.
'
' Output Shape:
'   Returns a tagged object Collection where:
'     (1) = TAG_OBJECT
'     (2..n) = Array(path As String, value As Variant)
'
' Path Format:
'   - Root: "$"
'   - Object: "$.a.b"
'   - Array index: "$.items[0].id"
'   - Keys with dots are escaped during flatten via Json_EscapePathSegment.
'
' Determinism:
'   - Pair emission order follows deterministic traversal of tagged objects and arrays.
'
' Parameters:
'   maxDepth:
'     When exceeded, the remaining subtree is stored as JSON text at that path.
'
'   tableRootToExpand + arrayMode:
'     arrayMode=0 (legacy): expand all arrays
'     arrayMode=1: expand ONLY arrays that are the tableRoot or ancestors of it; exclude all others
'     arrayMode=2: expand ONLY tableRoot/ancestors; stringify all other arrays into the cell
'
' Notes:
'   - For table-aware modes (1/2), we normalize tableRootToExpand by removing indices
'     so roots like "$[0].items" or "$.orders[0].items" still correctly expand ancestors.
' =============================================================================
Public Function Json_Flatten( _
    ByVal parsedJson As Variant, _
    Optional ByVal maxDepth As Long = 12, _
    Optional ByVal tableRootToExpand As String = vbNullString, _
    Optional ByVal arrayMode As Long = 0 _
) As Collection

    Dim flat As New Collection
    flat.Add TAG_OBJECT

    Dim tableRootNorm As String
    tableRootNorm = Trim$(tableRootToExpand)

    ' Normalize only for table-aware flattening (modes 1/2).
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

' =============================================================================
' Json_FlatGet
'
' Purpose:
'   Retrieve the primitive value at an exact path from a flattened tagged object.
'
' Parameters:
'   flatObj:
'     Tagged object from Json_Flatten (TAG_OBJECT at slot 1).
'   path:
'     Exact path key to find (case-sensitive, binary compare).
'
' Returns:
'   The stored value (must be non-object).
'
' Errors:
'   vbObjectError + 880  flatObj is not a tagged object
'   vbObjectError + 881  path exists but refers to an object/array (IsObject=True)
'   vbObjectError + 882  path not found
'
' Notes:
'   - This is an O(n) scan over flatObj pairs (2..count).
' =============================================================================
Public Function Json_FlatGet(ByVal flatObj As Collection, ByVal path As String) As Variant
    If Not Json_IsObject(flatObj) Then
        Err.Raise vbObjectError + 880, ERR_SRC, "FlatGet expects tagged object"
    End If

    Dim i As Long
    For i = 2 To flatObj.count
        Dim pair As Variant
        pair = flatObj(i)

        If CStr(pair(0)) = path Then
            If IsObject(pair(1)) Then
                Err.Raise vbObjectError + 881, ERR_SRC, "Path refers to object"
            End If
            Json_FlatGet = pair(1)
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 882, ERR_SRC, "Path not found: " & path
End Function

' =============================================================================
' Json_FlatContains
'
' Purpose:
'   Check whether a flattened tagged object contains an exact path key.
'
' Parameters:
'   flatObj:
'     Tagged object from Json_Flatten (TAG_OBJECT at slot 1).
'   path:
'     Exact path key to find (case-sensitive, binary compare).
'
' Returns:
'   True if a pair exists with pair(0)=path; False otherwise.
'
' Errors:
'   vbObjectError + 890  flatObj is not a tagged object
'
' Notes:
'   - This is an O(n) scan over flatObj pairs (2..count).
' =============================================================================
Public Function Json_FlatContains(ByVal flatObj As Collection, ByVal path As String) As Boolean
    If Not Json_IsObject(flatObj) Then
        Err.Raise vbObjectError + 890, ERR_SRC, "FlatContains expects tagged object"
    End If

    Dim i As Long
    For i = 2 To flatObj.count
        Dim pair As Variant
        pair = flatObj(i)
        If CStr(pair(0)) = path Then
            Json_FlatContains = True
            Exit Function
        End If
    Next i
End Function

' =============================================================================
' Json_Unflatten
'
' Purpose:
'   Reconstruct a nested tagged object from a flattened tagged object of [path,value] pairs.
'
' Input Shape:
'   flatObj must be a tagged object:
'     (1)=TAG_OBJECT
'     (2..n)=Array(path,value)
'
' Output Shape:
'   Returns a tagged object root. Special case:
'     - A flat pair with path="$" is stored under key "$" in the returned object.
'
' Limitations:
'   - Array indices in paths are NOT supported and will raise.
'
' Errors:
'   vbObjectError + 900  flatObj not tagged object
'   vbObjectError + 905  array index paths encountered (raised by internals)
'   vbObjectError + 907+ unflatten collision / invalid existing type while descending
'
' Notes:
'   - Keys with escaped dots and backslashes are handled via Json_(Un)escapePathSegment.
' =============================================================================
Public Function Json_Unflatten(ByVal flatObj As Collection) As Collection
    If Not Json_IsObject(flatObj) Then
        Err.Raise vbObjectError + 900, ERR_SRC, "Unflatten expects tagged object"
    End If

    Dim root As New Collection
    root.Add TAG_OBJECT

    Dim i As Long
    For i = 2 To flatObj.count
        Dim pair As Variant
        pair = flatObj(i)

        Dim path As String
        path = CStr(pair(0))

        Dim value As Variant
        VarAssign value, pair(1)

        If path = "$" Then
            Dim vv As Variant
            If IsObject(value) Then
                Set vv = value
            Else
                vv = value
            End If

            root.Add Array("$", vv)
        Else
            Json_UnflattenInsert root, path, value
        End If
    Next i

    Set Json_Unflatten = root
End Function

' =============================================================================
' Public API: Array-of-Object Root Discovery
' =============================================================================

' Scan flat paths and return candidate roots for array-of-object tables.
' Returned roots are unique, insertion-ordered.
'
' Errors:
'   vbObjectError + 910 flatObj not tagged object
Public Function Json_FindArrayObjectRoots( _
    ByVal flatObj As Collection, _
    Optional ByVal stopAfterFirst As Boolean = False _
) As Collection

    If Not Json_IsObject(flatObj) Then
        Err.Raise vbObjectError + 910, ERR_SRC, "FindArrayObjectRoots expects tagged object"
    End If

    Dim roots As New Collection

    Dim cap As Long
    Dim slotHash() As Long
    Dim slotIdx() As Long
    cap = 0 ' lazy init

    Dim i As Long
    For i = 2 To flatObj.count
        Dim pair As Variant
        pair = flatObj(i)

        Dim path As String
        path = CStr(pair(0))

        Json_CollectArrayObjectRootsFromPath_Fast roots, cap, slotHash, slotIdx, path, stopAfterFirst

        If stopAfterFirst Then
            If roots.count > 0 Then Exit For
        End If
    Next i

    Set Json_FindArrayObjectRoots = roots
End Function

' =============================================================================
' Public API: Table Extraction and 2D Conversion
' =============================================================================

' =============================================================================
' Json_ObjSet
'
' Purpose:
'   Set key=value on a tagged JSON object Collection, overwriting if the key exists.
'
' Model Contract:
'   - obj must be a tagged object Collection: obj(1)=TAG_OBJECT
'   - Each member is stored as Array(key,value) (preferred).
'
' Behavior:
'   - If key exists, removes existing entry and reinserts at the same position
'     to preserve deterministic relative order.
'   - If key does not exist, appends the new pair at the end.
'
' Notes:
'   - Uses binary compare for keys (case-sensitive).
' =============================================================================
Public Sub Json_ObjSet(ByVal obj As Collection, ByVal key As String, ByVal value As Variant)

    Dim i As Long
    Dim vv As Variant

    If IsObject(value) Then
        Set vv = value
    Else
        vv = value
    End If

    For i = 2 To obj.count

        Dim entry As Variant
        entry = obj(i)

        If IsArray(entry) Then

            If CStr(entry(LBound(entry))) = key Then

                obj.Remove i

                If i - 1 >= 1 Then
                    obj.Add Array(key, vv), , , i - 1
                Else
                    obj.Add Array(key, vv), , 2
                End If

                Exit Sub

            End If

        End If

    Next i

    obj.Add Array(key, vv)

End Sub

' Convert a Collection of tagged row objects into:
'   - headers: 1-based Variant array of column names (first-seen order)
'   - return : 2D Variant array (1..rowCount, 1..colCount) or Empty if no rows
'
' Behavior for 0 rows:
'   headers => ["value"]
'   return  => Empty
Public Function Json_TableTo2D(ByVal rows As Collection, ByRef headers As Variant) As Variant
    Const DBG As Boolean = False

    Dim rowCount As Long
    rowCount = rows.count

    If rowCount = 0 Then
        ReDim headers(1 To 1) As Variant
        headers(1) = "value"
        Json_TableTo2D = Empty
        Exit Function
    End If

    ' 1) Collect headers (first-seen order)
    Dim hdrs() As String
    Dim hdrCount As Long
    hdrCount = 0

    Dim cap As Long
    cap = 64

    Dim slotHash() As Long
    Dim slotIdx() As Long
    ReDim slotHash(0 To cap - 1) As Long
    ReDim slotIdx(0 To cap - 1) As Long

    Dim r As Long
    For r = 1 To rowCount
        Dim rowObj As Collection
        Set rowObj = rows(r)

        Dim p As Long
        For p = 2 To rowObj.count
            Dim pair As Variant
            pair = rowObj(p)

            Dim k As String
            k = CStr(pair(0))

            HeaderTable_Ensure k, hdrs, hdrCount, slotHash, slotIdx, cap, DBG
        Next p
    Next r

    ' Rows exist but no keys
    If hdrCount = 0 Then
        ReDim headers(1 To 1) As Variant
        headers(1) = "value"

        Dim data0 As Variant
        ReDim data0(1 To rowCount, 1 To 1) As Variant
        Json_TableTo2D = data0
        Exit Function
    End If

    ReDim headers(1 To hdrCount) As Variant
    Dim c As Long
    For c = 1 To hdrCount
        headers(c) = hdrs(c)
    Next c

    ' 2) Allocate and fill data
    Dim data As Variant
    ReDim data(1 To rowCount, 1 To hdrCount) As Variant

    For r = 1 To rowCount
        Dim rowObj2 As Collection
        Set rowObj2 = rows(r)

        Dim p2 As Long
        For p2 = 2 To rowObj2.count
            Dim pair2 As Variant
            pair2 = rowObj2(p2)

            Dim k2 As String
            k2 = CStr(pair2(0))

            Dim col2 As Long
            col2 = HeaderTable_Find(k2, hdrs, slotHash, slotIdx, cap)

            If col2 > 0 Then
                data(r, col2) = pair2(1)
            Else
                If DBG Then Debug.Print "WARN: header not found for key=" & k2
            End If
        Next p2
    Next r

    Json_TableTo2D = data
End Function

' Extract table rows (tagged objects) from a flattened object, using tableRoot.
'
' tableRoot examples:
'   "$"                for root array
'   "$.orders.items"   for nested arrays (supports parent indices in paths)
'
' Errors:
'   vbObjectError + 920 flatObj not tagged object
Public Function Json_ExtractTableRows(ByVal flatObj As Collection, ByVal tableRoot As String) As Collection
    If Not Json_IsObject(flatObj) Then
        Err.Raise vbObjectError + 920, ERR_SRC, "ExtractTableRows expects tagged object"
    End If

    Dim rows As New Collection
    Dim map As RowKeyMap

    ' Compile tableRoot once for nested roots performance
    Dim rootSegs() As String
    Dim rootSegCount As Long
    Json_BuildRootSegs tableRoot, rootSegs, rootSegCount

    Dim i As Long
    For i = 2 To flatObj.count
        Dim kv As Variant
        kv = flatObj(i)

        Dim path As String
        path = CStr(kv(0))

        Dim idx As Long
        Dim colPath As String
        Dim rowKey As String
        Dim ok As Boolean

        Dim usedIndexedFastPath As Boolean
        usedIndexedFastPath = False

        ' Fast path: root indexed immediately after tableRoot
        If Left$(path, Len(tableRoot) + 1) = (tableRoot & "[") Then
            ok = Json_TryParseIndexedPath(path, tableRoot, idx, colPath, rowKey)
            usedIndexedFastPath = ok
        Else
            ok = Json_TryParseTableRowPath_Compiled(path, tableRoot, rootSegs, rootSegCount, idx, colPath, rowKey)
        End If

        If ok Then
            ' Exclude child-table columns
            If InStr(1, colPath, "[", vbBinaryCompare) = 0 Then
                Dim rowObj As Collection

                If usedIndexedFastPath Then
                    Set rowObj = Json_EnsureRow(rows, idx)
                Else
                    Set rowObj = RowKeyMap_GetOrAdd(map, rowKey, rows)
                End If

                Dim v As Variant
                VarAssign v, kv(1)
                Json_ObjSet rowObj, colPath, v
            End If
        End If
    Next i

    Set Json_ExtractTableRows = rows
End Function

' =============================================================================
' Public API: Excel ListObject Upsert
' =============================================================================

' =============================================================================
' Excel_GetListObject
'
' Purpose:
'   Locate a ListObject on a worksheet by name (case-insensitive).
'
' Returns:
'   - The ListObject if found
'   - Nothing if not found
'
' Notes:
'   - Only searches ws.ListObjects (does not search other sheets).
' =============================================================================
Public Function Excel_GetListObject(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If StrComp(lo.name, tableName, vbTextCompare) = 0 Then
            Set Excel_GetListObject = lo
            Exit Function
        End If
    Next lo
    Set Excel_GetListObject = Nothing
End Function

' =============================================================================
' Excel_EnsureListObject
'
' Purpose:
'   Ensure a ListObject exists on ws with the given tableName.
'   If missing, create it starting at topLeft with the provided headers.
'
' Parameters:
'   headers:
'     1D array of header names; validated for blanks and duplicates.
'
' Returns:
'   The existing or newly created ListObject.
'
' Notes:
'   - New table is created with a single header row range; body is empty.
' =============================================================================
Public Function Excel_EnsureListObject( _
    ByVal ws As Worksheet, _
    ByVal tableName As String, _
    ByVal topLeft As Range, _
    ByVal headers As Variant _
) As ListObject

    Dim lo As ListObject
    Set lo = Excel_GetListObject(ws, tableName)

    If lo Is Nothing Then
        Excel_ValidateHeaders headers, "Excel_EnsureListObject"

        Dim colCount As Long
        colCount = UBound(headers) - LBound(headers) + 1

        Dim headerRange As Range
        Set headerRange = ws.Range(topLeft, topLeft.Offset(0, colCount - 1))

        headerRange.Value2 = Excel_HeadersTo2D(headers)

        Set lo = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=headerRange, XlListObjectHasHeaders:=xlYes)
        lo.name = tableName
    End If

    Set Excel_EnsureListObject = lo
End Function

Private Sub Excel_ListObjectUpsertData( _
    ByVal lo As ListObject, _
    ByVal headers As Variant, _
    ByVal data2D As Variant, _
    Optional ByVal clearExisting As Boolean = True, _
    Optional ByVal addMissingColumns As Boolean = True, _
    Optional ByVal removeMissingColumns As Boolean = False, _
    Optional ByVal preserveFormulaColumns As Boolean = True, _
    Optional ByVal fillFormulasOnAppend As Boolean = True _
)
    Const ERR_SRC As String = "Excel_ListObjectUpsertData"
    Const ERR_SHEET_BOUNDS As Long = vbObjectError + 1102

    If removeMissingColumns And (Not clearExisting) Then
        Err.Raise vbObjectError + 1101, ERR_SRC, _
            "removeMissingColumns=True requires clearExisting=True (schema shrink would corrupt existing rows)."
    End If

    Dim calcOld As XlCalculation
    Dim eventsOld As Boolean
    Dim updatingOld As Boolean
    Dim statusBarOld As Variant

    calcOld = Application.Calculation
    eventsOld = Application.EnableEvents
    updatingOld = Application.ScreenUpdating
    statusBarOld = Application.StatusBar

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = False

    Dim fHdrs() As String
    Dim fFmls() As String
    Dim fCount As Long
    If preserveFormulaColumns Then
        Excel_CaptureFormulaTemplates lo, fHdrs, fFmls, fCount
    Else
        fCount = 0
    End If

    Dim oldCols As Long
    oldCols = lo.ListColumns.count

    Dim oldBodyRows As Long
    Dim oldBody As Range
    If lo.DataBodyRange Is Nothing Then
        oldBodyRows = 0
    Else
        Set oldBody = lo.DataBodyRange
        oldBodyRows = oldBody.rows.count
    End If

    Dim existingHeaders As Variant
    existingHeaders = Excel_ListObjectHeadersTo1D(lo)

    Dim finalHeaders As Variant
    Dim finalData As Variant

    Dim incomingCount As Long
    incomingCount = UBound(headers) - LBound(headers) + 1

    Dim existingCount As Long
    existingCount = UBound(existingHeaders) - LBound(existingHeaders) + 1

    Dim sameSchema As Boolean
    sameSchema = False

    If incomingCount = existingCount Then
        sameSchema = True

        Dim hs As Long
        For hs = 1 To existingCount
            If StrComp( _
                Trim$(CStr(headers(LBound(headers) + hs - 1))), _
                Trim$(CStr(existingHeaders(LBound(existingHeaders) + hs - 1))), _
                vbTextCompare _
            ) <> 0 Then
                sameSchema = False
                Exit For
            End If
        Next hs
    End If

    If removeMissingColumns Then
        finalHeaders = headers
        finalData = data2D
    ElseIf addMissingColumns Then
        If sameSchema Then
            finalHeaders = existingHeaders
            finalData = data2D
        Else
            finalHeaders = Excel_UnionHeadersFromListObject(lo, headers)

            Dim finalCount1 As Long
            finalCount1 = UBound(finalHeaders) - LBound(finalHeaders) + 1

            If finalCount1 = incomingCount Then
                Dim unionMatchesIncoming As Boolean
                unionMatchesIncoming = True

                Dim hu As Long
                For hu = 1 To incomingCount
                    If StrComp( _
                        Trim$(CStr(headers(LBound(headers) + hu - 1))), _
                        Trim$(CStr(finalHeaders(LBound(finalHeaders) + hu - 1))), _
                        vbTextCompare _
                    ) <> 0 Then
                        unionMatchesIncoming = False
                        Exit For
                    End If
                Next hu

                If unionMatchesIncoming Then
                    finalData = data2D
                Else
                    finalData = Excel_ReshapeDataToHeaders(headers, finalHeaders, data2D)
                End If
            Else
                finalData = Excel_ReshapeDataToHeaders(headers, finalHeaders, data2D)
            End If
        End If
    Else
        finalHeaders = existingHeaders

        If sameSchema Then
            finalData = data2D
        Else
            finalData = Excel_ReshapeDataToHeaders(headers, finalHeaders, data2D)
        End If
    End If

    Excel_ValidateHeaders finalHeaders, ERR_SRC

    Dim newBodyRows As Long
    newBodyRows = Excel_RowCount2D(finalData)

    If removeMissingColumns Then
        If newBodyRows = 0 Then
            If Excel_IsDefaultValueOnlyHeaders(headers) Then
                finalHeaders = existingHeaders
                finalData = Empty
                newBodyRows = 0
                removeMissingColumns = False
                addMissingColumns = False
                clearExisting = True
            Else
                finalHeaders = headers
                finalData = Empty
                newBodyRows = 0
            End If
        End If
    End If

    Dim newCols As Long
    newCols = UBound(finalHeaders) - LBound(finalHeaders) + 1

    Dim targetBodyRows As Long
    If clearExisting Then
        targetBodyRows = newBodyRows
    Else
        targetBodyRows = oldBodyRows + newBodyRows
    End If

    Dim headerRow As Long
    headerRow = lo.HeaderRowRange.row

    If newCols > lo.parent.Columns.count Then
        Err.Raise ERR_SHEET_BOUNDS, ERR_SRC, _
            "Target column count exceeds worksheet limit. cols=" & newCols & ", max=" & lo.parent.Columns.count
    End If

    If headerRow + targetBodyRows > lo.parent.rows.count Then
        Err.Raise ERR_SHEET_BOUNDS, ERR_SRC, _
            "Target row count exceeds worksheet limit. required_last_row=" & (headerRow + targetBodyRows) & _
            ", max=" & lo.parent.rows.count
    End If

    If newCols < oldCols Then
        Excel_ClearOrphanedColumns lo, newCols, oldCols, oldBodyRows
    End If

    Dim writeHeaders As Boolean
    writeHeaders = clearExisting Or (newCols <> oldCols)

    Dim header2D As Variant
    If writeHeaders Then
        header2D = Excel_HeadersTo2D(finalHeaders)
    End If

    Dim rowsPerChunk As Long
    rowsPerChunk = Excel_GetRowsPerChunk(newCols)

    If clearExisting Then
        If oldBodyRows > 0 Then oldBody.ClearContents

        Excel_ResizeTableToRowCol lo, finalHeaders, newBodyRows

        If writeHeaders Then
            lo.HeaderRowRange.Value2 = header2D
        End If

        If newBodyRows > 0 Then
            Dim fullBody As Range
            Set fullBody = lo.DataBodyRange

            If newBodyRows <= rowsPerChunk Then
                fullBody.Value2 = finalData
            Else
                Dim srcRowLb As Long
                Dim srcColLb As Long
                srcRowLb = LBound(finalData, 1)
                srcColLb = LBound(finalData, 2)

                Dim writeStart As Long
                writeStart = 1

                Do While writeStart <= newBodyRows
                    Dim takeRows As Long
                    takeRows = rowsPerChunk
                    If writeStart + takeRows - 1 > newBodyRows Then
                        takeRows = newBodyRows - writeStart + 1
                    End If

                    Dim chunkData As Variant
                    ReDim chunkData(1 To takeRows, 1 To newCols)

                    Dim rr As Long
                    Dim cc As Long
                    For rr = 1 To takeRows
                        For cc = 1 To newCols
                            chunkData(rr, cc) = finalData(srcRowLb + writeStart + rr - 2, srcColLb + cc - 1)
                        Next cc
                    Next rr

                    fullBody.Cells(writeStart, 1).Resize(takeRows, newCols).Value2 = chunkData
                    Erase chunkData
                    writeStart = writeStart + takeRows
                Loop
            End If
        End If

        If IsArray(finalData) Then Erase finalData

        If preserveFormulaColumns And fCount > 0 Then
            Excel_ApplyFormulasToBody lo, finalHeaders, newBodyRows, fHdrs, fFmls, fCount
        End If
    Else
        Dim startRow As Long
        startRow = oldBodyRows

        Excel_ResizeTableToRowCol lo, finalHeaders, targetBodyRows

        If writeHeaders Then
            lo.HeaderRowRange.Value2 = header2D
        End If

        If newBodyRows > 0 Then
            Dim appendBody As Range
            Set appendBody = lo.DataBodyRange

            If newBodyRows <= rowsPerChunk Then
                appendBody.Cells(startRow + 1, 1).Resize(newBodyRows, newCols).Value2 = finalData
            Else
                Dim srcRowLb2 As Long
                Dim srcColLb2 As Long
                srcRowLb2 = LBound(finalData, 1)
                srcColLb2 = LBound(finalData, 2)

                Dim appendStart As Long
                appendStart = 1

                Do While appendStart <= newBodyRows
                    Dim takeRows2 As Long
                    takeRows2 = rowsPerChunk
                    If appendStart + takeRows2 - 1 > newBodyRows Then
                        takeRows2 = newBodyRows - appendStart + 1
                    End If

                    Dim chunkData2 As Variant
                    ReDim chunkData2(1 To takeRows2, 1 To newCols)

                    Dim rr2 As Long
                    Dim cc2 As Long
                    For rr2 = 1 To takeRows2
                        For cc2 = 1 To newCols
                            chunkData2(rr2, cc2) = finalData(srcRowLb2 + appendStart + rr2 - 2, srcColLb2 + cc2 - 1)
                        Next cc2
                    Next rr2

                    appendBody.Cells(startRow + appendStart, 1).Resize(takeRows2, newCols).Value2 = chunkData2
                    Erase chunkData2
                    appendStart = appendStart + takeRows2
                Loop
            End If
        End If

        If IsArray(finalData) Then Erase finalData

        If preserveFormulaColumns And fillFormulasOnAppend And fCount > 0 Then
            Excel_ApplyFormulasToAppendedRows lo, finalHeaders, startRow, newBodyRows, fHdrs, fFmls, fCount
        End If
    End If

    If newCols < oldCols Then
        Excel_ClearOrphanedHeaderOnly lo, newCols, oldCols
    End If

CleanExit:
    Application.StatusBar = statusBarOld
    Application.Calculation = calcOld
    Application.EnableEvents = eventsOld
    Application.ScreenUpdating = updatingOld
    Exit Sub

CleanFail:
    Application.StatusBar = statusBarOld
    Application.Calculation = calcOld
    Application.EnableEvents = eventsOld
    Application.ScreenUpdating = updatingOld
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub Excel_UpsertListObjectOnSheet( _
    ByVal ws As Worksheet, _
    ByVal tableName As String, _
    ByVal topLeft As Range, _
    ByVal headers As Variant, _
    ByVal data2D As Variant, _
    Optional ByVal clearExisting As Boolean = True, _
    Optional ByVal addMissingColumns As Boolean = True, _
    Optional ByVal removeMissingColumns As Boolean = False, _
    Optional ByVal preserveFormulaColumns As Boolean = True, _
    Optional ByVal fillFormulasOnAppend As Boolean = True _
)
    Dim lo As ListObject
    Set lo = Excel_GetListObject(ws, tableName)

    If lo Is Nothing Then
        Set lo = Excel_EnsureListObject(ws, tableName, topLeft, headers)
    End If

    Excel_ListObjectUpsertData lo, headers, data2D, _
        clearExisting, addMissingColumns, removeMissingColumns, _
        preserveFormulaColumns, fillFormulasOnAppend
End Sub

' =============================================================================
' Excel_ResizeTableToRowCol
'
' Purpose:
'   Resize a ListObject to the requested header/body shape while preserving
'   existing compensation for Excel table materialization edge cases.
'
' Behavior:
'   - Uses a single Resize call for the target shape.
'   - Preserves explicit DataBodyRange materialization fallback when Excel does
'     not fully realize body rows after Resize.
'   - For bodyRowCount=0, uses a temporary body row during resize, then deletes it.
'
' Errors:
'   vbObjectError + 1140 listobject has no HeaderRowRange
' =============================================================================
Public Sub Excel_ResizeTableToRowCol( _
    ByVal lo As ListObject, _
    ByVal finalHeaders As Variant, _
    ByVal bodyRowCount As Long _
)
    If Not lo.ShowHeaders Then lo.ShowHeaders = True
    If lo.HeaderRowRange Is Nothing Then
        Err.Raise vbObjectError + 1140, "Excel_ResizeTableToRowCol", _
            "ListObject has no HeaderRowRange (headers hidden or table corrupted): " & lo.name
    End If

    Dim headerTopLeft As Range
    Set headerTopLeft = lo.HeaderRowRange.Cells(1, 1)

    Dim colCount As Long
    colCount = UBound(finalHeaders) - LBound(finalHeaders) + 1

    Dim resizeBodyRows As Long
    resizeBodyRows = bodyRowCount
    If resizeBodyRows < 1 Then resizeBodyRows = 1

    Dim newRange As Range
    Set newRange = headerTopLeft.Resize(1 + resizeBodyRows, colCount)

    lo.Resize newRange

    If bodyRowCount <= 0 Then
        If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
        lo.HeaderRowRange.Value2 = Excel_HeadersTo2D(finalHeaders)
        Exit Sub
    End If

    Dim haveRows As Long
    If lo.DataBodyRange Is Nothing Then
        haveRows = 0
    Else
        haveRows = lo.DataBodyRange.rows.count
    End If

    If haveRows < bodyRowCount Then
        Dim needRows As Long
        needRows = bodyRowCount - haveRows

        Dim i As Long
        For i = 1 To needRows
            lo.ListRows.Add
        Next i
    End If
End Sub

' =============================================================================
' Excel_UpsertListObjectFromJsonAtRoot
'
' Purpose:
'   Parse jsonText, resolve a tableRoot path to an array-of-objects (or null),
'   extract rows, convert to 2D data, and upsert into an Excel ListObject.
'
' Key Behavior:
'   - Table extraction is driven by tableRoot (JSONPath-like).
'   - Nested arrays that are NOT part of the tableRoot path can be:
'       * excluded entirely (nonTableArraysAsJson=False)
'       * stored as JSON text in the cell (nonTableArraysAsJson=True)
'
' Parameters:
'   nonTableArraysAsJson:
'     False => exclude non-table arrays from flattening (prevents explosion)
'     True  => stringify non-table arrays into the cell as JSON text
'
' Notes:
'   - tableRoot may include indices (e.g., "$[0].items"); flattening normalizes
'     indices away for ancestor detection, but extraction still uses the real root.
' =============================================================================
Public Sub Excel_UpsertListObjectFromJsonAtRoot( _
    ByVal ws As Worksheet, _
    ByVal tableName As String, _
    ByVal topLeft As Range, _
    ByVal jsonText As String, _
    ByVal tableRoot As String, _
    Optional ByVal clearExisting As Boolean = True, _
    Optional ByVal addMissingColumns As Boolean = True, _
    Optional ByVal removeMissingColumns As Boolean = False, _
    Optional ByVal preserveFormulaColumns As Boolean = True, _
    Optional ByVal fillFormulasOnAppend As Boolean = True, _
    Optional ByVal nonTableArraysAsJson As Boolean = False _
)
    Const SRC As String = "Excel_UpsertListObjectFromJsonAtRoot"

    On Error GoTo Fail

    Dim parsed As Variant
    Json_ParseInto jsonText, parsed

    If (Not IsObject(parsed)) Or (TypeName(parsed) <> "Collection") Then
        Err.Raise vbObjectError + 1130, SRC, _
            "JSON root must be an object or array (Collection). Primitive root is not supported for table upsert."
    End If

    Dim resolved As Variant
    If Not Json_TryResolvePath(parsed, tableRoot, resolved) Then
        Err.Raise vbObjectError + 1160, SRC, "tableRoot not found: " & tableRoot
    End If

    If Not IsNull(resolved) Then
        If (Not IsObject(resolved)) _
            Or (TypeName(resolved) <> "Collection") _
            Or Json_IsObject(resolved) Then
            Err.Raise vbObjectError + 1162, SRC, _
                "tableRoot must resolve to an array-of-objects (or null): " & tableRoot
        End If
    End If

    Dim arr As Collection
    If IsNull(resolved) Then
        Set arr = New Collection
    Else
        Set arr = resolved
    End If

    Dim rowCount As Long
    rowCount = arr.count

    Dim i As Long
    For i = 1 To rowCount
        Dim elem As Variant
        VarAssign elem, arr(i)

        If (Not IsObject(elem)) _
            Or (TypeName(elem) <> "Collection") _
            Or (Not Json_IsObject(elem)) Then
            Err.Raise vbObjectError + 1163, SRC, _
                "Array element at index " & (i - 1) & " is not an object for root: " & tableRoot
        End If
    Next i

    Dim hdrs() As String
    Dim hdrCount As Long
    Dim cap As Long
    Dim slotHash() As Long
    Dim slotIdx() As Long

    cap = 64
    ReDim slotHash(0 To cap - 1) As Long
    ReDim slotIdx(0 To cap - 1) As Long
    hdrCount = 0

    For i = 1 To rowCount
        Dim rowObj As Collection
        Set rowObj = arr(i)
        Json_RowObjectCollectHeaders rowObj, vbNullString, nonTableArraysAsJson, hdrs, hdrCount, slotHash, slotIdx, cap
    Next i

    Dim headersOut As Variant
    If hdrCount = 0 Then
        ReDim headersOut(1 To 1) As Variant
        headersOut(1) = "value"
    Else
        ReDim headersOut(1 To hdrCount) As Variant

        Dim hc As Long
        For hc = 1 To hdrCount
            headersOut(hc) = hdrs(hc)
        Next hc
    End If

    If removeMissingColumns And rowCount = 0 Then
        Dim loExisting As ListObject
        Set loExisting = Excel_GetListObject(ws, tableName)

        If Not loExisting Is Nothing Then
            headersOut = Excel_ListObjectHeadersTo1D(loExisting)
            addMissingColumns = False
            removeMissingColumns = False
            clearExisting = True
        End If
    End If

    If rowCount = 0 Then
        Dim emptyData As Variant
        emptyData = Empty

        Excel_UpsertListObjectOnSheet ws, tableName, topLeft, _
            headersOut, emptyData, _
            clearExisting, addMissingColumns, removeMissingColumns, _
            preserveFormulaColumns, fillFormulasOnAppend

        Exit Sub
    End If

    Dim colCount As Long
    colCount = UBound(headersOut) - LBound(headersOut) + 1

    Dim rowsPerChunk As Long
    rowsPerChunk = Excel_GetRowsPerChunk(colCount)

    Dim chunkData As Variant
    Dim chunkRows As Long
    Dim chunkRowPos As Long
    Dim rowIndex As Long
    Dim firstWrite As Boolean

    firstWrite = True
    chunkRows = rowCount
    If chunkRows > rowsPerChunk Then chunkRows = rowsPerChunk
    ReDim chunkData(1 To chunkRows, 1 To colCount)

    chunkRowPos = 0

    For rowIndex = 1 To rowCount
        If chunkRowPos = chunkRows Then
            Excel_UpsertListObjectOnSheet ws, tableName, topLeft, _
                headersOut, chunkData, _
                IIf(firstWrite, clearExisting, False), _
                IIf(firstWrite, addMissingColumns, False), _
                IIf(firstWrite, removeMissingColumns, False), _
                preserveFormulaColumns, fillFormulasOnAppend

            firstWrite = False
            Erase chunkData

            Dim remainingRows As Long
            remainingRows = rowCount - rowIndex + 1
            chunkRows = remainingRows
            If chunkRows > rowsPerChunk Then chunkRows = rowsPerChunk
            ReDim chunkData(1 To chunkRows, 1 To colCount)

            chunkRowPos = 0
        End If

        chunkRowPos = chunkRowPos + 1
        Set rowObj = arr(rowIndex)
        Json_RowObjectFillRow rowObj, vbNullString, nonTableArraysAsJson, hdrs, slotHash, slotIdx, cap, chunkData, chunkRowPos
    Next rowIndex

    If chunkRowPos > 0 Then
        If chunkRowPos < chunkRows Then
            Dim lastChunk As Variant
            ReDim lastChunk(1 To chunkRowPos, 1 To colCount)

            Dim rr As Long
            Dim cc As Long
            For rr = 1 To chunkRowPos
                For cc = 1 To colCount
                    lastChunk(rr, cc) = chunkData(rr, cc)
                Next cc
            Next rr

            Excel_UpsertListObjectOnSheet ws, tableName, topLeft, _
                headersOut, lastChunk, _
                IIf(firstWrite, clearExisting, False), _
                IIf(firstWrite, addMissingColumns, False), _
                IIf(firstWrite, removeMissingColumns, False), _
                preserveFormulaColumns, fillFormulasOnAppend

            Erase lastChunk
        Else
            Excel_UpsertListObjectOnSheet ws, tableName, topLeft, _
                headersOut, chunkData, _
                IIf(firstWrite, clearExisting, False), _
                IIf(firstWrite, addMissingColumns, False), _
                IIf(firstWrite, removeMissingColumns, False), _
                preserveFormulaColumns, fillFormulasOnAppend
        End If
    End If

    Erase chunkData
    Exit Sub

Fail:
    Dim n As Long
    Dim d As String
    Dim s As String

    n = Err.Number
    d = Err.Description
    s = Err.Source

    Err.Clear
    If Len(s) > 0 And StrComp(s, SRC, vbBinaryCompare) <> 0 Then
        d = d & " | inner_source=" & s
    End If

    Err.Raise n, SRC, d
End Sub

' =============================================================================
' JSON Reader: primitives, strings, arrays, objects
' =============================================================================

Private Sub JR_Init(ByRef r As JsonReader, ByVal jsonText As String)
    r.text = jsonText
    r.pos = 1
End Sub

Private Function JR_Eof(ByRef r As JsonReader) As Boolean
    JR_Eof = (r.pos > Len(r.text))
End Function

Private Function JR_Peek(ByRef r As JsonReader) As String
    If JR_Eof(r) Then
        JR_Peek = vbNullString
    Else
        JR_Peek = Mid$(r.text, r.pos, 1)
    End If
End Function

Private Function JR_Next(ByRef r As JsonReader) As String
    Dim ch As String
    ch = JR_Peek(r)
    If Not JR_Eof(r) Then r.pos = r.pos + 1
    JR_Next = ch
End Function

Private Sub JR_SkipWs(ByRef r As JsonReader)
    Do While Not JR_Eof(r)
        Select Case JR_Peek(r)
            Case " ", vbTab, vbCr, vbLf
                r.pos = r.pos + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

Private Sub JR_ExpectChar(ByRef r As JsonReader, ByVal expected As String)
    JR_SkipWs r

    Dim ch As String
    ch = JR_Next(r)

    If ch <> expected Then
        Err.Raise vbObjectError + 520, ERR_SRC, _
            "Expected '" & expected & "' at pos " & (r.pos - 1) & " but got '" & ch & "'"
    End If
End Sub

Private Sub JR_ExpectLiteral(ByRef r As JsonReader, ByVal lit As String)
    JR_SkipWs r

    Dim i As Long
    For i = 1 To Len(lit)
        If JR_Next(r) <> Mid$(lit, i, 1) Then
            Err.Raise vbObjectError + 525, ERR_SRC, _
                "Expected literal '" & lit & "' near pos " & (r.pos - 1)
        End If
    Next i
End Sub

Private Sub Json_ReadValue(ByRef r As JsonReader, ByRef outValue As Variant)
    JR_SkipWs r

    Dim ch As String
    ch = JR_Peek(r)

    Select Case ch
        Case """"
            outValue = JR_ReadJsonString(r)

        Case "t"
            JR_ExpectLiteral r, "true"
            outValue = True

        Case "f"
            JR_ExpectLiteral r, "false"
            outValue = False

        Case "n"
            JR_ExpectLiteral r, "null"
            outValue = Null

        Case "-", "0" To "9"
            outValue = JR_ReadNumber(r)

        Case "["
            Dim arr As Collection
            Set arr = JR_ReadArray(r)
            Set outValue = arr

        Case "{"
            Dim obj As Collection
            Set obj = JR_ReadObject(r)
            Set outValue = obj

        Case Else
            Err.Raise vbObjectError + 701, ERR_SRC, _
                "Unexpected token '" & ch & "' at pos " & r.pos
    End Select
End Sub

Private Function JR_ReadNumber(ByRef r As JsonReader) As Variant
    JR_SkipWs r

    Dim startPos As Long
    startPos = r.pos

    If JR_Peek(r) = "-" Then JR_Next r

    Dim ch As String
    ch = JR_Peek(r)

    If ch = "0" Then
        JR_Next r
    ElseIf ch >= "1" And ch <= "9" Then
        Do While JR_Peek(r) >= "0" And JR_Peek(r) <= "9"
            JR_Next r
        Loop
    Else
        Err.Raise vbObjectError + 710, ERR_SRC, "Invalid number at pos " & r.pos
    End If

    If JR_Peek(r) = "." Then
        JR_Next r
        If Not (JR_Peek(r) >= "0" And JR_Peek(r) <= "9") Then
            Err.Raise vbObjectError + 711, ERR_SRC, "Invalid fractional part"
        End If
        Do While JR_Peek(r) >= "0" And JR_Peek(r) <= "9"
            JR_Next r
        Loop
    End If

    If JR_Peek(r) = "e" Or JR_Peek(r) = "E" Then
        JR_Next r
        If JR_Peek(r) = "+" Or JR_Peek(r) = "-" Then JR_Next r

        If Not (JR_Peek(r) >= "0" And JR_Peek(r) <= "9") Then
            Err.Raise vbObjectError + 712, ERR_SRC, "Invalid exponent"
        End If

        Do While JR_Peek(r) >= "0" And JR_Peek(r) <= "9"
            JR_Next r
        Loop
    End If

    Dim numText As String
    numText = Mid$(r.text, startPos, r.pos - startPos)

    If InStr(1, numText, ".", vbBinaryCompare) = 0 And InStr(1, numText, "e", vbTextCompare) = 0 Then
        On Error Resume Next
        Dim L As Long
        L = CLng(numText)
        If Err.Number = 0 Then
            JR_ReadNumber = L
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
    End If

    JR_ReadNumber = CDbl(numText)
End Function

Private Function JR_ReadArray(ByRef r As JsonReader) As Collection
    JR_SkipWs r
    JR_ExpectChar r, "["

    Dim result As New Collection
    JR_SkipWs r

    If JR_Peek(r) = "]" Then
        JR_Next r
        Set JR_ReadArray = result
        Exit Function
    End If

    Do
        Dim value As Variant
        Json_ReadValue r, value
        result.Add value

        JR_SkipWs r

        Dim ch As String
        ch = JR_Peek(r)

        If ch = "," Then
            JR_Next r
        ElseIf ch = "]" Then
            JR_Next r
            Exit Do
        Else
            Err.Raise vbObjectError + 730, ERR_SRC, "Expected ',' or ']' at pos " & r.pos
        End If
    Loop

    Set JR_ReadArray = result
End Function

Private Function JR_ReadObject(ByRef r As JsonReader) As Collection
    JR_SkipWs r
    JR_ExpectChar r, "{"

    Dim obj As New Collection
    obj.Add TAG_OBJECT

    JR_SkipWs r

    If JR_Peek(r) = "}" Then
        JR_Next r
        Set JR_ReadObject = obj
        Exit Function
    End If

    Do
        Dim key As String
        key = JR_ReadJsonString(r)

        JR_SkipWs r
        JR_ExpectChar r, ":"

        Dim value As Variant
        Json_ReadValue r, value

        Dim vv As Variant
        If IsObject(value) Then
            Set vv = value
        Else
            vv = value
        End If

        ' IMPORTANT: use Array(...) not a fixed-size local array
        obj.Add Array(key, vv)

        JR_SkipWs r

        Dim ch As String
        ch = JR_Peek(r)

        If ch = "," Then
            JR_Next r
        ElseIf ch = "}" Then
            JR_Next r
            Exit Do
        Else
            Err.Raise vbObjectError + 760, ERR_SRC, "Expected ',' or '}' at pos " & r.pos
        End If
    Loop

    Set JR_ReadObject = obj
End Function

' =============================================================================
' JSON String Parsing
' =============================================================================

Private Function JR_ReadJsonString(ByRef r As JsonReader) As String
    JR_SkipWs r
    JR_ExpectChar r, """"

    Dim parts() As String
    Dim partCount As Long
    ReDim parts(0 To 31)
    partCount = 0

    Do While Not JR_Eof(r)
        Dim ch As String
        ch = JR_Next(r)

        If ch = """" Then
            JR_ReadJsonString = JR_JoinParts(parts, partCount)
            Exit Function
        End If

        If ch = "\" Then
            If JR_Eof(r) Then Err.Raise vbObjectError + 521, ERR_SRC, "Unterminated escape at end of input"

            Dim esc As String
            esc = JR_Next(r)

            Select Case esc
                Case """": JR_AddPart parts, partCount, """"
                Case "\":  JR_AddPart parts, partCount, "\"
                Case "/":  JR_AddPart parts, partCount, "/"
                Case "b":  JR_AddPart parts, partCount, Chr$(8)
                Case "f":  JR_AddPart parts, partCount, Chr$(12)
                Case "n":  JR_AddPart parts, partCount, vbLf
                Case "r":  JR_AddPart parts, partCount, vbCr
                Case "t":  JR_AddPart parts, partCount, vbTab
                Case "u":  JR_AddPart parts, partCount, JR_ReadUnicodeEscape(r)
                Case Else
                    Err.Raise vbObjectError + 522, ERR_SRC, _
                        "Invalid escape '\\" & esc & "' at pos " & (r.pos - 1)
            End Select
        Else
            Dim cc As Long
            cc = AscW(ch)
            If cc >= 0 And cc < 32 Then
                Err.Raise vbObjectError + 526, ERR_SRC, _
                    "Unescaped control character in string at pos " & (r.pos - 1)
            End If
            JR_AddPart parts, partCount, ch
        End If
    Loop

    Err.Raise vbObjectError + 523, ERR_SRC, "Unterminated string"
End Function

Private Function JR_ReadUnicodeEscape(ByRef r As JsonReader) As String
    Dim u1 As Long
    u1 = JR_ReadHex4ToLong(r)

    If u1 >= &HD800 And u1 <= &HDBFF Then
        If JR_Eof(r) Then Err.Raise vbObjectError + 527, ERR_SRC, "Invalid surrogate pair (incomplete)"

        If JR_Next(r) <> "\" Then Err.Raise vbObjectError + 527, ERR_SRC, "Invalid surrogate pair (expected \u)"
        If JR_Eof(r) Then Err.Raise vbObjectError + 527, ERR_SRC, "Invalid surrogate pair (incomplete)"
        If JR_Next(r) <> "u" Then Err.Raise vbObjectError + 527, ERR_SRC, "Invalid surrogate pair (expected \u)"

        Dim u2 As Long
        u2 = JR_ReadHex4ToLong(r)

        If u2 < &HDC00 Or u2 > &HDFFF Then
            Err.Raise vbObjectError + 527, ERR_SRC, "Invalid surrogate pair (low surrogate out of range)"
        End If

        JR_ReadUnicodeEscape = ChrW$(u1) & ChrW$(u2)
        Exit Function
    End If

    If u1 >= &HDC00 And u1 <= &HDFFF Then
        Err.Raise vbObjectError + 527, ERR_SRC, "Invalid surrogate pair (unexpected low surrogate)"
    End If

    JR_ReadUnicodeEscape = ChrW$(u1)
End Function

Private Function JR_ReadHex4ToLong(ByRef r As JsonReader) As Long
    Dim hex4 As String
    hex4 = vbNullString

    Dim i As Long
    For i = 1 To 4
        If JR_Eof(r) Then Err.Raise vbObjectError + 524, ERR_SRC, "Incomplete \uXXXX escape"

        Dim ch As String
        ch = JR_Next(r)

        If Not JR_IsHexDigit(ch) Then
            Err.Raise vbObjectError + 524, ERR_SRC, "Invalid \uXXXX escape"
        End If

        hex4 = hex4 & ch
    Next i

    On Error GoTo BadHex
    JR_ReadHex4ToLong = CLng("&H" & hex4)
    Exit Function

BadHex:
    Err.Clear
    Err.Raise vbObjectError + 524, ERR_SRC, "Invalid \uXXXX escape"
End Function

Private Function JR_IsHexDigit(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    Select Case ch
        Case "0" To "9", "a" To "f", "A" To "F"
            JR_IsHexDigit = True
    End Select
End Function

Private Sub JR_AddPart(ByRef parts() As String, ByRef partCount As Long, ByVal s As String)
    If partCount > UBound(parts) Then
        ReDim Preserve parts(0 To (UBound(parts) * 2) + 1)
    End If
    parts(partCount) = s
    partCount = partCount + 1
End Sub

Private Function JR_JoinParts(ByRef parts() As String, ByVal partCount As Long) As String
    If partCount = 0 Then
        JR_JoinParts = vbNullString
        Exit Function
    End If

    Dim tmp() As String
    ReDim tmp(0 To partCount - 1)

    Dim i As Long
    For i = 0 To partCount - 1
        tmp(i) = parts(i)
    Next i

    JR_JoinParts = Join(tmp, vbNullString)
End Function

' =============================================================================
' JSON Stringify internals
' =============================================================================

Private Function Json_StringifyArray(ByVal c As Collection) As String
    Dim parts() As String
    Dim partCount As Long
    ReDim parts(0 To 31)
    partCount = 0

    JS_AddPart parts, partCount, "["

    Dim i As Long
    For i = 1 To c.count
        If i > 1 Then JS_AddPart parts, partCount, ","
        JS_AddPart parts, partCount, Json_Stringify(c(i))
    Next i

    JS_AddPart parts, partCount, "]"
    Json_StringifyArray = JS_JoinParts(parts, partCount)
End Function

Private Function Json_StringifyObject(ByVal obj As Collection) As String
    Dim parts() As String
    Dim partCount As Long
    ReDim parts(0 To 63)
    partCount = 0
    
    If obj Is Nothing Then
        Err.Raise vbObjectError + 1134, ERR_SRC, _
            "Json_StringifyObject: object is Nothing."
    End If
    
    If obj.count < 1 Or CStr(obj(1)) <> TAG_OBJECT Then
        Err.Raise vbObjectError + 1134, ERR_SRC, _
            "Json_StringifyObject: collection is not a tagged object."
    End If
    
    JS_AddPart parts, partCount, "{"

    Dim first As Boolean
    first = True

    Dim i As Long
    i = 2 ' skip "__OBJ__"

    Do While i <= obj.count
        Dim entry As Variant
        entry = obj(i)

        Dim keyStr As String
        Dim val As Variant

        ' Case A: pair stored as 2-element array
        If IsArray(entry) Then
            Dim lb As Long, ub As Long
            lb = LBound(entry)
            ub = UBound(entry)

            If (ub - lb + 1) < 2 Then
                Err.Raise vbObjectError + 1136, ERR_SRC, _
                    "Json_StringifyObject: object pair at index " & CStr(i) & _
                    " must contain 2 elements (key,value)."
            End If

            keyStr = CStr(entry(lb))

            If IsObject(entry(lb + 1)) Then
                Set val = entry(lb + 1)
            Else
                val = entry(lb + 1)
            End If

            i = i + 1

        ' Case B: pair stored as 2-element Collection
        ElseIf IsObject(entry) And TypeName(entry) = "Collection" Then
            If entry.count < 2 Then
                Err.Raise vbObjectError + 1136, ERR_SRC, _
                    "Json_StringifyObject: object pair Collection at index " & CStr(i) & _
                    " must contain 2 elements (key,value)."
            End If

            keyStr = CStr(entry(1))

            If IsObject(entry(2)) Then
                Set val = entry(2)
            Else
                val = entry(2)
            End If

            i = i + 1

        ' Case C: alternating key/value representation: key is String at i, value at i+1
        ElseIf VarType(entry) = vbString Then
            keyStr = CStr(entry)

            If i = obj.count Then
                Err.Raise vbObjectError + 1136, ERR_SRC, _
                    "Json_StringifyObject: dangling key at final index " & CStr(i) & _
                    " (missing value)."
            End If

            If IsObject(obj(i + 1)) Then
                Set val = obj(i + 1)
            Else
                val = obj(i + 1)
            End If

            i = i + 2

        Else
            Err.Raise vbObjectError + 1135, ERR_SRC, _
                "Json_StringifyObject: object entry at index " & CStr(i) & _
                " is not Array(key,value) or Collection(key,value) or String(key). Found type=" & TypeName(entry)
        End If

        ' Emit JSON member
        If Not first Then JS_AddPart parts, partCount, ","
        first = False

        JS_AddPart parts, partCount, """"
        JS_AddPart parts, partCount, Json_EscapeString(keyStr)
        JS_AddPart parts, partCount, """:"
        JS_AddPart parts, partCount, Json_Stringify(val)
    Loop

    JS_AddPart parts, partCount, "}"
    Json_StringifyObject = JS_JoinParts(parts, partCount)
End Function

Private Sub JS_AddPart(ByRef parts() As String, ByRef partCount As Long, ByVal s As String)
    If partCount > UBound(parts) Then
        ReDim Preserve parts(0 To (UBound(parts) * 2) + 1)
    End If
    parts(partCount) = s
    partCount = partCount + 1
End Sub

Private Function JS_JoinParts(ByRef parts() As String, ByVal partCount As Long) As String
    If partCount = 0 Then
        JS_JoinParts = vbNullString
        Exit Function
    End If

    Dim tmp() As String
    ReDim tmp(0 To partCount - 1)

    Dim i As Long
    For i = 0 To partCount - 1
        tmp(i) = parts(i)
    Next i

    JS_JoinParts = Join(tmp, vbNullString)
End Function

Private Function Json_EscapeString(ByVal s As String) As String
    Dim i As Long, ch As String, code As Long, out As String
    out = vbNullString

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)

        Select Case ch
            Case """": out = out & "\"""   ' quote
            Case "\": out = out & "\\"     ' backslash
            Case "/": out = out & "\/"     ' optional
            Case vbBack: out = out & "\b"
            Case vbFormFeed: out = out & "\f"
            Case vbCr: out = out & "\r"
            Case vbLf: out = out & "\n"
            Case vbTab: out = out & "\t"
            Case Else
                If code >= 0 And code < 32 Then
                    out = out & "\u" & Right$("0000" & Hex$(code), 4)
                Else
                    out = out & ch
                End If
        End Select
    Next i

    Json_EscapeString = out
End Function

Private Function Json_NumberToString(ByVal d As Double) As String
    Dim s As String
    s = CStr(d)

    Dim decSep As String
    decSep = Mid$(CStr(1.1), 2, 1)

    If decSep <> "." Then s = Replace$(s, decSep, ".")
    Json_NumberToString = s
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

    ' ----------------------------------------------------
    ' ARRAYS
    ' ----------------------------------------------------
    If Json_IsArray(v) Then

        Dim arr As Collection
        Set arr = v

        Dim basePath As String
        basePath = IIf(Len(prefix) = 0, "$", prefix)

        ' legacy behavior
        If arrayMode = 0 Then

            Dim i As Long
            For i = 1 To arr.count

                Dim idxPath As String
                idxPath = basePath & "[" & (i - 1) & "]"

                Dim elem As Variant
                VarAssign elem, arr(i)

                If IsObject(elem) Then
                    If Json_IsObject(elem) Or Json_IsArray(elem) Then
                        Json_FlattenInto flat, idxPath, elem, depth + 1, maxDepth, tableRootNorm, arrayMode
                    Else
                        AddFlat flat, idxPath, Json_Stringify(elem)
                    End If
                Else
                    AddFlat flat, idxPath, elem
                End If

            Next i

            Exit Sub

        End If


        ' ----------------------------------------------------
        ' table-aware logic
        ' ----------------------------------------------------

        Dim baseNoIdx As String
        baseNoIdx = Json_RemoveIndices(basePath)

        Dim expandArray As Boolean
        expandArray = False

        If Len(tableRootNorm) > 0 Then

            ' EXACT table root
            If StrComp(baseNoIdx, tableRootNorm, vbBinaryCompare) = 0 Then
                expandArray = True

            ' ancestor of table root
            ElseIf Left$(tableRootNorm, Len(baseNoIdx) + 1) = (baseNoIdx & ".") Then
                expandArray = True

            End If

        End If


        If expandArray Then

            Dim j As Long
            For j = 1 To arr.count

                Dim idxPath2 As String
                idxPath2 = basePath & "[" & (j - 1) & "]"

                Dim elem2 As Variant
                VarAssign elem2, arr(j)

                If IsObject(elem2) Then
                    If Json_IsObject(elem2) Or Json_IsArray(elem2) Then
                        Json_FlattenInto flat, idxPath2, elem2, depth + 1, maxDepth, tableRootNorm, arrayMode
                    Else
                        AddFlat flat, idxPath2, Json_Stringify(elem2)
                    End If
                Else
                    AddFlat flat, idxPath2, elem2
                End If

            Next j

        Else

            ' array not part of table root path
            If arrayMode = 2 Then
                AddFlat flat, basePath, Json_Stringify(arr)
            End If

        End If

        Exit Sub

    End If


    ' ----------------------------------------------------
    ' OBJECTS
    ' ----------------------------------------------------
    If Json_IsObject(v) Then

        Dim obj As Collection
        Set obj = v

        Dim k As Long
        For k = 2 To obj.count

            Dim pair As Variant
            pair = obj(k)

            Dim seg As String
            seg = Json_EscapePathSegment(CStr(pair(0)))

            Dim nextPrefix As String
            If Len(prefix) = 0 Then
                nextPrefix = seg
            Else
                nextPrefix = prefix & "." & seg
            End If

            Dim child As Variant
            VarAssign child, pair(1)

            If IsObject(child) Then
                If Json_IsObject(child) Or Json_IsArray(child) Then
                    Json_FlattenInto flat, nextPrefix, child, depth + 1, maxDepth, tableRootNorm, arrayMode
                Else
                    AddFlat flat, nextPrefix, Json_Stringify(child)
                End If
            Else
                AddFlat flat, nextPrefix, child
            End If

        Next k

        Exit Sub

    End If


    ' fallback
    AddFlat flat, IIf(Len(prefix) = 0, "$", prefix), Json_Stringify(v)

End Sub

Private Sub AddFlat(ByVal flat As Collection, ByVal key As String, ByVal value As Variant)
    Dim vv As Variant
    If IsObject(value) Then
        Set vv = value
    Else
        vv = value
    End If

    flat.Add Array(key, vv)   ' <<< IMPORTANT
End Sub

Private Function Json_EscapePathSegment(ByVal s As String) As String
    s = Replace$(s, "\", "\\")
    s = Replace$(s, ".", "\.")
    Json_EscapePathSegment = s
End Function

Private Sub VarAssign(ByRef dest As Variant, ByVal SRC As Variant)
    If IsObject(SRC) Then
        Set dest = SRC
    Else
        dest = SRC
    End If
End Sub

' =============================================================================
' Unflatten internals
' =============================================================================

Private Sub Json_UnflattenInsert(ByVal root As Collection, ByVal path As String, ByVal value As Variant)
    If Left$(path, 2) = "$." Then
        path = Mid$(path, 3)
    End If

    If InStr(1, path, "[", vbBinaryCompare) > 0 Or InStr(1, path, "]", vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 905, ERR_SRC, "Unflatten does not support array index paths: " & path
    End If

    Dim tokens As Collection
    Set tokens = Json_TokenizePath(path)

    Dim current As Collection
    Set current = root

    Dim i As Long
    For i = 1 To tokens.count
        Dim key As String
        key = Json_UnescapePathSegment(CStr(tokens(i)))

        If i = tokens.count Then
            Json_ObjSet current, key, value
        Else
            Dim child As Collection
            Set child = Json_FindOrCreateChild(current, key)
            Set current = child
        End If
    Next i
End Sub

Private Function Json_TokenizePath(ByVal path As String) As Collection
    Dim tokens As New Collection
    Dim current As String
    current = vbNullString

    Dim i As Long
    i = 1

    Do While i <= Len(path)
        Dim ch As String
        ch = Mid$(path, i, 1)

        If ch = "\" Then
            If i < Len(path) Then
                current = current & ch & Mid$(path, i + 1, 1)
                i = i + 2
            Else
                current = current & ch
                i = i + 1
            End If
        ElseIf ch = "." Then
            tokens.Add current
            current = vbNullString
            i = i + 1
        Else
            current = current & ch
            i = i + 1
        End If
    Loop

    If Len(current) > 0 Then tokens.Add current
    Set Json_TokenizePath = tokens
End Function

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
    newObj.Add TAG_OBJECT

    parent.Add Array(key, newObj)  ' <<< IMPORTANT: Array(...), not Dim p(0 To 1)

    Set Json_FindOrCreateChild = newObj
End Function

Private Function Json_UnescapePathSegment(ByVal s As String) As String
    s = Replace$(s, "\.", ".")
    s = Replace$(s, "\\", "\")
    Json_UnescapePathSegment = s
End Function

' =============================================================================
' Root discovery internals: open addressing roots set
' =============================================================================

Private Sub RootsSet_Init(ByRef cap As Long, ByRef slotHash() As Long, ByRef slotIdx() As Long)
    cap = 64
    ReDim slotHash(0 To cap - 1) As Long
    ReDim slotIdx(0 To cap - 1) As Long
End Sub

Private Sub RootsSet_Rehash( _
    ByVal newCap As Long, _
    ByRef cap As Long, _
    ByRef slotHash() As Long, _
    ByRef slotIdx() As Long, _
    ByVal roots As Collection _
)
    Dim pow2 As Long
    pow2 = 1
    Do While pow2 < newCap
        pow2 = pow2 * 2
    Loop
    newCap = pow2

    Dim newHash() As Long
    Dim newIdx() As Long
    ReDim newHash(0 To newCap - 1) As Long
    ReDim newIdx(0 To newCap - 1) As Long

    Dim mask As Long
    mask = newCap - 1

    Dim i As Long
    For i = 1 To roots.count
        Dim s As String
        s = CStr(roots(i))

        Dim h As Long
        h = Json_Hash32_FNV1a(s)

        Dim pos As Long
        pos = (h And mask)

        Do
            If newIdx(pos) = 0 Then
                newHash(pos) = h
                newIdx(pos) = i
                Exit Do
            End If
            pos = (pos + 1) And mask
        Loop
    Next i

    cap = newCap
    slotHash = newHash
    slotIdx = newIdx
End Sub

Private Sub RootsSet_AddIfMissing( _
    ByVal s As String, _
    ByRef cap As Long, _
    ByRef slotHash() As Long, _
    ByRef slotIdx() As Long, _
    ByVal roots As Collection _
)
    If cap = 0 Then RootsSet_Init cap, slotHash, slotIdx

    If (roots.count + 1) * 10 > cap * 7 Then
        RootsSet_Rehash cap * 2, cap, slotHash, slotIdx, roots
    End If

    Dim h As Long
    h = Json_Hash32_FNV1a(s)

    Dim mask As Long
    mask = cap - 1

    Dim pos As Long
    pos = (h And mask)

    Do
        Dim existingIdx As Long
        existingIdx = slotIdx(pos)

        If existingIdx = 0 Then
            roots.Add s
            slotHash(pos) = h
            slotIdx(pos) = roots.count
            Exit Sub
        End If

        If slotHash(pos) = h Then
            If CStr(roots(existingIdx)) = s Then Exit Sub
        End If

        pos = (pos + 1) And mask
    Loop
End Sub

Private Sub Json_CollectArrayObjectRootsFromPath_Fast( _
    ByVal roots As Collection, _
    ByRef cap As Long, _
    ByRef slotHash() As Long, _
    ByRef slotIdx() As Long, _
    ByVal path As String, _
    ByVal stopAfterFirst As Boolean _
)
    ' Common case: root array-of-objects, paths like "$[0].id"
    If Len(path) >= 5 Then
        If Mid$(path, 1, 2) = "$[" Then
            If InStr(3, path, "].", vbBinaryCompare) > 0 Then
                RootsSet_AddIfMissing "$", cap, slotHash, slotIdx, roots
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
                    RootsSet_AddIfMissing rootPath, cap, slotHash, slotIdx, roots
                    If stopAfterFirst Then Exit Sub
                End If
            End If
        End If

        p = closePos + 1
    Loop
End Sub

Private Function Json_RemoveIndices(ByVal s As String) As String
    Dim out As String
    out = vbNullString

    Dim i As Long
    i = 1

    Do While i <= Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)

        If ch = "[" Then
            Dim j As Long
            j = InStr(i + 1, s, "]")
            If j = 0 Then
                out = out & Mid$(s, i)
                Exit Do
            End If

            Dim inside As String
            inside = Mid$(s, i + 1, j - i - 1)

            If Len(inside) > 0 And Json_IsAllDigits(inside) Then
                i = j + 1
            Else
                out = out & Mid$(s, i, (j - i + 1))
                i = j + 1
            End If
        Else
            out = out & ch
            i = i + 1
        End If
    Loop

    Json_RemoveIndices = out
End Function

Private Function Json_IsAllDigits(ByVal s As String) As Boolean
    Dim k As Long
    For k = 1 To Len(s)
        Dim ch As String
        ch = Mid$(s, k, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next k
    Json_IsAllDigits = (Len(s) > 0)
End Function

' =============================================================================
' Table path parsing internals
' =============================================================================

Private Function Json_TryParseIndexedPath( _
    ByVal fullPath As String, _
    ByVal tableRoot As String, _
    ByRef outIndex As Long, _
    ByRef outColPath As String, _
    ByRef outRowKey As String _
) As Boolean

    Json_TryParseIndexedPath = False
    outIndex = 0
    outColPath = vbNullString
    outRowKey = vbNullString

    Dim openPos As Long
    openPos = Len(tableRoot) + 1

    If openPos > Len(fullPath) Then Exit Function
    If Mid$(fullPath, openPos, 1) <> "[" Then Exit Function

    Dim closePos As Long
    closePos = InStr(openPos + 1, fullPath, "]")
    If closePos = 0 Then Exit Function

    Dim idxText As String
    idxText = Mid$(fullPath, openPos + 1, closePos - openPos - 1)
    If Len(idxText) = 0 Or Not Json_IsAllDigits(idxText) Then Exit Function

    outIndex = CLng(idxText)
    outRowKey = tableRoot & "[" & CStr(outIndex) & "]"

    Dim remainder As String
    remainder = Mid$(fullPath, closePos + 1)

    If Len(remainder) = 0 Then
        outColPath = "value"
    ElseIf Left$(remainder, 1) = "." Then
        outColPath = Mid$(remainder, 2)
        If Len(outColPath) = 0 Then outColPath = "value"
    Else
        Exit Function
    End If

    Json_TryParseIndexedPath = True
End Function

Private Sub Json_BuildRootSegs(ByVal tableRoot As String, ByRef rootSegs() As String, ByRef rootSegCount As Long)
    rootSegCount = 0

    tableRoot = Trim$(tableRoot)
    If Len(tableRoot) = 0 Then Exit Sub
    If Left$(tableRoot, 2) <> "$." Then Exit Sub

    Dim remainder As String
    remainder = Mid$(tableRoot, 3)

    Dim toks As Collection
    Set toks = Json_TokenizePath(remainder)
    If toks.count = 0 Then Exit Sub

    rootSegCount = toks.count
    ReDim rootSegs(1 To rootSegCount) As String

    Dim i As Long
    For i = 1 To rootSegCount
        rootSegs(i) = CStr(toks(i))
    Next i
End Sub

Private Function Json_TryParseTableRowPath_Compiled( _
    ByVal fullPath As String, _
    ByVal tableRoot As String, _
    ByRef rootSegs() As String, _
    ByVal rootSegCount As Long, _
    ByRef outIndex As Long, _
    ByRef outColPath As String, _
    ByRef outRowKey As String _
) As Boolean

    Json_TryParseTableRowPath_Compiled = False
    outIndex = 0
    outColPath = vbNullString
    outRowKey = vbNullString

    If rootSegCount = 0 Then Exit Function
    If Len(fullPath) = 0 Or Len(tableRoot) = 0 Then Exit Function
    If Left$(tableRoot, 2) <> "$." Then Exit Function
    If Left$(fullPath, 2) <> "$." Then Exit Function

    Dim pos As Long
    pos = 3 ' after "$." in fullPath

    Dim i As Long
    For i = 1 To rootSegCount
        Dim seg As String
        seg = rootSegs(i)

        If Mid$(fullPath, pos, Len(seg)) <> seg Then Exit Function
        pos = pos + Len(seg)

        If pos <= Len(fullPath) Then
            If Mid$(fullPath, pos, 1) = "[" Then
                Dim closePos As Long
                closePos = InStr(pos + 1, fullPath, "]")
                If closePos = 0 Then Exit Function

                Dim idxText As String
                idxText = Mid$(fullPath, pos + 1, closePos - pos - 1)
                If Len(idxText) = 0 Or Not Json_IsAllDigits(idxText) Then Exit Function

                If i = rootSegCount Then
                    outIndex = CLng(idxText)
                End If

                pos = closePos + 1
            Else
                If i = rootSegCount Then Exit Function
            End If
        Else
            Exit Function
        End If

        If i < rootSegCount Then
            If pos > Len(fullPath) Then Exit Function
            If Mid$(fullPath, pos, 1) <> "." Then Exit Function
            pos = pos + 1
        End If
    Next i

    outRowKey = Left$(fullPath, pos - 1)

    If pos > Len(fullPath) Then
        outColPath = "value"
        Json_TryParseTableRowPath_Compiled = True
        Exit Function
    End If

    Dim remainder As String
    remainder = Mid$(fullPath, pos)

    If Len(remainder) = 0 Then
        outColPath = "value"
    ElseIf Left$(remainder, 1) = "." Then
        outColPath = Mid$(remainder, 2)
        If Len(outColPath) = 0 Then outColPath = "value"
    Else
        Exit Function
    End If

    Json_TryParseTableRowPath_Compiled = True
End Function

' =============================================================================
' RowKeyMap internals (open addressing, stable row order)
' =============================================================================

Private Sub RowKeyMap_Init(ByRef m As RowKeyMap, Optional ByVal initialCap As Long = 64)
    Dim capPow2 As Long
    capPow2 = 1
    Do While capPow2 < initialCap
        capPow2 = capPow2 * 2
    Loop

    m.cap = capPow2
    ReDim m.slotHash(0 To m.cap - 1) As Long
    ReDim m.slotIdx(0 To m.cap - 1) As Long

    m.count = 0

    ReDim m.rowKeys(1 To 16) As String
    ReDim m.rowObjs(1 To 16) As Collection
End Sub

Private Sub RowKeyMap_Rehash(ByRef m As RowKeyMap, ByVal newCap As Long)
    Dim capPow2 As Long
    capPow2 = 1
    Do While capPow2 < newCap
        capPow2 = capPow2 * 2
    Loop
    newCap = capPow2

    Dim newHash() As Long
    Dim newIdx() As Long
    ReDim newHash(0 To newCap - 1) As Long
    ReDim newIdx(0 To newCap - 1) As Long

    Dim mask As Long
    mask = newCap - 1

    Dim i As Long
    For i = 1 To m.count
        Dim h As Long
        h = Json_Hash32_FNV1a(m.rowKeys(i))

        Dim pos As Long
        pos = (h And mask)

        Do
            If newIdx(pos) = 0 Then
                newHash(pos) = h
                newIdx(pos) = i
                Exit Do
            End If
            pos = (pos + 1) And mask
        Loop
    Next i

    m.cap = newCap
    m.slotHash = newHash
    m.slotIdx = newIdx
End Sub

Private Function RowKeyMap_GetOrAdd( _
    ByRef m As RowKeyMap, _
    ByVal rowKey As String, _
    ByVal rows As Collection _
) As Collection

    If m.cap = 0 Then RowKeyMap_Init m, 64

    If (m.count + 1) * 10 > m.cap * 7 Then
        RowKeyMap_Rehash m, (m.cap * 2)
    End If

    Dim h As Long
    h = Json_Hash32_FNV1a(rowKey)

    Dim mask As Long
    mask = m.cap - 1

    Dim pos As Long
    pos = (h And mask)

    Do
        If m.slotIdx(pos) = 0 Then
            m.count = m.count + 1

            If m.count > UBound(m.rowKeys) Then
                ReDim Preserve m.rowKeys(1 To UBound(m.rowKeys) * 2) As String
                ReDim Preserve m.rowObjs(1 To UBound(m.rowObjs) * 2) As Collection
            End If

            Dim o As New Collection
            o.Add TAG_OBJECT
            rows.Add o

            m.rowKeys(m.count) = rowKey
            Set m.rowObjs(m.count) = o

            m.slotHash(pos) = h
            m.slotIdx(pos) = m.count

            Set RowKeyMap_GetOrAdd = o
            Exit Function
        Else
            If m.slotHash(pos) = h Then
                Dim idx As Long
                idx = m.slotIdx(pos)
                If m.rowKeys(idx) = rowKey Then
                    Set RowKeyMap_GetOrAdd = m.rowObjs(idx)
                    Exit Function
                End If
            End If
            pos = (pos + 1) And mask
        End If
    Loop
End Function

Private Function Json_EnsureRow(ByVal rows As Collection, ByVal idx As Long) As Collection
    Dim needCount As Long
    needCount = idx + 1

    Do While rows.count < needCount
        Dim o As New Collection
        o.Add TAG_OBJECT
        rows.Add o
    Loop

    Set Json_EnsureRow = rows(needCount)
End Function

' =============================================================================
' Header hash table internals (no Dictionary)
' =============================================================================

Private Sub HeaderTable_Ensure( _
    ByVal key As String, _
    ByRef hdrs() As String, _
    ByRef hdrCount As Long, _
    ByRef slotHash() As Long, _
    ByRef slotIdx() As Long, _
    ByRef cap As Long, _
    ByVal DBG As Boolean _
)
    If (hdrCount + 1) * 10 > cap * 7 Then
        HeaderTable_Rehash hdrs, hdrCount, slotHash, slotIdx, cap, (cap * 2), DBG
    End If

    Dim h As Long
    h = Json_Hash32_FNV1a(key)

    Dim mask As Long
    mask = cap - 1

    Dim pos As Long
    pos = (h And mask)

    Do
        If slotIdx(pos) = 0 Then
            hdrCount = hdrCount + 1
            If hdrCount = 1 Then
                ReDim hdrs(1 To 16) As String
            ElseIf hdrCount > UBound(hdrs) Then
                ReDim Preserve hdrs(1 To UBound(hdrs) * 2) As String
            End If

            hdrs(hdrCount) = key
            slotHash(pos) = h
            slotIdx(pos) = hdrCount
            Exit Sub
        Else
            If slotHash(pos) = h Then
                Dim idx As Long
                idx = slotIdx(pos)
                If hdrs(idx) = key Then Exit Sub
            End If
            pos = (pos + 1) And mask
        End If
    Loop
End Sub

Private Function HeaderTable_Find( _
    ByVal key As String, _
    ByRef hdrs() As String, _
    ByRef slotHash() As Long, _
    ByRef slotIdx() As Long, _
    ByVal cap As Long _
) As Long
    Dim h As Long
    h = Json_Hash32_FNV1a(key)

    Dim mask As Long
    mask = cap - 1

    Dim pos As Long
    pos = (h And mask)

    Do
        If slotIdx(pos) = 0 Then
            HeaderTable_Find = 0
            Exit Function
        End If

        If slotHash(pos) = h Then
            Dim idx As Long
            idx = slotIdx(pos)
            If hdrs(idx) = key Then
                HeaderTable_Find = idx
                Exit Function
            End If
        End If

        pos = (pos + 1) And mask
    Loop
End Function

Private Sub HeaderTable_Rehash( _
    ByRef hdrs() As String, _
    ByVal hdrCount As Long, _
    ByRef slotHash() As Long, _
    ByRef slotIdx() As Long, _
    ByRef cap As Long, _
    ByVal newCap As Long, _
    ByVal DBG As Boolean _
)
    Dim pow2 As Long
    pow2 = 1
    Do While pow2 < newCap
        pow2 = pow2 * 2
    Loop
    newCap = pow2

    If DBG Then Debug.Print "Rehash: cap " & cap & " -> " & newCap & " (hdrCount=" & hdrCount & ")"

    Dim newHash() As Long
    Dim newIdx() As Long
    ReDim newHash(0 To newCap - 1) As Long
    ReDim newIdx(0 To newCap - 1) As Long

    Dim mask As Long
    mask = newCap - 1

    Dim i As Long
    For i = 1 To hdrCount
        Dim key As String
        key = hdrs(i)

        Dim h As Long
        h = Json_Hash32_FNV1a(key)

        Dim pos As Long
        pos = (h And mask)

        Do
            If newIdx(pos) = 0 Then
                newHash(pos) = h
                newIdx(pos) = i
                Exit Do
            End If
            pos = (pos + 1) And mask
        Loop
    Next i

    cap = newCap
    slotHash = newHash
    slotIdx = newIdx
End Sub

' =============================================================================
' Hash: FNV-1a 32-bit (safe in VBA Long via LongLong)
' =============================================================================

Private Function Json_Hash32_FNV1a(ByVal s As String) As Long
    Const FNV_OFFSET As Long = &H811C9DC5
    Const FNV_PRIME  As Long = &H1000193

    Dim MASK32 As LongLong
    MASK32 = (CLngLng(&H7FFFFFFF) * 2) + 1          ' 4294967295

    Dim TWO32 As LongLong
    TWO32 = (CLngLng(&H7FFFFFFF) + 1) * 2           ' 4294967296

    Dim h As Long
    h = FNV_OFFSET

    Dim i As Long
    For i = 1 To Len(s)
        Dim cc As Long
        cc = AscW(Mid$(s, i, 1)) And &HFFFF&

        Dim t As LongLong
        t = (CLngLng(h) Xor CLngLng(cc)) * CLngLng(FNV_PRIME)

        Dim u As LongLong
        u = (t And MASK32)

        If u > 2147483647# Then
            h = CLng(u - TWO32)
        Else
            h = CLng(u)
        End If
    Next i

    Json_Hash32_FNV1a = h
End Function

' =============================================================================
' Excel helpers
' =============================================================================

Private Function Excel_IsDefaultValueOnlyHeaders(ByVal headers As Variant) As Boolean
    On Error GoTo Nope

    Dim lb As Long, ub As Long
    lb = LBound(headers)
    ub = UBound(headers)

    If (ub - lb + 1) <> 1 Then GoTo Nope

    Dim h As String
    h = LCase$(Trim$(CStr(headers(lb))))

    Excel_IsDefaultValueOnlyHeaders = (h = "value")
    Exit Function

Nope:
    Excel_IsDefaultValueOnlyHeaders = False
End Function

Private Function Excel_RowCount2D(ByVal data2D As Variant) As Long
    If IsEmpty(data2D) Then
        Excel_RowCount2D = 0
    Else
        Excel_RowCount2D = (UBound(data2D, 1) - LBound(data2D, 1) + 1)
    End If
End Function

Private Function Excel_ColCount2D(ByVal data2D As Variant) As Long
    If IsEmpty(data2D) Then
        Excel_ColCount2D = 0
    Else
        Excel_ColCount2D = (UBound(data2D, 2) - LBound(data2D, 2) + 1)
    End If
End Function

Private Function Excel_ListObjectHeadersTo1D(ByVal lo As ListObject) As Variant
    Dim n As Long
    n = lo.ListColumns.count

    Dim arr As Variant
    ReDim arr(1 To n)

    Dim i As Long
    For i = 1 To n
        arr(i) = lo.ListColumns(i).name
    Next i

    Excel_ListObjectHeadersTo1D = arr
End Function

Private Function Excel_UnionHeadersFromListObject(ByVal lo As ListObject, ByVal incomingHeaders As Variant) As Variant
    Dim existing As Variant
    existing = Excel_ListObjectHeadersTo1D(lo)

    Dim outList As New Collection

    Dim i As Long
    For i = 1 To UBound(existing)
        Dim ex As String
        ex = Trim$(CStr(existing(i)))
        outList.Add ex
    Next i

    Dim lb As Long, ub As Long
    lb = LBound(incomingHeaders)
    ub = UBound(incomingHeaders)

    For i = lb To ub
        Dim h As String
        h = Trim$(CStr(incomingHeaders(i)))
        If Not Excel_CollectionContainsText(outList, h) Then outList.Add h
    Next i

    Excel_UnionHeadersFromListObject = Excel_CollectionTo1D(outList)
End Function

Private Function Excel_CollectionContainsText(ByVal c As Collection, ByVal s As String) As Boolean
    Dim needle As String
    needle = Trim$(CStr(s))

    Dim i As Long
    For i = 1 To c.count
        If StrComp(Trim$(CStr(c(i))), needle, vbTextCompare) = 0 Then
            Excel_CollectionContainsText = True
            Exit Function
        End If
    Next i
End Function

Private Function Excel_CollectionTo1D(ByVal c As Collection) As Variant
    Dim arr As Variant
    ReDim arr(1 To c.count)

    Dim i As Long
    For i = 1 To c.count
        arr(i) = CStr(c(i))
    Next i

    Excel_CollectionTo1D = arr
End Function

Private Function Excel_ReshapeDataToHeaders( _
    ByVal inHeaders As Variant, _
    ByVal outHeaders As Variant, _
    ByVal inData As Variant _
) As Variant
    If IsEmpty(inData) Then
        Excel_ReshapeDataToHeaders = Empty
        Exit Function
    End If

    Dim inRows As Long
    Dim inCols As Long
    Dim outCols As Long

    inRows = Excel_RowCount2D(inData)
    inCols = Excel_ColCount2D(inData)
    outCols = (UBound(outHeaders) - LBound(outHeaders) + 1)

    Dim outArr As Variant
    ReDim outArr(1 To inRows, 1 To outCols)

    Dim cap As Long
    cap = 1
    Do While cap < (inCols * 2)
        cap = cap * 2
    Loop
    If cap < 16 Then cap = 16

    Dim slotHash() As Long
    Dim slotIdx() As Long
    ReDim slotHash(0 To cap - 1) As Long
    ReDim slotIdx(0 To cap - 1) As Long

    Dim i As Long
    For i = LBound(inHeaders) To UBound(inHeaders)
        Dim key As String
        key = Trim$(CStr(inHeaders(i)))

        Dim h As Long
        h = Json_Hash32_FNV1a(key)

        Dim pos As Long
        Dim mask As Long
        mask = cap - 1
        pos = (h And mask)

        Do
            If slotIdx(pos) = 0 Then
                slotHash(pos) = h
                slotIdx(pos) = (i - LBound(inHeaders) + 1)
                Exit Do
            End If

            If slotHash(pos) = h Then
                If StrComp(Trim$(CStr(inHeaders(LBound(inHeaders) + slotIdx(pos) - 1))), key, vbTextCompare) = 0 Then
                    Exit Do
                End If
            End If

            pos = (pos + 1) And mask
        Loop
    Next i

    Dim oc As Long
    For oc = 1 To outCols
        Dim outKey As String
        outKey = Trim$(CStr(outHeaders(LBound(outHeaders) + oc - 1)))

        Dim outHash As Long
        outHash = Json_Hash32_FNV1a(outKey)

        Dim foundIdx As Long
        foundIdx = 0

        mask = cap - 1
        pos = (outHash And mask)

        Do
            If slotIdx(pos) = 0 Then Exit Do

            If slotHash(pos) = outHash Then
                If StrComp(Trim$(CStr(inHeaders(LBound(inHeaders) + slotIdx(pos) - 1))), outKey, vbTextCompare) = 0 Then
                    foundIdx = slotIdx(pos)
                    Exit Do
                End If
            End If

            pos = (pos + 1) And mask
        Loop

        If foundIdx > 0 And foundIdx <= inCols Then
            Dim srcCol As Long
            srcCol = LBound(inData, 2) + foundIdx - 1

            Dim r As Long
            For r = 1 To inRows
                outArr(r, oc) = inData(LBound(inData, 1) + r - 1, srcCol)
            Next r
        End If
    Next oc

    Excel_ReshapeDataToHeaders = outArr
End Function

Private Function Excel_FindHeaderIndex(ByVal headers As Variant, ByVal headerName As String) As Long
    Dim needle As String
    needle = Trim$(CStr(headerName))

    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If StrComp(Trim$(CStr(headers(i))), needle, vbTextCompare) = 0 Then
            Excel_FindHeaderIndex = (i - LBound(headers) + 1)
            Exit Function
        End If
    Next i
End Function

Private Function Excel_HeadersTo2D(ByVal headers As Variant) As Variant
    Dim lb As Long, ub As Long
    lb = LBound(headers)
    ub = UBound(headers)

    Dim outArr As Variant
    ReDim outArr(1 To 1, 1 To (ub - lb + 1))

    Dim c As Long
    c = 1

    Dim i As Long
    For i = lb To ub
        outArr(1, c) = CStr(headers(i))
        c = c + 1
    Next i

    Excel_HeadersTo2D = outArr
End Function

Private Sub Excel_ClearOrphanedColumns( _
    ByVal lo As ListObject, _
    ByVal newColCount As Long, _
    ByVal oldColCount As Long, _
    ByVal oldBodyRows As Long _
)
    Dim tl As Range
    Set tl = lo.Range.Cells(1, 1)

    Dim orphanHeader As Range
    Set orphanHeader = tl.Offset(0, newColCount).Resize(1, oldColCount - newColCount)
    orphanHeader.ClearContents

    If oldBodyRows > 0 Then
        Dim orphanBody As Range
        Set orphanBody = tl.Offset(1, newColCount).Resize(oldBodyRows, oldColCount - newColCount)
        orphanBody.ClearContents
    End If
End Sub

Private Sub Excel_ClearOrphanedHeaderOnly( _
    ByVal lo As ListObject, _
    ByVal newColCount As Long, _
    ByVal oldColCount As Long _
)
    If newColCount >= oldColCount Then Exit Sub

    Dim tl As Range
    Set tl = lo.Range.Cells(1, 1)

    tl.Offset(0, newColCount).Resize(1, oldColCount - newColCount).ClearContents
End Sub

Private Sub Excel_ValidateHeaders(ByRef headers As Variant, ByVal sourceName As String)
    Dim i As Long
    Dim j As Long

    For i = LBound(headers) To UBound(headers)
        Dim hi As String
        hi = Trim$(CStr(headers(i)))

        If Len(hi) = 0 Then
            Err.Raise vbObjectError + 1120, sourceName, "Header at index " & i & " is blank."
        End If

        headers(i) = hi
    Next i

    For i = LBound(headers) To UBound(headers)
        For j = i + 1 To UBound(headers)
            If StrComp(CStr(headers(i)), CStr(headers(j)), vbTextCompare) = 0 Then
                Err.Raise vbObjectError + 1121, sourceName, _
                    "Duplicate header (case-insensitive): '" & CStr(headers(i)) & "' at indices " & i & " and " & j & "."
            End If
        Next j
    Next i
End Sub

' =============================================================================
' JSONPath resolve (minimal, deterministic)
' =============================================================================

Public Function Json_TryResolvePath( _
    ByVal root As Variant, _
    ByVal path As String, _
    ByRef outValue As Variant _
) As Boolean

    Json_TryResolvePath = False
    VarAssign outValue, Null

    path = Trim$(path)
    If Len(path) = 0 Then Exit Function
    If path = "$" Then
        VarAssign outValue, root
        Json_TryResolvePath = True
        Exit Function
    End If

    If Left$(path, 2) <> "$." Then Exit Function
    If Not IsObject(root) Then Exit Function
    If TypeName(root) <> "Collection" Then Exit Function

    Dim cur As Variant
    VarAssign cur, root

    Dim i As Long
    i = 3 ' after "$."

    Do While i <= Len(path)
        Dim seg As String
        seg = vbNullString

        Do While i <= Len(path)
            Dim ch As String
            ch = Mid$(path, i, 1)
            If ch = "." Or ch = "[" Then Exit Do
            seg = seg & ch
            i = i + 1
        Loop

        If Len(seg) > 0 Then
            If Not IsObject(cur) Then Exit Function
            If TypeName(cur) <> "Collection" Then Exit Function
            If Not Json_IsObject(cur) Then Exit Function

            Dim nextVal As Variant
            If Not Json_TryObjGet(cur, seg, nextVal) Then Exit Function
            VarAssign cur, nextVal
        End If

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
            VarAssign elem, arr(oneBased)
            VarAssign cur, elem
        Loop

        If i <= Len(path) Then
            If Mid$(path, i, 1) = "." Then
                i = i + 1
            ElseIf Mid$(path, i, 1) <> "[" Then
                Exit Function
            End If
        End If
    Loop

    VarAssign outValue, cur
    Json_TryResolvePath = True
End Function

Public Function Json_ObjGet(ByVal obj As Collection, ByVal key As String) As Variant
    Const ERR_SRC As String = "zz_ModernJsonInVba.Json_ObjGet"
    
    Dim v As Variant
    
    If Json_TryObjGet(obj, key, v) Then
        VarAssign Json_ObjGet, v
        Exit Function
    End If
    
    Err.Raise vbObjectError + 5320, ERR_SRC, _
        "JSON object key not found: '" & key & "'"
End Function

Public Function Json_TryObjGet(ByVal obj As Collection, ByVal key As String, ByRef outValue As Variant) As Boolean
    Json_TryObjGet = False
    VarAssign outValue, Null

    Dim i As Long
    For i = 2 To obj.count
        Dim pair As Variant
        pair = obj(i)
        If StrComp(CStr(pair(0)), key, vbBinaryCompare) = 0 Then
            VarAssign outValue, pair(1)
            Json_TryObjGet = True
            Exit Function
        End If
    Next i
End Function

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

' =============================================================================
' Excel_ListObjectToJson
' =============================================================================
' Converts an Excel ListObject (table) into a JSON array-of-objects.
'
' Each row becomes a JSON object. Column headers define JSON property paths.
'
' Value precedence:
'
'   1. JSON structure parsed from cell text (if parseJsonInCells=True)
'   2. Formula text (if preserveFormulas=True and JSON parsing did not occur)
'   3. Raw Value2 cell value
'
' Nested paths using dot notation (customer.name) are supported.
' Array index paths (items[0].sku) are NOT supported.
'
' -----------------------------------------------------------------------------
' PARAMETERS
' -----------------------------------------------------------------------------
' lo
'   Source ListObject
'
' includeBlanksAsNull (default False)
'   Blank cells omitted or written as JSON null
'
' parseJsonInCells (default False)
'   Detect and parse JSON arrays/objects inside cell text
'
' parseArraysOnly (default False)
'   Only parse arrays if True
'
' preserveFormulas (default False)
'   Serialize formula text instead of evaluated value
'
' -----------------------------------------------------------------------------
' ERROR CONDITIONS
' -----------------------------------------------------------------------------
' 1120 blank header
' 1121 duplicate header
' 1170 Excel error value
' 1171 column mismatch
' 1172 TAG_OBJECT missing
' 905  array index path unsupported
' =============================================================================
Public Function Excel_ListObjectToJson( _
    ByVal lo As ListObject, _
    Optional ByVal includeBlanksAsNull As Boolean = False, _
    Optional ByVal parseJsonInCells As Boolean = False, _
    Optional ByVal parseArraysOnly As Boolean = False, _
    Optional ByVal preserveFormulas As Boolean = False _
) As String

    Const SRC As String = "Excel_ListObjectToJson"

    If Len(TAG_OBJECT) = 0 Then
        Err.Raise vbObjectError + 1172, SRC, "TAG_OBJECT is blank or not initialized."
    End If

    ' -----------------------------
    ' Headers
    ' -----------------------------
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

    Dim i As Long, j As Long
    For i = 1 To colCount
        For j = i + 1 To colCount
            If StrComp(headers(i), headers(j), vbTextCompare) = 0 Then
                Err.Raise vbObjectError + 1121, SRC, _
                    "Duplicate header: '" & headers(i) & _
                    "' at indices " & i & " and " & j
            End If
        Next j
    Next i

    If lo.DataBodyRange Is Nothing Then
        Excel_ListObjectToJson = "[]"
        Exit Function
    End If

    ' -----------------------------
    ' Bulk read
    ' -----------------------------
    Dim data As Variant
    data = lo.DataBodyRange.Value2

    Dim formulas As Variant
    If preserveFormulas Then
        formulas = lo.DataBodyRange.Formula
    End If

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

    ' -----------------------------
    ' Build JSON
    ' -----------------------------
    Dim arr As Collection
    Set arr = New Collection

    Dim r As Long
    For r = 1 To rowCount

        Dim rowObj As Collection
        Set rowObj = New Collection
        rowObj.Add TAG_OBJECT

        For c = 1 To colCount

            Dim keyPath As String
            keyPath = headers(c)

            If InStr(keyPath, "[") > 0 Or InStr(keyPath, "]") > 0 Then
                Err.Raise vbObjectError + 905, SRC, _
                    "Array index paths unsupported: " & keyPath
            End If

            Dim v As Variant
            v = data(LBound(data, 1) + r - 1, LBound(data, 2) + c - 1)

            If IsError(v) Then
                Err.Raise vbObjectError + 1170, SRC, _
                    "Excel error value at row " & r & ", col " & c
            End If

            Dim parsedJson As Boolean
            parsedJson = False

            ' ---------------------------------
            ' JSON parsing (highest priority)
            ' ---------------------------------
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

                            If Excel_ListObjectToJson_TryParseJsonCell(s, parsedCell) Then

                                If IsObject(parsedCell) Then

                                    If TypeName(parsedCell) = "Collection" Then

                                        If Json_IsObject(parsedCell) Or Json_IsArray(parsedCell) Then
                                            VarAssign v, parsedCell
                                            parsedJson = True
                                        End If

                                    End If

                                End If

                            End If

                        End If

                    End If

                End If

            End If

            ' ---------------------------------
            ' Formula preservation (secondary)
            ' ---------------------------------
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

            Dim isBlank As Boolean
            
            If VarType(v) = vbString Then
                isBlank = (LenB(v) = 0)
            Else
                isBlank = IsEmpty(v)
            End If

            If isBlank Then

                If includeBlanksAsNull Then
                    Excel_ListObjectToJson_InsertValue rowObj, keyPath, Null
                End If

            Else

                Excel_ListObjectToJson_InsertValue rowObj, keyPath, v

            End If

        Next c

        arr.Add rowObj

    Next r

    Excel_ListObjectToJson = Json_Stringify(arr)

End Function

Private Sub Excel_ListObjectToJson_InsertValue( _
    ByVal rowObj As Collection, _
    ByVal keyPath As String, _
    ByVal value As Variant _
)
    Dim toks As Collection
    Set toks = Json_TokenizePath(keyPath)

    If toks.count > 1 Then
        Json_UnflattenInsert rowObj, keyPath, value
    Else
        Json_ObjSet rowObj, Json_UnescapePathSegment(CStr(toks(1))), value
    End If
End Sub

Private Function Excel_ListObjectToJson_TryParseJsonCell( _
    ByVal s As String, _
    ByRef outValue As Variant _
) As Boolean
    ' Parse cell text as JSON, but never throw from here.
    ' Return True only when parse succeeded.
    Excel_ListObjectToJson_TryParseJsonCell = False
    VarAssign outValue, Null

    On Error GoTo Fail

    ' Must use the engine parser so you get the same deterministic model.
    Dim v As Variant
    Json_ParseInto s, v

    VarAssign outValue, v
    Excel_ListObjectToJson_TryParseJsonCell = True
    Exit Function

Fail:
    ' swallow parse errors: treat as ordinary string cell
    Err.Clear
End Function

' =============================================================================
' Formula preservation for ListObject refresh/append
'
' Behavior:
'   - Captures existing formula templates per header before any resize/clear.
'   - After writing data:
'       * clearExisting=True  => reapply formulas down entire body for those columns
'       * clearExisting=False => apply formulas only to newly appended rows
'
' Notes:
'   - "Formula column" is detected if ANY cell in that column has a formula.
'     Template = first formula found scanning top-down.
'   - Incoming data for a formula column is overwritten by the formula.
'   - Deterministic: header match is case-insensitive; first-found formula wins.
' =============================================================================

Private Sub Excel_CaptureFormulaTemplates( _
    ByVal lo As ListObject, _
    ByRef outHdrs() As String, _
    ByRef outFmlR1C1() As String, _
    ByRef outCount As Long _
)
    outCount = 0
    Erase outHdrs
    Erase outFmlR1C1

    If lo Is Nothing Then Exit Sub
    If lo.ListColumns.count = 0 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim c As Long
    For c = 1 To lo.ListColumns.count
        Dim colRng As Range
        Set colRng = lo.DataBodyRange.Columns(c)

        Dim f As String
        If Excel_TryFindFirstFormulaR1C1(colRng, f) Then
            outCount = outCount + 1
            If outCount = 1 Then
                ReDim outHdrs(1 To 8) As String
                ReDim outFmlR1C1(1 To 8) As String
            ElseIf outCount > UBound(outHdrs) Then
                ReDim Preserve outHdrs(1 To UBound(outHdrs) * 2) As String
                ReDim Preserve outFmlR1C1(1 To UBound(outFmlR1C1) * 2) As String
            End If

            outHdrs(outCount) = CStr(lo.ListColumns(c).name)
            outFmlR1C1(outCount) = f
        End If
    Next c
End Sub

Private Function Excel_TryFindFirstFormulaR1C1(ByVal colRng As Range, ByRef outFormulaR1C1 As String) As Boolean
    Excel_TryFindFirstFormulaR1C1 = False
    outFormulaR1C1 = vbNullString

    If colRng Is Nothing Then Exit Function

    Dim r As Long
    For r = 1 To colRng.rows.count
        Dim cell As Range
        Set cell = colRng.Cells(r, 1)

        If cell.HasFormula Then
            outFormulaR1C1 = cell.FormulaR1C1
            Excel_TryFindFirstFormulaR1C1 = (Len(outFormulaR1C1) > 0)
            Exit Function
        End If
    Next r
End Function

Private Function Excel_TryGetFormulaForHeader( _
    ByRef fHdrs() As String, _
    ByRef fFmls() As String, _
    ByVal fCount As Long, _
    ByVal headerName As String, _
    ByRef outFormulaR1C1 As String _
) As Boolean
    Excel_TryGetFormulaForHeader = False
    outFormulaR1C1 = vbNullString

    If fCount <= 0 Then Exit Function

    Dim i As Long
    For i = 1 To fCount
        If StrComp(fHdrs(i), headerName, vbTextCompare) = 0 Then
            outFormulaR1C1 = fFmls(i)
            Excel_TryGetFormulaForHeader = (Len(outFormulaR1C1) > 0)
            Exit Function
        End If
    Next i
End Function

Private Sub Excel_ApplyFormulasToBody( _
    ByVal lo As ListObject, _
    ByRef finalHeaders As Variant, _
    ByVal bodyRowCount As Long, _
    ByRef fHdrs() As String, _
    ByRef fFmls() As String, _
    ByVal fCount As Long _
)
    If lo Is Nothing Then Exit Sub
    If bodyRowCount <= 0 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim newCols As Long
    newCols = (UBound(finalHeaders) - LBound(finalHeaders) + 1)

    Dim c As Long
    For c = 1 To newCols
        Dim h As String
        h = CStr(finalHeaders(LBound(finalHeaders) + c - 1))

        Dim f As String
        If Excel_TryGetFormulaForHeader(fHdrs, fFmls, fCount, h, f) Then
            lo.DataBodyRange.Columns(c).FormulaR1C1 = f
        End If
    Next c
End Sub

Private Sub Excel_ApplyFormulasToAppendedRows( _
    ByVal lo As ListObject, _
    ByRef finalHeaders As Variant, _
    ByVal startRowZeroBased As Long, _
    ByVal appendedRowCount As Long, _
    ByRef fHdrs() As String, _
    ByRef fFmls() As String, _
    ByVal fCount As Long _
)
    If lo Is Nothing Then Exit Sub
    If appendedRowCount <= 0 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim newCols As Long
    newCols = (UBound(finalHeaders) - LBound(finalHeaders) + 1)

    Dim c As Long
    For c = 1 To newCols
        Dim h As String
        h = CStr(finalHeaders(LBound(finalHeaders) + c - 1))

        Dim f As String
        If Excel_TryGetFormulaForHeader(fHdrs, fFmls, fCount, h, f) Then
            lo.DataBodyRange.Cells(startRowZeroBased + 1, c).Resize(appendedRowCount, 1).FormulaR1C1 = f
        End If
    Next c
End Sub

' =============================================================================
' Excel_UpsertListObjectFromSource
'
' Purpose
'   Unified ingestion entry point for loading structured text formats into an
'   Excel ListObject using the deterministic ModernJsonInVBA upsert pipeline.
'
'   The function accepts structured source text (JSON, CSV, XML), converts it
'   into the canonical JSON representation expected by the core engine, and
'   then delegates the actual table ingestion to:
'
'       Excel_UpsertListObjectFromJsonAtRoot
'
'   This preserves all deterministic behaviors of the core engine including:
'       - stable header discovery
'       - schema evolution controls
'       - formula column preservation
'       - deterministic ListObject updates
'
'
' Architecture
'
'       Source Text
'           ¦
'           ?
'     Format Adapter
'   (JSON / CSV / XML)
'           ¦
'           ?
'        JSON Text
'           ¦
'           ?
'   Excel_UpsertListObjectFromJsonAtRoot
'           ¦
'           ?
'     Flatten + Table Extraction
'           ¦
'           ?
'      2D Array Conversion
'           ¦
'           ?
'   Deterministic ListObject Upsert
'
'
' Supported Formats
'
'   ExcelSourceFormat_JSON
'       Input is already JSON text. The caller supplies the tableRoot path.
'
'   ExcelSourceFormat_CSV
'       CSV text is converted to a JSON array-of-objects via CsvTextToJson().
'       The resulting table root is "$".
'
'   ExcelSourceFormat_XML
'       XML text is converted to JSON via XmlTextToJson(). The resulting JSON
'       structure is expected to contain an array under the property "item".
'       The resulting table root is "$.item".
'
'
' Parameters
'
'   ws
'       Target worksheet containing the ListObject.
'
'   tableName
'       Name of the ListObject to create or update.
'
'   topLeft
'       Anchor cell used when creating the table if it does not exist.
'
'   sourceText
'       Raw input text representing the structured data source.
'
'   format
'       ExcelSourceFormat enum specifying the input format.
'
'   tableRoot
'       JSON path identifying the array-of-objects that forms the table.
'       Used only when format = ExcelSourceFormat_JSON.
'
'   clearExisting
'       If True, existing rows are cleared before inserting new data.
'
'   addMissingColumns
'       If True, new columns discovered in the input schema are added.
'
'   removeMissingColumns
'       If True, columns not present in the input schema are removed.
'
'   preserveFormulaColumns
'       If True, columns containing formulas are preserved during upsert.
'
'   fillFormulasOnAppend
'       If True, formula columns automatically propagate when rows expand.
'
'   nonTableArraysAsJson
'       If True, nested arrays not part of the table root are serialized as
'       JSON text in their respective cells instead of being excluded.
'
'
' Determinism Guarantees
'
'   - Header discovery preserves first-seen order.
'   - Stable column ordering across runs.
'   - Schema evolution controlled via explicit flags.
'   - No Scripting.Dictionary dependencies.
'   - Consistent ListObject structure regardless of source format.
'
'
' Notes
'
'   This function acts strictly as a routing and format-adaptation layer.
'   All parsing, flattening, schema discovery, and ListObject update logic
'   remains centralized inside the ModernJsonInVBA ingestion pipeline.
'
' =============================================================================
Public Sub Excel_UpsertListObjectFromSource( _
    ByVal ws As Worksheet, _
    ByVal tableName As String, _
    ByVal topLeft As Range, _
    ByVal sourceText As String, _
    ByVal format As ExcelSourceFormat, _
    Optional ByVal tableRoot As String = "$", _
    Optional ByVal clearExisting As Boolean = True, _
    Optional ByVal addMissingColumns As Boolean = True, _
    Optional ByVal removeMissingColumns As Boolean = False, _
    Optional ByVal preserveFormulaColumns As Boolean = True, _
    Optional ByVal fillFormulasOnAppend As Boolean = True, _
    Optional ByVal nonTableArraysAsJson As Boolean = False _
)

    Const ERR_SRC As String = "Excel_UpsertListObjectFromSource"
    Const ERR_BADFORMAT As Long = vbObjectError + 1400

    Dim jsonText As String
    Dim resolvedRoot As String

    Select Case format

        Case ExcelSourceFormat_JSON
            
            jsonText = sourceText
            resolvedRoot = tableRoot
    
        Case ExcelSourceFormat_CSV
            
            jsonText = CsvTextToJson(sourceText)
            resolvedRoot = "$"
    
        Case ExcelSourceFormat_XML

            jsonText = XmlTextToJson(sourceText)
            resolvedRoot = tableRoot
    
        Case Else
            
            Err.Raise ERR_BADFORMAT, ERR_SRC, "Unsupported source format."
    
    End Select


    Excel_UpsertListObjectFromJsonAtRoot _
        ws, _
        tableName, _
        topLeft, _
        jsonText, _
        resolvedRoot, _
        clearExisting, _
        addMissingColumns, _
        removeMissingColumns, _
        preserveFormulaColumns, _
        fillFormulasOnAppend, _
        nonTableArraysAsJson

End Sub

' =============================================================================
' CsvFileToJson
'
' Reads a CSV file and converts it to JSON by delegating to CsvTextToJson.
' =============================================================================
Public Function CsvFileToJson(ByVal filePath As String) As String

    Dim f As Integer
    f = FreeFile
    
    Dim txt As String
    Open filePath For Input As #f
    txt = Input$(LOF(f), f)
    Close #f
    
    CsvFileToJson = CsvTextToJson(txt)

End Function

' =============================================================================
' CsvTextToJson
'
' Converts raw CSV text into a JSON array-of-objects.
'
' This implements an RFC-4180 style state-machine CSV parser.
' =============================================================================
Public Function CsvTextToJson(ByVal txt As String) As String

    Dim rows As New Collection
    Dim fields As New Collection
    
    Dim field As String
    Dim inQuotes As Boolean
    
    Dim i As Long
    Dim ch As String
    Dim nextCh As String
    Dim L As Long
    
    L = Len(txt)

    For i = 1 To L
    
        ch = Mid$(txt, i, 1)
        
        If inQuotes Then
        
            If ch = """" Then
            
                If i < L Then
                    nextCh = Mid$(txt, i + 1, 1)
                    
                    If nextCh = """" Then
                        field = field & """"
                        i = i + 1
                    Else
                        inQuotes = False
                    End If
                    
                Else
                    inQuotes = False
                End If
                
            Else
                field = field & ch
            End If
            
        Else
        
            Select Case ch
            
                Case """"
                    inQuotes = True
                    
                Case ","
                    fields.Add field
                    field = ""
                    
                Case vbCr
                    ' ignore
                    
                Case vbLf
                    fields.Add field
                    field = ""
                    
                    Dim row() As String
                    ReDim row(0 To fields.count - 1)
                    
                    Dim j As Long
                    For j = 1 To fields.count
                        row(j - 1) = fields(j)
                    Next j
                    
                    rows.Add row
                    Set fields = New Collection
                    
                Case Else
                    field = field & ch
                    
            End Select
            
        End If
        
    Next i

    If field <> "" Or fields.count > 0 Then
    
        fields.Add field
        
        Dim row2() As String
        ReDim row2(0 To fields.count - 1)
        
        Dim k As Long
        For k = 1 To fields.count
            row2(k - 1) = fields(k)
        Next k
        
        rows.Add row2
        
    End If

    If rows.count = 0 Then
        CsvTextToJson = "[]"
        Exit Function
    End If

    Dim headers() As String
    headers = rows(1)

    Dim r As Long, c As Long
    Dim json As String
    json = "["

    For r = 2 To rows.count
    
        Dim cols() As String
        cols = rows(r)
        
        json = json & "{"
        
        For c = LBound(headers) To UBound(headers)
        
            Dim key As String
            key = headers(c)
            
            Dim val As String
            If c <= UBound(cols) Then
                val = cols(c)
            Else
                val = ""
            End If
            
            key = Replace(key, "\", "\\")
            key = Replace(key, """", "\""")
            
            val = Replace(val, "\", "\\")
            val = Replace(val, """", "\""")
            val = Replace(val, vbCrLf, "\n")
            val = Replace(val, vbLf, "\n")
            val = Replace(val, vbCr, "\n")
            
            json = json & """" & key & """:""" & val & """"
            
            If c < UBound(headers) Then json = json & ","
            
        Next c
        
        json = json & "},"
        
    Next r

    If Right$(json, 1) = "," Then
        json = Left$(json, Len(json) - 1)
    End If
    
    json = json & "]"
    
    CsvTextToJson = json

End Function


' =============================================================================
' Function:   XmlFileToJson
'
' Author:     William Smith
' Created:    2026
'
' Summary
'   Reads an XML file from disk and converts it into JSON by delegating
'   parsing to XmlTextToJson.
'
' Behavior
'   • Reads the entire XML file into memory.
'   • Passes the raw text to XmlTextToJson for parsing and conversion.
'
' Parameters
'   filePath (String)
'       Full path to the XML file to read.
'
' Returns
'   String
'       JSON representation of the XML document.
'
' Notes
'   This function is a thin I/O wrapper. All parsing logic resides in
'   XmlTextToJson so that XML content can also be converted directly from
'   raw text (for example from HTTP responses).
' =============================================================================
Public Function XmlFileToJson(ByVal filePath As String) As String

    Dim f As Integer
    f = FreeFile
    
    Dim txt As String
    Open filePath For Input As #f
    txt = Input$(LOF(f), f)
    Close #f
    
    XmlFileToJson = XmlTextToJson(txt)

End Function

' =============================================================================
' Function:   XmlTextToJson
'
' Author:     William Smith
' Created:    2026
'
' Summary
'   Converts raw XML text into a JSON representation using a lightweight
'   pure-VBA XML parser.
'
' Behavior
'   • Parses XML elements using a character-level recursive descent parser.
'   • Produces JSON objects representing nested XML elements.
'   • Text nodes are serialized as JSON string values.
'
' XML Support
'   Supported:
'       • Nested elements
'       • Text nodes
'       • Self-closing tags
'
'   Not supported:
'       • XML namespaces
'       • DTDs
'       • Processing instructions
'       • Comments
'
' Parameters
'   txt (String)
'       Raw XML document text.
'
' Returns
'   String
'       JSON representation of the XML structure.
'
' Example
'
'   XML
'       <person>
'           <id>1</id>
'           <name>Alice</name>
'       </person>
'
'   Result
'       {"id":{"value":"1"},"name":{"value":"Alice"}}
'
' Notes
'   Designed for portability across Excel environments without requiring
'   platform-specific libraries such as MSXML.
' =============================================================================
Public Function XmlTextToJson(ByVal txt As String) As String

    Dim pos As Long
    pos = 1
    
    ' -------------------------------------------------
    ' Strip UTF-8 / UTF-16 BOM if present
    ' -------------------------------------------------
    If Len(txt) > 0 Then
        If AscW(Left$(txt, 1)) = &HFEFF Then
            txt = Mid$(txt, 2)
        End If
    End If
    
    Xml_SkipWhitespace txt, pos
    
    ' Parse root XML node directly into JSON
    XmlTextToJson = Xml_ParseNode(txt, pos)

End Function

Private Function Xml_ParseNode(ByVal txt As String, ByRef pos As Long) As String

    Dim name As String
    Dim L As Long
    L = Len(txt)

    pos = pos + 1
    name = Xml_ReadName(txt, pos)

    Xml_SkipAttributes txt, pos

    If Mid$(txt, pos, 2) = "/>" Then
        pos = pos + 2
        Xml_ParseNode = "null"
        Exit Function
    End If

    If Mid$(txt, pos, 1) = ">" Then
        pos = pos + 1
    Else
        Err.Raise vbObjectError + 824, "XmlTextToJson", "Malformed XML: expected '>'."
    End If


    Dim childNames As New Collection
    Dim childValues As New Collection
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
                Err.Raise vbObjectError + 822, "XmlTextToJson", "Unterminated CDATA section."
            End If

            textBuffer = textBuffer & Mid$(txt, pos + 9, cEnd - (pos + 9))
            pos = cEnd + 3

            GoTo ContinueLoop

        End If

        If Mid$(txt, pos, 1) = "<" Then

            Dim childName As String
            Dim tmp As Long
            tmp = pos + 1

            childName = Xml_ReadName(txt, tmp)

            Dim childJson As String
            childJson = Xml_ParseNode(txt, pos)

            childNames.Add childName
            childValues.Add childJson

        Else

            Dim textVal As String
            textVal = Xml_ReadText(txt, pos)

            If Len(Trim$(textVal)) > 0 Then
                textBuffer = textBuffer & textVal
            End If

        End If

ContinueLoop:

        If pos = startPos Then
            Err.Raise vbObjectError + 821, "XmlTextToJson", "Malformed XML: parser stalled."
        End If

    Loop


    ' -------------------------------------------------
    ' TEXT ONLY NODE  ? primitive
    ' -------------------------------------------------
    If childNames.count = 0 Then

        If Len(textBuffer) = 0 Then
            Xml_ParseNode = "null"
        Else
            Xml_ParseNode = """" & Xml_EscapeJson(textBuffer) & """"
        End If

        Exit Function

    End If


    ' -------------------------------------------------
    ' OBJECT NODE
    ' -------------------------------------------------
    Dim json As String
    json = "{"

    Dim first As Boolean
    first = True

    ' preserve text when mixed content exists
    If Len(textBuffer) > 0 Then
        json = json & """value"":""" & Xml_EscapeJson(textBuffer) & """"
        first = False
    End If


    Dim used() As Boolean
    ReDim used(1 To childNames.count)

    Dim i As Long, j As Long

    For i = 1 To childNames.count

        If used(i) Then GoTo NextChild

        Dim nm As String
        nm = childNames(i)

        Dim count As Long
        count = 0

        For j = 1 To childNames.count
            If childNames(j) = nm Then count = count + 1
        Next j

        If Not first Then json = json & ","
        first = False

        json = json & """" & Xml_EscapeJson(nm) & """:"

        If count = 1 Then

            json = json & childValues(i)
            used(i) = True

        Else

            json = json & "["
            Dim firstArr As Boolean
            firstArr = True

            For j = 1 To childNames.count

                If childNames(j) = nm Then

                    If Not firstArr Then json = json & ","
                    firstArr = False

                    json = json & childValues(j)
                    used(j) = True

                End If

            Next j

            json = json & "]"

        End If

NextChild:

    Next i

    json = json & "}"

    Xml_ParseNode = json

End Function

Private Function Xml_ReadName(ByVal txt As String, ByRef pos As Long) As String

    Dim startPos As Long
    startPos = pos
    
    Dim ch As String
    
    ' --- first character rules ---
    If pos > Len(txt) Then
        Err.Raise vbObjectError + 822, "XmlTextToJson", "Unexpected end while reading tag name."
    End If
    
    ch = Mid$(txt, pos, 1)
    
    ' XML names cannot start with number, dash, or dot
    If Not (ch Like "[A-Za-z_:]") Then
        Err.Raise vbObjectError + 823, "XmlTextToJson", "Invalid start of XML name."
    End If
    
    pos = pos + 1
    
    ' --- subsequent characters ---
    Do While pos <= Len(txt)
    
        ch = Mid$(txt, pos, 1)
        
        If ch Like "[A-Za-z0-9_.:-]" Then
            pos = pos + 1
        Else
            Exit Do
        End If
        
    Loop
    
    Xml_ReadName = Mid$(txt, startPos, pos - startPos)

End Function

Private Function Xml_ReadText(ByVal txt As String, ByRef pos As Long) As String

    Dim startPos As Long
    startPos = pos

    Do While pos <= Len(txt)

        If Mid$(txt, pos, 1) = "<" Then Exit Do

        pos = pos + 1

    Loop

    Dim raw As String
    raw = Mid$(txt, startPos, pos - startPos)

    Xml_ReadText = Xml_DecodeEntities(raw)

End Function

Private Sub Xml_SkipWhitespace(ByVal txt As String, ByRef pos As Long)

    Do While pos <= Len(txt)
    
        Dim ch As String
        ch = Mid$(txt, pos, 1)
        
        If ch = " " Or ch = vbCr Or ch = vbLf Or ch = vbTab Then
            pos = pos + 1
        Else
            Exit Do
        End If
        
    Loop

End Sub

Private Function Xml_EscapeJson(ByVal s As String) As String

    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim out As String

    For i = 1 To Len(s)

        ch = Mid$(s, i, 1)
        code = AscW(ch)

        Select Case code

            Case 34 ' "
                out = out & "\"""

            Case 92 ' \
                out = out & "\\"

            Case 8
                out = out & "\b"

            Case 9
                out = out & "\t"

            Case 10
                out = out & "\n"

            Case 12
                out = out & "\f"

            Case 13
                out = out & "\r"

            Case 0 To 31
                out = out & "\u" & Right$("000" & Hex$(code), 4)

            Case Else
                out = out & ch

        End Select

    Next i

    Xml_EscapeJson = out

End Function

Private Function Xml_DecodeEntities(ByVal s As String) As String

    ' -------------------------------------------------------------
    ' Decodes XML built-in entities and numeric entities.
    '
    ' Supported:
    '   &lt;   <
    '   &gt;   >
    '   &amp;  &
    '   &apos; '
    '   &quot; "
    '   &#NNN;
    '   &#xHH;
    ' -------------------------------------------------------------

    Dim i As Long
    Dim out As String
    Dim n As Long

    n = Len(s)
    i = 1

    Do While i <= n

        Dim ch As String
        ch = Mid$(s, i, 1)

        If ch <> "&" Then
            out = out & ch
            i = i + 1
        Else

            Dim semi As Long
            semi = InStr(i, s, ";")

            If semi = 0 Then
                out = out & "&"
                i = i + 1
                GoTo ContinueLoop
            End If

            Dim ent As String
            ent = Mid$(s, i + 1, semi - i - 1)

            Select Case ent

                Case "lt"
                    out = out & "<"

                Case "gt"
                    out = out & ">"

                Case "amp"
                    out = out & "&"

                Case "apos"
                    out = out & "'"

                Case "quot"
                    out = out & """"

                Case Else

                    ' numeric entities
                    If Left$(ent, 1) = "#" Then

                        Dim code As Long

                        If LCase$(Mid$(ent, 2, 1)) = "x" Then
                            code = CLng("&H" & Mid$(ent, 3))
                        Else
                            code = CLng(Mid$(ent, 2))
                        End If

                        out = out & ChrW$(code)

                    Else
                        ' unknown entity: leave literal
                        out = out & "&" & ent & ";"
                    End If

            End Select

            i = semi + 1

        End If

ContinueLoop:
    Loop

    Xml_DecodeEntities = out

End Function

Private Sub Xml_SkipAttributes(ByVal txt As String, ByRef pos As Long)

    Dim ch As String
    Dim quote As String

    Do While pos <= Len(txt)

        ch = Mid$(txt, pos, 1)

        Select Case ch

            Case """", "'"
                quote = ch
                pos = pos + 1

                Do While pos <= Len(txt) And Mid$(txt, pos, 1) <> quote
                    pos = pos + 1
                Loop

            Case ">", "/"
                Exit Sub

        End Select

        pos = pos + 1

    Loop

End Sub

Private Function XmlJson_CollapseValueNodes(ByVal jsonText As String) As String

    ' Converts patterns like:
    '   "sku":{"value":"A100"}
    ' into:
    '   "sku":"A100"

    Dim i As Long
    Dim result As String
    
    result = jsonText
    
    result = Replace(result, "{""value"":", "")
    
    ' remove closing braces that belong to collapsed value nodes
    result = Replace(result, "}}", "}")
    
    XmlJson_CollapseValueNodes = result

End Function

' =============================================================================
' Json_CoalesceChildArrays
'
' Purpose:
'   Extract nested child arrays from a parent JSON array and merge them into
'   a single array. Supports optional injection of parent values, literals,
'   and Excel formulas via a key map.
'
' Example:
'   Parent JSON
'       $.orders
'           + items[]
'
'   Result:
'       flattened array of items with optional injected fields
'
' Key Map Format:
'   Collection of Array(source, destination)
'
'   source modes:
'       "orderId"      -> copy parent field
'       "'=ROW()"      -> inject Excel formula
'       "'orders"      -> inject literal constant
'
' Parameters:
'   parentJson      - Source JSON document
'   parentRoot      - Path to parent array (ex: "$.orders")
'   childProperty   - Name of child array property (ex: "items")
'   strictMode      - If True, validates child object shape consistency
'   parentKeyMap    - Optional mapping collection for injected fields
'
' Returns:
'   JSON string representing merged child array
'
' Notes:
'   - Deterministic traversal
'   - No Excel dependencies
'   - Literal injection uses leading `'`
' =============================================================================
Public Function Json_CoalesceChildArrays( _
    ByVal parentJson As String, _
    ByVal parentRoot As String, _
    ByVal childProperty As String, _
    Optional ByVal strictMode As Boolean = False, _
    Optional ByVal parentKeyMap As Collection = Nothing _
) As String

    Const ERR_SRC As String = "Json_CoalesceChildArrays"

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

    Dim found As Boolean
    Dim row As Collection
    Dim i As Long
    Dim pair As Variant

    Dim parentVal As Variant
    Dim keyPair As Variant
    Dim SRC As String

    For Each parentObj In parents

        ' ------------------------------------------------
        ' child array lookup
        ' ------------------------------------------------
        found = Json_TryObjGet(parentObj, childProperty, childArr)

        If Not found Then GoTo NextParent
        If IsNull(childArr) Then GoTo NextParent

        If (Not IsObject(childArr)) Or (TypeName(childArr) <> "Collection") Then
            Err.Raise vbObjectError + 5302, ERR_SRC, _
                "Child property is not an array: '" & childProperty & "'"
        End If

        For Each childObj In childArr

            ' ------------------------------------------------
            ' strict validation
            ' ------------------------------------------------
            If strictMode Then

                If childObj.count = 0 Or childObj(1) <> "__OBJ__" Then
                    Err.Raise vbObjectError + 5303, ERR_SRC, _
                        "Strict mode requires arrays of objects"
                End If

                If Not shapeCaptured Then
                    Set firstShape = Json_ObjectShape(childObj)
                    shapeCaptured = True
                Else
                    If Not Json_ObjectShapeMatches(childObj, firstShape) Then
                        Err.Raise vbObjectError + 5304, ERR_SRC, _
                            "Child array shapes are inconsistent"
                    End If
                End If

            End If

            ' ------------------------------------------------
            ' clone child object
            ' ------------------------------------------------
            Set row = New Collection
            row.Add "__OBJ__"

            For i = 2 To childObj.count
                pair = childObj(i)
                row.Add pair
            Next i

            ' ------------------------------------------------
            ' parent / literal injection
            ' ------------------------------------------------
            If Not parentKeyMap Is Nothing Then

                For Each keyPair In parentKeyMap

                    SRC = CStr(keyPair(0))

                    ' literal injection
                    If Left$(SRC, 1) = "'" Then

                        Json_ObjSet row, keyPair(1), Mid$(SRC, 2)

                    Else

                        If Not Json_TryObjGet(parentObj, SRC, parentVal) Then
                            Err.Raise vbObjectError + 5301, ERR_SRC, _
                                "Parent key not found: '" & SRC & "'"
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

' =============================================================================
' Json_CoalesceArraysFromStrings
'
' Purpose:
'   Merge multiple JSON arrays (provided as strings) into a single array.
'
' Parameters:
'   jsonStrings - Collection of JSON array strings
'   strictMode  - If True, ensures all objects share identical shape
'
' Returns:
'   JSON string containing the merged array
'
' Notes:
'   - Each string must represent a JSON array
'   - Deterministic ordering preserved
'   - Fast-fails if strict shape validation fails
' =============================================================================
Public Function Json_CoalesceArraysFromStrings( _
    ByVal jsonStrings As Collection, _
    Optional ByVal strictMode As Boolean = False _
) As String

    Const ERR_SRC As String = "Json_CoalesceArraysFromStrings"
    
    Dim result As New Collection
    
    Dim firstShape As Collection
    Dim shapeCaptured As Boolean
    
    Dim i As Long
    
    For i = 1 To jsonStrings.count
    
        Dim txt As String
        txt = jsonStrings(i)
        
        Dim parsed As Variant
        Json_ParseInto txt, parsed
        
        If Not TypeOf parsed Is Collection Then
            Err.Raise vbObjectError + 5201, ERR_SRC, _
                "Value is not a JSON array"
        End If
        
        Dim arr As Collection
        Set arr = parsed
        
        Dim obj As Variant
        
        For Each obj In arr
        
            If strictMode Then
            
                If Not TypeOf obj Is Collection Then
                    Err.Raise vbObjectError + 5202, ERR_SRC, _
                        "Strict mode requires arrays of objects"
                End If
                
                If obj.count = 0 Or obj(1) <> "__OBJ__" Then
                    Err.Raise vbObjectError + 5203, ERR_SRC, _
                        "Strict mode requires arrays of objects"
                End If
                
                If Not shapeCaptured Then
                
                    Set firstShape = Json_ObjectShape(obj)
                    shapeCaptured = True
                    
                Else
                
                    If Not Json_ObjectShapeMatches(obj, firstShape) Then
                        Err.Raise vbObjectError + 5204, ERR_SRC, _
                            "Array object shapes are inconsistent"
                    End If
                    
                End If
                
            End If
            
            result.Add obj
            
        Next obj
        
    Next i
    
    Json_CoalesceArraysFromStrings = Json_Stringify(result)

End Function

' =============================================================================
' Json_CoalesceArraysFromRange
'
' Purpose:
'   Merge JSON arrays stored in an Excel range into a single array.
'
' Parameters:
'   rng        - Range containing JSON array strings
'   strictMode - Optional strict object shape validation
'
' Returns:
'   JSON string containing merged arrays
'
' Notes:
'   - Empty cells are ignored
'   - Uses Value2 bulk read for performance
' =============================================================================
Public Function Json_CoalesceArraysFromRange( _
    ByVal rng As Range, _
    Optional ByVal strictMode As Boolean = False _
) As String

    Dim jsonStrings As Collection
    Set jsonStrings = Excel_RangeToJsonStrings(rng)
    
    Json_CoalesceArraysFromRange = _
        Json_CoalesceArraysFromStrings(jsonStrings, strictMode)

End Function

' =============================================================================
' Excel_RangeToJsonStrings
'
' Purpose:
'   Convert a worksheet range into a collection of JSON strings.
'
' Parameters:
'   rng - Range containing JSON text values
'
' Returns:
'   Collection of non-empty trimmed strings
'
' Notes:
'   - Uses bulk Value2 read for performance
'   - Ignores empty cells
'   - Handles single-cell ranges
' =============================================================================
Public Function Excel_RangeToJsonStrings( _
    ByVal rng As Range _
) As Collection

    Dim result As New Collection
    Dim data As Variant
    Dim r As Long, c As Long
    
    data = rng.Value2
    
    If IsArray(data) Then
    
        For r = LBound(data, 1) To UBound(data, 1)
            For c = LBound(data, 2) To UBound(data, 2)
            
                Dim txt As String
                txt = Trim$(CStr(data(r, c)))
                
                If Len(txt) > 0 Then
                    result.Add txt
                End If
                
            Next c
        Next r
        
    Else
    
        Dim singleVal As String
        singleVal = Trim$(CStr(data))
        
        If Len(singleVal) > 0 Then
            result.Add singleVal
        End If
        
    End If
    
    Set Excel_RangeToJsonStrings = result

End Function

' =============================================================================
' Json_ResolveArrayPath
'
' Purpose:
'   Resolve a simple JSON path (ex: "$.products") to an array collection.
'
' Parameters:
'   root - Parsed JSON root object
'   path - Path expression starting with "$"
'
' Returns:
'   Collection representing the resolved array
'
' Notes:
'   - Supports simple dot traversal only
'   - Does not implement full JSONPath
' =============================================================================
Public Function Json_ResolveArrayPath( _
    ByVal root As Variant, _
    ByVal path As String _
) As Collection

    Const ERR_SRC As String = "Json_ResolveArrayPath"

    If Len(path) = 0 Then
        Err.Raise vbObjectError + 5310, ERR_SRC, "Path cannot be empty"
    End If

    If path = "$" Then
        If TypeOf root Is Collection Then
            Set Json_ResolveArrayPath = root
            Exit Function
        End If
        
        Err.Raise vbObjectError + 5311, ERR_SRC, _
            "Root is not an array"
    End If

    If Left$(path, 2) <> "$." Then
        Err.Raise vbObjectError + 5312, ERR_SRC, _
            "Path must begin with '$.'"
    End If

    Dim parts() As String
    parts = Split(Mid$(path, 3), ".")

    Dim current As Object
    Set current = root
    
    Dim i As Integer
    For i = LBound(parts) To UBound(parts)
    
        If Not TypeOf current Is Collection Then
            Err.Raise vbObjectError + 5313, ERR_SRC, _
                "Path traversal encountered non-object"
        End If
    
        Set current = Json_ObjGet(current, parts(i))
    
    Next i

    If Not TypeOf current Is Collection Then
        Err.Raise vbObjectError + 5315, ERR_SRC, _
            "Resolved path is not an array"
    End If

    Set Json_ResolveArrayPath = current

End Function

Private Function Json_ObjectShape(ByVal obj As Collection) As Collection

    Dim shape As New Collection
    Dim i As Long
    
    For i = 2 To obj.count
        shape.Add obj(i)(0)
    Next i
    
    Set Json_ObjectShape = shape

End Function

Private Function Json_ObjectShapeMatches( _
    ByVal obj As Collection, _
    ByVal shape As Collection _
) As Boolean

    Dim i As Long
    
    If obj.count - 1 <> shape.count Then Exit Function
    
    For i = 1 To shape.count
        If obj(i + 1)(0) <> shape(i) Then Exit Function
    Next i
    
    Json_ObjectShapeMatches = True

End Function

Private Sub Excel_ReleaseVariantObject(ByRef v As Variant)
    If IsObject(v) Then
        Set v = Nothing
    Else
        v = Empty
    End If
End Sub

Private Sub Json_RowObjectCollectHeaders( _
    ByVal obj As Collection, _
    ByVal prefix As String, _
    ByVal nonTableArraysAsJson As Boolean, _
    ByRef hdrs() As String, _
    ByRef hdrCount As Long, _
    ByRef slotHash() As Long, _
    ByRef slotIdx() As Long, _
    ByRef cap As Long _
)
    Dim i As Long
    For i = 2 To obj.count
        Dim pair As Variant
        pair = obj(i)

        Dim seg As String
        seg = Json_EscapePathSegment(CStr(pair(0)))

        Dim path As String
        If Len(prefix) = 0 Then
            path = seg
        Else
            path = prefix & "." & seg
        End If

        Json_RowValueCollectHeaders pair(1), path, nonTableArraysAsJson, hdrs, hdrCount, slotHash, slotIdx, cap
    Next i
End Sub

Private Sub Json_RowValueCollectHeaders( _
    ByVal v As Variant, _
    ByVal path As String, _
    ByVal nonTableArraysAsJson As Boolean, _
    ByRef hdrs() As String, _
    ByRef hdrCount As Long, _
    ByRef slotHash() As Long, _
    ByRef slotIdx() As Long, _
    ByRef cap As Long _
)
    If Not IsObject(v) Then
        HeaderTable_Ensure path, hdrs, hdrCount, slotHash, slotIdx, cap, False
        Exit Sub
    End If

    If TypeName(v) <> "Collection" Then
        HeaderTable_Ensure path, hdrs, hdrCount, slotHash, slotIdx, cap, False
        Exit Sub
    End If

    If Json_IsObject(v) Then
        Json_RowObjectCollectHeaders v, path, nonTableArraysAsJson, hdrs, hdrCount, slotHash, slotIdx, cap
        Exit Sub
    End If

    If nonTableArraysAsJson Then
        HeaderTable_Ensure path, hdrs, hdrCount, slotHash, slotIdx, cap, False
    End If
End Sub

Private Sub Json_RowObjectFillRow( _
    ByVal obj As Collection, _
    ByVal prefix As String, _
    ByVal nonTableArraysAsJson As Boolean, _
    ByRef hdrs() As String, _
    ByRef slotHash() As Long, _
    ByRef slotIdx() As Long, _
    ByVal cap As Long, _
    ByRef outData As Variant, _
    ByVal rowNumber As Long _
)
    Dim i As Long
    For i = 2 To obj.count
        Dim pair As Variant
        pair = obj(i)

        Dim seg As String
        seg = Json_EscapePathSegment(CStr(pair(0)))

        Dim path As String
        If Len(prefix) = 0 Then
            path = seg
        Else
            path = prefix & "." & seg
        End If

        Json_RowValueFill pair(1), path, nonTableArraysAsJson, hdrs, slotHash, slotIdx, cap, outData, rowNumber
    Next i
End Sub

Private Sub Json_RowValueFill( _
    ByVal v As Variant, _
    ByVal path As String, _
    ByVal nonTableArraysAsJson As Boolean, _
    ByRef hdrs() As String, _
    ByRef slotHash() As Long, _
    ByRef slotIdx() As Long, _
    ByVal cap As Long, _
    ByRef outData As Variant, _
    ByVal rowNumber As Long _
)
    Dim col As Long

    If Not IsObject(v) Then
        col = HeaderTable_Find(path, hdrs, slotHash, slotIdx, cap)
        If col > 0 Then outData(rowNumber, col) = v
        Exit Sub
    End If

    If TypeName(v) <> "Collection" Then
        col = HeaderTable_Find(path, hdrs, slotHash, slotIdx, cap)
        If col > 0 Then outData(rowNumber, col) = CStr(TypeName(v))
        Exit Sub
    End If

    If Json_IsObject(v) Then
        Json_RowObjectFillRow v, path, nonTableArraysAsJson, hdrs, slotHash, slotIdx, cap, outData, rowNumber
        Exit Sub
    End If

    If nonTableArraysAsJson Then
        col = HeaderTable_Find(path, hdrs, slotHash, slotIdx, cap)
        If col > 0 Then outData(rowNumber, col) = Json_Stringify(v)
    End If
End Sub

Private Function Excel_GetRowsPerChunk(ByVal colCount As Long) As Long
    If colCount <= 0 Then
        Excel_GetRowsPerChunk = 1
        Exit Function
    End If

    Excel_GetRowsPerChunk = EXCEL_CHUNK_MAX_CELLS \ colCount
    If Excel_GetRowsPerChunk < 1 Then Excel_GetRowsPerChunk = 1
End Function
