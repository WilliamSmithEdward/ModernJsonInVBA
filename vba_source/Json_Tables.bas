Attribute VB_Name = "Json_Tables"
Option Explicit

' =============================================================================
' Module:      Json_Tables
' Project:     ModernJsonInVBA
' Version:     3.8.2
' Released:    2026-07-09
'
' Turns parsed/flattened JSON into tabular data:
'
'   Json_ExtractTableRows        flattened object -> Collection of row objects
'   Json_TableTo2D               row objects -> headers + 2D Variant array
'
' Determinism:
'   Header discovery is first-seen order across rows. Row order follows path
'   order in the flattened input. Both are stable across runs.
'
' No Excel dependencies live here; the ListObject layer is Json_Excel and
' array merging is Json_Coalesce.
'
' Error numbers (all vbObjectError + n):
'   920  ExtractTableRows expects tagged object
' =============================================================================

Private Const ERR_SRC As String = "ModernJsonInVBA"

' =============================================================================
' Public API: table extraction
' =============================================================================

' Extract table rows (tagged objects) from a flattened object.
'
' tableRoot examples:
'   "$"                root array
'   "$.orders.items"   nested arrays (parent indices in paths are supported)
'
' Columns whose path still contains "[" (nested child tables) are excluded.
' Rows are created on first sight of any of their columns; rows addressed
' positionally (root immediately indexed) are padded so indices line up.
Public Function Json_ExtractTableRows(ByVal flatObj As Collection, ByVal tableRoot As String) As Collection
    If Not Json_IsObject(flatObj) Then
        Err.Raise vbObjectError + 920, ERR_SRC, "ExtractTableRows expects tagged object"
    End If

    Dim rows As New Collection

    ' rowKey -> row object, first-seen order (for nested, index-bearing keys)
    Dim rowIdx As JsonStringIndex
    Dim rowObjs() As Collection

    ' Positional rows (fast path) in a plain array: Collection.Item(i) walks
    ' a linked list, so per-pair indexed lookups would be quadratic.
    Dim rowsByIdx() As Collection
    Dim rowsByIdxCount As Long
    rowsByIdxCount = 0

    ' Compile tableRoot segments once; reused for every path.
    Dim rootSegs() As String
    Dim rootSegCount As Long
    Json_BuildRootSegs tableRoot, rootSegs, rootSegCount

    Dim fastPrefix As String
    fastPrefix = tableRoot & "["

    Dim fastPrefixLen As Long
    fastPrefixLen = Len(fastPrefix)

    ' Column-set tracking for the row currently being filled. Flatten emits
    ' each array element's paths contiguously, so pairs for a row arrive as
    ' one run; a per-row hash check makes each insert O(1) instead of the
    ' O(columns) scan Json_ObjSet would need. Duplicate column paths (only
    ' possible when the source JSON object repeats a key) fall back to
    ' Json_ObjSet so overwrite-in-place semantics are preserved exactly.
    Dim curRow As Collection
    Dim curRowFresh As Boolean
    Dim colSeen As JsonStringIndex

    ' For Each (enumerator) instead of flatObj(i): indexed access walks the
    ' Collection's linked list and turns this loop quadratic on big inputs.
    Dim isFirst As Boolean
    isFirst = True

    Dim kv As Variant
    For Each kv In flatObj
        If isFirst Then
            ' Slot (1) is the object tag, not a pair.
            isFirst = False
        Else
            Dim path As String
            path = CStr(kv(0))

            Dim idx As Long
            Dim colPath As String
            Dim rowKey As String
            Dim ok As Boolean

            Dim usedIndexedFastPath As Boolean
            usedIndexedFastPath = False

            ' Fast path: the row index directly follows tableRoot ("root[n]...").
            If Left$(path, fastPrefixLen) = fastPrefix Then
                ok = Json_TryParseIndexedPath(path, tableRoot, idx, colPath, rowKey)
                usedIndexedFastPath = ok
            Else
                ok = Json_TryParseTableRowPath(path, tableRoot, rootSegs, rootSegCount, idx, colPath, rowKey)
            End If

            If ok Then
                ' Exclude child-table columns (nested arrays under this row).
                If InStr(1, colPath, "[", vbBinaryCompare) = 0 Then
                    Dim rowObj As Collection

                    If usedIndexedFastPath Then
                        Set rowObj = RowsByIdx_Ensure(rowsByIdx, rowsByIdxCount, rows, idx)
                    Else
                        Set rowObj = RowMap_GetOrAdd(rowIdx, rowObjs, rowKey, rows)
                    End If

                    If Not (rowObj Is curRow) Then
                        Set curRow = rowObj
                        ' Blind-append mode only for rows with no pairs yet; a
                        ' revisited row (defensive: should not happen with
                        ' contiguous emission) uses the safe Json_ObjSet path.
                        curRowFresh = (curRow.count <= 1)
                        If curRowFresh Then JsonIdx_Init colSeen, 16
                    End If

                    Dim v As Variant
                    Json_VarAssign v, kv(1)

                    If curRowFresh Then
                        Dim seenBefore As Long
                        seenBefore = colSeen.count
                        JsonIdx_Ensure colSeen, colPath

                        If colSeen.count > seenBefore Then
                            rowObj.Add Array(colPath, v)
                        Else
                            Json_ObjSet rowObj, colPath, v
                        End If
                    Else
                        Json_ObjSet rowObj, colPath, v
                    End If
                End If
            End If
        End If
    Next kv

    Set Json_ExtractTableRows = rows
End Function

' Positional row access for root-indexed paths: pads rows (and the parallel
' array) with empty tagged objects so 1-based position (idx + 1) exists,
' then returns it via the O(1) array.
Private Function RowsByIdx_Ensure( _
    ByRef rowsByIdx() As Collection, _
    ByRef rowsByIdxCount As Long, _
    ByVal rows As Collection, _
    ByVal idx As Long _
) As Collection

    Dim needCount As Long
    needCount = idx + 1

    If rowsByIdxCount = 0 Then
        Dim initialCap As Long
        initialCap = 16
        Do While initialCap < needCount
            initialCap = initialCap * 2
        Loop
        ReDim rowsByIdx(1 To initialCap) As Collection
    ElseIf needCount > UBound(rowsByIdx) Then
        Dim newCap As Long
        newCap = UBound(rowsByIdx)
        Do While newCap < needCount
            newCap = newCap * 2
        Loop
        ReDim Preserve rowsByIdx(1 To newCap) As Collection
    End If

    Do While rowsByIdxCount < needCount
        Dim o As New Collection
        o.Add JSON_TAG_OBJECT
        rows.Add o

        rowsByIdxCount = rowsByIdxCount + 1
        Set rowsByIdx(rowsByIdxCount) = o

        Set o = Nothing
    Loop

    Set RowsByIdx_Ensure = rowsByIdx(needCount)
End Function

' Convert a Collection of tagged row objects into:
'   - headers: 1-based Variant array of column names (first-seen order)
'   - return:  2D Variant array (1..rowCount, 1..colCount), or Empty when
'              there are no rows
'
' Zero rows        => headers = ["value"], returns Empty.
' Rows but no keys => headers = ["value"], returns rowCount x 1 empty cells.
Public Function Json_TableTo2D(ByVal rows As Collection, ByRef headers As Variant) As Variant
    Dim rowCount As Long
    rowCount = rows.count

    If rowCount = 0 Then
        ReDim headers(1 To 1) As Variant
        headers(1) = "value"
        Json_TableTo2D = Empty
        Exit Function
    End If

    ' Pass 1: discover headers in first-seen order. Rows enumerate with
    ' For Each (indexed Collection access is O(index) per hit); the pairs
    ' inside one row also enumerate, skipping the leading tag.
    Dim hdrIdx As JsonStringIndex
    JsonIdx_Init hdrIdx, 64

    Dim rowVar As Variant
    Dim pair As Variant
    Dim rowFirst As Boolean

    For Each rowVar In rows
        rowFirst = True
        For Each pair In rowVar
            If rowFirst Then
                rowFirst = False
            Else
                JsonIdx_Ensure hdrIdx, CStr(pair(0))
            End If
        Next pair
    Next rowVar

    If hdrIdx.count = 0 Then
        ReDim headers(1 To 1) As Variant
        headers(1) = "value"

        Dim data0 As Variant
        ReDim data0(1 To rowCount, 1 To 1) As Variant
        Json_TableTo2D = data0
        Exit Function
    End If

    ReDim headers(1 To hdrIdx.count) As Variant
    Dim c As Long
    For c = 1 To hdrIdx.count
        headers(c) = hdrIdx.keys(c)
    Next c

    ' Pass 2: place values by header index.
    Dim data As Variant
    ReDim data(1 To rowCount, 1 To hdrIdx.count) As Variant

    Dim r As Long
    r = 0

    For Each rowVar In rows
        r = r + 1
        rowFirst = True

        For Each pair In rowVar
            If rowFirst Then
                rowFirst = False
            Else
                Dim col2 As Long
                col2 = JsonIdx_Find(hdrIdx, CStr(pair(0)))

                If col2 > 0 Then
                    data(r, col2) = pair(1)
                End If
            End If
        Next pair
    Next rowVar

    Json_TableTo2D = data
End Function

' =============================================================================
' Internal: single-pass row filling with header discovery
'
' Used by Json_Excel to stream array-of-object rows into ONE 2D array in a
' single pass: unseen column paths register in headerIdx on the fly, and the
' array's column dimension grows as needed (ReDim Preserve is legal on the
' last dimension). Earlier rows keep Empty in late-appearing columns, which
' matches the old two-pass collect/fill semantics exactly.
'
' Nested objects contribute dotted column paths; nested arrays are included
' (as JSON text) only when nonTableArraysAsJson is True.
' =============================================================================

Public Sub Json_RowObjectFillRow( _
    ByVal obj As Collection, _
    ByVal prefix As String, _
    ByVal nonTableArraysAsJson As Boolean, _
    ByRef headerIdx As JsonStringIndex, _
    ByRef outData As Variant, _
    ByVal rowNumber As Long _
)
    Dim isFirst As Boolean
    isFirst = True

    Dim pair As Variant
    For Each pair In obj
        If isFirst Then
            isFirst = False
        Else
            Json_RowValueFill pair(1), _
                RowPath_Append(prefix, CStr(pair(0))), _
                nonTableArraysAsJson, headerIdx, outData, rowNumber
        End If
    Next pair
End Sub

Private Function RowPath_Append(ByRef prefix As String, ByVal key As String) As String
    Dim seg As String
    seg = Json_EscapePathSegment(key)

    If Len(prefix) = 0 Then
        RowPath_Append = seg
    Else
        RowPath_Append = prefix & "." & seg
    End If
End Function

Private Sub Json_RowValueFill( _
    ByVal v As Variant, _
    ByVal path As String, _
    ByVal nonTableArraysAsJson As Boolean, _
    ByRef headerIdx As JsonStringIndex, _
    ByRef outData As Variant, _
    ByVal rowNumber As Long _
)
    Dim col As Long

    If Not IsObject(v) Then
        col = JsonIdx_Ensure(headerIdx, path)
        Json_Grow2DCols outData, col
        outData(rowNumber, col) = v
        Exit Sub
    End If

    If TypeName(v) <> "Collection" Then
        col = JsonIdx_Ensure(headerIdx, path)
        Json_Grow2DCols outData, col
        outData(rowNumber, col) = CStr(TypeName(v))
        Exit Sub
    End If

    If Json_IsObject(v) Then
        Json_RowObjectFillRow v, path, nonTableArraysAsJson, headerIdx, outData, rowNumber
        Exit Sub
    End If

    ' JSON array: only a column when arrays are kept as JSON text.
    If nonTableArraysAsJson Then
        col = JsonIdx_Ensure(headerIdx, path)
        Json_Grow2DCols outData, col
        outData(rowNumber, col) = Json_Stringify(v)
    End If
End Sub

' =============================================================================
' Internal: row keying
' =============================================================================

' Map rowKey -> row object, creating (and appending to rows) on first sight.
Private Function RowMap_GetOrAdd( _
    ByRef rowIdx As JsonStringIndex, _
    ByRef rowObjs() As Collection, _
    ByVal rowKey As String, _
    ByVal rows As Collection _
) As Collection

    Dim before As Long
    before = rowIdx.count

    Dim k As Long
    k = JsonIdx_Ensure(rowIdx, rowKey)

    If rowIdx.count > before Then
        ' New row.
        If before = 0 Then
            ReDim rowObjs(1 To 16) As Collection
        ElseIf k > UBound(rowObjs) Then
            ReDim Preserve rowObjs(1 To UBound(rowObjs) * 2) As Collection
        End If

        Dim o As New Collection
        o.Add JSON_TAG_OBJECT
        rows.Add o

        Set rowObjs(k) = o
    End If

    Set RowMap_GetOrAdd = rowObjs(k)
End Function

' =============================================================================
' Internal: table path parsing
' =============================================================================

' Parse a path of the form tableRoot & "[idx]" & optional "." & colPath.
' Returns the row index, the column path, and the row key
' (tableRoot & "[idx]"). A row with no member path gets colPath "value".
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

' Pre-split tableRoot ("$.a.b") into segments so per-path matching does not
' re-tokenize the root.
Private Sub Json_BuildRootSegs(ByVal tableRoot As String, ByRef rootSegs() As String, ByRef rootSegCount As Long)
    rootSegCount = 0

    tableRoot = Trim$(tableRoot)
    If Len(tableRoot) = 0 Then Exit Sub
    If Left$(tableRoot, 2) <> "$." Then Exit Sub

    Dim toks As Collection
    Set toks = Json_TokenizePath(Mid$(tableRoot, 3))
    If toks.count = 0 Then Exit Sub

    rootSegCount = toks.count
    ReDim rootSegs(1 To rootSegCount) As String

    Dim i As Long
    For i = 1 To rootSegCount
        rootSegs(i) = CStr(toks(i))
    Next i
End Sub

' Match a full path against the compiled root segments, tolerating "[n]"
' after any segment (parent arrays). The LAST segment must carry an index -
' that index identifies the row. The row key is the matched path prefix
' including all indices, which keeps rows distinct across parents.
Private Function Json_TryParseTableRowPath( _
    ByVal fullPath As String, _
    ByVal tableRoot As String, _
    ByRef rootSegs() As String, _
    ByVal rootSegCount As Long, _
    ByRef outIndex As Long, _
    ByRef outColPath As String, _
    ByRef outRowKey As String _
) As Boolean

    Json_TryParseTableRowPath = False
    outIndex = 0
    outColPath = vbNullString
    outRowKey = vbNullString

    If rootSegCount = 0 Then Exit Function
    If Len(fullPath) = 0 Or Len(tableRoot) = 0 Then Exit Function
    If Left$(tableRoot, 2) <> "$." Then Exit Function
    If Left$(fullPath, 2) <> "$." Then Exit Function

    Dim pos As Long
    pos = 3   ' after "$." in fullPath

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
                ' The table root segment itself must be indexed.
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
        Json_TryParseTableRowPath = True
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

    Json_TryParseTableRowPath = True
End Function
