Option Explicit

' =============================================================================
' mJsonTests
' -----------------------------------------------------------------------------
' Purpose
'   Lightweight, zero-dependency test harness for mJsonOneFile JSON library.
'
' Usage
'   Run: Json_RunAllTests
'
' Conventions
'   - Each test is a single Sub named Test_<Area>_<Scenario>
'   - Each test has a doc header describing:
'       * Goal
'       * Input
'       * Expected result
'       * Notes / known limitations
'   - Assertions throw a consistent error number range (vbObjectError + 610+)
' =============================================================================

' =============================================================================
' PUBLIC RUNNER
' =============================================================================

Public Sub Json_RunAllTests()
    ' Parse
    Test_Parse_Literals
    Test_Parse_Number
    Test_Parse_Number_RejectLeadingZero
    Test_Parse_TrailingCharacters
    Test_Parse_Number_Exponent_Negative

    Test_Parse_Array_Empty
    Test_Parse_Array_Simple
    Test_Parse_Array_Nested
    Test_Parse_Object_Empty
    Test_Parse_Object_Simple
    Test_Parse_Object_Nested
    Test_Parse_StringRejectsRawNewline
    Test_Parse_IntegerType

    ' Stringify
    Test_Stringify_RoundTrip_Primitives
    Test_Stringify_ObjectAndArray

    ' Flatten
    Test_Flatten_Simple
    Test_Flatten_Nested
    Test_Flatten_ArrayIndexed_Primitives
    Test_Flatten_ArrayIndexed_Objects
    Test_Flatten_KeyEscaping
    Test_Flatten_RootPrimitiveStoredAtDollar
    Test_Flatten_RootArrayIndexed
    Test_Flatten_PathEscapes_DotAndBackslash

    ' Flat access helpers
    Test_FlatGet
    Test_FlatContains

    ' Unflatten
    Test_Unflatten_Simple
    Test_Unflatten_KeyEscaping_RoundTrip
    Test_Unflatten_RejectsArrayPaths
    Test_Unflatten_RoundTrip_DotAndBackslash

    ' Find array-of-object roots
    Test_FindArrayObjectRoots
    Test_FindArrayObjectRoots_RootArray
    Test_FindArrayObjectRoots_DoesNotFalsePositive_OnArrayOfPrimitives

    ' ExtractTableRows
    Test_ExtractTableRows_Simple
    Test_ExtractTableRows_RootArray
    Test_ExtractTableRows_ExcludesNestedArrays_ChildTablesOnly

    ' TableTo2D
    Test_TableTo2D_EmptyRows
    Test_TableTo2D_ColumnPlacement_VariedOrder
    Test_TableTo2D_Basic
    Test_TableTo2D_WideRow_ManyHeaders
    Test_TableTo2D_WideRows_VariedPresence
    Test_TableTo2D_WideSchema_DoesNotHang_AndIsComplete
    Test_TableTo2D_HeaderOrder_FirstSeenAcrossRows
    Test_TableTo2D_RowsExist_ButNoKeys

    ' ObjSet
    Test_ObjSet_PreservesOrder_OnOverwrite
    Test_ObjSet_Overwrite_PreservesPosition

    ' Unicode / Emoji
    Test_Parse_UnicodeEscape4_BasicBMP
    Test_Parse_UnicodeEscape_SurrogatePair_Emoji
    Test_Parse_UnicodeEscape_InvalidHighWithoutLow
    Test_Parse_UnicodeEscape_InvalidLowWithoutHigh
    Test_Stringify_DoesNotEscapeNegativeAscW
    Test_Parse_UnicodeEscape_SurrogatePair_RoundTrip_WithText
    Test_Parse_String_Long
    Test_Parse_String_Long_WithEscapes

    ' Stringify (perf/correctness on large structures)
    Test_Stringify_LargeArray_RoundTrip
    Test_Stringify_LargeObject_RoundTrip
    
    ' Primitive Parsing
    Test_Parse_Primitive_Root_Set_ShouldFail_AndIsCaught
    Test_Excel_UpsertListObjectFromJsonAtRoot_PrimitiveRoot_Raises1130
    Test_ParseInto_Primitive_DoesNotBecomeObject
    Test_ParseInto_Object_IsObject
    Test_ObjGetObject_ReturnsObject_AndCanBeReSet
    Test_TableTo2D_Null_WritesAsNullVariant
    
    ' Nested Child Rows
    Test_ExtractTableRows_NestedChildRows_DoNotCollideAcrossParents
    Test_ExtractTableRows_NestedChildRows_OrderIsFirstSeenPathOrder
    Test_ExtractTableRows_NestedChildRows_AllowsObjectColumns_ButExcludesNestedArrays
    Test_ExtractTableRows_NestedChildRows_TableTo2D_ProducesExpectedHeaders
    Test_ExtractTableRows_NestedChildRows_EmptyChildArray_ReturnsZeroRows
    
    ' Path parsing edge cases
    Test_FlatGet_EscapedDotAndBackslash_ExactMatch
    Test_FlatContains_DoesNotConfuseEscapes

    ' ExtractTableRows root validation
    Test_ExtractTableRows_TableRootMustBeArrayOfObjects
    Test_ExtractTableRows_TableRootMustExist
    Test_ExtractTableRows_RootArray_PrimitiveRowsRejected

    MsgBox "All JSON Parser tests passed.", vbInformation
End Sub


' =============================================================================
' ASSERTIONS
' =============================================================================

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 610, "mJsonTests", "ASSERT FAIL: " & message
End Sub

Private Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal message As String)
    ' ---- Null handling ----
    If IsNull(expected) Then
        If IsNull(actual) Then Exit Sub
        Err.Raise vbObjectError + 611, "mJsonTests", _
            "ASSERT FAIL: " & message & " | expected=<Null> actual=" & SafeToString(actual)
    End If

    If IsNull(actual) Then
        Err.Raise vbObjectError + 611, "mJsonTests", _
            "ASSERT FAIL: " & message & " | expected=" & SafeToString(expected) & " actual=<Null>"
    End If

    ' ---- Empty handling (optional, but avoids “Invalid use of Empty” edge cases) ----
    If IsEmpty(expected) Then
        If IsEmpty(actual) Then Exit Sub
        Err.Raise vbObjectError + 611, "mJsonTests", _
            "ASSERT FAIL: " & message & " | expected=<Empty> actual=" & SafeToString(actual)
    End If

    If IsEmpty(actual) Then
        Err.Raise vbObjectError + 611, "mJsonTests", _
            "ASSERT FAIL: " & message & " | expected=" & SafeToString(expected) & " actual=<Empty>"
    End If

    ' ---- Normal compare ----
    If expected <> actual Then
        Err.Raise vbObjectError + 611, "mJsonTests", _
            "ASSERT FAIL: " & message & " | expected=" & SafeToString(expected) & " actual=" & SafeToString(actual)
    End If
End Sub

Private Function SafeToString(ByVal v As Variant) As String
    If IsNull(v) Then
        SafeToString = "<Null>"
    ElseIf IsEmpty(v) Then
        SafeToString = "<Empty>"
    ElseIf IsObject(v) Then
        SafeToString = "<Object:" & TypeName(v) & ">"
    Else
        On Error GoTo fallback
        SafeToString = CStr(v)
        Exit Function
fallback:
        SafeToString = "<Unprintable>"
    End If
End Function

Private Sub AssertIsCollection(ByVal v As Variant, ByVal message As String)
    AssertTrue IsObject(v), message & " (not object)"
    AssertEquals "Collection", TypeName(v), message & " (not Collection)"
End Sub

Private Sub AssertIsTaggedObject(ByVal v As Variant, ByVal message As String)
    AssertTrue Obj_IsObject(v), message & " (not tagged object Collection)"
End Sub

Private Sub AssertCollectionContainsString(ByVal c As Collection, ByVal expected As String, ByVal message As String)
    Dim i As Long
    For i = 1 To c.count
        If CStr(c(i)) = expected Then Exit Sub
    Next i
    Err.Raise vbObjectError + 612, "mJsonTests", "ASSERT FAIL: " & message & " | missing=" & expected
End Sub

Private Sub Assert1DArrayEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal message As String)
    AssertTrue IsArray(expected), message & " expected not array"
    AssertTrue IsArray(actual), message & " actual not array"

    Dim expLen As Long, actLen As Long
    expLen = (UBound(expected) - LBound(expected) + 1)
    actLen = (UBound(actual) - LBound(actual) + 1)

    AssertEquals expLen, actLen, message & " length mismatch"

    Dim i As Long
    For i = 1 To expLen
        Dim expVal As String, actVal As String
        expVal = CStr(expected(LBound(expected) + i - 1))
        actVal = CStr(actual(LBound(actual) + i - 1))
        AssertEquals expVal, actVal, message & " item[" & i & "]"
    Next i
End Sub


' =============================================================================
' TESTS: Parse
' =============================================================================

Private Sub Test_Parse_Literals()
    ' Goal: Verify JSON literals are parsed into correct VBA primitives.
    ' Input: "true", "false", "null"
    ' Expect:
    '   - True -> VBA Boolean True
    '   - False -> VBA Boolean False
    '   - Null -> VBA Null Variant
    AssertEquals True, Json_Parse("true"), "true literal"
    AssertEquals False, Json_Parse("false"), "false literal"
    AssertTrue IsNull(Json_Parse("null")), "null literal"
End Sub

Private Sub Test_Parse_Number()
    ' Goal: Verify number parsing including integer, negative, decimal, exponent.
    ' Notes: Integers should come back as Long when possible (see Test_Parse_IntegerType).
    AssertEquals 5, Json_Parse("5"), "integer"
    AssertEquals -12, Json_Parse("-12"), "negative integer"
    AssertEquals 3.14, Json_Parse("3.14"), "decimal"
    AssertEquals 1200, Json_Parse("1.2e3"), "exponent"
End Sub

Private Sub Test_Parse_Array_Empty()
    ' Goal: Empty array returns Collection with Count=0.
    Dim v As Variant
    Json_ParseInto "[]", v
    AssertIsCollection v, "empty array returns Collection"
    AssertEquals 0, v.count, "empty array count"
End Sub

Private Sub Test_Parse_Array_Simple()
    ' Goal: Mixed primitive array parses in order.
    Dim arr As Collection
    Set arr = Json_Parse("[1,true,""x"",null]")

    AssertEquals 4, arr.count, "array count"
    AssertEquals 1, arr(1), "elem0"
    AssertEquals True, arr(2), "elem1"
    AssertEquals "x", arr(3), "elem2"
    AssertTrue IsNull(arr(4)), "elem3"
End Sub

Private Sub Test_Parse_Array_Nested()
    ' Goal: Nested arrays parse as nested Collections.
    Dim outer As Collection
    Set outer = Json_Parse("[[1,2],[3]]")

    AssertEquals 2, outer.count, "outer count"
    AssertEquals 2, outer(1).count, "inner0 count"
    AssertEquals 3, outer(2)(1), "inner1 elem0"
End Sub

Private Sub Test_Parse_Object_Empty()
    ' Goal: Empty object parses as tagged object Collection ("__OBJ__").
    Dim obj As Collection
    Set obj = Json_Parse("{}")

    AssertIsTaggedObject obj, "empty object"
    AssertEquals 0, Obj_CountPairs(obj), "empty object pair count"
End Sub

Private Sub Test_Parse_Object_Simple()
    ' Goal: Simple object parses as tagged object Collection with correct primitive values.
    Dim obj As Collection
    Set obj = Json_Parse("{""a"":1,""b"":true}")

    AssertIsTaggedObject obj, "object"
    AssertEquals 2, Obj_CountPairs(obj), "pair count"
    AssertEquals 1, Obj_GetValue(obj, "a"), "prop a"
    AssertEquals True, Obj_GetValue(obj, "b"), "prop b"
End Sub

Private Sub Test_Parse_Object_Nested()
    ' Goal: Object containing an array yields nested Collection as an object value.
    Dim obj As Collection
    Set obj = Json_Parse("{""x"":[1,2]}")

    Dim x As Object
    Set x = Obj_GetObject(obj, "x")

    AssertEquals "Collection", TypeName(x), "nested x is Collection"
    AssertEquals 2, x.count, "nested x count"
End Sub

Private Sub Test_Parse_StringRejectsRawNewline()
    ' Goal: Raw newline characters inside JSON strings should error (must be escaped as \n).
    On Error GoTo expected
    Call Json_Parse("[""" & vbLf & """]")
    Err.Raise vbObjectError + 699, "mJsonTests", "Expected error for raw newline in JSON string"
expected:
    AssertTrue (Err.Number <> 0), "raw newline should error"
    Err.Clear
End Sub

Private Sub Test_Parse_IntegerType()
    ' Goal: Integer tokens without decimal/exponent should come back as Long when representable.
    Dim v As Variant
    v = Json_Parse("5")
    AssertEquals vbLong, VarType(v), "integer should be Long"
End Sub


' =============================================================================
' TESTS: Stringify
' =============================================================================

Private Sub Test_Stringify_RoundTrip_Primitives()
    ' Goal: Parse -> Stringify returns canonical JSON for primitives.
    AssertEquals "true", Json_Stringify(Json_Parse("true")), "stringify true"
    AssertEquals "false", Json_Stringify(Json_Parse("false")), "stringify false"
    AssertEquals "null", Json_Stringify(Json_Parse("null")), "stringify null"
    AssertEquals """x""", Json_Stringify(Json_Parse("""x""")), "stringify string"
    AssertEquals "12.5", Json_Stringify(Json_Parse("12.5")), "stringify number"
End Sub

Private Sub Test_Stringify_ObjectAndArray()
    ' Goal: Stringify nested object/array and ensure stable output for known input.
    Dim v As Object
    Set v = Json_Parse("{""a"":1,""b"":[2,3],""c"":{""d"":4}}")
    AssertEquals "{""a"":1,""b"":[2,3],""c"":{""d"":4}}", Json_Stringify(v), "stringify nested object/array"
End Sub


' =============================================================================
' TESTS: Flatten
' =============================================================================

Private Sub Test_Flatten_Simple()
    ' Goal: Flatten simple object into "$.<key>" paths.
    Dim obj As Collection
    Set obj = Json_Parse("{""a"":1,""b"":true}")

    Dim flat As Collection
    Set flat = Json_Flatten(obj)

    AssertIsTaggedObject flat, "flatten returns tagged object"
    AssertEquals 2, Obj_CountPairs(flat), "flat pair count"
    AssertEquals 1, Flat_GetValue(flat, "$.a"), "flat $.a"
    AssertEquals True, Flat_GetValue(flat, "$.b"), "flat $.b"
End Sub

Private Sub Test_Flatten_Nested()
    ' Goal: Flatten nested object produces dotted paths.
    Dim obj As Collection
    Set obj = Json_Parse("{""a"":{""b"":2}}")

    Dim flat As Collection
    Set flat = Json_Flatten(obj)

    AssertEquals 1, Obj_CountPairs(flat), "flat pair count nested"
    AssertEquals 2, Flat_GetValue(flat, "$.a.b"), "flat $.a.b"
End Sub

Private Sub Test_Flatten_ArrayIndexed_Primitives()
    ' Goal: Flatten arrays produce index paths [0], [1], ...
    Dim obj As Collection
    Set obj = Json_Parse("{""x"":[1,2]}")

    Dim flat As Collection
    Set flat = Json_Flatten(obj)

    AssertEquals 2, Obj_CountPairs(flat), "flat pair count array primitives"
    AssertEquals 1, Flat_GetValue(flat, "$.x[0]"), "flat $.x[0]"
    AssertEquals 2, Flat_GetValue(flat, "$.x[1]"), "flat $.x[1]"
End Sub

Private Sub Test_Flatten_ArrayIndexed_Objects()
    ' Goal: Flatten arrays-of-objects produce indexed + dotted paths.
    Dim obj As Collection
    Set obj = Json_Parse("{""x"":[{""a"":1},{""a"":2}]}")

    Dim flat As Collection
    Set flat = Json_Flatten(obj)

    AssertEquals 2, Obj_CountPairs(flat), "flat pair count array objects"
    AssertEquals 1, Flat_GetValue(flat, "$.x[0].a"), "flat $.x[0].a"
    AssertEquals 2, Flat_GetValue(flat, "$.x[1].a"), "flat $.x[1].a"
End Sub

Private Sub Test_Flatten_KeyEscaping()
    ' Goal: Keys containing '.' and '\' are escaped in flattened path segments.
    Dim obj As Collection
    Set obj = Json_Parse("{""a.b"":1,""c\\d"":2}")

    Dim flat As Collection
    Set flat = Json_Flatten(obj)

    AssertEquals 1, Flat_GetValue(flat, "$.a\.b"), "dot escaped"
    AssertEquals 2, Flat_GetValue(flat, "$.c\\d"), "backslash escaped"
End Sub

Private Sub Test_Flatten_RootPrimitiveStoredAtDollar()
    ' Goal: Root primitive stored at "$".
    Dim v As Variant
    Json_ParseInto "5", v

    Dim flat As Collection
    Set flat = Json_Flatten(v)

    AssertIsTaggedObject flat, "flatten root primitive returns tagged object"
    AssertEquals 1, Obj_CountPairs(flat), "flat pair count root primitive"
    AssertEquals 5, Flat_GetValue(flat, "$"), "root primitive at $"
End Sub

Private Sub Test_Flatten_RootArrayIndexed()
    ' Goal: Root array stored at "$[n]" paths.
    Dim v As Variant
    Json_ParseInto "[1,2]", v

    Dim flat As Collection
    Set flat = Json_Flatten(v)

    AssertIsTaggedObject flat, "flatten root array returns tagged object"
    AssertEquals 2, Obj_CountPairs(flat), "flat pair count root array"
    AssertEquals 1, Flat_GetValue(flat, "$[0]"), "root [0]"
    AssertEquals 2, Flat_GetValue(flat, "$[1]"), "root [1]"
End Sub


' =============================================================================
' TESTS: Flat access helpers
' =============================================================================

Private Sub Test_FlatGet()
    ' Goal: FlatGet returns primitive at exact path.
    Dim obj As Collection
    Set obj = Json_Parse("{""a"":{""b"":5}}")

    Dim flat As Collection
    Set flat = Json_Flatten(obj)

    AssertEquals 5, Json_FlatGet(flat, "$.a.b"), "FlatGet nested"
End Sub

Private Sub Test_FlatContains()
    ' Goal: FlatContains returns True for present path, False otherwise.
    Dim obj As Collection
    Set obj = Json_Parse("{""a"":{""b"":5}}")

    Dim flat As Collection
    Set flat = Json_Flatten(obj)

    AssertTrue Json_FlatContains(flat, "$.a.b"), "contains existing path"
    AssertTrue Not Json_FlatContains(flat, "$.a.c"), "does not contain missing path"
End Sub


' =============================================================================
' TESTS: Unflatten
' =============================================================================

Private Sub Test_Unflatten_Simple()
    ' Goal: Flatten -> Unflatten reconstructs nested object structure (object-only paths).
    Dim obj As Collection
    Set obj = Json_Parse("{""a"":{""b"":5}}")

    Dim flat As Collection
    Set flat = Json_Flatten(obj)

    Dim rebuilt As Collection
    Set rebuilt = Json_Unflatten(flat)

    Dim inner As Object
    Set inner = Obj_GetObject(rebuilt, "a")

    AssertEquals 5, Obj_GetValue(inner, "b"), "unflatten nested"
End Sub

Private Sub Test_Unflatten_KeyEscaping_RoundTrip()
    ' Goal: Keys with '.' and '\' survive flatten/unflatten by escape rules.
    Dim original As Collection
    Set original = Json_Parse("{""a.b"":1,""c\\d"":2}")

    Dim flat As Collection
    Set flat = Json_Flatten(original)

    Dim rebuilt As Collection
    Set rebuilt = Json_Unflatten(flat)

    AssertIsTaggedObject rebuilt, "rebuilt is tagged object"
    AssertEquals 1, Obj_GetValue(rebuilt, "a.b"), "dot key survives"
    AssertEquals 2, Obj_GetValue(rebuilt, "c\d"), "backslash key survives"
End Sub

Private Sub Test_Unflatten_RejectsArrayPaths()
    ' Goal: Current Unflatten rejects any path containing array index segments.
    Dim v As Variant
    Json_ParseInto "{""x"":[1,2]}", v

    Dim flat As Collection
    Set flat = Json_Flatten(v)

    On Error GoTo expected
    Call Json_Unflatten(flat)
    Err.Raise vbObjectError + 699, "mJsonTests", "Expected Unflatten to reject array paths"
expected:
    AssertTrue (Err.Number <> 0), "Unflatten should raise on array paths"
    Err.Clear
End Sub


' =============================================================================
' TESTS: Array-of-object root detection
' =============================================================================

Private Sub Test_FindArrayObjectRoots()
    ' Goal: Detect roots that are arrays-of-objects by path pattern: <root>[n].<prop>
    Dim obj As Collection
    Set obj = Json_Parse("{""x"":[1,2],""orders"":[{""id"":1,""items"":[{""sku"":""A""},{""sku"":""B""}]},{""id"":2,""items"":[]}]}")

    Dim flat As Collection
    Set flat = Json_Flatten(obj)

    Dim roots As Collection
    Set roots = Json_FindArrayObjectRoots(flat)

    AssertCollectionContainsString roots, "$.orders", "detect root array-of-objects"
    AssertCollectionContainsString roots, "$.orders.items", "detect nested array-of-objects"

    Dim i As Long
    For i = 1 To roots.count
        AssertTrue (CStr(roots(i)) <> "$.x"), "should not include primitive array root $.x"
    Next i
End Sub

Private Sub Test_FindArrayObjectRoots_RootArray()
    ' Goal: Root array-of-objects should be detected as "$".
    Dim v As Variant
    Json_ParseInto "[{""a"":1},{""a"":2}]", v

    Dim flat As Collection
    Set flat = Json_Flatten(v)

    Dim roots As Collection
    Set roots = Json_FindArrayObjectRoots(flat)

    AssertCollectionContainsString roots, "$", "detect root array-of-objects at $"
End Sub


' =============================================================================
' TESTS: ExtractTableRows + TableTo2D
' =============================================================================

Private Sub Test_ExtractTableRows_Simple()
    ' Goal: Extract rows from "$.orders" and exclude nested array columns from the parent table.
    ' Input: orders is array-of-objects; each order contains customer object and items array.
    ' Expect:
    '   - 2 rows extracted
    '   - columns include "id", "customer.name"
    '   - columns containing "[" are excluded (items[0].sku not present)
    Dim obj As Collection
    Set obj = Json_Parse("{""orders"":[{""id"":1,""customer"":{""name"":""A""}, ""items"":[{""sku"":""X""}]},{""id"":2,""customer"":{""name"":""B""}, ""items"":[]}]}")

    Dim flat As Collection
    Set flat = Json_Flatten(obj)

    Dim rows As Collection
    Set rows = Json_ExtractTableRows(flat, "$.orders")

    AssertEquals 2, rows.count, "rows count"

    Dim r0 As Collection
    Set r0 = rows(1)
    AssertIsTaggedObject r0, "row0 is tagged object"
    AssertEquals 1, Obj_GetValue(r0, "id"), "row0 id"
    AssertEquals "A", Obj_GetValue(r0, "customer.name"), "row0 customer.name"

    Dim r1 As Collection
    Set r1 = rows(2)
    AssertEquals 2, Obj_GetValue(r1, "id"), "row1 id"
    AssertEquals "B", Obj_GetValue(r1, "customer.name"), "row1 customer.name"

    On Error GoTo ExpectedMissing
    Call Obj_GetValue(r0, "items[0].sku")
    Err.Raise vbObjectError + 699, "mJsonTests", "Expected missing nested array column"
ExpectedMissing:
    Err.Clear
End Sub

Private Sub Test_ExtractTableRows_RootArray()
    ' Goal: Root array-of-objects extraction uses tableRoot="$".
    Dim v As Variant
    Json_ParseInto "[{""a"":1},{""a"":2}]", v

    Dim flat As Collection
    Set flat = Json_Flatten(v)

    Dim rows As Collection
    Set rows = Json_ExtractTableRows(flat, "$")

    AssertEquals 2, rows.count, "root rows count"
    AssertEquals 1, Obj_GetValue(rows(1), "a"), "root row0 a"
    AssertEquals 2, Obj_GetValue(rows(2), "a"), "root row1 a"
End Sub

Private Sub Test_TableTo2D_EmptyRows()
    ' Goal: TableTo2D returns default header "value" and Empty data when rows are empty.
    Dim rows As New Collection
    Dim headers As Variant
    Dim data As Variant

    data = Json_TableTo2D(rows, headers)

    AssertEquals "value", CStr(headers(1)), "default header for empty rows"
    AssertTrue IsEmpty(data), "data should be Empty when rows are empty"
End Sub

Private Sub Test_TableTo2D_ColumnPlacement_VariedOrder()
    ' Goal: Header order is first-seen across rows; data lands in correct columns.
    Dim rows As New Collection

    Dim r0 As New Collection
    r0.Add "__OBJ__"
    Dim p0a(0 To 1) As Variant: p0a(0) = "b": p0a(1) = 2
    Dim p0b(0 To 1) As Variant: p0b(0) = "a": p0b(1) = 1
    r0.Add p0a
    r0.Add p0b
    rows.Add r0

    Dim r1 As New Collection
    r1.Add "__OBJ__"
    Dim p1a(0 To 1) As Variant: p1a(0) = "a": p1a(1) = 10
    r1.Add p1a
    rows.Add r1

    Dim headers As Variant
    Dim data As Variant
    data = Json_TableTo2D(rows, headers)

    AssertEquals "b", CStr(headers(1)), "header1"
    AssertEquals "a", CStr(headers(2)), "header2"

    AssertEquals 2, data(1, 1), "r0 b"
    AssertEquals 1, data(1, 2), "r0 a"
    AssertTrue IsEmpty(data(2, 1)), "r1 b missing"
    AssertEquals 10, data(2, 2), "r1 a"
End Sub

Private Sub Test_TableTo2D_Basic()
    ' Goal: TableTo2D produces headers and body for extracted rows with nested object columns.
    Dim obj As Collection
    Set obj = Json_Parse("{""orders"":[{""id"":1,""customer"":{""name"":""A""}},{""id"":2}]}")

    Dim flat As Collection
    Set flat = Json_Flatten(obj)

    Dim rows As Collection
    Set rows = Json_ExtractTableRows(flat, "$.orders")

    Dim headers As Variant
    Dim data As Variant
    data = Json_TableTo2D(rows, headers)

    Dim expHeaders As Variant
    expHeaders = Array("id", "customer.name")

    Assert1DArrayEquals expHeaders, headers, "headers"
    AssertEquals 2, UBound(data, 1), "row count"
    AssertEquals 2, UBound(data, 2), "col count"

    AssertEquals 1, data(1, 1), "r1 id"
    AssertEquals "A", data(1, 2), "r1 customer.name"
    AssertEquals 2, data(2, 1), "r2 id"
    AssertTrue IsEmpty(data(2, 2)), "r2 missing customer.name"
End Sub


' =============================================================================
' TESTS: ObjSet
' =============================================================================

Private Sub Test_ObjSet_PreservesOrder_OnOverwrite()
    ' Goal: Overwriting an existing key keeps original key order.
    Dim o As New Collection
    o.Add "__OBJ__"

    Dim p1(0 To 1) As Variant: p1(0) = "a": p1(1) = 1
    Dim p2(0 To 1) As Variant: p2(0) = "b": p2(1) = 2
    Dim p3(0 To 1) As Variant: p3(0) = "c": p3(1) = 3

    o.Add p1
    o.Add p2
    o.Add p3

    Json_ObjSet o, "b", 99

    Dim kA As String, kB As String, kC As String
    kA = CStr(o(2)(0))
    kB = CStr(o(3)(0))
    kC = CStr(o(4)(0))

    AssertEquals "a", kA, "order key 1"
    AssertEquals "b", kB, "order key 2"
    AssertEquals "c", kC, "order key 3"

    AssertEquals 99, Obj_GetValue(o, "b"), "overwrite value"
End Sub

Private Function Obj_IsObject(ByVal v As Variant) As Boolean
    Obj_IsObject = (IsObject(v) And TypeName(v) = "Collection" And v.count >= 1 And VarType(v(1)) = vbString And v(1) = "__OBJ__")
End Function

Private Function Obj_CountPairs(ByVal obj As Collection) As Long
    Obj_CountPairs = obj.count - 1
End Function

Private Function Obj_GetValue(ByVal obj As Collection, ByVal key As String) As Variant
    Dim i As Long
    For i = 2 To obj.count
        Dim pair As Variant
        pair = obj(i)
        If CStr(pair(0)) = key Then
            If IsObject(pair(1)) Then
                Err.Raise vbObjectError + 651, "mJsonTests", "Key is an object (use Obj_GetObject): " & key
            End If
            Obj_GetValue = pair(1)
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 650, "mJsonTests", "Key not found: " & key
End Function

Private Function Obj_GetObject(ByVal obj As Collection, ByVal key As String) As Object
    Dim i As Long
    For i = 2 To obj.count
        Dim pair As Variant
        pair = obj(i)
        If CStr(pair(0)) = key Then
            If Not IsObject(pair(1)) Then
                Err.Raise vbObjectError + 652, "mJsonTests", "Key is not an object (use Obj_GetValue): " & key
            End If
            Set Obj_GetObject = pair(1)
            Exit Function
        End If
    Next i
    Err.Raise vbObjectError + 650, "mJsonTests", "Key not found: " & key
End Function

Private Function Flat_GetValue(ByVal flatObj As Collection, ByVal key As String) As Variant
    Flat_GetValue = Obj_GetValue(flatObj, key)
End Function

Private Function Flat_GetObject(ByVal flatObj As Collection, ByVal key As String) As Object
    Set Flat_GetObject = Obj_GetObject(flatObj, key)
End Function

' =============================================================================
' TESTS: Unicode / Emoji (Surrogate pairs)
' =============================================================================

Private Sub Test_Parse_UnicodeEscape_SurrogatePair_Emoji()
    ' Goal: Surrogate pairs decode into a VBA string (2 UTF-16 code units) and round-trip safely.
    ' Input: JSON string containing "\uD83D\uDE00" (??)
    ' Expect:
    '   - Parsed string Len=2 (UTF-16 surrogate pair)
    '   - Stringify returns same 2 code units wrapped in quotes (not \u escapes)
    Dim s As String
    s = CStr(Json_Parse("""\uD83D\uDE00"""))

    AssertEquals 2, Len(s), "emoji should be 2 UTF-16 code units in VBA"

    Dim expected As String
    expected = """" & ChrW$(&HD83D) & ChrW$(&HDE00) & """"

    AssertEquals expected, Json_Stringify(s), "stringify preserves surrogate pair code units"
End Sub

Private Sub Test_Stringify_DoesNotEscapeNegativeAscW()
    ' Goal: Json_EscapeString must NOT treat negative AscW values as control chars.
    ' Input: A string containing a surrogate code unit (ChrW(&HD83D))
    ' Expect:
    '   - Stringify includes that code unit verbatim, not as \u???? escape.
    Dim s As String
    s = ChrW$(&HD83D) & "X"  ' high surrogate + "X"

    Dim out As String
    out = Json_Stringify(s)

    Dim expected As String
    expected = """" & ChrW$(&HD83D) & "X" & """"

    AssertEquals expected, out, "negative AscW must not trigger \\u escaping"
End Sub

Private Sub Test_Parse_UnicodeEscape4_BasicBMP()
    ' Goal: \uXXXX decodes BMP character correctly.
    ' Input: "\u00E9" (é)
    ' Expect: parsed string equals ChrW(&H00E9)
    Dim s As String
    s = CStr(Json_Parse("""\u00E9"""))

    AssertEquals ChrW$(&HE9), s, "BMP unicode escape decoded"
End Sub

Private Sub Test_Parse_UnicodeEscape_SurrogatePair_RoundTrip_WithText()
    ' Goal: Emoji within surrounding text survives Parse -> Stringify -> Parse.
    ' Input: "A\uD83D\uDE00B"
    ' Expect:
    '   - After round-trip, string is identical (by code units)
    Dim s1 As String, s2 As String
    s1 = CStr(Json_Parse("""A\uD83D\uDE00B"""))

    Dim json2 As String
    json2 = Json_Stringify(s1)

    s2 = CStr(Json_Parse(json2))

    AssertEquals s1, s2, "round-trip preserves emoji inside text"
    AssertEquals 4, Len(s1), "A + surrogate pair + B => 4 UTF-16 code units"
End Sub

Private Sub Test_Parse_String_Long()
    ' Goal: Parser handles large strings without quadratic slowdown or corruption.
    ' Input: JSON string with 10,000 'a' characters.
    ' Expect: Exact same string returned.
    Dim s As String
    s = String$(10000, "a")

    Dim jsonText As String
    jsonText = """" & s & """"

    Dim out As Variant
    out = Json_Parse(jsonText)

    AssertEquals s, CStr(out), "parse long string 10k"
End Sub

Private Sub Test_Parse_String_Long_WithEscapes()
    ' Goal: Large string + escapes still parse correctly.
    ' Input: long base + embedded escapes \n \t \" \\ and a unicode emoji.
    ' Expect: Exact reconstructed string.
    Dim base As String
    base = String$(4000, "x")

    Dim expected As String
    expected = base & vbLf & vbTab & """" & "\" & " " & ChrW$(&HD83D) & ChrW$(&HDE00)

    ' Build JSON:
    '   - \n, \t, \", \\ are literal escapes
    '   - emoji ?? as surrogate pair \uD83D\uDE00
    Dim jsonText As String
    jsonText = """" & base & "\n\t" & "\""" & "\\" & " " & "\uD83D\uDE00" & """"

    Dim out As Variant
    out = Json_Parse(jsonText)

    AssertEquals expected, CStr(out), "parse long string with escapes + emoji"
End Sub

Private Sub Test_Parse_UnicodeEscape_InvalidHighWithoutLow()
    ' Goal: High surrogate not followed by valid low surrogate must error.
    On Error GoTo expected
    Call Json_Parse("""\uD83D""")
    Err.Raise vbObjectError + 699, "mJsonTests", "Expected error for lone high surrogate"
expected:
    AssertTrue Err.Number <> 0, "lone high surrogate should error"
    Err.Clear
End Sub

Private Sub Test_Parse_UnicodeEscape_InvalidLowWithoutHigh()
    ' Goal: Low surrogate without preceding high surrogate must error.
    On Error GoTo expected
    Call Json_Parse("""\uDE00""")
    Err.Raise vbObjectError + 699, "mJsonTests", "Expected error for lone low surrogate"
expected:
    AssertTrue Err.Number <> 0, "lone low surrogate should error"
    Err.Clear
End Sub

Private Sub Test_Parse_Number_RejectLeadingZero()
    ' Goal: JSON does not allow leading zero before additional digits.
    On Error GoTo expected
    Call Json_Parse("0123")
    Err.Raise vbObjectError + 699, "mJsonTests", "Expected error for leading zero"
expected:
    AssertTrue Err.Number <> 0, "leading zero should error"
    Err.Clear
End Sub

Private Sub Test_Parse_TrailingCharacters()
    ' Goal: Parser rejects valid JSON followed by extra characters.
    On Error GoTo expected
    Call Json_Parse("{""a"":1} x")
    Err.Raise vbObjectError + 699, "mJsonTests", "Expected trailing character error"
expected:
    AssertTrue Err.Number <> 0, "trailing characters should error"
    Err.Clear
End Sub

Private Sub Test_Parse_Number_Exponent_Negative()
    AssertEquals 0.0012, Json_Parse("1.2e-3"), "negative exponent"
End Sub

Private Sub Test_Stringify_LargeArray_RoundTrip()
    ' Goal: Large array stringify does not corrupt output; round-trip preserves values.
    ' Input: JSON array [1..2000] built as Collection.
    ' Expect:
    '   - Stringify -> Parse returns Collection count=2000
    '   - spot-check first/middle/last values

    Dim arr As New Collection
    Dim i As Long
    For i = 1 To 2000
        arr.Add CLng(i)
    Next i

    Dim jsonText As String
    jsonText = Json_Stringify(arr)

    Dim parsed As Variant
    Set parsed = Json_Parse(jsonText)

    AssertIsCollection parsed, "parsed large array is collection"
    AssertEquals 2000, parsed.count, "large array count"
    AssertEquals 1, parsed(1), "large array first"
    AssertEquals 1000, parsed(1000), "large array middle"
    AssertEquals 2000, parsed(2000), "large array last"
End Sub

Private Sub Test_Stringify_LargeObject_RoundTrip()
    ' Goal: Large object stringify does not drop keys; round-trip preserves values.
    ' Input: Tagged object with 500 pairs: k1..k500 => 1..500
    ' Expect:
    '   - Parse(Stringify(obj)) returns tagged object with 500 pairs
    '   - spot-check keys

    Dim o As New Collection
    o.Add "__OBJ__"

    Dim i As Long
    For i = 1 To 500
        Json_ObjSet o, "k" & CStr(i), CLng(i)
    Next i

    Dim jsonText As String
    jsonText = Json_Stringify(o)

    Dim parsedObj As Collection
    Set parsedObj = Json_Parse(jsonText)

    AssertIsTaggedObject parsedObj, "parsed large object is tagged object"
    AssertEquals 500, Obj_CountPairs(parsedObj), "large object pair count"
    AssertEquals 1, Obj_GetValue(parsedObj, "k1"), "k1"
    AssertEquals 250, Obj_GetValue(parsedObj, "k250"), "k250"
    AssertEquals 500, Obj_GetValue(parsedObj, "k500"), "k500"
End Sub

' =============================================================================
' TESTS: TableTo2D (hash-table stress)
' =============================================================================

Private Sub Test_TableTo2D_WideRow_ManyHeaders()
    ' Goal: Stress Json_TableTo2D with many distinct headers in a single row.
    ' Input: 1 row object with 350 keys: k001..k350
    ' Expect:
    '   - headers count = 350
    '   - data(1, col(k001))=1, data(1, col(k175))=175, data(1, col(k350))=350
    ' Notes:
    '   This covers hash table load/probing under "few rows, very wide schema".

    Dim rows As New Collection
    Dim r0 As New Collection
    r0.Add "__OBJ__"

    Dim i As Long
    For i = 1 To 350
        Dim p(0 To 1) As Variant
        p(0) = "k" & Right$("000" & CStr(i), 3)
        p(1) = CLng(i)
        r0.Add p
    Next i

    rows.Add r0

    Dim headers As Variant
    Dim data As Variant
    data = Json_TableTo2D(rows, headers)

    ' Header count
    AssertEquals 350, (UBound(headers) - LBound(headers) + 1), "wide row header count"
    AssertEquals 1, UBound(data, 1), "wide row data row count"
    AssertEquals 350, UBound(data, 2), "wide row data col count"

    ' Spot-check header order (first-seen)
    AssertEquals "k001", CStr(headers(1)), "wide row header(1)"
    AssertEquals "k175", CStr(headers(175)), "wide row header(175)"
    AssertEquals "k350", CStr(headers(350)), "wide row header(350)"

    ' Spot-check values align
    AssertEquals 1, data(1, 1), "wide row value k001"
    AssertEquals 175, data(1, 175), "wide row value k175"
    AssertEquals 350, data(1, 350), "wide row value k350"
End Sub

Private Sub Test_TableTo2D_WideRows_VariedPresence()
    ' Goal: Ensure sparse presence across rows still places values correctly.
    ' Input:
    '   Row1: a=1, b=2
    '   Row2: b=20, c=30
    '   Row3: c=300, a=100
    ' Expect:
    '   - Headers first-seen order: a, b, c
    '   - Missing cells remain Empty

    Dim rows As New Collection

    Dim r1 As New Collection: r1.Add "__OBJ__"
    Dim p1a(0 To 1) As Variant: p1a(0) = "a": p1a(1) = 1
    Dim p1b(0 To 1) As Variant: p1b(0) = "b": p1b(1) = 2
    r1.Add p1a: r1.Add p1b
    rows.Add r1

    Dim r2 As New Collection: r2.Add "__OBJ__"
    Dim p2b(0 To 1) As Variant: p2b(0) = "b": p2b(1) = 20
    Dim p2c(0 To 1) As Variant: p2c(0) = "c": p2c(1) = 30
    r2.Add p2b: r2.Add p2c
    rows.Add r2

    Dim r3 As New Collection: r3.Add "__OBJ__"
    Dim p3c(0 To 1) As Variant: p3c(0) = "c": p3c(1) = 300
    Dim p3a(0 To 1) As Variant: p3a(0) = "a": p3a(1) = 100
    r3.Add p3c: r3.Add p3a
    rows.Add r3

    Dim headers As Variant
    Dim data As Variant
    data = Json_TableTo2D(rows, headers)

    AssertEquals 3, (UBound(headers) - LBound(headers) + 1), "sparse header count"
    AssertEquals "a", CStr(headers(1)), "header1"
    AssertEquals "b", CStr(headers(2)), "header2"
    AssertEquals "c", CStr(headers(3)), "header3"

    AssertEquals 1, data(1, 1), "r1 a"
    AssertEquals 2, data(1, 2), "r1 b"
    AssertTrue IsEmpty(data(1, 3)), "r1 c missing"

    AssertTrue IsEmpty(data(2, 1)), "r2 a missing"
    AssertEquals 20, data(2, 2), "r2 b"
    AssertEquals 30, data(2, 3), "r2 c"

    AssertEquals 100, data(3, 1), "r3 a"
    AssertTrue IsEmpty(data(3, 2)), "r3 b missing"
    AssertEquals 300, data(3, 3), "r3 c"
End Sub

Public Sub Test_TableTo2D_WideSchema_DoesNotHang_AndIsComplete()
    Dim rows As New Collection

    Dim o As New Collection
    o.Add "__OBJ__"

    Dim i As Long
    For i = 1 To 512
        Dim p(0 To 1) As Variant
        p(0) = "k" & CStr(i)
        p(1) = i
        o.Add p
    Next i

    rows.Add o

    Dim headersOut As Variant
    Dim data2D As Variant
    data2D = Json_TableTo2D(rows, headersOut)

    AssertEquals 512, (UBound(headersOut) - LBound(headersOut) + 1), "wide schema header count"
    AssertEquals 1, (UBound(data2D, 1) - LBound(data2D, 1) + 1), "wide schema row count"
    AssertEquals 512, (UBound(data2D, 2) - LBound(data2D, 2) + 1), "wide schema col count"
    AssertEquals 1, CLng(data2D(1, 1)), "wide schema first cell value"
    AssertEquals 512, CLng(data2D(1, 512)), "wide schema last cell value"
End Sub

Public Sub Test_TableTo2D_HeaderOrder_FirstSeenAcrossRows()
    Dim rows As New Collection

    Dim r1 As New Collection: r1.Add "__OBJ__"
    Dim p1(0 To 1) As Variant: p1(0) = "b": p1(1) = 2: r1.Add p1
    Dim p2(0 To 1) As Variant: p2(0) = "a": p2(1) = 1: r1.Add p2
    rows.Add r1

    Dim r2 As New Collection: r2.Add "__OBJ__"
    Dim p3(0 To 1) As Variant: p3(0) = "c": p3(1) = 3: r2.Add p3
    Dim p4(0 To 1) As Variant: p4(0) = "a": p4(1) = 10: r2.Add p4
    rows.Add r2

    Dim headersOut As Variant
    Dim data2D As Variant
    data2D = Json_TableTo2D(rows, headersOut)

    AssertEquals "b", CStr(headersOut(1)), "header order (1)"
    AssertEquals "a", CStr(headersOut(2)), "header order (2)"
    AssertEquals "c", CStr(headersOut(3)), "header order (3)"

    AssertEquals 2, CLng(data2D(1, 1)), "row1 b"
    AssertEquals 1, CLng(data2D(1, 2)), "row1 a"
    AssertTrue IsEmpty(data2D(1, 3)), "row1 c empty"

    AssertTrue IsEmpty(data2D(2, 1)), "row2 b empty"
    AssertEquals 10, CLng(data2D(2, 2)), "row2 a"
    AssertEquals 3, CLng(data2D(2, 3)), "row2 c"
End Sub

Public Sub Test_Flatten_PathEscapes_DotAndBackslash()
    Dim jsonText As String
    jsonText = "{""a.b"":{""c\\d"":1}}"

    Dim parsed As Variant
    Set parsed = Json_Parse(jsonText)

    Dim flat As Collection
    Set flat = Json_Flatten(parsed)

    AssertTrue Json_FlatContains(flat, "$.a\.b.c\\d"), "flatten path escape should exist"
    AssertEquals 1, CLng(Json_FlatGet(flat, "$.a\.b.c\\d")), "flatten path escape value"
End Sub

Public Sub Test_Unflatten_RoundTrip_DotAndBackslash()
    Dim jsonText As String
    jsonText = "{""a.b"":{""c\\d"":1}}"

    Dim parsed As Variant
    Set parsed = Json_Parse(jsonText)

    Dim flat As Collection
    Set flat = Json_Flatten(parsed)

    Dim rebuilt As Collection
    Set rebuilt = Json_Unflatten(flat)

    Dim outJson As String
    outJson = Json_Stringify(rebuilt)

    Dim reparsed As Variant
    Set reparsed = Json_Parse(outJson)

    Dim flat2 As Collection
    Set flat2 = Json_Flatten(reparsed)

    AssertTrue Json_FlatContains(flat2, "$.a\.b.c\\d"), "roundtrip flatten contains escaped path"
    AssertEquals 1, CLng(Json_FlatGet(flat2, "$.a\.b.c\\d")), "roundtrip value"
End Sub

Public Sub Test_FindArrayObjectRoots_DoesNotFalsePositive_OnArrayOfPrimitives()
    Dim jsonText As String
    jsonText = "{""x"":[1,2,3],""y"":[{""a"":1},{""a"":2}]}"

    Dim parsed As Variant
    Set parsed = Json_Parse(jsonText)

    Dim flat As Collection
    Set flat = Json_Flatten(parsed)

    Dim roots As Collection
    Set roots = Json_FindArrayObjectRoots(flat)

    Dim i As Long
    Dim hasY As Boolean, hasX As Boolean

    For i = 1 To roots.count
        If CStr(roots(i)) = "$.y" Then hasY = True
        If CStr(roots(i)) = "$.x" Then hasX = True
    Next i

    AssertTrue hasY, "should detect $.y (array-of-objects)"
    AssertTrue (Not hasX), "should NOT detect $.x (array-of-primitives)"
End Sub

Public Sub Test_ExtractTableRows_ExcludesNestedArrays_ChildTablesOnly()
    Dim jsonText As String
    jsonText = "{""orders"":[{""id"":1,""items"":[{""sku"":""A""}]}]}"

    Dim parsed As Variant
    Set parsed = Json_Parse(jsonText)

    Dim flat As Collection
    Set flat = Json_Flatten(parsed)

    Dim rows As Collection
    Set rows = Json_ExtractTableRows(flat, "$.orders")

    Dim headersOut As Variant
    Dim data2D As Variant
    data2D = Json_TableTo2D(rows, headersOut)

    ' Must include id; must NOT include items[0].sku because it contains "["
    Dim i As Long
    Dim hasId As Boolean, hasItems As Boolean
    For i = LBound(headersOut) To UBound(headersOut)
        If CStr(headersOut(i)) = "id" Then hasId = True
        If InStr(1, CStr(headersOut(i)), "items", vbBinaryCompare) > 0 Then hasItems = True
    Next i

    AssertTrue hasId, "orders table should include id"
    AssertTrue (Not hasItems), "orders table should not include nested array columns"
    AssertEquals 1, CLng(data2D(1, 1)), "id value"
End Sub

Public Sub Test_ObjSet_Overwrite_PreservesPosition()
    Dim o As New Collection
    o.Add "__OBJ__"

    Dim p1(0 To 1) As Variant: p1(0) = "a": p1(1) = 1: o.Add p1
    Dim p2(0 To 1) As Variant: p2(0) = "b": p2(1) = 2: o.Add p2
    Dim p3(0 To 1) As Variant: p3(0) = "c": p3(1) = 3: o.Add p3

    Json_ObjSet o, "b", 20

    ' ensure a,b,c order preserved and b updated
    AssertEquals "a", CStr(o(2)(0)), "objset order a"
    AssertEquals "b", CStr(o(3)(0)), "objset order b"
    AssertEquals "c", CStr(o(4)(0)), "objset order c"
    AssertEquals 20, CLng(o(3)(1)), "objset value updated"
End Sub

Private Sub Test_TableTo2D_RowsExist_ButNoKeys()
    ' Goal: If rows exist but each row object has no pairs, TableTo2D should not crash.
    ' Input: 2 row objects: { } and { } (represented as tagged row objects with only "__OBJ__")
    ' Expect:
    '   - headers defaults to ["value"] OR some defined behavior
    '   - data is 2x1 (or Empty if you choose), but must not raise

    Dim rows As New Collection

    Dim r1 As New Collection: r1.Add "__OBJ__"
    Dim r2 As New Collection: r2.Add "__OBJ__"

    rows.Add r1
    rows.Add r2

    Dim headers As Variant
    Dim data As Variant

    On Error GoTo Fail
    data = Json_TableTo2D(rows, headers)
    On Error GoTo 0

    ' Pick the behavior you want as contract.
    ' If you choose "value" default:
    AssertEquals "value", CStr(headers(1)), "headers default for keyless rows"
    AssertEquals 2, (UBound(data, 1) - LBound(data, 1) + 1), "row count"
    AssertEquals 1, (UBound(data, 2) - LBound(data, 2) + 1), "col count"
    AssertTrue IsEmpty(data(1, 1)), "row1 empty"
    AssertTrue IsEmpty(data(2, 1)), "row2 empty"

    Exit Sub

Fail:
    Err.Raise vbObjectError + 699, "mJsonTests", "Json_TableTo2D crashed on keyless rows: " & Err.Description
End Sub

Private Sub Test_Parse_Primitive_Root_Set_ShouldFail_AndIsCaught()
    ' Goal: Prove that "Set o = Json_Parse(<primitive>)" raises a runtime error,
    '       and also clarify why you see no popup (because the test catches it).
    '
    ' Expect:
    '   - Err.Number <> 0 inside expected:
    '   - Typically Err.Number = 91 ("Object variable or With block variable not set")
    '
    On Error GoTo expected

    Dim o As Object
    Set o = Json_Parse("null")  ' <-- should error at Set assignment

    Err.Raise vbObjectError + 699, "mJsonTests", "Expected Set on primitive to fail (should not reach here)"
expected:
    AssertTrue Err.Number <> 0, "Set on primitive should error"
    ' Optional: if you want to lock it down harder, uncomment:
    ' AssertEquals 91, Err.Number, "Expected runtime error 91 for Set on primitive"
    Err.Clear
End Sub


Private Sub Test_Excel_UpsertListObjectFromJsonAtRoot_PrimitiveRoot_Raises1130()
    ' Contract:
    '   Excel_UpsertListObjectFromJsonAtRoot requires JSON root
    '   to be an object or array (Collection).
    '   Primitive roots (e.g. "null") must raise vbObjectError + 1130.

    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets.Add

    On Error GoTo ExpectedError

    ' --- Act ---
    Excel_UpsertListObjectFromJsonAtRoot _
        ws, "tX", ws.Range("A1"), _
        "null", "$", _
        True, True, False

    ' If we reach here, no error occurred (FAIL)
    Err.Raise vbObjectError + 699, "mJsonTests", _
        "Expected error 1130 for primitive root, but no error was raised."

ExpectedError:
    ' --- Assert ---
    AssertEquals vbObjectError + 1130, Err.Number, _
        "Excel_UpsertListObjectFromJsonAtRoot primitive root should raise 1130"
    AssertEquals "Excel_UpsertListObjectFromJsonAtRoot", Err.Source, _
        "Error source should be Excel_UpsertListObjectFromJsonAtRoot"

    Err.Clear

    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
End Sub

Private Sub Test_ParseInto_Primitive_DoesNotBecomeObject()
    Dim v As Variant
    Json_ParseInto "null", v
    AssertTrue Not IsObject(v), "ParseInto primitive must not be object"
    AssertTrue IsNull(v), "ParseInto null must be Null"
End Sub

Private Sub Test_ParseInto_Object_IsObject()
    Dim v As Variant
    Json_ParseInto "{""a"":1}", v
    AssertTrue IsObject(v), "ParseInto object must be object"
    AssertEquals "Collection", TypeName(v), "ParseInto object type"
    AssertTrue Obj_IsObject(v), "ParseInto object must be tagged object"
End Sub

Private Sub Test_ObjGetObject_ReturnsObject_AndCanBeReSet()
    Dim obj As Collection
    Set obj = Json_Parse("{""x"":{""a"":1}}")

    Dim child As Object
    Set child = Obj_GetObject(obj, "x")

    Dim o2 As New Collection
    o2.Add "__OBJ__"
    Json_ObjSet o2, "child", child

    Dim out As String
    out = Json_Stringify(o2)

    AssertEquals "{""child"":{""a"":1}}", out, "object reference preserved through Variant pair"
End Sub

Private Sub Test_TableTo2D_Null_WritesAsNullVariant()
    Dim rows As New Collection
    Dim r1 As New Collection: r1.Add "__OBJ__"
    Dim p(0 To 1) As Variant: p(0) = "a": p(1) = Null
    r1.Add p
    rows.Add r1

    Dim headers As Variant, data As Variant
    data = Json_TableTo2D(rows, headers)

    AssertEquals "a", CStr(headers(1)), "header"
    AssertTrue IsNull(data(1, 1)), "cell should be Null (decide contract)"
End Sub

Private Sub Test_ExtractTableRows_NestedChildRows_DoNotCollideAcrossParents()
    ' Goal:
    '   Prove nested child table extraction does NOT collide when each parent has items[0].
    '
    ' Input:
    '   orders[0].items[0].sku = "A0"
    '   orders[1].items[0].sku = "B0"
    '
    ' Expected:
    '   Extract rows for "$.orders.items" returns 2 rows:
    '     row1 sku = "A0"
    '     row2 sku = "B0"
    '
    Dim jsonText As String
    jsonText = "{""orders"":[{""id"":1,""items"":[{""sku"":""A0"",""qty"":1}]},{""id"":2,""items"":[{""sku"":""B0"",""qty"":2}]}]}"

    Dim parsed As Variant
    Set parsed = Json_Parse(jsonText)

    Dim flat As Collection
    Set flat = Json_Flatten(parsed)

    Dim rows As Collection
    Set rows = Json_ExtractTableRows(flat, "$.orders.items")

    AssertEquals 2, rows.count, "nested child rows count (no collision across parents)"

    AssertEquals "A0", Obj_GetValue(rows(1), "sku"), "row1 sku from orders[0].items[0]"
    AssertEquals 1, CLng(Obj_GetValue(rows(1), "qty")), "row1 qty"

    AssertEquals "B0", Obj_GetValue(rows(2), "sku"), "row2 sku from orders[1].items[0]"
    AssertEquals 2, CLng(Obj_GetValue(rows(2), "qty")), "row2 qty"
End Sub

Private Sub Test_ExtractTableRows_NestedChildRows_OrderIsFirstSeenPathOrder()
    ' Goal:
    '   Verify deterministic ordering remains "first time a rowKey is seen".
    '
    ' Input:
    '   orders[0].items => [A0, A1]
    '   orders[1].items => [B0]
    '
    ' Expected order:
    '   A0, A1, B0
    '
    Dim jsonText As String
    jsonText = "{""orders"":[{""id"":1,""items"":[{""sku"":""A0""},{""sku"":""A1""}]},{""id"":2,""items"":[{""sku"":""B0""}]}]}"

    Dim parsed As Variant
    Set parsed = Json_Parse(jsonText)

    Dim flat As Collection
    Set flat = Json_Flatten(parsed)

    Dim rows As Collection
    Set rows = Json_ExtractTableRows(flat, "$.orders.items")

    AssertEquals 3, rows.count, "row count"
    AssertEquals "A0", Obj_GetValue(rows(1), "sku"), "order row1"
    AssertEquals "A1", Obj_GetValue(rows(2), "sku"), "order row2"
    AssertEquals "B0", Obj_GetValue(rows(3), "sku"), "order row3"
End Sub

Private Sub Test_ExtractTableRows_NestedChildRows_AllowsObjectColumns_ButExcludesNestedArrays()
    ' Goal:
    '   Child rows should include dotted object columns (e.g. product.name)
    '   but still exclude any nested arrays under the child row (e.g. tags[0]).
    '
    ' Input:
    '   orders.items includes:
    '     sku
    '     product.name
    '     tags = ["x","y"]   (nested array under child row)
    '
    ' Expected:
    '   headers include "sku" and "product.name"
    '   does NOT include "tags[0]" etc
    '
    Dim jsonText As String
    jsonText = "{""orders"":[{""items"":[{""sku"":""A0"",""product"":{""name"":""Widget""},""tags"":[""x"",""y""]}]}]}"

    Dim parsed As Variant
    Set parsed = Json_Parse(jsonText)

    Dim flat As Collection
    Set flat = Json_Flatten(parsed)

    Dim rows As Collection
    Set rows = Json_ExtractTableRows(flat, "$.orders.items")

    AssertEquals 1, rows.count, "row count"
    AssertEquals "A0", Obj_GetValue(rows(1), "sku"), "sku"
    AssertEquals "Widget", Obj_GetValue(rows(1), "product.name"), "product.name"

    On Error GoTo ExpectedMissing
    Call Obj_GetValue(rows(1), "tags[0]")
    Err.Raise vbObjectError + 699, "mJsonTests", "Expected tags[0] to be excluded from child rows"
ExpectedMissing:
    Err.Clear
End Sub

Private Sub Test_ExtractTableRows_NestedChildRows_TableTo2D_ProducesExpectedHeaders()
    ' Goal:
    '   End-to-end: Extract nested child rows -> TableTo2D produces stable headers and correct cells.
    '
    ' Expected:
    '   headers: ["sku","qty"] (first-seen)
    '   data:
    '     A0, 1
    '     B0, 2
    '
    Dim jsonText As String
    jsonText = "{""orders"":[{""items"":[{""sku"":""A0"",""qty"":1}]},{""items"":[{""sku"":""B0"",""qty"":2}]}]}"

    Dim parsed As Variant
    Set parsed = Json_Parse(jsonText)

    Dim flat As Collection
    Set flat = Json_Flatten(parsed)

    Dim rows As Collection
    Set rows = Json_ExtractTableRows(flat, "$.orders.items")

    Dim headers As Variant
    Dim data As Variant
    data = Json_TableTo2D(rows, headers)

    Dim expHeaders As Variant
    expHeaders = Array("sku", "qty")
    Assert1DArrayEquals expHeaders, headers, "headers"

    AssertEquals 2, (UBound(data, 1) - LBound(data, 1) + 1), "row count"
    AssertEquals 2, (UBound(data, 2) - LBound(data, 2) + 1), "col count"

    AssertEquals "A0", CStr(data(1, 1)), "r1 sku"
    AssertEquals 1, CLng(data(1, 2)), "r1 qty"
    AssertEquals "B0", CStr(data(2, 1)), "r2 sku"
    AssertEquals 2, CLng(data(2, 2)), "r2 qty"
End Sub

Private Sub Test_ExtractTableRows_NestedChildRows_EmptyChildArray_ReturnsZeroRows()
    ' Goal:
    '   If some parents have empty child arrays, we only get rows from non-empty ones.
    '
    ' Input:
    '   orders[0].items = []
    '   orders[1].items = [{sku:"B0"}]
    '
    ' Expected:
    '   rows.Count = 1 and sku="B0"
    '
    Dim jsonText As String
    jsonText = "{""orders"":[{""id"":1,""items"":[]},{""id"":2,""items"":[{""sku"":""B0""}]}]}"

    Dim parsed As Variant
    Set parsed = Json_Parse(jsonText)

    Dim flat As Collection
    Set flat = Json_Flatten(parsed)

    Dim rows As Collection
    Set rows = Json_ExtractTableRows(flat, "$.orders.items")

    AssertEquals 1, rows.count, "rows count with one empty parent child array"
    AssertEquals "B0", Obj_GetValue(rows(1), "sku"), "remaining child row sku"
End Sub

Private Sub Test_FlatGet_EscapedDotAndBackslash_ExactMatch()
    ' Goal: Json_FlatGet must resolve escaped segments exactly (no partial/greedy matching).
    ' Input: { "a.b": { "c\\d": 1 }, "a": { "b": { "c\\d": 2 } } }
    ' Expect:
    '   $.a\.b.c\\d == 1
    '   $.a.b.c\\d  == 2
    Dim jsonText As String
    jsonText = "{""a.b"":{""c\\d"":1},""a"":{""b"":{""c\\d"":2}}}"

    Dim parsed As Variant
    Set parsed = Json_Parse(jsonText)

    Dim flat As Collection
    Set flat = Json_Flatten(parsed)

    AssertEquals 1, CLng(Json_FlatGet(flat, "$.a\.b.c\\d")), "escaped-dot path resolves"
    AssertEquals 2, CLng(Json_FlatGet(flat, "$.a.b.c\\d")), "unescaped-dot path resolves"
End Sub

Private Sub Test_FlatContains_DoesNotConfuseEscapes()
    ' Goal: Contains must not treat escaped and unescaped paths as interchangeable.
    Dim jsonText As String
    jsonText = "{""a.b"":1,""a"":{""b"":2}}"

    Dim parsed As Variant
    Set parsed = Json_Parse(jsonText)

    Dim flat As Collection
    Set flat = Json_Flatten(parsed)

    AssertTrue Json_FlatContains(flat, "$.a\.b"), "contains escaped key a.b"
    AssertTrue Json_FlatContains(flat, "$.a.b"), "contains nested a.b"
    AssertTrue (CLng(Json_FlatGet(flat, "$.a\.b")) = 1), "escaped value"
    AssertTrue (CLng(Json_FlatGet(flat, "$.a.b")) = 2), "nested value"
End Sub

Private Sub Test_ExtractTableRows_TableRootMustBeArrayOfObjects()
    ' Goal: Json_ExtractTableRows should fail clearly when tableRoot does not refer to array-of-objects.
    ' Input:
    '   $.x is array of primitives
    ' Expect:
    '   raises error (any nonzero), rather than returning nonsense rows
    Dim jsonText As String
    jsonText = "{""x"":[1,2,3],""y"":[{""a"":1}]}"

    Dim parsed As Variant
    Set parsed = Json_Parse(jsonText)

    Dim flat As Collection
    Set flat = Json_Flatten(parsed)

    On Error GoTo expected
    Call Json_ExtractTableRows(flat, "$.x")
    Err.Raise vbObjectError + 699, "mJsonTests", "Expected Json_ExtractTableRows to reject $.x (array of primitives)"
expected:
    AssertTrue Err.Number <> 0, "ExtractTableRows should error on array-of-primitives root"
    Err.Clear
End Sub

Private Sub Test_ExtractTableRows_TableRootMustExist()
    ' Goal: Json_ExtractTableRows should error if the root path does not exist in the flattened object.
    Dim jsonText As String
    jsonText = "{""orders"":[{""id"":1}]}"

    Dim parsed As Variant
    Set parsed = Json_Parse(jsonText)

    Dim flat As Collection
    Set flat = Json_Flatten(parsed)

    On Error GoTo expected
    Call Json_ExtractTableRows(flat, "$.missing")
    Err.Raise vbObjectError + 699, "mJsonTests", "Expected Json_ExtractTableRows to error for missing root"
expected:
    AssertTrue Err.Number <> 0, "ExtractTableRows should error on missing root"
    Err.Clear
End Sub

Private Sub Test_ExtractTableRows_RootArray_PrimitiveRowsRejected()
    ' Goal: Root array that is primitives should be rejected even at "$".
    Dim v As Variant
    Json_ParseInto "[1,2,3]", v

    Dim flat As Collection
    Set flat = Json_Flatten(v)

    On Error GoTo expected
    Call Json_ExtractTableRows(flat, "$")
    Err.Raise vbObjectError + 699, "mJsonTests", "Expected Json_ExtractTableRows to reject root array-of-primitives"
expected:
    AssertTrue Err.Number <> 0, "ExtractTableRows should error on root array-of-primitives"
    Err.Clear
End Sub
