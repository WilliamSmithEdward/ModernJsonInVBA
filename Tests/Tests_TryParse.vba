Attribute VB_Name = "Tests_TryParse"
Option Explicit

' =============================================================================
' Json_TryParse tests
'
' Validates the non-raising parse:
'   - valid objects, arrays, and scalars return True with the right model
'   - malformed input returns False instead of raising
'   - outValue is reset to Null and outError carries a reason on failure
'   - outError is empty on success
'   - a failure leaves no lingering error state (the next parse still works)
'   - the outError argument is optional
' =============================================================================

Public Sub RunAll_TryParseTests_StopOnFail()
    On Error GoTo Fail

    Test_TryParse_ValidObject
    Test_TryParse_ValidArray
    Test_TryParse_ValidScalars
    Test_TryParse_Malformed
    Test_TryParse_EmptyAndBlank
    Test_TryParse_Garbage
    Test_TryParse_OutValueNullOnFail
    Test_TryParse_ErrorMessageOverwritten
    Test_TryParse_RecoversAfterFailure
    Test_TryParse_OptionalErrorOmitted

    MsgBox "All Json_TryParse tests passed.", vbInformation
    Exit Sub

Fail:
    Err.Raise vbObjectError + 780, "mTryParseTests", _
        "Json_TryParse test run failed. Err " & Err.Number & ": " & Err.Description
End Sub


' =============================================================================
' ASSERTS
' =============================================================================

Private Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal message As String)
    If expected <> actual Then
        Err.Raise vbObjectError + 781, "mTryParseTests", _
            message & " expected=" & CStr(expected) & " actual=" & CStr(actual)
    End If
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 782, "mTryParseTests", message
End Sub


' =============================================================================
' TESTS
' =============================================================================

Private Sub Test_TryParse_ValidObject()
    Dim v As Variant
    Dim e As String
    AssertTrue Json_TryParse("{""a"":1,""b"":2}", v, e), "valid object returns True"
    AssertEquals "", e, "no error text on success"
    AssertTrue Json_IsObject(v), "result is an object"
    AssertEquals 1, Json_ObjGet(v, "a"), "member a"
    AssertEquals 2, Json_ObjGet(v, "b"), "member b"
End Sub

Private Sub Test_TryParse_ValidArray()
    Dim v As Variant
    AssertTrue Json_TryParse("[10,20,30]", v), "valid array returns True"
    AssertTrue Json_IsArray(v), "result is an array"
    AssertEquals 3, v.count, "array length"
    AssertEquals 20, v(2), "element 2"
End Sub

Private Sub Test_TryParse_ValidScalars()
    Dim v As Variant
    AssertTrue Json_TryParse("42", v), "number scalar returns True"
    AssertEquals 42, v, "number value"
    AssertTrue Json_TryParse("true", v), "boolean scalar returns True"
    AssertEquals True, v, "boolean value"
    AssertTrue Json_TryParse("null", v), "null scalar returns True"
    AssertTrue IsNull(v), "null value"
    AssertTrue Json_TryParse("""hi""", v), "string scalar returns True"
    AssertEquals "hi", v, "string value"
End Sub

Private Sub Test_TryParse_Malformed()
    Dim v As Variant
    AssertTrue Not Json_TryParse("{}x", v), "trailing characters"
    AssertTrue Not Json_TryParse("{", v), "unclosed object"
    AssertTrue Not Json_TryParse("[1,2", v), "unclosed array"
    AssertTrue Not Json_TryParse("{""a"" 1}", v), "missing colon"
    AssertTrue Not Json_TryParse("{""a"":""oops", v), "unterminated string"
    AssertTrue Not Json_TryParse("[1,,2]", v), "double comma"
    AssertTrue Not Json_TryParse("nul", v), "bad literal"
End Sub

Private Sub Test_TryParse_EmptyAndBlank()
    Dim v As Variant
    AssertTrue Not Json_TryParse("", v), "empty string returns False"
    AssertTrue Not Json_TryParse("   ", v), "whitespace-only returns False"
End Sub

Private Sub Test_TryParse_Garbage()
    Dim v As Variant
    Dim e As String
    AssertTrue Not Json_TryParse("this is not json", v, e), "garbage returns False"
    AssertTrue Len(e) > 0, "garbage yields an error message"
End Sub

Private Sub Test_TryParse_OutValueNullOnFail()
    ' outValue must be reset even when it held a prior value and even when the
    ' parser assigned a partial value before raising.
    Dim v As Variant
    v = "sentinel"
    AssertTrue Not Json_TryParse("{""a"":1 bad", v), "partial-then-bad returns False"
    AssertTrue IsNull(v), "outValue reset to Null on failure"
End Sub

Private Sub Test_TryParse_ErrorMessageOverwritten()
    Dim v As Variant
    Dim e As String
    e = "preset"
    AssertTrue Not Json_TryParse("[1,,2]", v, e), "malformed returns False"
    AssertTrue Len(e) > 0, "error message populated"
    AssertTrue e <> "preset", "error message overwrites the caller's value"
End Sub

Private Sub Test_TryParse_RecoversAfterFailure()
    ' A failed parse must leave no lingering error state.
    Dim v As Variant
    Dim e As String
    Json_TryParse "{bad", v, e
    AssertTrue Json_TryParse("{""ok"":true}", v, e), "success after a prior failure"
    AssertEquals "", e, "error text cleared on the subsequent success"
    AssertTrue Json_IsObject(v), "recovered model is an object"
End Sub

Private Sub Test_TryParse_OptionalErrorOmitted()
    Dim v As Variant
    AssertTrue Json_TryParse("{""x"":1}", v), "works without outError (success)"
    AssertTrue Not Json_TryParse("{bad", v), "works without outError (failure)"
End Sub
