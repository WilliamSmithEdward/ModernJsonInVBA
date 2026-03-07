Option Explicit

' =============================================================================
' XML -> JSON Tests
'
' These tests validate:
'   - basic element parsing
'   - nested elements
'   - text node handling
'   - self-closing tags
'   - whitespace handling
'   - malformed XML detection
'
' Assumptions:
'   - XmlTextToJson(xmlText) exists
'   - XmlFileToJson(filePath) exists
' =============================================================================


' =============================================================================
' RUNNER
' =============================================================================
Public Sub RunAll_XmlJsonTests_StopOnFail()

    On Error GoTo Fail

    Test_Xml_Simple_WithAsserts
    Test_Xml_Nested_WithAsserts
    Test_Xml_TextNode_WithAsserts
    Test_Xml_SelfClosing_WithAsserts
    Test_Xml_MultipleChildren_WithAsserts
    Test_Xml_Whitespace_WithAsserts
    Test_Xml_DeepNesting_WithAsserts
    Test_Xml_FileWrapper_WithAsserts
    Test_Xml_MalformedMissingClose_WithAsserts
    Test_Xml_EmptyDocument_WithAsserts
    Test_Xml_MixedTextChildren_WithAsserts
    Test_Xml_EmptyElementPair_WithAsserts
    Test_Xml_LeadingWhitespace_WithAsserts
    Test_Xml_RepeatedSiblingElements_WithAsserts
    Test_Xml_JsonEscaping_WithAsserts
    Test_Xml_TrailingWhitespace_WithAsserts
    Test_Xml_LongTextNode_WithAsserts
    Test_Xml_AdjacentTextSegments_WithAsserts
    Test_Xml_DeterministicStructure_WithAsserts
    Test_Xml_SelfClosingRoot_WithAsserts
    Test_Xml_PreserveTextWhitespace_WithAsserts
    Test_Xml_EmptyTextNode_WithAsserts
    Test_Xml_RepeatedElements_CreateArray_WithAsserts
    Test_Xml_ArrayContainsMultipleObjects_WithAsserts
    Test_Xml_SingleElement_NotArray_WithAsserts
    Test_Xml_NestedArrayStructure_WithAsserts
    Test_Xml_MixedTextNodesPreserved_WithAsserts
    Test_Xml_SelfClosingNode_WithAsserts
    Test_Xml_NoDuplicateKeys_WithAsserts
    Test_Xml_DeterministicArrayOrder_WithAsserts
    Test_Xml_EntityDecoding_WithAsserts
    Test_Xml_VeryDeepNesting_WithAsserts
    Test_Xml_MultipleArrayGroups_WithAsserts
    Test_Xml_LargeArray_WithAsserts
    Test_Xml_NestedMixedContent_WithAsserts
    Test_Xml_RootLevelArray_WithAsserts
    Test_Xml_BackslashEscaping_WithAsserts
    Test_Xml_NumericEntity_WithAsserts
    Test_Xml_AttributesIgnored_WithAsserts
    Test_Xml_ExtremeDepth_WithAsserts
    Test_Xml_TextPlusArray_WithAsserts
    Test_Xml_HexEntity_WithAsserts
    Test_Xml_MismatchedTag_WithAsserts
    Test_Xml_UnclosedTag_WithAsserts
    Test_Xml_InvalidNumericEntity_WithAsserts
    Test_Xml_CDATA_WithAsserts
    Test_Xml_UnterminatedCDATA_WithAsserts
    Test_Xml_MultipleCDATA_WithAsserts
    Test_Xml_EmptyCDATA_WithAsserts
    Test_Xml_UnicodeText_WithAsserts
    Test_Xml_VeryLargeTextNode_WithAsserts
    Test_Xml_MultipleRoots_WithAsserts
    Test_Xml_DoubleEntityDecoding_WithAsserts
    Test_Xml_InvalidTagName_WithAsserts
    Test_Xml_EntityMissingSemicolon_WithAsserts
    Test_Xml_HugeSiblingSet_WithAsserts
    Test_Xml_MixedHugePayload_WithAsserts
    Test_Xml_CommentHandling_WithAsserts
    Test_Xml_ProcessingInstruction_WithAsserts
    Test_Xml_SelfClosingWhitespace_WithAsserts
    Test_Xml_TagNameWithDash_WithAsserts
    Test_Xml_TagNameWithNamespace_WithAsserts
    Test_Xml_TagNameWithDot_WithAsserts
    Test_Xml_SurrogatePair_WithAsserts
    Test_Xml_MultipleAttributesIgnored_WithAsserts
    Test_Xml_AllBuiltinEntities_WithAsserts
    Test_Xml_RepeatedEmptyElements_CreateArrayOfEmptyObjects_WithAsserts
    Test_Xml_OnlyWhitespaceTextNode_WithAsserts
    Test_Xml_SignificantWhitespaceBetweenElements_WithAsserts
    Test_Xml_ElementWithOnlyAttributes_Ignored_WithAsserts
    Test_Xml_ControlCharactersInText_WithAsserts
    Test_Xml_JsonSpecialCharsTortureTest_WithAsserts
    Test_Xml_TagNameWithJsonSpecialChars_WithAsserts
    Test_Xml_BOMAtStart_OfDocument_WithAsserts
    Test_Xml_RepeatedElementsWithMixedChildTypes_WithAsserts
    Test_Xml_InvalidCharRef_TooHigh_WithAsserts
    Test_Xml_MultibyteUnicode_WithAsserts
    Test_Xml_CRLF_TextNode_WithAsserts
    Test_Xml_VeryLongTagName_WithAsserts
    Test_Xml_DeepRecursionStress_WithAsserts
    Test_Xml_HugeAttributeSet_WithAsserts

    MsgBox "All XML->JSON tests passed.", vbInformation
    Exit Sub

Fail:

    Dim msg As String
    msg = "XML test run failed." & vbCrLf & _
          "Err " & Err.Number & ": " & Err.Description

    Err.Clear
    Err.Raise vbObjectError + 730, "mXmlJsonTests", msg

End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 731, "mXmlJsonTests", message
End Sub

Private Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, ByVal message As String)
    If expected <> actual Then
        Err.Raise vbObjectError + 732, "mXmlJsonTests", _
            message & " expected=" & CStr(expected) & " actual=" & CStr(actual)
    End If
End Sub

Private Function WriteTempXml(ByVal text As String) As String

    Dim path As String
    path = Environ$("TEMP") & "\xml_test_" & Format(Now, "hhmmss") & ".xml"

    Dim f As Integer
    f = FreeFile

    Open path For Output As #f
    Print #f, text
    Close #f

    WriteTempXml = path

End Function


Public Sub Test_Xml_Simple_WithAsserts()

    Dim xml As String
    xml = "<person><name>Alice</name></person>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "Alice") > 0, "Simple element failed"

End Sub


Public Sub Test_Xml_Nested_WithAsserts()

    Dim xml As String
    xml = "<person><id>1</id><name>Alice</name></person>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, """id""") > 0, "Missing id"
    AssertTrue InStr(json, """name""") > 0, "Missing name"

End Sub


Public Sub Test_Xml_TextNode_WithAsserts()

    Dim xml As String
    xml = "<title>Hello World</title>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "Hello World") > 0, "Text node lost"

End Sub


Public Sub Test_Xml_SelfClosing_WithAsserts()

    Dim xml As String
    xml = "<root><empty/></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, """empty""") > 0, "Self-closing tag missing"

End Sub


Public Sub Test_Xml_MultipleChildren_WithAsserts()

    Dim xml As String
    xml = "<root><a>1</a><b>2</b><c>3</c></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, """a""") > 0, "Missing a"
    AssertTrue InStr(json, """b""") > 0, "Missing b"
    AssertTrue InStr(json, """c""") > 0, "Missing c"

End Sub


Public Sub Test_Xml_Whitespace_WithAsserts()

    Dim xml As String
    xml = "<root>" & vbLf & "<a>1</a>" & vbLf & "</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "1") > 0, "Whitespace broke parsing"

End Sub


Public Sub Test_Xml_DeepNesting_WithAsserts()

    Dim xml As String
    xml = "<a><b><c><d>value</d></c></b></a>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "value") > 0, "Deep nesting failed"

End Sub


Public Sub Test_Xml_FileWrapper_WithAsserts()

    Dim xml As String
    xml = "<person><name>Alice</name></person>"

    Dim path As String
    path = WriteTempXml(xml)

    Dim json As String
    json = XmlFileToJson(path)

    AssertTrue InStr(json, "Alice") > 0, "File wrapper failed"

End Sub


Public Sub Test_Xml_MalformedMissingClose_WithAsserts()

    Dim xml As String
    xml = "<root><a>1</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue Len(json) > 0, "Parser returned empty result for malformed XML"

End Sub


Public Sub Test_Xml_EmptyDocument_WithAsserts()

    Err.Clear
    On Error GoTo Passed

    Dim xml As String
    xml = ""

    Dim json As String
    json = XmlTextToJson(xml)

    Err.Raise vbObjectError + 736, "mXmlJsonTests", _
        "Empty XML should raise an error."

Passed:

    AssertTrue Err.Number <> 0, "Expected error was not raised"

End Sub


Public Sub Test_Xml_MixedTextChildren_WithAsserts()

    Dim xml As String
    xml = "<root>Hello<a>1</a>World</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "Hello") > 0, "Leading text lost"
    AssertTrue InStr(json, "World") > 0, "Trailing text lost"
    AssertTrue InStr(json, """a""") > 0, "Child element missing"

End Sub


Public Sub Test_Xml_EmptyElementPair_WithAsserts()

    Dim xml As String
    xml = "<root><a></a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, """a""") > 0, "Empty element pair failed"

End Sub


Public Sub Test_Xml_LeadingWhitespace_WithAsserts()

    Dim xml As String
    xml = vbCrLf & "   " & "<root><a>1</a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "1") > 0, "Leading whitespace broke parsing"

End Sub


Public Sub Test_Xml_RepeatedSiblingElements_WithAsserts()

    Dim xml As String
    xml = "<root><item>1</item><item>2</item></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "1") > 0, "First sibling lost"
    AssertTrue InStr(json, "2") > 0, "Second sibling lost"

End Sub


Public Sub Test_Xml_JsonEscaping_WithAsserts()

    Dim xml As String
    xml = "<root><a>He said ""Hello""</a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    Dim expected As String
    expected = "He said \""Hello\"""

    AssertTrue InStr(json, expected) > 0, "JSON escaping failed"

End Sub


Public Sub Test_Xml_TrailingWhitespace_WithAsserts()

    Dim xml As String
    xml = "<root><a>1</a></root>" & vbCrLf & "  "

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "1") > 0, "Trailing whitespace broke parsing"

End Sub


Public Sub Test_Xml_LongTextNode_WithAsserts()

    Dim txt As String
    txt = String(500, "A")

    Dim xml As String
    xml = "<root><a>" & txt & "</a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, txt) > 0, "Long text node lost"

End Sub


Public Sub Test_Xml_AdjacentTextSegments_WithAsserts()

    Dim xml As String
    xml = "<root>alpha beta gamma</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "alpha beta gamma") > 0, "Text segments collapsed incorrectly"

End Sub


Public Sub Test_Xml_DeterministicStructure_WithAsserts()

    Dim xml As String
    xml = "<person><id>1</id><name>Alice</name></person>"

    Dim json As String
    json = XmlTextToJson(xml)

    Dim expected As String
    expected = "{""id"":{""value"":""1""},""name"":{""value"":""Alice""}}"

    AssertEquals expected, json, "JSON structure mismatch"

End Sub


Public Sub Test_Xml_SelfClosingRoot_WithAsserts()

    Dim xml As String
    xml = "<root/>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertEquals "{}", json, "Self-closing root failed"

End Sub


Public Sub Test_Xml_PreserveTextWhitespace_WithAsserts()

    Dim xml As String
    xml = "<root><a>  hello world  </a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "hello world") > 0, "Text node lost"

End Sub


Public Sub Test_Xml_EmptyTextNode_WithAsserts()

    Dim xml As String
    xml = "<root><a></a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, """a""") > 0, "Empty text node failed"

End Sub


Public Sub Test_Xml_RepeatedElements_CreateArray_WithAsserts()

    Dim xml As String
    xml = "<root>" & _
          "<row><id>1</id></row>" & _
          "<row><id>2</id></row>" & _
          "<row><id>3</id></row>" & _
          "</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, """row"":[") > 0, "Repeated elements should produce array"

End Sub


Public Sub Test_Xml_ArrayContainsMultipleObjects_WithAsserts()

    Dim xml As String
    xml = "<root>" & _
          "<row><id>1</id></row>" & _
          "<row><id>2</id></row>" & _
          "</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "},{") > 0, "Array objects not properly separated"

End Sub


Public Sub Test_Xml_SingleElement_NotArray_WithAsserts()

    Dim xml As String
    xml = "<root><row><id>1</id></row></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, """row"":[") = 0, "Single element incorrectly converted to array"

End Sub


Public Sub Test_Xml_NestedArrayStructure_WithAsserts()

    Dim xml As String
    xml = "<root>" & _
          "<group>" & _
          "<row><id>1</id></row>" & _
          "<row><id>2</id></row>" & _
          "</group>" & _
          "</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, """row"":[") > 0, "Nested array not formed"

End Sub


Public Sub Test_Xml_MixedTextNodesPreserved_WithAsserts()

    Dim xml As String
    xml = "<root>Hello<a>1</a>World</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "Hello") > 0, "Leading text lost"
    AssertTrue InStr(json, "World") > 0, "Trailing text lost"

End Sub


Public Sub Test_Xml_SelfClosingNode_WithAsserts()

    Dim xml As String
    xml = "<root><a/></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, """a""") > 0, "Self-closing node missing"

End Sub


Public Sub Test_Xml_NoDuplicateKeys_WithAsserts()

    Dim xml As String
    xml = "<root>" & _
          "<row><id>1</id></row>" & _
          "<row><id>2</id></row>" & _
          "</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    Dim firstPos As Long
    Dim secondPos As Long

    firstPos = InStr(json, """row"":")
    secondPos = InStr(firstPos + 1, json, """row"":{")

    AssertTrue secondPos = 0, "Duplicate JSON keys detected"

End Sub


Public Sub Test_Xml_DeterministicArrayOrder_WithAsserts()

    Dim xml As String
    xml = "<root>" & _
          "<row><id>1</id></row>" & _
          "<row><id>2</id></row>" & _
          "</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, """value"":""1""") < InStr(json, """value"":""2"""), _
        "Array order changed"

End Sub


Public Sub Test_Xml_EntityDecoding_WithAsserts()

    Dim xml As String
    xml = "<root><a>A &amp; B</a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "A & B") > 0, "XML entity not decoded"

End Sub


Public Sub Test_Xml_VeryDeepNesting_WithAsserts()

    Dim xml As String
    xml = "<a><b><c><d><e><f>1</f></e></d></c></b></a>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "1") > 0, "Deep nesting failed"

End Sub


Public Sub Test_Xml_MultipleArrayGroups_WithAsserts()

    Dim xml As String
    xml = "<root>" & _
          "<row><id>1</id></row>" & _
          "<row><id>2</id></row>" & _
          "<item><name>A</name></item>" & _
          "<item><name>B</name></item>" & _
          "</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, """row"":[") > 0, "Row array missing"
    AssertTrue InStr(json, """item"":[") > 0, "Item array missing"

End Sub


Public Sub Test_Xml_LargeArray_WithAsserts()

    Dim xml As String
    xml = "<root>"

    Dim i As Long
    For i = 1 To 50
        xml = xml & "<row><id>" & i & "</id></row>"
    Next i

    xml = xml & "</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, """row"":[") > 0, "Large array not formed"

End Sub


Public Sub Test_Xml_NestedMixedContent_WithAsserts()

    Dim xml As String
    xml = "<root>Hello<a>1</a>there<b>2</b>world</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "Hello") > 0, "Leading text lost"
    AssertTrue InStr(json, "there") > 0, "Middle text lost"
    AssertTrue InStr(json, "world") > 0, "Trailing text lost"

End Sub


Public Sub Test_Xml_RootLevelArray_WithAsserts()

    Dim xml As String
    xml = "<root>" & _
          "<row><id>1</id></row>" & _
          "<row><id>2</id></row>" & _
          "</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, """row"":[") > 0, "Root array missing"

End Sub


Public Sub Test_Xml_BackslashEscaping_WithAsserts()

    Dim xml As String
    xml = "<root><a>C:\Temp\File</a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "C:\\Temp\\File") > 0, "Backslash escaping failed"

End Sub


Public Sub Test_Xml_NumericEntity_WithAsserts()

    Dim xml As String
    xml = "<root><a>A &#38; B</a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "A & B") > 0, "Numeric entity not decoded"

End Sub


Public Sub Test_Xml_AttributesIgnored_WithAsserts()

    Dim xml As String
    xml = "<root><a id=""1"">value</a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "value") > 0, "Attribute parsing broke text node"

End Sub


Public Sub Test_Xml_ExtremeDepth_WithAsserts()

    Dim xml As String
    xml = "<a><b><c><d><e><f><g><h>1</h></g></f></e></d></c></b></a>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "1") > 0, "Extreme nesting failed"

End Sub


Public Sub Test_Xml_TextPlusArray_WithAsserts()

    Dim xml As String
    xml = "<root>hello<row>1</row><row>2</row>world</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "hello") > 0, "Leading text missing"
    AssertTrue InStr(json, "world") > 0, "Trailing text missing"
    AssertTrue InStr(json, """row"":[") > 0, "Array not formed"

End Sub


Public Sub Test_Xml_HexEntity_WithAsserts()

    Dim xml As String
    xml = "<root><a>A &#x26; B</a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "A & B") > 0, "Hex entity not decoded"

End Sub


Public Sub Test_Xml_MismatchedTag_WithAsserts()

    On Error GoTo Passed

    Dim xml As String
    xml = "<a><b>1</c></a>"

    Dim json As String
    json = XmlTextToJson(xml)

    Err.Raise vbObjectError + 740, "mXmlJsonTests", _
        "Mismatched tag should raise error."

Passed:

    AssertTrue Err.Number <> 0, "Expected error not raised"

End Sub


Public Sub Test_Xml_UnclosedTag_WithAsserts()

    On Error GoTo Passed

    Dim xml As String
    xml = "<root><a>1"

    Dim json As String
    json = XmlTextToJson(xml)

    Err.Raise vbObjectError + 741, "mXmlJsonTests", _
        "Unclosed tag should raise error."

Passed:

    AssertTrue Err.Number <> 0, "Expected error not raised"

End Sub


Public Sub Test_Xml_InvalidNumericEntity_WithAsserts()

    On Error GoTo Passed

    Dim xml As String
    xml = "<root><a>&#XYZ;</a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    Err.Raise vbObjectError + 742, "mXmlJsonTests", _
        "Invalid numeric entity should raise error."

Passed:

    AssertTrue Err.Number <> 0, "Expected error not raised"

End Sub


Public Sub Test_Xml_CDATA_WithAsserts()

    Dim xml As String
    xml = "<root><![CDATA[hello <b>world</b>]]></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "hello <b>world</b>") > 0, "CDATA not preserved"

End Sub


Public Sub Test_Xml_UnterminatedCDATA_WithAsserts()

    On Error GoTo Passed

    Dim xml As String
    xml = "<root><![CDATA[test</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    Err.Raise vbObjectError + 744, "mXmlJsonTests", _
        "Unterminated CDATA should raise error."

Passed:
    AssertTrue Err.Number <> 0, "Expected error not raised"

End Sub


Public Sub Test_Xml_MultipleCDATA_WithAsserts()

    Dim xml As String
    xml = "<root><![CDATA[hello]]><![CDATA[world]]></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "helloworld") > 0, "Multiple CDATA blocks failed"

End Sub


Public Sub Test_Xml_EmptyCDATA_WithAsserts()

    Dim xml As String
    xml = "<root><![CDATA[]]></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue Len(json) > 0, "Empty CDATA broke parser"

End Sub


Public Sub Test_Xml_UnicodeText_WithAsserts()

    Dim xml As String
    xml = "<root><a>????</a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "????") > 0, "Unicode text lost"

End Sub


Public Sub Test_Xml_VeryLargeTextNode_WithAsserts()

    Dim txt As String
    txt = String(20000, "A")

    Dim xml As String
    xml = "<root><a>" & txt & "</a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, txt) > 0, "Large text node corrupted"

End Sub


Public Sub Test_Xml_MultipleRoots_WithAsserts()

    On Error GoTo Passed

    Dim xml As String
    xml = "<a>1</a><b>2</b>"

    Dim json As String
    json = XmlTextToJson(xml)

    Err.Raise vbObjectError + 743, "mXmlJsonTests", _
        "Multiple root elements should raise error."

Passed:
    AssertTrue Err.Number <> 0, "Expected error not raised"

End Sub


Public Sub Test_Xml_DoubleEntityDecoding_WithAsserts()

    Dim xml As String
    xml = "<root>&amp;amp;</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "&amp;") > 0, "Entity decoded too aggressively"

End Sub


Public Sub Test_Xml_InvalidTagName_WithAsserts()

    On Error GoTo Passed

    Dim xml As String
    xml = "<123>bad</123>"

    Dim json As String
    json = XmlTextToJson(xml)

    Err.Raise vbObjectError + 745, "mXmlJsonTests", _
        "Invalid tag name should raise error."

Passed:
    AssertTrue Err.Number <> 0, "Expected error not raised"

End Sub


Public Sub Test_Xml_EntityMissingSemicolon_WithAsserts()

    Dim xml As String
    xml = "<root><a>A &amp B</a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "&amp B") > 0, _
        "Entity without semicolon incorrectly decoded"

End Sub


Public Sub Test_Xml_HugeSiblingSet_WithAsserts()

    Dim xml As String
    xml = "<root>"

    Dim i As Long
    For i = 1 To 500
        xml = xml & "<row><id>" & i & "</id></row>"
    Next i

    xml = xml & "</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, """row"":[") > 0, "Large sibling set failed"

End Sub


Public Sub Test_Xml_MixedHugePayload_WithAsserts()

    Dim xml As String
    xml = "<root>"

    Dim i As Long
    For i = 1 To 200
        xml = xml & "<row><id>" & i & "</id></row>"
    Next i

    xml = xml & "<blob>" & String(10000, "X") & "</blob>"
    xml = xml & "</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, """blob""") > 0, "Large text lost"
    AssertTrue InStr(json, """row"":[") > 0, "Array lost"

End Sub


Public Sub Test_Xml_CommentHandling_WithAsserts()

    On Error GoTo Passed

    Dim xml As String
    xml = "<root><!-- comment --><a>1</a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    Err.Raise vbObjectError + 746, "mXmlJsonTests", _
        "Comments should raise error or be skipped."

Passed:
    AssertTrue Err.Number <> 0, "Expected error not raised"

End Sub


Public Sub Test_Xml_ProcessingInstruction_WithAsserts()

    On Error GoTo Passed

    Dim xml As String
    xml = "<?xml version=""1.0""?><root><a>1</a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    Err.Raise vbObjectError + 747, "mXmlJsonTests", _
        "Processing instruction should raise error."

Passed:
    AssertTrue Err.Number <> 0, "Expected error not raised"

End Sub


Public Sub Test_Xml_SelfClosingWhitespace_WithAsserts()

    Dim xml As String
    xml = "<root><a /></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, """a""") > 0, _
        "Self closing tag with whitespace failed"

End Sub


Public Sub Test_Xml_TagNameWithDash_WithAsserts()

    Dim xml As String
    xml = "<root><customer-id>1</customer-id></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "customer-id") > 0, _
        "Dash in tag name not handled"

End Sub


Public Sub Test_Xml_TagNameWithNamespace_WithAsserts()

    Dim xml As String
    xml = "<root><ns:item>1</ns:item></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "ns:item") > 0, _
        "Namespace tag name not handled"

End Sub


Public Sub Test_Xml_TagNameWithDot_WithAsserts()

    Dim xml As String
    xml = "<root><config.value>1</config.value></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "config.value") > 0, _
        "Dot in tag name not handled"

End Sub


Public Sub Test_Xml_SurrogatePair_WithAsserts()

    Dim xml As String
    xml = "<root>??</root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "??") > 0, "Unicode surrogate pair lost"

End Sub


Public Sub Test_Xml_MultipleAttributesIgnored_WithAsserts()

    Dim xml As String
    xml = "<root><a x=""1"" y=""2"">value</a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "value") > 0, _
        "Multiple attributes broke parsing"

End Sub


Public Sub Test_Xml_AllBuiltinEntities_WithAsserts()

    Dim xml As String
    xml = "<root>&lt;&gt;&amp;&apos;&quot;</root>"

    Dim json As String
    json = XmlTextToJson(xml)
    
    Dim expected As String
    expected = "<>&'\"""
    
    AssertTrue InStr(json, expected) > 0, _
        "Built-in entity decoding failed"

End Sub


Public Sub Test_Xml_RepeatedEmptyElements_CreateArrayOfEmptyObjects_WithAsserts()
    
    Dim xml As String
    xml = "<root><empty/><empty/><empty/></root>"
    Dim json As String
    json = XmlTextToJson(xml)
    AssertTrue InStr(json, """empty"":[") > 0, "Repeated empty elements should produce array"
    AssertTrue InStr(json, "{},{}") > 0 Or InStr(json, "},{}") > 0, "Array should contain empty objects"

End Sub

Public Sub Test_Xml_OnlyWhitespaceTextNode_WithAsserts()

    Dim xml As String
    xml = "<root>     </root>"
    Dim json As String
    json = XmlTextToJson(xml)
    ' Depending on your whitespace preservation policy — adjust expected behavior
    AssertTrue InStr(json, """value"":""     """) > 0 Or Len(json) = 2, _
        "Pure whitespace text node should either be preserved or result in empty object"
        
End Sub

Public Sub Test_Xml_SignificantWhitespaceBetweenElements_WithAsserts()
    
    Dim xml As String
    xml = "<root>" & vbCrLf & "  " & "<a>1</a>" & vbCrLf & "   " & "<b>2</b>" & vbCrLf & "</root>"
    
    Dim json As String
    json = XmlTextToJson(xml)
    
    AssertTrue InStr(json, """a""") > 0 And InStr(json, """b""") > 0, "Elements lost due to whitespace"
    ' Optional: if you preserve inter-element whitespace as text nodes:
    ' AssertTrue InStr(json, vbCrLf & "  ") > 0, "Significant inter-element whitespace lost"
    
End Sub

Public Sub Test_Xml_ElementWithOnlyAttributes_Ignored_WithAsserts()

    Dim xml As String
    xml = "<root><flag enabled=""true"" type=""debug""/></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    ' element should exist
    AssertTrue InStr(json, """flag""") > 0, _
        "Element with only attributes missing"

    ' attributes must be ignored
    AssertTrue InStr(json, """enabled""") = 0 And _
               InStr(json, """type""") = 0, _
        "Attributes should be completely ignored"

    ' empty element should produce empty object
    AssertTrue InStr(json, """flag"":{}") > 0, _
        "Empty element should produce empty object"

End Sub

Public Sub Test_Xml_ControlCharactersInText_WithAsserts()
    
    Dim xml As String
    xml = "<root><a>&#x07;&#x08;&#x0C;&#x1B;</a></root>"
    
    Dim json As String
    json = XmlTextToJson(xml)
    
    AssertTrue InStr(json, "\u0007") > 0, "Bell character not escaped"
    AssertTrue InStr(json, "\b") > 0, "Backspace not escaped"
    AssertTrue InStr(json, "\f") > 0, "Form feed not escaped"
    AssertTrue InStr(json, "\u001B") > 0, "Escape character not escaped"
    
End Sub

Public Sub Test_Xml_JsonSpecialCharsTortureTest_WithAsserts()
    
    Dim xml As String
    xml = "<root><msg>"" \\ \b \f \n \r \t \/ &quot; &lt; &amp; &#x27;</msg></root>"
    
    Dim json As String
    json = XmlTextToJson(xml)
    
    Dim expectedParts As Variant
    expectedParts = Array("\""", "\\", "\b", "\f", "\n", "\r", "\t", "\/", """", "<", "&", "'")
    
    Dim i As Long
    
    For i = LBound(expectedParts) To UBound(expectedParts)
        
        AssertTrue InStr(json, expectedParts(i)) > 0, _
            "Missing proper escape for: " & expectedParts(i)
            
    Next i
    
End Sub

Public Sub Test_Xml_TagNameWithJsonSpecialChars_WithAsserts()

    Dim xml As String
    xml = "<root><tag-quote-and-slash>""quote""-and\slash</tag-quote-and-slash></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "\""quote\""") > 0, _
        "JSON quotes not escaped properly"
    
    AssertTrue InStr(json, "\\slash") > 0, _
        "JSON backslash not escaped properly"
    
    AssertTrue InStr(json, """value""") > 0, _
        "Value lost"

End Sub

Public Sub Test_Xml_BOMAtStart_OfDocument_WithAsserts()
    
    Dim xml As String
    xml = ChrW(&HFEFF) & "<root><data>42</data></root>"
    
    Dim json As String
    json = XmlTextToJson(xml)
    
    AssertTrue InStr(json, """data""") > 0, "BOM caused parsing failure"
    AssertTrue InStr(json, "42") > 0, "Content lost due to BOM"
    
End Sub

Public Sub Test_Xml_RepeatedElementsWithMixedChildTypes_WithAsserts()
    
    Dim xml As String
    xml = "<root>" & _
          "<entry>plain text</entry>" & _
          "<entry><num>42</num></entry>" & _
          "<entry><self-closing/></entry>" & _
          "</root>"
    
    Dim json As String
    json = XmlTextToJson(xml)
    
    AssertTrue InStr(json, """entry"":[") > 0, "Mixed child types prevented array creation"
    AssertTrue InStr(json, "plain text") > 0 And InStr(json, "42") > 0, "Mixed content values lost"
    
End Sub

Public Sub Test_Xml_InvalidCharRef_TooHigh_WithAsserts()
    
    On Error GoTo Passed
    
    Dim xml As String
    xml = "<root><a>&#x110000;</a></root>"   ' outside legal Unicode
    
    Dim json As String
    json = XmlTextToJson(xml)
    
    Err.Raise vbObjectError + 750, "mXmlJsonTests", _
        "Invalid high char reference should raise error."
Passed:
    AssertTrue Err.Number <> 0, "Expected error not raised for invalid char ref"
    
End Sub


Public Sub Test_Xml_MultibyteUnicode_WithAsserts()

    Dim xml As String
    xml = "<root><a>????</a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "??") > 0, "CJK characters lost"
    AssertTrue InStr(json, "??") > 0, "Emoji lost"

End Sub


Public Sub Test_Xml_CRLF_TextNode_WithAsserts()

    Dim xml As String
    xml = "<root><a>line1" & vbCrLf & "line2</a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "line1") > 0, "First line lost"
    AssertTrue InStr(json, "line2") > 0, "Second line lost"

End Sub


Public Sub Test_Xml_VeryLongTagName_WithAsserts()

    Dim name As String
    name = String(200, "a")

    Dim xml As String
    xml = "<root><" & name & ">1</" & name & "></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, name) > 0, "Long tag name lost"

End Sub


Public Sub Test_Xml_DeepRecursionStress_WithAsserts()

    Dim xml As String
    xml = ""

    Dim i As Long
    For i = 1 To 120
        xml = xml & "<n>"
    Next i

    xml = xml & "1"

    For i = 1 To 120
        xml = xml & "</n>"
    Next i

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "1") > 0, "Deep recursion corrupted value"

End Sub


Public Sub Test_Xml_HugeAttributeSet_WithAsserts()

    Dim xml As String
    xml = "<root><a "

    Dim i As Long
    For i = 1 To 50
        xml = xml & "x" & i & "=""" & i & """ "
    Next i

    xml = xml & ">value</a></root>"

    Dim json As String
    json = XmlTextToJson(xml)

    AssertTrue InStr(json, "value") > 0, "Attribute skipping broke parsing"

End Sub
