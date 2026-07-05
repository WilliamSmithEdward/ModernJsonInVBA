Attribute VB_Name = "Json_Common"
Option Explicit

' =============================================================================
' Module:      Json_Common
' Project:     ModernJsonInVBA
'
' Shared foundation for all ModernJsonInVBA modules: the object-tag constant,
' Variant assignment helper, JSON path utilities, a growable text builder,
' and an open-addressing string index used for fast header/key lookups.
'
' Everything here is host-agnostic (no Excel references) and dependency-free:
' no Scripting.Dictionary, no external libraries, no LongLong (compiles on
' both 32-bit and 64-bit Office).
'
' Members are Public so sibling modules can use them. Types and procedures in
' this module are internal plumbing for the library; the supported public API
' is documented in the README.
' =============================================================================

' Tag stored in slot (1) of a Collection to mark it as a JSON object.
' JSON arrays are UNTAGGED Collections. This tag is the authoritative signal
' that distinguishes the two; see Json_Model for the full model contract.
Public Const JSON_TAG_OBJECT As String = "__OBJ__"

' JsonTextBuilder: growable string buffer.
'
' VBA string concatenation (s = s & x) inside a loop is O(n^2) because every
' append copies the whole string. The builder pre-allocates a buffer, writes
' with Mid$ assignment, and doubles capacity as needed, for O(n) total cost.
Public Type JsonTextBuilder
    buffer As String
    used As Long
End Type

' JsonStringIndex: open-addressing hash index over strings.
'
' Maps key -> 1-based insertion index while preserving first-seen order in
' keys(1..count). One shared implementation backs header discovery, row-key
' mapping, root de-duplication, and schema reshaping across the library.
'
' textCompare:
'   False (default) => binary, case-sensitive keys
'   True            => case-insensitive (hashes LCase$, compares vbTextCompare)
'
' Load factor is kept under 0.7; capacity is always a power of two so probe
' positions can be computed with And-masking instead of Mod.
Public Type JsonStringIndex
    cap As Long          ' power-of-two slot count (0 = not yet initialized)
    slotHash() As Long
    slotIdx() As Long    ' 0 = empty slot, else 1-based index into keys()
    keys() As String
    count As Long
    textCompare As Boolean
End Type

' =============================================================================
' JsonTextBuilder procedures
' =============================================================================

Public Sub JsonSB_Init(ByRef sb As JsonTextBuilder, Optional ByVal initialCapacity As Long = 256)
    If initialCapacity < 16 Then initialCapacity = 16
    sb.buffer = Space$(initialCapacity)
    sb.used = 0
End Sub

Public Sub JsonSB_Append(ByRef sb As JsonTextBuilder, ByRef s As String)
    Dim addLen As Long
    addLen = Len(s)
    If addLen = 0 Then Exit Sub

    Dim capNow As Long
    capNow = Len(sb.buffer)

    If capNow = 0 Then
        capNow = 256
        Do While capNow < addLen
            capNow = capNow * 2
        Loop
        sb.buffer = Space$(capNow)
    ElseIf sb.used + addLen > capNow Then
        Do While sb.used + addLen > capNow
            capNow = capNow * 2
        Loop
        sb.buffer = sb.buffer & Space$(capNow - Len(sb.buffer))
    End If

    Mid$(sb.buffer, sb.used + 1, addLen) = s
    sb.used = sb.used + addLen
End Sub

Public Function JsonSB_Text(ByRef sb As JsonTextBuilder) As String
    If sb.used = 0 Then
        JsonSB_Text = vbNullString
    Else
        JsonSB_Text = Left$(sb.buffer, sb.used)
    End If
End Function

' =============================================================================
' JsonStringIndex procedures
' =============================================================================

Public Sub JsonIdx_Init( _
    ByRef m As JsonStringIndex, _
    Optional ByVal initialCapacity As Long = 64, _
    Optional ByVal textCompare As Boolean = False _
)
    Dim capPow2 As Long
    capPow2 = 16
    Do While capPow2 < initialCapacity
        capPow2 = capPow2 * 2
    Loop

    m.cap = capPow2
    ReDim m.slotHash(0 To m.cap - 1) As Long
    ReDim m.slotIdx(0 To m.cap - 1) As Long
    ReDim m.keys(1 To 16) As String
    m.count = 0
    m.textCompare = textCompare
End Sub

' Return the index of key, adding it if missing. Lazily initializes.
Public Function JsonIdx_Ensure(ByRef m As JsonStringIndex, ByVal key As String) As Long
    If m.cap = 0 Then JsonIdx_Init m

    If (m.count + 1) * 10 > m.cap * 7 Then
        JsonIdx_Rehash m, m.cap * 2
    End If

    Dim h As Long
    h = JsonIdx_HashKey(m, key)

    Dim mask As Long
    mask = m.cap - 1

    Dim pos As Long
    pos = (h And mask)

    Do
        If m.slotIdx(pos) = 0 Then
            m.count = m.count + 1
            If m.count > UBound(m.keys) Then
                ReDim Preserve m.keys(1 To UBound(m.keys) * 2) As String
            End If

            m.keys(m.count) = key
            m.slotHash(pos) = h
            m.slotIdx(pos) = m.count

            JsonIdx_Ensure = m.count
            Exit Function
        End If

        If m.slotHash(pos) = h Then
            If JsonIdx_KeysEqual(m, m.keys(m.slotIdx(pos)), key) Then
                JsonIdx_Ensure = m.slotIdx(pos)
                Exit Function
            End If
        End If

        pos = (pos + 1) And mask
    Loop
End Function

' Return the index of key, or 0 if absent.
Public Function JsonIdx_Find(ByRef m As JsonStringIndex, ByVal key As String) As Long
    If m.cap = 0 Then Exit Function

    Dim h As Long
    h = JsonIdx_HashKey(m, key)

    Dim mask As Long
    mask = m.cap - 1

    Dim pos As Long
    pos = (h And mask)

    Do
        If m.slotIdx(pos) = 0 Then Exit Function

        If m.slotHash(pos) = h Then
            If JsonIdx_KeysEqual(m, m.keys(m.slotIdx(pos)), key) Then
                JsonIdx_Find = m.slotIdx(pos)
                Exit Function
            End If
        End If

        pos = (pos + 1) And mask
    Loop
End Function

Private Sub JsonIdx_Rehash(ByRef m As JsonStringIndex, ByVal newCap As Long)
    Dim capPow2 As Long
    capPow2 = 16
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
        h = JsonIdx_HashKey(m, m.keys(i))

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

Private Function JsonIdx_HashKey(ByRef m As JsonStringIndex, ByRef key As String) As Long
    If m.textCompare Then
        JsonIdx_HashKey = Json_Hash32_FNV1a(LCase$(key))
    Else
        JsonIdx_HashKey = Json_Hash32_FNV1a(key)
    End If
End Function

Private Function JsonIdx_KeysEqual(ByRef m As JsonStringIndex, ByRef a As String, ByRef b As String) As Boolean
    If m.textCompare Then
        JsonIdx_KeysEqual = (StrComp(a, b, vbTextCompare) = 0)
    Else
        JsonIdx_KeysEqual = (a = b)
    End If
End Function

' FNV-1a 32-bit over UTF-16 code units.
'
' The 32-bit modular multiply is done with Double intermediates split into
' 16-bit halves (exact: every product stays below 2^42 < 2^53), avoiding
' LongLong so the code compiles on 32-bit Office.
Private Function Json_Hash32_FNV1a(ByRef s As String) As Long
    Const FNV_PRIME As Double = 16777619#
    Const TWO16 As Double = 65536#
    Const TWO32 As Double = 4294967296#

    Dim h As Long
    h = &H811C9DC5

    Dim i As Long
    For i = 1 To Len(s)
        h = h Xor (AscW(Mid$(s, i, 1)) And &HFFFF&)

        Dim lo As Double
        Dim hi As Double
        lo = (h And &HFFFF&) * FNV_PRIME
        hi = (((h And &HFFFF0000) \ &H10000) And &HFFFF&) * FNV_PRIME
        hi = hi - Int(hi / TWO16) * TWO16

        Dim u As Double
        u = lo + hi * TWO16
        u = u - Int(u / TWO32) * TWO32

        If u > 2147483647# Then
            h = CLng(u - TWO32)
        Else
            h = CLng(u)
        End If
    Next i

    Json_Hash32_FNV1a = h
End Function

' =============================================================================
' Variant assignment
' =============================================================================

' Assign src to dest using Set when src is an object. VBA requires different
' assignment statements for objects and values; this hides that split at call
' sites that handle both.
Public Sub Json_VarAssign(ByRef dest As Variant, ByVal src As Variant)
    If IsObject(src) Then
        Set dest = src
    Else
        dest = src
    End If
End Sub

' =============================================================================
' Path utilities
'
' Paths use a JSONPath-like syntax: "$" root, ".key" object member,
' "[0]" array index. Literal dots and backslashes inside keys are escaped
' as "\." and "\\" so path strings stay unambiguous.
' =============================================================================

Public Function Json_EscapePathSegment(ByVal s As String) As String
    s = Replace$(s, "\", "\\")
    s = Replace$(s, ".", "\.")
    Json_EscapePathSegment = s
End Function

Public Function Json_UnescapePathSegment(ByVal s As String) As String
    s = Replace$(s, "\.", ".")
    s = Replace$(s, "\\", "\")
    Json_UnescapePathSegment = s
End Function

' Split a dotted path into segments, honoring backslash escapes.
' Escapes are preserved in the returned tokens; callers unescape per segment.
Public Function Json_TokenizePath(ByVal path As String) As Collection
    Dim tokens As New Collection

    ' Fast path: no escapes, so native Split does the work. A trailing "."
    ' historically produced no empty final token; replicate that.
    If InStr(1, path, "\", vbBinaryCompare) = 0 Then
        If Len(path) > 0 Then
            Dim parts() As String
            parts = Split(path, ".")

            Dim n As Long
            n = UBound(parts)
            If Len(parts(n)) = 0 Then n = n - 1

            Dim j As Long
            For j = 0 To n
                tokens.Add parts(j)
            Next j
        End If

        Set Json_TokenizePath = tokens
        Exit Function
    End If

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

Public Function Json_IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    Json_IsAllDigits = (Len(s) > 0)
End Function

' Remove numeric array indices from a path: "$.a[0].b[12]" => "$.a.b".
' Non-numeric bracket content is preserved verbatim.
Public Function Json_RemoveIndices(ByVal s As String) As String
    ' Fast path: nothing to remove.
    Dim openPos As Long
    openPos = InStr(1, s, "[", vbBinaryCompare)
    If openPos = 0 Then
        Json_RemoveIndices = s
        Exit Function
    End If

    Dim sb As JsonTextBuilder
    JsonSB_Init sb, Len(s)

    Dim i As Long
    i = 1

    Do While openPos > 0
        ' Copy the run before "[" in one shot.
        If openPos > i Then
            JsonSB_Append sb, Mid$(s, i, openPos - i)
        End If

        Dim closePos As Long
        closePos = InStr(openPos + 1, s, "]", vbBinaryCompare)

        If closePos = 0 Then
            ' Unterminated bracket: keep the remainder verbatim.
            JsonSB_Append sb, Mid$(s, openPos)
            Json_RemoveIndices = JsonSB_Text(sb)
            Exit Function
        End If

        Dim inside As String
        inside = Mid$(s, openPos + 1, closePos - openPos - 1)

        If Not (Len(inside) > 0 And Json_IsAllDigits(inside)) Then
            JsonSB_Append sb, Mid$(s, openPos, closePos - openPos + 1)
        End If

        i = closePos + 1
        openPos = InStr(i, s, "[", vbBinaryCompare)
    Loop

    If i <= Len(s) Then
        JsonSB_Append sb, Mid$(s, i)
    End If

    Json_RemoveIndices = JsonSB_Text(sb)
End Function
