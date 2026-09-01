VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsJson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Whitespace, NumberRegex, StringChunk
Private b, f, r, n, t

Private Sub Class_Initialize()
    Whitespace = " " & vbTab & vbCr & vbLf
    b = ChrW(8)
    f = vbFormFeed
    r = vbCr
    n = vbLf
    t = vbTab
    Set NumberRegex = CreateObject("vbscript.regexp")
    NumberRegex.Pattern = "(-?(?:0|[1-9]\d*))(\.\d+)?([eE][-+]?\d+)?"
    NumberRegex.Global = False
    NumberRegex.MultiLine = True
    NumberRegex.IgnoreCase = True
    Set StringChunk = CreateObject("vbscript.regexp")
    StringChunk.Pattern = "([\s\S]*?)([""\\\x00-\x1f])"
    StringChunk.Global = False
    StringChunk.MultiLine = True
    StringChunk.IgnoreCase = True
End Sub

Public Function encode(ByRef obj)
    Dim buf, i, c, g
    Set buf = CreateObject("Scripting.Dictionary")
    Select Case VarType(obj)
    Case vbNull
        buf.Add buf.count, "null"
    Case vbBoolean
        If obj Then
            buf.Add buf.count, "true"
        Else
            buf.Add buf.count, "false"
        End If
    Case vbInteger, vbLong, vbSingle, vbDouble
        buf.Add buf.count, obj
    Case vbString
        buf.Add buf.count, """"
        For i = 1 To Len(obj)
            c = Mid(obj, i, 1)
            Select Case c
            Case """": buf.Add buf.count, "\"""
            Case "\": buf.Add buf.count, "\\"
            Case "/": buf.Add buf.count, "/"
            Case b: buf.Add buf.count, "\b"
            Case f: buf.Add buf.count, "\f"
            Case r: buf.Add buf.count, "\r"
            Case n: buf.Add buf.count, "\n"
            Case t: buf.Add buf.count, "\t"
            Case Else
                If AscW(c) >= 0 And AscW(c) <= 31 Then
                    c = Right("0" & Hex(AscW(c)), 2)
                    buf.Add buf.count, "\u00" & c
                Else
                    buf.Add buf.count, c
                End If
            End Select
        Next
        buf.Add buf.count, """"
    Case vbArray + vbVariant
        g = True
        buf.Add buf.count, "["
        For Each i In obj
            If g Then g = False Else buf.Add buf.count, ","
            buf.Add buf.count, encode(i)
        Next
        buf.Add buf.count, "]"
    Case vbObject
        If TypeName(obj) = "Dictionary" Then
            g = True
            buf.Add buf.count, "{"
            For Each i In obj
                If g Then g = False Else buf.Add buf.count, ","
                buf.Add buf.count, """" & i & """" & ":" & encode(obj(i))
            Next
            buf.Add buf.count, "}"
        Else
            Err.Raise 8732, , "None dictionary object"
        End If
    Case Else
        buf.Add buf.count, """" & CStr(obj) & """"
    End Select
    encode = Join(buf.items, "")
End Function

Public Function Decode(ByRef str)
    Dim idx
    idx = SkipWhitespace(str, 1)
    If Mid(str, idx, 1) = "{" Then
        Set Decode = ScanOnce(str, 1)
    Else
        Decode = ScanOnce(str, 1)
    End If
End Function

Private Function ScanOnce(ByRef str, ByRef idx)
    Dim c, ms
    idx = SkipWhitespace(str, idx)
    c = Mid(str, idx, 1)
    If c = "{" Then
        idx = idx + 1
        Set ScanOnce = parseObject(str, idx)
        Exit Function
    ElseIf c = "[" Then
        idx = idx + 1
        ScanOnce = parseArray(str, idx)
        Exit Function
    ElseIf c = """" Then
        idx = idx + 1
        ScanOnce = parseString(str, idx)
        Exit Function
    ElseIf c = "n" And StrComp("null", Mid(str, idx, 4)) = 0 Then
        idx = idx + 4
        ScanOnce = Null
        Exit Function
    ElseIf c = "t" And StrComp("true", Mid(str, idx, 4)) = 0 Then
        idx = idx + 4
        ScanOnce = True
        Exit Function
    ElseIf c = "f" And StrComp("false", Mid(str, idx, 5)) = 0 Then
        idx = idx + 5
        ScanOnce = False
        Exit Function
    End If
    Set ms = NumberRegex.Execute(Mid(str, idx))
    If ms.count = 1 Then
        idx = idx + ms(0).length
        ScanOnce = CDbl(ms(0))
        Exit Function
    End If
    Err.Raise 8732, , "No JSON object could be ScanOnced"
End Function

Private Function parseObject(ByRef str, ByRef idx)
    Dim c, key, value
    Set parseObject = CreateObject("Scripting.Dictionary")
    idx = SkipWhitespace(str, idx)
    c = Mid(str, idx, 1)
    If c = "}" Then
        Exit Function
    ElseIf c <> """" Then
        Err.Raise 8732, , "Expecting property name"
    End If
    idx = idx + 1
    Do
        key = parseString(str, idx)
        idx = SkipWhitespace(str, idx)
        If Mid(str, idx, 1) <> ":" Then
            Err.Raise 8732, , "Expecting : delimiter"
        End If
        idx = SkipWhitespace(str, idx + 1)
        If Mid(str, idx, 1) = "{" Then
            Set value = ScanOnce(str, idx)
        Else
            value = ScanOnce(str, idx)
        End If
        parseObject.Add key, value
        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)
        If c = "}" Then
            Exit Do
        ElseIf c <> "," Then
            Err.Raise 8732, , "Expecting , delimiter"
        End If
        idx = SkipWhitespace(str, idx + 1)
        c = Mid(str, idx, 1)
        If c <> """" Then
            Err.Raise 8732, , "Expecting property name"
        End If
        idx = idx + 1
    Loop
    idx = idx + 1
End Function

Private Function parseArray(ByRef str, ByRef idx)
    Dim c, values, value
    Set values = CreateObject("Scripting.Dictionary")
    idx = SkipWhitespace(str, idx)
    c = Mid(str, idx, 1)
    If c = "]" Then
        parseArray = values.items
        Exit Function
    End If
    Do
        idx = SkipWhitespace(str, idx)
        If Mid(str, idx, 1) = "{" Then
            Set value = ScanOnce(str, idx)
        Else
            value = ScanOnce(str, idx)
        End If
        values.Add values.count, value
        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)
        If c = "]" Then
            Exit Do
        ElseIf c <> "," Then
            Err.Raise 8732, , "Expecting , delimiter"
        End If
        idx = idx + 1
    Loop
    idx = idx + 1
    parseArray = values.items
End Function

Private Function parseString(ByRef str, ByRef idx)
    Dim chunks, content, terminator, ms, esc, char
    Set chunks = CreateObject("Scripting.Dictionary")
    Do
        Set ms = StringChunk.Execute(Mid(str, idx))
        If ms.count = 0 Then
            Err.Raise 8732, , "Unterminated string starting"
        End If
        content = ms(0).SubMatches(0)
        terminator = ms(0).SubMatches(1)
        If Len(content) > 0 Then
            chunks.Add chunks.count, content
        End If
        idx = idx + ms(0).length
        If terminator = """" Then
            Exit Do
        ElseIf terminator <> "\" Then
            Err.Raise 8732, , "Invalid control character"
        End If
        esc = Mid(str, idx, 1)
        If esc <> "u" Then
            Select Case esc
            Case """": char = """"
            Case "\": char = "\"
            Case "/": char = "/"
            Case "b": char = b
            Case "f": char = f
            Case "n": char = n
            Case "r": char = r
            Case "t": char = t
            Case Else: Err.Raise 8732, , "Invalid escape"
            End Select
            idx = idx + 1
        Else
: char = ChrW("&H" & Mid(str, idx + 1, 4))
            idx = idx + 5
        End If
        chunks.Add chunks.count, char
    Loop
    parseString = Join(chunks.items, "")
End Function

Private Function SkipWhitespace(ByRef str, ByVal idx)
    Do While idx <= Len(str) And _
       InStr(Whitespace, Mid(str, idx, 1)) > 0
        idx = idx + 1
    Loop
    SkipWhitespace = idx
End Function

Function ParseJson(strJson)
    Dim HTML As Object, Window
    Set HTML = CreateObject("htmlfile")
    Set Window = HTML.parentWindow
    Window.execScript "var Json = " & strJson & ";", "JScript"
    'Set ParseJson = Window.eval("Json")
    Set ParseJson = Window.Json
End Function



