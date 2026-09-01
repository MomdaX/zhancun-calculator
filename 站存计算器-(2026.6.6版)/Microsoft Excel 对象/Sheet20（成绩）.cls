VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Declare Function InternetSetCookie Lib "wininet.dll" Alias "InternetSetCookieA" (ByVal lpszUrl As String, ByVal lpszCookieName As String, ByVal lpszCookieData As String) As Long

Private Sub CommandButton4_Click()
    Dim url$, ck$, 试卷$, 批注$, 试卷编号$
    Dim js As Object, Json As Object, cls As New ClsJson, trr(), jrr, dt As New Dictionary, tm, sj As String
    Dim 题型 As String, 题干 As String, 答案 As String, 选项 As String
    
    t = Timer
    批注 = Me.Cells(Selection.Row, "J").Comment.text
    试卷编号 = Split(Split(批注, "=")(1), "&")(0)
    试卷 = 提取试卷(试卷编号, Me.Cells(Selection.Row, "Q")) '同步
    
    Application.ScreenUpdating = False
    
    If InStr(试卷, "topicArray.push") > 0 Then
        jrr = Split(试卷, "topicArray.push(")
        For i = 1 To UBound(jrr)
            If Left(jrr(i), 1) = "{" Then
                sj = Split(jrr(i), ");")(0) & "," & sj '拼接JS字符串tm = Split(jrr(i), ");")(0)
            End If
        Next
    End If
    
    ReDim trr(1 To UBound(jrr), 5)
    试卷 = "[" & Mid(sj, 1, Len(sj) - 1) & "]"
    Set Json = cls.ParseJson(试卷)
'    Set Json = JsonConverter.ParseJson(试卷)
'    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
'    ow.execScript "var js =JSON.stringify(" & sj & ");"
'    Set Json = ow.eavl("js[0]")
    For Each js In Json
'       Set js = 提取(tm, "(\w+?):(""[^""]*""|[^""},]*)")
        题型 = js.basetype
        题干 = Replace(js.topic, Chr(10), "")
'        If InStr(题干, "子系统包括钩计划编制和钩计划管理") Then Stop
        答案 = js.topicAnswer2
        选项 = Replace(js.topicOption, Chr(10), "")
        If Not dt.Exists(题型 & 题干 & 选项) Then
'           If Not dt.Exists(js(3).SubMatches(1) & js(4).SubMatches(1) & js(7).SubMatches(1)) Then
            k = k + 1
            trr(k, 0) = "=ROW()-1"
            trr(k, 1) = 题型 '题型js(3).SubMatches(1)
            trr(k, 2) = 题干 '题干js(4).SubMatches(1)
            trr(k, 3) = 答案 '答案js(9).SubMatches(1)
            If trr(k, 1) = "判断题" Then
                trr(k, 4) = "对$$错"
            Else
                trr(k, 4) = 选项 '选项js(7).SubMatches(1)
            End If
            dt(题型 & 题干 & 选项) = 答案 '加入字典去重
        End If
    Next
    
    With Sheet21
        .Range("A2").CurrentRegion.Offset(1).Clear '清空
        .Range("A2").Resize(UBound(trr, 1), UBound(trr, 2)) = trr '写入
        With .Cells
'           .Replace """", "", 2, , 0, 0
'           .WrapText = True '自动换行
            .HorizontalAlignment = xlCenter '居中
            .VerticalAlignment = xlCenter '居中
            .CurrentRegion.Borders.LineStyle = xlContinuous '设置样式边框
            Application.Intersect(.CurrentRegion.Offset(1), .Columns(3)).HorizontalAlignment = 1 '交集区域左对齐
        End With
        
        '分列
        Application.DisplayAlerts = False
        .Columns("E:E").TextToColumns DataType:=xlDelimited, TextQualifier:=xlNone, ConsecutiveDelimiter:=True, Other:=True, OtherChar:="$$", TrailingMinusNumbers:=True
        .Range("A1:L1") = Array("序号", "试题类型", "试题内容", "正确选项", "A", "B", "C", "D", "E", "F", "G", "H") '标题
        Application.DisplayAlerts = True
    End With
    
    '排序
    升序 Sheet21, Sheet21.Range("A1:L1"), Sheet21.Range("C1"), 1 '题干排序
    升序 Sheet21, Sheet21.Range("A1:L1"), Sheet21.Range("B1"), 1 '题型排序
    Application.ScreenUpdating = True
    
    MsgBox "完成！" & Chr(10) & "用时：" & Format(Timer - t, "0.00秒")
    
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim rng As Range, rn As Range, objData As Object
    If Target.address = "$F$1" Then
        Cancel = True
        试卷成绩查询.试卷列表
    ElseIf Target.address = "$F$4:$F$5" Then
        Cancel = True
        Sheet20.Range("H4").CurrentRegion.Offset(1).Clear
    ElseIf Target.Column = 3 And Target.Cells(1) <> "" Then
        Cancel = True
        Set rng = Target.CurrentRegion
        Set rn = rng.Rows(Target.Row)
        试卷成绩查询.试卷试题 rn
    ElseIf Target.Column = 1 And Target.Cells(1) <> "" Then
        Cancel = True
        Set rng = Target.CurrentRegion
        Set rn = rng.Rows(Target.Row)
       试卷成绩查询.成绩列表 rn
    ElseIf Target.Column = 10 And Target.Cells(1) <> "" Then
        Dim url As String, browser As Object
        Cancel = True
        url = Selection.Comment.text '获取批注中的网址
        Set browser = CreateObject("WScript.shell")
        browser.Run url '打开网站
        
'        UserForm6.Show 0
'        Set objData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
'        objData.SetText Target.Comment.Text
'        objData.PutInClipboard
'        MsgBox "已复制！！！"
    End If
End Sub


Private Sub CommandButton1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim rn As Range
    r = CommandButton1.TopLeftCell.Row
    Set rng = Me.Cells(r, "A").CurrentRegion
    Set rn = rng.Rows(Me.Cells(r, "A").Row)
    试卷成绩查询.成绩列表 rn
End Sub

Private Sub CommandButton2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim rn As Range
    r = CommandButton2.TopLeftCell.Row
    Set rng = Me.Cells(r, "A").CurrentRegion
    Set rn = rng.Rows(Me.Cells(r, "A").Row)
    试卷成绩查询.试卷试题 rn
End Sub

Private Sub CommandButton3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim r%
    r = CommandButton3.TopLeftCell.Row
    Sheet21.Range("O1").value = ActiveSheet.Cells(r, "A").value
    试卷成绩查询.题库获取 ActiveSheet.Cells(r, "C").value
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim rng As Range
    Set rng = Me.Range("H1").CurrentRegion
    If Not Intersect(rng, Target) Is Nothing Then
        With CommandButton4
            .Visible = True
            .width = 60
            .height = 30
            .Left = Me.Cells(Target.Row, rng.Column + rng.Columns.count).Left + 5
            .Top = Me.Cells(Target.Row, rng.Column + rng.Columns.count).Top
            .Caption = "单份导出" '(请选择要导出的试卷)
        End With
    Else
        CommandButton4.Visible = False
    End If
    If Target.Column = 1 And Target.Rows.count = 1 And Target.Row <> 1 Then
        If Target.count = 1 Then
            If Target <> "" Then
                With CommandButton1
                    .Visible = True
                    .width = 30
                    .height = 20
                    .Left = Target.Left + Target.width
                    .Top = Target.Top
                    .Caption = "成绩"
                End With
                With CommandButton2
                    .Visible = True
                    .width = 30
                    .height = 20
                    .Left = CommandButton1.Left + CommandButton1.width
                    .Top = Target.Top
                    .Caption = "试卷"
                End With
                With CommandButton3
                    .Visible = True
                    .width = 30
                    .height = 20
                    .Left = CommandButton2.Left + CommandButton2.width
                    .Top = Target.Top
                    .Caption = "题库"
                End With
            End If
        Else
            CommandButton1.Visible = False
            CommandButton2.Visible = False
            CommandButton3.Visible = False
        End If
    Else
        CommandButton1.Visible = False
        CommandButton2.Visible = False
        CommandButton3.Visible = False
    End If
End Sub
Sub HTML()
    UserForm6.Show 0
End Sub
