VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim ws As Worksheet, rng As Range, rngs As Range, HTML As String
    If Target.address(0, 0) = "K22" Then
        Cancel = True '取消单元格编辑模式
        Set ws = ActiveSheet
        Set rng = ws.Cells.Find("永鑫取送车记录").CurrentRegion
        Set rngs = rng.Resize(rng.Rows.count + 3, rng.Columns.count)
        rngs.Cells(rngs.Rows.count, rngs.Columns.count - 1).value = "打印时间:" & Now() '设置打印时间
        ws.PageSetup.PrintArea = rngs.address(0, 0) '设置打印区域
        ws.PrintOut
        ws.PageSetup.PrintArea = ws.Range("A1").CurrentRegion.address(0, 0) '重置打印区域
        rngs.Cells(rngs.Rows.count, rngs.Columns.count - 1).value = "" '删除打印时间
    ElseIf Target.address(0, 0) = "K2" Then
        Cancel = True '取消单元格编辑模式
        Set ws = ActiveSheet
        ws.PageSetup.PrintArea = ws.Range("A1").CurrentRegion.address(0, 0) '设置打印区域
        ws.PrintOut
    ElseIf Target.Cells(1).address(0, 0) = "M2" Then
        Cancel = True '取消单元格编辑模式
        调车计划.刷新计划列表
    ElseIf Target.Cells(1).address(0, 0) = "A1" Then
        Cancel = True '取消单元格编辑模式
        调车计划.读取调车钩数
    ElseIf Target <> "" And Target.count = 1 Then
        Cancel = True '取消单元格编辑模式
        Dim htmlpath As String
        If Not Target.Comment Is Nothing Then
            htmlpath = Target.Comment.text
            
            With 计划HTML.WebBrowser1
                .width = 560
                .height = 计划HTML.height + 10
                .navigate htmlpath
            End With
            
            计划HTML.Show
        End If
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
Dim ft As Date, et As Date, tf As String, reg As New RegExp
reg.Pattern = "\d+"
Application.EnableEvents = False
    Select Case Target.Row
        Case 3, 4, 5, 6, 7, 9, 10, 11, 12, 13, 15, 16, 17, 18, 19
        If Target.count = 1 And Target.Column = 3 Then
            If Target = "" Then Target.Offset(0, 1).Resize(1, 5).ClearContents '计划为空时，清空整行
        ElseIf Target.count = 1 Then
            '时间格式调整
            If IsNumeric(Target.value) And Target.value > 0 And Target.value < 1 Then
                If InStr(Target.value, "：") > 0 Then Target.value = Replace(Target.value, "：", ":")
                Target.NumberFormat = "hh:mm"
            ElseIf Target.value = "" Then
                Target.value = ""
            ElseIf Target.Column = 5 Or Target.Column = 6 Then
                ts = Format(Target.value, "0000")
                Target.NumberFormatLocal = "hh:mm"
                Target.value = Left(ts, 2) & ":" & Right(ts, 2)
            End If
            
            If Cells(Target.Row, "E") <> "" And Cells(Target.Row, "F") <> "" Then
                With Sheet15
                    '计算时间差
                    ft = TimeValue(Format(.Cells(Target.Row, "F"), "hh:mm"))
                    et = TimeValue(Format(.Cells(Target.Row, "E"), "hh:mm"))
                    
                    If DateDiff("n", et, ft) < 0 Then
                        tf = DateDiff("n", DateValue(Date) + et, DateValue(Date + 1) + ft)
                    Else
                        tf = DateDiff("n", et, ft)
                    End If
                    '作业时间
                    With .Cells(Target.Row, "G")
                        .NumberFormatLocal = "G/通用格式"
                        .value = tf + 15
                        .Interior.ColorIndex = IIf(tf > 389, 3, xlNone) '3红色(表示可能错误)
                    End With
                    '平均钩分
                    If .Cells(Target.Row, "D") <> "" Then
                    .Cells(Target.Row, "H").NumberFormatLocal = "0.0"
                    .Cells(Target.Row, "H") = Round(.Cells(Target.Row, "G") / .Cells(Target.Row, "D"), 1)
                    Else
                    .Cells(Target.Row, "H") = ""
                    .Cells(Target.Row, "H").NumberFormatLocal = "G/通用格式"
                    End If
                End With
            End If
        End If
        Case Is > 24
        If Target.Column = 1 And IsNumeric(Target.value) Then
            Target.NumberFormat = "hh:mm"
            If InStr(Target.value, "：") > 0 Then
                Target.value = Replace(Target.value, "：", ":")
                Application.EnableEvents = True
                Exit Sub
            ElseIf InStr(Target.value, ":") > 0 Then
                Application.EnableEvents = True
                Exit Sub
            ElseIf Target.value > 1 Then
                ts = Format(Target.value, "0000")
                Target.value = Left(ts, 2) & ":" & Right(ts, 2)
                Application.EnableEvents = True
                Exit Sub
            End If
        Else
            Application.EnableEvents = True
            Exit Sub
        End If
    End Select
    '合计
    With Sheet15
        .Cells(20, "C").value = "1调：" & WorksheetFunction.Sum(Range("D3:D7")) & "钩"
        .Cells(20, "D").value = "2调：" & WorksheetFunction.Sum(Range("D9:D13")) & "钩"
        .Cells(20, "E").value = "3调：" & WorksheetFunction.Sum(Range("D15:D19")) & "钩"
        .Cells(20, "F").value = "共：" & WorksheetFunction.Sum(Range("D3:D7,D9:D13,D15:D19")) & "钩"
        .Cells(20, "G").value = WorksheetFunction.Sum(Range("G3:G7,G9:G13,G15:G19"))
'        .Cells(1, "A").Value = Date$'设置日期
        .Cells(20, "H").NumberFormatLocal = "0.0"
        If reg.Execute(.Cells(20, "F").value)(0).value = 0 Or reg.Execute(.Cells(20, "F").value)(0).value = "" Then
            .Cells(20, "H").value = 0
        Else
'            Set mmk = reg.Execute(.Cells(20, "F").Value)
            .Cells(20, "H").value = val(.Cells(20, "G").value) / val(reg.Execute(.Cells(20, "F").value)(0).value)
        End If
    End With
Application.EnableEvents = True
End Sub

