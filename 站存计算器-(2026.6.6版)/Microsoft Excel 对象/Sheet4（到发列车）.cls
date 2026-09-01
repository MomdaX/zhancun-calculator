VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton_保存_Click()
    If Sheet14.UsedRange.Row = 1 And Sheet14.UsedRange.Rows.count = 1 Then
        Sheet4.UsedRange.Copy Sheet14.Cells(1, 1)
    Else
        Sheet4.UsedRange.Copy Sheet14.Cells(Sheet14.UsedRange.Row + Sheet14.UsedRange.Rows.count + 1, 1)
    End If
    Range("I3").RowHeight = 15
    With CommandButton_保存
        .Top = Range("I3").Top
        .Left = Range("I3").Left
        .width = Range("I3").width + 1
        .height = Range("I3").height + 5
    End With
End Sub

Private Sub CommandButton_清空_Click()
qr = MsgBox("是否要清空接发车信息", vbYesNo, "确认")
If qr = 6 Then
    With Sheet4
        .Cells.ClearComments '清空批注
        .Range("A1").value = Format(Date, "yyyy年mm月dd日")
        .Range("B4:V30").ClearContents
    End With
    Range("I2").RowHeight = 25
    With CommandButton_清空
        .Top = Range("I2").Top
        .Left = Range("I2").Left
        .width = Range("I2").width + 1
        .height = Range("I2").height
    End With
End If
End Sub

Private Sub CommandButton1_Click()
    违编核对.出发刷新
    违编核对.去向导出
End Sub

Private Sub CommandButton2_Click()
    违编核对.查询编组
End Sub

Private Sub CommandButton3_Click()
    违编核对.登录系统
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim rng As Range, M1 As Boolean, M2 As Boolean, M3 As Boolean
    If Target.Column = 1 Then End
    M1 = Target.Column = 12 Or Target.Column = 16
    M2 = Target.value <> ""
    M3 = Target.Column = 8
    If M1 And M2 Then
        Cancel = True ' 防止 Excel 默认执行双击编辑操作
        Sheet7.Select
        Sheet7.Range("A3:V3").AutoFilter 1, Sheet4.Cells(Target.Row, "L").value 'Target.Value
    ElseIf M2 And M3 Then
        Dim pk$
        Cancel = True ' 防止 Excel 默认执行双击编辑操作
        Sheet30.Select
        If Not Target.Comment Is Nothing Then
            pk = Split(Target.Comment.text, ":")(1)
            违编核对.获取编组 pk
        End If
    End If
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim btnArray() As Variant ' 存储按钮名称
    Dim btnName As Variant ' 单个按钮
    Dim btn As Object ' 按钮对象
    
    ' 隐藏按钮
    Me.sc.Visible = False
    Me.up.Visible = False
    Me.down.Visible = False
    
    '条件判断
    r = Selection.Row
    M1 = Target.Row < 4 ' 判断选中区域是否在第1~100行
    M2 = Sheet4.Cells(Rows.count, "D").End(3).Row < Target.Row And Sheet4.Cells(Rows.count, "L").End(3).Row < Target.Row
    M3 = Target.Column = 9 ' 判断选中区域是否在第K列
    M4 = Target.Column > 17
'    M6 = WorksheetFunction.CountA(Range("B" & r & ":" & "H" & r)) = 0 ' 判断数据区域是否为空
'    M7 = WorksheetFunction.CountA(Range("J" & r & ":" & "Q" & r)) = 0
    
    If M1 Or M2 Or M3 Or M4 Then Exit Sub

    ' 设置按钮位置
    btnArray = Array("sc", "up", "down")
    
    For Each btnName In btnArray
        Set btn = Me.Shapes(btnName)
        With btn
            .Left = Me.Range("I" & Target.Row).Left
            .Top = Me.Range("I" & Target.Row).Top + 偏移
            .width = Me.Range("I" & Target.Row).width
            .height = 20
            .Visible = True
        End With
        偏移 = 偏移 + 20
    Next btnName
End Sub


Private Sub sc_Click() '删除

    Application.ScreenUpdating = False
    r = Selection.Row
    
    With Me
        r1 = .Cells(Rows.count, "A").End(3).Row
        If Selection.Column < 9 Then
            .Range("B" & r & ":" & "H" & r).Delete Shift:=xlUp
            .Range("B4:H" & r1).Borders.LineStyle = 1
        ElseIf 9 < Selection.Column < 18 Then
            
            .Range("J" & r & ":" & "V" & r).Delete Shift:=xlUp
            .Range("J4:V" & r1).Borders.LineStyle = 1
        End If
        .sc.Visible = False
    End With
    Application.ScreenUpdating = True

End Sub

Private Sub up_Click() '上移
    Dim rng As Range
    Set rng = Selection
    r = rng.Row
    Application.ScreenUpdating = 0
    If rng.Rows.count <> 1 Then Exit Sub
    If r = 4 Then rng.Select: Exit Sub ' 如果选择区域的行号为4，则无法再向上移动，退出子程序

    With Me ' 引用当前窗体或工作表
        If Selection.Column < 9 Then ' 如果选择区域在第9列之前
            .Range("B" & r & ":" & "H" & r).Cut ' 将选择区域的内容剪切
            .Range("B" & r - 1 & ":" & "H" & r - 1).Insert Shift:=xlDown ' 在上方插入一行，并粘贴剪切的内容,其它行下移
            .Range("B" & r - 1 & ":" & "H" & r - 1).Select ' 选择插入的新行
        ElseIf 9 < Selection.Column < 18 Then ' 如果选择区域在第10列到第17列之间
            .Range("J" & r & ":" & "V" & r).Cut ' 将选择区域的内容剪切
            .Range("J" & r - 1 & ":" & "V" & r - 1).Insert Shift:=xlDown ' 在上方插入一行，并粘贴剪切的内容,其它行下移
            .Range("J" & r - 1 & ":" & "V" & r - 1).Select ' 选择插入的新行
        End If
    End With
    Application.ScreenUpdating = True
End Sub



Private Sub down_Click() '下移
Dim rng As Range
Set rng = Selection
    Application.ScreenUpdating = 0
    r = rng.Row
    If rng.Rows.count <> 1 Then Exit Sub
    If r + 1 > Me.UsedRange.Rows.count Then
        MsgBox "哥们别点了，到头了！"
        rng.Select
        Exit Sub
    End If
    With Me
        If Selection.Column < 9 Then
            If WorksheetFunction.CountA(.Range("B" & r + 1 & ":" & "H" & r + 1)) = 0 Then
                .Range("B" & r & ":" & "H" & r).Cut .Range("B" & r + 1 & ":" & "H" & r + 1)
                .Range("B4:H" & .UsedRange.Rows.count).Borders.LineStyle = 1
                .Range("B" & r + 1 & ":" & "H" & r + 1).Select
            Else
                .Range("B" & r & ":" & "H" & r).Cut
                .Range("B" & r + 2 & ":" & "H" & r + 2).Insert Shift:=xlDown '使用Insert时要保证置换区间大小一样
                .Range("B" & r + 1 & ":" & "H" & r + 1).Select
            End If
        ElseIf 9 < Selection.Column < 19 Then
            If WorksheetFunction.CountA(.Range("J" & r + 1 & ":" & "S" & r + 1)) = 0 Then
                .Range("J" & r & ":" & "V" & r).Cut .Range("J" & r + 1 & ":" & "V" & r + 1)
                .Range("J4:V" & .UsedRange.Rows.count).Borders.LineStyle = 1
                .Range("J" & r + 1 & ":" & "S" & r + 1).Select
            Else
                .Range("J" & r & ":" & "V" & r).Cut
                .Range("J" & r + 2 & ":" & "V" & r + 2).Insert Shift:=xlDown '使用Insert时要保证置换区间大小一样
                .Range("J" & r + 1 & ":" & "V" & r + 1).Select
            End If
        End If
    End With
    Application.ScreenUpdating = 1
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rngs As Range, sj As Date
    Application.EnableEvents = False
    M1 = Target.Row > 3
    If Target.Rows.count = 1 Then
        If Target.Columns.count = 1 Then M2 = Len(Target.value) > 4 Or val(Target.value) > 2359 Or Target.value = ""
    Else
        M2 = False
    End If
    M3 = Target.Row < Me.Cells(Me.Rows.count, "A").End(3).Row + 1
    M4 = Target.Column = 3 Or Target.Column = 18 Or Target.Column = 20 Or Target.Column = 21
    If M1 And M2 And M3 And M4 Then
        GoTo s
    ElseIf M4 And M3 And Target.Rows.count = 1 And Target.Columns.count = 1 Then
        '时间格式调整
        If InStr(Target.value, "：") > 0 Then Target.value = Replace(Target.value, "：", ":")

        ts = Format(Target.value, "0000")
        Target.NumberFormatLocal = "hh:mm"
        Target.HorizontalAlignment = xlCenter '设置上下居中
        Target.VerticalAlignment = xlCenter '设置左右居中
        On Error GoTo s
        sj = Left(ts, 2) & ":" & Right(ts, 2)
        Target.value = sj
        On Error GoTo 0
    Else
        If Target.Rows.count = 1 And Target.Columns.count = 1 Then Target.NumberFormatLocal = "G/通用格式"
    End If
    Application.EnableEvents = True
    Exit Sub
s:
    Target.value = ""
    Target.HorizontalAlignment = xlCenter '设置上下居中
    Target.VerticalAlignment = xlCenter '设置左右居中
    On Error GoTo 0
    Application.EnableEvents = True
End Sub
