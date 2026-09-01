VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim btn As Object ' 按钮对象
    Set btn = Me.da
    
    ' 隐藏按钮
    Me.da.Visible = False
    
    ' 判断选中区域是否为单个单元格,选中罐车结存时显示罐车结存按钮
    If Target.Rows.count = ActiveSheet.Rows.count Or Target.Columns.count = ActiveSheet.Columns.count Then End

    ' 判断选中区域是否在第1~83行
    rs = Sheet30.Cells(6, "R").End(4).Row
    If Target.Row < 6 Or rs < Target.Row Then Exit Sub
    
    ' 判断选中区域是否在第Q~X列
    If Target.Column < 17 Or Target.Column > 24 Or Cells(Target.Row, 24).value = "" Then Exit Sub
    
    '到达按钮
    With btn
        .Left = Me.Range("X" & Target.Row).Left + Target.width - 37
        .Top = Me.Range("X" & Target.Row + 1).Top
        .width = 36
        .height = Target.RowHeight + 10
        .Visible = True
    End With
End Sub

Private Sub da_Click()
    r = Selection.Row
    srr = Sheet30.Range("R" & r & ":X" & r)
    车次 = srr(1, 1)
    到达时间 = Format(srr(1, 2), "hh:mm")
    到站 = srr(1, 3)
    车种 = srr(1, 4)
    车数 = srr(1, 5)
    换长 = srr(1, 6)
    编组 = srr(1, 7)
    批注 = Range("X" & r).Comment.text
    
    '用途=手动填
    合并 = 车次 & "$" & 到达时间 & "$" & 到站 & "$" & 车种 & "$" & 车数 & "$" & 换长 & "$" & 编组
    
    r = Sheet4.Cells(15, "B").End(3).Row + 1
    srr = Split(合并, "$")
    With Sheet4
        .Range("B" & r).Resize(1, UBound(srr) + 1) = srr
        .Range("H" & r).AddComment 批注
    End With
    
End Sub

'双击执行
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim 车次 As Boolean, pk As String, rng As Range
    Cancel = True
    车次 = Me.Range("R3").value
    Set rng = Me.Range("R5").CurrentRegion.Offset
    Set rn = Application.Intersect(Target, rng)
    If Target.address(0, 0) = "S3" Then
        Me.da.Visible = False '隐藏到达按钮
        If 车次 Then
            pk = 列车到达计划.查询转发运统一(Me.Range("R3").value)
            If pk = "" Then
                MsgBox "请核对列车车次" & Chr(10) & "无编组内容！！！"
            Else
                If 违编核对.获取编组(pk) Then
                    Call 列车到达计划.变更查询
                End If
            End If
        Else
            列车到达计划.查询转发运统一
        End If
    ElseIf Target.address(0, 0) = "F2" Then '成功
        With Sheet30
            .Range("F2").Interior.Color = 49407 '黄
            .Range("F3").value = val(GetSetting("Momda", "login", "计数"))
            .Range("F2").Interior.Color = 16777215 '无
        End With
    ElseIf Target.address(0, 0) = "G2" Then '失败
        With Sheet30
            .Range("G2").Interior.Color = 49407 '黄
            .Range("G3").value = val(GetSetting("Momda", "login", "失败"))
            .Range("G2").Interior.Color = 16777215 '无
        End With
    ElseIf Not rn Is Nothing Then '读取编组
        If Target.Column = 24 And Target.Row > 5 And Target.value <> "" Then
            pk = Split(Target.Comment.text, ":")(1)
            违编核对.获取编组 pk
            Call 列车到达计划.变更查询
        Else
            Cancel = False
        End If
    End If
    
End Sub

