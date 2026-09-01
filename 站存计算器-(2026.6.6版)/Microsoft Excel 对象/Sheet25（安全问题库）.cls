VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    
    If Target.address(0, 0) = "R1" Then
        Cancel = True ' 防止 Excel 默认执行双击编辑操作
        OA安全库.获取
        With Target
            .value = "刷新"
            .Borders.LineStyle = xlContinuous '设置样式边框
            .HorizontalAlignment = xlCenter '设置上下居中
            .VerticalAlignment = xlCenter '设置左右居中
            .Interior.ColorIndex = 14
            .Font.ColorIndex = 2
            .Font.Bold = True
        End With
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim rng As Range
    Application.ScreenUpdating = False
    On Error Resume Next
    With Me
        Set rng = .Range("A1").CurrentRegion
        If Not Intersect(Target, rng) Is Nothing Then
            .Cells.WrapText = False
            With rng
                .Offset(1).Font.Size = 10
                .Offset(1).Interior.ColorIndex = 0
                With .Rows(Target.Row)
                    .Font.Size = 12
                    .Interior.ColorIndex = 8
                    Me.Cells(Target.Row, 7).Font.Size = 14
                End With
            End With
            .Cells(Target.Row, 7).WrapText = True
        ElseIf Target.Column = 14 And Target.Cells(1) <> "" And Not Target.Row = 1 Then
            If .AutoFilterMode Then
                .AutoFilterMode = False
            End If
            rng.AutoFilter 9, Target.value ', xlFilterValues
'            OA安全库.SQL筛选 (Target.value)
        Else
            .Cells.WrapText = False
            With .Range("A1").CurrentRegion.Offset(1)
                .Font.Size = 10
                .Interior.ColorIndex = 0
            End With
            If .AutoFilterMode Then
                .AutoFilterMode = False
            End If
            .Range("A1").CurrentRegion.Rows(1).AutoFilter
        End If
        With rng.Rows(1)
            .Font.Size = 14
            .Font.Bold = True
            .Font.ColorIndex = 2
            .Interior.ColorIndex = 14
        End With
    End With
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub
