VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    到卸查询.查询卸车专用线
    设置按钮
End Sub

Private Sub CommandButton_历装_Click()
    Dim rng As Range, rngs1 As Range, rngs2 As Range, rn As Range, r%, c%
    r = Selection.Row
    
    If r = 1 Then GoTo 选择车辆

    c = Selection.Rows.count - 1
    Me.Columns("O").ClearComments
    Set rn = Me.Range(Me.Cells(r, "A"), Me.Cells(r + c, "A"))
    
    If Me.Range("A1").value <> "" Then
        Set rng = Me.Range("A1").CurrentRegion
    End If
    Set rngs1 = Intersect(rn, rng)
    Set rngs2 = Intersect(Selection, rng)
    If Not rngs1 Is Nothing And Not rngs2 Is Nothing Then

        Call 查询卸线.查询货装历史(rn)
    Else
        GoTo 选择车辆

    End If
    
    Exit Sub
    
选择车辆:
    MsgBox "请选择要查询的车辆！"
End Sub

Private Sub CommandButton_预报查询_Click()
    Dim rng As Range, rngs As Range
    With Sheet13
        Set rng = .Range("A2:A" & Sheet13.Cells(Sheet13.Rows.count, "A").End(3).Row)
        Set rngs = .Range("P2").Resize(rng.Rows.count, 3)
    End With
    到卸查询.查询卸车专用线 rng, rngs
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim 文件名 As String
    If Target.count = 1 Then
        If Target.Column = 6 Or Target.Column = 16 Then
            车号 = Sheet13.Cells(Target.Row, "A").value
            显示信息.读取票据 车号
'            If Dir(文件名, vbDirectory) <> "" Then
'                ' 取消单元格编辑模式
'                Cancel = True
'                UserForm3.Show 0
'            Else
'                MsgBox "查无票据：" & Sheet13.Cells(Target.Row, "A").value
'            End If
        End If
    End If
    Cancel = True ' 取消单元格编辑模式
    Application.EnableEvents = True
End Sub
Private Sub CommandButton_查询到卸_Click()
    查询卸线.查询卸车地点
    设置按钮
End Sub

Private Sub CommandButton_结束_Click()
    设置按钮
    Application.EnableEvents = True
    End
End Sub

Private Sub CommandButton_提取编组_Click()
    查询卸线.查询编组
    设置按钮
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    With Me.Range("A:R")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .EntireColumn.AutoFit
    End With
End Sub

Sub 设置按钮()
With Me.CommandButton_提取编组
    .Top = Me.Range("T3").Top
    .Left = Me.Range("T3").Left
    .width = 100
    .height = 50
End With

With Me.CommandButton_查询到卸
    .Top = Me.CommandButton_提取编组.Top + Me.CommandButton_提取编组.height + 1
    .Left = Me.Range("T3").Left
    .width = 100
    .height = 50
End With

With Me.CommandButton_预报查询
    .Top = Me.CommandButton_查询到卸.Top + Me.CommandButton_查询到卸.height + 1
    .Left = Me.Range("T3").Left
    .width = 100
    .height = 50
End With

With Me.CommandButton_结束
    .Top = Me.CommandButton_预报查询.Top + Me.CommandButton_预报查询.height + 1
    .Left = Me.Range("T3").Left
    .width = 100
    .height = 50
End With

With Me.CommandButton_历装
    .Caption = "装车记录"
    .Top = Me.CommandButton_结束.Top + Me.CommandButton_结束.height + 1
    .Left = Me.Range("T3").Left
    .width = 100
    .height = 50
End With


End Sub

'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'    If Not Target.Comment Is Nothing Then
'        Target.Comment.Visible = False
'    End If
'End Sub
