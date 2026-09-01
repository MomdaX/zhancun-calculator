VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CommandButton_港务局_Click()
eet7.Range("A3:Q3").AutoFilter 1, Array("G1", "G2"), xlFilterValues
End Sub

Private Sub CommandButton_货场_Click()
Sheet7.Range("A3:Q3").AutoFilter 1, Array("H1", "H2", "H3", "H4", "H5"), xlFilterValues
End Sub

Private Sub CommandButton_中粮大洋_Click()
Sheet7.Range("A3:Q3").AutoFilter 1, Array("ZL1", "ZL2", "ZL3", "LZ"), xlFilterValues
End Sub

Private Sub CommandButton_YX_Click()
Sheet7.Range("A3:Q3").AutoFilter 1, Array("YX2", "YX3"), xlFilterValues
End Sub

Private Sub CommandButton_TS_Click()
Sheet7.Range("A3:Q3").AutoFilter 1, Array("TS1", "TS2"), xlFilterValues
End Sub

Private Sub CommandButton_GT_Click()
Sheet7.Range("A3:Q3").AutoFilter 1, Array("GT1", "GT2"), xlFilterValues
End Sub

Private Sub CommandButton_SH_Click()
Sheet7.Range("A3:Q3").AutoFilter 1, Array("SH1", "SH2", "SH3"), xlFilterValues
End Sub

Private Sub CommandButton_CZ_Click()
Sheet7.Range("A3:Q3").AutoFilter 1, Array("CZ3", "CZ4"), xlFilterValues
End Sub

Private Sub CommandButton_GM_Click()
Sheet7.Range("A3:Q3").AutoFilter 1, Array("GM1", "GM2", "GM3", "GM4"), xlFilterValues
End Sub

Private Sub CommandButton_敞顶箱_Click()
Sheet7.Range("A3:Q3").AutoFilter Field:=14, Criteria1:="=*空箱*", Operator:=xlOr, Criteria2:="=*敞顶箱*"
End Sub

Private Sub CommandButton1_Click()
    Dim rng As Range, rn As Range, 辆数 As Double, ls As Double, hc As Double, 换长 As Double
    Set rng = Selection
    For Each rn In rng
        ls = ls + 1
        hc = hc + rn.value
    Next
    换长 = Split(Me.Label1.Caption, "/")(0) + hc
    辆数 = Split(Me.Label1.Caption, "/")(1) + ls
    Me.Label1.Caption = 换长 & "/" & 辆数
    Me.Label1.BackColor = IIf(换长 > 70, 8421631, 8454016)
    rng.Select
End Sub

Private Sub CommandButton2_Click()
    Dim rng As Range, rn As Range, 辆数 As Double, ls As Double, hc As Double, 换长 As Double
    Set rng = Selection
    For Each rn In rng
        ls = ls + 1
        hc = hc + rn.value
    Next
    换长 = Split(Me.Label1.Caption, "/")(0) - hc
    辆数 = Split(Me.Label1.Caption, "/")(1) - ls
    Me.Label1.Caption = 换长 & "/" & 辆数
    Me.Label1.BackColor = IIf(换长 > 70, 8421631, 8454016)
    rng.Select
End Sub

Private Sub CommandButton3_Click()
    Dim rng As Range, rngs As Range
    With Sheet7
        Set rng = .Range(.Cells(Selection.Row, "D"), .Cells(Selection.Row + Selection.Rows.count - 1, "D"))
        If rng.Column <> 4 Or rng.Row < 4 Or rng.Row > .Cells(Rows.count, 4).End(3).Row Then
            MsgBox "请选择车号的范围！"
            End
        End If
        Set rngs = .Range("N" & Selection.Row).Resize(rng.Rows.count, 1)
    End With
    到卸查询.查询卸车专用线 rng, rngs
End Sub

Private Sub CommandButton4_Click()
    显示信息.查询载重
End Sub

Private Sub CommandButton5_Click()
    Dim rng As Range, rn As Range, 车号 As String
    Application.EnableEvents = False
    车号 = InputBox("请输入车号!")
'    Set rng = Me.Columns("D")
'    Set rn = rng.Find(车号)
'
'    If Not rn Is Nothing Then
'        rn.EntireRow.Select
'        With Me.CommandButton6
'            .Width = 50
'            .Height = 30
'            .Left = rn.Offset(0, 14).Left
'            .Top = rn.Offset(0, 14).Top
'            .Caption = "返回"
''            .Interior.ColorIndex = 3
'            .Visible = True
'        End With
'    Else
'        Me.CommandButton6.Visible = False
'        Debug.Print "未找到：" & 车号
'    End If
    
    Me.Range("$A$3:$Q$3").AutoFilter Field:=4, Criteria1:="=*" & 车号 & "*", Operator:=xlAnd
    
    
    
    Application.EnableEvents = True
End Sub

Private Sub CommandButton6_Click()
    Sheet6.Select
    With Sheet7.Range("O3")
        .value = "卸车地点"
        .Interior.Color = 16777215
    End With
'    Application.EnableEvents = False
'        Me.Rows(3).EntireRow.Select
'    Application.EnableEvents = True
End Sub

Private Sub CommandButton7_Click()
    查询站内车的到卸和变更.变更查询
End Sub

Private Sub Label1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim rng As Range
    Set rng = Selection
    Me.Label1.Caption = "0/0"
    Me.Label2.Visible = False
    rng.Select
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim 文件名 As String
    If Target.count = 1 Then
        If Target.Column = 14 And Not Target.Comment Is Nothing Then
            车号 = Me.Cells(Target.Row, "D").value
            显示信息.读取票据 车号
        End If
    End If
    Cancel = True ' 取消单元格编辑模式
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim rng As Range, rn As Range, ls As Double, hc As Double
    If Target.Column = 6 And WorksheetFunction.CountA(Target) > 0 Then
        Set rng = Selection
        For Each rn In rng
            If IsNumeric(rn.value) Then
                ls = ls + 1
                hc = hc + rn.value
            End If
        Next
        With Me.Label1
            .Left = Target.Cells(1).Offset(0, 1).Left
            .Top = Target.Cells(1).Offset(0, 1).Top
            .TextAlign = fmTextAlignCenter
            .BorderStyle = fmBorderStyleSingle
            .Font.Size = 20
            .width = 100
            .height = 26
            .Visible = True
        End With
        With Me.Label2
            .Left = Me.Label1.Left
            .Top = Me.Label1.Top + Me.Label1.height + 1
            .TextAlign = fmTextAlignCenter
            .BorderStyle = fmBorderStyleSingle
            .Font.Size = 20
            .width = 100
            .height = 26
            .Caption = hc & "/" & ls
            .Visible = True
        End With
        With Me.CommandButton1
            .Left = Me.Label1.Left + Me.Label1.width
            .Top = Me.Label1.Top
            .Caption = "+"
            .Visible = True
        End With
        With Me.CommandButton2
            .Left = Me.Label1.Left + Me.Label1.width
            .Top = Me.Label1.Top + Me.CommandButton1.height + 1
            .Caption = "-"
            .Visible = True
        End With
    Else
    With Me
        .Label1.Visible = False
        .Label2.Visible = False
        .CommandButton1.Visible = False
        .CommandButton2.Visible = False
        .CommandButton6.Visible = True
    End With
    End If
End Sub
