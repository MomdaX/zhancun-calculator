VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CommandButton_罐车结存_Click()
    Dim rn As Range, reg As New RegExp
    Me.CommandButton_罐车结存.Visible = False
    For Each rn In [G2:G92]
        If rn.value <> "" Then
            Dim 自备罐%, 路罐%, g%
            reg.Pattern = "\d+"
'            mmk = Split(rn.Value, " ")
            For Each tx In Split(rn.value, " ")
                If InStr(tx, "自备罐") > 0 And tx <> "" Then
                    g = reg.Execute(tx)(0)
                    自备罐 = 自备罐 + g
                End If
                If InStr(tx, "路罐") > 0 And tx <> "" Then
                    g = reg.Execute(tx)(0)
                    路罐 = 路罐 + g
                End If
                g = 0
            Next
        End If
    Next
    Sheet6.Range("G93").value = 自备罐 & "(自)/" & 路罐 & "(路)"
End Sub

Private Sub CommandButton_结存按钮_Click()
    Dim t1%, t2%, 路罐%, 自备罐%
    Dim reg As New RegExp: reg.Pattern = "\d+": reg.Global = True
    Set rng = Selection
    If InStr(rng.value, "G") > 0 Then
        '罐车
        If reg.Execute(rng.value).count > 1 Then
'            Set mmk = reg.Execute(rng.Value)
            自备罐 = reg.Execute(rng.value)(0)
             路罐 = reg.Execute(rng.value)(1)
        Else
            If InStr(rng.value, "自") > 0 Then
                自备罐 = reg.Execute(rng.value)(0)
            Else
                路罐 = reg.Execute(rng.value)(0)
            End If
        End If
        If InStr(Me.TextBox_结存文本框.text, "/") > 0 Then
            t1 = val(Split(Me.TextBox_结存文本框.text, "/")(0)) + 自备罐
            t2 = val(Split(Me.TextBox_结存文本框.text, "/")(1)) + 路罐
            Me.TextBox_结存文本框.text = t1 & "/" & t2
        Else
            t1 = 自备罐
            t2 = 路罐
            Me.TextBox_结存文本框.text = t1 & "/" & t2
        End If
    Else
        '非罐车
        If InStr(Me.TextBox_结存文本框.text, "/") > 0 Then
            Me.TextBox_结存文本框.text = reg.Execute(rng.value)(0)
        Else
            hj = reg.Execute(rng.value)(0)
            Me.TextBox_结存文本框.text = val(Me.TextBox_结存文本框.text) + hj
        End If
    End If
End Sub

Private Sub CommandButton_结存按钮_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.TextBox_结存文本框.text = ""
End Sub

Private Sub CommandButton_自备罐_Click()
    Dim t1%, t2%, 路罐%, 自备罐%
    Dim reg As New RegExp: reg.Pattern = "\d+": reg.Global = True
    Set rng = Selection
    If InStr(rng.value, "G") > 0 Then
        If InStr(Me.TextBox_结存文本框.text, "/") > 0 Then
            t1 = val(Split(Me.TextBox_结存文本框.text, "/")(0)) + val(reg.Execute(rng.value)(0))
            t2 = val(Split(Me.TextBox_结存文本框.text, "/")(1))
            Me.TextBox_结存文本框.text = t1 & "/" & t2
        Else
            t1 = 自备罐 + val(reg.Execute(rng.value))
            t2 = 路罐
            Me.TextBox_结存文本框.text = t1 & "/" & t2
        End If
    Else
        hj = reg.Execute(rng.value)(0)
        Me.TextBox_结存文本框.text = val(Me.TextBox_结存文本框.text) + hj
    End If
End Sub

Private Sub CommandButton_路罐_Click()
    Dim t1%, t2%, 路罐%, 自备罐%
    Dim reg As New RegExp: reg.Pattern = "\d+": reg.Global = True
    Set rng = Selection
    If InStr(rng.value, "G") > 0 Then
        If InStr(Me.TextBox_结存文本框.text, "/") > 0 Then
            t1 = val(Split(Me.TextBox_结存文本框.text, "/")(0))
            t2 = val(Split(Me.TextBox_结存文本框.text, "/")(1)) + val(reg.Execute(rng.value)(0))
            Me.TextBox_结存文本框.text = t1 & "/" & t2
        Else
            t1 = 自备罐
            t2 = 路罐 + val(reg.Execute(rng.value)(0))
            Me.TextBox_结存文本框.text = t1 & "/" & t2
        End If
    Else
        hj = reg.Execute(rng.value)(0)
        Me.TextBox_结存文本框.text = val(Me.TextBox_结存文本框.text) + hj
    End If
    
End Sub

Private Sub SpinButton_调节按钮_SpinDown()
    If InStr(Me.TextBox_结存文本框.text, "/") = 0 Then
        Me.TextBox_结存文本框.text = val(Me.TextBox_结存文本框.text) - 1
    End If
End Sub

Private Sub SpinButton_调节按钮_SpinUp()
    If InStr(Me.TextBox_结存文本框.text, "/") = 0 Then
        Me.TextBox_结存文本框.text = val(Me.TextBox_结存文本框.text) + 1
    End If
End Sub

'Private Sub Worksheet_Change(ByVal Target As Range)
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
Dim rng As Range
    M1 = Target.Column = 2 Or Target.Column = 7
    M2 = Target.value <> ""
    If M1 And M2 Then
        Cancel = True ' 防止 Excel 默认执行双击编辑操作
        Sheet7.Select
        Sheet7.Range("A3:Q3").AutoFilter 1, Sheet6.Cells(Target.Row, "B").value 'Target.Value
    End If
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim btn As Object ' 按钮对象

    ' 隐藏按钮
    Set btn_bh = Me.bh
    Set btn_da = Me.da
    Set btn_jc = Me.CommandButton_结存按钮
    Set btn_jct = Me.TextBox_结存文本框
    Set btn_lg = Me.CommandButton_路罐
    Set btn_zbg = Me.CommandButton_自备罐
    Set btn_tj = SpinButton_调节按钮
    btn_bh.Visible = False
    btn_da.Visible = False
    btn_jc.Visible = False
    btn_jct.Visible = False
    btn_lg.Visible = False
    btn_zbg.Visible = False
    btn_tj.Visible = False
    Me.CommandButton_罐车结存.Visible = False
    
    ' 判断选中区域是否为单个单元格,选中罐车结存时显示罐车结存按钮
    If Target.Rows.count = ActiveSheet.Rows.count Or Target.Columns.count = ActiveSheet.Columns.count Then End
    If Target.Cells.count = 1 And Target.address = "$G$93" Then
        With Me.CommandButton_罐车结存
            .Left = Target.Left
            .Top = Target.Top
            .width = 60
            .height = Target.height
            .Visible = True
        End With
    ElseIf Target.Cells.count <> 1 Then
        Exit Sub
    End If
    
    ' 判断选中区域是否在第1~83行
    If Target.Row < 2 Or 85 < Target.Row Then Exit Sub
    
    ' 判断选中区域是否在第K列
    If Target.Column > 11 Or Cells(Target.Row, 7).value = "" Then Exit Sub
    
    ' 设置按钮位置
    With btn_bh '编好
        .Left = Me.Range("L" & Target.Row).Left
        .Top = Me.Range("L" & Target.Row).Top
        .width = 50
        .height = 30
        .Visible = True
    End With
    With btn_da '到达
        .Left = Me.Range("L" & Target.Row).Left + 50
        .Top = Me.Range("L" & Target.Row).Top
        .width = 50
        .height = 30
        .Visible = True
    End With
    
    If (Target.Column = 8 Or Target.Column = 9) And Target.value <> "" Then
        With btn_jc '结存按钮
            .Left = Target.Offset(0, 1).Left
            .Top = Target.Offset(0, 1).Top - 5
            .width = 30
            .height = 25
            .Visible = True
        End With
        With btn_jct '结存文本框
            .Left = Target.Offset(0, 1).Left + btn_jc.width
            .Top = Target.Offset(0, 1).Top - 5
            .width = 72
            .height = 25
            .Visible = True
        End With
        With btn_lg '路罐
            .Left = Target.Offset(0, 1).Left + btn_jc.width + btn_jct.width
            .Top = Target.Offset(0, 1).Top + 7
            .width = 15
            .height = 12.5
            .Visible = True
        End With
        With btn_zbg '自备罐
            .Left = Target.Offset(0, 1).Left + btn_jc.width + btn_jct.width
            .Top = Target.Offset(0, 1).Top - 6
            .width = 15
            .height = 12.5
            .Visible = True
        End With
        With btn_tj '调节按钮
            .Left = Target.Offset(0, 1).Left + btn_jc.width + btn_jct.width + btn_lg.width
            .Top = Target.Offset(0, 1).Top - 6.5
            .width = 20
            .height = 25
            .Visible = True
        End With
        
    End If
End Sub

Private Sub bh_Click()
r = Selection.Row
srr = Sheet6.Range("A" & r & ":K" & r)
车次 = srr(1, 1)
方向 = srr(1, 3)
发车股道 = srr(1, 2)
车种 = srr(1, 5)
车数 = srr(1, 4)
换长 = srr(1, 6)
到站 = srr(1, 7)
预计编好时间 = srr(1, 11)
合并 = 车次 & "$" & 方向 & "$" & 发车股道 & "$" & 车种 & "$" & 车数 & "$" & 换长 & "$" & 到站 & "$" & 预计编好时间

r = Sheet4.Cells(15, "L").End(3).Row + 1
srr = Split(合并, "$")
Sheet4.Range("J" & r).Resize(1, UBound(srr) + 1) = srr

Dim rng As Range
Set rng = Sheet4.Range("P4:P16")
Call 显示信息.标记到站方向颜色(rng)
End Sub
Private Sub da_Click()
    r = Selection.Row
    srr = Sheet6.Range("A" & r & ":K" & r)
    车次 = srr(1, 10)
    '到达时间=手动填
    接车股道 = srr(1, 2)
    车种 = srr(1, 5)
    车数 = srr(1, 4)
    换长 = srr(1, 6)
    '用途=手动填
    
    合并 = 车次 & "$" & 到达时间 & "$" & 接车股道 & "$" & 车种 & "$" & 车数 & "$" & 换长
    
    r = Sheet4.Cells(15, "B").End(3).Row + 1
    srr = Split(合并, "$")
    Sheet4.Range("B" & r).Resize(1, UBound(srr) + 1) = srr

End Sub
