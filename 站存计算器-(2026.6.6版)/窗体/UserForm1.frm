VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Momda-莫"
   ClientHeight    =   7128
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   8220
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton载入文件_Click()

    Dim filePath As String
    Application.ScreenUpdating = False '关闭屏幕刷新
    Application.DisplayAlerts = False '关闭警告界面
    '打开文件对话框
    With Application.fileDialog(msoFileDialogOpen)
        .InitialFileName = "C:\Program Files\SMIS2.6\TranData"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls*,*.xls*"
        .AllowMultiSelect = False
        If .Show = -1 Then
            filePath = .SelectedItems(1)
            '显示文件名
            Me.Label3.Caption = filePath
        End If
    End With
    
Dim wb As Workbook
Dim ws As Worksheet

'如果没有选择文件，路径为空时，结束程序
If filePath = "" Then Exit Sub
path = filePath

' 打开工作簿并将数据装入数组中
'Set wb = Workbooks.Open(filePath)
Set wb = GetObject(filePath)
Set ws = wb.ActiveSheet 's("第1页")

' 删除前两行数据
ws.Range("1:2").Delete Shift:=xlUp

' 装入数组中
arr = ws.Range("A1").CurrentRegion '模块1.arr

' 关闭工作簿
wb.Close False
Set wb = Nothing

Application.ScreenUpdating = True '开启屏幕刷新
Application.DisplayAlerts = True '开启警告界面
End Sub


Private Sub CommandButton开始计算_Click()
    
    Application.ScreenUpdating = False '取消屏幕刷新
    Application.DisplayAlerts = False '取消警告界面
        '如果没有选择文件
            If path = "" Then
                MsgBox ("请选择要计算站存的源文件")
                Exit Sub
            End If
            
        '当任何一个CheckBox控件的状态发生变化时，就会触发CheckBox_Click()事件，该事件会遍历窗体中所有的CheckBox控件，并检查它们的Value属性值是否为True。如果是True，则将它们的Caption属性值装入数组drr中。
        Dim chk As MSForms.CheckBox
        
        j = 0
        For i = 0 To Me.Controls.count - 1
            If TypeName(Me.Controls(i)) = "CheckBox" Then
                Set chk = Me.Controls(i)
                If chk.value = True Then
                    ReDim Preserve drr(j)
                    drr(j) = chk.Caption
                    j = j + 1
                End If
            End If
        Next i
    
    Call 设置车流属性.设置车流属性
    Call 计算.统计
    
    Me.TextBox站存.value = UBound(arr) - 1
    
    Debug.Print "当前时间：" & Time
    If Time < TimeValue("12:00:00") Then
    Sheet1.Range("A1").value = "钦州港 " & Month(Date) & " 月 " & Day(Date) & " 日 早6点 站存"
    Else
    Sheet1.Range("A1").value = "钦州港 " & Month(Date) & " 月 " & Day(Date) & " 日 18点 站存"
    End If
    
    Application.ScreenUpdating = True '开启屏幕刷新
    Application.DisplayAlerts = True '开启警告界面
    
    '延迟2秒
    Application.Wait (Now + TimeValue("0:00:1"))
    Unload Me
    
    Erase arr
    Erase drr
    Set d = Nothing
    Set gds = Nothing
    
    '保存工作簿
    'ThisWorkbook.Save

End Sub

Private Sub CommandButton3_Click()
    Unload Me
End Sub

Private Sub ListBox待发_Change()
    Dim i As Long
    Dim count As Long
    count = 0
    
    '计算ListBox1中选中项的个数
    For i = 0 To ListBox待发.ListCount - 1
        If ListBox待发.Selected(i) Then
            count = count + 1
        End If
    Next i
    
    If count <> 0 Then
    '将选中项的值存入数组crr中
        ReDim crr(1 To count)
        count = 0
        For i = 0 To ListBox待发.ListCount - 1
            If ListBox待发.Selected(i) Then
                count = count + 1
                crr(count) = ListBox待发.list(i)
            End If
        Next i
    End If
End Sub

Private Sub ToggleButton超智_Click()
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.Caption = "CZ3" Or ctrl.Caption = "CZ4" Then
                ctrl.value = Me.ToggleButton超智.value
                '设置颜色
                设置颜色 Me.ToggleButton超智
            End If
        End If
    Next ctrl
End Sub

Private Sub ToggleButton广明_Click()
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.Caption = "GM1" Or ctrl.Caption = "GM2" Or ctrl.Caption = "GM3" Or ctrl.Caption = "GM4" Then
                ctrl.value = Me.ToggleButton广明.value
                '设置颜色
                设置颜色 ToggleButton广明
            End If
        End If
    Next ctrl
End Sub

Private Sub ToggleButton货场_Click()
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.Caption = "H1" Or ctrl.Caption = "H2" Or ctrl.Caption = "H3" Or ctrl.Caption = "H4" Or ctrl.Caption = "H5" Then
                ctrl.value = Me.ToggleButton货场.value
                '设置颜色
                设置颜色 ToggleButton货场
            End If
        End If
    Next ctrl
End Sub


Private Sub ToggleButton大洋_Click()
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.Caption = "LZ" Then
                ctrl.value = Me.ToggleButton大洋.value
                '设置颜色
                设置颜色 ToggleButton大洋
            End If
        End If
    Next ctrl
End Sub

Private Sub ToggleButton港务局_Click()
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.Caption = "G1" Or ctrl.Caption = "G2" Then
                ctrl.value = Me.ToggleButton港务局.value
                '设置颜色
                设置颜色 ToggleButton港务局
            End If
        End If
    Next ctrl
End Sub

Private Sub ToggleButton国投_Click()
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.Caption = "GT1" Or ctrl.Caption = "GT2" Then
                ctrl.value = Me.ToggleButton国投.value
                '设置颜色
                设置颜色 ToggleButton国投
            End If
        End If
    Next ctrl
End Sub

Private Sub ToggleButton默认待装_Click()
ToggleButton中粮.value = Me.ToggleButton默认待装.value
ToggleButton货场.value = Me.ToggleButton默认待装.value
ToggleButton天煤.value = Me.ToggleButton默认待装.value
ToggleButton永鑫.value = Me.ToggleButton默认待装.value
ToggleButton石化.value = Me.ToggleButton默认待装.value
ToggleButton国投.value = Me.ToggleButton默认待装.value
ToggleButton中油.value = Me.ToggleButton默认待装.value
ToggleButton港务局.value = Me.ToggleButton默认待装.value

'    If ToggleButton默认待装.Value = True Then
'        ToggleButton默认待装.BackColor = RGB(0, 0, 255) '红色,黄色，蓝色
'    Else
'        ToggleButton默认待装.BackColor = RGB(255, 255, 255) '白色
'    End If
    
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is ToggleButton Then
            If ctrl.value = True And ctrl.Caption <> "默认待装" Then
                ctrl.BackColor = RGB(255, 0, 0)  '红色,黄色，蓝色
            Else
                ctrl.BackColor = RGB(255, 255, 255) '白色
            End If
        End If
    Next ctrl
End Sub
Function 设置颜色(按钮 As Control)

    If 按钮 Then
        按钮.BackColor = RGB(255, 0, 0)  '红色
    Else
        按钮.BackColor = RGB(255, 255, 255) '白色
    End If
End Function

Private Sub ToggleButton石化_Click()
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.Caption = "SH1" Or ctrl.Caption = "SH2" Then
                ctrl.value = Me.ToggleButton石化.value
                '设置颜色
                设置颜色 ToggleButton石化
            End If
        End If
    Next ctrl

End Sub


Private Sub ToggleButton天煤_Click()
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.Caption = "TS1" Or ctrl.Caption = "TS2" Then
                ctrl.value = Me.ToggleButton天煤.value
                '设置颜色
                设置颜色 ToggleButton天煤
            End If
        End If
    Next ctrl
End Sub

Private Sub ToggleButton天油_Click()
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.Caption = "TSY1" Or ctrl.Caption = "TSY2" Or ctrl.Caption = "TSY3" Or ctrl.Caption = "TSY4" Then
                ctrl.value = Me.ToggleButton天油.value
                '设置颜色
                设置颜色 ToggleButton天油
            End If
        End If
    Next ctrl
End Sub

Private Sub ToggleButton永鑫_Click()
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.Caption = "YX2" Or ctrl.Caption = "YX3" Then
                ctrl.value = Me.ToggleButton永鑫.value
                '设置颜色
                设置颜色 ToggleButton永鑫
            End If
        End If
    Next ctrl
End Sub

Private Sub ToggleButton中粮_Click()

    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.Caption = "ZL1" Or ctrl.Caption = "ZL2" Or ctrl.Caption = "ZL3" Then
                ctrl.value = Me.ToggleButton中粮.value
                '设置颜色
                设置颜色 ToggleButton中粮
            End If
        End If
    Next ctrl

End Sub

Private Sub ToggleButton中油_Click()
    Dim ctrl As Control
    
    For Each ctrl In Me.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.Caption = "Y5" Or ctrl.Caption = "Y6" Or ctrl.Caption = "Y9" Or ctrl.Caption = "Y10" Or ctrl.Caption = "Y11" Or ctrl.Caption = "Y12" Or ctrl.Caption = "Y13" Or ctrl.Caption = "Y14" Or ctrl.Caption = "Y15" Or ctrl.Caption = "Y16" Then
                ctrl.value = Me.ToggleButton中油.value
                '设置颜色
                设置颜色 ToggleButton中油
            End If
        End If
    Next ctrl
End Sub

Private Sub UserForm_Initialize()
    path = ""
    '向左侧ListBox添加带复选框的项
    '实际场
    For i = 1 To 15
        Me.ListBox待发.AddItem i
    Next i
    
    '虚拟场
    For i = 1 To 10
        Me.ListBox待发.AddItem "X" & i
    Next i

   '默认待装
   ToggleButton默认待装.value = False
End Sub
