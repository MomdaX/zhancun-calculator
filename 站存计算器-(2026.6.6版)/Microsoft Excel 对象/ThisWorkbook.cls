VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean) '自动保存
    Application.ScreenUpdating = False '关闭屏幕刷新
    Application.DisplayAlerts = False '关闭警告界面
    Application.AlertBeforeOverwriting = False '关闭提示框

    Sheet5.Range("K2") = "自动保存时间"
    Sheet5.Range("L2") = Now

    Application.ScreenUpdating = True '开启屏幕刷新
    Application.DisplayAlerts = True '开启警告界面
    Application.AlertBeforeOverwriting = True '开启提示框

    '保存工作簿
    ThisWorkbook.Save
End Sub


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    '另存为时
    If SaveAsUI Then
        Cancel = True
        Debug.Print "保存事件:" & Now()
        MsgBox "未授权，该功能无法使用！", vbOKOnly, "Momda-莫"
    Else
        'Cancel = True
        Debug.Print "保存事件:" & Now()
    End If
    
End Sub

Private Sub Workbook_Open()
    Dim 授权 As Boolean, t As Date, 权限 As Boolean, sht As Worksheet, srr
    Dim wn As String, qx As Long
    t = GetSetting("Momda", "站存计算器", "权限") '获取授权
    't = "2026/10/20"
    Application.ScreenUpdating = False '关闭屏幕刷新
    
    授权 = IIf(DateDiff("d", Date, t) > 0, True, False)
    srr = Sheet19.Range("M1").CurrentRegion
    If 授权 Then
        权限 = 授权
        For i = 2 To UBound(srr)
            wn = srr(i, 1)
            qx = srr(i, 3)
            Sheets(wn).Visible = qx
        Next i
    Else
        权限 = 授权
        For i = 2 To UBound(srr)
            wn = srr(i, 1)
            qx = srr(i, 2)
            Sheets(wn).Visible = qx
        Next i
        Application.ScreenUpdating = True '打开屏幕刷新
        MsgBox "授权：" & 权限 & Chr(10) & t, vbOKOnly, "权限"
    End If
    
    Application.ScreenUpdating = True '打开屏幕刷新
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub


Function sk() '用于设置权限

    Dim t As Date, 权限 As Boolean
    
    't = Sheet19.Range("Q6")
    t = "2026/10/20"
    SaveSetting "Momda", "站存计算器", "权限", t
    
    MsgBox "授权：" & "OK" & Chr(10) & t, vbOKOnly, "权限"
End Function

