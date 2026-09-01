VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim t As Date, 权限 As Boolean
    If Target(1).address(0, 0) = "Q6" Then
        t = Target(1).value
        SaveSetting "Momda", "站存计算器", "权限", t
        Cancel = True
        MsgBox "授权：" & "OK" & Chr(10) & t, vbOKOnly, "权限"
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    If Not Intersect(Target, Me.Range("E2:E3")) Is Nothing And Me.Range("E3") <> "" Then
    
    
        Call 查询用户.查询用户
        
        
    ElseIf Me.Range("E3") = "" And Me.Range("E2") = "" Then

        Me.Range("E1") = "请输入帐号密码！"
        
    ElseIf Me.Range("E2") <> "" And Me.Range("E3") = "" Then

        Me.Range("E1") = "请输入密码！"
        
    End If
End Sub

