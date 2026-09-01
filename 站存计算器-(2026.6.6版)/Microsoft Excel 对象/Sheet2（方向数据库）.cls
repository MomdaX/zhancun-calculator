VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()

    aa = MsgBox("是否执行更新到站信息！", vbYesNo, "到站数据库")
    If aa = 6 Then
        更新.刷新
    End If
End Sub
