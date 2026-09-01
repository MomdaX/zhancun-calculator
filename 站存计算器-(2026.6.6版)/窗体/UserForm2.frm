VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "选择要查询的车种类型"
   ClientHeight    =   1956
   ClientLeft      =   48
   ClientTop       =   156
   ClientWidth     =   4524
   OleObjectBlob   =   "UserForm2.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function 车种筛选(车种 As String)
Application.ScreenUpdating = False '取消屏幕刷新
Application.DisplayAlerts = False '取消警告界面
    Sheet6.Range("H2:I83").ClearContents
    Call 统计股道存车.重空车按车种分类(车种)
    Label2.Caption = CommandButton_C.Caption
    
    '筛选非空
        Sheet6.AutoFilterMode = False
'        With Sheet6.Range("A1").CurrentRegion
'            .AutoFilter Field:=8, Criteria1:="<>"
'            .AutoFilter Field:=9, Criteria1:="<>"
'        End With
        
    ' 关闭窗体
    UserForm2.Hide
Application.ScreenUpdating = True '启用屏幕刷新
Application.DisplayAlerts = True '启用警告界面
End Function

Private Sub CommandButton_C_Click()
车种筛选 CommandButton_C.Caption
End Sub

Private Sub CommandButton_X_Click()
车种筛选 CommandButton_X.Caption
End Sub

Private Sub CommandButton_P_Click()
车种筛选 CommandButton_P.Caption
End Sub

Private Sub CommandButton_G_Click()
车种筛选 CommandButton_G.Caption
End Sub

Private Sub UserForm_Click()

End Sub
