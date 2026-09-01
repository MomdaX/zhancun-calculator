VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm4 
   Caption         =   "验证码"
   ClientHeight    =   2580
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   2208
   OleObjectBlob   =   "UserForm4.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
    UserForm4.Caption = UserForm4.TextBox1.value
End Sub
