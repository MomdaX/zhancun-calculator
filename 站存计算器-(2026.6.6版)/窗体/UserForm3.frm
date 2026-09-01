VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm2"
   ClientHeight    =   8772
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   12936
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()

    Do While WebBrowser1.Busy
        DoEvents
    Loop

    ' 设置 WebBrowser 控件位置和大小
    With WebBrowser1
'        .navigate "about:blank"
        .navigate ThisWorkbook.path & "\票据缓存\" & Sheet13.Cells(Selection.Row, "A").value & ".html" '文件
        .Visible = True
        .width = 700
        .height = 550
    End With
    
    UserForm3.width = 700
    UserForm3.height = 550
    UserForm3.Caption = "运单号码：" & Sheet13.Cells(Selection.Row, "A").value
    
    ' 等待 WebBrowser 控件加载完毕
    Do While WebBrowser1.Busy Or WebBrowser1.readyState <> 4
        DoEvents
    Loop
    
    '滑动到最底部
    On Error Resume Next
    WebBrowser1.document.parentWindow.scrollTo 0, WebBrowser1.document.Body.ScrollHeight
    On Error GoTo 0
'    ' 加载 HTML 代码
''    WebBrowser1.document.body.innerHTML = Selection.Comment.Text '代码段
'     WebBrowser1.navigate "F:\Momda\货票缓存\" & Sheet13.Cells(Selection.Row, "F").Value & ".html" '文件
'    ' 等待用户关闭 WebBrowser 控件
'    Do While WebBrowser1.Visible
'        DoEvents
'    Loop
    
    ' 释放对象
    Set WebBrowser = Nothing
Application.EnableEvents = True
End Sub

