VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm6 
   Caption         =   "UserForm6"
   ClientHeight    =   9084
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   16188
   OleObjectBlob   =   "UserForm6.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Private Declare PtrSafe Function InternetSetCookie Lib "wininet.dll" Alias "LnternetSetCookieA" (ByVal lpszUrl As String,ByVal lpszCookieName As String,ByVal lpszCookieData As String) As Long
'加载Cookie库
'Private Declare Function InternetSetCookie Lib "wininet.dll" Alias "InternetSetCookieA" (ByVal lpszUrl As String, ByVal lpszCookieName As String, ByVal lpszCookieData As String) As Long

Private Sub UserForm_Initialize()
'Private Sub UserForm_Activate()
    Dim path As String
    Do While Me.WebBrowser1.Busy
        DoEvents
    Loop
    
'    With Sheet20
'        cookie = .Range("F3").value
'        试卷编号 = .Cells(6, "C").value
'        试卷ID = .Range("Q6").value
'
'        '设置批注
'        With .Cells(6, "J")
'            url1 = .Comment.Text '页面
'            url2 = "http://10.190.168.62/nnjzj/user/exam/topic_answer.html?tpid=" & 试卷编号 & "&egid=" & 试卷ID '题目和答案
'            '下载试卷到批注中
'            With CreateObject("WinHttp.WinHttpRequest.5.1")
'                .Open "GET", url1, False
'                .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
'                .setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=UTF-8"
'                .setRequestHeader "Origin", "http://10.190.168.62"
'                .setRequestHeader "Referer", "http://10.190.168.62/nnjzj/manager/index.html"
'                .setRequestHeader "Cookie", cookie
'                .send
'                res = .responseText
'                保存.文件保存 res, "view", "html", "C:\Users\Administrator\Desktop\每日站存\试卷缓存\"
'                '要保存为UFT-8格式
'
'                .Open "GET", url2, False
'                .setRequestHeader "Cookie", cookie
'                .send
'                res = .responseText
'                保存.文件保存 res, "topic_answer", "html", "C:\Users\Administrator\Desktop\每日站存\试卷缓存\"
'                '要保存为UFT-8格式
'
'                path = "C:\Users\Administrator\Desktop\每日站存\试卷缓存\view.html"
'            End With
'        End With
'    End With
    ' 设置 WebBrowser 控件位置和大小
    With Me.WebBrowser1
        .navigate "about:blank"
        .navigate "C:\Users\Administrator\Desktop\每日站存\试卷缓存\考试练习系统\考试练习系统.html"
'        .navigate "C:\Users\Administrator\Desktop\每日站存\试卷缓存\topic_answer.html"
'        .document.body.innerHTML = res 'HTML页面没在题目数据
        .Visible = True
'        .Width = 700
'        .Height = 550
    End With
    
    UserForm6.Caption = Selection.value
    
'   等待 WebBrowser 控件加载完毕
    Do While Me.WebBrowser1.Busy Or Me.WebBrowser1.readyState <> 4
        DoEvents
    Loop
    
    '滑动到最底部
'    On Error Resume Next
'    WebBrowser1.document.parentWindow.scrollTo 0, WebBrowser1.document.body.ScrollHeight
'    On Error GoTo 0

    ' 加载 HTML 代码
'    With WebBrowser1
'        .Visible = True
'        .navigate "about:blank"
'        .document.body.innerHTML = Selection.Comment.Text '代码段
'    End With

'     WebBrowser1.navigate "F:\Momda\货票缓存\" & Sheet13.Cells(Selection.Row, "F").Value & ".html" '文件
    
    ' 释放对象
'    Set WebBrowser = Nothing
    Application.EnableEvents = True
End Sub


