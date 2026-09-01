VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm7 
   Caption         =   "发送文件"
   ClientHeight    =   4620
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   5928
   OleObjectBlob   =   "UserForm7.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public 路径$, datas, bdy$, uid$, userid$, userna$
Private Sub CommandButton1_Click() '退出按钮
    Label14.Caption = "即将退出程序！"
    'UserForm7.Hide'隐藏
    Unload Me
End Sub

Sub 接收人信息重置()
    路径 = ""
    datas = ""
    bdy = ""
    uid = ""
    userid = ""
    userna = ""
End Sub

Private Sub CommandButton2_Click() '发送按钮

    With UserForm7
        Debug.Print .userna
        Debug.Print .uid
        Debug.Print .userid
        Debug.Print .路径
        If .路径 <> "" Then
            Call OA发送文件.OA_发送文件(.uid, .userid, .路径, .datas, .bdy) '接收人ID，用户ID，路径，文件体,分隔符
            Label14.Caption = "发送成功！"
        Else
            MsgBox "发送失败！"
            Label3.Caption = "双击加载文件"
            Label14.Caption = "加载要发送的文件！"
        End If
    End With
    'Call 接收人信息重置
End Sub

Private Sub Label3_DblClick(ByVal Cancel As MSForms.ReturnBoolean) '双击加载上传文件
    
    ' 生成boundary
    bdy = "----WebKitFormBoundary" & Format(Now, "yyyymmddhhnnss") & Right("0000" & Int(Rnd() * 10000), 4)
    
    '接收人信息
    r = ListBox1.ListIndex
    If r > 0 Then
        uid = ListBox1.list(r, 1)
        userid = Label9.Caption
        
        If IsNumeric(uid) Then
            
            If ActiveSheet.CodeName = "Sheet1" Then
                路径 = 站存导出.导出到桌面
            ElseIf ActiveSheet.CodeName = "Sheet22" Then
                路径 = 站存导出.车流表导出
            Else
                Label3.Caption = "双击加载文件"
                Label14.Caption = "导出文件出错!"
                Call 接收人信息重置
            End If
            If 路径 = "" Then
                Label14.Caption = "文件加载失败!"
            Else
                Label14.Caption = "文件加载完成!"
                Label3.Caption = Mid(路径, InStrRev(路径, "\") + 1)
                
                '获取返回的请求体
                datas = OA发送文件.UploadFile(路径, bdy)
                Kill 路径 '删除文件
            End If
        Else
            Label14.Caption = "接收人信息有误！"
        End If
    Else
        Label14.Caption = "未选择接收人!"
    End If
End Sub

Private Sub ListBox1_Change()
    Dim r
    r = ListBox1.ListIndex
    uid = ListBox1.list(r, 1) '选中的人员ID
    If IsNumeric(uid) Then
        Set data = 获取人员信息(uid)
    
        xm = data.user_name
        userna = xm '用户名
        gw = data.priv_name
        zt = CallByName(data, "online", VbGet)
        
        Label11.Caption = xm
        Label13.Caption = zt
        Label9.Caption = gw
        CheckBox1.value = IIf(zt = "离线", False, True)
    End If
End Sub

Private Sub UserForm_Initialize()
    OA发送文件.获取人员列表
End Sub
