Attribute VB_Name = "OA发送文件"
Function OA_发送文件(uid$, userid$, 路径$, datas, bdy$)
    Dim Cookie$, url$, 姓名$, id$, oDom As Object, jss As Object, js As Object, http As Object, url_Base64$, txt_Base64$
    
    '获取Cookie
    Cookie = 公司OA登录_获取cookie
    'Debug.Print cookie
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    With http
        '上传文件
        url = "http://10.190.128.231/module/upload/upload.php?module=im"
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "multipart/form-data; boundary=" & bdy
        .setRequestHeader "Origin", "http://10.190.128.231"
        .setRequestHeader "Referer", "http://10.190.128.231/general/index.php?isIE=0&modify_pwd=0"
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "Cookie", Cookie
        .send datas
        'Debug.Print .responseText
        strJson = .responseText
        Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
        ow.execScript "var js =" & strJson & ";", "JScript"
        'Set jss = ow.eval("js")
        Set jss = ow.js
        
        '接收人信息
        'uid = 3416
        'userid = "钦州港站调"
        fileid = CallByName(jss, "id", VbGet)
        fileName = Mid(路径, InStrRev(路径, "\") + 1)
        
        '发送文件
        url = "http://10.190.128.231/general/index_simple_submit.inc.php?uid=" & uid & "&userid=" & userid & "&fileid=" & fileid & "&filename=" & fileName & "&flag=0"
        .Open "GET", url, True
        .setRequestHeader "Cookie", Cookie
        .setRequestHeader "Host", "10.190.128.231"
        .setRequestHeader "Connection", "keep-alive"
        .setRequestHeader "Referer", "http://10.190.128.231/general/index.php?isIE=0&modify_pwd=0"
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .send
        'Debug.Print .responseText
        
        Debug.Print "接收人：" & UserForm7.userna & ",文件：" & fileName
    End With


End Function


Function 公司OA登录_获取cookie()
    Dim http As Object, Cookie As String, reg As New RegExp, ckr(), oDom As Object, js As Object, url_Base64$, txt_Base64$
    
'    url_Base64 = "http://10.190.128.231/static/js/base64/base64.min.js"
'    txt_Base64 = 加载js(url_Base64)
'    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
'    ow.execScript txt_Base64
'    ow.execScript "function Number(){return Math.random()}"
'    tx = ow.eval("new Date().getTime()") '时间戳
'    boundary = "WebKitFormBoundary" & Left(ow.Base64.encode(tx), 16)
    
    用户ID = "8002"
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = "http://10.190.128.231/logincheck.php"
    data = "UNAME=" & 用户ID & "&PASSWORD=cXpjd2RxemcxMjMu&encode_type=1"
        With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'        .setRequestHeader "Origin", "http://10.190.136.8"
'        .setRequestHeader "Referer", "http://10.190.136.8/"
'        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .send data
        'cor = Split(.getAllResponseHeaders(), "Set-Cookie: ")
        'Debug.Print .responseText

        ReDim ckr(1 To 5)
        ckr(1) = Split(Split(.getAllResponseHeaders(), "KEY_RANDOMDATA=")(1), ";")(0)
        ckr(2) = Split(Split(.getAllResponseHeaders(), "PHPSESSID=")(1), "; ")(0)
        ckr(3) = Split(Split(.getAllResponseHeaders(), "USER_NAME_COOKIE=")(1), "; ")(0)
        ckr(4) = Split(Split(.getAllResponseHeaders(), "OA_USER_ID=")(1), Chr(13))(0)
        ckr(5) = Split(Split(.getAllResponseHeaders(), "SID_3417=")(1), "; ")(0)
        Cookie = "USER_NAME_COOKIE=" & ckr(3) & "; OA_USER_ID=" & ckr(4) & "; PHPSESSID=" & ckr(2) & "; SID_3417=" & ckr(5)
    End With
    
    公司OA登录_获取cookie = Cookie
End Function
Function 获取人员列表()
    Dim Cookie$, url$, 姓名$, id$, oDom As Object, jss As Object, js As Object, http As Object
    
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    url = "http://10.190.128.231/inc/user_list/tree.php?DEPT_ID=660&PARA_URL1=&PARA_URL2=/general/ipanel/user/user_info.php&PARA_TARGET=_blank&PRIV_NO=&PARA_ID=WINDOW&PARA_VALUE=1&MANAGE_FLAG=0&MODULE_ID=2&SHOW_IP=&PWD=&DEPT_PRIV=1&ROLE_PRIV=2&PRIV_NO_FLAG=0&OP_SMS=1&rand=1242624128&_=" & tx
        
      Cookie = 公司OA登录_获取cookie
      Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        With http
            .Open "GET", url, False
            '.setRequestHeader "Referer", "http://10.190.128.231/inc/user_tree.php?FROM=WebOS&TREE_ID=orgTree1&SHOW_IP=0&SHOW_BUTTON=0&JSON_URL=/inc/user_list/tree.php?MANAGE_FLAG=0&DEPT_ID=0&PARA_URL1=&PARA_URL2=/general/ipanel/user/user_info.php&e=&PARA_TARGET=_blank&PRIV_NO=&PARA_ID=WINDOW&PARA_VALUE=1&MODULE_ID=2&SHOW_IP=&PWD=&DEPT_PRIV=1&ROLE_PRIV=2&PRIV_NO_FLAG=0&OP_SMS=1&rand=59314421"
            .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
            .setRequestHeader "Cookie", Cookie
            .send
            'Debug.Print .responseText
            strJson = .responseText
            ow.execScript "var js = " & strJson, "JScript"
            Set jss = ow.js
            With UserForm7.ListBox1
                .ColumnCount = 2
                .AddItem "姓名"
                .list(0, 1) = "id"
                .ColumnWidths = "60;30"
            End With
            i = 0
            For Each js In jss
                姓名 = CallByName(js, "title", VbGet)
                id = js.user_id
                'Debug.Print 姓名 & "：" & id
            
                With UserForm7.ListBox1
                    i = i + 1
                    .AddItem 姓名
                    .list(i, 1) = id
                End With
            Next
        End With
        
        '增加额外人员
        With UserForm7.ListBox1
            .AddItem "钦州港站调"
            .list(.ListCount - 1, 1) = "3416"
        End With

End Function

Function 获取人员信息(ryid)
'curl 'http://10.190.128.231/general/userinfo.php?UID=2908' \
'  -H 'Accept: */*' \
'  -H 'Accept-Language: zh-CN,zh;q=0.9' \
'  -H 'Connection: keep-alive' \
'  -H 'Cookie: PHPSESSID=88uriium5k6ponrq2i7l17a0e0; USER_NAME_COOKIE=8002; OA_USER_ID=3417; SID_3417=4d4eed4' \
'  -H 'Referer: http://10.190.128.231/general/index.php?isIE=0&modify_pwd=0' \
'  -H 'User-Agent: Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.88 Safari/537.36' \
'  -H 'X-Requested-With: XMLHttpRequest' \
'    Dim cookie$, url$, 姓名$, id$, oDom As Object, jss As Object, js As Object, http As Object
    Dim Cookie$, url$, 姓名$, id$, oDom As Object, jss As Object, js As Object, http As Object
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    
    'ryid = "3417"
    url = "http://10.190.128.231/general/userinfo.php?UID=" & ryid
      Cookie = 公司OA登录_获取cookie
      Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        With http
            .Open "GET", url, False
            .setRequestHeader "Referer", "http://10.190.128.231/general/index.php?isIE=0&modify_pwd=0"
            .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
            .setRequestHeader "Cookie", Cookie
            .send
            'Debug.Print .responseText
            strJson = .responseText
            ow.execScript "var js = " & strJson, "JScript"
        End With
        
        Set jss = ow.js
        Set 获取人员信息 = CallByName(jss, "data", VbGet)
    
End Function

Function 加载js(urlPath As String)
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "get", urlPath, False
        .send
'        Debug.Print .responseText
        加载js = .responseText
    End With
End Function
Function UploadFile(filePath As String, boundary As String)
    Dim http As Object, stream As Object, reqStream As Object
    Dim formData As String, fileData() As Byte
    Dim fileName As String, fileType As String, lastModifiedDate As String, fileSize As Long
    Dim errorMessage As String

    On Error Resume Next
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
    fileType = "application/vnd.ms-excel"
    fileSize = FileLen(filePath)
    lastModifiedDate = Format(FileDateTime(filePath), "ddd mmm dd yyyy hh:mm:ss") & " GMT+0800 (中国标准时间)"
    
    '检查文件是否存在
    If Dir(filePath) = "" Then
        Debug.Print "文件不存在: " & filePath
        Exit Function
    End If

    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 1
        .Open
        .LoadFromFile filePath
        If Err.Number <> 0 Then
            errorMessage = "读取文件失败: " & Err.Description
            Debug.Print errorMessage
            .Close
            Exit Function
        End If
        fileData = .Read
        .Close
    End With

    ' 创建请求体
    Set reqStream = CreateObject("ADODB.Stream")
    reqStream.Type = 1
    reqStream.Open

    ' 构建完整的请求体
    Dim requestBody As String
    requestBody = ""
    requestBody = requestBody & "--" & boundary & vbCrLf & _
        "Content-Disposition: form-data; name=""id""" & vbCrLf & vbCrLf & _
        "WU_FILE_1" & vbCrLf
        
    requestBody = requestBody & "--" & boundary & vbCrLf & _
        "Content-Disposition: form-data; name=""name""" & vbCrLf & vbCrLf & _
        fileName & vbCrLf
        
    requestBody = requestBody & "--" & boundary & vbCrLf & _
        "Content-Disposition: form-data; name=""type""" & vbCrLf & vbCrLf & _
        fileType & vbCrLf
        
    requestBody = requestBody & "--" & boundary & vbCrLf & _
        "Content-Disposition: form-data; name=""lastModifiedDate""" & vbCrLf & vbCrLf & _
        lastModifiedDate & vbCrLf
        
    requestBody = requestBody & "--" & boundary & vbCrLf & _
        "Content-Disposition: form-data; name=""size""" & vbCrLf & vbCrLf & _
        fileSize & vbCrLf
        
    ' 添加文件部分
    requestBody = requestBody & "--" & boundary & vbCrLf & _
        "Content-Disposition: form-data; name=""file""; filename=""" & fileName & """" & vbCrLf & _
        "Content-Type: " & fileType & vbCrLf & vbCrLf

    
    reqStream.Write StringToBytes(requestBody) ' 写入请求体头部
    reqStream.Write fileData ' 写入文件数据
    reqStream.Write StringToBytes(vbCrLf & "--" & boundary & "--" & vbCrLf) ' 写入结束标记

    ' 重置流位置
    reqStream.Position = 0
    
    ' 读取完整请求体
    UploadFile = reqStream.Read
    reqStream.Close
    On Error GoTo 0
End Function

Function StringToBytes(str As String) As Byte()
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText str
    stream.Position = 0
    stream.Type = 1
    stream.Position = 0
    StringToBytes = stream.Read
    stream.Close
End Function



