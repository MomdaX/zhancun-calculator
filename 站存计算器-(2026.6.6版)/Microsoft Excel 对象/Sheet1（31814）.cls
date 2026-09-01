VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub 铁路十八点统计系统()
    Dim oDom As Object, ow As Object, result As String, pathDes As String
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    ow.execScript "function reloadCode(){return new Date().getTime();}"
    randomTime = ow.reloadCode("") '验证码请求的参数
'    randomTime = ow.eval("new Date().getTime()")
    url = "http://10.190.166.37:8010/rimss-upss-web/upss/login/CheckCodeImage?randomTime=" & randomTime
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "get", url, False
        .setRequestHeader "Accept", "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8"
        .setRequestHeader "Referer", "http://10.190.166.37:8010/rimss-upss-web/jsp/login.html"
        .send
        
'        cookie = Split(Split(.getAllResponseHeaders(), "Set-Cookie: ")(1), ";")(0)
'        Debug.Print cookie
        
        path = ThisWorkbook.path & "\验证码.jpg"
        fn = Dir(path)
        If fn <> "" Then
            Kill path '删除前验证码图片
        End If
        If .Status = 200 Then
            Dim stream As Object
            Set stream = CreateObject("ADODB.Stream")
            stream.Type = 1
            stream.Open
            stream.Write .responseBody
            stream.SaveToFile path, 1 '下载验证码图片
            stream.Close
        Else
            Debug.Print "下载失败：" & .Status & " " & .statusText
            End
        End If
        With UserForm4
'            验证码 = ""
            .Show 0
            .Caption = "验证码"
            .Image1.Picture = LoadPicture(path)
'            验证码 = InputBox("请输入图片中的验证码！")
            Do
                DoEvents
                验证码 = .Caption
            Loop While 验证码 = "验证码"
        End With
        Unload UserForm4
        
        checkCode = 验证码 '验证码图片内容
        Kill path '删除验证码图片
    End With

    url = "http://10.190.166.37:8010/rimss-upss-web/upss/login/login"
    data = "conditionJson=searchType%3D1%26code%3Dqzgyz%26password%3DCKYHKbJZPMgCcbdapcGbDQ%3D%3D&checkCode=" & checkCode & "&randomTime=" & randomTime
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .setRequestHeader "Origin", "http://10.190.166.37:8010"
        .setRequestHeader "Referer", "http://10.190.166.37:8010/rimss-upss-web/jsp/login.html"
'        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
'        .setRequestHeader "Cookie", cookie
        .send data
        Debug.Print .responseText
        ssoToken = JsonConverter.ParseJson(.responseText)("ssoToken")
'        Debug.Print .getAllResponseHeaders
        Cookie = Split(Split(.getAllResponseHeaders(), "Set-Cookie: ")(1), ";")(0)
        Debug.Print Cookie
        
        
        url = "http://10.190.166.37:8010/rimss-rsas-web/rsas/rsas-main.jsp"
        .Open "GET", url, False
        .setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
        .setRequestHeader "Referer", "http://10.190.166.37:8010/rimss-upss-web/jsp/login.html"
        .setRequestHeader "Cookie", Cookie
        .send data
        Debug.Print .responseText
        Cookie = Split(Split(.getAllResponseHeaders(), "Set-Cookie: ")(1), ";")(0)
        Debug.Print Cookie
        'var ssoToken = '8a3e26a592281ede019278b762d47d04';
        
        
        
        url = "http://10.190.166.37:8010/rimss-rsas-web/rsas/login/login"
        data = "conditionJson=searchType%3D2%26ssoToken%3D" & ssoToken & "&searchType=2&ssoToken=" & ssoToken
        .Open "GET", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .setRequestHeader "Origin", "http://10.190.166.37:8010"
        .setRequestHeader "Referer", "http://10.190.166.37:8010/rimss-rsas-web/rsas/login_new.jsp?ssoToken=" & ssoToken
        .setRequestHeader "Cookie", Cookie
        .send data
        Debug.Print .responseText
    End With
    

    With CreateObject("WinHttp.WinHttpRequest.5.1")
        url = "http://10.190.166.37:8010/rimss-rsas-web/rsas/qbjr/listGdml"
'        data = "conditionJson=ch%3D"
        .Open "GET", url, False
        .setRequestHeader "Cookie", Cookie
        .send
        Debug.Print .responseText
    End With

End Sub

Sub aes加密()
    'var key = CryptoJS.enc.Utf8.parse("abcdefgabcdefg12");
    'var srcs = CryptoJS.enc.Utf8.parse(password);
    'password = CryptoJS.AES.encrypt(srcs, key, {mode:CryptoJS.mode.ECB,padding: CryptoJS.pad.Pkcs7});
    'conditionJson = "searchType=1&code=" & code & "&password=" & password
    'conditionJson: searchType=1&code=qzgyz&password=CKYHKbJZPMgCcbdapcGbDQ==
    'function reloadCode() {randomTime = new Date().getTime();}

    Dim oDom As Object, ow As Object, result As String, pathDes As String
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    pathDes = "http://10.190.168.62:80/nnjzj/resources/script/common/aes.js" '网络读取
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "get", pathDes, False
        .send
    '        Debug.Print .responseText
        DES = .responseText
        ow.execScript DES
        code = "qzgyz"
        password = ow.encrypt("Nn_123456", "abcdefgabcdefg12")
    End With
End Sub

Sub 清空任务栏按钮()
    
    Dim shp As Shape
    On Error Resume Next
    For Each shp In Sheet1.Shapes
        If shp.Type = msoFormControl Then
            shp.ControlFormat.value = 0 'xlOn
        End If
        'shp.Select
    Next
    On Error GoTo 0
End Sub
