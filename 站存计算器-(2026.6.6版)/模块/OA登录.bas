Attribute VB_Name = "OA登录"
Function 公司OA登录()
    Dim http As Object, Cookie As String, reg As New RegExp, ckr()
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = "http://10.190.128.231/logincheck.php"
    data = "UNAME=8001&PASSWORD=cXpjd2RxemcxMjMu&encode_type=1"
        With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'        .setRequestHeader "Origin", "http://10.190.136.8"
'        .setRequestHeader "Referer", "http://10.190.136.8/"
'        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .send data
        cor = Split(.getAllResponseHeaders(), "Set-Cookie: ")
        'Cookie = "PHPSESSID=sufvjaqpc7gv33eo9kkbinmjk6; USER_NAME_COOKIE=8001; OA_USER_ID=3416; SID_3416=34ea9217"
        Debug.Print .responseText
        
        ReDim ckr(1 To 5)
        ckr(1) = Split(Split(.getAllResponseHeaders(), "PHPSESSID=")(1), ";")(0)
        ckr(2) = Split(Split(.getAllResponseHeaders(), "USER_NAME_COOKIE=")(1), ";")(0)
        ckr(3) = Split(Split(.getAllResponseHeaders(), "OA_USER_ID=")(1), Chr(13))(0)
        ckr(4) = Split(Split(.getAllResponseHeaders(), "SID_1774=")(1), ";")(0)
        ckr(5) = Split(Split(.getAllResponseHeaders(), "UI_COOKIE=")(1), ";")(0)
    
        End With
    获取Cookie = "USER_NAME_COOKIE=" & ckr(2) & "; OA_USER_ID=" & ckr(3) & "; UI_COOKIE=" & ckr(5) & "; PHPSESSID=" & ckr(1) & "; SID_1774=" & ckr(4)
End Function



Function 获取Cookie() '成功了
    Dim http As Object, Cookie As String, reg As New RegExp, ckr()
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = "http://10.190.136.8/logincheck.php"
    data = "UNAME=%CE%A4%CE%C4%BF%B5&PASSWORD=Wwk1002-&submit=%B5%C7+%C2%BC"
        With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'        .setRequestHeader "Origin", "http://10.190.136.8"
'        .setRequestHeader "Referer", "http://10.190.136.8/"
'        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .send data
'        cor = Split(.getAllResponseHeaders(), "Set-Cookie: ")
        'Cookie = "USER_NAME_COOKIE=%CE%A4%CE%C4%BF%B5; OA_USER_ID=%CE%A4%CE%C4%BF%B5; UI_COOKIE=0; PHPSESSID=91f8bafb57acc3ef91a1ab15544ce51f; SID_1774=449928a3"
        
        ReDim ckr(1 To 5)
        ckr(1) = Split(Split(.getAllResponseHeaders(), "PHPSESSID=")(1), ";")(0)
        ckr(2) = Split(Split(.getAllResponseHeaders(), "USER_NAME_COOKIE=")(1), ";")(0)
        ckr(3) = Split(Split(.getAllResponseHeaders(), "OA_USER_ID=")(1), Chr(13))(0)
        ckr(4) = Split(Split(.getAllResponseHeaders(), "SID_1774=")(1), ";")(0)
        ckr(5) = Split(Split(.getAllResponseHeaders(), "UI_COOKIE=")(1), ";")(0)
    
        End With
    获取Cookie = "USER_NAME_COOKIE=" & ckr(2) & "; OA_USER_ID=" & ckr(3) & "; UI_COOKIE=" & ckr(5) & "; PHPSESSID=" & ckr(1) & "; SID_1774=" & ckr(4)
End Function


Sub 从OA上获取钦州港站股道存车情况()

    Dim xmlhttp As Object, oDom As Object, reg As New RegExp
    Dim HTML As String
    Dim divs As IHTMLElementCollection
    Dim parser As HTMLDocument
    
    Dim url As String, Cookie As String
    Dim Body As HTMLBody
    
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    ow.execScript "function Number(){return Math.random()}"
    tx1 = ow.eval("new Date().getTime()") '时间戳
    
'    url = "http://10.190.136.222:8099/WebReport/ReportServer?_=" & tx & "&__boxModel__=true&op=fr_write&cmd=read_w_content&sessionID=3370&reportIndex=0&browserWidth=1098&iid=" & num & "&__cutpage__="
    url1 = "http://10.190.136.222:8099/WebReport/ReportServer?reportlet=/zbs/[94a6][5dde][6e2f][8282][70b9][7ad9][5b58][8f66][63a8][9001].cpt&__parameters__={%22_%22:%22" & tx1 & "%22,%22__pi__%22:%22true%22,%22op%22:%22write%22}"
'    Cookie = 获取Cookie
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    With xmlhttp
        .Open "get", url1, False
'        .setRequestHeader "Host", "10.190.136.222:8099"
        .setRequestHeader "Referer", "http://10.190.136.222:8099/WebReport/ReportServer?reportlet=sr/sr.cpt"
        .send

        HTML = .responseText
        reg.Pattern = "FR.SessionMgr.register\('(\d+)', contentPane\)" '正则
        sid = reg.Execute(HTML)(0).SubMatches(0) '返回结果中取sid
        tx2 = ow.eval("new Date().getTime()") '时间戳
        num = VBA.Format(ow.Number(""), "0.000000000000000") '随机数
        
        url2 = "http://10.190.136.222:8099/WebReport/ReportServer?_=" & tx2 & "&__boxModel__=true&op=fr_write&cmd=read_w_content&sessionID=" & sid & "&reportIndex=0&iid=" & num
        .Open "get", url2, False
        .setRequestHeader "Referer", url1
        .send
        HTML = .responseText
        Debug.Print .responseText
'        Set parser = New HTMLDocument
'        Set body = parser.createElement("body")
'        body.innerHTML = html
'
'        Set trs = body.getElementsByTagName("tr")
    End With
End Sub
Sub 从十八点子系统上获取钦州港站股道存车情况()

    Dim xmlhttp As Object, oDom As Object, reg As New RegExp
    Dim HTML As String
    Dim divs As IHTMLElementCollection
    Dim parser As HTMLDocument
    
    Dim url As String, Cookie As String
    Dim Body As HTMLBody
    
'    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
'    ow.execScript "function Number(){return Math.random()}"
'    tx1 = ow.eval("new Date().getTime()") '时间戳
    
    url = "http://10.190.166.37:8010/rimss-rsas-web/rsas/qbjr/listGdxc?gdm=gdm%3D6%26pxfs%3D%26falg%3D"
'    Cookie = 获取Cookie
    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    With xmlhttp
        .Open "get", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .setRequestHeader "Referer", "http://10.190.166.37:8010/rimss-rsas-web/rsas/jsp/collectData/qbjr.jsp?token=8a3e26a59304c66101931be8fb025141"
        .setRequestHeader "Origin", "http://10.190.166.37:8010"
        .setRequestHeader "Cookie", "JSESSIONID=89014989A19EB7A95BDEF1F901AA700B; operAuthButtonPCode=132"
        .send

        HTML = .responseText
        Debug.Print .responseText
        reg.Pattern = "FR.SessionMgr.register\('(\d+)', contentPane\)" '正则
        sid = reg.Execute(HTML)(0).SubMatches(0) '返回结果中取sid
        tx2 = ow.eval("new Date().getTime()") '时间戳
        num = VBA.Format(ow.Number(""), "0.000000000000000") '随机数
        
        url2 = "http://10.190.136.222:8099/WebReport/ReportServer?_=" & tx2 & "&__boxModel__=true&op=fr_write&cmd=read_w_content&sessionID=" & sid & "&reportIndex=0&iid=" & num
        .Open "get", url2, False
        .setRequestHeader "Referer", url1
        .send
        HTML = .responseText
        Debug.Print .responseText
'        Set parser = New HTMLDocument
'        Set body = parser.createElement("body")
'        body.innerHTML = html
'
'        Set trs = body.getElementsByTagName("tr")
    End With
End Sub
