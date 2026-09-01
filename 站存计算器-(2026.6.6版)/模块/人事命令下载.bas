Attribute VB_Name = "人事命令下载"
Sub 下载人事命令()
    
    Dim xmlhttp As New MSXml2.ServerXMLHTTP
    Dim link As String, path As String
    Dim HTML As String
    'Dim Xhtml As New HTMLDocument
    Dim divs As IHTMLElementCollection
    Dim parser As HTMLDocument
    
    Dim url As String, Cookie As String
    Dim Body As HTMLBody
    
    url = "http://10.190.136.8/general/file_folder/folder.php?SORT_ID=75&FILE_SORT=1&start=0&TOTAL_ITEMS=428&PAGE_SIZE=1000"
    path = Sheet16.Range("G2").value
    
    'Cookie = "USER_NAME_COOKIE=%CE%A4%CE%C4%BF%B5; OA_USER_ID=%CE%A4%CE%C4%BF%B5; UI_COOKIE=0; PHPSESSID=91f8bafb57acc3ef91a1ab15544ce51f; SID_1774=449928a3"
    'Cookie = "USER_NAME_COOKIE=%CE%A4%CE%C4%BF%B5; OA_USER_ID=%CE%A4%CE%C4%BF%B5; UI_COOKIE=0; PHPSESSID=4d889a4061862d009910fcdc84ce39e8; SID_1774=462ebd62; mytable_notify_1774=57183bd3"
    'Cookie = "USER_NAME_COOKIE=%CE%A4%CE%C4%BF%B5; OA_USER_ID=%CE%A4%CE%C4%BF%B5; UI_COOKIE=0; PHPSESSID=c5dab1c694bb61f687f0d32883060784; SID_1774=709e17e5"
    Cookie = OA登录
    
    xmlhttp.Open "get", url, False
    xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
    xmlhttp.setRequestHeader "Host", "10.190.136.8"
    xmlhttp.setRequestHeader "Cookie", Cookie
    xmlhttp.send
    
    HTML = xmlhttp.responseText
    Debug.Print xmlhttp.responseText
    
    Set parser = New HTMLDocument
    Set Body = parser.createElement("body")
    Body.innerHTML = HTML
    
    Set trs = Body.getElementsByTagName("tr")
    
    'For Each div In divs
    For Each tr In trs
    
    Set tds = tr.getElementsByTagName("td")
    For Each td In tds
    'Debug.Print td.Align
        If td.Align = "center" Then
    
            时间 = Split(td.innerText, " ", 2)(0)
            Exit For
        End If
    Next
    
    Set divs = tr.getElementsByTagName("div")
        For Each div In divs
            Debug.Print div.Title
            If div.Title Like "*人生*" Then
                Set aas = div.getElementsByTagName("a")
                For Each aa In aas
                'Debug.Print aa.innerText
                    If aa.innerText = "下载" Then
                        文件名 = 时间 & "-" & div.Title
                        hurl = Split(aa.href, ":", 2)(1)
                        link = "http://10.190.136.8" & hurl
                        Debug.Print link
                        Call 下载(link, path & 文件名, Cookie)
                    End If
                Next
            End If
        
        Next
    Next
    MsgBox "下载完成"
End Sub

Function 下载(url As String, path As String, Cookie As String)
    Dim xhttp As Object
    Set xhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
'    Cookie = "PHPSESSID=17388f27684c70c9a9ff096944228e9d; USER_NAME_COOKIE=%CE%A4%CE%C4%BF%B5; OA_USER_ID=%CE%A4%CE%C4%BF%B5; SID_1774=3aa28f95; UI_COOKIE=0; mytable_notify_1774=57183bd3"
    With xhttp
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "Host", "10.190.136.8"
        .setRequestHeader "Cookie", Cookie
        .send
        '测试数据类型
'        tearr = .responseText
'        tebrr = .responseBody
        '下载方式1
'        ByteToFile .responseBody, path
        '下载方式2
        If .Status = 200 Then
            With CreateObject("ADODB.Stream")
                .Type = 1
                .Open
                .Write xhttp.responseBody
                .SaveToFile path, 2 '下载图片/文件
                .Close
            End With
            Debug.Print "下载完成"
        Else
            Debug.Print "下载失败：" & .Status & " " & .statusText
        End If
    End With
    
    Set xhttp = Nothing
End Function


Sub ByteToFile(arrByte, strFileName)
    With CreateObject("Adodb.Stream")
        .Type = adTypeBinary
        .Open
        .Write arrByte
        .SaveToFile strFileName, adSaveCreateOverWrite
        .Close
    End With
End Sub

Function OA登录() '成功了
Dim http As Object, Cookie As String, reg As New RegExp, ckr()
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
url = "http://10.190.136.8/logincheck.php"
data = "UNAME=%CE%A4%CE%C4%BF%B5&PASSWORD=Wwk1002-&submit=%B5%C7+%C2%BC"
    With http
    .Open "POST", url, False
    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    .setRequestHeader "Origin", "http://10.190.136.8"
    .setRequestHeader "Referer", "http://10.190.136.8/"
    .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
'    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=UTF-8"
    .send data
'    cor = Split(.getAllResponseHeaders(), "Set-Cookie: ")
    'Cookie = "USER_NAME_COOKIE=%CE%A4%CE%C4%BF%B5; OA_USER_ID=%CE%A4%CE%C4%BF%B5; UI_COOKIE=0; PHPSESSID=91f8bafb57acc3ef91a1ab15544ce51f; SID_1774=449928a3"
    
    ReDim ckr(1 To 5)
    ckr(1) = Split(Split(.getAllResponseHeaders(), "PHPSESSID=")(1), ";")(0)
    ckr(2) = Split(Split(.getAllResponseHeaders(), "USER_NAME_COOKIE=")(1), ";")(0)
    ckr(3) = Split(Split(.getAllResponseHeaders(), "OA_USER_ID=")(1), Chr(13))(0)
    ckr(4) = Split(Split(.getAllResponseHeaders(), "SID_1774=")(1), ";")(0)
    ckr(5) = Split(Split(.getAllResponseHeaders(), "UI_COOKIE=")(1), ";")(0)

    End With
    OA登录 = "USER_NAME_COOKIE=" & ckr(2) & "; OA_USER_ID=" & ckr(3) & "; UI_COOKIE=" & ckr(5) & "; PHPSESSID=" & ckr(1) & "; SID_1774=" & ckr(4)
End Function


