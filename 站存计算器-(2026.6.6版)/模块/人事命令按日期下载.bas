Attribute VB_Name = "人事命令按日期下载"
Sub 下载人事命令()
    Dim xmlhttp As Object
    Dim link As String, path As String
    Dim HTML As String, rn_ywj As Range
    'Dim Xhtml As New HTMLDocument
    Dim divs As IHTMLElementCollection
    Dim parser As HTMLDocument
    Dim 开始日期 As Date, 结束日期 As Date, 时间 As Date, Cookie As String
    Dim url As String
    Dim Body As HTMLBody
    
    Sheet16.UsedRange.Offset(2).ClearContents
    
    url = "http://10.190.136.8/general/file_folder/folder.php?SORT_ID=75&FILE_SORT=1&start=0&TOTAL_ITEMS=428&PAGE_SIZE=1000"
    path = Sheet16.Range("G2").value & "\"
    开始日期 = Sheet16.Range("H2").value
    结束日期 = Sheet16.Range("I2").value
    'cookie = "SID_245=4fb885b6; UI_COOKIE=0; mytable_notify_245=57183bd3; mytable_notify_1774=57183bd3; SID_1212=8315d919; PHPSESSID=04b1a5b7d131349b8a0b0cff1e49f5f4; USER_NAME_COOKIE=%CE%A4%CE%C4%BF%B5; OA_USER_ID=%CE%A4%CE%C4%BF%B5; SID_1774=599b8266"
    'cookie = "USER_NAME_COOKIE=%CE%A4%CE%C4%BF%B5; OA_USER_ID=%CE%A4%CE%C4%BF%B5; UI_COOKIE=0; PHPSESSID=91f8bafb57acc3ef91a1ab15544ce51f; SID_1774=449928a3"
    Cookie = OA登录() '登录后获取cookie

    Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
    With xmlhttp
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "Host", "10.190.136.8"
        .setRequestHeader "Cookie", Cookie
        .send
    
        If InStr(.responseText, "用户未登录") Then
        
            MsgBox "用户未登录!"
            End
        End If
    
    HTML = .responseText
    'Debug.Print xmlhttp.responseText
    End With
    
    Set parser = New HTMLDocument
    Set Body = parser.createElement("body")
    Body.innerHTML = HTML
    
    Set trs = Body.getElementsByTagName("tr")
    
    For Each tr In trs
    
        Set tds = tr.getElementsByTagName("td")
        For Each td In tds
        'Debug.Print td.Align
            If td.Align = "center" Then
                If td.innerText = "文件夹操作：" Then Exit For
                时间 = Split(td.innerText, " ", 2)(0)
                Exit For
            End If
        Next

        Set divs = tr.getElementsByTagName("div")
        For Each div In divs
            Debug.Print div.Title
            M1 = 开始日期 < 时间 Or Sheet16.Range("H2").value = ""
            M2 = 时间 < 结束日期 Or Sheet16.Range("I2").value = ""
            M3 = 时间 < 开始日期
            If div.Title Like "*人生*" And M1 And M2 Then
            '下载记录写入Sheet16表中
            Set rn_ywj = Sheet16.Range("A" & Sheet16.Cells(Sheet16.Rows.count, "A").End(3).Row + 1)
            If M3 Then MsgBox "下载完成": End
                rn_ywj.Offset(0, 0).value = 时间
                rn_ywj.Offset(0, 1).value = Split(div.Title, ".")(0)
                rn_ywj.Offset(0, 4).value = Split(div.Title, ".")(1)
                Set aas = div.getElementsByTagName("a")
                For Each aa In aas
                    If aa.innerText = "下载" Then
                    'Debug.Print aa.innerText
                        文件名 = 时间 & "-" & div.Title
                        hurl = Split(aa.href, ":", 2)(1)
                        link = "http://10.190.136.8" & hurl '得到图片/文件链接
                        Debug.Print link
                        Call 下载(link, path & 文件名, Cookie) '下载与保存
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
                .SaveToFile path, 1 '下载图片/文件
                .Close
            End With
            Debug.Print "下载完成"
        Else
            Debug.Print "下载失败：" & .Status & " " & .statusText
        End If
    End With
    
    Set xhttp = Nothing
End Function


Function ByteToFile(arrByte, strFileName) '文件保存程序
    With CreateObject("ADODB.Stream")
        .Type = 1
        .Open
        .Write arrByte
        .SaveToFile strFileName, 2 '下载图片
        .Close
    End With
End Function

Function OA登录() '成功了
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
    OA登录 = "USER_NAME_COOKIE=" & ckr(2) & "; OA_USER_ID=" & ckr(3) & "; UI_COOKIE=" & ckr(5) & "; PHPSESSID=" & ckr(1) & "; SID_1774=" & ckr(4)
End Function
