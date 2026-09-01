Attribute VB_Name = "虚拟场"
Sub 现车刷新_虚拟场()

    If Sheet7.AutoFilterMode Then Sheet7.AutoFilterMode = False '取消筛选状态
    Call 获取现车
    xrr = Sheet7.Range("A1").CurrentRegion.value
    Call 统计股道存车.股道存车(xrr)
    
End Sub

Function 虚拟场(股道, ow As Object, reg As Object)
    Dim 站名 As String, zm$, xhp As Object

    '股道 = "2"
    站名 = "钦州港"
    t = ow.reloadCode("")
    iid = VBA.Format(ow.Number(""), "0.0000000000000000") '随机数
    zm = ow.encode(站名)
    
    With reg
        .Pattern = "'(\w{8}-\w{4}-\w{4}-\w{4}-\w{12})'" '"FR.SessionMgr.register('89c5c3c1-e710-47cc-9ae3-b8306d71ef0b', contentPane);"
        .Global = False
        .MultiLine = False
    End With
    
    Set xhp = CreateObject("MSXML2.XMLHTTP.6.0") 'CreateObject("WinHttp.WinHttpRequest.5.1")
    With xhp
        url = "http://10.190.5.143:8081/webroot/decision/view/form?viewlet=cw/czgdxq.cpt&table=V_B_VGDSYCLK&ZMLM=QVZ&ZM=" & zm & "&op=write&GDM=" & 股道
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "Referer", "http://10.190.5.143:8081/webroot/decision/view/form?viewlet=cw/czgdxcqk.cpt&op=write&dwName=钦州车务段&RoleID=2"
        .send
        HTML = .responseText
        '保存.桌面 html

        sessionID = reg.Execute(HTML)(0).SubMatches(0)

        url = "http://10.190.5.143:8081/webroot/decision/view/form?op=fr_write&cmd=read_w_content&&pn=0&iid=" & iid
        .Open "GET", url, True
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "sessionID", reg.Execute(HTML)(0).SubMatches(0)
        .send
        
'        res = .responseText
'        ow.execScript "var rjs =" & res & ";"
'        html = ow.eval("rjs[""html""]")
'        reg.Pattern = "<script>.*</script>"
'        hStr = reg.Replace(html, "")
'
'        With 计划HTML.WebBrowser1
'            .width = 560
'            .navigate "about:blank"
'            Do While .Busy Or .readyState <> 4
'                DoEvents
'            Loop
'            .document.Open
'            .document.Write hStr
'            .document.Close
'        End With
'
'        计划HTML.Show
    End With
'
    Set 虚拟场 = xhp
End Function
    
Sub 获取现车()
    Dim oDom As Object, ow As Object, 股道 As String, 站名 As String, reg As New RegExp
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    ow.execScript "function Number(){return Math.random()}", "jscript"
    ow.execScript "function reloadCode(){return new Date().getTime();}", "jscript"
    ow.execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
    
    Sheet7.Cells(2, 1).CurrentRegion.Offset(2).ClearContents
    grr = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "B1", "B2", "X1", "X2", "X3", "X4", "X5", "X6", "X7", "X8", "X9", "X10", "X11", "X12", "X13", "X14", "X15", "H1", "H2", "H3", "H4", "H5", "L1", "L2", "L3", "ZL1", "ZL2", "ZL3", "LZ", "G1", "G2", "YX1", "YX2", "YX3", "TS1", "TS2", "SH1", "SH2", "SH3", "CZ3", "CZ4", "GT1", "GT2", "GM1", "GM2", "GM3", "GM4", "Y1", "Y2", "Y3", "Y4", "Y5", "Y6", "Y7", "Y8", "Y9", "Y10", "Y11", "Y12", "Y13", "Y14", "Y15", "Y16", "YQX", "##1", "##2", "##3") '"DY1", "DY2", "DY3", "DY4", "TSY1", "TSY2", "TSY3", "TSY4", "YH1", "YH2",

    Dim ds As Object
    Set ds = CreateObject("Scripting.Dictionary")
    For Each gd In grr
        '发起请求
        股道 = gd
        Set ds(股道) = 虚拟场(股道, ow, reg)
    Next
    
    Do '异步查询1
        For Each dd In ds.keys
            DoEvents
            If ds(dd).readyState = 4 Then
                res = ds(dd).responseText
                
                ow.execScript "var rjs =" & res & ";"
                HTML = ow.eval("rjs[""html""]")
                
                reg.Pattern = "<script>.*</script>"
                hStr = reg.Replace(HTML, "")
'                保存.桌面 hStr
                Dim trr(), x%, y%
                With CreateObject("htmlfile")
                    .Open
                    .Write hStr
                    .Close
                    '.parentWindow.execScript "calcMark()"
                    Set tt = .getElementById("frozen-north")
                    Set tb1 = .getElementById("frozen-west")
                    Set tb2 = .getElementById("frozen-center")
                End With
                
                Set trs1 = tb1.getElementsByTagName("tr")
                Set trs2 = tb2.getElementsByTagName("tr")
                
                If Trim(trs1(0).Cells(0).innerText) <> "" Then
                
                    ReDim trr(1 To trs1.length, 1 To 17)
                    For x = 0 To trs1.length - 1
                        trr(x + 1, 1) = trs1(0).Cells(0).innerText '股道
                        
                        '第一个tb
                        f = IIf(x, 2, 0)
                        trr(x + 1, 2) = trs1(x).Cells(3 - f).innerText '序号
                        trr(x + 1, 3) = trs1(x).Cells(6 - f).innerText '车种
                        trr(x + 1, 4) = trs1(x).Cells(5 - f).innerText '车号
                        '第二个tb
                        trr(x + 1, 5) = trs2(x).Cells(0).innerText
                        trr(x + 1, 6) = trs2(x).Cells(1).innerText
                        trr(x + 1, 7) = trs2(x).Cells(3).innerText
                        trr(x + 1, 8) = trs2(x).Cells(4).innerText
                        trr(x + 1, 9) = trs2(x).Cells(5).innerText
                        trr(x + 1, 10) = trs2(x).Cells(9).innerText
                        trr(x + 1, 11) = trs2(x).Cells(15).innerText
                        trr(x + 1, 12) = trs2(x).Cells(12).innerText
                        trr(x + 1, 13) = trs2(x).Cells(16).innerText
                        trr(x + 1, 14) = trs2(x).Cells(10).innerText
                        trr(x + 1, 15) = trs2(x).Cells(13).innerText 'trs2(x).Cells(18).innerText
                        trr(x + 1, 16) = trs2(x).Cells(14).innerText
                        trr(x + 1, 17) = ""
                    Next
                    With Sheet7
                        r = .Cells(.Rows.count, "A").End(3).Row + 1
                        .Cells(r, "A").Resize(UBound(trr, 1), UBound(trr, 2)) = trr
                    End With
                Else
                    Debug.Print dd
                End If
                ds.Remove dd
            End If
        Next
    Loop Until ds.count = 0

End Sub

Sub dl()
    Dim xhtp As Object
    url = "http://10.190.170.50:8080/api/call/PassportSvr.LoginForSSO"
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1")
    With xhtp
        .Open "POST", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "Origin", "http://10.190.170.50:8080"
        .setRequestHeader "callid", "f9792488-0cb8-f6f6-b49d-aac080243f12"
        .setRequestHeader "IDApp", "%E8%BF%90%E8%BE%93%E7%94%9F%E4%BA%A7%E6%8C%87%E6%8C%A5%E5%B9%B3%E5%8F%B0"
        
        data = "{""loginName"":""qzcwd_suzj"",""password"":""88CrlGHC9zpOs825Lkg8Pg=="",""client"":""jO130wA4R40o8jKvjQVyYg==""}"
        .send data
        
        res = .responseText
    End With
End Sub


Sub DES加密()
    'var key = CryptoJS.enc.Utf8.parse("abcdefgabcdefg12");
    'var srcs = CryptoJS.enc.Utf8.parse(password);
    'password = CryptoJS.AES.encrypt(srcs, key, {mode:CryptoJS.mode.ECB,padding: CryptoJS.pad.Pkcs7});
    'conditionJson = "searchType=1&code=" & code & "&password=" & password
    'conditionJson: searchType=1&code=qzgyz&password=CKYHKbJZPMgCcbdapcGbDQ==
    'function reloadCode() {randomTime = new Date().getTime();}

    Dim oDom As Object, ow As Object, result As String, pathDes As String
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    pathDes = "http://10.190.168.62:80/nnjzj/resources/script/common/des.js" '网络读取
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "get", pathDes, False
        .send
    '        Debug.Print .responseText
        DES = .responseText
        ow.execScript DES
        password = ow.strEnc("http://10.190.188.11:8085/#/gdnxcqk", "mnd2-=ad")
        pw = ow.strToBt(password)
    End With
End Sub


