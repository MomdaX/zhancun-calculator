Attribute VB_Name = "调车计划"
Function 登录()

    Dim oDom As Object, ow As Object, result As String
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow

    Dim pathAes As String, AES As String, js As Object, res As String, token As String
    ow.execScript 读取文件("H:\Momda\VBA_大文本缓存\调车计划HTML\登录RSA加密.txt")
    
    密码 = ow.rsa("Qzcwd@123")
    t = ow.eval("new Date().getTime()") '时间戳
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = "http://10.190.48.4:8080/webroot/decision/login"
    data = "{""username"":""曾祥昊"",""password"":""" & 密码 & """,""validity"":-1,""sliderToken"":"""",""origin"":"""",""encrypted"":true}"
    'Debug.Print data
    With xhtp
'        .Open "GET", "http://10.190.48.4:8080/webroot/decision/login/slider/info?_=" & t, False
'        .setRequestHeader "Accept", "application/json, text/javascript, */*; q=0.01"
'        .setRequestHeader "Connection", "keep-alive"
'        .setRequestHeader "Cookie", "tenantId=default"
'        .setRequestHeader "X-Requested-With", "XMLHttpRequest"
'        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
'        .setRequestHeader "Content-Type", "application/json"
'        .setRequestHeader "Referer", "http://10.190.48.4:8080/webroot/decision/login"
'        .send
'        res = .responseText
'        ow.execScript "var js =" & res & ";"
'        imageId = ow.eval("js.data.imageId")
'
'        .Open "GET", "http://10.190.48.4:8080/webroot/decision/login/image/" & imageId, False
'        .setRequestHeader "Accept", "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8"
'        .setRequestHeader "Connection", "keep-alive"
'        .setRequestHeader "Cookie", "tenantId=default"
'        .setRequestHeader "Referer", "http://10.190.48.4:8080/webroot/decision/login"
'        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
'        .send
'        保存.图片byt .responseBody, "C:\Users\Administrator\Desktop\验证码.jpg"
'        x = 缺口图片验证码识别.JPG模式(.responseBody)
        
        .Open "POST", url, False
        .setRequestHeader "Accept", "application/json, text/javascript, */*; q=0.01"
        .setRequestHeader "Connection", "keep-alive"
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "transencryptlevel", "1"
        .setRequestHeader "Origin", "http://10.190.48.4:8080"
        .setRequestHeader "Referer", "http://10.190.48.4:8080/webroot/decision/login"
        .send data

        res = .responseText
        ow.execScript "var js =" & res & ";"
        token = ow.eval("js.data.accessToken")
        
        '列表的sessionID
        url = "http://10.190.48.4:8080/webroot/decision/v10/entry/access/44ed73de-99dd-48c7-a97c-aa6a398c0d8a?width=1920&height=1080"
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "Cookie", "tenantId=default; fine_remember_login=-1; fine_auth_token=" & token
        .setRequestHeader "Referer", "http://10.190.48.4:8080/webroot/decision"
        .send
        sessionID = Split(Split(.getResponseHeader("Set-Cookie"), ";")(0), "=")(1)
        'Debug.Print .responseText
        
        SaveSetting "Momda", "调车计划", "sessionID", sessionID
        SaveSetting "Momda", "调车计划", "token", token
    End With
    登录 = token
End Function

Function byteToB64(ByRef bData() As Byte) As String
    Dim xml As Object
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    With xml.createElement("b64")
        .DataType = "bin.base64"
        .nodeTypedValue = bData
        byteToB64 = .text
    End With
    Set xml = Nothing
End Function

Function 刷新计划列表()
    Dim oDom As Object, ow As Object, result As String, urr()
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    
    Dim re As Object, sst As String, est As String, htmlstr As String, hStr As String, HTML As String, path As String
    Set re = CreateObject("VBScript.RegExp")
    
    re.Pattern = "<script>\s*\w+\._p\s*=\s*_\w*\(\);\s*\w+\._p\.\w+\s*=\s*\d+;\s*\w+\._p\.\w+\s*=\s*\d+;\s*delete\s+\w+\._p\s*;\s*</script\s*>\s*"
    re.IgnoreCase = True
    re.Global = True   '删多处
    
    ow.execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
    
    查询参数 = "{""ZMLM"":""QVZ"",""CZM"":""钦州港"",""ENTRYID"":""44ed73de-99dd-48c7-a97c-aa6a398c0d8a"",""DECISIONTEMPLATE"":""true"",""WIDTH"":""1920"",""ENCODINGFILTER.FILTERED"":""true"",""HEIGHT"":""1080"",""HEADERADDED"":""true"",""__ESCAPE_XSS__"":""false"",""URICHECKFILTER.FILTERED"":""true"",""COUNTER.FILTERED"":""true"",""ENCRYPTIONLEVELFILTER.FILTERED"":""true"",""SUPERUSER"":""[530a7191-c431-48de-9d28-83c68dee47d2]"",""FORMLETNAME"":""YSBGLPT/DCJH/QZCWD/钦州车务段调车作业计划.frm"",""ADDRESS"":""/10.191.152.60:49974"",""COOKIE-CHECK.FILTERED"":""true"",""animateType"":""NONE""}"
    data = "widgetName=REPORT1&pageIndex=1&__parameters__=" & ow.encode(查询参数) & "&noCache=lazy&simpleJson=false&arrayJson=false"
    'Debug.Print data
        
    sid = GetSetting("Momda", "调车计划", "sessionID")
    token = GetSetting("Momda", "调车计划", "token")
    
重新查询:
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1")
    With xhtp
        '列表查询
        ow.execScript "function reloadCode(){return new Date().getTime();}"
        randomTime = ow.reloadCode("")
        url = "http://10.190.48.4:8080/webroot/decision/view/fit/form/load/content?_=" & randomTime
        .Open "POST", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .setRequestHeader "Authorization", "Bearer " & token
        .setRequestHeader "Origin", "http://10.190.48.4:8080"
        .setRequestHeader "Referer", "http://10.190.48.4:8080/webroot/decision/v10/entry/access/44ed73de-99dd-48c7-a97c-aa6a398c0d8a"
        .setRequestHeader "sessionID", sid
        .send data
        'Debug.Print .responseText
        res = .responseText
        
        '一份计划的sessionID
        'http://10.190.48.4:8080/webroot/decision/view/report?viewlet=/YSBGLPT/DCJH/NNTLJ/调车作业通知单.cpt&__parameters__={"__pi__":true,"KSTIME":"19:30","GJHH":"A002","DJH":"3","LCCC":"7013","ZYGD":"X1","JSTIME":"23:00","DCZ":"王河","BZR":"蒋静静","GJHNAME":"QVZ35883","BZTIME":"2025-11-04 19:07:13","ZYSX":"1.压信号（绝缘）调车注意确认进路;6.编成列车，注意检查前后3位关门车。;7.由无电区向有电区调车，注意站位安全。;13.进港线轨道衡过磅限速35公里。;14.货场轨道衡过磅限速18公里。;","ZYMC":"取送","fine_hyperlink":"3fa9fc83-fc53-4c3a-b3f0-cfc3a7b8fd05"}&_=1762276727531'
        Dim i, n, x, y, c, k, l
        
        On Error Resume Next
        ow.execScript "var tbjs =" & res & ";"
        Set params = ow.eval("tbjs[""pageResult""]")
        On Error GoTo 0
        
        '检查登录是否成功
        If params = "" Then
            登录
            sid = GetSetting("Momda", "调车计划", "sessionID")
            token = GetSetting("Momda", "调车计划", "token")
            GoTo 重新查询:
        End If
        
        l = CallByName(params, "length", VbGet)

        With Sheet15
        
            If l < 3 Then
                MsgBox "无计划！！！"
                End
            End If
            
            ReDim urr(1 To l - 2, 1 To 7)
            
            Application.EnableEvents = False
            
            With .Range("M3").CurrentRegion.Offset(2)
                .Cells.ClearComments '清除批注
                .ClearContents '清除内容
                .Interior.Color = 16777215 '清除颜色
            End With
            For x = 2 To l - 1
                Set param = CallByName(params, x, VbGet)
                
                For Each c In Array(0, 4, 5, 6, 7, 10)
                
                    Set pm = CallByName(param, c, VbGet)
                    ps = CallByName(pm, "value", VbGet)
                    ow.execScript "var pv =" & ps & ";"
                    .Cells(x + 2, 13 + y).value = ow.eval("pv.value")
                    
                    y = y + 1
                Next
                
                Set parameters = ow.eval("tbjs[""pageResult""][" & x & "][0][""cellhyperlink""][""nxNameJavascriptGroup""][""javaScriptGroup""][0][""javaScript""][""parameters""]")
                ow.execScript "var reportPath = ""/YSBGLPT/DCJH/NNTLJ/调车作业通知单.cpt"";var encodedViewlet = encodeURIComponent(reportPath);"
                
                On Error Resume Next
                s = """__pi__"":true"
                For Each param In parameters
                   s = s & "," & """" & param.name & """:""" & param.value & """"
                Next
                On Error GoTo 0
                
                u = "http://10.190.48.4:8080/webroot/decision/view/report?"
                v = "viewlet=" & ow.eval("encodedViewlet")
                p = "__parameters__=" & ow.encode("{" & s & "}")
                t = "_=" & ow.reloadCode("")
                Rf = u & v & "&" & p & "&" & t
                
                
                Set Djh = CallByName(parameters, "2", VbGet): urr(x - 1, 1) = CallByName(Djh, "value", VbGet) '调机号
                Set Jhh = CallByName(parameters, "1", VbGet): urr(x - 1, 2) = CallByName(Jhh, "value", VbGet) '计划号
                Set Kssj = CallByName(parameters, "0", VbGet): urr(x - 1, 3) = CallByName(Kssj, "value", VbGet) '开始时间
                Set Gssj = CallByName(parameters, "5", VbGet): urr(x - 1, 4) = CallByName(Gssj, "value", VbGet) '结束时间
                urr(x - 1, 5) = Rf '计划地址
                
                HTML = 计划查询(Rf)
                htmlstr = re.Replace(HTML, "")
                sst = 读取("H:\Momda\VBA_大文本缓存\调车计划HTML\Html_s.txt", "UTF-8")
                est = 读取("H:\Momda\VBA_大文本缓存\调车计划HTML\Html_e.txt", "UTF-8")
                
                hStr = sst & htmlstr & est
                On Error GoTo toEnd
                
                '识别
                Dim rut As Boolean, span As Object, rng As Range
                With CreateObject("htmlfile")
                    .Open
                    .Write hStr
                    .Close
                    '.parentWindow.execScript "calcMark()"
                    Set span = .getElementById("auditResultSpan")
                End With

                '取出结果
                itxt = span.innerText
                rut = span.getAttribute("data-result")
            
                '保存
                path = "H:\Momda\VBA_大文本缓存\调车计划HTML\缓存\" & urr(x - 1, 2) & ".html" '△2
                urr(x - 1, 7) = path
                With CreateObject("ADODB.Stream")
                    .Type = 2
                    .Charset = "UTF-8"
                    .Open
                    .WriteText hStr
                    .SaveToFile path, 2
                    .Close
                End With
                
                Set rng = .Cells(x + 2, 13)
                With rng
                    .AddComment path
                    '.Comment.Visible = False'关闭批注显示
                    .Interior.ColorIndex = IIf(rut, -4142, 44)
                End With
                y = 0
            Next
        Application.EnableEvents = True
        End With
        
        'Sheet15.Comment.Visible = False
    End With
    
    刷新计划列表 = urr
    
    Exit Function
    
toEnd:
    Application.EnableEvents = True
    End
End Function


Function 计划查询(url) As String

    Dim oDom As Object, ow As Object, result As String
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    
    ow.execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
    ow.execScript "function reloadCode(){return new Date().getTime();}"
     
    token = GetSetting("Momda", "调车计划", "token")
    
重新查询:
    t = "_=" & ow.reloadCode("")
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1")
    With xhtp
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "Cookie", "fineMarkId=ccc8a1faed99c617c91f4180c2c2a633; tenantId=default; fine_remember_login=-1; fine_auth_token=" & token
        .setRequestHeader "Referer", "http://10.190.48.4:8080/webroot/decision/v10/entry/access/44ed73de-99dd-48c7-a97c-aa6a398c0d8a?width=1920&height=1080"
        .send
        
        On Error Resume Next
        sessionID = Split(Split(.getResponseHeader("Set-Cookie"), ";")(0), "=")(1)
        On Error GoTo 0
        
        If sessionID = "" Then
            token = 登录
            GoTo 重新查询:
        End If
        
        '计划查询
        cid = "195962aadacd01877d7ebf204c094b0e#" & t & "#" & md5(t)
        cid = ow.encode(cid)
        url = "http://10.190.48.4:8080/webroot/decision/view/report?" & t & "&__boxModel__=true&op=page_content&pn=1&cid=" & cid & "&__webpage__=true&_paperWidth=1920&_paperHeight=1080&__fit__=false"
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .setRequestHeader "Authorization", "Bearer " & token
        .setRequestHeader "Origin", "http://10.190.48.4:8080"
        .setRequestHeader "Referer", url
        .setRequestHeader "sessionID", sessionID
        .send
        'Debug.Print .responseText
        res = .responseText
        
        ow.execScript "var rjs =" & res & ";"
        retHTML = ow.eval("rjs[""html""]")
    End With
    计划查询 = retHTML
End Function

Function 显示计划(HTML As String)
    Dim re As Object, sst As String, est As String, htmlstr As String
    Set re = CreateObject("VBScript.RegExp")
    
    're.Pattern = "<script>\s*FR\._p\s*=\s*_g\(\);\s*FR\._p\.currentPageIndex\s*=\s*1;\s*FR\._p\.reportTotalPage\s*=\s*2;\s*delete\s+FR\._p;\s*</script>\s*"
    re.Pattern = "<script>\s*\w+\._p\s*=\s*_\w*\(\);\s*\w+\._p\.\w+\s*=\s*\d+;\s*\w+\._p\.\w+\s*=\s*\d+;\s*delete\s+\w+\._p\s*;\s*</script\s*>\s*"
    re.IgnoreCase = True
    re.Global = True   '如果可能出现多处，就删多处
    
    htmlstr = re.Replace(HTML, "")
    
    sst = 读取("H:\Momda\VBA_大文本缓存\调车计划HTML\Html_s.txt", "GBK")
    est = 读取("H:\Momda\VBA_大文本缓存\调车计划HTML\Html_eys.txt", "GBK")
    
    htmlstr = sst & htmlstr & est

    '加载到 WebBrowser
    With 计划HTML.WebBrowser1
        .width = 560
        .navigate "about:blank"
        Do While .Busy Or .readyState <> 4
            DoEvents
        Loop
        .document.Open
        .document.Write htmlstr
        .document.Close
    End With
    
    计划HTML.Show
    'DoEvents

End Function

Function 加载js(urlPath As String)
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "get", urlPath, False
        .send
'        Debug.Print .responseText
        加载js = .responseText
    End With
End Function

Function md5(ae) As String
    Dim md5Hasher As Object
    Set md5Hasher = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    
    Dim pwdBytes() As Byte
    pwdBytes = StrConv(ae, vbFromUnicode)
    
    Dim hashBytes() As Byte
    hashBytes = md5Hasher.ComputeHash_2(pwdBytes)

    Dim i As Long
    Dim tempStr As String
    For i = LBound(hashBytes) To UBound(hashBytes)
        tempStr = tempStr & Right("0" & Hex(hashBytes(i)), 2)
    Next i
    
    md5 = LCase(Right(tempStr, 8))
End Function

Function URL编码(strText)
    Static objHtmlfile As Object
    If objHtmlfile Is Nothing Then
        Set objHtmlfile = CreateObject("htmlfile")
        objHtmlfile.parentWindow.execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
        'objHtmlfile.parentWindow.execScript "function Decode(s) {return decodeURIComponent(s)}", "jscript"
    End If
    URL编码 = objHtmlfile.parentWindow.encode(strText)
End Function

Function 读取文件(fpth As String) As String
    Dim txt As String
'    filePath = "E:\桌面\Python\des加密.js" '文件读取
    Open fpth For Binary As #1
    txt = Input$(LOF(1), #1)
    Close #1 ' 关闭文件
    读取文件 = txt
End Function

Function 读取(path As String, gs As String) As String
    Dim hStr As String
    Set objStream = CreateObject("ADODB.Stream")
    With objStream
        .Charset = gs ' 假设文件是 UTF-8 编码（如果是 GBK 可改为 "GBK"）
        .Open
        .LoadFromFile path
        hStr = .ReadText ' 读取文本内容
        .Close
    End With
    Set objStream = Nothing
    
    读取 = hStr
End Function

Function 读取调车钩数()
    Dim hTxt As String, f As Long
    
    urs = 刷新计划列表
    With Sheet15
    
        Application.EnableEvents = False
        .Range("C3:I7,C9:I13,C15:I19").ClearContents '清除内容
        .Range("B2").CurrentRegion.ClearComments '清除批注
        
        Application.EnableEvents = True
        
        For n = 1 To UBound(urs)
        
            url = urs(n, 5)
            hTxt = 计划查询(url)
        
            Dim doc As Object: Set doc = CreateObject("htmlfile")
            doc.Open: doc.Write hTxt: doc.Close
            
            Dim tb As Object
            Set tb = doc.getElementsByTagName("tbody")(0)
            
            Dim trs As Object
            Set trs = tb.getElementsByTagName("tr")
            
            Dim rc As Long: rc = trs.length
            zgs = rc - 12 '主钩数
            
            Dim cnt As Long: cnt = 0
            Dim sRow As Long: sRow = 6
            Dim eRow As Long: eRow = rc - 7
            
            Dim i As Long, tds As Object
            Dim trkCell As Object, cntCell As Object
            Dim trkVal As String, cntVal As String, nTrk As String
            
            For i = sRow To eRow '副钩数
                Set tds = trs(i).getElementsByTagName("td")
                
                f = -(i = 6) '布尔值转数字
                
                Set trkCell = tds(2 + f)
                trkVal = Trim(trkCell.innerText)
                
                Set cntCell = tds(4 + f)
                cntVal = Trim(cntCell.innerText)
                
                nTrk = NormTrk(trkVal)
                
                If IsNumeric(nTrk) Then
                    Dim tNum As Integer: tNum = CInt(nTrk)
                    If tNum >= 1 And tNum <= 15 And cntVal = "0" Then
                        cnt = cnt + 1
                        'Debug.Print "第" & i - 5 & "钩: " & trkVal & "道， 车数为:" & cntVal
                    End If
                End If
        
            Next i
            
            urs(n, 5) = zgs '主钩数
            urs(n, 6) = cnt '副钩数
            cnt = 0
            
            '写表
            If urs(n, 1) = 3 Then
                x = 20
            ElseIf urs(n, 1) = 2 Then
                x = 14
            ElseIf urs(n, 1) = 1 Then
                x = 8
            Else
                x = ""
            End If
            
            Set ess = trs(4).getElementsByTagName("td")(1)
            esstr = ess.innerText
            'Debug.Print Len(esstr), esstr
            
            If x <> "" And Len(esstr) > 10 Then
                r = .Cells(x, 3).End(3).Row + 1
                .Cells(r, "C").value = urs(n, 2): .Cells(r, "C").AddComment urs(n, 7)
                .Cells(r, "D").value = urs(n, 5)
                .Cells(r, "E").value = urs(n, 3)
                .Cells(r, "F").value = urs(n, 4)
                .Cells(r, "I").value = urs(n, 6)
'                x = ""
            End If
        Next
    End With
End Function

Function NormTrk(v As String) As String
    If Len(v) >= 2 And Left(v, 1) = "X" Then
        Dim sfx As String: sfx = Right(v, Len(v) - 1)
        If IsNumeric(sfx) Then
            NormTrk = sfx: Exit Function
        End If
    End If
    NormTrk = v
End Function

Function ReadFile(p As String) As String
    Dim fso As Object, f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(p, 1)
    ReadFile = f.ReadAll
    f.Close
End Function

