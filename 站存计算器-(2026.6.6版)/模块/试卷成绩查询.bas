Attribute VB_Name = "试卷成绩查询"
Private Declare Function InternetSetCookie Lib "wininet.dll" Alias "InternetSetCookieA" (ByVal lpszUrl As String, ByVal lpszCookieName As String, ByVal lpszCookieData As String) As Long
Sub 设置IE浏览器cookie()
    Cookie = Split(Sheet20.Range("F3").value, "=")
   url = Selection.Comment.text
   Debug.Print url
    InternetSetCookie url, "login_loginid", ""
    InternetSetCookie url, "login_username", ""
    InternetSetCookie url, "login_rem_username", "false"
'   InternetSetCookie url, "learning.session.id", "9e60711f-1d06-431f-95c6-2d9efc8d9756"
    InternetSetCookie url, Cookie(0), Cookie(1)
End Sub

Function 试卷试题(rn As Range)
' 创建 HTML 文档对象和窗口对象
    Dim oDom As Object, ow As Object, result As String, pathDes As String
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    pathDes = "http://10.190.168.62:80/nnjzj/resources/script/common/des.js" '网络读取
    DES = 加载js(pathDes)
    ow.execScript DES
    
重试:
    With Sheet20
        试卷编号 = rn.Cells(3)
        secreKkey = .Cells(2, "F")
        Cookie = .Cells(3, "F")
        '表格初始化
'        Sheet21.Range("H1").CurrentRegion.Offset(1).Clear'注释后为追加模式
        Sheet21.Range("O1").value = rn.Cells(1)
    End With
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    With http
        uuid = getuuid()
        url = "http://10.190.168.62/nnjzj/exam/examreport/topiclist.html" '试卷试题地址
        data = "%24%40%24.ecryptedData=" & ow.strEnc("{""tmp8wsdF4"":""" & uuid & """,""ENTITYCLASSNAME"":"""",""ttop003"":""" & 试卷编号 & """}", secreKkey)
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=UTF-8"
        .setRequestHeader "Origin", "http://10.190.168.62"
        .setRequestHeader "Referer", "http://10.190.168.62/nnjzj/manager/index.html"
        .setRequestHeader "Cookie", Cookie
        .send data
        Debug.Print .responseText
        
        If InStr(.responseText, "Sorry..运行异常") Then
            MsgBox "Sorry..运行异常"
            Debug.Print "Sorry..运行异常"
            Exit Function
        ElseIf InStr(.responseText, "Sorry..登录信息失效") Then
'            MsgBox "Sorry..登录信息失效！"
            Debug.Print "Sorry..登录信息失效！,正在重试！"
            参数
            GoTo 重试
        End If

        Dim Json
        Set Json = JsonConverter.ParseJson(.responseText)
        
        '标题
        Sheet21.Range("A1:L1") = Array("序号", "试题类型", "试题内容", "正确选项", "A", "B", "C", "D", "E", "F", "G", "H")
            '设置格式
        With Sheet21.Range(Sheet21.Cells(2, "A"), Sheet21.Cells(Sheet21.Rows.count, "L"))
            .NumberFormatLocal = "G/通用格式"
            .HorizontalAlignment = xlCenter '居中
            .VerticalAlignment = xlCenter '居中
        End With
        
        rr = Sheet21.Cells(Rows.count, "A").End(3).Row
        For Each da In Json("testtxList")
            With Sheet21
                r = .Cells(.Rows.count, "A").End(3).Row + 1
                .Cells(r, "A") = "=row()-1"
                .Cells(r, "B") = da("ttx004")
                .Cells(r, "C") = da("testtopic")("ttop012")
                .Cells(r, "D") = da("testtopic")("ttop022")
                If da("ttx004") = "判断题" Then
                    .Range(.Cells(r, "E"), .Cells(r, "L")) = Array("对", "错")
                Else
                    .Range(.Cells(r, "E"), .Cells(r, "L")) = Split(da("testtopic")("ttop019"), "$$")
                End If
            End With
        Next

    End With
    
    With Sheet21.Range("A1:L1").CurrentRegion
        .Borders.LineStyle = xlContinuous
        .WrapText = True '自动换行
        Sheet21.Visible = xlSheetVisible
        Sheet21.Select
        .Replace "#N/A", ""
'        .Range("M1:U1").EntireColumn.AutoFit
    升序 Sheet21, .Rows(1), .Rows(1).Cells(2), 2 '表对象，标题行，排序列单元格，1升2降
    End With
    
    MsgBox "追加模式!" & Chr(10) & "提取到" & r - rr & "条题目！"
End Function
Function 成绩列表(rn As Range)
' 创建 HTML 文档对象和窗口对象
    Dim http As Object, xhtp As Object
    Dim oDom As Object, ow As Object, result As String, pathDes As String, 试卷URL As String
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    pathDes = "http://10.190.168.62:80/nnjzj/resources/script/common/des.js" '网络读取
    DES = 加载js(pathDes)
    ow.execScript DES
重试:
    With Sheet20
        试卷编号 = rn.Cells(3).value
        secreKkey = .Cells(2, "F")
        Cookie = .Cells(3, "F")
        '表格初始化
        .Range("H1").CurrentRegion.Offset(1).Clear
    End With
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    With http
        uuid = getuuid()
        url = "http://10.190.168.62/nnjzj/exam/examscorerecord/findByPage.html" '考试记录地址
        data = "%24%40%24.ecryptedData=" & ow.strEnc("{""tmp8wsdF4"":""" & uuid & """,""ENTITYCLASSNAME"":""ExamScoreRecordEntity"",""pageSize"":0,""currentPage"":1,""customColumn"":""egra008,egra013,egra022,result,user004,user005,organname,egratime,joinmode"",""customColumnName"":""参考时间,初始成绩,成绩,考试结果,账号,姓名,所属单位,考试用时,参考方式"",""egra004"":""" & 试卷编号 & """,""userqufen"":1,""orderType"":1}", secreKkey) '原pageSize=50
        .Open "POST", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=UTF-8"
        .setRequestHeader "Origin", "http://10.190.168.62"
        .setRequestHeader "Referer", "http://10.190.168.62/nnjzj/manager/index.html"
        .setRequestHeader "Cookie", Cookie
        .send data
'        Debug.Print .responseText

        If InStr(.responseText, "Sorry..运行异常") Then
            MsgBox "Sorry..运行异常"
            Debug.Print "Sorry..运行异常"
            Exit Function
        ElseIf InStr(.responseText, "Sorry..登录信息失效") Then
'            MsgBox "Sorry..登录信息失效！"
            Debug.Print "Sorry..登录信息失效！,正在重试！"
            参数
            GoTo 重试
        End If
        aa = .responseText
        Dim Json
        Set Json = JsonConverter.ParseJson(.responseText)
        '标题
        Sheet20.Range("H1:P1") = Array("参考时间", "账号", "姓名", "所属单位", "初始成绩", "成绩", "考试用时", "考试结果", "参考方式", "Egid")
            '设置格式
        With Sheet20.Range(Sheet20.Cells(2, "H"), Sheet20.Cells(Sheet20.Rows.count, "P"))
            .Columns(1).NumberFormatLocal = "yyyy-m-d h:mm;@"
            .Columns(2).NumberFormatLocal = "@"
            .Columns(3).NumberFormatLocal = "G/通用格式"
            .Columns(4).NumberFormatLocal = "G/通用格式"
            .Columns(5).NumberFormatLocal = "0.0"
            .Columns(6).NumberFormatLocal = "0.0"
            .Columns(7).NumberFormatLocal = "G/通用格式"
            .Columns(8).NumberFormatLocal = "G/通用格式"
            .Columns(9).NumberFormatLocal = "G/通用格式"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        For Each da In Json("dataSource")
            With Sheet20
                r = .Cells(.Rows.count, "I").End(3).Row + 1
                试卷URL = "http://10.190.168.62/nnjzj/user/exam/view.html?tpid=" & 试卷编号 & "&egid=" & da("egra001") '试卷URL
'                下载URL = "http://10.190.168.62/nnjzj/user/exam/topic_answer.html?tpid=" & 试卷编号 & "&egid=" & da("egra001") '试卷URL
                .Cells(r, "H") = Format(DateAdd("s", da("egra008") / 1000, #1/1/1970#), "yyyy-mm-dd hh:mm:ss")
                .Cells(r, "I") = da("user004")
                .Cells(r, "J") = da("user005")
                .Cells(r, "K") = da("organname")
                .Cells(r, "L") = da("egra013")
                .Cells(r, "M") = da("egra022")
                .Cells(r, "N") = da("egratime")
                .Cells(r, "O") = IIf(da("result") = 1, "通过", "未通过")
                .Cells(r, "P") = IIf(da("joinmode"), "APP", "PC")
                .Cells(r, "Q") = da("egra001") '用户ID
                
                '设置批注
                With .Cells(r, "J")
                    If Not .Comment Is Nothing Then .Comment.Delete
                    
                    '下载试卷到批注中
'                    With CreateObject("WinHttp.WinHttpRequest.5.1")
'                        .Open "GET", 下载URL, False
'                        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
'                        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=UTF-8"
'                        .setRequestHeader "Origin", "http://10.190.168.62"
'                        .setRequestHeader "Referer", "http://10.190.168.62/nnjzj/manager/index.html"
'                        .setRequestHeader "Cookie", Cookie
'                        .send
'                        res = .responseText
'                    End With
                    .AddComment 试卷URL
                End With
'                .Hyperlinks.Add .Cells(r, "Q"), 试卷URL, , , "查看试卷"'设置超链接
            End With
        Next
        
        '排序
        
        
    End With
    With Sheet20
'        .Range("A1").CurrentRegion.Offset(1).VerticalAlignment = 4
        .Range("H1").CurrentRegion.Borders.LineStyle = xlContinuous
        .Range("H1:P1").EntireColumn.AutoFit
        升序 Sheet20, .Range("H1:P1"), .Range("L1"), 2 '表对象，标题行，排序列单元格，1升2降
    End With
    
End Function

Sub 试卷列表()
    ' 创建 HTML 文档对象和窗口对象
    Dim oDom As Object, ow As Object, result As String, pathDes As String, 参与 As String
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    Dim Cookie$, secreKkey$, uuid$, DES$, data$, url$, da, r%
    'Dim http As Object
    Dim pathAes As String, 日期 As String, cs As Variant
    
重试:
    With Sheet20
        secreKkey = .Cells(2, "F")
        Cookie = .Cells(3, "F")
    End With
'    cs = 参数()
'    Cookie = cs(0)
'    secreKkey = cs(1)
    uuid = getuuid()
    
    '表格初始化
    Sheet20.Range("A1").CurrentRegion.Offset(1).Clear
    
    pathDes = "http://10.190.168.62:80/nnjzj/resources/script/common/des.js" '网络读取
    DES = 加载js(pathDes)
    ow.execScript DES
    
    'data加密需要密文和密钥
    日期1 = VBA.Format(DateAdd("m", -6, Now), "yyyy-mm-dd hh:mm:ss")
    日期2 = VBA.Format(Now(), "yyyy-mm-dd hh:mm:ss")
    If Sheet20.Range("F2").value = "" Or Sheet20.Range("F3").value = "" Then
        Call 参数
        GoTo 重试
    ElseIf Sheet20.Range("F4").value = 1 Or Sheet20.Range("F4").value = 2 Then
        参与 = Sheet20.Range("F4").value
        url = "http://10.190.168.62/nnjzj/exam/examreport/findByPage.html" '考试监控地址
        '//本单位参与的试卷：userqufen=2
        'data = "%24%40%24.ecryptedData=" & ow.strEnc("{""tmp8wsdF4"":""" & uuid & """,""ENTITYCLASSNAME"":""ExamReportEntity"",""pageSize"":0,""currentPage"":1,""customColumn"":""tact004,tact003name,user005,urre010,examnum,onlinenum,scorenum,noexamnum,rate,passRate"",""customColumnName"":""考试活动名称,主办单位,发布人,发布时间,应考人数,在考人数,已考人数,缺考人数,参考率,通过率"",""menuType"":""1"",""userqufen"":""" & 参与 & """,""BETWEEN_START_urre010"":""" & 日期1 & """,""BETWEEN_END_urre010"":""" & 日期2 & """}", secreKkey)
        data = "%24%40%24.ecryptedData=" & ow.strEnc("{""tmp8wsdF4"":""" & uuid & """,""ENTITYCLASSNAME"":""ExamReportEntity"",""pageSize"":""0"",""currentPage"":""1"",""tact002"":""-3"",""tact005"":""1"",""nodetId"":""ztree_1"",""isParent"":true,""menuType"":""0"",""userqufen"":""" & 参与 & """,""BETWEEN_START_urre010"":""" & 日期1 & """,""BETWEEN_END_urre010"":""" & 日期2 & """,""sortTableColumn"":""urre010"",""sortTableType"":""asc""}", secreKkey) '2025-02-20
    ElseIf Sheet20.Range("F4").value = 3 Then
        '//所有单位的试卷
        url = "http://10.190.168.62/nnjzj/exam/testactivity/findByPage.html" '所有单位的试卷
        data = "%24%40%24.ecryptedData=" & ow.strEnc("{""tmp8wsdF4"":""" & uuid & """,""ENTITYCLASSNAME"":""TestactivityEntity"",""pageSize"":100,""currentPage"":1,""customColumn"":""tact038,tact004,organname,ksfl,tact059,tact049,urre010,tact006,tact007"",""isArchive"":""1"",""tact002"":""-3"",""customColumnName"":""状态,考试名称,主办单位,考试分类,是否为抽考,创建人,最后修改时间,考试开始时间,考试结束时间"",""tact005"":""1"",""nodetId"":""ztree_1"",""isParent"":true}", secreKkey) '"5Ad8XJiY"，原pageSize=10,
    Else
        参与 = 1
        url = "http://10.190.168.62/nnjzj/exam/examreport/findByPage.html" '考试监控地址
        '//本单位参与的试卷：userqufen=2
    '    data = "%24%40%24.ecryptedData=" & ow.strEnc("{""tmp8wsdF4"":""" & uuid & """,""ENTITYCLASSNAME"":""ExamReportEntity"",""pageSize"":0,""currentPage"":1,""tact005"":1,""flag"":2,""start_urre010"":""" & 日期1 & """,""end_urre010"":""" & 日期2 & """,""menuType"":""1"",""userqufen"":""" & 参与 & """,""BETWEEN_START_urre010"":""" & 日期1 & """,""BETWEEN_END_urre010"":""" & 日期2 & """}", secreKkey)
        data = "%24%40%24.ecryptedData=" & ow.strEnc("{""tmp8wsdF4"":""" & uuid & """,""ENTITYCLASSNAME"":""ExamReportEntity"",""pageSize"":0,""currentPage"":1,""customColumn"":""tact004,tact003name,user005,urre010,examnum,onlinenum,scorenum,noexamnum,rate,passRate"",""customColumnName"":""考试活动名称,主办单位,发布人,发布时间,应考人数,在考人数,已考人数,缺考人数,参考率,通过率"",""menuType"":""1"",""userqufen"":""" & 参与 & """,""BETWEEN_START_urre010"":""" & 日期1 & """,""BETWEEN_END_urre010"":""" & 日期2 & """}", secreKkey)
    End If
    
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        With http
            '获取所有试卷
            .Open "POST", url, False
            .setRequestHeader "Referer", "http://10.190.168.62/nnjzj/manager/index.html"
            .setRequestHeader "Cookie", Cookie & "; login_rem_username=false; login_username=; login_loginid="
            .setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=UTF-8"
            .setRequestHeader "Origin", "http://10.190.168.62"
            .send data
            'Debug.Print .responseText
    
        If InStr(.responseText, "Sorry..运行异常") Then
            MsgBox "Sorry..运行异常"
            Debug.Print "Sorry..运行异常"
            Exit Sub
        ElseIf InStr(.responseText, "Sorry..登录信息失效") Then
'            MsgBox "Sorry..登录信息失效！"
            Debug.Print "Sorry..登录信息失效！,正在重试！"
            参数
            GoTo 重试
        ElseIf InStr(.responseText, """totalRow"":0") Then
            MsgBox "列表为空"
            Debug.Print "列表为空"
            Exit Sub
        End If
            
            '设置格式
            With Sheet20
                .Columns(1).NumberFormatLocal = "G/通用格式"
                .Columns(1).Cells(1).HorizontalAlignment = xlCenter
                With .Range(.Cells(2, 2), .Cells(.Rows.count, 4))
                    .Columns(2).NumberFormatLocal = "G/通用格式"
                    .Columns(3).NumberFormatLocal = "G/通用格式"
                    .Columns(4).NumberFormatLocal = "yyyy-m-d h:mm;@"
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With
            End With
            
            Dim Json
            Set Json = JsonConverter.ParseJson(.responseText)
            '保存.桌面 .responseText
            '标题
            Sheet20.Range("A1:D1") = Array("考试活动名称", "发布人", "ID", "发布时间")
            For Each da In Json("dataSource")
                With Sheet20
                    r = .Cells(.Rows.count, "A").End(3).Row + 1
                    .Cells(r, 1) = da("tact004")
                    .Cells(r, 2) = da("user005")
                    .Cells(r, 3) = da("tact001")
                    .Cells(r, 4) = VBA.Format(DateAdd("s", da("urre010") / 1000, #1/1/1970#), "yyyy-mm-dd hh:mm:ss")
                End With
            Next
        End With
        
        With Sheet20
    '        .Range("A1").CurrentRegion.Offset(1).VerticalAlignment = 4
            .Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous '设置边框
    '        .Range("A1:I1").EntireColumn.AutoFit
            升序 Sheet20, .Range("A1:D1"), .Range("D1"), 2 '表对象，标题行，排序列单元格，1升2降
        End With
'    MsgBox "完成！"
End Sub
Function 提取(HTML, 表达式 As String)
Dim reg As New RegExp
With reg
    .Pattern = 表达式 '"""([\d_]+)"""
    .Global = True
    .MultiLine = True
End With

Set 提取 = reg.Execute(HTML)
Set reg = Nothing
End Function

Function getuuid()
    Dim oDom As Object, ow As Object
    Set oDom = CreateObject("htmlfile")
    Set ow = oDom.parentWindow
    Dim pathuuid As String, uuid As String
'    pathuuid = "F:\Momda\试卷\DES加密\uuid.txt"
'    uuid = json(pathuuid)
    uuid = "function uuid(len,radix){var chars='0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'.split('');var uuid=[],i;radix=radix||chars.length;if(len){for(i=0;i<len;i++)uuid[i]=chars[0|Math.random()*radix];}else{var r;uuid[8]=uuid[13]=uuid[18]=uuid[23]='-';uuid[14]='4';for(i=0;i<36;i++){if(!uuid[i]){r=0|Math.random()*16;uuid[i]=chars[(i==19)?(r&0x3)|0x8:r];}}}return uuid.join('');}"
    ow.execScript uuid
    getuuid = ow.uuid(8, 26)
End Function

Function 参数() As Variant '登录
    Dim oDom As Object, ow As Object, result As String, Cookie As String
    Dim 验证码 As String, 帐号 As String, 密码 As String, num As String
    Dim http, url$, fn, data
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    ow.execScript "function Number(){return Math.random()}"
    num = VBA.Format(ow.Number(""), "0.000000000000000")
    url = "http://10.190.168.62/nnjzj/nnjzj/login/captcha.html?number=" & num
    帐号 = "qzcwd"
    密码 = "Zjk63951+"
    'DES加密
cx:
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    With http
        .Open "GET", url, False
        .setRequestHeader "Accept", "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8"
        .setRequestHeader "Referer", "http://10.190.168.62/nnjzj/nnjzj/newhome/index.html"
'        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .send
'        Debug.Print .getAllResponseHeaders
        Cookie = Split(Split(.getAllResponseHeaders(), "Set-Cookie: ")(1), ";")(0)
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
        Kill path '删除验证码图片
        
        data = "validate=" & 验证码
        url = "http://10.190.168.62/nnjzj/nnjzj/login/checkCode.html"
        .Open "POST", url, False
        .setRequestHeader "Referer", "http://10.190.168.62/nnjzj/nnjzj/newhome/index.html"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=UTF-8"
        .setRequestHeader "Origin", "http://10.190.168.62"
        .setRequestHeader "Cookie", Cookie
        .send data
'        Debug.Print .responseText

        If InStr(.responseText, "Sorry..运行异常") Then
            MsgBox "Sorry..运行异常"
            Debug.Print "Sorry..运行异常"
            If ei > 5 Then End
            Set http = Nothing
            GoTo cx
        End If

        '请求获取cookie
        'http://10.190.168.62/nnjzj/nnjzj/login/checkLogin.html
        Dim pathuuid As String, pathAes As String, uuid, AES, username, password, validate
        pathuuid = ThisWorkbook.path & "\uuid.txt"
'        uuid = json(pathuuid)
        uuid = "function uuid(len,radix){var chars='0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz'.split('');var uuid=[],i;radix=radix||chars.length;if(len){for(i=0;i<len;i++)uuid[i]=chars[0|Math.random()*radix];}else{var r;uuid[8]=uuid[13]=uuid[18]=uuid[23]='-';uuid[14]='4';for(i=0;i<36;i++){if(!uuid[i]){r=0|Math.random()*16;uuid[i]=chars[(i==19)?(r&0x3)|0x8:r];}}}return uuid.join('');}"
        ow.execScript uuid
        uuid = ow.uuid(8, 26)
        pathAes = "http://10.190.168.62:80/nnjzj/resources/script/common/aes.js" '网络读取
        AES = 加载js(pathAes)
        ow.execScript AES
        username = Replace(Replace(ow.encrypt(帐号, "A-16-Byte-keyVal"), "=", "%3D"), "/", "%2F")
        password = Replace(Replace(ow.encrypt(密码, "A-16-Byte-keyVal"), "=", "%3D"), "/", "%2F")
        validate = Replace(Replace(ow.encrypt(验证码, "A-16-Byte-keyVal"), "=", "%3D"), "/", "%2F") '正确
        data = "tmp8wsdF4=" & uuid & "&ENTITYCLASSNAME=&username=" & username & "&password=" & password & "&aesEncrypt=true&role=2ee071e8751ab4c618ba930443d00a9b&os=Windows&browser=Chrome&resolution=1920*1080&validate=" & validate
'       Debug.Print Data
        url = "http://10.190.168.62/nnjzj/nnjzj/login/checkLogin.html"
        .Open "POST", url, False
        .setRequestHeader "Referer", "http://10.190.168.62/nnjzj/nnjzj/newhome/index.html"
        .setRequestHeader "Origin", "http://10.190.168.62"
        .setRequestHeader "Cookie", Cookie
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=UTF-8"
        .send data
        res = .responseText
        ow.execScript "var res =" & res & ";"
        Status = ow.eval("res[""status""]")
'        Debug.Print .getAllResponseHeaders()

        If InStr(.responseText, "Sorry..运行异常") Then
            ei = ei + 1
            MsgBox "Sorry..运行异常"
            Debug.Print "Sorry..运行异常"
            If ei > 5 Then End
            Set http = Nothing
            GoTo cx
        ElseIf Status = "FAIL" Then
            msg = ow.eval("res[""msg""]")
            MsgBox msg
            End
        End If
        
        ck = Split(.getAllResponseHeaders(), "Set-Cookie: ")
        ck1 = Split(ck(4), ";")(0)
        ck2 = Split(ck(6), ";")(0)
        Cookie = ck1 & ";" & ck2
        Debug.Print Cookie
        
        '获取secreKkey
        Dim secreKkey As String
        url = "http://10.190.168.62/nnjzj/manager/index.html"
        .Open "GET", url, False
        .setRequestHeader "Cookie", Cookie
        .send
        
        secreKkey = Split(Split(.responseText, "secreKkey=""")(1), """></script>")(0)
'        Debug.Print .responseText
    End With
    Set http = Nothing
    With Sheet20
        .Cells(2, "F").value = secreKkey
        .Cells(3, "F").value = Cookie
    End With
    参数 = Array(Cookie, secreKkey)
    
End Function

Sub js读取()

    ' 创建 HTML 文档对象和窗口对象
    Dim oDom As Object, ow As Object, result As String
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    
'    FilePath = "E:\桌面\Python\des加密.js"'文件读取
'    Open FilePath For Binary As #1
'    json = Input$(LOF(1), #1)
'    Close #1 ' 关闭文件
    Dim pathBase64 As String, pathDes As String, pathAes As String
'    pathBase64 = "http://10.190.168.62:80/nnjzj/resources/script/common/jquery/jquery.base64.js"
    pathBase64 = "http://10.190.128.231/static/js/base64/base64.min.js"
'    pathDes = "http://10.190.168.62:80/nnjzj/resources/script/common/des.js" '网络读取
    pathAes = "http://10.190.168.62:80/nnjzj/resources/script/common/aes.js" '网络读取
    AES = 加载js(pathAes)
    Base64 = 加载js(pathBase64)
'    DES = 加载js(pathDes)

    ' 定义 JavaScript 函数
'    ow.execScript DES
'    Debug.Print ow.strEnc("fdadf21s1a5", "1df5a3a1e3")
    ow.execScript Base64
'    Debug.Print ow.strEnc("fdadf21s1a5", "1df5a3a1e3")
    ow.execScript AES
    Debug.Print ow.encrypt("zqzcwd", "RailsYdjszz46495")
End Sub

Function 加载js(urlPath As String)
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "get", urlPath, False
        .send
'        Debug.Print .responseText
        加载js = .responseText
    End With
End Function

Sub 题库获取(试卷编号$) '传coolkie进来,答案要分列
    Dim js As Object, Json As Object, cls As New ClsJson, trr(), jrr, dt As New Dictionary, tm, sj As String
    Dim 题型 As String, 题干 As String, 答案 As String, 选项 As String, 试卷 As Dictionary
    Dim oDom As Object, ow As Object
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow

    t = Timer
    
    '遍历考生试卷
'    With Sheet20
'        For n = 2 To .Cells(Rows.count, "Q").End(3).Row
'            试卷 = 提取试卷(试卷编号, .Cells(n, "Q")) & Chr(13) & 试卷
'        Next
'    End With
    
    '异步
    If Sheet20.Cells(Rows.count, "H").End(3).Row > 1 Then
        Set 试卷 = 提取试卷_异步(试卷编号)
    Else
        MsgBox "没有试卷！"
        End
    End If
    
    Application.ScreenUpdating = False
    
    If 试卷.count > 0 Then
        ReDim trr(1 To 试卷.count * 50 + 50, 5)
        For Each rest In 试卷.keys
        
        topic = "function topic() {" & Trim(rest) & "return topicArray;};"
        ow.execScript topic
        
        Set jss = ow.eval("topic()")
        
            For Each js In jss
                题型 = js.basetype
                题干 = Replace(js.topic, Chr(10), "")
        '        If InStr(题干, "子系统包括钩计划编制和钩计划管理") Then Stop
                答案 = js.topicAnswer2
                选项 = Replace(js.topicOption, Chr(10), "")
                If dt.Exists(题型 & 题干) Then
                    For Each ss In Split(选项, "$$") '这里会过滤掉判断题
                        If InStr(dt(题型 & 题干), Trim(ss)) = 0 Then
                            k = k + 1
                            trr(k, 0) = "=ROW()-1"
                            trr(k, 1) = 题型
                            trr(k, 2) = 题干
                            trr(k, 3) = 答案
                            trr(k, 4) = 选项
                            dt(题型 & 题干) = dt(题型 & 题干) & "$$" & 选项  '加入字典去重
                        End If
                    Next
                Else
                    k = k + 1
                    trr(k, 0) = "=ROW()-1"
                    trr(k, 1) = 题型
                    trr(k, 2) = 题干
                    trr(k, 3) = 答案
            
                    If trr(k, 1) = "判断题" Then
                        trr(k, 4) = "对$$错"
                    Else
                        trr(k, 4) = 选项
                    End If
                    dt(题型 & 题干) = 选项  '加入字典去重
                End If
            Next
        Next
    End If
    
    
    '补一份试卷
'    Dim result As String, pathDes As String
'    pathDes = "http://10.190.168.62:80/nnjzj/resources/script/common/des.js" '网络读取
'    DES = 加载js(pathDes)
'    ow.execScript DES
'    With Sheet20 '登录信息
'        secreKkey = .Cells(2, "F").value
'        cookie = .Cells(3, "F").value
'    End With
'    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
'    With http
'        uuid = getuuid()
'        url = "http://10.190.168.62/nnjzj/exam/examreport/topiclist.html" '试卷试题地址
'        data = "%24%40%24.ecryptedData=" & ow.strEnc("{""tmp8wsdF4"":""" & uuid & """,""ENTITYCLASSNAME"":"""",""ttop003"":""" & 试卷编号 & """}", secreKkey)
'        .Open "POST", url, False
'        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=UTF-8"
'        .setRequestHeader "Origin", "http://10.190.168.62"
'        .setRequestHeader "Referer", "http://10.190.168.62/nnjzj/manager/index.html"
'        .setRequestHeader "Cookie", cookie
'        .send data
'        'Debug.Print .responseText
'
'        '题目加入数组中
''        Set Json = JsonConverter.ParseJson(.responseText)
'        ow.execScript "var obj=" & .responseText & ";"
'        Set jss = ow.eval("obj.testtxList")
'
'        For Each da In jss
'            题型 = da.ttx004
'            题干 = Replace(da.testtopic.ttop012, Chr(10), "")
''            If InStr(题干, "子系统包括钩计划编制和钩计划管理") Then Stop
'            答案 = da.testtopic.ttop022
'
'            On Error Resume Next
'            选项 = Replace(da.testtopic.ttop019, Chr(10), "")
'            If dt.Exists(题型 & 题干) Then
'                For Each ss In Split(选项, "$$")
'                    If InStr(dt(题型 & 题干), Trim(ss)) = 0 Then
'                        k = k + 1
'                        trr(k, 0) = "=ROW()-1"
'                        trr(k, 1) = 题型 '题型js(3).SubMatches(1)
'                        trr(k, 2) = 题干 '题干js(4).SubMatches(1)
'                        trr(k, 3) = 答案 '答案js(9).SubMatches(1)
'                        trr(k, 4) = 选项 '选项js(7).SubMatches(1)
'                        dt(题型 & 题干) = dt(题型 & 题干) & "$$" & 选项  '加入字典去重
'                    End If
'                Next
'            Else
'                k = k + 1
'                trr(k, 0) = "=ROW()-1"
'                trr(k, 1) = 题型 '题型js(3).SubMatches(1)
'                trr(k, 2) = 题干 '题干js(4).SubMatches(1)
'                trr(k, 3) = 答案 '答案js(9).SubMatches(1)
'
'                If trr(k, 1) = "判断题" Then
'                    trr(k, 4) = "对$$错"
'                Else
'                    trr(k, 4) = 选项 '选项js(7).SubMatches(1)
'                End If
'                dt(题型 & 题干) = 选项  '加入字典去重
'            End If
'        Next
'
'        On Error GoTo 0
'    End With

    With Sheet21
        .Range("A2").CurrentRegion.Offset(1).Clear '清空
        .Range("A2").Resize(UBound(trr, 1), UBound(trr, 2)) = trr '写入
        With .Cells
'           .Replace """", "", 2, , 0, 0
'           .WrapText = True '自动换行
            .HorizontalAlignment = xlCenter '居中
            .VerticalAlignment = xlCenter '居中
            .CurrentRegion.Borders.LineStyle = xlContinuous '设置样式边框
            Application.Intersect(.CurrentRegion.Offset(1), .Columns(3)).HorizontalAlignment = 1 '交集区域左对齐
        End With
        
        '分列
        Application.DisplayAlerts = False
        .Columns("E:E").TextToColumns DataType:=xlDelimited, TextQualifier:=xlNone, ConsecutiveDelimiter:=True, Other:=True, OtherChar:="$$", TrailingMinusNumbers:=True
        .Range("A1:L1") = Array("序号", "试题类型", "试题内容", "正确选项", "A", "B", "C", "D", "E", "F", "G", "H") '标题
        Application.DisplayAlerts = True
    End With
    
    '排序
    升序 Sheet21, Sheet21.Range("A1:L1"), Sheet21.Range("C1"), 1 '题干排序
    升序 Sheet21, Sheet21.Range("A1:L1"), Sheet21.Range("B1"), 1 '题型排序
    Application.ScreenUpdating = True
    MsgBox "完成！" & Chr(10) & Sheet20.CommandButton4.Caption & Chr(10) & "用时：" & Format(Timer - t, "0.00秒")
End Sub


Function 提取试卷(试卷编号$, id)
    ck = Sheet20.Range("F3").value
cx:
'    url = "http://10.190.168.62/nnjzj/user/exam/topic_preview.html?tpid=" & 试卷编号 & "&no=" & n '& "&rnd=2" '无答案
    url = "http://10.190.168.62/nnjzj/user/exam/topic_answer.html?tpid=" & 试卷编号 & "&egid=" & id '有答案
'    url = "http://10.190.168.62/nnjzj/user/exam/preview.html?tpid=" & 试卷编号 & "&no=" & n
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    With http
        '登录取Cookie
        .Open "GET", url, False
'        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
'        .setRequestHeader "Referer", "http://10.190.168.62/nnjzj/user/exam/preview.html?tpid=01HTYG413H8W6EAXPSEEQXZKPN&no=3"
        .setRequestHeader "Cookie", ck
        .send
        If InStr(.responseText, "查询失败") Then
            ck = 参数()(0)
            Set http = Nothing
            GoTo cx
        End If
        '保存.桌面 .responseText
        提取试卷 = .responseText
    End With
End Function

Function 提取试卷_异步(试卷编号$)
    Dim http() As Object, id$, ck$, k%, j%, rs%, ds As New Dictionary, dc As New Dictionary
    Dim stream As Object, path As String, fso As Object, ss As String
'    Dim fso As New FileSystemObject
    Set fso = CreateObject("scripting.FileSystemObject")

    path = "C:\Users\Administrator\Desktop\每日站存\试卷缓存\试卷\"
    If Dir(path, vbDirectory) <> "" Then
        fn = Dir(path & "\.")
        Do While fn <> ""
            Kill path & "\" & fn
            fn = Dir
        Loop
    Else
        '创建文件夹
        MkDir path
    End If
    
    With Sheet20
        
        ck = .Range("F3").value
'        试卷编号 = "01JDVKJVFV9JW6ACESM0N5TMBB"
        rs = .Cells(Rows.count, "Q").End(3).Row
        
        ReDim http(2 To rs)
        For n = 2 To rs
cx:
            姓名 = .Cells(n, "J").value
            id = .Cells(n, "Q").value
            url = "http://10.190.168.62/nnjzj/user/exam/topic_answer.html?tpid=" & 试卷编号 & "&egid=" & id '有答案
            Set http(n) = CreateObject("WinHttp.WinHttpRequest.5.1") 'CreateObject("MSXML2.XMLHTTP.6.0")
            With http(n)
                .Open "GET", url, True
                .setRequestHeader "Cookie", ck
                .send
                ds(http(n)) = 姓名 '对象加入字典
            End With
        Next
        t = Timer
        Do
            For Each dd In ds.keys
                On Error Resume Next
                ss = dd.responseText
                If InStr(ss, "查询失败") Then
                    Debug.Print "未登录！！！"
                    Application.DisplayAlerts = True
                    Application.ScreenUpdating = True
                    提示 = "提取到：" & j & "/" & rs - 1 & "份试卷。"
                    Sheet20.CommandButton4.Caption = 提示
                    End
                End If
                On Error GoTo 0
               If ss <> "" Then
                    If dc.Exists(ss) Then
'                        Debug.Print ds(dd) & "；" & "试卷已存在！"
                        k = k + 1
                        If k > 50 Then Exit Do
                    Else
                        j = j + 1
                        
                        '保存试卷
                        '试卷 = ss & Chr(13) & 试卷
                        With CreateObject("ADODB.Stream")
                            .Type = 2
                            .Charset = "UTF-8"
                            .Open
                            .WriteText ss
                            .SaveToFile path & ds(dd) & ".txt", 2
                            .Close
                        End With
                        Debug.Print "下载=：" & ds(dd)
                        
                        '加入字典中去中，重复试卷不要
                        dc(ss) = ds(dd)
                    End If
                    
                    '移除试卷
'                    Debug.Print ds(dd) & "；" & "试卷移除！"
                    ss = ""
                    ds.Remove dd
                    
                End If
                DoEvents '交回控制权给系统，防卡死！
            Next
            If Timer - t > 50 Then Exit Do '取50秒的试卷
            DoEvents '交回控制权给系统，防卡死！
        Loop Until ds.count = 0
    End With
    
    '输出
    If dc Is Nothing Then
        提示 = "提取到：" & j & "/" & rs - 1 & "份试卷。"
        Sheet20.CommandButton4.Caption = 提示
        MsgBox "没有获取到试卷！！！"
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        End
    Else
        提示 = "提取到：" & j & "/" & rs - 1 & "份试卷。"
        Sheet20.CommandButton4.Caption = 提示
        Set 提取试卷_异步 = dc
    End If
    
    '清空对象
    Set fso = Nothing
End Function

Function 读取试卷并运行Json()

    Dim http As Object, ck$, id$, 试卷编号$, topic$
    ck = Sheet20.Range("F3").value
    试卷编号 = "01JX2585RMW0R9B3QG5XNZ69GQ"

    id = Sheet20.Cells(3, "Q").value
    url = "http://10.190.168.62/nnjzj/user/exam/topic_answer.html?tpid=" & 试卷编号 & "&egid=" & id '有答案
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    With http
        .Open "GET", url, False
        .setRequestHeader "Cookie", ck
        .send
        rest = .responseText
    End With

    topic = "function topic() {" & Trim(rest) & "return topicArray;};"
    
    Dim oDom As Object, ow As Object
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    ow.execScript topic
    Set jss = ow.eval("topic()")

End Function


Function Json(path As String)

'    filePath = "E:\桌面\Python\des加密.js" '文件读取
    Open path For Binary As #1
    Json = Input$(LOF(1), #1)
    Close #1 ' 关闭文件
    
'    filePath = "C:\Users\Administrator\Desktop\HTML.txt" '文件保存
'    Open filePath For Output As #1
'    Print #1, .responseText
'    Close #1
End Function

Function 升序(ws As Worksheet, rng As Range, rn As Range, n) '标题，要升序的单元格
    '排序-升
    With ws
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
'        Set rng = .Range("A1:L1")
        rng.AutoFilter
        .AutoFilter.Sort.SortFields.Clear
        .AutoFilter.Sort.SortFields.Add key:=rn, SortOn:=xlSortOnValues, Order:=n, DataOption:=xlSortNormal '1升2降
        With .AutoFilter.Sort
            .Header = xlYes '包唅标题
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
'        If ws.AutoFilterMode Then ws.AutoFilterMode = False
    End With
End Function

Function 试卷预览()

' 创建 HTML 文档对象和窗口对象
    Dim oDom As Object, ow As Object, result As String, pathDes As String
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    pathDes = "http://10.190.168.62:80/nnjzj/resources/script/common/des.js" '网络读取
    DES = 加载js(pathDes)
    ow.execScript DES
重试:
    With Sheet20
'        Text = "http://10.190.168.62/nnjzj/user/exam/view.html?tpid=01J2K34XBE75QBTENP610P309C&egid=01J2YYWHX4H9H3CJ681K6EGH0K"
'        secreKkey = .Cells(2, "F")
        Cookie = .Cells(3, "F")
        '表格初始化
'        .Range("H1").CurrentRegion.Offset(1).Clear
    End With
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    With http
'        uuid = getuuid()
        url = "http://10.190.168.62/nnjzj/user/exam/view.html?tpid=01J2K34XBE75QBTENP610P309C&egid=01J2YYWHX4H9H3CJ681K6EGH0K" '考试记录地址
'        data = "%24%40%24.ecryptedData=" & ow.strEnc("{""tmp8wsdF4"":""" & uuid & """,""ENTITYCLASSNAME"":""ExamScoreRecordEntity"",""pageSize"":0,""currentPage"":1,""customColumn"":""egra008,egra013,egra022,result,user004,user005,organname,egratime,joinmode"",""customColumnName"":""参考时间,初始成绩,成绩,考试结果,账号,姓名,所属单位,考试用时,参考方式"",""egra004"":""" & 试卷编号 & """,""userqufen"":1,""orderType"":1}", secreKkey) '原pageSize=50
        .Open "GET", url, False
'        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
'        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=UTF-8"
'        .setRequestHeader "Origin", "http://10.190.168.62"
'        .setRequestHeader "Referer", "http://10.190.168.62/nnjzj/manager/index.html"
        .setRequestHeader "Cookie", Cookie
        .send
'        Debug.Print .responseText

    '设置Cookie
    '    Cookie = Split(Sheet20.Range("F3").value, "=")
    '    url = Selection.Comment.Text
    '    InternetSetCookie url, "login_loginid", ""
    '    InternetSetCookie url, "login_username", ""
    '    InternetSetCookie url, "login_rem_username", "false"
    '    InternetSetCookie url, Cookie(0), Cookie(1)
    '    InternetSetCookie url, "learning.session.id", "29f94358-21d8-4001-8655-0f5cc19a8bc3"

        Dim ie As Object
        Set ie = CreateObject("InternetExplorer.application") '测试不可行
        With ie
            .Visible = True
            .navigate "about:blank"
            .document.Cookie = Cookie '设置Cookie
            .navigate url
        End With



        With UserForm6
            .Show 0
            With .WebBrowser1
                .Navigate2 "about:blank"
                Do While .Busy
                    DoEvents
                Loop
                .document.Body.innerHTML = http.responseText
                
                Do While .Busy Or .readyState <> 4
                    DoEvents
                Loop
            End With
        End With
    
        If InStr(.responseText, "Sorry..运行异常") Then
            MsgBox "Sorry..运行异常"
            Debug.Print "Sorry..运行异常"
            Exit Function
        ElseIf InStr(.responseText, "Sorry..登录信息失效") Then
'            MsgBox "Sorry..登录信息失效！"
            Debug.Print "Sorry..登录信息失效！,正在重试！"
            参数
            GoTo 重试
        End If
    End With
End Function

Sub 导入题库_字典()
    Dim arr(), brr(), 题型, 题目, 选项, 答案
    Dim rng As Range, rn As Range
    Dim A As New Dictionary, b As New Dictionary
    
    '待入库的题目_a
    arr = Sheet21.Range("A1").CurrentRegion.value
    
    '题库中的题目_b
    Set rng = Sheet28.Range("A5").CurrentRegion
    brr = rng.value
    
    '题库_b
    For i = 2 To UBound(brr, 1)
        题型 = Trim(brr(i, 2))
        题目 = Trim(brr(i, 3))
        答案 = Trim(brr(i, 4))
        For k = 5 To 12
            If 选项 = "" Then
                选项 = Trim(brr(i, k))
            ElseIf Trim(brr(i, k)) <> "" Then
                选项 = 选项 & "$$" & Trim(brr(i, k))
            End If
        Next
        b(题型 & "$$" & 题目) = 答案 & "$$" & 选项
        
        题型 = ""
        题目 = ""
        答案 = ""
        选项 = ""
    Next
    
    '待入库的题目_a
    For i = 2 To UBound(arr, 1)
        题型 = Trim(arr(i, 2))
        题目 = Trim(arr(i, 3))
        答案 = Trim(arr(i, 4))
        For k = 5 To 12
            If 选项 = "" Then
                选项 = Trim(arr(i, k))
            ElseIf Trim(arr(i, k)) <> "" Then
                选项 = 选项 & "$$" & Trim(arr(i, k))
            End If
        Next

        If b.Exists(题型 & "$$" & 题目) Then

            Debug.Print 题型 & "$$" & 题目
            For k = 5 To 12
                If InStr(b(题型 & "$$" & 题目), Trim(arr(i, k))) = 0 Then
                    A(题型 & "$$" & 题目) = 答案 & "$$" & 选项
                    Debug.Print "加入题库"
                    Exit For
                Else
                    Debug.Print "题库中已有该题目！"
                End If
            Next
        Else
            A(题型 & "$$" & 题目) = 答案 & "$$" & 选项
'            Debug.Print "加入题库"
        End If
        
        题型 = ""
        题目 = ""
        答案 = ""
        选项 = ""
    Next
    
    Application.EnableEvents = False
    With Sheet28
        Set rn = .Cells(rng.Row + rng.Rows.count, 2).Resize(A.count)
        With rn
            .value = Application.Transpose(A.keys) 'WorksheetFunction
            .TextToColumns DataType:=xlDelimited, TextQualifier:=xlNone, ConsecutiveDelimiter:=True, Other:=True, OtherChar:="$$", TrailingMinusNumbers:=True '分裂
        End With
        
        Set rn = .Cells(rng.Row + rng.Rows.count, 4).Resize(A.count)
        For Each Item In A.items
            j = j + 1
            rn.Cells(j) = Item
        Next
        rn.TextToColumns DataType:=xlDelimited, TextQualifier:=xlNone, ConsecutiveDelimiter:=True, Other:=True, OtherChar:="$$", TrailingMinusNumbers:=True '分裂
        
        With rn.Offset(0, -3)
            .Formula = "=ROW()-5" '序号
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        rn.CurrentRegion.Borders.LineStyle = xlContinuous '边框
        
        With rn.Resize(A.count, 9) '居中
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End With
    Application.EnableEvents = True
End Sub

Sub 导入题库_数组()
    Dim arr(), brr(), crr(), 题型, 题目, 选项, 答案
    Dim rng As Range, rn As Range, k%, j%
    Dim A As New Dictionary, b As New Dictionary
    
    Application.ScreenUpdating = False
    
    '待入库的题目_a
    arr = Sheet21.Range("A1").CurrentRegion.value
    ReDim crr(1 To UBound(arr, 1), 1 To UBound(arr, 2))
    
    '题库中的题目_b
    With Sheet28
        If .AutoFilterMode Then
            .AutoFilterMode = False
        End If
        Set rng = .Range("A8").CurrentRegion
        brr = rng.value
    End With
    
    '题库_b
    For i = 2 To UBound(brr, 1)
        题型 = Trim(brr(i, 2))
        题目 = Trim(brr(i, 3))
        答案 = Trim(brr(i, 4))
        For k = 5 To 12
            If 选项 = "" Then
                选项 = Trim(brr(i, k))
            ElseIf Trim(brr(i, k)) <> "" Then
                选项 = 选项 & "$$" & Trim(brr(i, k))
            End If
        Next
        If b.Exists(题型 & "$$" & 题目) Then
            b(题型 & "$$" & 题目) = b(题型 & "$$" & 题目) & 答案 & "$$" & 选项
        Else
            b(题型 & "$$" & 题目) = 答案 & "$$" & 选项
        End If
        题型 = ""
        题目 = ""
        答案 = ""
        选项 = ""
    Next
    
    '待入库的题目_a
    For i = 2 To UBound(arr, 1)
        题型 = Trim(arr(i, 2))
        题目 = Trim(arr(i, 3))
        答案 = Trim(arr(i, 4))
        For k = 5 To 12
            If 选项 = "" Then
                选项 = Trim(arr(i, k))
            ElseIf Trim(arr(i, k)) <> "" Then
                选项 = 选项 & "$$" & Trim(arr(i, k))
            End If
        Next
        
'        Debug.Print 题型 & "$$" & 题目
        If b.Exists(题型 & "$$" & 题目) Then
            For k = 5 To 12
                If InStr(b(题型 & "$$" & 题目), Trim(arr(i, k))) = 0 Then
                    j = j + 1
                    crr(j, 1) = "=ROW()-5"
                    crr(j, 2) = 题型
                    crr(j, 3) = 题目
                    crr(j, 4) = 答案
                    crr(j, 5) = 选项
'                    'Debug.Print "加入题库"
                    b(题型 & "$$" & 题目) = b(题型 & "$$" & 题目) & 答案 & "$$" & 选项 '20250202更新
                    Exit For
                Else
                    'Debug.Print "题库中已有该题目！"
                End If
                If Trim(arr(i, k)) = "" Then Exit For
            Next
        Else
            j = j + 1
            crr(j, 1) = "=ROW()-8"
            crr(j, 2) = 题型
            crr(j, 3) = 题目
            crr(j, 4) = 答案
            crr(j, 5) = 选项
'            'Debug.Print "加入题库"
            b(题型 & "$$" & 题目) = 答案 & "$$" & 选项
        End If
        
        题型 = ""
        题目 = ""
        答案 = ""
        选项 = ""
    Next
    
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    With Sheet28
        If j > 0 Then '当有新题目时导入
            Set rn = .Cells(rng.Row + rng.Rows.count, 1).Resize(UBound(crr, 1), 5)
            rn.value = crr
    
            rn.Columns(5).TextToColumns DataType:=xlDelimited, TextQualifier:=xlNone, ConsecutiveDelimiter:=True, Other:=True, OtherChar:="$$", TrailingMinusNumbers:=True '分裂
    
            With rn.Columns(1)
'                .Formula = "=ROW()-5" '序号
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            rn.CurrentRegion.Borders.LineStyle = xlContinuous '边框
            
            With .Columns("A:B") '居中
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            With Sheet28.Columns("D:L") '居中
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            .Cells.Replace "#N/A", ""
        End If
    End With
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "入库：" & j & "条题目！"
End Sub

Function 选中行显示(sht As Worksheet, rngs As Range, Target As Range)
Dim rng As Range
With sht
    If Target.Rows.count = 1 Then
        Set rng = rngs.CurrentRegion
        If Not Intersect(Target, rng) Is Nothing And Target.count = 1 Then
            .Cells.WrapText = False
            With rng.Offset(1)
                .Font.Size = 12
                .Interior.ColorIndex = 0
                With .Rows(Target.Row - rng.Row)
                    .Font.Size = 15
                    .Interior.ColorIndex = 8 '背景色-青
                    .WrapText = True
                    
                    '答案标记
                    Dim arr(), d As New Dictionary, rn As Range
                    arr = rng.Offset(0, 4).Resize(1, 8)
                    For Each ar In arr
                        k = k + 1
                        d(ar) = k
                    Next
                    
                    '答案
                    Set rn = .Cells(4)
                    rn.Interior.ColorIndex = 19 '米黄
                    
                    '选项
                    For i = 1 To Len(rn.value)
                        st = Mid(rn.value, i, 1)
                        rn.Offset(0, d(st)).Interior.ColorIndex = 6 '黄
                    Next
                End With
            End With
        Else
            With rng.Offset(1)
                .Cells.WrapText = False
                .Font.Size = 12
                .Interior.ColorIndex = 2 '背景色-白
            End With
        End If
'        With rng.Rows(1)'标题
'            .Font.Size = 14
'            .Font.Bold = True
'            .Font.ColorIndex = 2
'            .Interior.ColorIndex = 14
'            .RowHeight = 30
'        End With
    End If
    End With
    
    Set d = Nothing
End Function

Sub 查看请求休()
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    pathDes = "http://10.190.168.62:80/nnjzj/resources/script/common/des.js" '网络读取
    DES = 加载js(pathDes)
    ow.execScript DES
    '读取txt文件内容
    txt = 1
    明文 = ow.strDec(txt)
    
End Sub

'中国铁路南宁局集团有限公司职工培训部
Sub 登录()
    Dim http As Object, oDom As Object, ow As Object
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    url = "http://jypx-xc.nann.cr/jypx-api/v1/front/login/verificationCode"
    With http
        .Open "GET", url, False
        .setRequestHeader "Accept", "application/json;charset=UTF-8"
        .setRequestHeader "Accept-Language", "zh-CN,zh;q=0.9"
        .setRequestHeader "Referer", "http://jypx-xc.nann.cr/jypx/home"
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
        .send
        res = .responseText
        
        ow.execScript "var res=" & res & ";"
        suc = ow.eval("res.success")
        If suc Then
            captchaBase64 = ow.eval("res.result.captchaBase64")
            captchaKey = ow.eval("res.result.captchaKey")
            timestamp = ow.eval("res.timestamp")

            With UserForm3
                .WebBrowser1.navigate "about:blank"
                .WebBrowser1.document.Write "<img src='" & captchaBase64 & "' style='position: relative; top: -20px;'>"
                .Show 0

                验证码 = InputBox("请输入图片中的验证码！")

            End With
            Unload UserForm3
            
            url = "http://jypx-xc.nann.cr/jypx-api/v1/front/login/login"
            data = "{""username"":""IspHAZXD3EoW2LOMCyhzDg=="",""password"":""VCjJwBPi8txjOkUsJ21Lqg=="",""checkKey"":""" & captchaKey & """,""captcha"":""" & 验证码 & """,""roleId"":""5346ec69c5dd45bea6bab9429eafb505""}"
            .Open "POST", url, False
            .setRequestHeader "Accept", "application/json;charset=UTF-8"
            .setRequestHeader "Content-Type", "application/json;charset=UTF-8"
            .setRequestHeader "Origin", "http://jypx-xc.nann.cr"
            .setRequestHeader "Referer", "http://jypx-xc.nann.cr/jypx/home"
            .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
            .setRequestHeader "roleId", "5346ec69c5dd45bea6bab9429eafb505"
            .send data
                
            res = .responseText
            
            ow.execScript "var res=" & res & ";"
            suc = ow.eval("res.success")
            If suc Then
                token = ow.eval("res.result.token")
                SaveSetting "Momda", "职工教育平台", "token", token
                Debug.Print token
               
            End If
        End If
    End With

End Sub

Sub 职工教育平台_试卷()

    Dim http As Object, oDom As Object, ow As Object, token As String
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    token = GetSetting("Momda", "职工教育平台", "token")
    url = "http://jypx-xc.nann.cr/jypx-api/v1/front/login/verificationCode"
    With http
        'url = "http://jypx-xc.nann.cr/jypx-api/v1/ddy/examMonitor/gradePage?pageSize=10&pageNo=1&entityClassName=ExamScoreEntity&activityId=ff8080819bbc38fc019bc0de1d2202af&_t=1769541901521"
        url = "http://jypx-xc.nann.cr/jypx-api/v1/ddy/examMonitor/activityPage?userqufen=1&urre010=2025-12-01+00:00:00,2026-01-31+23:59:59&pageSize=50&pageNo=1&starTime=2025-12-01+00:00:00&endTime=2026-01-31+23:59:59&_t=" & timestamp
         .Open "GET", url, False
         .setRequestHeader "Accept", "application/json;charset=UTF-8"
         .setRequestHeader "Referer", "http://jypx-xc.nann.cr/jypxAdmin/"
         .setRequestHeader "X-Access-Token", token
         .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
         .send data
         res = .responseText
         ow.execScript "var res=" & res & ";"
         suc = ow.eval("res.success")
        If suc Then
             Set recs = ow.eval("res.result.records")
             Sheet20.Range("A1:D1").CurrentRegion.Offset(1).ClearContents
             Sheet20.Range("A1:D1") = Array("考试活动名称", "发布人", "ID", "发布时间")
             
             For Each rec In recs
                 Debug.Print rec.tact004
                 With Sheet20
                     r = .Cells(.Rows.count, "A").End(3).Row + 1
                     .Cells(r, 1) = rec.tact004
                     .Cells(r, 2) = rec.createBy
                     .Cells(r, 3) = rec.id
                     .Cells(r, 4) = rec.urre010
                 End With
             Next
             
             With Sheet20
         '        .Range("A1").CurrentRegion.Offset(1).VerticalAlignment = 4
                 .Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous '设置边框
         '        .Range("A1:I1").EntireColumn.AutoFit
                 升序 Sheet20, .Range("A1:D1"), .Range("D1"), 2 '表对象，标题行，排序列单元格，1升2降
             End With
        End If
    End With
End Sub

Function 职工教育平台_成绩()

    Dim http As Object, oDom As Object, ow As Object, token As String
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    token = GetSetting("Momda", "职工教育平台", "token")
    url = "http://jypx-xc.nann.cr/jypx-api/v1/front/login/verificationCode"
    With http
        url = "http://jypx-xc.nann.cr/jypx-api/v1/ddy/examMonitor/gradePage?pageSize=10000&pageNo=1&entityClassName=ExamScoreEntity&activityId=ff8080819bbc38fc019bc0de1d2202af"
         .Open "GET", url, False
         .setRequestHeader "Accept", "application/json;charset=UTF-8"
         .setRequestHeader "Referer", "http://jypx-xc.nann.cr/jypxAdmin/"
         .setRequestHeader "X-Access-Token", token
         .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
         .send data
         res = .responseText
         ow.execScript "var res=" & res & ";"
         Set rec = ow.eval("res.result.records")
        '标题
        Sheet20.Range("H1:P1").CurrentRegion.Offset(1).ClearContents
        Sheet20.Range("H1:P1") = Array("参考时间", "账号", "姓名", "所属单位", "初始成绩", "成绩", "考试用时", "考试结果", "参考方式", "Egid")
            '设置格式
        With Sheet20.Range(Sheet20.Cells(2, "H"), Sheet20.Cells(Sheet20.Rows.count, "P"))
            .Columns(1).NumberFormatLocal = "yyyy-m-d h:mm;@"
            .Columns(2).NumberFormatLocal = "@"
            .Columns(3).NumberFormatLocal = "G/通用格式"
            .Columns(4).NumberFormatLocal = "G/通用格式"
            .Columns(5).NumberFormatLocal = "0.0"
            .Columns(6).NumberFormatLocal = "0.0"
            .Columns(7).NumberFormatLocal = "G/通用格式"
            .Columns(8).NumberFormatLocal = "G/通用格式"
            .Columns(9).NumberFormatLocal = "G/通用格式"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        For Each da In rec
            With Sheet20
                r = .Cells(.Rows.count, "I").End(3).Row + 1
                试卷编号 = da.id
                试卷URL = "http://jypx-xc.nann.cr/jypx/examPaper?id=" & 试卷编号 '试卷URL
'                下载URL = "http://10.190.168.62/nnjzj/user/exam/topic_answer.html?tpid=" & 试卷编号 & "&egid=" & da("egra001") '试卷URL
                .Cells(r, "H") = da.egra008
                .Cells(r, "I") = da.idNumber
                .Cells(r, "J") = CallByName(da, "fullName", VbGet)
                .Cells(r, "K") = da.orgName
                .Cells(r, "L") = da.egra022
                .Cells(r, "M") = da.egra013
                .Cells(r, "N") = da.egra010
                .Cells(r, "O") = IIf(da.result = 1, "通过", "未通过")
                .Cells(r, "P") = IIf(da.joinMode, "APP", "PC")
                .Cells(r, "Q") = CallByName(da, "userId", VbGet) '用户ID
                
                '设置批注
'                With .Cells(r, "J")
'                    If Not .Comment Is Nothing Then .Comment.Delete
'
'                    下载试卷到批注中
'                    With CreateObject("WinHttp.WinHttpRequest.5.1")
'                        .Open "GET", 下载URL, False
'                        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
'                        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded;charset=UTF-8"
'                        .setRequestHeader "Origin", "http://10.190.168.62"
'                        .setRequestHeader "Referer", "http://10.190.168.62/nnjzj/manager/index.html"
'                        .setRequestHeader "Cookie", cookie
'                        .send
'                        res = .responseText
'                    End With
'                    .AddComment 试卷URL
'                End With
'                .Hyperlinks.Add .Cells(r, "Q"), 试卷URL, , , "查看试卷"'设置超链接
            End With
        Next
        
        '排序
        
        
    End With
    With Sheet20
'        .Range("A1").CurrentRegion.Offset(1).VerticalAlignment = 4
        .Range("H1").CurrentRegion.Borders.LineStyle = xlContinuous
        .Range("H1:P1").EntireColumn.AutoFit
        升序 Sheet20, .Range("H1:P1"), .Range("L1"), 2 '表对象，标题行，排序列单元格，1升2降
    End With
    
End Function
