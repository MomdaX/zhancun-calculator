Attribute VB_Name = "查询卸线"
Sub 查询卸车地点()
    Application.EnableEvents = False
    Dim http() As Object, https() As Object
    ReDim http(0 To 1)
    ReDim https(0 To 1)
    ReDim 变更(0 To 1)
    ReDim 运单号(0 To 1)
    ReDim 要求书号(0 To 1)
    Dim HTML As String, url As String
    Dim d As New Dictionary
    Dim divs As IHTMLElementCollection
    Dim ht As New HTMLDocument 'Dim ht As New HTMLDocument
    Dim Body As HTMLBody
    Dim paths As String
    
    paths = ThisWorkbook.path & "\票据缓存"
    Call 创建文件夹(paths)
    
    Dim rng As Range
    Set rng = Sheet13.Range("F2:F" & Sheet13.Cells(Sheet13.Rows.count, "F").End(3).Row)
    If WorksheetFunction.CountA(rng) = 1 And rng.address = Range("F1:F2").address Then
        MsgBox "无查询内容！"
        End
    End If
    Set http(0) = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = "http://10.3.64.102:8080/web/checkLogin.htm"
    data = "username=qvz01&password=3XXLAg5rywhaTB16UOIGjQ%3D%3D&loginFlag=1"
    With http(0)
        '登录取Cookie
        .Open "POST", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "Origin", "http://10.3.64.102:8080"
'        .setRequestHeader "Host", "10.190.136.155:8888"
        .setRequestHeader "Referer", "http://10.3.64.102:8080/web/login.htm"
        .setRequestHeader "Cookie", "critc_chk=1; critc_username=qvz01"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"

        .send data
        Cookie = Split(.getResponseHeader("Set-Cookie"), ";")(0) & "; critc_chk=1; critc_username=qvz01; current=module_5282"
    End With
    
    '取票据号
    Dim hts As String, M1 As Boolean, M2 As Boolean, M3 As Boolean, M4 As Boolean
    ReDim Preserve http(0 To rng.Rows.count)
    ReDim Preserve https(0 To rng.Rows.count)
    ReDim Preserve 变更(0 To rng.Rows.count)
    ReDim Preserve 运单号(0 To rng.Rows.count)
    ReDim Preserve 要求书号(0 To rng.Rows.count)
    For Each rn In rng
        i = i + 1
        M1 = rn.value <> "" '运单号码空
        M2 = rn.Offset(0, 9).value Like "*重*" Or rn.Offset(0, 9).value = ""
        M3 = rn.Offset(0, 8).value = "钦州港" '到站名
'       M4 = rn.Offset(0, 11).Value = "" '车号记事栏空
        
        If M1 And M2 Then
            If M3 Then
                Set http(i) = CreateObject("WinHttp.WinHttpRequest.5.1")
                Set https(i) = CreateObject("WinHttp.WinHttpRequest.5.1")
                    With http(i)
                        票据号 = rn.value
                        url = "http://10.1.203.120:8105/dzpjfw/dzpjzzqd.htm?appid=10001&pjdh=" & 票据号
                        .Open "get", url, True
                        .setRequestHeader "Cookie", Cookie
                        .setRequestHeader "Referer", "http://10.3.64.102:8080/"
                        .send '久
                        rn.Offset(0, 10).value = "发送请求..."
                    End With
            Else
               rn.Offset(0, 10).value = rn.Offset(0, 8).value
            End If
        End If
    Next
    
    
    Sheet13.Cells.ClearComments
    
    '请求数据
       For i = 1 To Cells(ActiveSheet.Rows.count, 6).End(3).Row - 1
        hts = ""
        If Not http(i) Is Nothing Then
        With http(i)
'        While .readyState <> 4
'        DoEvents
'        Wend
        Do
            On Error Resume Next
    '            st = .Status
                hts = .responseText
                ht.Body.innerHTML = .responseText
                DoEvents
'                Debug.Print "1"
            On Error GoTo 0
        Loop Until hts <> ""
    
        Set A = ht.getElementsByTagName("a")
        If A.length > 0 Then '集装箱

            On Error Resume Next
            运单号(i) = Split(Split(A(1).outerHTML, "ydhm=")(1), "'")(0)
            要求书号(i) = Split(Split(A(2).outerHTML, "arg1=")(1), "'")(0)
            On Error GoTo 0
            If A.length > 3 Then
            
                变更(i) = "运输变更"
            Else
            url = "http://10.1.203.120:8105/hyreport/reports/reportJsp/showReport.jsp?raq=hwyd1.raq&ydhm=" & 运单号(i)
            With https(i)
                .Open "get", url, True
                .setRequestHeader "Referer", "http://10.1.203.120:8105/dzpjfw/dzpjzzqd.htm?appid=10001&pjdh=" & 票据号
                .setRequestHeader "Cookie", Cookie
                .send '久/2
                Sheet13.Cells(i + 1, "p") = "查询运单..."
            End With
            End If
        Else '散货
            url = "http://10.1.203.120:8105/hyreport/reports/reportJsp/showReport.jsp?raq=hwyd1.raq&ydhm=" & 票据号
            With https(i)
                .Open "get", url, True
                .setRequestHeader "Referer", "http://10.1.203.120:8105/dzpjfw/dzpjyd.htm?appid=10001&pjdh=" & 票据号
                .setRequestHeader "Cookie", Cookie
                .send '久/2
                Sheet13.Cells(i + 1, "p") = "查询运单..."
            End With
        
        End If
        End With
        End If
    Next
    
       For i = 1 To Cells(ActiveSheet.Rows.count, 6).End(3).Row - 1
       hts = ""
'       If Not Sheet13.Cells(i + 1, "P").Comment Is Nothing Then Sheet13.Cells(i + 1, "P").Comment.Delete '删除批注
'       If Not Sheet13.Cells(i + 1, "F").Comment Is Nothing Then Sheet13.Cells(i + 1, "F").Comment.Delete '删除批注
       If Not https(i) Is Nothing Then
        Sheet13.Cells(i + 1, "p") = "正在查询专用线..."
            With https(i)
                Do
                On Error Resume Next
        '            st = .Status
                    hts = .responseText
'                    Debug.Print hts
                    ht.Body.innerHTML = .responseText
                    DoEvents
'                    Debug.Print "2"
                On Error GoTo 0
                Loop Until hts <> ""
                
                '运单导出
                文件名 = paths & "\" & Sheet13.Cells(i + 1, "A").value & ".html"
                Open 文件名 For Output As #1
                Print #1, .responseText
                Close #1
                
                '写入批注
                Dim Ct As Comment
                Set Ct = Sheet13.Cells(i + 1, "P").AddComment
                    Ct.text text:="Momda提示：" & Chr(10) & "双击查看运单详情！"
                
                '提取信息到表格中
    '            ht.body.innerHTML = hts
                Set tds = ht.getElementById("report1_A26")
                Set trs = ht.getElementsByTagName("tr")
                Set tds = trs(28).getElementsByTagName("td")
        '            Debug.Print tds(5).innerText
        
                Dim reg As New RegExp, arghtml$, 变更原因$
                
                If 变更(i) <> "" Then
                    With CreateObject("WinHttp.WinHttpRequest.5.1")
                        argurl = "http://10.1.203.120:8105/hyreport/reports/reportJsp/showReport.jsp?raq=hwysbg.raq&arg1=" & 要求书号(i)
                        .Open "get", argurl, True
                        .setRequestHeader "Referer", "http://10.1.203.120:8105/dzpjfw/dzpjzzqd.htm?appid=10001&pjdh=" & 票据号
                        .setRequestHeader "Cookie", Cookie
                        .send '久/2
                        Do
                            On Error Resume Next
                            arghtml = .responseText
                            DoEvents
                            On Error GoTo 0
                        Loop Until arghtml <> ""
'                        reg.Pattern = "<td rowspan=""4"" colspan=""88"" class=""report1_13"">(.*?)<\/td>"
'                        Set Mat = reg.Execute(arghtml)

                        ht.Body.innerHTML = arghtml
                        Set Table = ht.getElementById("report1")
                        Set trs = ht.getElementsByTagName("tr")
                        变更原因 = trs(34).getElementsByTagName("TD")(10).innerText

                        Sheet13.Cells(i + 1, "p") = 变更原因
                    End With
                    'Sheet13.Cells(i + 1, "p").value = 变更(i)
                Else
                    If tds(5).innerText = "" Then
                        Sheet13.Cells(i + 1, "p").value = IIf(tds(3).innerText = "钦州港（宁）", "钦州港-货场", tds(3).innerText)
                    Else
                        Sheet13.Cells(i + 1, "p").value = tds(5).innerText
                    End If
                End If
            End With
        End If
        Next
'        文件名 = "C:\Users\Administrator\Desktop\2.txt"
'        Open 文件名 For Output As #1
'        Print #1, .responseText
'        Close #1
        
    Sheet13.UsedRange.EntireColumn.AutoFit
    Application.EnableEvents = True
'    MsgBox "查询完成！"
End Sub

Sub 查询编组()
    Application.EnableEvents = False
'Dim xmlhttp As New MSXML2.ServerXMLHTTP
    Dim xmlhttp As New WinHttpRequest
    Dim HTML As String, url As String
    Dim d As New Dictionary
    Dim divs As IHTMLElementCollection
    Dim ht As New HTMLDocument
    Dim Body As HTMLBody, Cookie As String
    
    车号 = InputBox("车号")
    '车号 = "1661919"
    If 车号 = "" Then
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        End
    End If
    url = "http://10.3.64.102:8080/web/checkLogin.htm"
    data = "username=qvz01&password=3XXLAg5rywhaTB16UOIGjQ%3D%3D&loginFlag=1"
    With xmlhttp
        '登录
        .Open "POST", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "Origin", "http://10.3.64.102:8080"
'        .setRequestHeader "Host", "10.190.136.155:8888"
        .setRequestHeader "Referer", "http://10.3.64.102:8080/web/login.htm"
        .setRequestHeader "Cookie", "critc_chk=1; critc_username=qvz01"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"

        .send data
        Cookie = Split(.getResponseHeader("Set-Cookie"), ";")(0) & "; critc_chk=1; critc_username=qvz01; current=module_5282"
        'Debug.Print cookie

        '查询到卸
        url = "http://10.3.64.102:8080/web/qlc/fx/index.htm?ch=" & 车号
        .Open "get", url, False
        .setRequestHeader "Cookie", Cookie
        .setRequestHeader "Referer", "http://10.3.64.102:8080/web/qlc/fx/index.htm?ch=" & 车号
        .send '请求时间久2s
        HTML = .responseText
'        Debug.Print html
        
        href = Split(HTML, "<a target=""_blank""")
        If UBound(href) < 1 Then
            MsgBox "请核对车号！" & Chr(10) & 车号
            Application.ScreenUpdating = True
            Application.DisplayAlerts = True
            Application.EnableEvents = True
            End
        Else
            Debug.Print "正在查询卸车地点！"
            Sheet13.Cells.Clear '清空内容
        End If
        
        url = Split(Split(href(1), "href=""")(1), """>查看编组</a>")(0)
        .Open "get", url, False
        .send
        ht.Body.innerHTML = .responseText
        
        '写表
        Dim k, j
        With Sheet13
            Set Table = ht.getElementById("showTable") '表
            Set ths = ht.getElementsByTagName("th") '标题
            For Each TH In ths
                k = k + 1
                .Cells(1, k).value = TH.innerText
            Next
            Set trs = ht.getElementsByTagName("tbody")(0).getElementsByTagName("tr") '内容
            For Each tr In trs
                Set tds = tr.getElementsByTagName("td")
                r = .Cells(Rows.count, 1).End(3).Row + 1
                For Each td In tds
                    j = j + 1
                    .Cells(r, j).value = td.innerText
                Next
                j = 0
            Next
            
            查询卸线.查询卸车地点 '开始查询卸车地点
            
            '设置格式
            With .Range("A1").CurrentRegion '查询时的车号定位
                k = .Columns(1).Find(车号).Row
                .Rows(k).Interior.ColorIndex = 17 '蓝色
                
                With .Rows(1)
                    .Font.Bold = True '加粗
                End With
                .HorizontalAlignment = xlCenter '居中
            End With
            
                
                
        End With
        
        '读表
'        With ActiveSheet.QueryTables.Add(Connection:="URL;" & url, Destination:=Sheet13.Range("$A$1"))
'
'            .WebFormatting = xlWebFormattingNone
'            .Refresh 'BackgroundQuery:=False
'        End With
          
'        文件名 = "C:\Users\Administrator\Desktop\2.txt"
'        Open 文件名 For Output As #1
'        Print #1, .responseText
'        Close #1
        
    End With

    Application.EnableEvents = True
End Sub
Sub 结束查询()
    Application.EnableEvents = True
    End
End Sub
Sub 创建文件夹(path As String)
    
    If ActiveSheet.name = "到卸查询" Then
        If Dir(path, vbDirectory) <> "" Then
            Application.DisplayAlerts = False
    '        RmDir path
            fn = Dir(path & "\.")
            Do While fn <> ""
                Kill path & "\" & fn
                fn = Dir
            Loop
            Application.DisplayAlerts = True
        Else
            '创建文件夹
            MkDir path
        End If
    End If
End Sub
Sub cz()
Call 创建文件夹("C:\Users\Administrator\Desktop\每日站存\试卷缓存\试卷")
End Sub

Sub 查询货装历史(rng As Range)

    Dim url$, HTML$, 车号$, rn As Range, n%, k%, h%
    Dim http() As Object, ds As New Dictionary, hts As New HTMLDocument
    '参数
'    cookie = "SESSION=460b616b-5dce-40d0-995b-250b44260484; critc_chk=1; critc_username=qvz01; current=module_5282"
'    Set rng = Sheet13.Range("A2:A10")

'    With Sheet13
'        Set rng = Intersect(.Range("A1").CurrentRegion, .Range("A1").CurrentRegion.Offset(1).Columns(1))
'    End With
    
    ReDim http(0 To rng.Rows.count)
    '登录
    url = "http://10.3.64.102:8080/web/checkLogin.htm"
    data = "username=qvz01&password=3XXLAg5rywhaTB16UOIGjQ%3D%3D&loginFlag=1"
    Set http(0) = CreateObject("WinHttp.WinHttpRequest.5.1")
    With http(0)
        .Open "POST", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "Origin", "http://10.3.64.102:8080"
        .setRequestHeader "Referer", "http://10.3.64.102:8080/web/login.htm"
        .setRequestHeader "Cookie", "critc_chk=1; critc_username=qvz01"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"

        .send data
        Cookie = Split(.getResponseHeader("Set-Cookie"), ";")(0) & "; critc_chk=1; critc_username=qvz01; current=module_5282"
        'Debug.Print cookie
    End With
    
    
    For Each rn In rng '使用异步加速
        If rn.value <> "" Then
            n = n + 1
            车号 = rn.value
            Set http(n) = CreateObject("WinHttp.WinHttpRequest.5.1")
            With http(n)
                url = "http://10.3.64.102:8080/web/qlc/fx/index.htm?ch=" & 车号
                .Open "get", url, True
                .setRequestHeader "Cookie", Cookie
                .setRequestHeader "Referer", "http://10.3.64.102:8080/web/qlc/fx/index.htm?ch=" & 车号
                .send '请求时间久2s
                Set ds(http(n)) = rn '对象加入字典
                Set rn = Sheet13.Cells(rn.Row, "O")
'                If Not rn.Comment Is Nothing Then rn.Comment.Delete
            End With
        End If
    Next
    
    Do
        For Each dd In ds.keys
            On Error Resume Next
            HTML = dd.responseText
            If InStr(HTML, "查询失败") Then
                'Sheet20.CommandButton4.Caption = 提示
                End
            End If
            On Error GoTo 0
            If HTML <> "" Then
                内容 = "车号：" & ds(dd).value
            
                '保存.桌面 html
                hts.Body.innerHTML = HTML
                Set div = hts.getElementById("portlet_tab10") ' 现车到报正文 表div
                Set tbody = div.getElementsByTagName("tbody") '内容
                Set trs = tbody(0).getElementsByTagName("tr") '行信息
                For Each tr In trs
                
                Dim 车次 As String * 8, 当前站名 As String * 10, 载重 As String * 5, 品名 As String * 10, 到发日期, 货票记事栏
                    
                    当前站名 = tr.getElementsByTagName("td")(3).innerText
                    车次 = tr.getElementsByTagName("td")(6).innerText
                    载重 = tr.getElementsByTagName("td")(9).innerText
                    品名 = tr.getElementsByTagName("td")(14).innerText
                    到发日期 = tr.getElementsByTagName("td")(7).innerText & "  "
                    货票记事栏 = tr.getElementsByTagName("td")(15).innerText
                    DoEvents
                    '拼接信息
                    内容 = 内容 & Chr(10) & 车次 & 当前站名 & 载重 & 品名 & 到发日期 & 货票记事栏
                    'Debug.Print 内容
                    
                    '退出条件
                    k = k + 1
                    If k > 30 Then '取前30条数据
                        h = k
                        k = 0
                        Exit For
                    End If
                Next
                
                '写入批注
                Set rn = Sheet13.Cells(ds(dd).Row, "O")
                With rn
                    .AddComment 内容
                    .Comment.Shape.width = 360
                    .Comment.Shape.height = h * 12
                End With
                
                '删除对应字典对象
                内容 = ""
                HTML = ""
                Set rn = Nothing
                ds.Remove dd
            End If
        Next
        DoEvents '交回控制权给系统，防卡死！
    Loop Until ds.count = 0
    
    MsgBox "查询完成！"
End Sub

