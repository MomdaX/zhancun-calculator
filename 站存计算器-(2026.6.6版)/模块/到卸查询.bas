Attribute VB_Name = "到卸查询"
Sub 查询卸车专用线(rng As Range, rngs As Range)
    Application.EnableEvents = False
    Dim http() As Object, https() As Object, 票据号() As String, 车号() As String
    ReDim http(0 To 1)
    ReDim https(0 To 1)
    Dim HTML As String, url As String
    Dim d As New Dictionary
    Dim divs As IHTMLElementCollection
    Dim ht As New HTMLDocument 'Dim ht As New HTMLDocument
    Dim Body As HTMLBody
    Dim paths As String
    
    paths = ThisWorkbook.path & "\票据缓存"
    创建文件夹 paths
    
'    Dim rng As Range
'    Set rng = Sheet13.Range("A2:A" & Sheet13.Cells(Sheet13.Rows.count, "A").End(3).Row)
    If WorksheetFunction.CountA(rng) = 1 And rng.address = Range("A1:A2").address Then
        MsgBox "无查询内容！"
        End
    End If
    
    '登录取Cookie
    Set http(0) = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = "http://10.3.64.102:8080/web/checkLogin.htm"
    data = "username=qvz01&password=3XXLAg5rywhaTB16UOIGjQ%3D%3D&loginFlag=1"
    With http(0)
        .Open "POST", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "Origin", "http://10.3.64.102:8080"
        .setRequestHeader "Referer", "http://10.3.64.102:8080/web/login.htm"
        .setRequestHeader "Cookie", "critc_chk=1; critc_username=qvz01"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .send data
        Cookie = Split(.getResponseHeader("Set-Cookie"), ";")(0) & "; critc_chk=1; critc_username=qvz01; current=module_5282"
    End With
    
    '逐个车号查询
    Dim hts As String, M1 As Boolean, M2 As Boolean, M3 As Boolean, M4 As Boolean
    ReDim Preserve http(0 To rng.Rows.count)
    ReDim Preserve https(0 To rng.Rows.count)
    ReDim Preserve 票据号(0 To rng.Rows.count)
    ReDim Preserve 车号(0 To rng.Rows.count)
    For Each rn In rng '车号查询-摘要
        i = i + 1
        Set http(i) = CreateObject("WinHttp.WinHttpRequest.5.1") 'CreateObject("MSXML2.XMLHTTP.6.0") '
        With http(i)
            车号(i) = Format(rn.value, "0000000")
            If Left(车号(i), 1) > 0 And Left(车号(i), 1) <> 6 Then
'                Sheet13.Cells(i + 1, "P").value = "发送请求..."
                rngs.Cells(i, 1).value = "发送请求..."
                url = "http://10.3.64.102:8080/web/qlc/fx/index.htm?ch=" & 车号(i)
                .Open "get", url, True
                .setRequestHeader "Cookie", Cookie
                .setRequestHeader "Referer", url
                .send
            Else
                Set http(i) = Nothing
            End If
        End With
    Next
    For i = 1 To rng.Rows.count '获取票据号
        hts = ""
        If Not http(i) Is Nothing Then
            With http(i)
                Do
                On Error Resume Next
        '            st = .Status
                    hts = .responseText
                    ht.Body.innerHTML = .responseText
                    DoEvents
                    
    '                Debug.Print "1"
                On Error GoTo 0
                Loop Until hts <> ""
'                保存.桌面 hts
                Set Table = ht.getElementById("portlet_tab1") '时1时2
                Set A = Table.getElementsByTagName("a")
                If A.length > 0 Then
'                    Sheet13.Cells(i + 1, "P").value = "识别票据号..."
                    rngs.Cells(i, 1).value = "识别票据号..."
                    票据号(i) = A(0).innerText
                    url = "http://10.1.203.120:8105/dzpjfw/dzpjzzqd.htm?appid=10001&pjdh=" & 票据号(i)
                    Set http(i) = CreateObject("WinHttp.WinHttpRequest.5.1")
                    With http(i)
                        .Open "get", url, True
                        .setRequestHeader "Cookie", Cookie
                        .setRequestHeader "Referer", "http://10.3.64.102:8080/"
                        .send '久
                    End With
                Else
                    Set http(i) = Nothing
'                    Sheet13.Cells(i + 1, "P").value = ""
                    rngs.Cells(i, 1).value = ""
                End If
            End With
        End If
    Next
    
    '请求数据
       For i = 1 To rng.Rows.count '装载清单/运单1
        hts = ""
        If Not http(i) Is Nothing Then
            With http(i)
            Do
                On Error Resume Next
                hts = .responseText
                ht.Body.innerHTML = .responseText
                DoEvents
                On Error GoTo 0
            Loop Until hts <> ""
            If InStr(hts, "未授权或无效访问") = 0 Then
'                Sheet13.Cells(i + 1, "p") = "查询运单..."
                rngs.Cells(i, 1).value = "查询运单..."
                Set A = ht.getElementsByTagName("a")
                If A.length > 0 Then '集装箱
                    运单号 = Split(Split(A(1).outerHTML, "ydhm=")(1), "'")(0)
                    url = "http://10.1.203.120:8105/hyreport/reports/reportJsp/showReport.jsp?raq=hwyd1.raq&ydhm=" & 运单号
                    Set https(i) = CreateObject("WinHttp.WinHttpRequest.5.1")
                    With https(i)
                        .Open "get", url, True
                        .setRequestHeader "Referer", "http://10.1.203.120:8105/dzpjfw/dzpjzzqd.htm?appid=10001&pjdh=" & 票据号(i)
                        .setRequestHeader "Cookie", Cookie
                        .send '久/2
                        
                    End With
                Else 'If a.Length = 1 Then '散货
                    url = "http://10.1.203.120:8105/hyreport/reports/reportJsp/showReport.jsp?raq=hwyd1.raq&ydhm=" & 票据号(i)
                    Set https(i) = CreateObject("WinHttp.WinHttpRequest.5.1")
                    With https(i)
                        .Open "get", url, True
                        .setRequestHeader "Referer", "http://10.1.203.120:8105/dzpjfw/dzpjyd.htm?appid=10001&pjdh=" & 票据号(i)
                        .setRequestHeader "Cookie", Cookie
                        .send '久/2
                    End With
                End If
            Else
                Set http(i) = Nothing
'                Sheet13.Cells(i + 1, "P").value = ""
                rngs.Cells(i, 1).value = ""
            End If
            End With
        End If
    Next
    
       For i = 1 To rng.Rows.count
       hts = ""
'       If Not Sheet13.Cells(i + 1, "P").Comment Is Nothing Then Sheet13.Cells(i + 1, "P").Comment.Delete '删除批注
       If Not rngs.Cells(i, 1).Comment Is Nothing Then rngs.Cells(i, 1).Comment.Delete '删除批注
           If Not https(i) Is Nothing Then
'            Sheet13.Cells(i + 1, "p") = "正在查询专用线..."
                rngs.Cells(i, 1) = "正在查询专用线..."
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
                    文件名 = paths & "\" & 车号(i) & ".html"
                    Open 文件名 For Output As #1
                    Print #1, .responseText
                    Close #1
                    
                    '写入批注
                    Dim Ct As Comment
'                    Set Ct = Sheet13.Cells(i + 1, "P").AddComment
                    Set Ct = rngs.Cells(i, 1).AddComment
                    Ct.text text:="Momda提示：" & Chr(10) & "双击查看运单详情！"
                    
                    '提取信息到表格中
        '            ht.body.innerHTML = hts
                    Set tds = ht.getElementById("report1_A26")
                    Set trs = ht.getElementsByTagName("tr")
                    Set tds1 = trs(28).getElementsByTagName("td")
                    Set tds2 = trs(31).getElementsByTagName("td")
            '            Debug.Print tds(5).innerText
                    If tds1(3).innerText = "钦州港（宁）" Then
                        If tds1(5).innerText = "" Then
'                            Sheet13.Cells(i + 1, "p").value = "钦州港-货场"
                            rngs.Cells(i, 1).value = "钦州港-货场"
                        Else
'                            Sheet13.Cells(i + 1, "p").value = tds1(5).innerText
'                            Sheet13.Cells(i + 1, "R").value = tds2(2).innerText
                            rngs.Cells(i, 1).value = Split(tds1(5).innerText, "]")(1) '专用线
                            If rngs.Columns.count > 1 Then rngs.Cells(i, 3).value = tds2(2).innerText '货主
                        End If
                    Else
'                        Sheet13.Cells(i + 1, "p").value = tds1(3).innerText
                        rngs.Cells(i, 1).value = tds1(3).innerText
                    End If
                End With
            End If
        Next
        
'    ActiveSheet.UsedRange.EntireColumn.AutoFit
    Application.EnableEvents = True
    MsgBox "查询完成！"
End Sub

Sub 结束查询()
    Application.EnableEvents = True
    End
End Sub

Function 创建文件夹(path As String)

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
    
End Function


