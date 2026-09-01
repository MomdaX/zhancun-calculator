Attribute VB_Name = "OA安全库"
'curl 'http://10.190.136.222:8099/WebReport/ReportServer?op=fr_dialog&cmd=parameters_d&sessionID=81745' \ POST
'  -H 'Connection: keep-alive' \
'  -H 'Accept: */*' \
'  -H 'X-Requested-With: XMLHttpRequest' \
'  -H 'User-Agent: Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36' \
'  -H 'Content-Type: application/x-www-form-urlencoded; charset=UTF-8' \
'  -H 'Origin: http://10.190.136.222:8099' \
'  -H 'Referer: http://10.190.136.222:8099/WebReport/ReportServer?reportlet=gwjc/wtk/wtcx.cpt' \
'  -H 'Accept-Language: zh-CN,zh;q=0.9' \
'  --data-raw '__parameters__=%7B%22LABEL1_C_C_C_C_C%22%3A%22%5B5f00%5D%5B59cb%5D%5B65f6%5D%5B95f4%5D%22%2C%22%5B8d77%5D%5B65e5%5D%5B671f%5D%22%3A%222024-10-07+19%3A10%3A21%22%2C%22LABEL1_C_C_C_C%22%3A%22%5B7ed3%5D%5B675f%5D%5B65f6%5D%5B95f4%5D%22%2C%22%5B6b62%5D%5B65e5%5D%5B671f%5D%22%3A%222024-10-17+19%3A10%3A21%22%2C%22LABEL1_C_C_C_C_C_C_C_C_C_C%22%3A%22%5B4e13%5D%5B4e1a%5D%5B5206%5D%5B7c7b%5D%22%2C%22%5B4e13%5D%5B4e1a%5D%5B5206%5D%5B7c7b%5D%22%3A%22%22%2C%22LABEL1_C_C_C_C_C_C_C_C_C%22%3A%22%5B6d41%5D%5B6c34%5D%22%2C%22%5B6d41%5D%5B6c34%5D%22%3A%22%22%2C%22LABEL0%22%3A%22%5B68c0%5D%5B67e5%5D%5B4eba%5D%5B59d3%5D%5B540d%5D%22%2C%22JCRXM%22%3A%22%22%2C%22LABEL1_C_C_C%22%3A%22%5B88ab%5D%5B68c0%5D%5B8f66%5D%5B95f4%5D%22%2C%22%5B8f66%5D%5B95f4%5D%22%3A%2265%22%2C%22LABEL1_C_C_C_C_C_C%22%3A%22%5B88ab%5D%5B68c0%5D%5B90e8%5D%5B95e8%5D%22%2C%22LABEL1_C_C_C_C_C_C_C%22%3A%22%5B9879%5D%5B76ee%5D%22%2C%22%5B9879%5D%5B76ee%5D%22%3A%22%22%2C%22LABEL1_C_C_C_C_C_C_C_C%22%3A%22%5B5b9a%5D%5B6027%5D%22%2C%22LABEL1_C_C_C_C_C_C
'C_C_C_C_C%22%3A%22%5B653f%5D%5B6cbb%5D%5B9762%5D%5B8c8c%5D%22%2C%22%5B68c0%5D%5B67e5%5D%5B90e8%5D%5B95e8%5D%22%3A%22204%22%2C%22%5B653f%5D%5B6cbb%5D%5B9762%5D%5B8c8c%5D%22%3A%22%22%2C%22%5B5b9a%5D%5B6027%5D%22%3A%22%22%2C%22LABEL1%22%3A%22%5B8d23%5D%5B4efb%5D%5B4eba%5D%5B59d3%5D%5B540d%5D%22%2C%22ZRRXM%22%3A%22%22%2C%22LABEL1_C_C_C_C_C_C_C_C_C_C_C_C%22%3A%22%5B5e74%5D%5B9f84%5D%22%2C%22LABEL1_C_C_C_C_C_C_C_C_C_C_C_C_C%22%3A%22%5B81f3%5D%22%2C%22KSSJ%22%3A%222024-10-07+19%3A10%3A21%22%2C%22LABEL2_C%22%3A%22%5B81f3%5D%22%2C%22JSSJ%22%3A%222024-10-17+19%3A10%3A21%22%2C%22LABEL1_C%22%3A%22%5B5f55%5D%5B5165%5D%5B65f6%5D%5B95f4%5D%22%2C%22DI%22%3A%22%22%2C%22GAO%22%3A%22%22%2C%22YXWT%22%3Afalse%2C%22LABEL3_C%22%3A%22%5B53ea%5D%5B67e5%5D%5B6709%5D%5B6548%5D%5B95ee%5D%5B9898%5D%22%7D' \
'  --compressed \
'  --insecure

'岗位检查统计
'curl 'http://10.190.136.222:8099/WebReport/ReportServer?_=1737743366147&__boxModel__=true&op=fr_write&cmd=read_w_content&sessionID=22932&reportIndex=0&browserWidth=1093&iid=0.21494656926313316&__cutpage__=&pn=1' \
'  -H 'Connection: keep-alive' \
'  -H 'Accept: text/html, */*; q=0.01' \
'  -H 'User-Agent: Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36' \
'  -H 'X-Requested-With: XMLHttpRequest' \
'  -H 'Referer: http://10.190.136.222:8099/WebReport/ReportServer?reportlet=/gwjc/gwjc/[5b89][5168][5206][6790].cpt&__parameters__={%22_%22:%221737743260377%22,%22__pi__%22:%22true%22,%22tdname%22:%22%25E9%259F%25A6%25E6%2596%2587%25E5%25BA%25B7%22,%22op%22:%22write%22,%22tdpriv%22:%22102%22,%22tddept%22:%22206%22}' \
'  -H 'Accept-Language: zh-CN,zh;q=0.9' \
'  --compressed \
'  --insecure
'_=1737743366147&__boxModel__=true&op=fr_write&cmd=read_w_content&sessionID=22932&reportIndex=0&browserWidth=1093&iid=0.21494656926313316&__cutpage__=&pn=1

Sub 获取() '安全问题库-获取指定字段
    Dim dt As New Dictionary, dr As New Dictionary, dx As New Dictionary, t0, 月%
    Dim http As Object, Cookie As String, reg As New RegExp, ckr(), oDom As Object, parameters As String
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    ow.execScript "function Number(){return Math.random()}"
    ow.execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
    tx = ow.eval("new Date().getTime()") '时间戳
    num = VBA.Format(ow.Number(""), "0.000000000000000") '随机数

    '设置要保留的标题
    trr = Split("录入时间,检查项目,检查方式,检查人,检查人职务,检查人属性,责任职名,责任人,政治面貌,年龄,考核定性,检查内容", ",")
    For Each tr In trr
        dt(tr) = 0
    Next
    
    月 = Sheet25.Range("Q1").value
    If 月 = 0 Then
        月 = -1
    ElseIf IsNumeric(月) Then
        月 = 月 * -1
    Else
        MsgBox "设置参数有误"
    End If
    t0 = Timer
    
    Application.ScreenUpdating = False
'   Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    With http
        '登录
        url = "http://10.190.136.222:8099/WebReport/ReportServer?reportlet=gwjc/wtk/wtcx.cpt"
        data = "tdname=%5B97e6%5D%5B6587%5D%5B5eb7%5D&tddept=206&tdpriv=102&op="
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .setRequestHeader "Referer", "http://10.190.136.8/"
        .send data
        
        '获取sid
        resHtml = .responseText
        reg.Pattern = "FR.SessionMgr.register\('(\d+)', contentPane\)" '正则
        sid = reg.Execute(resHtml)(0).SubMatches(0) '返回结果中取sid
        Kssj = VBA.Format(DateAdd("m", 月, Now), "yyyy-mm-dd hh:mm:ss") '开始时间
        jssj = VBA.Format(Now(), "yyyy-mm-dd hh:mm:ss") '结束时间
        
        '设置请求体data
        parameters = "{""LABEL1_C_C_C_C_C"":""[5f00][59cb][65f6][95f4]"",""[8d77][65e5][671f]"":""" & Kssj & """,""LABEL1_C_C_C_C"":""[7ed3][675f][65f6][95f4]"",""[6b62][65e5][671f]"":""" & jssj & """,""LABEL1_C_C_C_C_C_C_C_C_C_C"":""[4e13][4e1a][5206][7c7b]"",""[4e13][4e1a][5206][7c7b]"":""381"",""LABEL1_C_C_C_C_C_C_C_C_C"":""[6d41][6c34]"",""[6d41][6c34]"":"""",""LABEL0"":""[68c0][67e5][4eba][59d3][540d]"",""JCRXM"":"""",""LABEL1_C_C_C"":""[88ab][68c0][8f66][95f4]"",""[8f66][95f4]"":""65"",""LABEL1_C_C_C_C_C_C"":""[88ab][68c0][90e8][95e8]"",""LABEL1_C_C_C_C_C_C_C"":""[9879][76ee]"",""[9879][76ee]"":"""",""LABEL1_C_C_C_C_C_C_C_C"":""[5b9a][6027]"",""LABEL1_C_C_C_C_C_C_C_C_C_C_C"":""[653f][6cbb][9762][8c8c]"",""[68c0][67e5][90e8][95e8]"":""204"",""[653f][6cbb][9762][8c8c]"":"""",""[5b9a][6027]"":"""",""LABEL1"":""[8d23][4efb][4eba][59d3][540d]"",""ZRRXM"":"""",""LABEL1_C_C_C_C_C_C_C_C_C_C_C_C"":""[5e74][9f84]"",""LABEL1_C_C_C_C_C_C_C_C_C_C_C_C_C"":""[81f3]""," & _
"""KSSJ"":""" & Kssj & """,""LABEL2_C"":""[81f3]"",""JSSJ"":""" & jssj & """,""LABEL1_C"":""[5f55][5165][65f6][95f4]"",""DI"":"""",""GAO"":"""",""YXWT"":false,""LABEL3_C"":""[53ea][67e5][6709][6548][95ee][9898]""}"
        url = "http://10.190.136.222:8099/WebReport/ReportServer?op=fr_dialog&cmd=parameters_d&sessionID=" & sid
        data = "__parameters__=" & ow.encode(parameters)
        
        '发起请求
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .setRequestHeader "Origin", "http://10.190.136.222:8099"
        .setRequestHeader "Referer", "Referer: http://10.190.136.222:8099/WebReport/ReportServer?reportlet=gwjc/wtk/wtcx.cpt"
        .send data
'        Debug.Print "返回：" & .responseText '无返回
    End With
    
    Dim 页码%, 总页码%, t%, t1, t2, t3, t4, x%, y%, ht As New HTMLDocument, brr()
    
    '记录登录用时
    t1 = Format(Timer - t0, "0.00秒")
    t2 = Timer
    
    Sheet25.Cells.Clear '内容清空

    With http
        '逐页请求读取
        For 页码 = 1 To 100
'            url = "http://10.190.136.222:8099/WebReport/ReportServer?_=" & tx & "&__boxModel__=true&op=page_fit&sessionID=" & sid & "&pn=" & 页码
            url = "http://10.190.136.222:8099/WebReport/ReportServer?_=" & tx & "&op=page_fit&sessionID=" & sid & "&pn=" & 页码
            .Open "get", url, False
            .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            .setRequestHeader "Referer", "http://10.190.136.222:8099/WebReport/ReportServer?reportlet=gwjc/wtk/wtcx.cpt"
    '        .setRequestHeader "cookie", cookie
            .send
            
            resHtml = .responseText
            ht.Body.innerHTML = .responseText
            
            With Sheet25
            
                Set Table = ht.getElementById("frozen-center") '内容表
                
                If t = 0 Then
                    Set tags = ht.getElementById("r-3-0").getElementsByTagName("tr") '设置标题
                    reg.Pattern = "FR._p.reportTotalPage=(\d+);" '正则
                    总页码 = reg.Execute(resHtml)(0).SubMatches(0) '返回总页码
                    
                    For Each Tag In tags
                        t = t + 1
'                        .Cells(1, t).value = Tag.innerText
                        tr = Tag.innerText
                        If dt.Exists(tr) Then
                            dt(t) = tr
                            y = y + 1
                            ReDim Preserve brr(1 To tags.length * 总页码, 1 To y)
                            brr(1, y) = Tag.innerText
                        End If
                        DoEvents
                    Next
                    x = 1
                    y = 0
                End If
                
                Set trs = Table.getElementsByTagName("tr") '遍历内容
                On Error Resume Next
                For Each tr In trs
    '                x = .Cells(Rows.count, "A").End(3).Row + 1
                    Set tds = tr.getElementsByTagName("tr")
                    If tds.length > 0 Then
                        If tds(20).innerText <> "外单位" Then '责任职名
                            x = x + 1
'                            If x = 56 Then Stop
                            For Each td In tds
                                i = i + 1
'                                .Cells(x, i).value = td.innerText
                                If dt.Exists(i) Then '表头对比
                                    y = y + 1
                                    If y = 1 Then
                                        brr(x, y) = Format(td.innerText, "yyyy年mm月dd日")
                                    Else
                                        If y = 9 Then
                                            dr(td.innerText) = dr(td.innerText) + 1 '记录责任人用于统计次数
                                        End If
                                        brr(x, y) = td.innerText
                                    End If
                                End If
                                DoEvents
                            Next
                            y = 0
                            i = 0
                        End If
                    End If
                Next
                On Error GoTo 0
            End With
            If 页码 > 总页码 Then
                Exit For
            End If
        Next
    End With

    Dim rng As Range, rn As Range, rngs As Range
    With Sheet25
        .Range("N1:P1") = Array("姓名", "统计", "星级")
        Set rng = .Range("N1:P1").CurrentRegion '表头
        With rng
            .Cells(2, 1).Resize(dr.count, 1) = WorksheetFunction.Transpose(dr.keys) '姓名
            .Cells(2, 2).Resize(dr.count, 1) = WorksheetFunction.Transpose(dr.items) '统计
            Set dx = 获取星级()
            On Error Resume Next
            For Each rn In .CurrentRegion.Columns(1).Cells
                姓名 = rn.value
                rn.Offset(0, 2).value = dx(Trim(姓名))(1)
                '测试
'                Sheet25.Cells(2, "M").Resize(dx.count, 1) = WorksheetFunction.Transpose(dx.Keys) '姓名
            Next
            On Error GoTo 0
            With .CurrentRegion
                .Borders.LineStyle = xlContinuous '设置样式边框
                .HorizontalAlignment = xlCenter '设置上下居中
                .VerticalAlignment = xlCenter '设置左右居中
            End With
            升序 Sheet25, .Rows(1), .Cells(1, 2), 2 '表对象，标题行，排序列单元格，1升2降
            With .Rows(1)
                .Interior.ColorIndex = 14
                .Font.ColorIndex = 2
            End With
        End With
        
        '取消筛选
        If .AutoFilterMode Then
            .AutoFilterMode = False
        End If
        .Range("A1").Resize(UBound(brr, 1), UBound(brr, 2)) = brr
        Set rng = .Cells(1, 1).CurrentRegion.Rows(1)
'        升序 Sheet25, rng, .Cells(1, 1), 2 '表对象，标题行，排序列单元格，1升2降
        With .Cells(1, 1)
            .EntireRow.Font.Bold = True '标题加粗
            .RowHeight = 30 '行高
            .CurrentRegion.Borders.LineStyle = xlContinuous '设置样式边框
    '        .Cells(1, 1).CurrentRegion.EntireColumn.AutoFit
        End With
        With rng
            .HorizontalAlignment = xlCenter '设置上下居中
            .VerticalAlignment = xlCenter '设置左右居中
            .Interior.ColorIndex = 14
            .Font.ColorIndex = 2
            .AutoFilter
        End With
    End With
    
    Application.ScreenUpdating = True
'    Application.DisplayAlerts = True
    Application.EnableEvents = True
    '记录处理用时
    t3 = Format(Timer - t2, "0.00秒")
    t4 = Format(Timer - t0, "0.00秒")
    Sheet25.Range("Q1").value = 1
    MsgBox "查询完成！" & Chr(10) & "登录用时：" & t1 & Chr(10) & "查询处理用时：" & t3 & Chr(10) & "总的用时：" & t4 & Chr(10) & "近" & Abs(月) & "个月内共查询到：" & x - 1 & "条记录。"
End Sub


Sub SQL查询()
    Dim cnn As Object, trr(), rst As Object, i%, rs%, rng As Range
    Set cnn = CreateObject("adodb.connection")
    
    path = ThisWorkbook.FullName
    str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & path
    cnn.Open str_cnn '打开表格
    
    strSQL = "Select 录入时间,检查项目,检查方式,检查人,检查人职务,检查人属性,责任职名,责任人,政治面貌,年龄,考核定性,检查内容 FROM [Sheet1$]"
    Set rst = cnn.Execute(strSQL) '查找
    
    With ActiveSheet
            .Cells.Clear '重置
            For i = 1 To rst.fields.count
                .Cells(1, i) = rst.fields(i - 1).name '标题
            Next
            .Cells(1, 1).CurrentRegion.Font.Bold = True '标题加粗
            .Cells(2, 1).CopyFromRecordset rst '写入查找到的数据
            .Cells(1, 1).CurrentRegion.Borders.LineStyle = xlContinuous '设置样式边框
            .Cells(1, 1).CurrentRegion.EntireColumn.AutoFit
    End With
    
    rst.Close
    Set cnn = Nothing
End Sub
    
Function 保存(text As String)
    文件名 = "C:\Users\Administrator\Desktop\HTML.txt"
    Open 文件名 For Output As #1
    Print #1, text
    Close #1
End Function


Sub 获取全部内容() '安全问题库

    Dim http As Object, Cookie As String, reg As New RegExp, ckr(), oDom As Object, parameters As String
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    ow.execScript "function Number(){return Math.random()}"
    ow.execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
    tx = ow.eval("new Date().getTime()") '时间戳
'   num = VBA.Format(ow.Number(""), "0.000000000000000") '随机数
'   cookie = OA登录.获取Cookie

    Application.ScreenUpdating = False
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    With http
        url = "http://10.190.136.222:8099/WebReport/ReportServer?reportlet=gwjc/wtk/wtcx.cpt"
        data = "tdname=%5B97e6%5D%5B6587%5D%5B5eb7%5D&tddept=206&tdpriv=102&op="
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .setRequestHeader "Referer", "http://10.190.136.8/"
        .send data

        HTML = .responseText
        reg.Pattern = "FR.SessionMgr.register\('(\d+)', contentPane\)" '正则
        sid = reg.Execute(HTML)(0).SubMatches(0) '返回结果中取sid
        Kssj = VBA.Format(DateAdd("m", -1, Now), "yyyy-mm-dd hh:mm:ss")
        jssj = VBA.Format(Now(), "yyyy-mm-dd hh:mm:ss")

        parameters = "{""LABEL1_C_C_C_C_C"":""[5f00][59cb][65f6][95f4]"",""[8d77][65e5][671f]"":""" & Kssj & """,""LABEL1_C_C_C_C"":""[7ed3][675f][65f6][95f4]"",""[6b62][65e5][671f]"":""" & jssj & """,""LABEL1_C_C_C_C_C_C_C_C_C_C"":""[4e13][4e1a][5206][7c7b]"",""[4e13][4e1a][5206][7c7b]"":""381"",""LABEL1_C_C_C_C_C_C_C_C_C"":""[6d41][6c34]"",""[6d41][6c34]"":"""",""LABEL0"":""[68c0][67e5][4eba][59d3][540d]"",""JCRXM"":"""",""LABEL1_C_C_C"":""[88ab][68c0][8f66][95f4]"",""[8f66][95f4]"":""65"",""LABEL1_C_C_C_C_C_C"":""[88ab][68c0][90e8][95e8]"",""LABEL1_C_C_C_C_C_C_C"":""[9879][76ee]"",""[9879][76ee]"":"""",""LABEL1_C_C_C_C_C_C_C_C"":""[5b9a][6027]"",""LABEL1_C_C_C_C_C_C_C_C_C_C_C"":""[653f][6cbb][9762][8c8c]"",""[68c0][67e5][90e8][95e8]"":""204"",""[653f][6cbb][9762][8c8c]"":"""",""[5b9a][6027]"":"""",""LABEL1"":""[8d23][4efb][4eba][59d3][540d]"",""ZRRXM"":"""",""LABEL1_C_C_C_C_C_C_C_C_C_C_C_C"":""[5e74][9f84]"",""LABEL1_C_C_C_C_C_C_C_C_C_C_C_C_C"":""[81f3]""," & _
"""KSSJ"":""" & Kssj & """,""LABEL2_C"":""[81f3]"",""JSSJ"":""" & jssj & """,""LABEL1_C"":""[5f55][5165][65f6][95f4]"",""DI"":"""",""GAO"":"""",""YXWT"":false,""LABEL3_C"":""[53ea][67e5][6709][6548][95ee][9898]""}"
        url = "http://10.190.136.222:8099/WebReport/ReportServer?op=fr_dialog&cmd=parameters_d&sessionID=" & sid
        data = "__parameters__=" & ow.encode(parameters)

        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .setRequestHeader "Origin", "http://10.190.136.222:8099"
        .setRequestHeader "Referer", "Referer: http://10.190.136.222:8099/WebReport/ReportServer?reportlet=gwjc/wtk/wtcx.cpt"
        .send data
'        Debug.Print .responseText'无返回
    End With
    
    Dim 页码%, 总页码%, t%, ht As New HTMLDocument
    
    Sheet25.Cells.ClearContents '清空
    With http
        For 页码 = 1 To 100
        
        url = "http://10.190.136.222:8099/WebReport/ReportServer?_=" & tx & "&__boxModel__=true&op=page_fit&sessionID=" & sid & "&pn=" & 页码
'        url = "http://10.190.136.222:8099/WebReport/ReportServer?_=" & tx & "&__boxModel__=true&op=page_fit&sessionID=" & sid & "&pn=" & 页码 & "&__webpage__=true&_paperWidth=1920&_paperHeight=1080&__fit__=false"
        .Open "get", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .setRequestHeader "Referer", "http://10.190.136.222:8099/WebReport/ReportServer?reportlet=gwjc/wtk/wtcx.cpt"
'        .setRequestHeader "cookie", cookie
        .send
        
        HTML = .responseText
        ht.Body.innerHTML = .responseText
        
        
        With Sheet25
            Set Table = ht.getElementById("frozen-center") '内容表
            Set trs = Table.getElementsByTagName("tr") '内容
            If t = 0 Then
                Set trs = ht.getElementById("r-3-0").getElementsByTagName("tr") '标题
                For Each tr In trs
                    t = t + 1
                    .Cells(1, t).value = tr.innerText
                Next
            End If
            
            Set Table = ht.getElementById("frozen-center") '内容表
            Set trs = Table.getElementsByTagName("tr") '内容
            For Each tr In trs
                x = .Cells(Rows.count, "A").End(3).Row + 1
                For Each td In tr.getElementsByTagName("tr")
                    y = y + 1
                    .Cells(x, y).value = td.innerText
                    DoEvents
                Next
                y = 0
            Next
        End With
        
        If 总页码 = 0 Then
            reg.Pattern = "FR._p.reportTotalPage=(\d+);" '正则
            总页码 = reg.Execute(HTML)(0).SubMatches(0) '返回总页码
        ElseIf 页码 > 总页码 Then
            Exit For
        End If
        Next
    End With

    Dim rng As Range, rn As Range, rngs As Range
    With Sheet25
    Set rng = .Cells(1, 1).CurrentRegion.Rows(1)
        cs = rng.Columns.count
        For i = cs To 1 Step -1

        If Not "录入时间,检查项目,检查方式,检查人,检查人职务,检查人属性,责任职名,责任人,政治面貌,年龄,考核定性,检查内容" Like "*" & rng.Cells(i).value & "*" Then
'            Debug.Print rng.Cells(i).value
            rng.Cells(i).EntireColumn.Delete
        End If
        Next

        
        .Cells(1, 1).EntireRow.Font.Bold = True '标题加粗
        .Cells(1, 1).CurrentRegion.Borders.LineStyle = xlContinuous '设置样式边框
'       .Cells(1, 1).CurrentRegion.EntireColumn.AutoFit
    End With
    Application.ScreenUpdating = True

    '锁定首行
    Rows("1:1").Select
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    ActiveWindow.FreezePanes = True

    MsgBox "查询完成！"
End Sub

Function 获取星级()

    Dim xmlhttp As Object, oDom As Object, reg As New RegExp, d As New Dictionary, tx As Double
    Dim HTML As String, url As String, Body As HTMLBody, ht As New HTMLDocument
    
    Set oDom = CreateObject("htmlfile")
    Set ow = oDom.parentWindow
    ow.execScript "function Number(){return Math.random()}"
    ow.execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
'    tx = ow.eval("new Date().getTime()") '当前的时间戳
    Yf = DateSerial(Year(Date), Month(Date) - 1, 1) - TimeSerial(8, 0, 0)
    tx = (Yf - DateSerial(1970, 1, 1)) * 86400000 '指定时间的时间戳
    Debug.Print tx
    
    '获取sid
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    With http
        url = "http://10.190.136.16:8080/WebReport/ReportServer?reportlet=152.19/136oa3.cpt&op=write"
        .Open "get", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .setRequestHeader "Referer", "http://10.190.136.36/"
        .send

        HTML = .responseText
        reg.Pattern = "FR.SessionMgr.register\('(\d+)', contentPane\)" '正则
        sid = reg.Execute(HTML)(0).SubMatches(0) '返回结果中取sid
    End With
    
    
    '设置专业
    crr = Array("[8c03][8f66]", "[8c03][8f66] [52b3][52a1][6d3e][9063]", "[63a5][53d1][8f66]-[5927][7ad9]", "[8d27][8fd0]", "[63a5][53d1][8f66]-[5355][53cc][7ebf][7ec4]")
    For Each 专业 In crr
    
    '按专业查询
    parameters = "{""LABEL0_C"":""[4e13][4e1a]"",""ZY"":""" & 专业 & """,""RQ1111_C"":""[6708][4efd]"",""D1"":{""__time__"":" & tx & "}}" '1725120000000
'    Debug.Print parameters
    data = "__parameters__=" & ow.encode(parameters)
    
        url = "http://10.190.136.16:8080/WebReport/ReportServer?op=fr_dialog&cmd=parameters_d&sessionID=" & sid
        Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
        With xmlhttp
            .Open "POST", url, False
            .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
            .setRequestHeader "Referer", "http://10.190.136.16:8080/WebReport/ReportServer?reportlet=152.19%2F136oa3.cpt&op=write"
            .setRequestHeader "Origin", "http://10.190.136.16:8080"
            .send data
    
            HTML = .responseText
    '        Debug.Print .responseText
    
'            tx = ow.eval("new Date().getTime()") '时间戳
            iid = VBA.Format(ow.Number(""), "0.000000000000000") '随机数
            
            url = "http://10.190.136.16:8080/WebReport/ReportServer?_=" & tx & "&__boxModel__=true&op=fr_write&cmd=read_w_content&sessionID=" & sid & "&reportIndex=0&browserWidth=495&iid=" & iid & "&__cutpage__=&pn=1"
            .Open "get", url, False
            .setRequestHeader "Referer", "http://10.190.136.16:8080/WebReport/ReportServer?reportlet=152.19%2F136oa3.cpt&op=write"
            .send
            
            HTML = .responseText
            ht.Body.innerHTML = HTML
'            Call 保存(html)'调试用
        
            Set Table = ht.getElementsByTagName("tbody") '表内容
            Set trs = Table(0).getElementsByTagName("tr")
            '标题
'            Debug.Print trs(0).innerText
            '内容
            For i = 5 To trs.length - 1
                Set tds = trs(i).getElementsByTagName("td")
                姓名 = tds(1).innerText
                积分 = tds(13).innerText
                星级 = tds(14).innerText
                d(姓名) = Array(积分, 星级)
            Next
        End With
    Next
    Set 获取星级 = d
End Function

Sub 获取时间戳()
    Dim oDom As Object, ow As Object, tx As Double, Yf As Double

    Yf = DateSerial(Year(Date), Month(Date) - 1, 1) - TimeSerial(8, 0, 0)
    tx = (Yf - DateSerial(1970, 1, 1)) * 86400000
    Debug.Print tx
End Sub

Sub SQL筛选(姓名 As String)

    Dim cnn As Object, trr(), rst As Object, i%, rs%, rng As Range
    Set cnn = CreateObject("adodb.connection")
    With Sheet25
    If .Range("XES1") = "" Then
        trr = .Range("A1").CurrentRegion.value
        .Range("XES1").Resize(UBound(trr, 1), UBound(trr, 2)) = trr
    End If
'        姓名 = "蒋静静"
        path = ThisWorkbook.FullName
        str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & path
        cnn.Open str_cnn '打开表格
    
        strSQL = "Select * FROM [安全问题库$XES1:XFD] WHERE 责任人='" & 姓名 & "'"
        Set rst = cnn.Execute(strSQL) '查找
        
        With .Range("A1")
            .CurrentRegion.Offset(1).ClearContents
            .Offset(1).CopyFromRecordset rst
        End With
     End With
'    k = 2
'    With Sheet26
'
'        rst.MoveFirst
'        .Columns(16).Clear  '清除
'        With .Columns(15)
'            .Clear '清除
'            .Cells(k).value = "选择股道"
'            .HorizontalAlignment = xlCenter '设置上下居中
'            .VerticalAlignment = xlCenter '设置左右居中
'        End With
'        Do Until rst.EOF
'            k = k + 1
'            .Cells(k, "O").value = rst.fields(0).value
'            .Cells(k, "P").value = rst.fields(1).value
''            .Cells(k, "Q").value = rst.fields(2).value
'            rst.MoveNext
'        Loop
'    End With
    rst.Close
    Set cnn = Nothing
End Sub
