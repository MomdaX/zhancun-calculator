Attribute VB_Name = "年休假统计表_异步"
Sub 年休假统计_文件读取()
    Dim d As New Dictionary, rn As Range, rng As Range
    Dim wb As Workbook, path As String
    On Error Resume Next
    
    '//选择要导入数据的工作簿文件
    path = Application.GetOpenFilename("textfiles(*.xls*),*.xls*")
    If path = "False" Then Exit Sub  '没有选择到文件时，退出程序
      
    'Set wb = Workbooks.Open(Path)     '打开工作薄获取表格内容
    Set wb = GetObject(path)         '不打开工作薄获取表格内容
    prr = wb.Sheets("sheet1").[B5].CurrentRegion
    Set wb = Nothing                 '//释放工作簿变量内存
    
    'brr = array("姓名", "职务", "假别", "请假始-假期止", "假期天数", "所在部门", "当前状态")
    
    
    
    '//筛选符合条件的数据装入数组(钦州港运转)
    For i = 4 To UBound(prr, 1) - 1
        If prr(i, 5) = "年休假" And prr(i, 12) = "段人力科同意" And prr(i, 9) = "钦州港运转" Then
        
            If d.Exists(prr(i, 3)) Then
                If d(prr(i, 3))(3) Like "*" & prr(i, 6) & "*" Then
                Else
                '姓名=姓名，岗位，假期，开始时间-结束时间 休：天数
                d(prr(i, 3)) = Array(prr(i, 3), prr(i, 4), prr(i, 5), prr(i, 6) & "-" & prr(i, 7) & "休:" & prr(i, 8) & Chr(10) & d(prr(i, 3))(3), d(prr(i, 3))(4) + prr(i, 8), prr(i, 9), prr(i, 12))
                End If
            Else
                d(prr(i, 3)) = Array(prr(i, 3), prr(i, 4), prr(i, 5), prr(i, 6) & "-" & prr(i, 7) & "休:" & prr(i, 8), prr(i, 8), prr(i, 9), prr(i, 12), prr(i, 6))
            End If
            
        End If
    Next
    
    '//筛选钦州港东
    For i = 3 To UBound(prr, 1) - 1
        If prr(i, 5) = "年休假" And prr(i, 12) = "段人力科同意" And prr(i, 10) = "钦州港东" Then
        
            If d.Exists(prr(i, 3)) Then
                If d(prr(i, 3))(3) Like "*" & prr(i, 6) & "*" Then
                Else
                d(prr(i, 3)) = Array(prr(i, 3), prr(i, 4), prr(i, 5), prr(i, 6) & "-" & prr(i, 7) & "休:" & prr(i, 8) & Chr(10) & d(prr(i, 3))(3), d(prr(i, 3))(4) + prr(i, 8), prr(i, 10), prr(i, 12))
                End If
            Else
            d(prr(i, 3)) = Array(prr(i, 3), prr(i, 4), prr(i, 5), prr(i, 6) & "-" & prr(i, 7) & "休:" & prr(i, 8), prr(i, 8), prr(i, 10), prr(i, 12), prr(i, 6))
            End If
            
        End If
    Next
    
    '//筛选马皇
    For i = 3 To UBound(prr, 1) - 1
        If prr(i, 5) = "年休假" And prr(i, 12) = "段人力科同意" And prr(i, 10) = "马皇" Then
        
            If d.Exists(prr(i, 3)) Then
                If d(prr(i, 3))(3) Like "*" & prr(i, 6) & "*" Then
                Else
                d(prr(i, 3)) = Array(prr(i, 3), prr(i, 4), prr(i, 5), prr(i, 6) & "-" & prr(i, 7) & "休:" & prr(i, 8) & Chr(10) & d(prr(i, 3))(3), d(prr(i, 3))(4) + prr(i, 8), prr(i, 10), prr(i, 12))
                End If
            Else
            d(prr(i, 3)) = Array(prr(i, 3), prr(i, 4), prr(i, 5), prr(i, 6) & "-" & prr(i, 7) & "休:" & prr(i, 8), prr(i, 8), prr(i, 10), prr(i, 12), prr(i, 6))
            End If
            
        End If
    Next
    
    '//测试数据写入Sheet3中
    'Sheets("sheet3").[a1].Resize(1, 7) = brr
    'Sheets("sheet3").[a2].Resize(d.Count, 7) = WorksheetFunction.Transpose(WorksheetFunction.Transpose(d.Items))
    
    For y = 2 To 16 Step 4
    '//获取姓名列区域
    Set Qy = Range(Cells(3, y), Cells(Rows.count, y).End(3))
    For Each rn In Qy
        '//将数据写入对应的单元格
        rn.Offset(0, 1) = d(rn.value)(1) '//职名
        If rn.Offset(0, 2) = "" Then rn.Offset(0, 2) = d(rn.value)(3) '//年休时间
        If rn.Offset(0, 3) = "" Then rn.Offset(0, 3) = d(rn.value)(4) '//假期天数
    
    Next
    
    
    For Each rng In Qy.Offset(0, 2)
        If rng <> "" Then           '//非空的单元格设置为青色
            With rng.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlauttomatic
            .Color = 5296274 '//青色
    '        .TintAndShade = 0
    '        .PatternTintAndShade = 0
            End With
            Else                     '//否则不设置颜色
             With rng.Interior
            .Pattern = xlNone
    '        .TintAndShade = 0
    '        .PatternTintAndShade = 0
        End With
        End If
    Next
    Next
    
    Set Qy = Nothing

'//设置单元格-格式
    With Sheets("sheet1").[A2].CurrentRegion
        .HorizontalAlignment = xlCenter '纵向居中
        .VerticalAlignment = xlCenter '横向居中
        .Borders.LineStyle = xlContinuous   '设置边框
    End With

End Sub

Sub 年休假统计表_异步_重置()
    '//清空所有颜色
    With Sheet27
        With .[A2].CurrentRegion.Interior
        .Pattern = xlNone
        End With
        
        '//清空假期天数
        For n = 5 To 17 Step 4
            sm = WorksheetFunction.Sum(.Range(Cells(3, n), .Cells(.Rows.count, n).End(3)))
            If sm <> 0 Then Range(.Cells(3, n), .Cells(Rows.count, n).End(3)).Clear
        Next
        
        '//清空年休时间
        For n = 4 To 17 Step 4
            sm = WorksheetFunction.CountA(.Range(Cells(3, n), .Cells(Rows.count, n).End(3)))
            If sm <> 0 Then Range(.Cells(3, n), .Cells(Rows.count, n).End(3)).Clear
        Next
    
    '//设置单元格-格式
        With .UsedRange
            .HorizontalAlignment = xlCenter '纵向居中
            .VerticalAlignment = xlCenter '横向居中
            .Borders.LineStyle = xlContinuous   '设置边框
        End With
        
    End With
End Sub

Sub 年休假统计表_异步_获取OA数据()
    Dim d As New Dictionary, t0, 页码%, 总页码%, t%, x%, y%, ht As New HTMLDocument, 假期天数%
    Dim http As Object, https() As Object, Cookie As String, reg As New RegExp, ckr(), oDom As Object, parameters As String
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    ow.execScript "function Number(){return Math.random()}"
    ow.execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
    tx = ow.eval("new Date().getTime()") '时间戳
    num = VBA.Format(ow.Number(""), "0.000000000000000") '随机数
    
    t0 = Timer '记录时间
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    With Sheet27
    
        With http
            '登录
            url = "http://10.190.136.222:8099/WebReport/ReportServer?reportlet=rs%2FQingJiaKaoQin%2F4KSCJ%2FQingJiaDanSouYe3.cpt"
            data = "tdpriv=102&tddept=206&op=&tdname=%5B97e6%5D%5B6587%5D%5B5eb7%5D"
            .Open "POST", url, False
            .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            .setRequestHeader "Referer", "http://10.190.136.8/"
            .send data
            
            '获取sid
            html1 = .responseText
    '        保存.桌面 html
            reg.Pattern = "FR.SessionMgr.register\('(\d+)', contentPane\)" '正则
            sid = reg.Execute(html1)(0).SubMatches(0) '返回结果中取sid
        End With
    
        '定义数组

        ReDim https(1 To 100)
'        On Error Resume Next
        For p = 1 To 100
            Set https(p) = CreateObject("MSXML2.ServerXMLHTTP") 'CreateObject("WinHttp.WinHttpRequest.5.1")
            With https(p)
                '发起请求
        '        url = "http://10.190.136.222:8099/WebReport/ReportServer?_=" & tx & "&__boxModel__=true&op=page_fit&sessionID=74451&pn=1&__webpage__=true&_paperWidth=580&_paperHeight=839&__fit__=false"
                url = "http://10.190.136.222:8099/WebReport/ReportServer?_=" & tx & "&__boxModel__=true&op=page_fit&sessionID=" & sid & "&pn=" & p
                .Open "GET", url, True
                .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
                .setRequestHeader "Referer", "http://10.190.136.222:8099/WebReport/ReportServer?reportlet=rs%2FQingJiaKaoQin%2F4KSCJ%2FQingJiaDanSouYe3.cpt"
                .send
        '        Debug.Print "返回：" & .responseText
        
'                html = .responseText
'                ht.body.innerHTML = .responseText
                Debug.Print "发送数据中....." & p
            End With
            DoEvents
        Next
        
        For i = 1 To 100
            DoEvents
            On Error Resume Next
            Do
            HTML = https(i).responseText
            ht.Body.innerHTML = https(i).responseText
            DoEvents
            .Range("A" & i).Select
            Debug.Print "等待数据中....." & i
            Loop Until HTML <> ""
            On Error GoTo 0
            
            If Not ht Is Nothing And HTML <> "" Then
            
                Debug.Print "开始写入数据中....."
                Set Table = ht.getElementById("frozen-center") '内容表
            
    '            Set tags = ht.getElementById("r-3-0").getElementsByTagName("tr") '取页码数
    '            reg.Pattern = "FR._p.reportTotalPage=(\d+);" '正则
    '            总页码 = reg.Execute(html)(0).SubMatches(0) '返回总页码
                
                Set trs = Table.getElementsByTagName("tr") '遍历内容
    
                For Each tr In trs
                    Set tds = tr.getElementsByTagName("td")
                    If tds.length > 1 Then 'brr = Array("姓名", "职务", "假别", "请假始-假期止", "假期天数", "所在部门", "当前状态")
                        姓名 = tds(5).innerText '姓名-3
                        职务 = tds(6).innerText '职务-4
                        假别 = tds(8).innerText '假别-5
                        开始时间 = tds(10).innerText '开始时间-6
                        结束时间 = tds(12).innerText '结束时间-7
                        假期天数 = tds(14).innerText '假期天数-8
                        所在部门 = tds(16).innerText '所在部门-9
                        当前状态 = tds(22).innerText '当前状态-12
                        Debug.Print 姓名
                        
                        '时间格式处理
                        If Year(开始时间) = Year(结束时间) Then
                            时间1 = Format(开始时间, "yyyy.mm.dd")
                            时间2 = Format(结束时间, "mm.dd")
                        Else
                            时间1 = Format(开始时间, "yyyy.mm.dd")
                            时间2 = Format(结束时间, "yyyy.mm.dd")
                        End If
                        
                        '条件判断
                        M1 = 假别 = "年休假" And 当前状态 = "段人力科同意"
                        M2 = 所在部门 = "钦州港运转" Or 所在部门 = "钦州港东" Or 所在部门 = "马皇"
                        M3 = val(Year(开始时间)) = val(.Range("A1").value)
    
                        If M1 And M2 And M3 Then
                        
                            If d.Exists(姓名) Then
                                If d(姓名)(3) Like "*" & 时间1 & "*" Then
                                Else
                                '姓名=姓名，岗位，假期，开始时间-结束时间 休：天数
    '                            If 姓名 = "詹鋆亮" Then Stop '调试
                                Debug.Print 姓名 & 假期天数
                                d(姓名) = Array(姓名, 职务, 假别, 时间1 & "-" & 时间2 & "休:" & 假期天数 & Chr(10) & d(姓名)(3), d(姓名)(4) + 假期天数, 所在部门, 当前状态)
                                End If
                            Else
                                d(姓名) = Array(姓名, 职务, 假别, 时间1 & "-" & 时间2 & "休:" & 假期天数, 假期天数, 所在部门, 当前状态, 时间1)
                            End If
                            
                        End If
    
                    End If
        
                    截止时间 = Format(DateSerial(val(.Range("A1").value), "01", "01") - 1, "yyyy-mm-dd")
                    If 结束时间 < 截止时间 Then
                        Exit For
                    End If
                Next
                
                '清除 key
                HTML = ""
                Set ht = Nothing
'                DoEvents
                
                If 结束时间 < 截止时间 Then
                    Exit For
                End If
                
            End If
            
        Next

        
        '重置表格区域
        Call 年休假统计表_重置
    
        For y = 2 To 16 Step 4
            '//获取姓名列区域
            Set Qy = .Range(Cells(3, y), .Cells(Rows.count, y).End(3))
            On Error Resume Next '错误处理
            For Each rn In Qy
                '//将数据写入对应的单元格
                rn.Offset(0, 1) = d(rn.value)(1) '//职名
                If rn.Offset(0, 2) = "" Then rn.Offset(0, 2) = d(rn.value)(3) '//年休时间
                If rn.Offset(0, 3) = "" Then rn.Offset(0, 3) = d(rn.value)(4) '//假期天数
            Next
            On Error GoTo 0
            
            For Each rng In Qy.Offset(0, 2)
                If rng <> "" Then           '//非空的单元格设置为青色
                    With rng.Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlauttomatic
                    .Color = 5296274 '//青色
            '        .TintAndShade = 0
            '        .PatternTintAndShade = 0
                    End With
                    Else                     '//否则不设置颜色
                     With rng.Interior
                    .Pattern = xlNone
            '        .TintAndShade = 0
            '        .PatternTintAndShade = 0
                End With
                End If
            Next
        Next
        DoEvents
    End With
    
    Set Qy = Nothing

'//设置单元格-格式
    With Sheet27.Range("A2").CurrentRegion
        .HorizontalAlignment = xlCenter '纵向居中
        .VerticalAlignment = xlCenter '横向居中
        .Borders.LineStyle = xlContinuous   '设置边框
    End With
    
    MsgBox "完成！" & Chr(10) & "用时:" & Format(Timer - t0, "0.00秒")
 End Sub
'curl 'http://10.190.136.222:8099/WebReport/ReportServer?reportlet=rs%2FQingJiaKaoQin%2F4KSCJ%2FQingJiaDanSouYe3.cpt' \
'  -H 'Connection: keep-alive' \
'  -H 'Cache-Control: max-age=0' \
'  -H 'Upgrade-Insecure-Requests: 1' \
'  -H 'Origin: http://10.190.136.8' \
'  -H 'Content-Type: application/x-www-form-urlencoded' \
'  -H 'User-Agent: Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36' \
'  -H 'Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9' \
'  -H 'Referer: http://10.190.136.8/' \
'  -H 'Accept-Language: zh-CN,zh;q=0.9' \
'  --data-raw 'tdpriv=102&tddept=206&op=&tdname=%5B97e6%5D%5B6587%5D%5B5eb7%5D' \
'  --compressed \
'  --insecure

'curl 'http://10.190.136.222:8099/WebReport/ReportServer?_=1733461937607&__boxModel__=true&op=page_fit&sessionID=24517&pn=1&__webpage__=true&_paperWidth=580&_paperHeight=839&__fit__=false' \
'  -H 'Connection: keep-alive' \
'  -H 'Accept: text/html, */*; q=0.01' \
'  -H 'User-Agent: Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36' \
'  -H 'X-Requested-With: XMLHttpRequest' \
'  -H 'Referer: http://10.190.136.222:8099/WebReport/ReportServer?reportlet=rs%2FQingJiaKaoQin%2F4KSCJ%2FQingJiaDanSouYe3.cpt' \
'  -H 'Accept-Language: zh-CN,zh;q=0.9' \
'  --compressed \
'  --insecure
