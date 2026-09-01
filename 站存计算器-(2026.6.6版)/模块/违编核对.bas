Attribute VB_Name = "违编核对"
Sub 出发刷新()
    Dim js As Object
    Application.EnableEvents = False
'    Application.ScreenUpdating = False '取消屏幕刷新
    
    '获取地磅的Cookie
'    url = "http://10.190.136.155:8888/Gsgz/login/userAuth"
'    用户名 = URL编码.编码("钦州港")
'    密码 = URL编码.编码("qzg123456#")
'    data = "User_Name=" & 用户名 & "&User_Password=" & 密码
'
'    With CreateObject("WinHttp.WinHttpRequest.5.1")
'        .Open "POST", url, False
'        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
'        .setRequestHeader "Origin", "http://10.190.136.155:8888"
'        .setRequestHeader "Host", "10.190.136.155:8888"
'        .setRequestHeader "Referer", "http://10.190.136.155:8888/Gsgz/loginPage"
'        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
'        .send data
'        dbck = Split(.getResponseHeader("Set-Cookie"), ";")(0)
'    End With
    
    Dim oDom As Object, ow As Object, result As String
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow

    Dim pathAes As String, AES As String
    pathAes = "http://10.190.168.62:80/nnjzj/resources/script/common/aes.js" '网络读取
    AES = 加载js(pathAes)
    ow.execScript AES
    帐号 = ow.encrypt("zqzcwd", "RailsYdjszz46495")
    密码 = ow.encrypt("Qcd123qwe!#@", "RailsYdjszz46495")
    
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = "http://yt1dzh.crc.cr:30001/authentication/login"
    data = "{""username"":""" & 帐号 & """,""password"":""" & 密码 & """}"
'    Debug.Print data
    With xhtp
        .Open "POST", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "Origin", "http://yt1dzh.crc.cr:30001"
        .setRequestHeader "Referer", "http://yt1dzh.crc.cr:30001/"
        .setRequestHeader "Cookie", "GUEST_LANGUAGE_ID=zh_CN"
        .send data
        
        '动态的参数
        Dim Cookie$, 开始时间$, 结束时间$
        Cookie = Split(.getResponseHeader("Set-Cookie"), ";")(0)
        开始时间 = VBA.Format(DateAdd("h", -15, Now), "yyyy-mm-dd%20hh:mm") '查询时间后推25小时
        结束时间 = VBA.Format(DateAdd("d", 1, Now), "yyyy-mm-dd%20hh:mm") '次日当前时间
        
        url = "http://yt1dzh.crc.cr:30001/qb/getqbmlforallByPage?qbType=2&begin=" & 开始时间 & "&end=" & 结束时间 & "&pageNum=1&pageSize=100&cdm=*&ljdm=10&departmentId=99113"
        .Open "POST", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "Origin", "http://yt1dzh.crc.cr:30001"
        .setRequestHeader "Referer", "http://yt1dzh.crc.cr:30001/"
'        .setRequestHeader "Authorization", "Bearer eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiJ6cXpjd2QiLCJjcmVhdGVkIjoxNzIwMjcyNjQxMTU0LCJleHAiOjIxNTIyNzI2NDF9.JH2aV6RKMsIubJ9_-qTCvx47HtmzO81bxuEmKg1e_SNPTX-HUdf5sWwpTcp-YKUdtTkp9H8A24S4y2MI1aaGqw"
        .setRequestHeader "Cookie", "GUEST_LANGUAGE_ID=zh_CN; " & Cookie
        .send "{""tm"":36871,""cc"":null,""khbz"":null}"
'        Debug.Print .responseText
        Set js = JsonConverter.ParseJson(.responseText)
    End With
    
    Dim items As Object, irr, d_gb As New Dictionary, 记录$
    Set items = js("row")("items")
    With Sheet4
        '清空数据区域
'        rs = .Cells(Rows.count, "L").End(3).Row + 1
'        .Range(.Cells(4, "J"), .Cells(rs, "V")).ClearContents
'        .Range(.Cells(4, "J"), .Cells(rs, "V")).Interior.Color = 15921906
        For Each Item In items
            r = .Cells(Rows.count, "L").End(3).Row + 1
            irr = Split(Item("chs"), "/") '车号
    '        For i = 0 To UBound(irr, 1) - 1
    '            记录 = ""
    '            记录 = 违编核对.查询过磅记录(irr(i), dbck)
    '            d_gb(记录) = d_gb(记录) + 1
    '        Next
            .Cells(r, "J").value = Item("cc") '车次
            .Cells(r, "K").value = Item("dzm") '到站名
            .Cells(r, "L").value = Item("gd") '股道
            .Cells(r, "N").value = Item("cs") '车数
            .Cells(r, "O").value = Item("hc") / 10 '换长
            .Cells(r, "Q").value = Item("zz") / 1000 '总重
            .Cells(r, "S").value = d_gb("已过磅") & "/" & d_gb("未过磅"): d_gb.RemoveAll '过磅情况
            .Cells(r, "T").value = VBA.Format(Item("jhrq"), "hh:mm") '计划开点
            .Cells(r, "U").value = VBA.Format(Item("dfrq"), "hh:mm") '实际开点
            
            '违编标记(换长<68.7,载重<4420,小运转不算)
            If (Item("hc") / 10) < 68.8 And (Item("zz") / 1000) < 4420 And Not Left(Item("cc"), 1) = 4 Then
                .Cells(r, "L").Interior.ColorIndex = 6
            End If
    
            
            '读取编组
            url = "http://yt1dzh.crc.cr:30001/qb/GetQbById/" & Item("pk") & "/2"
            With xhtp
                .Open "get", url, False
                .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
                .setRequestHeader "content-type", "application/json"
                .setRequestHeader "Referer", "http://yt1dzh.crc.cr:30001/"
                .setRequestHeader "Cookie", "GUEST_LANGUAGE_ID=zh_CN; " & Cookie
                .send data
                Set bz = JsonConverter.ParseJson(.responseText)
            End With
            Dim d As New Dictionary, dc As New Dictionary, 编组$, 车型$
            Set its = bz("row")("qbzw")
            For Each it In its
                dc(it("czjm")) = dc(it("czjm")) + 1 '车型
                If it("dzm") = "" Then
                    d(it("czjm")) = d(it("czjm")) + 1 '空车
                Else
                    d(it("dzm")) = d(it("dzm")) + 1 '有到站的车
                End If
            Next
            For Each dz In d.keys '编组
            编组 = 编组 & dz & d(dz) & " "
            Next
            .Cells(r, "P").value = 编组
            d.RemoveAll: 编组 = ""
            For Each dx In dc.keys '车型
            车型 = 车型 & dx & dc(dx) & " "
            Next
            .Cells(r, "M").value = 车型
            dc.RemoveAll: 车型 = ""
        Next
    End With
    Application.EnableEvents = True
    Application.ScreenUpdating = True '开启屏幕刷新
    
    Dim rng As Range
    Set rng = Sheet4.Range("P4:P" & r)
    rng.Font.ColorIndex = 1 '字体颜色
    显示信息.标记到站方向颜色 rng '设置颜色
    MsgBox "完成！"
End Sub

Function 查询过磅记录(车号, ck)
Dim HTML As String, url As String, rngc As Range, data As String, 用户名$, 密码$, 日期$
    Dim reg As New RegExp, Cookie$
    reg.Pattern = "\d+"
    
    开始日期 = Format(DateAdd("m", -5, Now), "yyyy-mm-dd")
    结束日期 = Format(DateAdd("d", 5, Now), "yyyy-mm-dd")

    With CreateObject("WinHttp.WinHttpRequest.5.1")
        url = "http://10.190.136.155:8888/Gsgz/GDH/searchGdh"
        data = "start_createtime=" & 开始日期 & "+18%3A00%3A00&end_createtime=" & 结束日期 & "+18%3A00%3A00&motorcycleTtype=&wagonNumber=" & 车号 & "&transceiver_Send=&transceiver_Collect=&loading_Name=&send_Collect=&cargo_Name=&direction=%3C--"
'       data = "start_createtime=2024-04-01+12%3A41%3A17&end_createtime=2024-08-01+12%3A41%3A22&motorcycleTtype=&wagonNumber=1515699&transceiver_Send=&transceiver_Collect=&loading_Name=&send_Collect=&cargo_Name=&direction=%3C--"

        .Open "POST", url, False
        .setRequestHeader "Cookie", ck
        .setRequestHeader "Origin", "http://10.190.136.155:8888"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .setRequestHeader "Referer", "http://10.190.136.155:8888/Gsgz/view/index"
        .send data
'        Debug.Print .responseText
        Set gb = JsonConverter.ParseJson(.responseText)
        If gb.count > 1 Then Debug.Print 车号 & "：" & gb.count
        If gb.count > 0 Then
            For i = gb.count To 1 Step -1
                If gb(i)("direction") = "<--" Then
    '            Debug.Print gb(i)("motorcycleTtype")
    '            Debug.Print gb(i)("rightFour")
    '            Debug.Print gb(i)("sequence")
    '            Debug.Print gb(i)("speed")
    '            Debug.Print gb(i)("roughWeight")
    '            Debug.Print gb(i)("suttle")
'                Debug.Print gb(i)("wagonNumber") & "：" & gb(i)("suttle")
    '            Debug.Print gb(i)("tare")
    '            Debug.Print gb(i)("weightOff")
                记录 = "已过磅"
                Exit For '退出循环
                End If
            Next
        Else
'            Debug.Print 车号 & "：" & "无过磅记录!"
            记录 = "未过磅"
        End If
    End With
    查询过磅记录 = 记录
End Function

Function 加载js(urlPath As String)
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "get", urlPath, False
        .send
'        Debug.Print .responseText
        加载js = .responseText
    End With
End Function

Function 去向导出()
    Dim r%, rng As Range, rn As Range, reg As New RegExp, 车次$
    Dim d_到站 As New Dictionary, d_车次 As New Dictionary, d_方向 As New Dictionary
    reg.Global = True
    reg.Pattern = "[一-龢]+"
    With Sheet4
        r = .Cells(Rows.count, "P").End(3).Row
        Set rng = .Range(.Cells(4, "J"), .Cells(r, "P"))
        For Each rn In rng.Columns(2).Cells
            Debug.Print rn.value
            For Each re In reg.Execute(rn.Offset(0, 5).value)
                d_到站(rn.value) = re.value & " " & d_到站(rn.value)
            Next
            车次 = Mid(rn.Offset(0, -1).value, 1, 3) & "XX"
            If d_方向.Exists(车次) Then
                Debug.Print 车次
            Else
                d_车次(rn.value) = 车次 & Chr(10) & d_车次(rn.value)
                d_方向(车次) = d_方向(车次) + 1
            End If
        Next
'        d_方向(rn.value) = Array(到站, 车次)
    End With
    With Sheet10
        r = .Cells(Rows.count, "Z").End(3).Row + 1
        If r > 80 Then
            Set rng = .Range(.Cells(81, "Z"), .Cells(r, "AC"))
            For Each rn In rng.Columns(1).Cells
                If d_车次.Exists(rn.value) Then
                    For Each cc In Split(d_车次(rn.value), Chr(10))
                        If InStr(rn.Offset(0, 1).value, cc) = 0 And cc <> "" Then
                            rn.Offset(0, 1).value = rn.Offset(0, 1).value & Chr(10) & cc
                        End If
                    Next
                    For Each dz In Split(d_到站(rn.value), " ")
                        If InStr(rn.Offset(0, 2).value, dz) = 0 And dz <> "" And dz <> "钦州港" Then
                            rn.Offset(0, 2).value = rn.Offset(0, 2).value & " " & dz
                        End If
                    Next
                    d_车次.Remove (rn.value)
                End If
            Next
            For Each key In d_车次.keys
                r = .Cells(Rows.count, "Z").End(3).Row + 1
                .Cells(r, "Z").value = key
                .Cells(r, "AA").value = d_车次(key)
                .Cells(r, "AB").value = d_到站(key)
            Next
        End If
        显示信息.标记到站方向颜色 rng.Columns(3).Cells '设置颜色
    End With
End Function


Sub 查询编组()
    Dim xhtp As Object, js As Object, rng As Range, rngs As Range, 车次$
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1")
    With Sheet4 '到发列车
        Cookie = .Range("A1").value
        Set rngs = .Range("B4", .Cells(.Rows.count, "B").End(3))
    End With
    
    Application.EnableEvents = False
    'Application.ScreenUpdating = False
    
    For Each rng In rngs 'B列：车次
    'Debug.Print rng.value
    If Sheet4.Cells(rng.Row, "H").value = "" Then
        车次 = rng.value
        Sheet4.Cells(rng.Row, "H").value = "正在查询..."
        With xhtp
            开始时间 = VBA.Format(DateAdd("h", -12, Now), "yyyy-mm-dd%20hh:mm") '查询时间后推12小时
            结束时间 = VBA.Format(DateAdd("d", 1, Now), "yyyy-mm-dd%2018:00") '次日18点
            
            url = "http://yt1dzh.crc.cr:30001/ytygctzd/queryLcbtQbml?ljdm=10&departmentId=99113&cdm=*&jch=*&startTime=" & 开始时间 & "&endTime=" & 结束时间 & "&pageNum=1&pageSize=100"
            .Open "POST", url, False
            .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
            .setRequestHeader "content-type", "application/json"
            .setRequestHeader "Origin", "http://yt1dzh.crc.cr:30001"
            .setRequestHeader "Referer", "http://yt1dzh.crc.cr:30001/"
            '.setRequestHeader "Cookie", cookie
            .send "{""tm"":36871,""cc"":" & 车次 & ",""khbz"":null}" '久
            '.send "{""tm"":null,""cc"":null,""khbz"":null}" '久
            'Debug.Print .responseText
            Set js = JsonConverter.ParseJson(.responseText)
        End With
        
        Dim items As Object, irr, d_gb As New Dictionary, 记录$
        Set items = js("row")("list")
        If items.count > 0 Then
            With Sheet4
                rs = .Cells(Rows.count, "B").End(3).Row
                '.Range(.Cells(4, "B"), .Cells(rs, "H")).ClearContents
                .Range(.Cells(4, "B"), .Cells(rs, "H")).Interior.Color = 15921906
                'For Each Item In items
                Set Item = items(1) '拿最新的编组
                    'Debug.Print Item("cc")
                    r = rng.Row
                    'irr = Split(Item("chs"), "/") '车号
                    '.Cells(r, "B").value = Item("cc") '车次
                    '.Cells(r, "C").value = VBA.Format(Item("dfrq"), "m-d hh:mm") '发报时间
                    With .Cells(r, "C")
                        .value = ""
                        .NumberFormatLocal = "hh:mm;@"
                    End With
                    .Cells(r, "D").value = Item("dzm") '终到站
                    .Cells(r, "F").value = Item("cs") '车数
                    .Cells(r, "G").value = Item("hc") / 10 '换长
                    '在批注中存入列车识别码
                    With .Cells(r, "H")
                        If Not .Comment Is Nothing Then .Comment.Delete
                        .AddComment Item("cc") & ":" & Item("pk")
                    End With
        
                    '读取编组
                    url = "http://yt1dzh.crc.cr:30001/qb/GetQbById/" & Item("pk") & "/20"
                    With xhtp
                        .Open "get", url, False
                        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
                        .setRequestHeader "content-type", "application/json"
                        .setRequestHeader "Referer", "http://yt1dzh.crc.cr:30001/"
                        '.setRequestHeader "Cookie", cookie
                        .send data
                        Set bz = JsonConverter.ParseJson(.responseText)
                    End With
                    Dim d As New Dictionary, dc As New Dictionary, 编组$, 车型$
                    '品名:pm，到站dzm，车型:czjm
                    Set its = bz("row")("qbzw")
                    For Each it In its
                        DoEvents
                        dc(it("czjm")) = dc(it("czjm")) + 1 '车型
                        If it("dzm") = "钦州港" Then
                            i = i + 1
                            If it("pm") Like "*[纯碱,尿素,小麦粉,棉粕,工业用盐,橡胶]*" Then '品名
                                d("货场") = d("货场") + 1 '货场重车
                            ElseIf it("chjsl") Like "*永鑫*" Then
                                d("永鑫") = d("永鑫") + 1 '记事栏:永鑫
                            ElseIf it("chjsl") Like "*[天盛,天益]*" Then
                                d("天盛") = d("天盛") + 1 '记事栏:天盛
                            ElseIf it("chjsl") Like "*货场*" Then
                                d("货场") = d("货场") + 1 '记事栏:货场
                            'ElseIf it("pm") Like "*[敞二空]*" Then
                                'd("漫游箱") = d("漫游箱") + 1 '记事栏:漫游箱
                            Else
                                If it("pm") = "" Then
                                    d(it("czjm")) = d(it("czjm")) + 1 '车型栏
                                Else
                                    'If it("pm") Like "*[自二重2,敞二重2]*" Then '品名
                                        'd("货场") = d("货场") + 1 '货场重车
                                    'ElseIf it("pm") = "敞二空2" Then
                                        'd("漫游箱") = d("漫游箱") + 1 '记事栏:漫游箱
                                    'Else
                                        d(it("pm")) = d(it("pm")) + 1 '品名栏
                                    'End If
                                End If
                            End If
                        ElseIf it("dzm") = "" Then
                            d(it("czjm")) = d(it("czjm")) + 1 '车型栏
                        Else
                            d(it("dzm")) = d(it("dzm")) + 1 '到站
                        End If
                    Next
                    For Each dz In d.keys '编组
                    编组 = 编组 & dz & ":" & d(dz) & " "
                    Next
                    .Cells(r, "H").value = 编组
                    d.RemoveAll: 编组 = ""
                    
                    For Each dx In dc.keys '车型
                    车型 = 车型 & dx & dc(dx) & " "
                    Next
                    .Cells(r, "E").value = 车型
                    dc.RemoveAll: 车型 = ""
                'Next
            End With
        Else
            Sheet4.Cells(rng.Row, "C").value = "请核对车次！"
            'Application.EnableEvents = True
        End If
    End If
    Next
    
    Application.EnableEvents = True
    'Application.ScreenUpdating = True '开启屏幕刷新

End Sub

Function 获取编组(pk$) As Boolean
    Dim xhtp As Object, Cookie$
    Dim d As New Dictionary, dc As New Dictionary, 编组$, 车型$
    Dim oDom As Object, ow As Object, res As String, rng As Range
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    url = "http://yt1dzh.crc.cr:30001/qb/GetQbById/" & pk & "/20"
    
    On Error GoTo ErrorHandler
    With Sheet30
        Cookie = Sheet4.Range("A1").value
        .Range("A3:O3").ClearContents '清空数据区域
        Set rng = .Range("A5").CurrentRegion.Offset(1)
        If rng.Rows.count > 2 Then
            With rng
                .ClearContents '清空数据区域
                .Interior.Color = 16777215 '16777215
                .Font.Bold = False
                .Font.Color = 0
            End With
        End If
        With xhtp
            .Open "get", url, False
            .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
            .setRequestHeader "content-type", "application/json"
            .setRequestHeader "Referer", "http://yt1dzh.crc.cr:30001/"
            .send
            res = .responseText
        End With
        
        
        ow.execScript "var jstr =" & res & ";"
        Set items = ow.eval("jstr.row.qbzw")
        On Error GoTo 0
        
        Application.ScreenUpdating = False
        On Error Resume Next
        For Each Item In items
            DoEvents
            r = .Cells(Rows.count, "C").End(3).Row + 1
            .Cells(r, "B").value = r - 5        '序号
            .Cells(r, "C").value = Format(Item.ch, "0000000") '车号
            .Cells(r, "D").value = Item.cz '车种
            .Cells(r, "E").value = Item.ziz / 1000 '自重
            .Cells(r, "F").value = Item.hc '换长
            .Cells(r, "G").value = Item.pjzaiz '载重
            .Cells(r, "H").value = Item.dzm '到站
            .Cells(r, "I").value = Item.pjysfs '方向
            .Cells(r, "J").value = Item.pm '品名
            .Cells(r, "K").value = Item.fzm '发站
            '.Cells(r, "L").value = Item.pjfzm '篷
            .Cells(r, "M").value = Item.pjshr '收货人
            .Cells(r, "N").value = Item.jsl '记事
            .Cells(r, "O").value = Item.chjsl '车号记事
        Next
        
        Set Item = ow.eval("jstr.row.qbjb[0]")
        .Cells(3, "D").value = Item.initTime '发车时间
        .Cells(3, "G").value = Item.item04 '局
        .Cells(3, "H").value = Item.item05 '重车
        .Cells(3, "I").value = Item.item06 '空车
        .Cells(3, "J").value = Item.item09 '其它
        .Cells(3, "K").value = Item.item10 '合计
        .Cells(3, "L").value = Item.item14 '自重
        .Cells(3, "M").value = Item.item02 '载重
        .Cells(3, "N").value = Item.item01 '总重
        .Cells(3, "O").value = Item.item11 / 10 '换长
        
        cc = ow.eval("jstr.row.qbml.cc")
        .Cells(3, "B").value = cc '车次
        
    End With
    
    On Error GoTo 0
    Application.ScreenUpdating = True '开启屏幕刷新
    
    '标记车型
    xrr = Array(Array("G", 12632256, 3, 0, 0), Array("X", 12632256, 4, 0, 0), Array("P", 49407, 4, 0, 0), Array("YW", 16777215, 4, 1, 255)) '车型，底色，位置，加粗，文字颜色
    For Each ar In xrr
        车型标记 ar(0), ar(1), ar(2), ar(3), ar(4)
    Next
    
    获取编组 = True
    
    Exit Function
    
ErrorHandler:

    Application.ScreenUpdating = True '开启屏幕刷新
    On Error GoTo 0
    MsgBox "查询失败！"
    获取编组 = False
    
End Function

Function 车型标记(cx, col, c, ft, fc)
    Dim reg As New RegExp, rng As Range, crr
    reg.Pattern = "\w+" '"[^\d]+"
    Set rng = Sheet30.Range("A5").CurrentRegion.Offset(1)
    crr = rng
    
    For i = 1 To UBound(crr) - 1
        Set t = reg.Execute(crr(i, 4))
        If reg.Test(crr(i, 4)) Then
            If cx = "G" And Left(crr(i, 3), 1) = 6 Then
                'Debug.Print "LG"
            ElseIf InStr(t(0), cx) > 0 Then
                With Sheet30.Cells(i + 5, c)
                    .Interior.Color = col '平板车(X/NX)标色-灰(15)/紫(14336204)
                    .Font.Bold = ft '字体加粗
                    .Font.Color = fc '字体颜色
                End With
            End If
        End If
    Next
End Function


Sub 登录系统()
    Dim oDom As Object, ow As Object, result As String, js As Object
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    
    Sheet4.Range("A1").value = "钦州港站调度-综合应用管理系统" & "  时间：" & Now()
    'End '结束
    
    Dim pathAes As String, AES As String
    pathAes = "http://10.190.168.62:80/nnjzj/resources/script/common/aes.js" '网络读取
    AES = 加载js(pathAes)
    ow.execScript AES
    帐号 = ow.encrypt("qvz", "RailsYdjszz46495")
    密码 = ow.encrypt("123qwe!@#", "RailsYdjszz46495")
    
    Application.EnableEvents = False
    'Application.ScreenUpdating = False '关闭屏幕刷新
    
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1")
    url = "http://yt1dzh.crc.cr:30001/authentication/login"
    data = "{""username"":""" & 帐号 & """,""password"":""" & 密码 & """}"
'    Debug.Print data
    With xhtp
        .Open "POST", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "Origin", "http://yt1dzh.crc.cr:30001"
        .setRequestHeader "Referer", "http://yt1dzh.crc.cr:30001/"
        .setRequestHeader "Cookie", "GUEST_LANGUAGE_ID=zh_CN"
        .send data
'        Debug.Print .responseText
'        res = .responseText
'
'        ow.execScript "var res=" & res
'        Set data = ow.eval("res")
'        Set ro = CallByName(data, "row", VbGet)
        
        '动态的参数
        Dim Cookie$, 开始时间$, 结束时间$
        Cookie = Split(.getResponseHeader("Set-Cookie"), ";")(0)
        Sheet4.Range("A1").value = Cookie
    End With
    Application.EnableEvents = True
End Sub


'async function 查询(){
'
'    var res=await fetch("http://yt1dzh.crc.cr:30001/qb/getqbmlforallByPage?qbType=2&begin=2025-09-21%2018:01&end=2025-09-22%2018:00&pageNum=1&pageSize=100&cdm=*&ljdm=10&departmentId=99113", {
'      "headers": {
'
'        "content-type": "application/json",
'        //"Referer": "http://yt1dzh.crc.cr:30001/",
'        //"Referrer-Policy": "strict-origin-when-cross-origin"
'      },
'      "body": `{"tm":null,"cc":null,"khbz":null}`,
'      "method": "POST"
'    });
'
'    var t=await res.json();
'    console.Log (Json.stringify(t.Row.items))
'
'}
