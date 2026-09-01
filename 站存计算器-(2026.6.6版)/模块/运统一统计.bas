Attribute VB_Name = "运统一统计"
Function 发运统一统计()

    Dim xhtp As Object, js As Object, rng As Range, rngs As Range, res As String, d_Qs As Scripting.Dictionary, d_Q As Scripting.Dictionary
    Dim oDom As Object, ow As Object, n As Long, 时间 As Long
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    
    时间 = 0 - val(Sheet32.Range("T3").value)
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1") 'CreateObject("MSXML2.XMLHTTP.6.0") '
    Set d_Q = CreateObject("Scripting.Dictionary") '初始化字典
    Set d_Qs = CreateObject("Scripting.Dictionary") '初始化字典
    With xhtp
'        开始时间 = VBA.Format(DateAdd("h", 时间, Now), "yyyy-mm-dd%20hh:mm") '查询时间后推12小时
'        结束时间 = VBA.Format(DateAdd("d", 1, Now), "yyyy-mm-dd%2018:00") '次日18点
        开始时间 = "2025-03-01%2018:01:00"
        结束时间 = "2025-03-31%2018:01:00"
        
        'DoEvents
        With Sheet32
'            rs = .Cells(Rows.count, "R").End(3).Row + 1
'            With .Range(.Cells(6, "Q"), .Cells(rs, "X"))
'                .ClearContents
'                .ClearComments '清除
'                '.Interior.Color = 16777215 '15921906
'            End With

            '钦港出发
            url = "http://yt1dzh.crc.cr:30001/qb/getqbmlforallByPage?qbType=2&begin=" & 开始时间 & "&end=" & 结束时间 & "&pageNum=1&cdm=*&ljdm=10&departmentId=99113&pageSize=1000"
            data = "{""tm"":36871,""cc"":null,""khbz"":null}"
            With xhtp
                .Open "post", url, False
                .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
                .setRequestHeader "content-type", "application/json"
                .setRequestHeader "Origin", "http://yt1dzh.crc.cr:30001"
                .setRequestHeader "Referer", "http://yt1dzh.crc.cr:30001/"
                .send data '久
                
                res = .responseText
            End With
            
            Dim stu As String, page As String
            ow.execScript "var js =" & res & ";"
            'Debug.Print res
            stu = ow.eval("js.success")
            If stu Then
                lics = ow.eval("js.row.items.length")
                Debug.Print "出发列数：" & lics
            Else
                Debug.Print "出错啦！：" & stu
            End If
            
        End With
            
    End With
End Function

Function 发送运统一请求(url, gt, data, tyb)
    Dim xhtp As Object
    Set xhtp = CreateObject("MSXML2.XMLHTTP.6.0") 'CreateObject("WinHttp.WinHttpRequest.5.1")
   
    With xhtp
        .Open gt, url, tyb
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "Origin", "http://yt1dzh.crc.cr:30001"
        .setRequestHeader "Referer", "http://yt1dzh.crc.cr:30001/"
        .send data '久
    End With
    
    Set 发送运统一请求 = xhtp
End Function

'比较久,需要Cookie
Function 钦州港站出发情况()
    Dim xhtp As Object
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1") 'CreateObject("MSXML2.XMLHTTP.6.0") 'CreateObject("WinHttp.WinHttpRequest.5.1")

    开始时间 = "2026-01-04+18:00:00"
    结束时间 = "2026-01-04+18:00:00"
        
    UA = "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
    Ct = "application/x-www-form-urlencoded;charset=UTF-8"
    Og = "http://10.190.5.143:8081"
    'ck = "fineMarkId=ccc8a1faed99c617c91f4180c2c2a633"
    Rf = "http://10.190.5.143:8081/webroot/decision/view/form?viewlet=cw/dd_cf.frm&op=view&dwName=钦州车务段&RoleID=2"
    'sid = "4285dcf4-6fb2-42fc-995d-a73a3447de84"
    sid = 登录
    
    With xhtp
        url = "http://10.190.5.143:8081/webroot/decision/view/form?op=fr_dialog&cmd=parameters_d"
        'data = "__parameters__={""DW"":""[8f66][52a1][6bb5][ff1a]"",""DWNAME"":""[94a6][5dde][8f66][52a1][6bb5]"",""LABEL"":""[8f66][7ad9][ff1a]"",""ZM"":"""",""LABEL3_C_"":""[5230][53d1][65e5][671f][ff1a]"",""START"":""" & 开始时间 & """,""LABEL2_C_C_C_C_C_C_C"":""-"",""END"":""" & 结束时间 & """,""DZ"":""'QVZ'""}"
        data0 = "__parameters__={""DW"":""[8f66][52a1][6bb5][ff1a]"",""DWNAME"":""[94a6][5dde][8f66][52a1][6bb5]"",""LABEL"":""[8f66][7ad9][ff1a]"",""ZM"":"""",""LABEL3_C_"":""[5230][53d1][65e5][671f][ff1a]"",""START"":""" & 开始时间 & """,""LABEL2_C_C_C_C_C_C_C"":""-"",""END"":""" & 结束时间 & """,""CZ"":"""",""DZ"":""'QVZ'"",""STATUS"":""'F'""}&_=" & 1770238997098#
        .Open "POST", url, False
        .setRequestHeader "User-Agent", UA
        .setRequestHeader "Content-Type", Ct
        .setRequestHeader "Origin", Og
        .setRequestHeader "Referer", Rf
        .setRequestHeader "sessionID", sid
        .send data0
        'res = .responseText
        'Debug.Print res
        '保存.HTML res
        
        url = "http://10.190.5.143:8081/webroot/decision/view/form?op=fr_form&cmd=load_content&widgetVersion=1&_=" & 1770245281239#
        .Open "GET", url, False
        .setRequestHeader "User-Agent", UA
        .setRequestHeader "Referer", Rf
        .setRequestHeader "sessionID", sid
        .send
        'res = .responseText
        'Debug.Print res
        '保存.HTML res
        
        url = "http://10.190.5.143:8081/webroot/decision/view/form"
        'data = "op=fr_form&cmd=load_report_content&widgetName=REPORT1&pageIndex=1" & data0
        data = "op=fr_form&cmd=load_report_content&widgetName=REPORT1&__parameters__={""LABEL3_C_"":""[5230][53d1][65e5][671f][ff1a]"",""FINEMARKID"":""ccc8a1faed99c617c91f4180c2c2a633"",""GZIPPED"":""true"",""HEADERADDED"":""true"",""DW"":""[8f66][52a1][6bb5][ff1a]"",""LABEL2_C_C_C_C_C_C_C"":""-"",""DZ"":""'QVZ'"",""WIDGETVERSION"":""1"",""LABEL"":""[8f66][7ad9][ff1a]"",""CMD"":""load_content"",""_"":""1770245281239"",""DWNAME"":""[94a6][5dde][8f66][52a1][6bb5]"",""COUNTER.FILTERED"":""true"",""FORMLETNAME"":""cw/dd_cf.frm"",""__WIDGETVALUE__"":""[5b]\\""CZ\\""[5d]"",""START"":""" & 开始时间 & """,""END"":""" & 结束时间 & """,""ROLEID"":""2"",""ZM"":""QVZ"",""CZ"":""[94a6][5dde][6e2f]"",""STATUS"":""'F'"",""animateType"":""none""}&noCache=lazy&pageIndex=1&_=1770247495593&__boxModel__=true&reload=null&_PAPERWIDTH=1014&_PAPERHEIGHT=553&_SHOWPARA=true&_SHOWPARATEMPLATE=false"
        .Open "POST", url, False
        .setRequestHeader "User-Agent", UA
        .setRequestHeader "Content-Type", Ct
        .setRequestHeader "Origin", Og
        .setRequestHeader "Referer", Rf
        .setRequestHeader "sessionID", sid
        .send data
        res = .responseText
        'Debug.Print res
        保存.HTML res
    End With

End Function

Function 登录()
    Dim xhtp As Object, reg As New RegExp
    With reg
        .Pattern = """(\w{8}-\w{4}-\w{4}-\w{4}-\w{12})""" '"FR.SessionMgr.register('89c5c3c1-e710-47cc-9ae3-b8306d71ef0b', contentPane);"
        .Global = False
        .MultiLine = False
    End With
    
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1")
    With xhtp
        url = "http://10.190.5.143:8081/webroot/decision/view/form?viewlet=cw/dd_cf.frm&op=view&dwName=%E9%92%A6%E5%B7%9E%E8%BD%A6%E5%8A%A1%E6%AE%B5&RoleID=2"
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "Cookie", "fineMarkId=ccc8a1faed99c617c91f4180c2c2a633"
        .send
        res = .responseText
'        保存.HTML res
'        Set mah = reg.Execute(res)
        sid = reg.Execute(res)(0).SubMatches(0)
    End With
    
    登录 = sid
End Function
