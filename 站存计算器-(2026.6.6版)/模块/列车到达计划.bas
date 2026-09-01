Attribute VB_Name = "列车到达计划"
Dim dr As New Dictionary
Function 查询转发运统一(Optional ByVal 车次 As String = "null")
    Dim xhtp As Object, js As Object, rng As Range, rngs As Range, res As String, d_Qs As Scripting.Dictionary, d_Q As Scripting.Dictionary
    Dim oDom As Object, ow As Object, n As Long, 时间 As Long
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    
    时间 = 0 - val(Sheet30.Range("T3").value)
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1") 'CreateObject("MSXML2.XMLHTTP.6.0") '
    Set d_Q = CreateObject("Scripting.Dictionary") '初始化字典
    Set d_Qs = CreateObject("Scripting.Dictionary") '初始化字典
    With xhtp
        开始时间 = VBA.Format(DateAdd("h", 时间, Now), "yyyy-mm-dd%20hh:mm") '查询时间后推12小时
        结束时间 = VBA.Format(DateAdd("d", 1, Now), "yyyy-mm-dd%2018:00") '次日18点
        
        DoEvents
        
        If 车次 = "null" Then
        
            With Sheet30
                rs = .Cells(Rows.count, "R").End(3).Row + 1
                With .Range(.Cells(6, "Q"), .Cells(rs, "X"))
                    .ClearContents
                    .ClearComments '清除
                    '.Interior.Color = 16777215 '15921906
                End With
            End With

            '黎塘转发
            url = "http://yt1dzh.crc.cr:30001/qb/getqbmlforallByPage?qbType=2&begin=" & 开始时间 & "&end=" & 结束时间 & "&pageNum=1&pageSize=100&cdm=*&ljdm=10&departmentId=14178"
            Set d_Qs("黎塘") = 发送运统一请求(url, "post", "{""tm"":36374,""cc"":null,""khbz"":null}", True) '黎塘转发
            
            '马皇出发
            url = "http://yt1dzh.crc.cr:30001/qb/getqbmlforallByPage?qbType=2&begin=" & 开始时间 & "&end=" & 结束时间 & "&pageNum=1&pageSize=100&cdm=*&ljdm=10&departmentId=99113"
            Set d_Qs("吴圩") = 发送运统一请求(url, "post", "{""tm"":36818,""cc"":null,""khbz"":null}", True) '吴圩转发
            Set d_Qs("南口") = 发送运统一请求(url, "post", "{""tm"":36815,""cc"":null,""khbz"":null}", True) '那罗转发
            Set d_Qs("六景东") = 发送运统一请求(url, "post", "{""tm"":36890,""cc"":null,""khbz"":null}", True) '六景东转发
            Set d_Qs("马皇1") = 发送运统一请求(url, "post", "{""tm"":36851,""cc"":null,""khbz"":null}", True) '马皇转发写入表格 res, "马皇"
            Set d_Qs("港东") = 发送运统一请求(url, "post", "{""tm"":36803,""cc"":null,""khbz"":null}", True) '钦州港东转发写入表格 res, "港东", , 1

            '马皇转发
            url = "http://yt1dzh.crc.cr:30001/ytygctzd/queryLcbtQbml?ljdm=10&departmentId=99113&cdm=*&jch=*&startTime=" & 开始时间 & "&endTime=" & 结束时间 & "&pageNum=1&pageSize=100"
            Set d_Qs("马皇2") = 发送运统一请求(url, "post", "{""tm"":36851,""cc"":null,""khbz"":null}", True) '马皇转发写入表格 res, "马皇"

            开始时间 = VBA.Format(DateAdd("h", 时间 - 5, Now), "yyyy-mm-dd%20hh:mm") '查询时间后推12小时
            结束时间 = VBA.Format(DateAdd("d", 1, Now), "yyyy-mm-dd%2018:00") '次日18点
            '南宁南出发
            url = "http://yt1dzh.crc.cr:30001/qb/getqbmlforallByPage?qbType=2&begin=" & 开始时间 & "&end=" & 结束时间 & "&pageNum=1&pageSize=100&cdm=*&ljdm=10&departmentId=14178"
            Set d_Qs("南宁南") = 发送运统一请求(url, "post", "{""tm"":36449,""cc"":null,""khbz"":null}", True) '南宁南转发

            '柳州南出发
            url = "http://yt1dzh.crc.cr:30001/qb/getqbmlforallByPage?qbType=2&begin=" & 开始时间 & "&end=" & 结束时间 & "&pageNum=1&pageSize=100&cdm=*&ljdm=10&departmentId=89509"
            Set d_Qs("柳州南") = 发送运统一请求(url, "post", "{""tm"":36266,""cc"":null,""khbz"":""H""}", True) '柳州南出发
            
            t = Now
            Do '异步查询1

                For Each dd In d_Qs.keys
                    DoEvents
'                    On Error Resume Next
'                    Set xhtp = d_Qs(dd)
'                    res = xhtp.responseText
'                    On Error GoTo 0
                    If d_Qs(dd).readyState = 4 Then
                        res = d_Qs(dd).responseText
                        If res = "" Then
                           Debug.Print "无返回数据！！！"
                        Else
                            ow.execScript "var js =" & res & ";"
                            'Debug.Print res
                            suc = ow.eval("js.success") '数据获取是否成功
                            If suc Then
                                ow.execScript "function sc(e){return e.row.items || e.row.list || []};sc(js)"
                                Set list = ow.eval("sc(js)")
                                For Each Item In list
                                    DoEvents
                                    
                                    '筛选条件
                                    Dim 到站 As Boolean
                                    车站 = Item.dzm
                                    转发站 = dd
                                    发站 = Item.fzm
                                    到站 = InStr("钦州港防城港马皇黎塘", 车站)
                                    
                                    If 到站 Then
                                        'Debug.Print 车站
                                        url = "http://yt1dzh.crc.cr:30001/qb/GetQbById/" & Item.pk & "/2"
                                        Set d_Q(发送运统一请求(url, "get", "", True)) = Array(发站, Item) '获取编组
                                    End If
                                    
                                    Set Item = Nothing
                                    Set list = Nothing
                                Next
                                d_Qs.Remove dd
                            Else
                                Debug.Print "1、请求失败：" & suc
                            End If
                        End If
                    End If
                    res = ""
                    
                Next dd
                
                DoEvents '交回控制权给系统，防卡死！
                jt = Format(Now - t, "s")
                 If jt > 10 Then
                    Debug.Print "1池:" & jt
                    Exit Do
                 End If
            Loop Until d_Qs.count = 0
            
            t = Now
            Do '异步查询2

                For Each dd In d_Q.keys
                    DoEvents
                    If dd.readyState = 4 Then
                        res = dd.responseText
                        If res = "" Then
                            Debug.Print "res等于空！" & d_Q(dd)(0)
                        Else
                            ow.execScript "var js =" & res & ";"
                            suc = ow.eval("js.success") '数据获取是否成功
                            If suc Then
                                车站 = d_Q(dd)(0)
                                Set obj = d_Q(dd)(1)
                                获取编组详情 res, 车站, obj
                                d_Q.Remove dd
                            Else
                                Debug.Print "2、请求失败：" & suc
                            End If
                        End If
                    End If
                    res = ""
                Next dd
                
                DoEvents '交回控制权给系统，防卡死！
                jt = Format(Now - t, "s")
                 If jt > 10 Then
                    Debug.Print "2池:" & jt
                    Exit Do
                 End If
            Loop Until d_Q.count = 0
        Else
'            按车次查询
            url = "http://yt1dzh.crc.cr:30001/ytygctzd/queryLcbtQbml?ljdm=10&departmentId=99113&cdm=*&jch=*&startTime=" & 开始时间 & "&endTime=" & 结束时间 & "&pageNum=1&pageSize=100"
            .Open "POST", url, False
            .setRequestHeader "content-type", "application/json"
            .setRequestHeader "Origin", "http://yt1dzh.crc.cr:30001"
            .setRequestHeader "Referer", "http://yt1dzh.crc.cr:30001/"
            .send "{""tm"":null,""cc"":" & 车次 & ",""khbz"":null}" '久
            res = .responseText

            '写入表格 res, 0, "list"
            ow.execScript "var js =" & res & ";"
            'Set list = ow.eval("js.row.list")
            l = ow.eval("js.row.list.length")
            If l = 0 Then Exit Function
            
            Set Item = ow.eval("js.row.list[""0""]")
            url = "http://yt1dzh.crc.cr:30001/qb/GetQbById/" & Item.pk & "/20"
            Set xhtp = 发送运统一请求(url, "get", "", False)
            res = xhtp.responseText

            ow.execScript "var js =" & res & ";"
            suc = ow.eval("js.success") '数据获取是否成功
            If suc Then
                查询转发运统一 = 获取编组详情(res, "查询", Item)
            Else
                Debug.Print "3、请求失败：" & suc
            End If
        End If

    End With
    
    '排序
    With Sheet30
        Application.ScreenUpdating = False
        If .Range("R5").CurrentRegion.Rows.count > 2 Then
            升序 Sheet30, .Range("Q5:X5"), .Range("R5"), 1
            .Range("Q5:X5").AutoFilter
        Else
            MsgBox "查询完成！" & Chr(10) & "无到达车流..."
        End If
        Application.ScreenUpdating = True
    End With
    
ftc:
    dr.RemoveAll '清空字典
    
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

Function 获取编组详情(res, ic, Item)
    Dim oDom As Object, ow As Object, its As Object
    Dim d As New Dictionary, dc As New Dictionary, 编组$, 车型$, 到站 As Boolean, 记录1 As Boolean, 记录2 As Boolean, 空到站 As Long
    
    Set oDom = CreateObject("htmlfile"): Set ow = oDom.parentWindow
    '读取编组
    With Sheet30
        On Error Resume Next
        Set bz = JsonConverter.ParseJson(res)
        Set its = bz("row")("qbzw")
        If Not its Is Nothing Then
            For Each it In its
                DoEvents
                dc(it("czjm")) = dc(it("czjm")) + 1 '车型
                If it("dzm") = "钦州港" Then
                    记录1 = True
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
                    空到站 = 空到站 + 1 '记录空到站
                    d(it("czjm")) = d(it("czjm")) + 1 '车型栏
                Else
                    d(it("dzm")) = d(it("dzm")) + 1 '到站
                End If
            Next
            
            '筛选钦港重车和空车
            If 空到站 = WorksheetFunction.Sum(dc.items) Then '空车
                记录2 = True
            Else
                'Debug.Print Item.dzm & ":" & Item.cc
                记录2 = False
            End If
            
            '筛选钦港方向
            If Item.dzm = "黎塘" And (val(Item.cc) Mod 2) = 0 Then
                'Debug.Print Item.dzm & ":" & Item.cc
                记录3 = False
            Else
                记录3 = True
            End If
    
            '记录2 = jsli = "list" Or d.count = dc.count
            If (记录1 Or 记录2) And 记录3 Then
            
                Dim 车次 As String
                车次 = str(Item.cc) ' 车次
                
                If Not dr.Exists(车次) Then
                
                    dr(车次) = ic '记录车次,决定是否显示,去重
    
                    If IsNumeric(ic) Then
                        ix = ix + 1 '序号
                    Else
                        ix = ic '转发站名
                    End If
                    
                    r = .Cells(Rows.count, "R").End(3).Row + 1 '新记录
                
                    .Cells(r, "Q").value = ix '序号
                    'irr = Split(Item("chs"), "/") '车号
                    .Cells(r, "R").value = Item.cc '车次
                    .Cells(r, "S").value = VBA.Format(Item.dfrq, "hh:mm") '出发时间
                    .Cells(r, "T").value = Item.dzm '终到站
                    .Cells(r, "V").value = Item.cs '车数
                    .Cells(r, "W").value = Item.hc / 10 '换长
                    
                    '在批注中存入列车识别码
                    With .Cells(r, "X")
                        If Not .Comment Is Nothing Then .Comment.Delete
                        .AddComment Item.cc & ":" & Item.pk
                        pk = Item.pk '列车识别码
                    End With
                    
                    For Each dz In d.keys '编组
                        DoEvents
                        编组 = 编组 & dz & ":" & d(dz) & " "
                    Next
                    .Cells(r, "X").value = 编组
                    d.RemoveAll: 编组 = ""
                    
                    For Each dx In dc.keys '车型
                        DoEvents
                        车型 = 车型 & dx & dc(dx) & " "
                    Next
                    .Cells(r, "U").value = 车型
                    dc.RemoveAll: 车型 = ""
                    
                    '开关
                    记录1 = False
                    记录2 = False
                    记录3 = False
                End If
            End If
        End If
        On Error GoTo 0
    End With
    
    获取编组详情 = pk
End Function

Function 登录()
    On Error GoTo ErrorHandler
    Dim oDom As Object, ow As Object, pw As String, uid$
    Set oDom = CreateObject("htmlfile")
    Set ow = oDom.parentWindow
    ow.execScript 读取文件("H:\Momda\VBA_大文本缓存\到达计划JS\RSA加密.txt")
    ow.execScript 读取文件("H:\Momda\VBA_大文本缓存\到达计划JS\getuid函数.txt")
    
    '帐号密码
    Dim x As String, tk As String, data As String, d As Object
    usid = "36871145"
    pw = ow.RSA_Public_Encrypt("Qvz123456#")
    uid = ow.getuid("")
    Set yzm = 获取验证码(uid)
    'y = CallByName(yzm, "Y", VbGet)
    Image = yzm.oriCopyImage
    tk = yzm.token
    x = 缺口图片验证码识别.PNG模式(Image)
    
    data = "{""userName"":""" & usid & """,""password"":""" & pw & """,""secType"":""RSA"",""type"":""inner"",""x"":" & x & ",""token"":""" & tk & """,""extData"":{""grant_type"":""password""},""agent"":""Chrome 108"",""deviceId"":""Windows64-6.1"",""deviceName"":""Chrome::0f5d844e80cda20b8bc6b88f9fb37859"",""osName"":""Windows64"",""osVersion"":""Windows64-6.1""}"
    
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1") 'CreateObject("MSXML2.XMLHTTP.6.0") '
    With xhtp

        url = "http://10.4.10.11/api/zuul/login"
        .Open "POST", url, False
        .setRequestHeader "Accept", "application/json, text/plain, */*"
        .setRequestHeader "Accept-Language", "zh-CN,zh;q=0.9"
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "channel", "P"
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "pragma", "no-cache"
        .setRequestHeader "rTrackId", uid
        .setRequestHeader "Origin", "http://10.4.10.11"
        .setRequestHeader "Referer", "http://10.4.10.11/login"
        .setRequestHeader "Referrer-Policy", "strict-origin-when-cross-origin"
        .setRequestHeader "type", "inner" '20260606新增
        .send data

        res = .responseText
        ow.execScript "var js =" & res & ";"
        
        Dim list As Object, l As Long
        If ow.js.msg = "OK" Then
            accessToken = ow.eval("js.data.accessToken")
        Else
            Debug.Print ow.js.msg
            GoTo ErrorHandler
        End If
    End With
    
    '存入程序级变量中
    Dim 累计 As Long: 累计 = 0
    累计 = val(GetSetting("Momda", "login", "计数")) + 1 '获取计数
    SaveSetting "Momda", "login", "accessToken", accessToken
    SaveSetting "Momda", "login", "计数", 累计
    
    登录 = True
    Debug.Print 登录
    
    Sheet30.Range("F3").value = 累计

    Exit Function
    
ErrorHandler:
    累计 = val(GetSetting("Momda", "login", "失败")) + 1 '获取计数
    SaveSetting "Momda", "login", "失败", 累计
    Sheet30.Range("G3").value = 累计
    
    登录 = False
    Debug.Print 登录
End Function

Function 获取验证码(uid As String)
    Dim xhtp As Object, oDom As Object, ow As Object
    Set oDom = CreateObject("htmlfile")
    Set ow = oDom.parentWindow
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1")
    With xhtp
        url = "http://10.4.10.11/api/yhzx/slug/getSliderImg"
        .Open "POST", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "accept", "application/json, text/plain, */*"
        .setRequestHeader "cache-control", "no-cache"
        .setRequestHeader "channel", "P"
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "rtrackid", uid
        .setRequestHeader "type", "inner"
        .setRequestHeader "Referer", "http://10.4.10.11/login"
        .setRequestHeader "Referrer-Policy", "strict-origin-when-cross-origin"
        .send "{}"

        res = .responseText
        ow.execScript "var jstr =" & res & ";"
        Dim data As Object, l As Long
        Set data = ow.eval("jstr.data")
    End With
    
    Set 获取验证码 = data
End Function

'设置程序级变量
Sub SetName(name As String, v)
    If VBA.IsNumeric(v) Then
        Application.ExecuteExcel4Macro "SET.NAME（""""" & name & """""," & v & ")"
    ElseIf VBA.TypeName(v) = "String" Then
        Application.ExecuteExcel4Macro "SET.NAME（""""" & name & """"",""" & v & """)"
    End If
End Sub

'读取程序级变量
Function GetName(name As String)
    On Error GoTo ErrorHandler
     GetName = Application.ExecuteExcel4Macro(name)
    If VBA.IsError(GetName) Then GetName = ""
ErrorHandler:
    GetName = ""
End Function

Sub 变更查询()
    
    Dim xhtp As Object, ds As Scripting.Dictionary, oDom As Object, ow As Object, token As String, data As String, s As String, res As String
    Set oDom = CreateObject("htmlfile")
    Set ow = oDom.parentWindow
    
    Dim http() As Object, 变更标志 As Boolean
    Dim rngs As Range, rng As Range, rn As Range
    Set rngs = Sheet30.Range("A5").CurrentRegion
    Set rng = rngs.Columns(3)
    
    If rng.Rows.count = 1 Then Exit Sub
    
    ReDim http(1 To rng.Rows.count - 1)
    rngs.ClearComments '清除批注
    
    变更标志 = False
    With Sheet30.Range("J3")
        .value = "查询..."
        .Interior.Color = 49407
    End With
    
    On Error GoTo ErrorHandler
    
重新查询:
    token = GetSetting("Momda", "login", "accessToken") '获取授权
    If token = "" Then
        登录
        GoTo 重新查询:
    End If
    
    Set ds = CreateObject("Scripting.Dictionary") '初始化字典
    For i = 2 To rng.Rows.count
        车号 = rng.Cells(i)
        到站 = rng.Cells(i, 6).value
        货主 = rng.Cells(i, 11).value
    
        If 到站 = "钦州港" And 货主 <> "" Then
        
            data = "{""ch"":""" & 车号 & """,""iftcbg"":"""",""njfzt"":""0"",""queryFlag"":""1"",""ydztgj"": ""70,35,40,60""}" '车号,变更Y,未交付0,
            n = i - 1
            Set http(n) = CreateObject("WinHttp.WinHttpRequest.5.1") 'CreateObject("MSXML2.XMLHTTP.6.0") '
            With http(n)
                url = "http://10.4.10.11/api/scjh/wayBillQuery/queryCargoArrival"
                .Open "POST", url, True
                .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
                .setRequestHeader "access_token", token
                .setRequestHeader "content-type", "application/json"
                .setRequestHeader "unitid", "36871"
                .setRequestHeader "userid", "36871145"
                .setRequestHeader "Referer", "http://10.4.10.11/hyManager/waybillQuery/arriveBillQuery"
                .setRequestHeader "Referrer-Policy", "strict-origin-when-cross-origin"
                .setRequestHeader "type", "inner" '20260606新增
                .send data
                
                Set rn = rng.Cells(i, 12)
                Set ds(http(n)) = rn '对象加入字典

            End With
        End If
        DoEvents
    Next
       
    Do '异步查询
        Dim dms(), m As Boolean
        dms = ds.keys '数组
        For Each dd In dms
            DoEvents
            On Error Resume Next
'            Set xhtp = d_Qs(dd)
            res = dd.responseText
            On Error GoTo 0
            If res <> "" Then
'                res = dd.responseText
                '登录失败时
                If dd.Status <> 200 Then
                    m = 登录
                    Debug.Print "授权：" & m
                    
                    '重置变量状态
                    s = ""
                    res = ""
                    Set rn = Nothing
                    Set ds = Nothing
                    
                    Exit For '退出for循环-关键
                End If
                
                '登录成功后
                ow.execScript "var jstr =" & res & ";"
                ow.execScript "function 变更类型(e){return '1' == e ? '取消托运' : '2' == e ? '变更收货人' : '3' == e ? '变更到站' : '4' == e ? '变更卸车地点' : ''}", "jscript"

                Dim list As Object, l As Long
                Set list = ow.eval("jstr.data.list")
                
                l = CallByName(list, "length", VbGet) '大于0时为True
                If l Then
                    For Each v In list
                        'Debug.Print v.dzyxhz
                        If IsNull(v.dzbgZtjgjc) Then
                            If s = "" Then
                                s = v.dzyxhz
                            Else
                                If s = v.dzyxhz Then
                                    'Debug.Print s
                                Else
                                    s = s & Chr(10) & v.dzyxhz
                                End If
                            End If
                            
                        '已交付
                        ElseIf v.ztgjjcend = "确认收货" Then
                            rn.value = v.ztgjjcend
                        
                        '变更
                        Else
                            Debug.Print v.ch & ":" & v.dzbgZtjgjc
                            'Debug.Print v.ch & ":" & v.dzbgZtjgjc
                            If s = "" Then
                                s = v.dzyxhz & "①"
                            Else
                                If s = v.dzyxhz & "①" Then
                                    s = s & "②"
                                Else
                                    s = s & Chr(10) & v.dzyxhz & "②"
                                End If
                            End If
                            
                            Set lit = 变更信息(v.ydid) '20260215更新
                            If lit Is Nothing Then
                                批注 = "请求错误：500"
                            Else '20260520更新
                                批注 = "、" & Chr(10) & _
                                "变更事项：" & ow.变更类型(lit.bglx) & Chr(10) & _
                                "变 更 人：" & lit.bgxx & Chr(10) & _
                                "变更原因：" & lit.bgyy & Chr(10) & _
                                "变更状态：" & lit.ztgjjc & Chr(10) & _
                                "原卸车地点：" & lit.dzyxhz & Chr(10) & _
                                "新卸车地点：" & lit.bgDzyxhz & Chr(10) & _
                                "原收货人：" & lit.shdwmc & Chr(10) & _
                                "新收货人：" & lit.bgxshrmc & Chr(10) & _
                                "原到站：" & lit.dzhzzm & Chr(10) & _
                                "新到站：" & lit.bghxdzzm
                            End If
                            
                            Set rn = ds(dd)
                            If rn.Comment Is Nothing Then
                                批注 = "第①组" & 批注
                                rn.AddComment 批注
                                rn.Comment.Shape.height = 150
                                rn.Comment.Shape.width = 260
                            Else
                                With rn.Comment
                                    批注 = .text & String(2, Chr(10)) & "第②组" & 批注
                                    .text text:=批注
                                    .Shape.height = 266
                                End With
                            End If
                                
                            批注 = ""
                            变更标志 = True
                            
                        End If
                        
                    Next
                    
                    '写入表格中
                    Set rn = ds(dd)
                    rn.value = s
                    
                End If
                
                '重置变量状态
                res = ""
                s = ""
                ds.Remove dd
                Set rn = Nothing
                Set dd = Nothing

            End If
        Next dd
        
        DoEvents '交回控制权给系统，防卡死！
            
        'token失效后重新查询
        If m Then
            m = False
            Debug.Print "重新查询！！！"
            GoTo 重新查询
        End If
    Loop Until ds.count = 0
    
    '变更状态判断
    With Sheet30.Range("J3")
        If 变更标志 Then
            .value = "变更！"
            .Interior.Color = 255
            MsgBox "请核对卸车地点！"
        Else
            .value = "无变更！"
            .Interior.Color = 5287936
        End If
    End With
    
    Set ds = Nothing
    Debug.Print "完成！！！"
    Exit Sub
    
ErrorHandler:
    Sheet30.Range("J3").value = "出错啦！"
     Set xhtp = Nothing
     GoTo 重新查询
End Sub

Function 变更信息(需求号)

    Dim http As Object, xhtp As Object, oDom As Object, ow As Object, token As String, data As String
    Set oDom = CreateObject("htmlfile")
    Set ow = oDom.parentWindow
    '票据号 = "364492509093690646"
    data = "{""xqdid"":""" & 需求号 & """}" '要求格式：{"xqdid":"202601RZ0528960001"}
    token = GetSetting("Momda", "login", "accessToken") '获取授权
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    With http
        url = "http://10.4.10.11/api/hp/wayBillChange/queryArriveChangeNoPermission"
        .Open "POST", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "access_token", token
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "unitid", "36871"
        .setRequestHeader "userid", "36871145"
        .setRequestHeader "Referer", "http://10.4.10.11/hyManager/waybillQuery/arriveBillQuery"
        .setRequestHeader "Referrer-Policy", "strict-origin-when-cross-origin"
        .send data
        
        res = .responseText
        ow.execScript "var jstr =" & res & ";"
        Dim list As Object, l As Long, s As String
        stu = ow.eval("jstr.msg || jstr.status")
    End With
    
    If stu = "OK" Then
        Set list = ow.eval("jstr.data['0']")
        Set 变更信息 = list
    ElseIf stu = 500 Then
        erg = ow.eval("jstr.error")
        Set 变更信息 = Nothing
        Debug.Print "请示错误：" & erg
    Else
        Set 变更信息 = Nothing
    End If
End Function

Sub 查询测试2()
    Dim xhtp As Object, js As Object, rng As Range, rngs As Range, res As String

    '车次 = "null"
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1")

    With Sheet30
        rs = .Cells(Rows.count, "R").End(3).Row + 1
        With .Range(.Cells(6, "Q"), .Cells(rs, "X"))
            .ClearContents
            .ClearComments '清除
            '.Interior.Color = 16777215 '15921906
        End With
    End With
        
        
    With xhtp
        开始时间 = VBA.Format(DateAdd("d", 0, Now), "yyyy-mm-dd%2018:01") '查询时间后推1天
        结束时间 = VBA.Format(DateAdd("d", 1, Now), "yyyy-mm-dd%2018:00") '次日18点

        url = "http://yt1dzh.crc.cr:30001/qb/getqbmlforallByPage?qbType=2&begin=" & 开始时间 & "&end=" & 结束时间 & "&pageNum=1&pageSize=100&cdm=*&ljdm=10&departmentId=99113"
        .Open "POST", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "Origin", "http://yt1dzh.crc.cr:30001"
        .setRequestHeader "Referer", "http://yt1dzh.crc.cr:30001/"
        
       
        '按车次查询
        .send "{""tm"":37436,""cc"":null,""khbz"":null}" '久
        res = .responseText
        写入表格 res, 0

    End With
    
    升序 Sheet30, Sheet30.Range("Q5:X5"), Sheet30.Range("R5"), 1
    Sheet30.Range("Q5:X5").AutoFilter
    
    'dr.RemoveAll '清空字典
End Sub

Sub 车站查询()
    Dim list As Object, l As Long, s As String
    Dim xhtp As Object, js As Object, rng As Range, rngs As Range, res As String, oDom As Object, ow As Object, data As String
    Set oDom = CreateObject("htmlfile")
    Set ow = oDom.parentWindow
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1")
    'Application.ScreenUpdating = False
    With xhtp
        
        'url = "http://yt1dzh.crc.cr:30001/crzm/selectByPage/1/100"'全路
        
        url = "http://yt1dzh.crc.cr:30001/crBzz/select" '车站
        data = "{""ljdm"":""10""}"
        
'        url = "http://yt1dzh.crc.cr:30001/department/select"
'        data = "{""smisUrl"":""department"",""lever"":3,""pid"":7111}" '车务段
        
        .Open "POST", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "Referer", "http://yt1dzh.crc.cr:30001/"
        .setRequestHeader "authorization", "Bearer eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiJ6cXpjd2QiLCJjcmVhdGVkIjoxNzYwNzYzMDg5ODA0LCJleHAiOjIxOTI3NjMwODl9.8Gvj-Pf3-6SZoGYCGQzaxFzbvaWXsbhPhNNqllGQQlKg9QeBbQBI0QKBSr4mz3Zp8tkfCE_pNTKkEP6iviZ6wQ"
        .setRequestHeader "Referrer-Policy", "strict-origin-when-cross-origin"
        .setRequestHeader "pragma", "no-cache"
        .send data

        res = .responseText
        ow.execScript "var jstr =" & res & ";"
        
        Set list = ow.eval("jstr.rows")
        For Each li In list
        车站 = li.zm '车站
        编号 = li.dblm
        
        '车站 = CallByName(li, "name", VbGet)'车务段
        '编号 = li.nameEn
        Debug.Print 车站 & ":" & 编号
        Next
    End With
    
    Set xhtp = Nothing
    Set oDom = Nothing
    'Application.ScreenUpdating = True
End Sub
Function 读取文件(fpth As String) As String
    Dim txt As String
'    filePath = "E:\桌面\Python\des加密.js" '文件读取
    Open fpth For Binary As #1
    txt = Input$(LOF(1), #1)
    Close #1 ' 关闭文件
    读取文件 = txt
End Function
