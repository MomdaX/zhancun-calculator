Attribute VB_Name = "查询站内车的到卸和变更"
Sub 变更查询()
    
    Dim xhtp As Object, ds As Scripting.Dictionary, oDom As Object, ow As Object, token As String, data As String, s As String, res As String
    Set oDom = CreateObject("htmlfile")
    Set ow = oDom.parentWindow
    
    Dim http() As Object, 变更标志 As Boolean
    Dim rngs As Range, rng As Range, rn As Range
'    Set rngs = Sheet30.Range("A5").CurrentRegion
'    Set rng = rngs.Columns(3)
    
    Set rngs = Sheet7.Cells.SpecialCells(xlCellTypeVisible)
    Set rng = rngs.Areas(2)
    
    If rng Is Nothing Then Exit Sub
    Set rng = rngs.Areas(2).Columns(4)
'
'    If rng.Rows.count = 0 Then Exit Sub
    
    ReDim http(1 To rng.Rows.count)
    rngs.ClearComments '清除批注
    
    变更标志 = False
    With Sheet7.Range("O3")
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
    For i = 1 To rng.Rows.count
        车号 = rng.Cells(i)
        到站 = rng.Cells(i, 5).value
        品名 = rng.Cells(i, 7).value
        货主 = rng.Cells(i, 10).value
    
        If 品名 <> "空" And 货主 <> "" Then
        
            data = "{""ch"":""" & 车号 & """,""iftcbg"":"""",""njfzt"":""0"",""queryFlag"":""1"",""ydztgj"": ""70,35,40,60""}" '车号,变更Y,未交付0,
            n = i
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
                            If s = "" Then
                                s = v.dzyxhz & "①"
                            Else
                                If s = v.dzyxhz & "①" Then
                                    s = s & "②"
                                Else
                                    s = s & Chr(10) & v.dzyxhz & "②"
                                End If
                            End If
                            
                            Set lit = 变更信息(v.ydid)
                                                    
                            批注 = "、" & Chr(10) & _
                            "变更事项：" & ow.变更类型(lit.bglx) & Chr(10) & _
                            "变 更 人：" & lit.bgxx & Chr(10) & _
                            "变更原因：" & lit.bgyy & Chr(10) & _
                            "原卸车地点：" & lit.dzyxhz & Chr(10) & _
                            "新卸车地点：" & lit.bgDzyxhz & Chr(10) & _
                            "原收货人：" & lit.shdwmc & Chr(10) & _
                            "新收货人：" & lit.bgxshrmc & Chr(10) & _
                            "原到站：" & lit.dzhzzm & Chr(10) & _
                            "新到站：" & lit.bghxdzzm

                            Set rn = ds(dd)
                            If rn.Comment Is Nothing Then
                                批注 = "第①组" & 批注
                                rn.AddComment 批注
                                rn.Comment.Shape.height = 130
                                rn.Comment.Shape.width = 200
                            Else
                                With rn.Comment
                                    批注 = .text & String(2, Chr(10)) & "第②组" & 批注
                                    .text text:=批注
                                    .Shape.height = 255
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
    With Sheet7.Range("O3")
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
    data = "{""xqdid"":" & 需求号 & "}"
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
        Set list = ow.eval("jstr.data['0']")
    End With

    Set 变更信息 = list
    
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
    
    data = "{""userName"":""" & usid & """,""password"":""" & pw & """,""secType"":""RSA"",""type"":""inner"",""x"":""" & x & """,""token"":""" & tk & """,""extData"":{""grant_type"":""password""},""agent"":""Chrome 100"",""deviceId"":""Windows64-6.1"",""deviceName"":""Chrome::0f5d844e80cda20b8bc6b88f9fb37859"",""osName"":""Windows64"",""osVersion"":""Windows64-6.1""}"
    
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1")
    With xhtp

        url = "http://10.4.10.11/api/zuul/login"
        .Open "POST", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "channel", "P"
        .setRequestHeader "content-type", "application/json"
        .setRequestHeader "pragma", "no-cache"
        .setRequestHeader "rtrackid", uid
        .setRequestHeader "Referer", "http://10.4.10.11/login"
        .setRequestHeader "Referrer-Policy", "strict-origin-when-cross-origin"
        .setRequestHeader "type", "inner" '20260606新增
        .send data

        res = .responseText
        ow.execScript "var js =" & res & ";"
        
        Dim list As Object, l As Long, msg As String
        msg = ow.js.msg
        If msg <> "OK" Then '账号或密码错误
            MsgBox msg
            End
        End If
        
        accessToken = ow.eval("js.data.accessToken")
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

Function 读取文件(fpth As String) As String
    Dim txt As String
'    filePath = "E:\桌面\Python\des加密.js" '文件读取
    Open fpth For Binary As #1
    txt = Input$(LOF(1), #1)
    Close #1 ' 关闭文件
    读取文件 = txt
End Function
