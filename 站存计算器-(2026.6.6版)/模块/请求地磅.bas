Attribute VB_Name = "请求地磅"
Sub 查询地磅数据(Optional Hms As String = "")
'    Dim xmlhttp As New MSXML2.ServerXMLHTTP
'    Dim xmlhttp As Object: Set xmlhttp = CreateObject("WinHttp.WinHttpRequest.5.1")
'    Dim xmlhttp As Object: Set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
    Dim HTML As String, url As String, rngc As Range, data As String, 用户名$, 密码$, 日期$, p, zz
    Dim d As New Dictionary, dc As New Dictionary, dz As New Dictionary, ds As New Dictionary, crr, reg As New RegExp, Cookie$, drs()
    reg.Pattern = "\d+"
'    Application.ScreenUpdating = False '取消屏幕刷新
    url = "http://10.190.136.155:8888/Gsgz/login/userAuth"
    用户名 = URL编码("钦州港")
    密码 = URL编码("qzg123456#")
    data = "User_Name=" & 用户名 & "&User_Password=" & 密码
    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "POST", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
        .setRequestHeader "Origin", "http://10.190.136.155:8888"
        .setRequestHeader "Host", "10.190.136.155:8888"
        .setRequestHeader "Referer", "http://10.190.136.155:8888/Gsgz/loginPage"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send data
        Cookie = Split(.getResponseHeader("Set-Cookie"), ";")(0)
        'Debug.Print cookie
        
        url = "http://10.190.136.155:8888/Gsgz/GDH/listDate"
        '例子：  日期 = "date=20231231"
        Dim cdz%
        cdz = Sheet12.Range("T1").value
        日期 = "date=" & Format(DateAdd("d", -cdz, Date), "yyyymmdd") '当天日期
        
        .Open "post", url, False
        .setRequestHeader "Cookie", Cookie
        .setRequestHeader "Referer", "http://10.190.136.155:8888/Gsgz/view/index"
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        .send 日期
'        Debug.Print .responseText
        Dim res As Object
        Set res = JsonConverter.ParseJson(.responseText)

        If res.count > 0 Then
            '清空
            Set rngc = Sheet12.Range("A1").CurrentRegion.Offset(2)
            rngc.CurrentRegion.Interior.Color = xlNone
            rngc.ClearContents
            
            With Sheet12
                .Range("S2:V2").value = Array("时间", "辆数", "地磅", "方向")
'                .Range("S2:V2").Interior.ColorIndex = 49
'                .Range("S2:V2").Font.Size = 11
'                .Range("S2:V2").Font.ColorIndex = 2
                If .Cells(3, "S") <> "" Then .Cells(2, "S").CurrentRegion.Offset(2).ClearContents
                For Each Item In res
                r = .Cells(Rows.count, "S").End(3).Row + 1
                    .Cells(r, "S").value = Item(1)
                    .Cells(r, "T").value = Item(2)
                    .Cells(r, "U").value = Item(4)
                    .Cells(r, "V").value = Item(7)
                Next
                
                If Hms = "" Then
                    Hms = "Hms=" & .Cells(r, "S").value
                Else
                    Hms = "Hms=" & Hms
                End If
            
            End With

        Else
            Application.ScreenUpdating = True '开启屏幕刷新
            MsgBox "没有该日期数据！！！"
            End
        End If
        
        url = "http://10.190.136.155:8888/Gsgz/GDH/listGDH?" & Hms
        .Open "get", url, False
        .setRequestHeader "Cookie", Cookie
        .send
'        Debug.Print .responseText
        Set res = Nothing
        Set res = JsonConverter.ParseJson(.responseText)
    End With

        If res.count > 0 Then
'            Debug.Print "正在加载数据..."
        Else
            Application.ScreenUpdating = True '开启屏幕刷新
            MsgBox "未查到过磅数据！！！"
            End
        End If

        For Each Item In res
            With Sheet12
                r = .Cells(.Rows.count, 2).End(3).Row + 1 '定位
'                If Item("wagonNumber") = 5479086 Then Stop
                .Cells(r, 1).value = r - 2 '序号
                .Cells(r, 2).value = Item("createDate") '时间
                .Cells(r, 3).value = Item("gdh_Station") '测点
                .Cells(r, 4).value = Item("wagonNumber") '车号
                .Cells(r, 5).value = Item("motorcycleTtype") '车型
                .Cells(r, 6).value = Item("roughWeight") '毛重(吨)
                .Cells(r, 7).value = Item("tare") '皮重(吨)
                .Cells(r, 8).value = Item("suttle") '净重(吨)
                .Cells(r, 9).value = Item("indicatedDeight") '标重(吨)
                '.Cells(r, 10).value = Item("weightOff") '允增(吨)
    
                .Cells(r, 13).value = Item("beforeAfterPartial") '前后偏(吨)
                .Cells(r, 14).value = Item("aboutPartial") '左右偏(毫米)
                .Cells(r, 15).value = Item("speed") '速度(km/h)
'                .Cells(r, 16).value = Item("direction") '股道
                .Cells(r, 17).value = Item("direction") '方向
            End With
        Next

    
    '过磅后的股道导入
    crr = Sheet7.Range("A4:J" & Sheet7.Cells(Sheet7.Rows.count, "D").End(3).Row) '站存测试表
    For i = 1 To UBound(crr)
    
        '收集编组信息(给J列用)2025/5/31
        If ds.Exists(crr(i, 1)) Then
            n = n + 1
            ds(crr(i, 1)).Add (Format(n, "00") & "、" & crr(i, 4) & " " & crr(i, 8) & " " & crr(i, 10)), "" '序号、车号 到站 品名
            'Debug.Print crr(i, 1) & "-" & (n & "、" & crr(i, 4) & " " & crr(i, 10))
        Else
            n = 1
            Set ds(crr(i, 1)) = CreateObject("scripting.Dictionary")
            ds(crr(i, 1)).Add (Format(n, "00") & "、" & crr(i, 4) & " " & crr(i, 8) & " " & crr(i, 10)), "" '序号、车号 到站 品名
        End If
        
        '2024/5/6更新
        d(crr(i, 4)) = crr(i, 1) '车号=股道
    Next
    
    drs = Sheet12.Range("A2").CurrentRegion '轨道衡表
    Dim 车号 As String, 股道 As String, rngs As Range, rng As Range, rn As Range, dt As New Dictionary
    For i = 3 To UBound(drs)
        车号 = drs(i, 4)
        dt(车号) = d(车号)  '车号=股道d
        dc(d(车号)) = dc(d(车号)) + 1 '股道计数
    Next
    
    '生成P列的数据
    Set rng = Sheet12.Cells(3, "P").Resize(dt.count, 1)
    rng = WorksheetFunction.Transpose(dt.items) '股道导入轨道衡表
    
    
    For Each T_max In dc.keys
        If dc(T_max) = WorksheetFunction.Max(dc.items) Then
            股道 = T_max
        End If
    Next
    
    '生成J列的数据
    If 股道 <> "" Then
        Dim sta(), Tng As Range, rr%, txch$
        sta = ds(股道).keys
        txch = Split(Split(sta(0), "、")(1), " ")(0)
        Set Tng = Columns(4).Find(txch)
        If Not Tng Is Nothing Then
            rr = Tng.Row
        Else
            rr = 3
        End If
        Set rngs = Sheet12.Cells(rr, "J").Resize(ds(股道).count, 1)
        rngs = WorksheetFunction.Transpose(ds(股道).keys)
        显示信息.标记到站方向颜色 rngs
         For Each rn In rng
            If rn.value <> 股道 Then
                rn.Interior.ColorIndex = 6 '黄
            End If
         Next
    End If
    Set d = Nothing
    Set dc = Nothing '清空股道
    Set ds = Nothing '清空品类
    Set dt = Nothing '清空品类
    
    '标题栏
    With Sheet12 '轨道衡数据表
        股道 = IIf(股道 <> "", 股道 & "道", "2.9")
        .Range("A2:Q2").value = Array("序号", "时间", "测点", "车号", "车型", "毛重(吨)", "自重(吨)", "载重(吨)", "标重(吨)", 股道 & "编组(序号、车号 品名)/允增", "总重(吨)", "超欠(吨)", "前后偏(吨)", "左右偏(毫米)", "速度(km/h)", "股道", "方向")
        .Range("A2:Q2").Interior.Color = 5287936
        
        '自重录入数据库
        crr = Sheet7.Range("C4:K" & Sheet7.Cells(Sheet7.Rows.count, "C").End(3).Row) '站存测试表
        For i = 1 To UBound(crr, 1)
            If val(crr(i, 3)) > 0 Then dc(crr(i, 1)) = val(crr(i, 3)) '自重
            If val(crr(i, 5)) > 0 And Not dz.Exists(crr(i, 1) & ":" & crr(i, 8)) Then '载重
                dz(crr(i, 1) & ":" & crr(i, 8)) = crr(i, 5)
            End If
        Next
        
        
        Dim rs%
        With Sheet10
            crr = .Range(.Range("F2"), .Cells(.Rows.count, "G").End(3).Offset(1)) '自重表
            For i = 1 To UBound(crr, 1)
                On Error Resume Next
                For Each zz In dz.keys '按品名获取载重
                    If Not zz = "" Then
                        If crr(i, 1) = Split(zz, ":")(0) And dz(zz) > 0 Then
                            If InStr(.Cells(i + 1, "H").value, Split(zz, ":")(1) & ":" & dz(zz)) = 0 Then
                                If .Cells(i + 1, "H").value = "" Then
                                    .Cells(i + 1, "H").value = Split(zz, ":")(1) & ":" & dz(zz)
                                Else
                                    If InStr(.Cells(i + 1, "H").value, Split(zz, ":")(1)) = 0 Then '判断是否有相同的品名
                                        .Cells(i + 1, "H").value = .Cells(i + 1, "H").value & " " & Split(zz, ":")(1) & ":" & dz(zz)
                                    Else
                                        pm = Split(.Cells(i + 1, "H").value, " ") '如有相同品名，取大更新
                                        For p = 0 To UBound(pm)
                                            If Split(pm(p), ":")(0) = Split(zz, ":")(1) Then '比较品名是否相同
                                                If dz(zz) > val(Split(pm(p), ":")(1)) Then
                                                    pm(p) = Split(zz, ":")(1) & ":" & dz(zz)
                                                    .Cells(i + 1, "H").value = Join(pm, " ")
                                                    Exit For '退出
                                                End If
                                            End If
                                        Next
                                        
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
                dc.Remove crr(i, 1) '移除已有的车型
                On Error GoTo 0
            Next
            For Each cx In dc.keys '添加新的车型入库
                rs = .Cells(.Rows.count, "F").End(3).Row + 1 '定位
                .Cells(rs, "E").Formula = "=row()-1"
                .Cells(rs, "F").value = cx
                .Cells(rs, "G").value = dc(cx)
                .Cells(rs, "I").value = Now()
            Next
            With .Cells(.Rows.count, "G").End(3).CurrentRegion
                .Borders.LineStyle = xlContinuous '边框
                .HorizontalAlignment = xlCenter '居中
                .VerticalAlignment = xlCenter '居中
            End With
        End With
        
        '自重纠正
        With Sheet10
            crr = .Range(.Range("F2"), .Cells(.Rows.count, "G").End(3)) '自重表
            For i = 1 To UBound(crr, 1)
                On Error Resume Next
                dc(crr(i, 1)) = val(crr(i, 2)) '加入dc字典中
                On Error GoTo 0
            Next
        End With
        For i = 3 To .Cells(.Rows.count, "G").End(3).Row '遍历轨道衡数据表
            If Not dc.Exists(.Cells(i, "E").value) And val(.Cells(i, "G").value) > 0 Then dc(.Cells(i, "E").value) = val(.Cells(i, "G").value) '当车型不存在，且该车型自重大于0时，录入
            If .Cells(i, "G") = 0 Or .Cells(i, "G") = "" Then
                If dc.Exists(.Cells(i, "E").value) Then '反录至轨道衡数据表中
                    .Cells(i, "G").value = dc(.Cells(i, "E").value) '自重
                    .Cells(i, "H").value = .Cells(i, "H").value - dc(.Cells(i, "E").value) '载重
                Else
                    .Cells(i, "G").value = 0 '自重
                End If
            End If
        Next
        
        '计算
        .Cells(2, "L").value = "超欠(吨)": .Cells(2, "K").value = "总重(吨)"
        For k = 3 To r
            '前后偏
            If Abs(.Cells(k, "M")) > 6.9 Then .Cells(k, "M").Interior.ColorIndex = 3 '超偏(红)
            '左右偏
            If Abs(.Cells(k, "N")) > 69 Then .Cells(k, "N").Interior.ColorIndex = 3 '超偏(红)
            '标重纠正
            If .Cells(k, "I") = 0 Then
'                Set mmk = reg.Execute(.Cells(k, "E").Value)
                If .Cells(k, "E").value <> "" And .Cells(k, "E").value <> "*" Then
                    .Cells(k, "I") = reg.Execute(.Cells(k, "E").value)(0).value
                End If
            End If
            '超吨
            .Cells(k, "L") = .Cells(k, "H") - .Cells(k, "I")
            If .Cells(k, "I") < 65 And .Cells(k, "L") > 3.5 Then    '60吨的车
                .Cells(k, "L").Interior.ColorIndex = 3 '超4吨(红)
                With .Cells(k, "L").Offset(0, -4)
                     .Interior.ColorIndex = IIf(val(.value) > 65, 6, IIf(val(.value) > 64.5, 3, xlNone)) '载总(黄)
                End With
            ElseIf .Cells(k, "I") > 65 And .Cells(k, "L") > 2 Then   '70吨的车
                .Cells(k, "L").Interior.ColorIndex = 3 '超4吨(红)
                With .Cells(k, "L").Offset(0, -4)
                     .Interior.ColorIndex = IIf(val(.value) > 72.5, 3, 6) '载总(黄)
                End With
            End If
            .Cells(k, "K") = .Cells(k, "H") + .Cells(k, "G") '总重=载重+自重
        Next
    '表头名
        If res.count > 0 Then
            .Cells(1, 1).value = res(1)("gdh_Station") & "-" & res(1)("createDate") & "轨道衡数据"
        End If
    End With
'    Application.ScreenUpdating = True '开启屏幕刷新
    '清空对象
    Set xmlhttp = Nothing
    Set rngc = Nothing
    Set reg = Nothing
    Set res = Nothing
    Set d = Nothing
    Set dc = Nothing
    crr = ""
    r = 0: rs = 0: i = 0: k = 0
    日期 = "": Hms = ""
'    MsgBox "查询完毕！"
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


Function URL编码(strText)
    Static objHtmlfile As Object
    If objHtmlfile Is Nothing Then
        Set objHtmlfile = CreateObject("htmlfile")
        objHtmlfile.parentWindow.execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
    End If
    URL编码 = objHtmlfile.parentWindow.encode(strText)
    Set objHtmlfile = Nothing
End Function

'Function 列号(key As String)
'
'Select Case key
'    Case "aboutPartial" '左右偏(毫米)
'    列号 = 13
'    Case "beforeAfterPartial" '前后偏(吨)
'    列号 = 12
'    Case "createDate" '时间
'    列号 = 1
'    Case "direction" '方向
'    列号 = 15
'    Case "gdh_Id" '车站ID
'    列号 = 16
'    Case "gdh_Station" '测点
'    列号 = 2
'    Case "indicatedDeight" '标重(吨)
'    列号 = 8
'    Case "motorcycleTtype" '车型
'    列号 = 4
'    Case "roughWeight" '毛重(吨)
'    列号 = 5
'    Case "speed" '速度(km/h)
'    列号 = 14
'    Case "suttle" '净重(吨)
'    列号 = 7
'    Case "tare" '皮重(吨)
'    列号 = 6
'    Case "wagonNumber" '车号
'    列号 = 3
'    Case "weightOff" '允增(吨)
'    列号 = 9
'End Select
'End Function
'
'Sub Macro1()
'
' Dim lb As ListBox, arr(5)
'
'    For Each lb In ActiveSheet.ListBoxes
'        lb.RemoveAllItems
'        lb.AddItem "it2"
'        Debug.Print lb.Selected
'        Debug.Print lb.Name
'        Debug.Print lb.Index
'        lb.Select
'    Next
'    ActiveSheet.ListBoxes (1)
'
'End Sub
