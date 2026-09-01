Attribute VB_Name = "显示信息"
Sub 查询载重()
    Dim rngs As Range, rng As Range, xrr, reg As New RegExp, dc As New Dictionary, z
    Set rngs = Sheet7.Cells.SpecialCells(xlCellTypeVisible)
'    Set rng = rngs.Areas(2)
'    rng.Select
    For Each rng In rngs.Areas '取筛选后的可见区域进数组中
        k = k + 1
        If k = 2 Then '取第二个可见区域：数据区域
            xrr = rng
            r = rng.Row
            Exit For
        End If
    Next
    
    reg.Global = True
    reg.Pattern = ":(\d+)" '按品名取对应的重量
    zrr = Sheet10.Range("F1").CurrentRegion '始发直达列车表
    For j = 2 To UBound(zrr)
        If zrr(j, 4) <> "" Then
            If Left(zrr(j, 2), 1) = "G" Then
               dc(zrr(j, 2)) = zrr(j, 4)
            Else
                Set zz = reg.Execute(zrr(j, 4))
                For Each z In zz
                    If val(z.SubMatches(0)) > dc(zrr(j, 2)) Then dc(zrr(j, 2)) = val(z.SubMatches(0))
                Next
            End If
        End If
    Next
    
    '统计
    For i = 1 To UBound(xrr)
        A = val(xrr(i, 5))
        If xrr(i, 7) = "" Then
            If Left(xrr(i, 3), 1) = "G" Then
                pms = Split(dc(xrr(i, 3)), " ")
                For Each pm In pms
                    If InStr(xrr(i, 14), Split(pm, ":")(0)) Then
                        b = val(Split(pm, ":")(1))
                        Exit For
                    ElseIf xrr(i, 1) = "Y11" Or xrr(i, 1) = "Y12" Or xrr(i, 1) = "Y13" Or xrr(i, 1) = "Y14" Then
                        If Split(pm, ":")(0) = "柴油" Then
                            b = val(Split(pm, ":")(1))
                        End If
                    ElseIf xrr(i, 1) = "Y9" Or xrr(i, 1) = "Y10" Then
                        If Split(pm, ":")(0) = "汽油" Then
                            b = val(Split(pm, ":")(1))
                        End If
                    ElseIf xrr(i, 1) = "Y5" Or xrr(i, 1) = "Y6" Then
                        If Split(pm, ":")(0) = "航空煤油" Then
                            b = val(Split(pm, ":")(1))
                        End If
                    End If
                Next
            Else
                b = dc(xrr(i, 3))
            End If
        Else
            If dc(xrr(i, 3)) > 0 Then '该车型有载重数据时
                If Left(xrr(i, 3), 1) = "G" Then '为罐车时
                    pms = Split(dc(xrr(i, 3)), " ") '拆分品名
                    For Each pm In pms '循环各个品名下的载重
                        If InStr(xrr(i, 14), Split(pm, ":")(0)) Then '记事栏
                            b = val(Split(pm, ":")(1))
                            Exit For
                        ElseIf InStr(xrr(i, 10), Split(pm, ":")(0)) Then '品名栏
                            b = val(Split(pm, ":")(1))
                            Exit For
                        End If
                    Next
                Else
                    b = dc(xrr(i, 3))
                End If
            Else
                b = val(xrr(i, 7))
            End If
        End If
        Sheet7.Cells(r, "L").value = b '写入
        s = s + A + b '累加
        r = r + 1
    Next
    
    
    MsgBox "总重: " & s '输出
End Sub



Sub 到站方向()
Attribute 到站方向.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim frr() As Variant, dv As New Dictionary, dh As New Dictionary, dw As New Dictionary, df As New Dictionary, rng As Range, rng1 As Range, rng2 As Range, rng3 As Range, rng4 As Range, rngs As Range
    Dim reg As New RegExp, rn As Range
    Application.ScreenUpdating = False '关闭屏幕刷新
    '读取车站方向数据
    frr = Sheet2.Range("A1:C2").CurrentRegion
    For i = 2 To UBound(frr)
        df(frr(i, 1)) = frr(i, 2) & "-" & frr(i, 3) & "-" & frr(i, 4)
    Next
    
    '设置Sheet7数据验证
    With Sheet7
        .Cells.Validation.Delete
        .Cells.Interior.Color = xlNone '整个表格的颜色重置
        r = .Cells(.Rows.count, 1).End(3).Row
        Set rng1 = .Range(.Cells(4, 8), .Cells(r, 8)) '到站列
        Set rng2 = .Range(.Cells(4, 11), .Cells(r, 11)) '发站列
        Set rng3 = .Range(.Cells(4, 17), .Cells(r, 17)) '到达时间列，计算大点车
'        Set rngs = Union(rng1, rng2)
'        rngs.Select
    End With
    
    For Each rng In rng1 '到站
        If rng.value <> "" And df(rng.value) <> "" Then
            With rng
                '方向
                If df(rng.value) Like "*沙*" Then
                    .Interior.ColorIndex = 8 '蓝
                ElseIf df(rng.value) Like "*南*" Then
                    .Interior.ColorIndex = 40 '黄
                Else
                    .Interior.ColorIndex = 19 '淡黄
                End If
                With .Validation
                    .Delete
                    .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
                    .InputTitle = "方向-局-口向"
                    .InputMessage = df(rng.value)
                End With

            End With
        End If
        
        '车型
        reg.Pattern = "\w+" '"[^\d]+"
        If Mid(rng.Offset(0, -5).value, 1, 2) = "DK" Then '大D车
            With rng.Offset(0, -5).Resize(1, 2).Font
                .Bold = True '加粗
                .ColorIndex = 3 '红
            End With
        ElseIf Mid(rng.Offset(0, -5).value, 1, 2) = "YW" Or Mid(rng.Offset(0, -5).value, 1, 2) = "YZ" Then '客车
            With rng.Offset(0, -5).Resize(1, 2).Font
                .Bold = True '加粗
                .ColorIndex = 3 '红
            End With
        ElseIf Left(rng.Offset(0, -5).value, 1) = "B" Then '机保车
            With rng.Offset(0, -5).Resize(1, 2).Font
                .Bold = True '加粗
                .ColorIndex = 21 '紫
            End With
        ElseIf Left(rng.Offset(0, -5).value, 1) = "K" Then '老K车
            With rng.Offset(0, -5).Resize(1, 2).Font
                .Bold = True '加粗
                .ColorIndex = 17 '蓝
            End With
        ElseIf Left(rng.Offset(0, -5).value, 1) = "T" Then '检衡车
            With rng.Offset(0, -5).Resize(1, 2).Font
                .Bold = True '加粗
                .ColorIndex = 22 '粉红
            End With
        ElseIf Left(rng.Offset(0, -5).value, 1) = "P" Then
            rng.Offset(0, -5).Interior.ColorIndex = 6 '盖车(P)标色-黄
            If Not Sheet6.Range("B43:B66").Find(what:=rng.Offset(0, -7).value, lookat:=xlWhole) Is Nothing Then '盖车不能进鹰岭
                r = Sheet6.Range("B43:B66").Find(what:=rng.Offset(0, -7).value, lookat:=xlWhole).Row
                Sheet6.Cells(r, "E").Interior.ColorIndex = 3   'E列标红
                Sheet6.Cells(r, "J").Interior.ColorIndex = 3   'J列标红
            ElseIf Not Sheet6.Range("B69:B82").Find(what:=rng.Offset(0, -7).value, lookat:=xlWhole) Is Nothing Then '盖车不能进栈桥
                r = Sheet6.Range("B69:B82").Find(what:=rng.Offset(0, -7).value, lookat:=xlWhole).Row
                Sheet6.Cells(r, "E").Interior.ColorIndex = 7   'E列标(粉红)
            End If
        ElseIf Left(rng.Offset(0, -5).value, 1) = "C" Then
            If Not Sheet6.Range("B67:B85").Find(what:=rng.Offset(0, -7).value, lookat:=xlWhole) Is Nothing Then '高边不能进中油
                r = Sheet6.Range("B67:B85").Find(what:=rng.Offset(0, -7).value, lookat:=xlWhole).Row
                Sheet6.Cells(r, "E").Interior.ColorIndex = 3   'E列标红
                Sheet6.Cells(r, "J").Interior.ColorIndex = 3   'J列标红
            End If
        Else
            Set t = reg.Execute(rng.Offset(0, -5).value)
            If reg.Test(rng.Offset(0, -5).value) Then
                If InStr(t(0), "X") > 0 Then rng.Offset(0, -5).Interior.ColorIndex = 15 '平板车(X/NX)标色-灰(15)/紫(14336204)
            End If
        End If
        
        '自备罐
        If Left(rng.Offset(0, -4).value, 1) = 0 Then
            If Mid(rng.Offset(0, -4).value, 2, 1) = 7 Then
                rng.Offset(0, -4).Font.Bold = True '加粗-中粮罐
            Else
                rng.Offset(0, -4).Interior.ColorIndex = 15 '灰-自备罐
            End If
        End If
        
        '品名
        If (InStr("汽油航煤", rng.Offset(0, 2).value) > 0 Or InStr(rng.Offset(0, 6).value, "汽油") > 0) And rng.Offset(0, 6).value <> "原装汽油" Then '20260201修复
            rng.Offset(0, 2).Interior.ColorIndex = 6
            If Not dw.Exists(rng.Offset(0, -7).value) Then '取汽油所在股道
                dw(rng.Offset(0, -7).value) = 6
            End If
        End If
        
        '点后开，重点事项备注
        If rng.Offset(0, 6).value Like "*[暂走去向点后开列扣检无计划坏超偏未脏不禁止有洗排磨损]*" Then
'            rng.Offset(0, 2).Interior.ColorIndex = 6
            For Each va In Split(Replace(rng.Offset(0, 6).value, ";", " "), " ")
                If va Like "*[暂不走去向点后开列扣检无计划坏超偏未脏禁止有洗排磨损]*" Then
                    If InStr(va, "不入扣") = 0 Then
                        If Not dv.Exists(rng.Offset(0, -7).value & va) Then
                            dh(rng.Offset(0, -7).value) = va & Chr(10) & dh(rng.Offset(0, -7).value) '取点后开所在股道和值
                            dv(rng.Offset(0, -7).value & va) = va '录入字典去重
                        End If
                    End If
                End If
            Next
        End If
        DoEvents
    Next

    For Each rng In rng2 '发站
        With rng.Validation
            .Delete
            .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator:=xlBetween
            .InputTitle = "方向-局-口向"
            .InputMessage = df(rng.value)
        End With
        DoEvents
    Next
    
    '大点车，老牌车
    For Each rn In rng3
        rn.HorizontalAlignment = xlCenter '居中
        If rn.value > 49 Then
           rn.Interior.ColorIndex = 19 '黄色
        End If
        DoEvents
    Next
    
    With Sheet6
'        .Range("J2:J90").Interior.ColorIndex = xlNone 'J列颜色重置
        For Each key In dw.keys '有汽油的股道路标黄
            If Mid(key, 1, 1) = "X" And Mid(key, 2) > 10 Then '2024-8-25修复=屏蔽11-15道
'                Debug.Print "11-15道"
            Else
                r = .Range("B2:B92").Find(what:=key, lookat:=xlWhole).Row
                .Cells(r, "J").Interior.ColorIndex = dw(key)
            End If
            DoEvents
        Next
        
        For Each key In dh.keys '点后开
            On Error Resume Next
            If Mid(key, 1, 1) = "X" And Mid(key, 2) > 10 Then '2024-8-25修复=屏蔽11-15道
'                Debug.Print "11-15道"
            Else
                r = .Range("B2:B92").Find(what:=key, lookat:=xlWhole).Row
                .Cells(r, "A").value = dh(key)
            End If
            On Error GoTo 0
            DoEvents
        Next
        
        '把最后面的回车去掉
        Set rng4 = .Range("A2:A92")
        For Each rn In rng4
            If rn <> "" Then
                If Asc(Right(rn.value, 1)) = 10 Then rn.value = Left(rn.value, Len(rn.value) - 1)
            End If
            DoEvents
        Next
        
        '重量核对，编车不得超5000吨
        Set rng5 = .Range("K2:K92")
        For Each rn In rng5
            If rn.value > 5000 Then
                rn.Interior.ColorIndex = 3 '超吨标记为红色
            End If
            DoEvents
        Next
        
        股道占用情况
    End With
    Application.ScreenUpdating = True '开启屏幕刷新
End Sub

Function 股道占用情况()
    Dim reg As New RegExp, d_zy As New Dictionary, rng As Range, rn As Range, tepm_fx As Double
    Dim M1 As Boolean, M2 As Boolean
    Set rng = Sheet6.Range("B2:B26") '1-X10道的换长录进字典
    reg.Pattern = "\d+"
    For Each rn In rng
        股道 = val(reg.Execute(rn)(0).value)
        d_zy(股道) = d_zy(股道) + rn.Offset(0, 4).value
    Next
    Set rng = Sheet6.Range("B2:B11") '1-10道填色
    With rng.Offset(0, 1).Resize(10, 4) '重置
        .Interior.ColorIndex = xlNone '无色
        .Font.Bold = False '不加粗
        .Font.ColorIndex = 0 '黑字
    End With
    For Each rn In rng
        股道 = val(reg.Execute(rn)(0).value)
'        Debug.Print 股道 & "道:" & d_zy(股道)
        If d_zy(股道) = 0 Then
            rn.Offset(0, 1).value = ""
            rn.Offset(0, 1).Interior.ColorIndex = 24
        Else
            tepm_fx = d_zy(股道)
            M1 = tepm_fx > rn.Offset(0, 4).value
            M2 = rn.Offset(0, 4).value <> ""
            If M1 And M2 Then
                With rn.Offset(0, 1)
                .value = tepm_fx '赋值
                .Font.Bold = True '加粗
                .Font.Color = 255 '红色
                End With
            End If
            
            '编长
            If IsNumeric(rn.Offset(0, 1).value) And rn.Offset(0, 1).value > 70 Then
                With rn.Offset(0, 2)
                    .Interior.ColorIndex = 3 '红色
                    .Font.Bold = True '加粗
                    .Font.ColorIndex = 2 '白色
                End With
            End If
            
            '超长
            If rn.Offset(0, 4).value > 70 Then
                With rn.Offset(0, 4)
                    .Interior.ColorIndex = 3 '红色
                    .Font.Bold = True '加粗
                    .Font.ColorIndex = 2 '白色
                End With
            End If
        End If
    Next
End Function

Sub 标记到站方向颜色(rng As Range)
    Dim frr() As Variant, df As New Dictionary, rn As Range, reg As New RegExp
    '读取车站方向数据
    frr = Sheet2.Range("A1:C2").CurrentRegion
    For i = 2 To UBound(frr)
        df(frr(i, 1)) = frr(i, 2)
    Next

    With rng.Font '全局文字重置
        .ColorIndex = xlAutomatic '颜色自动
        .Bold = False '不加粗
    End With
    
    '方案1
    For Each rn In rng
        If rn <> "" Then '股道不空
        For Each ss In Split(rn, " ")
            reg.Pattern = "[一-龢]+"
            If reg.Test(ss) Then
                k = reg.Execute(ss)(0)
    '            Debug.Print k
                reg.Pattern = k '到站匹配表达式
                Set t = reg.Execute(rn)
                If reg.Test(rn) And Not IsNumeric(k) Then
                    f = t(0).FirstIndex + 1 '开始位置
                    l = t(0).length '长度
                    If df(k) = "沙口" Then
'                        Debug.Print rn.Value & "沙口"
                        rn.Characters(f, l).Font.ColorIndex = 23 '蓝
                    ElseIf df(k) = "南口" Then
'                        Debug.Print rn.Value & "南口"
                        rn.Characters(f, l).Font.ColorIndex = 46 '黄
                    Else
'                        Debug.Print rn.Value & "其他（P、C、X、G、空车，中粮.....）"
                        rn.Characters(f, l).Font.ColorIndex = xlAutomatic '自动
                    End If
                End If
             End If
        Next
        
        '方案2
'            For Each k In df.Keys '到站标记
'                reg.Pattern = k '到站匹配表达式
'                Set t = reg.Execute(rn)
'                If reg.Test(rn) And Not IsNumeric(k) Then
'                    f = t(0).FirstIndex '开始位置
'                    l = t(0).Length '长度
'                    If df(k) = "沙口" Then
''                        Debug.Print rn.Value & "沙口"
'                        rn.Characters(f, l + 1).Font.ColorIndex = 23 '蓝
'                    ElseIf df(k) = "南口" Then
''                        Debug.Print rn.Value & "南口"
'                        rn.Characters(f, l + 1).Font.ColorIndex = 46 '黄
'                    Else
''                        Debug.Print rn.Value & "其他（P、C、X、G、空车，中粮.....）"
'                        rn.Characters(f, l + 1).Font.ColorIndex = xlAutomatic '自动
'                    End If
'                End If
'            Next
            
            For Each k In Array("YW", "D", "P", "到卸", "黑罐") '属性标记
                reg.Pattern = k 'K是属性匹配表达式
                Set t = reg.Execute(rn)
                If reg.Test(rn) Then
                    f = t(0).FirstIndex '开始位置
                    l = t(0).length '长度
                    If k = "P" Or k = "D" Or k = "YW" Then
                        With rn.Characters(f, l + 1)
                            .Font.ColorIndex = 3 '红
                            .Font.Bold = True
                        End With
'                        If Not Range("B41:B64").Find(what:=rn.Offset(0, -5).Value, lookat:=xlWhole) Is Nothing Then
'                            rn.Offset(0, 3).Interior.ColorIndex = 3
'                        End If
                    ElseIf k Like "*[黑罐到卸]*" Then
                        With rn.Characters(f, l + 1)
                            .Font.ColorIndex = 1 '黑
                            .Font.Bold = True
                        End With
                    End If
                End If
            Next
            
            '计算罐车的结存
            Dim 自备罐%, 路罐%, g%
            reg.Pattern = "\d+"
            On Error Resume Next
            For Each tx In Split(rn.value, " ")
                If InStr(tx, "自备罐") > 0 And tx <> "" Then
                    g = reg.Execute(tx)(0)
                    自备罐 = 自备罐 + g
                End If
                If InStr(tx, "路罐") > 0 And tx <> "" Then
                    g = reg.Execute(tx)(0)
                    路罐 = 路罐 + g
                End If
                g = 0
            Next
            On Error GoTo 0
        End If
    Next
    
    Sheet6.Range("G93").value = 自备罐 & "(自)/" & 路罐 & "(路)"
    
End Sub


Sub 读取票据(车号)
    Dim 文件名 As String
    文件名 = ThisWorkbook.path & "\票据缓存\" & 车号 & ".html"
    With UserForm3
        If Dir(文件名, vbDirectory) <> "" Then
            Do While .WebBrowser1.Busy
                DoEvents
            Loop
        
            ' 设置 WebBrowser 控件位置和大小
            With .WebBrowser1
        '        .navigate "about:blank"
                .navigate 文件名 '文件路径
                .Visible = True
                .width = 700
                .height = 550
            End With
            
            .width = 700
            .height = 550
            .Caption = "运单号码：" & 车号
            
            ' 等待 WebBrowser 控件加载完毕
            Do While .WebBrowser1.Busy Or .WebBrowser1.readyState <> 4
                DoEvents
            Loop
            
            '滑动到最底部
            On Error Resume Next
            .WebBrowser1.document.parentWindow.scrollTo 0, .WebBrowser1.document.Body.ScrollHeight
            On Error GoTo 0
    
            .Show 0
        Else
            MsgBox "查无票据：" & Sheet13.Cells(Target.Row, "A").value
        End If
        
    End With
    
    Cancel = True ' 取消单元格编辑模式
    Application.EnableEvents = True
End Sub
