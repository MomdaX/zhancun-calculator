Attribute VB_Name = "计算"
Sub 统计()
    Dim dis As New Dictionary, dic As New Dictionary, ls As New Dictionary, dcz As New Dictionary, dsk As New Dictionary, dnk As New Dictionary
    Dim 类型 As Variant
    
    '读取方向基础数据库
    Set rang = Sheet2.Range("A1:A" & Sheet2.Cells(Sheet2.Rows.count, 1).End(3).Row)
    
    'Sheet2.Range("H:R").Delete
    
    'For Each rg In rang
    'D(rg.Value) = rg.Offset(0, 1).Value
    'Next
    
    类型合计 = Array("沙口", "南口", "管内", "待卸", "待装", "空车", "待发", "待装自备罐", "自备")
    
    dn = Sheet2.Range("H2:Q" & Sheet2.Cells(Rows.count, "Q").End(3).Row) '数据源
    
    For Each 类型 In 类型合计
    
            For i = 1 To UBound(dn)
                If dis.Exists(dn(i, 4)) And dn(i, 9) = 类型 Then '到站存在与类型符合
                
                    dis(dn(i, 4)) = dis(dn(i, 4)) + 1 '到站累计
                    
                    '车型
                    If dcz.Exists(dn(i, 2)) Then
                        dcz(dn(i, 2)) = dcz(dn(i, 2)) + 1 '车型累计
                        'Debug.Print c(brr(i, 2))
                    Else
                        dcz(dn(i, 2)) = 1
                    End If
                    
                    '沙口与南口的车站统计
                    If dic.Exists(dn(i, 4)) Then '车站统计
                        If dn(i, 9) = "沙口" Or dn(i, 9) = "南口" Or dn(i, 9) = "管内" Then
                            dic(dn(i, 4)) = dic(dn(i, 4)) + 1
                        End If
                    Else
                        If dn(i, 9) = "沙口" Or dn(i, 9) = "南口" Or dn(i, 9) = "管内" Then
                        
                            dic(dn(i, 4)) = 1
                        End If
                    End If
                    
                ElseIf Not dis.Exists(dn(i, 4)) And dn(i, 9) = 类型 Then
                
                    dis(dn(i, 4)) = 1 '到站累计
                    
                    If dcz.Exists(dn(i, 2)) Then
                        dcz(dn(i, 2)) = dcz(dn(i, 2)) + 1
                        'Debug.Print c(brr(i, 2))
                    Else
                        dcz(dn(i, 2)) = 1
                    End If
                    
                    '沙口与南口的车站统计
                    If dic.Exists(dn(i, 4)) Then '车站统计
                        If dn(i, 9) = "沙口" Or dn(i, 9) = "南口" Or dn(i, 9) = "管内" Then
                            dic(dn(i, 4)) = dic(dn(i, 4)) + 1
                        End If
                    Else
                        If dn(i, 9) = "沙口" Or dn(i, 9) = "南口" Or dn(i, 9) = "管内" Then
                        
                            dic(dn(i, 4)) = 1
                        End If
                    End If
                    
                End If
            Next
    
    If dis.count = 0 Then
        ls(类型) = 0
    Else
        ls(类型) = WorksheetFunction.Sum(dis.items)
    End If
    
    For Each 车型 In dcz.keys
        ls(类型 & "车种") = ls(类型 & "车种") & 车型 & dcz(车型) & " " '2024-8-25：修复车型为空
    Next
    站存 = 站存 + ls(类型)
    'Debug.Print 类型 & ":" & ls(类型)
    Debug.Print 类型 & "车种：" & ls(类型 & "车种")
    dis.RemoveAll '清空dis字典
    dcz.RemoveAll
    Next
    Debug.Print "站存：" & 站存
    
    ars = dic.keys
    For Each ar In ars
        If d(ar) = "沙口" Then
            沙口 = ar & dic(ar) & " " & 沙口
            '沙口车流合计 = dic(ar) + 沙口车流合计
            
        ElseIf d(ar) = "南口" Then
            南口 = ar & dic(ar) & " " & 南口
            '南口车流合计 = dic(ar) + 南口车流合计
                
        ElseIf d(ar) = "管内" Then
            管内 = ar & ":" & dic(ar) & " " & 管内
            '管内车流合计 = dic(ar) + 管内车流合计
            
            For i = 1 To UBound(dn)
        '    If dn(i, 4) = "防城港" Then
                If ar = dn(i, 4) And dn(i, 3) > 4 And dcz.Exists(dn(i, 2)) And dn(i, 9) = "管内" Then '管内车种单项合计
                    dcz(dn(i, 2)) = dcz(dn(i, 2)) + 1
                    'Debug.Print zs(dn(i, 2))
                ElseIf ar = dn(i, 4) And dn(i, 3) > 4 And dn(i, 9) = "管内" Then
                    dcz(dn(i, 2)) = 1
                End If
        '    End If
            Next
            
            For Each gr In dcz.keys
                单项合计 = gr & dcz(gr) & 单项合计
            Next
            
            管内车种合计 = ar & ":" & 单项合计 & " " & 管内车种合计
            Set dcz = Nothing
            单项合计 = ""
                  
        End If
    Next
    
    For Each ar In gds.keys
    'Debug.Print gds(ar)
        待发股道合计 = ar & "道/" & gds(ar) & " " & 待发股道合计
    Next
    
    Debug.Print "沙口:" & 沙口
    Debug.Print "南口:" & 南口
    Debug.Print "管内:" & 管内
    
    '输出到sheet1中
    Sheet1.Range("B2").value = 沙口
    Sheet1.Range("B3").value = ls("沙口车种")
    Sheet1.Range("H2").value = ls("沙口")
    
    Sheet1.Range("B4").value = 南口
    Sheet1.Range("B5").value = ls("南口车种")
    Sheet1.Range("H4").value = ls("南口")
    
    Sheet1.Range("B6").value = 管内车种合计
    Sheet1.Range("H6").value = ls("管内")
    
    Sheet1.Range("B7").value = ls("待卸车种")
    Sheet1.Range("H7").value = ls("待卸")
    
    Sheet1.Range("E9").value = "G"
    待装车种合计 = Join(VBA.Filter(Split(ls("待装车种"), " "), "G", 0), " ")
    待装路罐 = Mid(Join(VBA.Filter(Split(ls("待装车种"), " "), "G", 1), ""), 2)
    Sheet1.Range("B8").value = 待装车种合计
    
    If ls("待装自备罐车种") = "" Then
        Sheet1.Range("H8").value = ls("待装")
        Sheet1.Range("F8").value = 0
    Else
        Sheet1.Range("H8").value = ls("待装") + Split(ls("待装自备罐车种"), "G")(1)
        Sheet1.Range("F8").value = Split(ls("待装自备罐车种"), "G")(1)
    End If
    
    If 待装路罐 = "" Then
        Sheet1.Range("F10").value = 0
    Else
        Sheet1.Range("F10").value = 待装路罐
    End If
    
    Sheet1.Range("E13").value = "G"
    空车车种合计 = Join(VBA.Filter(Split(ls("空车车种"), " "), "G", 0), " ")
    空车路罐 = Mid(Join(VBA.Filter(Split(ls("空车车种"), " "), "G", 1), ""), 2)
    Sheet1.Range("B12").value = 空车车种合计
    
    If ls("自备车种") = "" Then
        Sheet1.Range("H12").value = ls("空车")
        Sheet1.Range("F12").value = 0
    Else
        Sheet1.Range("H12").value = ls("空车") + Split(ls("自备车种"), "G")(1)
        Sheet1.Range("F12").value = Split(ls("自备车种"), "G")(1)
    End If
    
    If 空车路罐 = "" Then
        Sheet1.Range("F14").value = 0
    Else
        Sheet1.Range("F14").value = 空车路罐
    End If
    
    Sheet1.Range("B16").value = 待发股道合计
    Sheet1.Range("B17").value = ls("待发车种")
    Sheet1.Range("H16").value = ls("待发")
    
    '总数合计
    Dim zs As New Dictionary
        For i = 1 To UBound(dn)
            If zs.Exists(dn(i, 2)) Then '总数
                zs(dn(i, 2)) = zs(dn(i, 2)) + 1
                'Debug.Print zs(brr(i, 2))
                Else
                zs(dn(i, 2)) = 1
            End If
        Next
        For Each ar In zs.keys
        'Debug.Print zs(ar)
            总数车种合计 = ar & zs(ar) & " " & 总数车种合计
        Next
    Sheet1.Range("B18").value = 总数车种合计
    Sheet1.Range("H18").value = 站存
    
    '清空数组
    brr = ""
    crr = ""
    dn = ""
    ars = ""
    类型合计 = ""
    
    '清空字典
    Set gds = Nothing
    Set zs = Nothing
    Set d = Nothing
    Set dis = Nothing
    Set dic = Nothing
    Set ls = Nothing
    Set rang = Nothing
    
    '清空常量
    i = 0
    j = 0
    

End Sub


Sub 待卸纠正()
Dim dx, d1, d2, k(0 To 2), d As New Dictionary, cz
With Sheet1
    dx = .Range("K7:M7")
    If WorksheetFunction.CountA(.Range("K7:M7")) > 0 Then
        d1 = Split(.Range("B7"), " ")
        d2 = Split(.Range("B12"), " ")
        
        k(0) = .Range("K7").value
        k(1) = .Range("L7").value
        k(2) = .Range("M7").value
        
        For i = 0 To UBound(d1) - 1
            cx = 提取(d1(i), "[^\d]+")(0)
            cz = 提取(d1(i), "\d+")(0)
            d(cx) = k(i) - cz
            
            A = A & " " & cx & k(i)
        Next
        For i = 0 To UBound(d2) - 1
            cx = 提取(d2(i), "[^\d]+")(0)
            cz = 提取(d2(i), "\d+")(0)
            b = b & " " & cx & (cz - d(cx))
        Next
        .Range("B7").value = Trim(A)
        .Range("H7").value = .Range("H7").value + WorksheetFunction.Sum(d.items)
        .Range("B12").value = Trim(b)
        .Range("H12").value = .Range("H12").value - WorksheetFunction.Sum(d.items)
    End If
End With

End Sub
Function 提取(HTML, 表达式 As String)
Dim reg As New RegExp
With reg
    .Pattern = 表达式 '"""([\d_]+)"""
    .Global = True
    .MultiLine = True
End With

Set 提取 = reg.Execute(HTML)

End Function
