Attribute VB_Name = "统计股道存车"
Sub 选取文件()
    Dim filePath As String
    Application.ScreenUpdating = False '关闭屏幕刷新
    Application.DisplayAlerts = False '关闭警告界面
    '打开文件对话框
    With Application.fileDialog(msoFileDialogOpen)
        .InitialFileName = "C:\Program Files\SMIS2.6\TranData"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls*,*.xls*"
        .AllowMultiSelect = False
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        End If
    End With

    '如果没有选择文件，路径为空时，结束程序
    If filePath = "" Then
        Exit Sub
    Else
        '股道存车 filePath
        Dim wb As Workbook
        Dim ws As Worksheet
        Set wb = GetObject(filePath)
        Set ws = wb.ActiveSheet 's("第1页")
        xrr = ws.Range("A1").CurrentRegion.value
        
        '复制到测试表
        If Sheet7.AutoFilterMode Then Sheet7.AutoFilterMode = False '取消筛选状态
        Sheet7.Cells.Clear
        ws.Range("A1").CurrentRegion.Copy Sheet7.Range("A1")
        
        ' 关闭工作簿
        wb.Close False
        Set wb = Nothing
    
        股道存车 xrr
    End If
    Application.ScreenUpdating = False '关闭屏幕刷新
    Application.DisplayAlerts = False '关闭警告界面
End Sub

Sub 股道存车(xrr)
    '前期绑定
    Dim d As New Dictionary, d_老牌车 As New Dictionary, d_到站 As New Dictionary, d_方向 As New Dictionary, d_车种 As New Dictionary, d_车次 As New Dictionary
    Dim cz
    
    '读取车站方向数据
    srr = Sheet2.Range("A1:C2").CurrentRegion.value
    
    For i = 4 To UBound(xrr)
        '载重 and 到站
        If xrr(i, 7) = "" And xrr(i, 8) <> "" Then xrr(i, 8) = ""
    '    If xrr(i, 1) = 15 Then'查BUG用
    '        If xrr(i, 2) = 42 Then Stop
    '    End If
    
        '未做完系统的车，按备注栏标记
        '载重*1<15t and (到站="" or 钦州港) and 记事栏=""
        If xrr(i, 7) * 1 < 15 And (xrr(i, 8) = "" Or xrr(i, 8) Like "*[钦州港]*") And xrr(i, 14) <> "" Then
            For k = 2 To UBound(srr)
                If InStr(xrr(i, 14), srr(k, 1)) > 0 Then '优先到站
                    If InStr(xrr(i, 14), srr(k, 1) & "循环") = 0 And InStr(xrr(i, 14), "卸空后返回" & srr(k, 1)) = 0 Then
                        xrr(i, 8) = srr(k, 1)
                        Sheet7.Cells(i, 8) = srr(k, 1)
                    Else
                        If Left(xrr(i, 3), 1) = "G" And Left(xrr(i, 4), 1) = "6" Then
                            xrr(i, 8) = "路罐"
                        ElseIf Left(xrr(i, 3), 1) = "G" And Left(xrr(i, 4), 1) = "0" Then
                            If Mid(xrr(i, 4), 2, 1) = 7 Then
                                xrr(i, 8) = "黑罐"
                            Else
                                xrr(i, 8) = "自备罐"
                            End If
                        Else
                            xrr(i, 8) = 正则_车种(xrr(i, 3)) '空车
                        End If
                    End If
                    Exit For '匹配到车站后结束查找
                ElseIf k = UBound(srr) Then '其次品名
                    If Left(xrr(i, 3), 1) = "G" And Left(xrr(i, 4), 1) = "6" Then
                        If InStr(xrr(i, 14), "原装") = 0 Then
                            xrr(i, 8) = IIf(InStr(xrr(i, 14), "汽油") > 0, "汽油", IIf(InStr(xrr(i, 14), "柴油") > 0, "柴油", "路罐")) '路罐
                        Else
                            xrr(i, 8) = "路罐"
                        End If
                    ElseIf Left(xrr(i, 3), 1) = "G" And Left(xrr(i, 4), 1) = "0" Then
                        If Mid(xrr(i, 4), 2, 1) = 7 Then
                            xrr(i, 8) = "黑罐"
                        Else
                            If InStr(xrr(i, 14), "原装") = 0 Then
                                xrr(i, 8) = IIf(InStr(xrr(i, 14), "汽油") > 0, "汽油", IIf(InStr(xrr(i, 14), "柴油") > 0, "柴油", "自备罐")) '自备罐
                            Else
                                xrr(i, 8) = "自备罐"
                            End If
                        End If
                    Else
    '                    xrr(i, 8) = "空车"
                        xrr(i, 8) = 正则_车种(xrr(i, 3)) '空车
                    End If
                End If
            Next
        ElseIf xrr(i, 7) * 1 > 15 And xrr(i, 8) = "钦州港" Then '到重车
            xrr(i, 8) = "到卸"
        ElseIf xrr(i, 7) = "" And xrr(i, 8) = "" And xrr(i, 14) = "" Then '重量、到站、记事栏都为空时
            If Left(xrr(i, 3), 1) = "G" And Left(xrr(i, 4), 1) = "6" Then
                xrr(i, 8) = "路罐"
            ElseIf Left(xrr(i, 3), 1) = "G" And Left(xrr(i, 4), 1) = "0" Then
                If Mid(xrr(i, 4), 2, 1) = 7 Then
                    xrr(i, 8) = "黑罐"
                Else
                    xrr(i, 8) = "自备罐"
                End If
            Else
    '            xrr(i, 8) = "空车"
                xrr(i, 8) = 正则_车种(xrr(i, 3)) '空车
            End If
        End If
    
        '提取车型
        车型 = 正则_车种(xrr(i, 3))
        If d.Exists(xrr(i, 1)) Then 'd.E(股道)
            辆数 = 辆数 + 1
            '换长
            换长 = 换长 + val(xrr(i, 6))
            '载重
            载重 = 载重 + val(xrr(i, 7)) + val(xrr(i, 5))
            
            '方向分类
            If xrr(i, 9) = "3" And InStr(d_方向(xrr(i, 1)), "南口") = 0 Then
                d_方向(xrr(i, 1)) = d_方向(xrr(i, 1)) & "南口" & Chr(10)
            ElseIf xrr(i, 9) = "2" And InStr(d_方向(xrr(i, 1)), "管内") = 0 Then
                d_方向(xrr(i, 1)) = d_方向(xrr(i, 1)) & "管内" & Chr(10)
            ElseIf xrr(i, 9) = "6" And xrr(i, 7) > 20 And InStr(d_方向(xrr(i, 1)), "到卸") = 0 Then
                d_方向(xrr(i, 1)) = d_方向(xrr(i, 1)) & "到卸" & Chr(10)
            Else
                If InStr("236", xrr(i, 9)) = 0 And InStr(d_方向(xrr(i, 1)), "沙口") = 0 Then d_方向(xrr(i, 1)) = d_方向(xrr(i, 1)) & "沙口" & Chr(10)
            End If
            
            '识别到站
            If d_到站.Exists(xrr(i, 8)) Then
                d_到站(xrr(i, 8)) = d_到站(xrr(i, 8)) + 1
            Else
                d_到站(xrr(i, 8)) = 1
            End If
            到站 = ""
            For Each dz In d_到站.keys
'                Set 直达 = Sheet2.Range("A1").CurrentRegion.Find(dz)
'                If Not 直达 Is Nothing Then
'                    If 直达.Offset(0, 3).value <> "" Then
'                        到站 = 到站 & " " & dz & d_到站(dz) & "(" & 直达.Offset(0, 3).value & ")"
'                    Else
'                        到站 = 到站 & " " & dz & d_到站(dz)
'                    End If
'                Else
                    到站 = 到站 & " " & dz & d_到站(dz)
'                End If
            Next
            
            '识别车种
            If d_车种.Exists(车型) Then
                d_车种(车型) = d_车种(车型) + 1
            Else
                d_车种(车型) = 1
            End If
            车种 = ""
            For Each cz In d_车种.keys
                车种 = 车种 & " " & cz & d_车种(cz)
            Next
            
            '老牌车/大点车
            If d_老牌车.Exists(xrr(i, 1)) Then
                If DateDiff("h", xrr(i, 16), Now()) > 47 And Left(xrr(i, 4), 1) <> 0 Then d_老牌车(xrr(i, 1)) = d_老牌车(xrr(i, 1)) + 1 '老牌车/大点车：DateDiff("h",range("P4"),now())
            Else
                If DateDiff("h", xrr(i, 16), Now()) > 47 And Left(xrr(i, 4), 1) <> 0 Then d_老牌车(xrr(i, 1)) = 1
            End If
            
            '识别车次
            If d_车次.Exists(xrr(i, 15)) Then
                d_车次(xrr(i, 15)) = d_车次(xrr(i, 15)) + 1
            Else
                d_车次(xrr(i, 15)) = 1
            End If
            车次 = ""
            If d_车次.count = 1 Then
                车次 = xrr(i, 15)
            ElseIf d_车次.count = 0 Then
                车次 = ""
            Else
                maxVal = 0
                For Each cc In d_车次.keys
                    If d_车次(cc) > maxVal Then
                        maxVal = d_车次(cc)
                        车次 = cc
                    End If
                Next
            End If
            
            '存入字典中
    '        d(xrr(i, 1)) = 辆数 & "-" & 换长 & "-" & d_方向(xrr(i, 1)) & "-" & 到站 & "-" & 车种 & "-" & 车次 & "-" & 载重
            d(xrr(i, 1)) = Array(辆数, 换长, d_方向(xrr(i, 1)), 到站, 车种, 车次, 载重, d_老牌车(xrr(i, 1)))
        Else
            辆数 = 1
            换长 = val(xrr(i, 6))
            载重 = val(xrr(i, 7)) + val(xrr(i, 5))
            
            '方向分类
            If xrr(i, 9) = "3" And InStr(d_方向(xrr(i, 1)), "南口") = 0 Then
                d_方向(xrr(i, 1)) = d_方向(xrr(i, 1)) & "南口" & Chr(10)
            ElseIf xrr(i, 9) = "2" And InStr(d_方向(xrr(i, 1)), "管内") = 0 Then
                d_方向(xrr(i, 1)) = d_方向(xrr(i, 1)) & "管内" & Chr(10)
            ElseIf xrr(i, 9) = "6" And xrr(i, 7) > 20 And InStr(d_方向(xrr(i, 1)), "到卸") = 0 Then
                d_方向(xrr(i, 1)) = d_方向(xrr(i, 1)) & "到卸" & Chr(10)
            Else
                If InStr("236", xrr(i, 9)) = 0 And InStr(d_方向(xrr(i, 1)), "沙口") = 0 Then d_方向(xrr(i, 1)) = d_方向(xrr(i, 1)) & "沙口" & Chr(10)
            End If
            
            '识别到站
            d_到站.RemoveAll
            d_到站(xrr(i, 8)) = 1
            到站 = ""
            For Each dz In d_到站.keys
                Set 直达 = Sheet2.Range("A1").CurrentRegion.Find(dz)
                If Not 直达 Is Nothing Then
                    If 直达.Offset(0, 3).value <> "" Then
                        到站 = 到站 & " " & dz & d_到站(dz) & "(" & 直达.Offset(0, 3).value & ")"
                    Else
                        到站 = 到站 & " " & dz & d_到站(dz)
                    End If
                Else
                    到站 = 到站 & " " & dz & d_到站(dz)
                End If
            Next
            
            '识别车种
            d_车种.RemoveAll
            d_车种(车型) = 1
            车种 = ""
            For Each cz In d_车种.keys
                车种 = 车种 & " " & cz & d_车种(cz)
            Next
            
            '识别车次
            d_车次.RemoveAll
            d_车次(xrr(i, 15)) = 1
            车次 = ""
            If d_车次.count = 1 Then 车次 = xrr(i, 15)
            
            '老牌车/大点车：DateDiff("h",range("P4"),now())
            If DateDiff("h", xrr(i, 16), Now()) > 47 And Left(val(xrr(i, 4)), 1) <> 0 Then d_老牌车(xrr(i, 1)) = 1
            
            '存入字典中
    '        d(xrr(i, 1)) = 辆数 & "-" & 换长 & "-" & d_方向(xrr(i, 1)) & "-" & 到站 & "-" & 车种 & "-" & 车次 & "-" & 载重
            d(xrr(i, 1)) = Array(辆数, 换长, d_方向(xrr(i, 1)), 到站, 车种, 车次, 载重, d_老牌车(xrr(i, 1)))
        End If
        
        '计算每个车的停时，超过47小时的标记为橙色
        With Sheet7.Cells(i, 17)
            .value = DateDiff("h", xrr(i, 16), Now())
    '        If Val(.value) > 47 Then
    '            .Interior.ColorIndex = 19
    '        End If
        End With
    Next
    Set B1 = Sheet6.Cells(2, 2)
    Set B2 = Sheet6.Cells(Sheet6.Rows.count, 2).End(3)
    For Each b In Range(B1, B2)
        If Not IsEmpty(d(CStr(b))) Then
            b.Offset(0, 2) = d(CStr(b))(0) '辆数
            b.Offset(0, 4) = Round(d(CStr(b))(1), 1) '换长(保留1位小数)
            b.Offset(0, 5) = d(CStr(b))(3) '到站
            
            If Len(d(CStr(b))(2)) > 0 Then
                b.Offset(0, 1) = Left(d(CStr(b))(2), Len(d(CStr(b))(2)) - 1)
            Else
                b.Offset(0, 1) = d(CStr(b))(2) '方向
            End If
            b.Offset(0, 3) = d(CStr(b))(4) '车型
            b.Offset(0, 8) = d(CStr(b))(5) '车次
            If Left(b.Offset(0, 8), 1) = 6 Then b.Offset(0, 8).Interior.ColorIndex = 40 '循环车标色
            If d(CStr(b))(6) > 0 Then b.Offset(0, 9) = Round(d(CStr(b))(6), 1) '载重(保留1位小数)
            If d(CStr(b))(7) > 0 Then b.Offset(0, 10) = d(CStr(b))(7) '老牌车
         End If
    Next
    
    '设置数据验证信息（方向）
    Call 显示信息.到站方向
    Dim rng As Range
    Set rng = Sheet6.Range("G2:G92")
    Call 显示信息.标记到站方向颜色(rng)
    
    'Call 毛玻璃板.竖向毛玻璃板
    
    '合计
    Sheet6.Range("D93").value = WorksheetFunction.Sum(Sheet6.Range("D2:D92")) '车数
    Sheet6.Range("L93").value = WorksheetFunction.Sum(Sheet6.Range("L2:L92")) '老牌车
    
    Application.ScreenUpdating = True '开启屏幕刷新
    Application.DisplayAlerts = True '开启警告界面
End Sub

Function 正则_车种(车种)
    Dim reg As New RegExp
    With reg
        .Global = 1
        .MultiLine = 1
        .Pattern = "[^\d]+"
    End With
    Set 类型 = reg.Execute(车种)(0)
    For Each ls In Array("C", "X", "P", "G", "YW", "T", "B", "D", "K")
    '    If 类型 Like "*" & ls & "*" Then
        If InStr(类型, ls) > 0 Then
            If Mid(类型, 1, 1) = "N" And Mid(类型, 2, 1) = "X" Then
                正则_车种 = "X"
            Else
    '        If ls = "X" Then Stop
                正则_车种 = Mid(类型, 1, 1)
            End If
            Exit For
        Else
            正则_车种 = 类型
    '        Exit For
        End If
    Next
End Function



Sub 筛选_沙口()
    If Sheet6.AutoFilterMode Then
        Sheet6.AutoFilterMode = False
    End If
    'Sheet6.Range("A1").CurrentRegion.AutoFilter 3, "*沙口*"
    Sheet6.Range("A1").CurrentRegion.AutoFilter Field:=7, Criteria1:=RGB(0, 102, 204), Operator:=xlFilterFontColor
End Sub

Sub 筛选_南口()
    If Sheet6.AutoFilterMode Then
        Sheet6.AutoFilterMode = False
    End If
    'Sheet6.Range("A1").CurrentRegion.AutoFilter 3, "*南口*"
    Sheet6.Range("A1").CurrentRegion.AutoFilter Field:=7, Criteria1:=RGB(255, 102, 0), Operator:=xlFilterFontColor
End Sub

Sub 筛选_管内()
    If Sheet6.AutoFilterMode Then
        Sheet6.AutoFilterMode = False
    End If
    Sheet6.Range("A1").CurrentRegion.AutoFilter 3, "*管内*"
End Sub

Sub 筛选_到卸()
    If Sheet6.AutoFilterMode Then
        Sheet6.AutoFilterMode = False
    End If
    Sheet6.Range("A1").CurrentRegion.AutoFilter 3, "*到卸*"
End Sub

Sub 筛选_空车()
    If Sheet6.AutoFilterMode Then
        Sheet6.AutoFilterMode = False
    End If
    
    Sheet6.Range("A1").CurrentRegion.AutoFilter 7, "*空车*"
End Sub

Sub 取消筛选()
    If Sheet6.AutoFilterMode Then
        Sheet6.AutoFilterMode = False
    End If
End Sub

Sub 筛选_昆明及远()
    If Sheet6.AutoFilterMode Then
        Sheet6.AutoFilterMode = False
    End If
    Sheet6.Range("A1").CurrentRegion.AutoFilter 7, "*昆明及远*"
End Sub

Sub 筛选_麻口()
    If Sheet6.AutoFilterMode Then
        Sheet6.AutoFilterMode = False
    End If
    Sheet6.Range("A1").CurrentRegion.AutoFilter 7, "*麻口*"
End Sub

Sub 筛选_牙口()
    If Sheet6.AutoFilterMode Then
        Sheet6.AutoFilterMode = False
    End If
    Sheet6.Range("A1").CurrentRegion.AutoFilter 7, "*牙口*"
End Sub

Sub 筛选_永口()
    If Sheet6.AutoFilterMode Then
        Sheet6.AutoFilterMode = False
    End If
    Sheet6.Range("A1").CurrentRegion.AutoFilter 7, "*永口*"
End Sub

Sub 筛选_防城港()
    If Sheet6.AutoFilterMode Then
        Sheet6.AutoFilterMode = False
    End If
    Sheet6.Range("A1").CurrentRegion.AutoFilter 7, "*防城*"
End Sub

Sub 筛选_港东()
    If Sheet6.AutoFilterMode Then
        Sheet6.AutoFilterMode = False
    End If
    
    Sheet6.Range("A1").CurrentRegion.AutoFilter 7, "*港东*"
End Sub
Sub 筛选_钦州港()
    If Sheet6.AutoFilterMode Then
        Sheet6.AutoFilterMode = False
    End If
    
    Sheet6.Range("A1").CurrentRegion.AutoFilter 7, "*到卸*"
    'Sheet7.Range("A3:Q3").AutoFilter 1, Array("G1", "G2"), xlFilterValues
End Sub

Sub 重空车按车种分类(查询 As String)
    Dim d As New Dictionary
    Dim reg As New RegExp
    
    Call Sheet6_清空重空车
    
    With reg
        .Global = 1
        .MultiLine = 1
        .Pattern = "[^\d]+"
    End With
    
    zrr = Sheet7.Range("A1").CurrentRegion '读取车站站存文件
    srr = Sheet2.Range("A1:C2").CurrentRegion '读取车站方向数据

    For i = 4 To UBound(zrr)
        股道 = zrr(i, 1)
        车号 = zrr(i, 4)
        载重 = zrr(i, 7)
        到站 = zrr(i, 8)
        记事 = zrr(i, 14)
        If zrr(i, 7) = "" And zrr(i, 8) = "" And zrr(i, 14) <> "" Then
            For k = 2 To UBound(srr)
    '            If zrr(i, 14) Like "*" & srr(k, 1) & "*" Then
                If InStr(zrr(i, 14), srr(k, 1)) > 0 Then
                    zrr(i, 8) = srr(k, 1)
                    Sheet7.Cells(i, 8) = srr(k, 1)
                    Exit For
                End If
            Next
        End If
        
        For Each 车种 In Array("C", "X", "P", "G", "YW")
            If reg.Execute(zrr(i, 3))(0) Like "*" & 车种 & "*" Then
                If zrr(i, 14) Like "*箱*" And zrr(i, 7) * 1 < 10 Then
                    d(股道 & 车种 & "装箱") = d(股道 & 车种 & "装箱") + 1
                ElseIf zrr(i, 10) Like "*自备*" And zrr(i, 7) = "" Then
                    d(股道 & 车种 & "自备") = d(股道 & 车种 & "自备") + 1
                ElseIf zrr(i, 8) <> "" Then
                    If val(zrr(i, 7)) > 20 Then '1、载重和到站都有时
                        d(股道 & 车种 & "重车") = d(股道 & 车种 & "重车") + 1
                    Else
                        For k = 2 To UBound(srr)
                            If InStr(zrr(i, 14), srr(k, 1)) > 0 Then
                                If zrr(i, 8) = srr(k, 1) Then '2、到站和记事栏的到站相同时
                                    d(股道 & 车种 & "重车") = d(股道 & 车种 & "重车") + 1
                                    Exit For
                                ElseIf k = UBound(srr) Then '3、否则就是空车
                                    d(股道 & 车种 & "空车") = d(股道 & 车种 & "空车") + 1
                                End If
                            End If
                        Next
                    End If
                Else
                    d(股道 & 车种 & "空车") = d(股道 & 车种 & "空车") + 1
                End If
            End If
        Next
    Next
    
    With Sheet6
        Set B1 = .Cells(2, 2)
        Set B2 = .Cells(.Rows.count, 2).End(3)
        
        For Each b In .Range(B1, B2)
        
            If d(b & 查询 & "装箱") > 0 Then
                b.Offset(0, 6) = 查询 & d(b & 查询 & "空车") + d(b & 查询 & "装箱") & Chr(10) & "(装箱:" & d(b & 查询 & "装箱") & ")"
        '    ElseIf d(B & 查询 & "空车") > 0 Then
            Else
                If d(b & 查询 & "自备") > 0 Then
        '            B.Offset(0, 6) = 查询 & d(B & 查询 & "空车") + d(B & 查询 & "自备") & Chr(10) & "(自备:" & 查询 & d(B & 查询 & "自备") & ")"
                    If d(b & 查询 & "空车") > 0 Then
                        b.Offset(0, 6) = 查询 & d(b & 查询 & "自备") & "(自)" & Chr(10) & 查询 & d(b & 查询 & "空车") & "(路)"
                    Else
                        b.Offset(0, 6) = 查询 & d(b & 查询 & "自备") & "(自)"
                    End If
                ElseIf d(b & 查询 & "空车") > 0 Then
                    If 查询 = "G" Then
                        b.Offset(0, 6) = 查询 & d(b & 查询 & "空车") & "(路)"
                    Else
                        b.Offset(0, 6) = 查询 & d(b & 查询 & "空车")
                    End If
                End If
            End If
            
            If d(b & 查询 & "重车") > 0 Then
                b.Offset(0, 7) = 查询 & d(b & 查询 & "重车")
                
            End If
            
            合计_空车 = 合计_空车 + d(b & 查询 & "空车") + d(b & 查询 & "自备") ' + d(B & 查询 & "装箱")
            合计_重车 = 合计_重车 + d(b & 查询 & "重车")
            合计_装箱 = 合计_装箱 + d(b & 查询 & "装箱")
        Next
        
        If 合计_装箱 > 0 Then
            .Range("H93").value = 合计_装箱 & "/" & 合计_空车
            .Range("I93").value = 合计_重车
        Else
            .Range("H93").value = 合计_空车
            .Range("I93").value = 合计_重车
        End If
    End With
End Sub
