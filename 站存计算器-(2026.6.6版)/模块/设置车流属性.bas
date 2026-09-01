Attribute VB_Name = "设置车流属性"
Sub 设置车流属性()
    
    '读取方向基础数据库
    Dim rang As Range
    Set rang = Sheet2.Range("A1:A" & Sheet2.Cells(Sheet2.Rows.count, 1).End(3).Row)
    
    Sheet2.Range("H:R").Delete
    
    For Each rg In rang
        d(rg.value) = rg.Offset(0, 1).value
    Next
    
    'arr = Sheet2.Range("A1").CurrentRegion
    ReDim brr(1 To UBound(arr), 1 To 10)
    
    For i = 2 To UBound(arr)
        brr(i, 1) = arr(i, 1) '股道
        
        If Mid(arr(i, 3), 1, 1) = "N" And Mid(arr(i, 4), 1, 1) = 5 Then
            brr(i, 2) = "X" '平板车
        
        ElseIf Mid(arr(i, 3), 1, 1) = "B" Then
            If arr(i, 3) = "BH1" Then
                brr(i, 2) = "P" '橡胶盖-P车
            ElseIf Mid(arr(i, 4), 1, 1) = 5 Then
                brr(i, 2) = "X" '平板车
            ElseIf Mid(arr(i, 4), 1, 1) = 6 Then
                brr(i, 2) = "G" '罐车
            Else
                brr(i, 2) = Mid(arr(i, 3), 1, 1) '车种
            End If
        Else
        
            brr(i, 2) = Mid(arr(i, 3), 1, 1) '车种
        End If
        
        brr(i, 3) = arr(i, 7) '载重
        brr(i, 4) = arr(i, 8) '到站
        brr(i, 5) = arr(i, 9) '方向
        brr(i, 6) = arr(i, 10) '品名
        brr(i, 7) = arr(i, 13) '收货人
        brr(i, 8) = arr(i, 14) '记事
        brr(i, 10) = arr(i, 4) '车号
        
        'If arr(i, 8) = "钦州港" Then
            If arr(i, 8) = "防城港" And 4 < arr(i, 7) * 1 And arr(i, 10) Like "*空*" Then '防城港排空箱和重车
                brr(i, 9) = d(arr(i, 8))
            ElseIf d.Exists(arr(i, 8)) And arr(i, 7) * 1 > 25 Then '到站交口
                If brr(i, 8) Like "*[扣修]*" Then
                    brr(i, 9) = "空车"
                Else
                    brr(i, 9) = d(arr(i, 8)) '交口
                End If
            ElseIf d.Exists(arr(i, 8)) And arr(i, 7) * 1 < 10 And arr(i, 10) = "自备" And brr(i, 2) = "G" Then '自备车(罐车，其它自备车不列入自备)
                brr(i, 9) = "自备"
            ElseIf Not d.Exists(arr(i, 8)) And arr(i, 7) * 1 > 25 Then '当车站名不存在且载重大于25t时，按方向代号识别
        '    Debug.Print arr(i, 8) & arr(i, 7)
        '    Debug.Print brr(i, 5)
                If brr(i, 5) = "3" Then '方向代号为3时为南口
                    brr(i, 9) = "南口"
                ElseIf brr(i, 5) = "2" Then '方向代号为2时为管内,6为钦州港
                    brr(i, 9) = "管内"
                Else
                    brr(i, 9) = "沙口" '其它代号为沙口
                End If
            Else
                brr(i, 9) = "空车"
            End If
        'End If
    Next
    
    'Set Dzs = Sheet2.Range("E2:E" & Sheet2.Cells(Rows.count, "E").End(3).Row) '待装WorksheetFunction.Transpose(
    On Error Resume Next  '遇到错误时继续
        For Each dzg In drr 'Dzs
        'Debug.Print dzg
            For i = 2 To UBound(brr)
                If brr(i, 1) = dzg And brr(i, 3) < 25 Then
                    If brr(i, 2) = "G" And Mid(brr(i, 10), 1, 1) = 0 Then
                            brr(i, 9) = "待装自备罐"
                        Else
                            brr(i, 9) = "待装"
                      End If
                End If
            Next
        Next
    
    'Set Dfs = Sheet2.Range("F2:F" & Sheet2.Cells(Rows.count, "F").End(3).Row) '待发WorksheetFunction.Transpose(
    If IsArray(crr) Then
        For Each dfc In crr 'Dfs
        'Debug.Print dfc
            For i = 2 To UBound(brr)
                If brr(i, 1) = CStr(dfc) Then
                    brr(i, 9) = "待发"
                     If gds.Exists(brr(i, 1)) Then
                        gds(brr(i, 1)) = gds(brr(i, 1)) + 1
                        Else
                        gds(brr(i, 1)) = 1
                        End If
                End If
            Next
        Next
    End If
    
    With Sheet2
        .Range("H1").Resize(UBound(brr), UBound(brr, 2)) = brr
        .Range("Q:Q").NumberFormatLocal = "0000000"
        With .Range("H:Q")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Columns.AutoFit
        .AutoFilter
        End With
    End With

End Sub
