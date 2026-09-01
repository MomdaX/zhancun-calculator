Attribute VB_Name = "毛玻璃板"
Sub 竖向毛玻璃板()

    Dim cnn As Object, trr(), rst As Object, i%, rs%, r%, c%, rn As Range, rng As Range, 股道 As String
    Set cnn = CreateObject("adodb.connection")

    Application.ScreenUpdating = False
    path = ThisWorkbook.FullName
    str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & path
    cnn.Open str_cnn '打开表格
    
    xrr = Sheet6.Range("B1:G92")
    With Sheet31
        .UsedRange.Clear '重置
        For r = 2 To UBound(xrr)
            If xrr(r, 6) <> "" Then
                股道 = xrr(r, 1): c = c + 1
                'strSQL = "Select 股道,车种,车号,换长,载重,到站,方向,品名,收货人,记事 FROM [站存测试文件$A3:Q] WHERE [股道] IN ('" & 股道 & "')"
                strSQL = "Select 到站&'    '&品名&'  '&车号&'   '&记事 as 合并列 FROM [站存测试文件$A3:Q] WHERE [股道] IN ('" & 股道 & "')"
                Set rst = cnn.Execute(strSQL) '查找
                
                '设置表头
                With .Cells(1, c)
                    .value = 股道 & "道"
                    .EntireColumn.ColumnWidth = 6 '列宽
                    .Font.Bold = True '标题加粗
                End With
            
                '写入查找到的数据
                .Cells(2, c).CopyFromRecordset rst
                i = .Cells(Rows.count, c).End(3).Row '尾部行
                Set rng = .Range(.Cells(2, c), .Cells(i, c))
                rst.Close
                
                .Range(.Cells(i + 1, c), .Cells(70, c)) = " " '填充空值，防止溢出
                
                '结束位置，车数/换长
                strSQL = "Select SUM(换长) as Total FROM [站存测试文件$A3:Q] WHERE [股道] IN ('" & 股道 & "')"
                Set rst = cnn.Execute(strSQL) '查找
                With .Cells(i + 1, c)
                    .Interior.ColorIndex = 15 '绿
                    .value = (i - 1) & "/" & Round(rst.fields("Total").value, 1) '辆/换长
                    .Font.Bold = True '加粗
                End With
                rst.Close
                
                '标记重车
                For Each rn In rng
                    车站 = Split(rn.value, " ")(0)
                    If rn.value Like "*[货场橡胶小麦粉面粉纯碱肥料]*" And 车站 = "钦州港" Then
                        rn.Replace "钦州港", "货场", 2
                    ElseIf rn.value Like "*永鑫*" And 车站 = "钦州港" Then
                        rn.Replace "钦州港", "永鑫", 2
                    ElseIf rn.value Like "*天盛*" And 车站 = "钦州港" Then
                        rn.Replace "钦州港", "天盛", 2
                    ElseIf rn.value Like "*航煤*" And 车站 = "钦州港" Then
                        rn.Replace "钦州港", "航煤", 2
                    ElseIf rn.value Like "*自备*" And 车站 = "钦州港" Then
                        rn.Replace "钦州港", "自备", 2
                    End If
                    
                Next
                
                显示信息.标记到站方向颜色 rng
                
            End If
            
        Next
        
        '格式
        With .UsedRange
            .Borders.LineStyle = xlContinuous '设置样式边框
            With .Rows(1)
                .Interior.ColorIndex = 14 '绿
                .RowHeight = 20 '行高
                .Font.ColorIndex = 2 '白
                .Font.Size = 13 '字体
            End With
        End With
    End With
    
    Set cnn = Nothing
    Application.ScreenUpdating = True
End Sub
