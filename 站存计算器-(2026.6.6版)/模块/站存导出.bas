Attribute VB_Name = "站存导出"
Sub 启动发送文件窗口()
    UserForm7.Label14.Caption = "Momda：欢迎使用！！！"
    UserForm7.Show 0
End Sub

Function 导出到桌面()
Attribute 导出到桌面.VB_ProcData.VB_Invoke_Func = "2"

    Dim wb As Workbook, ws As Worksheet, rng As Range, rn As Range, path$, arr
    
    fileName = Replace(Sheet1.Range("A1").value, " ", "")
    path = "C:\Users\Administrator\Desktop\" & fileName & ".xls"
    Set wb = Workbooks.Add
    wb.Windows(1).Visible = False
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ws = wb.ActiveSheet
    Set rng = Sheet1.Range("A1:H18")
    rng.Copy
    Set rn = ws.Range("A1:H18")
    rn.PasteSpecial xlPasteAll
    rn.PasteSpecial xlPasteColumnWidths
    
    Application.CutCopyMode = False
    
    For i = 1 To rng.Rows.count
        ws.Rows(i).RowHeight = Sheet1.Rows(i).RowHeight
    Next
    
    ws.Cells.Interior.Pattern = xlNone
    wb.Windows(1).Visible = True
    ws.SaveAs path, xlExcel8
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True '开启警告界面
    
    wb.Close
    
    
    导出到桌面 = path '将文件路径传出去
End Function

Function 车流表导出()
    Dim wb As Workbook, ws As Worksheet, rng As Range, rn As Range, path$, arr
    
    path = "C:\Users\Administrator\Desktop\钦州港车流计划表-" & Date & ".xls"
    Set wb = Workbooks.Add
    wb.Windows(1).Visible = False
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set ws = wb.ActiveSheet
    Sheet22.Range("A1:E21").Copy ws.Range("A1:E21")
    Application.CutCopyMode = False
    
    ws.Cells.Interior.Pattern = xlNone
    ws.Columns("A:B").ColumnWidth = 10
    ws.Columns("C").ColumnWidth = 15
    ws.Columns("D:E").ColumnWidth = 20
    
    Set rng = ws.UsedRange.Columns(4)
    For Each rn In rng.Cells
        If rn.Row < 3 Or rn.Row = 3 Then
            rn.RowHeight = 30
        Else
            If rn <> "" Then
                If Len(rn.value) > 25 Then
                    rn.RowHeight = 60
                Else
                    rn.RowHeight = 25
                End If
            Else
                rn.RowHeight = 15
            End If
        End If
    Next
    
    wb.Windows(1).Visible = True
    'ws.SaveAs path, xlOpenXMLWorkbook
    ws.SaveAs path, xlExcel8
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True '开启警告界面
    
    wb.Close
    
    车流表导出 = path '将文件路径传出去
End Function
