Attribute VB_Name = "Excel转Word选择文件夹"
Sub 职培题库Excel转Word选文件夹()
    Dim fileDialog As Object, wb As Workbook, ws As Worksheet
    Dim wordApp As Word.Application, wdcx As Object
    Set fileDialog = Application.fileDialog(msoFileDialogFolderPicker)
    
    Application.ScreenUpdating = False '关闭屏幕刷新
    Application.DisplayAlerts = False '关闭警告界面
    
    If fileDialog.Show = -1 Then
    '获取所选路径
    selectedPath = fileDialog.SelectedItems(1)

    fipath = selectedPath & "\"
    End If
    
    '释放对象
    Set fileDialog = Nothing
    
    Set wordApp = New Word.Application
    Set wdcx = wordApp.Documents.Add
    'wdcx.Visible = True
    
    file = Dir(fipath & "*.*")
    Do While file <> ""
        ' 打开工作簿并将数据装入数组中
        Set wb = Workbooks.Open(fipath & file)
        'Set wb = GetObject(fipath & file)
        Set ws = wb.ActiveSheet
        wb.Visible = False
        
        '排序-升
        With ws
            Set rng = .Range("A2:T2")
            If ws.AutoFilterMode Then ws.AutoFilterMode = False
            rng.AutoFilter
            .AutoFilter.Sort.SortFields.Clear
            .AutoFilter.Sort.SortFields.Add key:=rng.Cells(4), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            With .AutoFilter.Sort
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End With
        
        trr = ws.Range("A2").CurrentRegion.value
        rs = ws.Cells(Rows.count, "C").End(3).Row
        文件名 = wb.name
        wb.Close False
        
        '清空文档内容
        wdcx.content.Delete
    
        For i = 3 To rs
        
        Wstr = Wstr & i - 2 & "、" & trr(i, 3)
        Wstr = Wstr & vbCrLf
            For k = 13 To 23
                If trr(i, k) <> "" Then
                Wstr = Wstr & Right(trr(2, k), 1) & "." & trr(i, k)
                Wstr = Wstr & vbCrLf
                End If
            Next
            
        Wstr = Wstr & "正确答案：" & trr(i, 7)
        Wstr = Wstr & vbCrLf
        Wstr = Wstr & vbCrLf
        Next
        
    'Excel职培题库
    '    For i = 3 To ws.Range("C3").CurrentRegion.Rows.count
    '        If ws.Range("D" & i).Value = "判断题" Then
    '            Wstr = Wstr & ws.Range("A" & i).Value & "、" & ws.Range("C" & i).Value
    '            Wstr = Wstr & vbCrLf
    '            If ws.Range("G" & i).Value = "A" Then
    '                Wstr = Wstr & "正确答案" & "：" & "对"
    '                Wstr = Wstr & vbCrLf
    '                Wstr = Wstr & vbCrLf
    '            ElseIf ws.Range("G" & i).Value = "B" Then
    '                Wstr = Wstr & "正确答案" & "：" & "错"
    '                Wstr = Wstr & vbCrLf
    '                Wstr = Wstr & vbCrLf
    '            End If
    '        ElseIf ws.Range("D" & i).Value = "单选题" Or ws.Range("D" & i).Value = "多选题" Then
    '            Wstr = Wstr & ws.Range("A" & i).Value & "、" & ws.Range("C" & i).Value
    '            Wstr = Wstr & vbCrLf
    '
    '        arr = Array("A", "B", "C", "D", "E", "F", "G")
    '        For k = 0 To UBound(arr)
    '            If ws.Cells(i, 13 + k).Value = "" Then: Exit For
    '                Wstr = Wstr & arr(k) & "、" & ws.Cells(i, 13 + k).Value
    '                Wstr = Wstr & vbCrLf
    '            Next
    '            Wstr = Wstr & "正确答案" & "：" & ws.Range("G" & i).Value
    '            Wstr = Wstr & vbCrLf
    '            Wstr = Wstr & vbCrLf
    '        ElseIf ws.Range("D" & i).Value <> "" Then
    '            Wstr = Wstr & ws.Range("A" & i).Value & "、" & ws.Range("C" & i).Value
    '            Wstr = Wstr & vbCrLf
    '            Wstr = Wstr & "正确答案" & "：" & ws.Range("G" & i).Value
    '            Wstr = Wstr & vbCrLf
    '            Wstr = Wstr & vbCrLf
    '
    '        End If
    '    Next
        
        '写入Word文档中
        wdcx.content.text = Wstr
    '    wordApp.Visible = True
        wdcx.SaveAs ThisWorkbook.path & "\" & Split(文件名, ".")(0) & ".docx"
    '    wb.Close False
        
        Wstr = ""
        file = Dir
    Loop
    
    格式处理 wordApp.ActiveDocument
    wdcx.Close False
    wordApp.Quit

    Application.ScreenUpdating = True '开启屏幕刷新
    Application.DisplayAlerts = True '开启警告界面
        
    MsgBox "转换完成！！", 64, "钦州港运转 - Momda"
End Sub
Sub Excel题库转到Word中()
    Dim wb As Workbook, ws As Worksheet, 选项 As String, path As String
    Dim wordApp As Word.Application, wdcx As Object
'
    Application.ScreenUpdating = False '关闭屏幕刷新
    Application.DisplayAlerts = False '关闭警告界面
    
    Set wordApp = New Word.Application
    Set wdcx = wordApp.Documents.Add

'    wdcx.Visible = True
    ' 打开工作簿并将数据装入数组中
    Set ws = Sheet21 'ActiveSheet
    
    '排序-升
    With ws
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
        Set rng = .Range("A1:L1")
        rng.AutoFilter
        .AutoFilter.Sort.SortFields.Clear
        .AutoFilter.Sort.SortFields.Add key:=rng.Cells(2), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With .AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
    trr = ws.UsedRange.CurrentRegion
    文件名 = ws.Range("O1")

    '清空文档内容
    wdcx.content.Delete
    
    For i = 2 To UBound(trr, 1)
    
    Wstr = Wstr & i - 1 & "、" & trr(i, 3) '题目
    Wstr = Wstr & vbCrLf '换行
        For k = 5 To 12 '选项
            If trr(i, k) <> "" Then
            选项 = Trim(trr(i, k))
            If Left(选项, 1) = "." Then 选项 = "0" & 选项
            Wstr = Wstr & trr(1, k) & ". " & 选项
            Wstr = Wstr & vbCrLf
            End If
        Next
        
    Wstr = Wstr & "正确答案：" & trr(i, 4) '答案
    Wstr = Wstr & vbCrLf
    Wstr = Wstr & vbCrLf
    Next
    
    '写入Word文档中
    wdcx.content.text = Wstr
'    wordApp.Visible = True
    path = ThisWorkbook.path & "\" & Split(文件名, ".")(0) & ".docx"
    wdcx.SaveAs ThisWorkbook.path & "\" & Split(文件名, ".")(0) & ".docx", 1
'    wb.Close False
    
    Wstr = ""

    格式处理 wordApp.ActiveDocument
    wdcx.Close True
    wordApp.Quit

Application.ScreenUpdating = True '开启屏幕刷新
Application.DisplayAlerts = True '开启警告界面
    
MsgBox "转换完成！！" & Chr(10) & path, 64, "钦州港运转 - Momda"
End Sub
Function 格式处理(doc) 'ActiveDocument
With doc.Range.PageSetup
        .LineNumbering.Active = False '
        .Orientation = wdOrientPortrait
        .TopMargin = 20 '
        .BottomMargin = 20 '
        .LeftMargin = 30 '
        .RightMargin = 30 '
        .Gutter = 0
        .HeaderDistance = 1.5
        .FooterDistance = 1.75
''        .PageWidth = 8
'        .PageHeight = 8
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .VerticalAlignment = wdAlignVerticalTop
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
'        .CharsLine = -10
'        .LinesPage = 30 '
        .LayoutMode = wdLayoutModeLineGrid
    End With

    With doc.PageSetup.TextColumns
        .SetCount NumColumns:=2
        .EvenlySpaced = True
        .LineBetween = True
'        .Width = 0.5
'        .Spacing = 2
    End With

    With doc.Range.ParagraphFormat
'        .SpaceBefore = 0
        .SpaceBeforeAuto = False
'        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 12
        .LineUnitBefore = 0
        .LineUnitAfter = 0
        .WordWrap = True
    End With
End Function
