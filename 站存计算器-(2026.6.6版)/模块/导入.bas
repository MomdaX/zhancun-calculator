Attribute VB_Name = "导入"
Sub 导入()
    Dim rng As Range, cg As Range, n As Integer
    Application.ScreenUpdating = False '取消屏幕刷新
    Application.DisplayAlerts = False '取消警告界面
    If Sheet5.Range("A1").CurrentRegion.Rows.count = 1 Then Exit Sub 'MsgBox "当前没有可导入数据":
    Set rng = Sheet5.Range("A2:C" & Sheet5.Cells(Rows.count, "A").End(3).Row) '导入前的区域
    'rng.Select
    Set cg = Sheet2.Range("A" & Sheet2.Cells(Rows.count, 1).End(3).Row).Offset(1, 0) '导入后的区域
    rng.Copy cg
    'cg.PasteSpecial = xlPasteFormulas
    Application.CutCopyMode = False
    rng.Clear '清空导入前的区域
    Call 设置条件格式

    '保存工作簿
    'ThisWorkbook.Save
    'ThisWorkbook.Saved = True
    Application.ScreenUpdating = True '开启屏幕刷新
    Application.DisplayAlerts = True '开启警告界面
End Sub
Sub 设置条件格式()
    Dim bkg As Range
    Application.ScreenUpdating = False '取消屏幕刷新
    Application.DisplayAlerts = False '取消警告界面
    Set bkg = Sheet2.Columns("A:C")
    
        Sheet2.Select
        bkg.Select
    
        bkg.FormatConditions.Delete  '清除条件格式
    
       With bkg.FormatConditions.Add(Type:=xlExpression, Formula1:="=$A1<>""""") '添加边框并居中
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
            For Each rng In bkg
                rng.HorizontalAlignment = xlCenter
                rng.VerticalAlignment = xlCenter
            Next
        End With
        '取消边框
        bkg.Borders.LineStyle = 0
        
    Application.ScreenUpdating = True '开启屏幕刷新
    Application.DisplayAlerts = True '开启警告界面
End Sub

