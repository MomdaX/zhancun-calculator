Attribute VB_Name = "单元格区域保存为图片"
Sub 保存为图片()

    Dim rng As Range
    Set rng = ActiveSheet.UsedRange '获取当前选定的单元格区域
    'rng.Select
    '创建新的图表对象，并将选定区域复制到图表上
    rng.CopyPicture xlScreen, xlPicture
    '创建新的图表对象并将复制的图像粘贴到该图表上
    Dim chartObj As ChartObject
    Set chartObj = ActiveSheet.ChartObjects.Add(0, 0, rng.width, rng.height)
    chartObj.Activate
    chartObj.Chart.Paste
    '设置图表的格式和大小
    chartObj.Chart.ChartArea.Border.LineStyle = xlNone
    chartObj.Chart.Parent.width = rng.width
    chartObj.Chart.Parent.height = rng.height
    '弹出文件对话框选择文件名
    Dim fileName As Variant
    Dim 图片名 As String
    图片名 = ActiveSheet.Range("A1")
    fileName = Application.GetSaveAsFilename(InitialFileName:=图片名 & ".jpg", fileFilter:="JPEG 文件 (*.jpg), *.jpg")
    If fileName <> False Then
        '将图表保存为 JPEG 图片
        chartObj.Chart.Export fileName:=fileName, FilterName:="JPEG"
        'MsgBox "图片已保存至：" & fileName
    End If
    '删除临时创建的图表对象
    chartObj.Delete
End Sub
Sub 站存保存为图片()

    Dim rng As Range
    Set rng = [A1:H18] '获取当前选定的单元格区域
    'rng.Select
    '创建新的图表对象，并将选定区域复制到图表上
    rng.CopyPicture xlScreen, xlPicture
    '创建新的图表对象并将复制的图像粘贴到该图表上
    Dim chartObj As ChartObject
    Set chartObj = ActiveSheet.ChartObjects.Add(0, 0, rng.width, rng.height)
    chartObj.Activate
    chartObj.Chart.Paste
    '设置图表的格式和大小
    chartObj.Chart.ChartArea.Border.LineStyle = xlNone
    chartObj.Chart.Parent.width = rng.width
    chartObj.Chart.Parent.height = rng.height
    '弹出文件对话框选择文件名
    Dim fileName As Variant
    Dim 图片名 As String
    图片名 = ActiveSheet.Range("A1")
    fileName = Application.GetSaveAsFilename(InitialFileName:=图片名 & ".jpg", fileFilter:="JPEG 文件 (*.jpg), *.jpg")
    If fileName <> False Then
        '将图表保存为 JPEG 图片
        chartObj.Chart.Export fileName:=fileName, FilterName:="JPEG"
        'MsgBox "图片已保存至：" & fileName
    End If
    '删除临时创建的图表对象
    chartObj.Delete
End Sub
