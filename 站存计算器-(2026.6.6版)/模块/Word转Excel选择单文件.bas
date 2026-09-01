Attribute VB_Name = "Word转Excel选择单文件"
Sub 提取()
Dim files As Object
Dim wordApp As Word.Application
  Application.StatusBar = "正在创建。。。"
  


Set wordApp = New Word.Application
Set files = CreateObject("Scripting.FileSystemObject")
If files.FileExists(ThisWorkbook.path & "\题库.docx") = True Then
'删除
Kill ThisWorkbook.path & "\题库.docx"
'MsgBox "目录下存在同名：《题库.docx》文档了" & vbCrLf & "请将其移动到别的地方" & vbCrLf & "否则会被覆盖" & vbCrLf & "当前文件夹下无《题库.docx》文档后，再重新执行本代码"
'Exit Sub
Else
With wordApp
'.Visible = True
' .Application.StatusBar = "正在创建word"
.Documents.Add
' .ActiveDocument.Paragraphs(1).Range.InsertBefore ("测试文档")
' .Application.StatusBar = "正在保存"
.ActiveDocument.SaveAs (ThisWorkbook.path & "\题库.docx")
' .Application.StatusBar = "正在退出"
.Quit
End With
Set wordApp = Nothing
Application.StatusBar = False
Call Excel题库提取到Word
End If
End Sub


Sub Excel题库提取到Word()
    Dim wdcx As Object, wd As Object
    Dim filePath As String
    
    Application.ScreenUpdating = False '关闭屏幕刷新
    Application.DisplayAlerts = False '关闭警告界面
    '打开文件对话框
    With Application.fileDialog(msoFileDialogOpen)
        .InitialFileName = "F:\Momda\试卷\2023-8-13\助理值班员（外勤）2023.8.8\"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls*,*.xls*"
        .AllowMultiSelect = False
        If .Show = -1 Then
            filePath = .SelectedItems(1)
            '显示文件名
        End If
    End With
    
'如果没有选择文件，路径为空时，结束程序
If filePath = "" Then Exit Sub
path = filePath

' 打开工作簿并将数据装入数组中
Dim wb As Workbook
Dim ws As Worksheet
Set wb = Workbooks.Open(filePath)
Set ws = wb.ActiveSheet

'Set wd = wdcx.Documents.Open(ThisWorkbook.Path & "\题库.docx")
Set wdcx = CreateObject("word.application")
Set wd = wdcx.Documents.Open(ThisWorkbook.path & "\题库.docx")
'wdcx.Visible = True
For i = 3 To ActiveSheet.UsedRange.Rows.count
If ws.Range("D" & i).value = "判断题" Then
wdcx.Selection.TypeText text:=ws.Range("A" & i).value & "、" & ws.Range("C" & i).value
wdcx.Selection.TypeParagraph
    If ws.Range("G" & i).value = "A" Then
    wdcx.Selection.TypeText text:="正确答案" & "：" & "对"
    wdcx.Selection.TypeParagraph
    wdcx.Selection.TypeParagraph
    ElseIf ws.Range("G" & i).value = "B" Then
    wdcx.Selection.TypeText text:="正确答案" & "：" & "错"
    wdcx.Selection.TypeParagraph
    wdcx.Selection.TypeParagraph
    End If
ElseIf ws.Range("D" & i).value = "单选题" Or ws.Range("D" & i).value = "多选题" Then
       
    wdcx.Selection.TypeText text:=ws.Range("A" & i).value & "、" & ws.Range("C" & i).value
    wdcx.Selection.TypeParagraph

arr = Array("A", "B", "C", "D", "E", "F", "G")
For k = 0 To UBound(arr)
    If ws.Cells(i, 13 + k).value = "" Then: Exit For
        wdcx.Selection.TypeText text:=arr(k) & "、" & ws.Cells(i, 13 + k).value
        wdcx.Selection.TypeParagraph
    Next
    wdcx.Selection.TypeText text:="正确答案" & "：" & ws.Range("G" & i).value
    wdcx.Selection.TypeParagraph
    wdcx.Selection.TypeParagraph
ElseIf ws.Range("D" & i).value = "填空题" Then
    wdcx.Selection.TypeText text:=ws.Range("A" & i).value & "、" & ws.Range("C" & i).value
    wdcx.Selection.TypeParagraph
    wdcx.Selection.TypeText text:="正确答案" & "：" & ws.Range("G" & i).value
    wdcx.Selection.TypeParagraph
    wdcx.Selection.TypeParagraph
ElseIf ws.Range("D" & i).value = "简答题" Then
    wdcx.Selection.TypeText text:=Range("A" & i).value & "、" & ws.Range("C" & i).value
    wdcx.Selection.TypeParagraph
    wdcx.Selection.TypeText text:="参考答案" & "：" & ws.Range("G" & i).value
    wdcx.Selection.TypeParagraph
    wdcx.Selection.TypeParagraph
End If
Next
wd.SaveAs ThisWorkbook.path & "\" & Split(wb.name, ".")(0) & ".docx"
wd.Close
wdcx.Quit

'删除
Kill ThisWorkbook.path & "\题库.docx"

'工作簿关闭
wb.Close
Set wdoc = Nothing
Set wapp = Nothing
Set wb = Nothing
MsgBox "Excel转Word完成，莫莫哒！" & vbCrLf & "题库保存在：" & vbCrLf & ThisWorkbook.path & "\" & Split(wb.name, ".")(0) & ".docx"
End Sub


