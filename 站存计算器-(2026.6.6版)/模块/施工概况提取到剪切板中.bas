Attribute VB_Name = "施工概况提取到剪切板中"
Sub 施工概况提取()
Dim rng As Range
Dim objData As Object

r = Cells(Rows.count, 1).End(3).Row
arr = ActiveSheet.Range("B3:C" & r)

'登记站=
For i = 1 To UBound(arr, 1)

reunt = reunt & i & "、" & 提取(arr(i, 2), "施工项目：\s(.*)\s")(0).SubMatches(0) & "施工  " & arr(i, 1) & Chr(13)


Next

set1 = Format(Date, "m月d日") & "钦州港站副站长韦文康：" & Chr(13)
set3 = "约16点00开会"
XX = set1 & Chr(13) & reunt & Chr(13) & set3

    Set objData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    objData.SetText XX
    objData.PutInClipboard
    MsgBox "“施工概况信息”已提取到剪贴板了！" & Chr(10) & "找个地方粘贴吧", 64, "钦州港运转-小莫提醒您"
    
End Sub

Function 提取(text, regex As String)
Dim reg As New RegExp
With reg
    .Pattern = regex
    .Global = True
    .MultiLine = True
End With

Set 提取 = reg.Execute(text)

End Function
