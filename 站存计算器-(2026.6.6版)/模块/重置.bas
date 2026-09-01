Attribute VB_Name = "重置"
Sub 重置()

Dim rng As Range
Set rng = Sheet1.Range("B2:H18")
rng = ""

End Sub

'Sub Sheet4_清空()
'With Sheet4
'    .Range("A1").Value = Format(Date, "yyyy年mm月dd日")
'    .Range("B4:S15").ClearContents
'End With
'End Sub

Sub Sheet6_清空股道()

    With Sheet6.Range("C2:L91")
        .ClearContents '内容重置
        .Interior.ColorIndex = xlNone '颜色重置
    End With
    
    Sheet6.Range("A2:A92").ClearContents  '车次重置
End Sub


Sub Sheet6_清空重空车()

Sheet6.Range("H2:I93").ClearContents
Sheet6.AutoFilterMode = False

End Sub

Sub Sheet9_货主清空()
    r = Sheet9.Cells(Rows.count, "D").End(3).Row
    If WorksheetFunction.CountA(Sheet9.Range("A2:H2")) > 0 Then
        Sheet9.Range("A2:H" & r).ClearContents
    End If
End Sub

Sub 文件选择()
    Dim file As String
    file = Application.GetOpenFilename("表格文件(*.xls*),*.xls*", , "momda")
End Sub
Sub qlss()
Dim rng As Range
Dim reg As New RegExp
With reg
    .Pattern = "\d+:\d+"
    .Global = True
    .MultiLine = True
End With
Set rng = Sheet10.Range("E1").CurrentRegion
trr = rng.Offset(1).Columns(4).value



End Sub
