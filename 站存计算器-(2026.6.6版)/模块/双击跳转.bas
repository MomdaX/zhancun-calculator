Attribute VB_Name = "双击跳转"
Sub 跳转(rng As Range)
'Application.Goto Sheet7.Range("N1"), True
'Application.Goto rng, True
Sheet7.Range("A3:Q3").AutoFilter 1, rng.value
End Sub

Sub 返回股道存车表()
    Dim rngs As Range, rng As Range
    Sheet6.Select
    Set rngs = Sheet7.Cells.SpecialCells(xlCellTypeVisible)
    For Each rng In rngs.Areas '取筛选后的可见区域进数组中
        k = k + 1
        If k = 2 Then
            rng.Columns(12).ClearContents
            Exit For
        End If
    Next
    Sheet7.AutoFilterMode = False
End Sub
Sub 返回到发列车表()
    Dim rngs As Range, rng As Range
    Sheet4.Select
    Set rngs = Sheet7.Cells.SpecialCells(xlCellTypeVisible)
    For Each rng In rngs.Areas '取筛选后的可见区域进数组中
        k = k + 1
        If k = 2 Then
            rng.Columns(12).ClearContents
            Exit For
        End If
    Next
    Sheet7.AutoFilterMode = False
End Sub

Sub 上一股道()
Dim rngs As Range, rng As Range, cell As Range
Set rngs = Sheet7.AutoFilter.Range.SpecialCells(xlCellTypeVisible)
Set cell = rngs.Areas(rngs.Areas.count)
If cell.Cells(1.1) = "股道" Then
    当前股道 = cell.Cells(2, 1)
Else
    当前股道 = cell.Cells(1, 1)
End If
Set rng = Sheet6.Columns(2).Find(当前股道)
    
    If rng.Offset(-1).Row = 1 Then
       MsgBox "当前已经是第一条股道了！"
    Else
        Sheet7.AutoFilterMode = False
        Sheet7.Range("A3:Q3").AutoFilter 1, rng.Offset(-1).value
    End If
End Sub


Sub 下一股道()
Dim rngs As Range, rng As Range, cell As Range
Set rngs = Sheet7.AutoFilter.Range.SpecialCells(xlCellTypeVisible)
Set cell = rngs.Areas(rngs.Areas.count)

If cell.Cells(1.1) = "股道" Then
    当前股道 = cell.Cells(2, 1)
Else
    当前股道 = cell.Cells(1, 1)
End If

Set rng = Sheet6.Columns(2).Find(当前股道)
    If rng.Offset(1).Row = 91 Or rng.Offset(1).value = "" Then
       MsgBox "当前已经是最后一条股道了！"
    Else
        Sheet7.AutoFilterMode = False
        Sheet7.Range("A3:Q3").AutoFilter 1, rng.Offset(1).value
    End If
End Sub

Sub Sheet7_筛选()

Sheet7.Range("A3:Q3").AutoFilter
'设置数据验证信息（方向）
'Call 显示信息.到站方向

End Sub

'Sub Sheet4_保存()
'    If Sheet14.UsedRange.Row = 1 And Sheet14.UsedRange.Rows.count = 1 Then
'        Sheet4.UsedRange.Copy Sheet14.Cells(1, 1)
'    Else
'        Sheet4.UsedRange.Copy Sheet14.Cells(Sheet14.UsedRange.Row + Sheet14.UsedRange.Rows.count + 1, 1)
'    End If
'End Sub

Sub 测试()

Set rng = Sheet4.UsedRange
rng.Select
Set rng = Sheet4.Range("A1").CurrentRegion
rng.Select

End Sub
