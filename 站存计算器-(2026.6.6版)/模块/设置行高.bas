Attribute VB_Name = "设置行高"
Sub 行高增()

Dim rng As Range, rngs As Range
Set rngs = ActiveSheet.UsedRange
'rngs.Select
On Error Resume Next
Set rng = Selection
    If Application.Intersect(rngs, rng) Is Nothing Then
        MsgBox "请选择一个有数据的单元格", 64, "选择了错误区域"
    Else
        
        If rng.EntireRow.RowHeight + 3 > 409 Then: Debug.Print "已达到最大行高！": rng.EntireRow.RowHeight = 409.5: End
        rng.EntireRow.RowHeight = rng.EntireRow.RowHeight + 3
    End If
On Error GoTo 0
End Sub
Sub 行高减()

Dim rng As Range, rngs As Range
Set rngs = ActiveSheet.UsedRange
'rngs.Select
On Error Resume Next
Set rng = Selection
    If Application.Intersect(rngs, rng) Is Nothing Then
        MsgBox "请选择一个有数据的单元格", 64, "选择了错误区域"
    Else
        If rng.EntireRow.RowHeight - 3 < 1 Then: MsgBox "行高无法再小啦": rng.EntireRow.RowHeight = 1: End
        rng.EntireRow.RowHeight = rng.EntireRow.RowHeight - 3
    End If
On Error GoTo 0
End Sub
Sub 行高还原()

    Dim rngs As Range
    Set rngs = ActiveSheet.UsedRange.Offset(2)
    'rngs.Select
        rngs.EntireRow.AutoFit
        
End Sub

Sub 行高自动()
Dim rnw As Range, rngs As Range, rng As Range
Set rngs = Range("A2").CurrentRegion.Offset(2)
arr = rngs.Rows
r = UBound(arr, 1)
c = UBound(arr, 2)

Set rnw = rngs.Resize(r - 2, c)
'rnw.Select

For i = 1 To r - 2
    
    Set rng = rnw.Rows(i)
    'rng.Select
    If rng.EntireRow.RowHeight + 10 > 409 Then: Debug.Print "已达到最大行高！": rng.EntireRow.RowHeight = 409.5: GoTo g:
    rng.EntireRow.RowHeight = rng.EntireRow.RowHeight + 10
Next
g:
End Sub

