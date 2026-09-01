Attribute VB_Name = "待发的到站"
Sub 载入到站()

Dim Tng As Range, rngs As Range, rng As Range, rns As Range, rn As Range, cell As Range
Dim d As New Dictionary

Set Tng = Sheet2.Range("H1:Q1")
If Sheet2.AutoFilterMode Then
    Sheet2.AutoFilterMode = False
End If
Sheet4.Select
Set rng = Application.InputBox("请选择对应的股道！", Type:=8)
For Each rn In rng
    If rn <> "" Then
        Tng.AutoFilter 1, rn.value
        Set rngs = Sheet2.AutoFilter.Range.SpecialCells(xlCellTypeVisible)
        'rngs.Areas(rngs.Areas.count).Select
        Set cell = rngs.Areas(rngs.Areas.count)
        'cell.Select
        If Not cell Is Nothing And rn.Offset(0, 4).value = "" Then
            '到站
            For Each rns In cell.Rows
                'rns.Select
                Debug.Print rns.Cells(1)
                If Not d.Exists(rns.Cells(4).value) And rns.Cells(4) <> "" Then
                    
                    d(rns.Cells(4).value) = rns.Cells(1).value
                End If
            Next
            
            If d.count > 0 Then rn.Offset(0, 4) = Join(d.keys, " ")
            d.RemoveAll
            '辆数
            If rn.Offset(0, 4) <> "" And cell.Rows.count - 1 <> 0 Then
                If rngs.Areas.count = 1 Then
                    rn.Offset(0, 2) = cell.Rows.count - 1 '包含标题
                Else
                    rn.Offset(0, 2) = cell.Rows.count '不包含标题
                End If
            End If
        End If
        Sheet2.AutoFilter.ShowAllData
    End If
Next
End Sub
