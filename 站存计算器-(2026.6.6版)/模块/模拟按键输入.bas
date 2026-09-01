Attribute VB_Name = "模拟按键输入"
Sub 模拟调试()
Dim rn As Range, rng As Range
    With Sheet13
        rw = .Cells(.Rows.count, "A").End(3).Row
        For i = 2 To rw
            SendKeys .Cells(i, "A").value
    '       SendKeys "{+}"
            SendKeys "{DOWN}"
            Application.Wait Now + TimeValue("0:00:01")
        Next
    End With
End Sub
