Attribute VB_Name = "待验证车站"
Sub 待验证车站()
    Dim d As New Dictionary
    Dim dic As New Dictionary
    Dim Qan As Range, rngs As Range
    Dim rg As Range, rang As Range
    '读取方向基础数据库
    Set rang = Sheet2.Range("A1:A" & Sheet2.Cells(Rows.count, 1).End(3).Row)
    'Sheet2.rang.Select
    Set rng = Sheet5.Range("A1:A" & Sheet5.Cells(Rows.count, 1).End(3).Row)
    'rng.Select
    For Each rg In rang
    d(rg.value) = rg.Offset(0, 1).value
    Next
    'For Each rg In rng
    'dic(rg.Value) = rg.Offset(0, 1).Value
    'Next
    'rg = ""
    'Sheet2.Range("D1").Resize(d.count, 1) = WorksheetFunction.Transpose(d.Items)
    
    For Each rg In rng
        If rg.value <> "到站" Then
        
            If d.Exists(rg.value) Then '将以有的车站条目删除
            Set rngs = Sheet5.Range(rg, rg.Offset(0, 2))
            'rngs.Select
                rngs.Delete Shift:=xlUp
            Else
                If Not dic.Exists(rg.value) Then
                    dic(rg.value) = rg.Offset(0, 1).value '在数据库中没有的车站存入dic字典中
                End If
            End If
        
        End If
    Next
    
    '收集车站数据
    czr = Sheet2.Range("H2").CurrentRegion
    
    For i = 2 To UBound(czr)
    Set Qan = Sheet5.Range("A" & Sheet5.Cells(Rows.count, 1).End(3).Row).Offset(1, 0)
    'Qan.Select
    If czr(i, 4) <> "" And Not d.Exists(czr(i, 4)) And Not dic.Exists(czr(i, 4)) Then
            Qan.value = czr(i, 4)
            Qan.Offset(0, 1).value = czr(i, 9)
            Qan.Offset(0, 2).value = czr(i, 5)
            dic(czr(i, 4)) = czr(i, 9)
    End If
    Next
    Set rng = Sheet5.Range("C2:C" & Sheet5.Cells(Rows.count, 1).End(3).Row)
    For Each rg In rng
        If rg.value <> "代号" And rg.value <> "" Then
            If rg.value = "3" Then '方向代号为3时为南口
                rg.Offset(0, -1).value = "南口"
            ElseIf rg.value = "6" Then '方向代号为6时为管内
                rg.Offset(0, -1).value = "管内"
            Else
               rg.Offset(0, -1).value = "沙口" '其它代号为沙口
            End If
        End If
    Next
    If Sheet5.Range("A2") <> "" Then
        Dim bkg As Range
        Set bkg = Sheet5.Range("A1").CurrentRegion

       With bkg
            .Borders.LineStyle = xlContinuous '设置样式边框
            .Borders.Weight = xlThin '边框粗细
            .HorizontalAlignment = xlCenter '设置上下居中
            .VerticalAlignment = xlCenter '设置左右居中
        End With
    End If
        
    '保存工作簿
    'ThisWorkbook.Save

End Sub


