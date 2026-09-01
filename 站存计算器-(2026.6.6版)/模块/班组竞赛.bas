Attribute VB_Name = "班组竞赛"
Function 班次(A As Integer, 余数 As Long, b As String)
    Dim d As New Dictionary
    If A = 4 Then
        d("甲") = Array("夜", "休", "休", "白")
        d("乙") = Array("休", "白", "夜", "休")
        d("丙") = Array("白", "夜", "休", "休")
        d("丁") = Array("休", "休", "白", "夜")
    ElseIf A = 3 Then
        d("甲") = Array("夜", "休", "白")
        d("乙") = Array("休", "白", "夜")
        d("丙") = Array("白", "夜", "休")
    End If
    For Each bb In d.keys
        If d(bb)(余数) = b Then
            班次 = bb
            Exit For
        End If
    Next
End Function
Sub 排名刷新()
    Dim ds As New Dictionary, d As New Dictionary, d_c As New Dictionary, d_t As New Dictionary
    Dim 三班 As String, 四班 As String
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    With Sheet23 '获取年月
        r1 = .Range("R1").value '年
        S1 = .Range("S1").value '月
        .Range("Q3") = S1 & "月份排日名表-四班"
    End With
    With Sheet24
        '刷新一下已获取排名
        d.RemoveAll
        arr = .UsedRange
'        rs = .Cells(Rows.Count, 2).End(3).Row
        For i = 3 To UBound(arr) '计算指定月份的排名
            If arr(i, 2) = DateSerial(r1, S1, 1) - 1 Then
                If arr(i, 1) = "夜" Then ds(arr(i, 2) & arr(i, 3)) = Application.Index(arr, i, 0) '上个月的夜班计入本月
            ElseIf Year(arr(i, 2)) = r1 And Month(arr(i, 2)) = S1 Then
                If arr(i, 2) = DateSerial(r1, S1 + 1, 1) - 1 And arr(i, 1) = "夜" Then
                    Debug.Print arr(i, 2) & arr(i, 1) & ":" & "剔除" '本月的夜班计入下个月（本月不统计）
                Else
                    ds(arr(i, 2) & arr(i, 3)) = Application.Index(arr, i, 0)
                End If
            End If
        Next
        For Each ars In ds.items
            If InStr(ars(3), "-") Then
                四班 = "四班|" & Split(ars(3), "-")(0)
                三班 = "三班|" & Split(ars(3), "-")(1)
            Else
                四班 = "四班|" & ars(3)
                三班 = "三班|" & ars(3)
            End If
            
            '四班计分
            d(四班) = d(四班) + ars(98) * 1 '分数累计
            d_t(ars(2) & 四班) = ars(98) '日期累计
            d_c(四班) = d_c(四班) + 1 '班数累计
            
            '三班计分
            d(三班) = d(三班) + ars(98) * 1 '分数累计
            d_t(ars(2) & 三班) = ars(98) '日期累计
            d_c(三班) = d_c(三班) + 1 '班数累计
        Next
    
        '计算总分数-排名
        With Sheet23
            '四班计分
            .Range("R6:S9").NumberFormatLocal = "0.0_ "
            .Range("R6:S9").ClearContents '清空
            .Range("R6").value = d("四班|甲")
            .Range("R7").value = d("四班|乙")
            .Range("R8").value = d("四班|丙")
            .Range("R9").value = d("四班|丁")
            If d_c("四班|甲") > 0 Then .Range("S6").value = d("四班|甲") / d_c("四班|甲")
            If d_c("四班|乙") > 0 Then .Range("S7").value = d("四班|乙") / d_c("四班|乙")
            If d_c("四班|丙") > 0 Then .Range("S8").value = d("四班|丙") / d_c("四班|丙")
            If d_c("四班|丁") > 0 Then .Range("S9").value = d("四班|丁") / d_c("四班|丁")
            
            '三班计分
            .Range("X6:Y9").NumberFormatLocal = "0.0_ "
            .Range("X6:Y9").ClearContents '清空
            .Range("X6").value = d("三班|甲")
            .Range("X7").value = d("三班|乙")
            .Range("X8").value = d("三班|丙")
            If d_c("三班|甲") > 0 Then .Range("Y6").value = d("三班|甲") / d_c("三班|甲")
            If d_c("三班|乙") > 0 Then .Range("Y7").value = d("三班|乙") / d_c("三班|乙")
            If d_c("三班|丙") > 0 Then .Range("Y8").value = d("三班|丙") / d_c("三班|丙")
        
            '每天分数详情-2025-1-17先到这里
            i = 16 '开始位置
            If .Cells(12, "Q") <> "" Then
                .Cells(12, "Q").CurrentRegion.Clear '清空
                .Cells(12, "P").value = "四班"
                
                '三班
                .Cells(12, "X").CurrentRegion.Clear '清空
                .Cells(12, "W").value = "三班"
            End If
            For Each bz In Array("四班|", "三班|")
                For Each t In Array("甲", "乙", "丙", "丁") '标头
                    i = i + 1
                    If b > 0 And i > 19 Then Exit For
                    .Cells(12, i + b) = t
                Next
                For i = 13 To 45
                    日期 = DateSerial(r1, S1, 1) - 1 + k
                    If Month(日期) <> S1 And i > 13 Then Exit For
                    .Cells(i, 16 + b) = Format(日期, "yyyy-m-d")
                    For c = 17 To 20
                        If b > 0 And c > 19 Then Exit For
                        If d_t.Exists(.Cells(i, "P").value & bz & .Cells(12, c + b).value) Then
                        .Cells(i, c + b) = d_t(.Cells(i, 16 + b).value & bz & .Cells(12, c + b).value)
                        End If
                    Next
                    k = k + 1
                Next
                
                r = .Cells(Rows.count, 16 + b).End(3).Row + 1
                .Range(.Cells(13, 17 + b), .Cells(r - 1, 20 + b)).NumberFormatLocal = "0.0_ "
                .Cells(r, 16 + b).Resize(1, 5) = Array("班数", d_c("四班|甲"), d_c("四班|乙"), d_c("四班|丙"), d_c("四班|丁")) '班数合计
                
                If .Cells(12, 17 + b) <> "" Then
                    With .Cells(12, 17 + b).CurrentRegion
                        .Borders.LineStyle = 1
                        .HorizontalAlignment = 3
                        .VerticalAlignment = 2
                        With .Rows(1)
                            .Interior.Color = 5287936
                            .Font.ColorIndex = 2
                            .Font.Bold = 1
                        End With
                    End With
                End If
                b = b + 7
                k = 0
                i = 16
            Next
        End With
    End With
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub 入位数()
Dim rng As Range, rn As Range
Set rng = Union(Range("C5:F16"), Range("H5:I16"))
For Each rn In rng
    If rn <> "" And tx <> "" Then
        tx = tx & "+" & rn.value
    ElseIf rn <> "" Then
        tx = rn.value
    End If
Next
If InStr(tx, "+") > 0 Then
    Range("L5").value = "货调" & Chr(10) & Application.Evaluate(tx)
Else
    Range("L5").value = "货调" & Chr(10) & tx
End If
End Sub
