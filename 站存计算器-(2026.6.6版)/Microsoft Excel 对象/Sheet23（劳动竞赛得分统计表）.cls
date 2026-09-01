VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    保存入库
    班组竞赛.排名刷新
End Sub

Private Sub CommandButton2_Click()
    Application.EnableEvents = False
    Dim rng As Range
    Set rng = 可编辑区域
    rng.ClearContents
    Application.EnableEvents = True
End Sub

Private Sub CommandButton3_Click() '刷新按钮
    班组竞赛.排名刷新
End Sub

Private Sub SpinButton1_SpinUp() '上
    Dim 余数 As Long, 初始日期 As Date, 查询日期 As Date, reg As New RegExp, key As String, 三班 As String, 四班 As String
    With reg
        .Global = True
    '    .IgnoreCase = True
        .Pattern = "(\d+)年\s*(\d+)\s*月\s*(\d+)\s*日\s+([甲乙丙丁])班（四班）\s+([甲乙丙丁])班（三班）\s+(白)班/(夜)班"
    End With
    初始日期 = DateValue("2019/12/31")

    With ActiveSheet
        Set tx = reg.Execute(.Range("A2"))(0).SubMatches
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        If Range("A2").Characters(36, 2).Font.Strikethrough Then
            查询日期 = DateSerial(tx(0), tx(1), tx(2))
            差值 = 查询日期 - 初始日期
'            余数 = 差值 Mod 4
            三班 = 班组竞赛.班次(3, 差值 Mod 3, "白")
            四班 = 班组竞赛.班次(4, 差值 Mod 4, "白")
            With .Range("A2")
                .value = tx(0) & "年 " & tx(1) & " 月 " & tx(2) & " 日  " & 四班 & "班（四班）   " & 三班 & "班（三班）    白班/夜班"
                .Font.Strikethrough = False
                .Font.Bold = False
                .Characters(Start:=39, length:=2).Font.Strikethrough = True
                .Characters(Start:=36, length:=2).Font.Bold = True
            End With
        Else
            查询日期 = DateSerial(tx(0), tx(1), tx(2) - 1)
            差值 = 查询日期 - 初始日期
'            余数 = 差值 Mod 4
            三班 = 班组竞赛.班次(3, 差值 Mod 3, "夜")
            四班 = 班组竞赛.班次(4, 差值 Mod 4, "夜")
            With .Range("A2")
                .value = Year(查询日期) & "年 " & Month(查询日期) & " 月 " & Day(查询日期) & " 日  " & 四班 & "班（四班）   " & 三班 & "班（三班）    白班/夜班"
                .Font.Strikethrough = False
                .Font.Bold = False
                .Characters(Start:=36, length:=2).Font.Strikethrough = True
                .Characters(Start:=39, length:=2).Font.Bold = True
            End With
        End If
        
        查询 Format(查询日期, "yyyy-m-d") & 四班
        班组竞赛.排名刷新
        定位 Format(查询日期, "yyyy-m-d"), 四班
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End With
End Sub

Private Sub SpinButton1_SpinDown() '下

    Dim 余数 As Long, 初始日期 As Date, 查询日期 As Date, reg As New RegExp, key As String, 四班 As String, 三班 As String
    With reg
        .Global = True
    '    .IgnoreCase = True
        .Pattern = "(\d+)年\s*(\d+)\s*月\s*(\d+)\s*日\s+([甲乙丙丁])班（四班）\s+([甲乙丙丁])班（三班）\s+(白)班/(夜)班"
    End With
    初始日期 = DateValue("2019/12/31")

    With ActiveSheet
        Set tx = reg.Execute(.Range("A2"))(0).SubMatches
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        If Range("A2").Characters(39, 2).Font.Strikethrough Then
            查询日期 = DateSerial(tx(0), tx(1), tx(2))
            差值 = 查询日期 - 初始日期
'            余数 = 差值 Mod 4
            三班 = 班组竞赛.班次(3, 差值 Mod 3, "夜")
            四班 = 班组竞赛.班次(4, 差值 Mod 4, "夜")
            With .Range("A2")
                .value = tx(0) & "年 " & tx(1) & " 月 " & tx(2) & " 日  " & 四班 & "班（四班）   " & 三班 & "班（三班）    白班/夜班"
                .Font.Strikethrough = False
                .Font.Bold = False
                .Characters(Start:=36, length:=2).Font.Strikethrough = True
                .Characters(Start:=39, length:=2).Font.Bold = True
            End With
        Else
            查询日期 = DateSerial(tx(0), tx(1), tx(2) + 1)
            差值 = 查询日期 - 初始日期
            余数 = 差值 Mod 4
            三班 = 班组竞赛.班次(3, 差值 Mod 3, "白")
            四班 = 班组竞赛.班次(4, 差值 Mod 4, "白")
            With .Range("A2")
                .value = Year(查询日期) & "年 " & Month(查询日期) & " 月 " & Day(查询日期) & " 日  " & 四班 & "班（四班）   " & 三班 & "班（三班）    白班/夜班"
                .Font.Strikethrough = False
                .Font.Bold = False
                .Characters(Start:=39, length:=2).Font.Strikethrough = True
                .Characters(Start:=36, length:=2).Font.Bold = True
            End With
        End If

        查询 Format(查询日期, "yyyy-m-d") & 四班
        班组竞赛.排名刷新
        定位 Format(查询日期, "yyyy-m-d"), 四班
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End With
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim arr, brr(1 To 28), rng As Range, rn As Range
    Dim reg As New RegExp
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error Resume Next
    '判断是否在可编辑区域
    Set rng = 可编辑区域
    Set rn = Intersect(Target, rng)
    If Target.Column > 11 Then
        End
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    ElseIf rn Is Nothing Then
        Application.Undo '撤消修改
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        End
    End If
    
    '正则
    With reg
        .Global = True
    '    .IgnoreCase = True
        .Pattern = "[\d]{4,}/(\d+)[^\d]?"
    End With
    
    '获取日期班次
    arr = ActiveSheet.Range("A1").CurrentRegion
    
    For i = 5 To 16 '对位装卸
       c = IIf(InStr(arr(i, 3), "+"), Application.Evaluate(arr(i, 3)), arr(i, 3))
       E = IIf(InStr(arr(i, 5), "+"), Application.Evaluate(arr(i, 5)), arr(i, 5))
       h = IIf(InStr(arr(i, 8), "+"), Application.Evaluate(arr(i, 8)), arr(i, 8))
       j = IIf(InStr(arr(i, 9), "+"), Application.Evaluate(arr(i, 9)), arr(i, 9))
       brr(i - 4) = ((c + E) * arr(i, 7) * 1) + ((h + j) * arr(i, 10) * 1)
    Next
    
    For i = 17 To 25 Step 3 '摘挂列车
        '方案1：49987/50 34102/54
'        For Each matc In reg.Execute(arr(i, 3))
'            brr(i - 4) = brr(i - 4) + matc.SubMatches(0) * 1
'        Next

        '方案2：50+54
        brr(i - 4) = IIf(InStr(arr(i, 3), "+"), Application.Evaluate(arr(i, 3)), arr(i, 3))
        
        If brr(i - 4) <> "" Then
            brr(i - 4) = brr(i - 4) * arr(i, 10) * 1
        End If
    Next
    
    For i = 18 To 25 Step 3 '管内直达列车
        '方案1：49987 34102
'        If arr(i, 6) <> "" Then
'            brr(i - 4) = (UBound(Split(Trim(arr(i, 6)), " ")) + 1) * arr(i, 10)
'        End If
        
        '方案2：3(列)
        If arr(i, 6) <> "" Then
            brr(i - 4) = arr(i, 6) * arr(i, 10)
        End If
    Next
    
    For i = 19 To 25 Step 3 '跨局直达列车
        '方案1：49987 34102
'        If arr(i, 6) <> "" Then
'            brr(i - 4) = (UBound(Split(Trim(arr(i, 6)), " ")) + 1) * arr(i, 10)
'        End If
        
        '方案2：3(列)
        If arr(i, 6) <> "" Then
            brr(i - 4) = arr(i, 6) * arr(i, 10)
        End If
    Next
    
    For i = 26 To 32 '其它
'        brr(i - 4) = arr(i, 3) * arr(i, 10) * 1
        brr(i - 4) = IIf(InStr(arr(i, 3), "+"), Application.Evaluate(arr(i, 3)), arr(i, 3)) * arr(i, 10) * 1
    Next
    
    '输出合计分数
    With ActiveSheet
        Set rng = .Range("K5:K32") '小计得分
        rng.ClearContents
        rng.value = WorksheetFunction.Transpose(brr)
        .Range("I34") = WorksheetFunction.Sum(rng) '合计得分
    End With
    On Error GoTo 0
    
    '统计入位车数
    Call 班组竞赛.入位数
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Function 查询(key As String)
    Dim d As New Dictionary
    '获取已有的记录
    arr = Sheet24.UsedRange
    rs = Sheet24.Cells(Rows.count, 2).End(3).Row
    For i = 3 To rs
        d(arr(i, 2) & arr(i, 3)) = i
    Next

    With ActiveSheet
        Call CommandButton2_Click '清空
        .Columns("A:B").Interior.ColorIndex = 0
        If d.Exists(key) Then '比对
            Application.EnableEvents = False
            brr = Application.Index(arr, d(key), 0)
           
            For i = 5 To 16
                .Cells(i, 3) = brr(i - 1)
                .Cells(i, 5) = brr(i + 11)
                .Cells(i, 8) = brr(i + 23)
                .Cells(i, 9) = brr(i + 35)
            Next
            
            For i = 17 To 25 Step 3 '摘挂列车
                .Cells(i, 3) = brr(i + 35)
            Next
            
            For i = 18 To 25 Step 3 '直达列车
                .Cells(i, 6) = brr(i + 35)
                .Cells(i + 1, 6) = brr(i + 36)
            Next

            '其他
            For i = 26 To 32
                .Cells(i, 3) = brr(i + 35)
            Next
            
            '小计得分
            For i = 5 To 32
                .Cells(i, 11) = brr(i + 63)
            Next
            
            .Cells(33, 3) = brr(96) '备注
            .Cells(34, 3) = brr(97) '运用车数
            .Cells(34, 9) = brr(98) '总分
            
            Application.EnableEvents = True
        End If
    End With

End Function


Private Sub Worksheet_SelectionChange(ByVal Target As Range) '定位
    r = Target.Row
    With ActiveSheet
        If Target.Column < 2 Or Target.Column > 11 Then
            .Columns("A:B").Interior.ColorIndex = 0
        Else
            .Columns("A:B").Interior.ColorIndex = 0
            .Cells(r, 2).MergeArea.Interior.ColorIndex = 4
        End If
    End With
End Sub

Function 可编辑区域()
    Dim rng As Range
    With ActiveSheet
        Set rng = .Range("C5:F16")
        Set rng = Union(.Range("H5:I16"), rng)
        Set rng = Union(.Range("K5:K32"), rng)
        Set rng = Union(.Range("C26:I32"), rng)
        Set rng = Union(.Range("C33:L33"), rng)
        Set rng = Union(.Range("C34:F34"), rng)
        Set rng = Union(.Range("I34:L34"), rng)
        Set rng = Union(.Range("C17:I17"), rng)
        Set rng = Union(.Range("C20:I20"), rng)
        Set rng = Union(.Range("C23:I23"), rng)
        Set rng = Union(.Range("F18:I18"), rng)
        Set rng = Union(.Range("F19:I19"), rng)
        Set rng = Union(.Range("F21:I21"), rng)
        Set rng = Union(.Range("F22:I22"), rng)
        Set rng = Union(.Range("F24:I24"), rng)
        Set rng = Union(.Range("F25:I25"), rng)
    End With
    Set 可编辑区域 = rng
End Function

Function 定位(A As String, b As String)
    Dim x, y, rn As Range
    With ActiveSheet
        With .Range("Q12").CurrentRegion
            .Offset(1).Interior.ColorIndex = 0
            For Each rn In .Columns(1).Cells
                If rn.value = A Then
                    x = rn.Row
                    Exit For
                End If
            Next
            y = .Find(b).Column
        End With
        If x > 0 And y > 0 Then
            .Cells(x, y).Interior.ColorIndex = 17
        End If
    End With
End Function
Function 保存入库()
    Dim d As New Dictionary, d_t As New Dictionary, d_p As New Dictionary, arr, brr(1 To 98)
    Dim reg As New RegExp, 班次 As String
    
    If Sheet23.Range("I34") = "" Or Sheet23.Range("I34") = 0 Then
        MsgBox "提示：请录入数据！"
        End
    End If
    '获取白夜班
    If Sheet23.Range("A2").Characters(Start:=39, length:=2).Font.Strikethrough Then
        brr(1) = "白"
    Else
        brr(1) = "夜"
    End If
    
    '获取日期班次
    arr = Sheet23.Range("A3").CurrentRegion.value
    reg.Global = True '多行匹配
    reg.Pattern = "\d+" '数字
    Set rq = reg.Execute(arr(2, 1))
    日期 = Join(Array(rq(0), rq(1), rq(2)), "-")
    reg.Pattern = ".*日\s+([甲乙丙丁])班（四班）\s+([甲乙丙丁])班（三班）.*" '班别
    With reg.Execute(arr(2, 1))(0)
        四班 = .SubMatches(0)
        三班 = .SubMatches(1)
        班次 = 四班 & "-" & 三班
    End With
'    班次 = Trim(reg.Execute(arr(2, 1))(0).SubMatches(0))
    brr(2) = 日期
    brr(3) = 班次
    
    
    For i = 5 To 16
        brr(i - 1) = arr(i, 3) '* arr(i, 10) '本班入位/本班装好:5-16
        brr(i + 11) = arr(i, 5) '* arr(i, 10) '本班入位/本班卸空:17-28
        brr(i + 23) = arr(i, 8) '* arr(i, 10) '本班入位待装:29-40
        brr(i + 35) = arr(i, 9) '* arr(i, 10) '本班入位待卸:41-52
    Next
    
    '本班编好本班出车
    brr(52) = arr(17, 3)
    brr(53) = arr(18, 6)
    brr(54) = arr(19, 6)
    
    '上班编好本班出车
    brr(55) = arr(20, 3)
    brr(56) = arr(21, 6)
    brr(57) = arr(22, 6)
    
    '本班编好下班出车
    brr(58) = arr(23, 3)
    brr(59) = arr(24, 6)
    brr(60) = arr(25, 6)
    
    '其他
    brr(61) = arr(26, 3)
    brr(62) = arr(27, 3)
    brr(63) = arr(28, 3)
    brr(64) = arr(29, 3)
    brr(65) = arr(30, 3)
    brr(66) = arr(31, 3)
    brr(67) = arr(32, 3)
    
    '小计得分
    For i = 5 To 32
        brr(i + 63) = arr(i, 11)
    Next
    
    '备注栏
    brr(96) = arr(33, 3)
    '运用车数
    brr(97) = arr(34, 3)
    '合计得分
    brr(98) = arr(34, 9)
    
    '查询是否有记录
    With Sheet24
        arr = .UsedRange
        rs = .Cells(Rows.count, 2).End(3).Row
        For i = 3 To rs
            d(Trim(Join(Application.Index(arr, i, 0)))) = i '录入字典中
            d_t(arr(i, 2) & arr(i, 3)) = i '2024-9-9甲
            d_p(arr(i, 3)) = d_p(arr(i, 3)) + arr(i, 98) * 1 '甲=甲+单日分数
        Next
        key = brr(2) & brr(3)
        If d.Exists(Trim(Join(brr))) Then '比对
            MsgBox key & "-" & brr(1) & Chr(10) & "该条记录已存在！"
        Else
            If d_t(key) Then
                A = MsgBox(key & Chr(10) & "是否修改记录？", vbYesNo)
                If A = 6 Then
                    r = d_t(brr(2) & brr(3)) '获取该记录所在行号
                    .Cells(r, 1).Resize(1, 98) = brr '修改
                    .UsedRange.Borders.LineStyle = 1 '边框
                    MsgBox "修改成功！"
                End If
            Else
                r = .Cells(Rows.count, 2).End(3).Row + 1 '新记录的行号
                .Cells(r, 1).Resize(1, 98) = brr '添加记录
                .UsedRange.Borders.LineStyle = 1 '边框
                MsgBox "保存成功！"
            End If
        End If
     End With
End Function

