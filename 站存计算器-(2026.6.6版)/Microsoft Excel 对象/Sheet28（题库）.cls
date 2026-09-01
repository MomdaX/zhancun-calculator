VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim rng As Range, rn As Range, m%, n%, r%, oDom As Object, 随机数
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    If Target.address(0, 0) = "E3" Then
        Cancel = True
        If Me.AutoFilterMode Then
            Me.AutoFilterMode = False
        End If
        Me.Range("C3").value = ""
    ElseIf Target.address(0, 0) = "B5" Or Target.address(0, 0) = "C6" Then
        Cancel = True
        m = Me.Cells(Rows.count, "C").End(3).Row
        n = 9

        Randomize '产生随机序列，配合Rnd使用
        r = Int((m - n + 1) * VBA.Rnd - n) '随机数
        Debug.Print Now() & ":   " & r
        
        Set rng = Me.Range("A8").CurrentRegion
        With Me.Range("A6:L6")
            .value = rng.Rows(r - rng.Row + 1).value
            .WrapText = True '自动换行
        End With
        Me.Range("A6:L6").Interior.ColorIndex = 0 '清除颜色
        Me.Range("D6").value = "" '清空答案
    ElseIf Target.address(0, 0) = "D6" Then
        Cancel = True
        Set rn = Me.Range("D6")
        If rn.value = "" Then '点击显示答案
            Set rng = Me.Range("A8").CurrentRegion
            r = Me.Range("A6").value + rng.Row
            rn.value = Me.Cells(r, 4).value
            
            '答案标记
            Dim arr(), d As New Dictionary
            arr = Me.Range("E8:L8")
            For Each ar In arr
                k = k + 1
                d(ar) = k
            Next
            
            Me.Range("A6:L6").Interior.ColorIndex = 0 '清除颜色
            rn.Interior.ColorIndex = 8 '米黄
            For i = 1 To Len(rn.value)
                st = Mid(rn.value, i, 1)
                rn.Offset(0, d(st)).Interior.ColorIndex = 6 '黄
            Next
        Else
            Me.Range("A6:L6").Interior.ColorIndex = 0 '清除颜色
            rn.value = ""
        End If
    End If
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Application.ScreenUpdating = False
    With Me
        If .AutoFilterMode Then
            .AutoFilterMode = False
        End If
        .Range("A8").CurrentRegion.AutoFilter 3, "*" & .Range("C3").value & "*"
    End With
    Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Cells(1).value <> "" Then
        Application.ScreenUpdating = False
        试卷成绩查询.选中行显示 Me, Me.Range("A8"), Target
        Application.ScreenUpdating = True
    End If
End Sub
