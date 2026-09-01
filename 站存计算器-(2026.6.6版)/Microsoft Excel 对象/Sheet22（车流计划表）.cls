VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton_插入行_Click()
    Dim xrr
    Application.EnableEvents = False
    With Me
        r1 = Selection.Row
        r2 = IIf(.Cells(24, "D").End(3).Row > 23, 23, .Cells(24, "D").End(3).Row)
        If r1 = 23 Then
            MsgBox "已是最大行！！！"
        ElseIf r2 = 3 Then
            Debug.Print "无效插入行！"
        Else
            xrr = .Range(.Cells(r1, "B"), .Cells(r2, "E"))
            .Range(.Cells(r1, "B"), .Cells(r2, "E")).ClearContents
            r3 = IIf(r2 + 1 > 23, 23, r2 + 1)
            .Range(.Cells(r1 + 1, "B"), .Cells(r3, "E")).value = xrr
        End If
    End With
    Me.CommandButton_插入行.Visible = False
    Application.EnableEvents = True
End Sub

Private Sub CommandButton1_Click()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    UserForm5.Show 0
End Sub

Private Sub CommandButton2_Click()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Private Sub CommandButton3_Click()
    站存导出.车流表导出
End Sub

Private Sub CommandButton4_Click()
    站存导出.启动发送文件窗口
End Sub

Private Sub SpinButton1_SpinUp()
    Dim y%, m%, d%, text$
    '预计2024年 4 月 25 日 03:00 - 次日 18:00 (出发)车流
    y = Year(Date)
    m = Month(Date)
    d = Day(Date)
    text = "预计" & y & "年" & m & "月" & d & "日 03:00 - 次日 18:00 (出发)车流"
    Me.Range("A2").value = text
End Sub
Private Sub SpinButton1_SpinDown()
    Dim y%, m%, d%, text$
    '预计2024年 4 月 25 日 03:00 - 次日 18:00 (出发)车流
    y = Year(Date)
    m = Month(Date)
    d = Day(Date)
    text = "预计" & y & "年" & m & "月" & d & "日 14:00 - 次日 00:00 (出发)车流"
    Me.Range("A2").value = text
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim rng As Range
    Set rng = Me.Range("B4:B23")
    If Target.count = 1 And Target.Column = 2 Then
        If Not Intersect(rng, Target) Then
            Cancel = True
            MsgBox "请直接输入3或4位的时间！" & Chr(10) & "如：820，即得到：08:20" & Chr(10) & "或：1236，即得到：12:36"
        End If
    End If
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rngs As Range, sj As Date, M1 As Boolean, M2 As Boolean, M3 As Boolean, M4 As Boolean, M5 As Boolean, ts
    Application.EnableEvents = False
    Set rngs = Me.Range("B4:B23")
    M1 = Target.Row > 3
    If Target.Rows.count = 1 Then
        If Target.Columns.count = 1 Then M2 = Len(Target.value) > 4 Or val(Target.value) > 2359 Or Target.value = ""
    Else
        M2 = False
    End If
    M3 = Target.Row < 24
    M4 = Target.Column = 2
    M5 = Target.Column = 4
    If M1 And M2 And M3 And M4 Then
        GoTo s
    ElseIf M1 And M3 And M5 Then
        With Me.Rows(Target.Row)
'            .AutoFit
'            If .RowHeight < 20 Then
'                .RowHeight = 20
'            End If
        End With
    ElseIf Not Intersect(rngs, Target) Is Nothing And Target.Rows.count = 1 And Target.Columns.count = 1 Then
        '时间格式调整
        If InStr(Target.value, "：") > 0 Then Target.value = Replace(Target.value, "：", ":")
        
        ts = VBA.Format(Target.value, "0000")
        Target.NumberFormatLocal = "hh:mm"
        Target.HorizontalAlignment = xlCenter '设置上下居中
        Target.VerticalAlignment = xlCenter '设置左右居中
        On Error GoTo s
        sj = VBA.Left(ts, 2) & ":" & VBA.Right(ts, 2)
        Target.value = sj
        On Error GoTo 0
    End If
    Application.EnableEvents = True
    Exit Sub
s:
    If Target.value = "" Then
        Target.Resize(1, 4).value = ""
    Else
        Target.value = ""
    End If
    Target.HorizontalAlignment = xlCenter '设置上下居中
    Target.VerticalAlignment = xlCenter '设置左右居中
    On Error GoTo 0
    Application.EnableEvents = True
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim rngs As Range
    
    Set rngs = Me.Range("A2:D2")
    If Not Intersect(rngs, Target) Is Nothing And Target.Rows.count = 1 Then
        Debug.Print "1"
        With SpinButton1
            .width = 36
            .height = Me.Range("E2").height
            .Left = Me.Range("E2").Left
            .Top = Me.Range("E2").Top
            .Visible = True
        End With
    Else
        SpinButton1.Visible = False
    End If
    
    Set rngs = Me.Range("B4:E20")
    If Not Intersect(rngs, Target) Is Nothing And Target.Rows.count = 1 And Target.Row > 3 Then
        With CommandButton_插入行
            .Visible = True
            .Left = Me.Cells(Selection.Row, "F").Left
            .Top = Me.Cells(Selection.Row, "F").Top
            .width = 55
            .height = 25
            .Caption = "插入一行"
        End With
    Else
        CommandButton_插入行.Visible = False
    End If
    
End Sub

