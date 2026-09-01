VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "Momda-钦州港运转(莫)"
   ClientHeight    =   8028
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   15540
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public d_py As Object, 开关 As Boolean

Private Sub CommandButton1_Click()
    Dim rn As Range
    If Selection.Column = 4 And Selection.Row > 3 And Selection.Row < 24 Then
        Set rn = Selection.Cells(1)
        Application.EnableEvents = False
        If TextBox1.value & TextBox2.value = "到站辆数" Then
            MsgBox "请输入到站和辆数!"
        ElseIf rn.value = "" Then
            rn.value = TextBox1.value & TextBox2.value
        Else
            rn.value = rn.value & " " & TextBox1.value & TextBox2.value
        End If
        Application.EnableEvents = True
    Else
        MsgBox "选中要输入信息的单元格!"
    End If
End Sub

Private Sub CommandButton2_Click()
    Dim rn As Range
    If Selection.Column = 4 And Selection.Row > 3 And Selection.Row < 24 Then
        Set rn = Selection.Cells(1)
        Application.EnableEvents = False
        If rn.value = "" Then
            rn.value = ListBox2.Column(2, ListBox2.ListIndex)
        Else
            rn.value = rn.value & " " & ListBox2.Column(2, ListBox2.ListIndex)
        End If
    Else
        MsgBox "选中要输入信息的单元格!"
    End If
End Sub

Private Sub CommandButton3_Click()
    Dim fso As New FileSystemObject, t1 As Date, t2 As Date, path$, 文件名$, file$
    
    Application.ScreenUpdating = False
    Call 重置.Sheet6_清空股道
        
    '判断路径是否存在
    path = "C:\Program Files\SMIS2.6\TranData"
    If Not fso.FolderExists(path) Then End
    
    For Each ft In fso.GetFolder(path).files
'
'        'Debug.Print Ft '得到路径
        t1 = FileDateTime(ft) '得到创建时间
        If t1 > t2 Then
            t2 = t1
            file = ft
            文件名 = ft.name
        End If
        
    'Debug.Print Int(Date - filedate) '得到差值
    'Debug.Print Ft.Name
    'Debug.Print Ft.DateCreated '创建时间
    'Debug.Print Ft.DateLastModified '修改时间
    Next
    If file = "" Then
        MsgBox "未找到文件！"
    Else
        Call 统计股道存车.股道存车(file)
        Call Me.启动加载
        MsgBox "刷新完成！" & Chr(10) & 文件名 & ":  " & t2
    End If
    Application.ScreenUpdating = True
    
End Sub


Private Sub ListBox1_Change()
    
    If ListBox1.ListIndex <> -1 Then ' 判断是否有行被选中
        
        Dim 换长 As Double, 辆数 As Double
        换长 = 0
        
        ' 循环遍历所有被选中的行
        Dim i%, k%
        For i = 1 To ListBox1.ListCount - 1
            If ListBox1.Selected(i) Then
                辆数 = 辆数 + 1
                换长 = 换长 + CDbl(ListBox1.Column(4, i)) ' 第五列的索引为4
            End If
        Next
        
    End If
    
    Label3.Caption = 换长
    Label4.Caption = 辆数
    
End Sub

Private Sub ListBox2_Click()
    Dim trr, k%, M1$, M2$, M3 As Boolean, i, 换长, reg As New RegExp, 敞顶箱%, c%, x%, p%, g%, d%, b%
    reg.Pattern = "[\u4E00-\u9FFF]"
    
    With Me
        If Not reg.Test(.ListBox2.Column(2, .ListBox2.ListIndex)) Then M3 = True
        .ListBox3.Clear
        .ListBox3.list = Split(.ListBox2.Column(2, .ListBox2.ListIndex), " ")
        If .ListBox3.ListCount = 2 Then
            .ListBox3.Selected(1) = True
'            Debug.Print .ListBox3.List(1)
            Call LB3双击(.ListBox3.list(1))
        End If
        .ListBox1.Clear
        .ListBox1.MultiSelect = fmMultiSelectExtended
        .ListBox1.ColumnCount = 10
        trr = Array("股道", "序号", "车种", "  车号", "换长", "载重", " 到站", "方向", " 品名", "   记事")
        .ListBox1.AddItem
        .ListBox1.list(k, 0) = trr(0)
        .ListBox1.list(k, 1) = trr(1)
        .ListBox1.list(k, 2) = trr(2)
        .ListBox1.list(k, 3) = trr(3)
        .ListBox1.list(k, 4) = trr(4)
        .ListBox1.list(k, 5) = trr(5)
        .ListBox1.list(k, 6) = trr(6)
        .ListBox1.list(k, 7) = trr(7)
        .ListBox1.list(k, 8) = trr(8)
        .ListBox1.list(k, 9) = trr(9)
        
        .ListBox1.ColumnWidths = "25,25,30,40,25,25,40,25,40,100"
        
'        .ListBox1.List = sql.SQL查询("5")
        
        For i = 4 To UBound(arr, 1)
            M1 = arr(i, 1)
            M2 = .ListBox2.value
            If M1 = M2 Then
                k = k + 1
                .ListBox1.AddItem
                .ListBox1.list(k, 0) = arr(i, 1)
                .ListBox1.list(k, 1) = arr(i, 2)
                .ListBox1.list(k, 2) = arr(i, 3)
                .ListBox1.list(k, 3) = arr(i, 4)
                .ListBox1.list(k, 4) = arr(i, 6)
                .ListBox1.list(k, 5) = arr(i, 7)
                .ListBox1.list(k, 6) = arr(i, 8)
                .ListBox1.list(k, 7) = arr(i, 9)
                .ListBox1.list(k, 8) = arr(i, 10)
                .ListBox1.list(k, 9) = arr(i, 14)
                换长 = 换长 + CDbl(arr(i, 6))
                If M3 And InStr(arr(i, 14), "箱") > 0 Then
                    敞顶箱 = 敞顶箱 + 1
                ElseIf M3 And InStr(arr(i, 3), "X") > 0 Then
                    x = x + 1
                ElseIf M3 And InStr(arr(i, 3), "C") > 0 Then
                    c = c + 1
                ElseIf M3 And InStr(arr(i, 3), "P") > 0 Then
                    p = p + 1
                ElseIf M3 And InStr(arr(i, 3), "G") > 0 Then
                    g = g + 1
                ElseIf M3 And InStr(arr(i, 3), "D") > 0 Then
                    d = d + 1
                ElseIf M3 And InStr(arr(i, 3), "B") > 0 Then
                    b = b + 1
                End If
            End If
        Next
        
        .Label3.Caption = 换长
        .Label4.Caption = k
        If M3 And 敞顶箱 > 0 Then
            开关 = True
            .TextBox1.value = "敞顶箱" & 敞顶箱 & IIf(x > 0, "X" & x, "") & IIf(c > 0, "C" & c, "") & IIf(p > 0, "P" & p, "") & IIf(g > 0, "G" & g, "") & IIf(d > 0, "D" & d, "") & IIf(b > 0, "B" & b, "")
            .TextBox2.value = ""
            开关 = False
        End If
        .ListBox1.height = 310
    End With
End Sub

Private Sub ListBox3_Click()
'Private Sub ListBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Me
        If .ListBox3.value = "" Then
            .TextBox1.value = ""
            .TextBox2.value = ""
        Else
            Call LB3双击(.ListBox3.value)
        End If
    End With
    
End Sub
Sub LB3双击(Lbv$)
    Dim reg As New RegExp
    With Me
        If Not IsNull(.ListBox3.value) Then
            开关 = True
            reg.Pattern = "[^\d]+"
            .TextBox1.value = reg.Execute(Lbv)(0).value

            .Label7.Font.Size = 13
            .Label7.Top = 46
            On Error Resume Next
            fx = WorksheetFunction.VLookup(.TextBox1.value, Sheet2.Columns("A:B"), 2, 0) '方向
            On Error GoTo 0
            
            .Label7 = .TextBox1.value & "：" & fx
            If fx = "南口" Then
                .Label8.BackColor = 26367 '橙
            ElseIf fx = "沙口" Then
                .Label8.BackColor = 13395456 '蓝
            Else
                .Label7 = .TextBox1.value & "：" & "无"
                .Label8.BackColor = -2147483633 '无
            End If

            
            reg.Pattern = "\d+"
            If Not IsNull(.ListBox3.value) Then
                If reg.Test(.ListBox3.value) Then
                    .TextBox2.value = reg.Execute(.ListBox3.value)(0).value
                Else
                    .TextBox2.value = ""
                End If
            End If
        Else
            .TextBox1.value = ""
            .TextBox2.value = ""
        End If
        开关 = False
    End With
End Sub
Private Sub TextBox1_Change()
    Dim prr(), dz, i
    Dim reg As Object: Set reg = CreateObject("VBScript.RegExp")
    If 开关 Then Exit Sub
    With reg
        .IgnoreCase = True
        .Global = True
        .Pattern = "^[a-zA-Z]+$"
    End With
    
    On Error Resume Next '忽略错误继续执行VBA代码,避免出现错误消息
    If reg.Test(TextBox1.value) Then '英文字母
        For Each dz In d_py.keys
            
            If InStr(dz, TextBox1.value) > 0 Then
                i = i + 1
                ReDim Preserve prr(i)
                prr(i) = d_py(dz)
                
            End If
        Next
        Me.ListBox3.list = prr
        
    Else '中文
        For Each dz In d_py.keys
            
            If InStr(1, dz, TextBox1.value, vbTextCompare) > 0 Then
                i = i + 1
                ReDim Preserve prr(i)
                prr(i) = d_py(dz)
                
            End If
        Next
        Me.ListBox3.list = prr
        
    End If
    On Error GoTo 0 '以下恢复捕捉代码出现错误消息
End Sub

Private Sub UserForm_Initialize()

    Call Me.启动加载
    
End Sub

Sub 启动加载()
    Dim trr, brr(), i, k, rngs As Range, rng As Range
    If Sheet7.AutoFilterMode = True Then Sheet7.AutoFilterMode = False
    If Sheet6.AutoFilterMode = True Then Sheet6.AutoFilterMode = False
    Set rngs = Sheet6.UsedRange
    Set rng = rngs.Columns(7)
    arr = Sheet7.UsedRange
    
    trr = rngs
    rs = WorksheetFunction.CountA(rng)
'    ReDim brr(1 To UBound(trr), 1 To 3)
    ReDim brr(1 To rs, 1 To 3)
    For i = 1 To UBound(trr) - 1
        If trr(i, 7) <> "" Then
            k = k + 1
            brr(k, 1) = trr(i, 2)
            brr(k, 2) = trr(i, 4)
            brr(k, 3) = trr(i, 7)
        End If
    Next
    With Me.ListBox2
        .BorderStyle = fmBorderStyleSingle
        .ColumnCount = 5
        .ColumnWidths = "25;25;300"
        .list = brr
    End With
    
    Dim rn As Range, py$
    Set rng = Sheet2.Cells(Rows.count, 1).End(3).CurrentRegion.Columns(1)
    Set d_py = CreateObject("Scripting.Dictionary")
    For Each rn In rng.Cells

        py = 汉字转拼音.getqs(rn.value)
        d_py(rn.value & py) = rn.value
        
    Next
End Sub
