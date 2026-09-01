VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Activate()

    Call SQL刷新
'    Call sql.获取最新的文件
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Application.ScreenUpdating = False
    If Target.Column = 15 And Target.Cells(1) <> "" And Not Target.Row = 2 Then
        Dim cnn As Object, trr(), rst As Object, i%, rs%, rng As Range
        Set cnn = CreateObject("adodb.connection")
    
        股道 = Target.value
        path = ThisWorkbook.FullName
        str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & path
        cnn.Open str_cnn '打开表格
    
        strSQL = "Select 股道,车种,车号,换长,载重,到站,方向,品名,收货人,记事 FROM [站存测试文件$A3:Q] WHERE [股道] IN ('" & 股道 & "')"
        Set rst = cnn.Execute(strSQL) '查找
            With Sheet26
                .Cells.Clear '重置
            For i = 1 To rst.fields.count
                .Cells(1, i) = rst.fields(i - 1).name '标题
            Next
                .Cells(1, 1).CurrentRegion.Font.Bold = True '标题加粗
                .Cells(2, 1).CopyFromRecordset rst '写入查找到的数据
                .Cells(1, 1).CurrentRegion.Borders.LineStyle = xlContinuous '设置样式边框
'                .Cells(1, 1).CurrentRegion.EntireColumn.AutoFit
                Call SQL刷新
        End With
    
        rst.Close
        Set cnn = Nothing
    End If
    Application.ScreenUpdating = True
End Sub

Sub SQL刷新()

    Dim cnn As Object, trr(), rst As Object, i%, rs%, rng As Range, rn As Range
    Set cnn = CreateObject("adodb.connection")
    
    Application.ScreenUpdating = False
    
    path = ThisWorkbook.FullName
    str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & path
    cnn.Open str_cnn '打开表格

    strSQL = "Select distinct 股道,count(股道)as 车数 FROM [站存测试文件$A3:Q] GROUP BY 股道 order by 股道 asc"
    Set rst = cnn.Execute(strSQL) '查找
    k = 2
    With Sheet26

        rst.MoveFirst
        .Columns(16).Clear  '清除
        .Cells.Font.ColorIndex = 1
        With .Columns(15)
            .Clear '清除
            .Cells(k).value = "选择股道"
            .HorizontalAlignment = xlCenter '设置上下居中
            .VerticalAlignment = xlCenter '设置左右居中
        End With
        Do Until rst.EOF
            k = k + 1
            .Cells(k, "O").value = rst.fields(0).value
            .Cells(k, "P").value = rst.fields(1).value
'            .Cells(k, "Q").value = rst.fields(2).value
            rst.MoveNext
        Loop
        
        Set rng = .Range("A1").CurrentRegion.Columns(6)
        For Each rn In rng.Cells
            Debug.Print rn.value
            If rn.value <> "" And rn.value <> "到站" Then
                显示信息.标记到站方向颜色 rn
            End If
        Next
    End With
    
    rst.Close
    Set cnn = Nothing
    
    
    Application.ScreenUpdating = True
End Sub
