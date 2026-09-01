Attribute VB_Name = "SQL"
Sub 获取最新的文件()
    
    Dim fso As New FileSystemObject, t1 As Date, t2 As Date, path$, 文件名$, file$

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
        
'        'Debug.Print Int(Date - filedate) '得到差值
'
'        'Debug.Print Ft.Name
'        'Debug.Print Ft.DateCreated '创建时间
'        'Debug.Print Ft.DateLastModified '修改时间
    Next
    Debug.Print 文件名 & ":" & t2
    Debug.Print file
    Call 统计股道存车.股道存车(file)
End Sub

Sub SQL查询()
    Dim cnn As Object, trr(), rst As Object, i%, rs%, rng As Range
    Set cnn = CreateObject("adodb.connection")
    
    股道 = InputBox("请输入股道号码")
    path = ThisWorkbook.FullName
    str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & path
    cnn.Open str_cnn '打开表格
    
    strSQL = "Select 股道,车种,车号,换长,载重,到站,方向,品名,收货人,记事 FROM [站存测试文件$A3:Q] WHERE [股道] IN ('" & 股道 & "')"
    Set rst = cnn.Execute(strSQL) '查找
    
    With ActiveSheet
            .Cells.Clear '重置
        For i = 1 To rst.fields.count
            .Cells(1, i) = rst.fields(i - 1).name '标题
        Next
            .Cells(1, 1).CurrentRegion.Font.Bold = True '标题加粗
            .Cells(2, 1).CopyFromRecordset rst '写入查找到的数据
            .Cells(1, 1).CurrentRegion.Borders.LineStyle = xlContinuous '设置样式边框
            .Cells(1, 1).CurrentRegion.EntireColumn.AutoFit
    End With
    
    rst.Close
    Set cnn = Nothing
End Sub
