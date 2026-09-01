Attribute VB_Name = "更新"
Sub 刷新()
    Application.ScreenUpdating = False '取消屏幕刷新
    Application.DisplayAlerts = False '取消警告界面
    Application.AlertBeforeOverwriting = False '禁用提示框

    当前日期 = Date
    '触发条件=文件夹中的文件创建时间最新的比当前时间大于50分钟时执行
    pth = "C:\Program Files\SMIS2.6\TranData"

    If Sjc(pth) > 0 Then
        Call 更新.加载
        Call 导入.导入
        Sheet5.Range("L1") = 当前日期
        Call 删除过期文件.删除过期文件
        'Sheet1.Range("K4") = "更新时间：" & Now
        'Sheet1.Select '选中数据库表
    End If
    '保存工作簿
    'ThisWorkbook.Save

    Application.ScreenUpdating = True '开启屏幕刷新
    Application.DisplayAlerts = True '开启警告界面
    Application.AlertBeforeOverwriting = True '开启提示框
End Sub

'自定义涵数，获取文件创建时间
Function Sjc(paths)
    Dim fso As New FileSystemObject, t1
        t2 = Now
        Sjc = 0
    '判断路径是否存在
    If Not fso.FolderExists(paths) Then End

    For Each ft In fso.GetFolder(paths).files
'
'        'Debug.Print Ft '得到路径
        filedate = FileDateTime(ft)
        Debug.Print filedate '得到创建时间
'        'Debug.Print Int(Date - filedate) '得到差值
'
'        'Debug.Print Ft.Name
'        'Debug.Print Ft.DateCreated '创建时间
'        'Debug.Print Ft.DateLastModified '修改时间
        t = Now - filedate
     'Debug.Print Format(t, "HH")
        If Format(t, "HH") < 1 Then
            t1 = filedate
            Sjc = 0
        If t1 < t2 Then Sheet1.Range("K4") = "更新时间：" & filedate
            t2 = filedate
            Sjc = 1
        End If
    Next

End Function

Sub 加载()
    Dim d As New Dictionary, d1 As New Dictionary, d2 As New Dictionary
    Dim book As Object, sheet As Worksheet, fileName As String
    Dim dic As New Dictionary
    Dim Qan As Range, rngs As Range
    Dim rg As Range, rang As Range
    
    Application.ScreenUpdating = False '取消屏幕刷新
    Application.DisplayAlerts = False '取消警告界面
    
    path = "C:\Program Files\SMIS2.6\TranData"
    fileName = Dir(path & "\*.xls*")
    
    Do While fileName <> ""
        Set book = GetObject(path & "\" & fileName)
        'Set book = Workbooks.Open(path & "\" & fileName, ReadOnly:=True)
        Set sheet = book.Sheets(1)
        trr = sheet.UsedRange
    
        '提取货主信息
    '    Call 货主与装卸线.货主信息载入(sheet)
    
        '收集车站数据
        For i = 4 To UBound(trr)
        
            If Not d.Exists(trr(i, 8)) And trr(i, 8) <> "" Then
                ma = trr(i, 8)
                d(trr(i, 8)) = trr(i, 9)
            End If
        Next
        
        book.Close savechanges:=False
        fileName = Dir
    Loop
    
    '读取方向基础数据库
    Set rang = Sheet2.Range("A2:A" & Sheet2.Cells(Rows.count, 1).End(3).Row) 'sheet2中的数据库
    
    '将数据库载入字典d1中
    For Each rg In rang
        d1(rg.value) = rg.Offset(0, 1).value
    Next
    
    Set rng = Sheet5.Range("A1:A" & Sheet5.Cells(Rows.count, 1).End(3).Row) 'sheet5中的数据
    
    'Sheet2.Range("D1").Resize(d.count, 1) = WorksheetFunction.Transpose(d.Items)
    
    '载入sheet5中的数据
    For Each rg In rng
        If rg.value <> "到站" Then
        
            If d1.Exists(rg.value) Then '将以有的车站条目删除
            Set rngs = Sheet5.Range(rg, rg.Offset(0, 2))
            'rngs.Select
                rngs.Delete Shift:=xlUp
            Else
                If Not dic.Exists(rg.value) Then
                    d2(rg.value) = rg.Offset(0, 2).value '在数据库中没有的车站存入dic字典中
                End If
            End If
        
        End If
    Next
    
    For Each rn In d.keys
        Set Qan = Sheet5.Range("A" & Sheet5.Cells(Rows.count, 1).End(3).Row).Offset(1, 0) 'sheet5中的数据
    
        If Not d1.Exists(rn) And Not d2.Exists(rn) Then
            Qan.value = rn '到站
            Qan.Offset(0, 2).value = d(rn) '方向代号
            d2(rn) = d(rn) '新加入站存入d2数组中
        End If
    Next
    
    Set rng = Sheet5.Range("C2:C" & Sheet5.Cells(Rows.count, 1).End(3).Row)
    '根据代号分配方向
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
    
    '设置'sheet5中的格式
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
    
    Sheet5.Range("L1") = Date
    
    Application.ScreenUpdating = True '开启屏幕刷新
    Application.DisplayAlerts = True '开启警告界面
End Sub

