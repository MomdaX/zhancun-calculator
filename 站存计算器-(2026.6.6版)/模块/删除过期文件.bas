Attribute VB_Name = "删除过期文件"
Sub 删除过期文件()
    pth = "C:\Program Files\SMIS2.6\TranData"
    Dim d As New Dictionary
    Dim fso As New FileSystemObject
    Set sh = ActiveSheet
    
    For Each ft In fso.GetFolder(pth).files
        
        'Debug.Print Ft '得到路径
        'filedate = FileDateTime(Ft)
        'Debug.Print filedate '得到创建时间
        'Debug.Print Int(Date - filedate) '得到差值
        
        'Debug.Print Ft.Name
        'Debug.Print Ft.DateCreated '创建时间
        'Debug.Print Ft.DateLastModified '修改时间
        Debug.Print "该文件创建至今：" & DateDiff("d", ft.DateCreated, Now) & "天"
        If DateDiff("d", ft.DateCreated, Now) > 3 Then
            
            Kill ft '删除文件
            Debug.Print "文件已清除"
            
        End If
        
    Next
End Sub

