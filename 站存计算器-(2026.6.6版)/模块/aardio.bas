Attribute VB_Name = "aardio"
Public ad As Object
Public Function Getadobj(aardio As Object)
    Set ad = aardio
    Debug.Print "[" & Format(Now, "HH:mm:ss") & "]：" & ad.msg
End Function
Sub CallAnything()
    ad.cx "新晃"
End Sub


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim val As String
    
    ' 只响应 H 列
    If Target.Column <> 8 Then Exit Sub
    Cancel = True
    
    ' 获取单元格内容
    val = Trim(Target.value)
    
    ' 不为空时执行
    If val <> "" Then
        Dim exePath As String, exeName As String
        exeName = "铁路地图.exe"
        exePath = ThisWorkbook.path & "\" & exeName
    
        ' 强制关闭同名进程
        If ad Is Nothing Then
            CreateObject("WScript.Shell").Run "taskkill /f /im " & Chr(34) & exeName & Chr(34), 0, True
            ' 启动新程序
            CreateObject("WScript.Shell").Run Chr(34) & exePath & Chr(34), 1, False
        Else
            If IsPR(exeName) Then
                Set ad = Nothing
                CreateObject("WScript.Shell").Run Chr(34) & exePath & Chr(34), 1, False
            Else
                ad.cx val
            End If
        End If
    End If
End Sub

' 辅助函数：判断进程是否运行
Function IsPR(exeName As String) As Boolean
    On Error Resume Next
    With GetObject("winmgmts:\\.\root\cimv2").ExecQuery("select * from Win32_Process where Name='" & exeName & "'")
        Debug.Print .count
        IsPR = .count = 0
    End With
    On Error GoTo 0
End Function

