Attribute VB_Name = "保存"
Function text(内容, fname)
    文件名 = "C:\Users\Administrator\Desktop\" & fname & ".txt"
    Open 文件名 For Output As #1
    Print #1, 内容
    Close #1
End Function

Function 自定义(内容, 文件名, 格式, 路径)
    path = 路径 & 文件名 & "." & 格式
    Open path For Output As #1
    Print #1, 内容
    Close #1
End Function

Function HTML(hStr)

    path = "C:\Users\Administrator\Desktop\HTML.html"
    
    With CreateObject("ADODB.Stream")
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .WriteText hStr
        .SaveToFile path, 2
        .Close
    End With
End Function


Function 图片Boy(img, path)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Type = 1
        .Open
        .Write img
        .SaveToFile path, 1 '下载验证码图片
        .Close
    End With
End Function
    
    
Function 图片byt(img, path)
    Open path For Binary As #1
    Put #1, , img
    Close #1 ' 关闭文件
End Function
