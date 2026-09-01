Attribute VB_Name = "查询用户"
Sub 查询用户()
Dim xmlhttp As New MSXml2.ServerXMLHTTP
Dim HTML As String, url As String
Dim d As New Dictionary

    '帐号密码 = 帐号校验.校对(用户)
    帐号密码 = WorksheetFunction.Transpose(Sheet2.Range("E2:E3")) '启用时要改下标
'    Debug.Print "帐号：" & 帐号密码(1)
'    Debug.Print "密码：" & 帐号密码(2)
'    Debug.Print "加密密码：" & MD5加密.MD5Hash(帐号密码(2))
    加密密码 = MD5加密.MD5Hash(帐号密码(2))

    url = "http://10.190.128.96:8080/gaotie/login.do?empno=" & 帐号密码(1) & "&password=" & 加密密码
    key = 帐号密码(1) & 加密密码
    
    xmlhttp.Open "get", url, False
    xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
    xmlhttp.send

    url = "http://10.190.128.96:8080/gaotie/getSysnUserCnt.do"

    xmlhttp.Open "get", url, False
    xmlhttp.send

    HTML = xmlhttp.responseText
    'Debug.Print html
    
    Set Users = 提取(HTML, """users"":\[(.*)\]")
    If Users.count = 0 Then: Debug.Print "密码错误！": Sheet2.Range("E1").value = "密码错误！": End
    Set arr = 提取(Users(0).SubMatches(0), "(\{.*?\})")

    If arr.count = 0 Then
        Dim 提示 As String
        提示 = 提示 & Chr(10) & rng & rq & ":" & "未找到计划"
        Debug.Print 提示
    Else
        For i = 0 To arr.count - 1
        text = arr(i).SubMatches(0)
        'Debug.Print text
        
            Set brr = 提取(text, """(\w+?)"":(""[^""]*""|[^"",]*)")
    'd()
    'Debug.Print brr(4).SubMatches(1)
    姓名 = Replace(brr(4).SubMatches(1), """", "")
    帐号 = Replace(brr(3).SubMatches(1), """", "")
    密码 = Replace(brr(2).SubMatches(1), """", "")
    'IP = brr(8).SubMatches(1)
    'd(姓名) = 帐号 & " " & 密码
    
    d(帐号 & 密码) = 姓名
    Next
    End If
    
'    ActiveSheet.Range("G2").Resize(d.Count) = Application.Transpose(d.Keys)
'    ActiveSheet.Range("H2").Resize(d.Count) = Application.Transpose(d.Items)
    Debug.Print d(key)
    Sheet2.Range("E1").value = d(key)
    
End Sub


Function 提取(HTML, regex As String)
Dim reg As New RegExp
With reg
    .Pattern = regex
    .Global = True
    .MultiLine = True
End With

Set 提取 = reg.Execute(HTML)

End Function

Function 提取2(str As String, regex As String)
Dim reg As New RegExp
With reg
    .Pattern = regex
    .Global = True
    .MultiLine = True
End With
If reg.Test(str) Then
    提取2 = reg.Execute(str)(0).SubMatches(0)
Else
    提取2 = str
End If
End Function
