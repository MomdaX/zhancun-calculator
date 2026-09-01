Attribute VB_Name = "帐号校验"
Function 校对(用户名 As String)

'用户名 = Sheet2.Range("E1").Value

On Error Resume Next
帐号 = WorksheetFunction.VLookup(用户名, Sheet2.Range("G:I"), 2, 0)
密码 = WorksheetFunction.VLookup(用户名, Sheet2.Range("G:I"), 3, 0)
On Error GoTo 0

If 用户名 = "" Then
    MsgBox "请在Sheet2的E1单元格中输入用户名"
    End
ElseIf Not IsError(帐号) Then

    If 帐号 = "" Then: MsgBox "用户名：" & 用户名 & Chr(10) & "该用户没有权限！": End
    
    Debug.Print "用户名：" & 用户名 & "有权"
    校对 = Array(帐号, 密码)
    Sheet2.Range("E2").value = 帐号
    Sheet2.Range("E3").value = 密码

Else
    MsgBox "用户名：" & 用户名 & Chr(10) & "该用户没有权限！"
    End
End If


End Function
