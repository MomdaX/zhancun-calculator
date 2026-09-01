Attribute VB_Name = "MD5加密"
Option Explicit
' 示例用法
Sub 加密()
    Dim password As String
'    password = Sheet2.Range("E3")
    password = "韦文康"
    Dim hashedPwd As String
    hashedPwd = MD5Hash(password)
    
    Debug.Print hashedPwd
End Sub
' 密码字符串的MD5哈希、转换为大写并截取指定长度的子字符串
Function MD5Hash(pwd) As String
    ' 创建MD5CryptoServiceProvider对象
    Dim md5Hasher As Object
    Set md5Hasher = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    
    ' 将密码字符串转换为字节数组
    Dim pwdBytes() As Byte
    pwdBytes = StrConv(pwd, vbFromUnicode)
    
    ' 计算MD5哈希值
    Dim hashBytes() As Byte
    hashBytes = md5Hasher.ComputeHash_2(pwdBytes)
    
    ' 将字节数组转换为十六进制字符串
    Dim i As Long
    Dim tempStr As String
    For i = LBound(hashBytes) To UBound(hashBytes)
        tempStr = tempStr & Right("0" & Hex(hashBytes(i)), 2)
    Next i
    
    ' 转换为大写并截取子字符串
    MD5Hash = UCase(Left(tempStr, 16))
End Function

'    url_Base64 = "http://10.190.128.231/static/js/base64/base64.min.js"
