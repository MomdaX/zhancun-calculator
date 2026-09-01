Attribute VB_Name = "铁路地图"
Function 获取getToken()

    Dim xhtp As Object, token As String
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1")
    With xhtp

        url = "http://10.208.2.72:8080/getToken?ip=10.208.2.72"
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
        .setRequestHeader "Referer", "http://10.208.2.72:8080/cljl"
        .send

        token = .responseText
    End With
    'Debug.Print token
    SaveSetting "Momda", "login", "铁路地图", token
    获取getToken = token
End Function


Function 获取中间站(编号) '编号 As Long
    Dim xhtp As Object, token As String
    
    'token = 获取getToken
    token = GetSetting("Momda", "login", "铁路地图") '获取授权
    '编号 = 1
    Set xhtp = CreateObject("WinHttp.WinHttpRequest.5.1")
    With xhtp

        url = "http://10.208.2.72:8080/getMidStationByLj?map_version=C260201&lj=" & 编号
        .Open "GET", url, False
        .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36"
        .setRequestHeader "Referer", "http://10.208.2.72:8080/cljl"
        .setRequestHeader "Authorization", token
        .send

        res = .responseText
    End With
    保存.text res, 编号
    获取中间站 = res

End Function

Sub 查询各铁路局中间站()
    Dim res As String, data As String
    
    data = "{"
    For i = 1 To 18
        data = data & i & ":"
        res = 获取中间站(i)
        data = data & res & ","
    Next
    
    保存.text data, "合并"
End Sub
