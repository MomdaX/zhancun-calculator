VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'curl 'http://10.190.166.62:5099/czhd/ctrlDl' \
'  -H 'Connection: keep-alive' \
'  -H 'Cache-Control: max-age=0' \
'  -H 'Upgrade-Insecure-Requests: 1' \
'  -H 'Origin: http://10.190.166.62:5099' \
'  -H 'Content-Type: application/x-www-form-urlencoded' \
'  -H 'User-Agent: Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36' \
'  -H 'Accept: text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9' \
'  -H 'Referer: http://10.190.166.62:5099/czhd/login.jsp' \
'  -H 'Accept-Language: zh-CN,zh;q=0.9' \
'  -H $'Cookie: JSESSIONID=pu-ePqtaRZDOu8lyp8rG7AJHlyWxY__Kko2qQMN7BPiAyyda5CjR\u0021-2045277198' \
'  --data-raw 'ysrq=20250316&dwlx=L&dw=QVZ&yh=1000&password=qzcwd12.' \
'  --compressed \
'  --insecure
Sub 测试()
    
    Dim http As Object
    
    url = "http://10.190.166.62:5099/czhd/ctrlDl"
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    With http

        data = "ysrq=20250316&dwlx=L&dw=QVZ&yh=1000&password=qzcwd12."
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .setRequestHeader "Origin", "http://10.190.166.62:5099"
        .setRequestHeader "Referer", "http://10.190.166.62:5099/czhd/login.jsp"
        '.setRequestHeader "Cookie", cookie
        .send data
        'Debug.Print .responseText
    End With
End Sub
