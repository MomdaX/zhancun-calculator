Attribute VB_Name = "OCR"
'Option Explicit PtrSafe
'Declare Sub Init Lib "C:\Users\Administrator\Desktop\每日站存\附件\DdddOcr.dll" ()
'Declare Sub Shutdown Lib "C:\Users\Administrator\Desktop\每日站存\附件\DdddOcr.dll" Alias "Close" ()
'Declare Function Classification Lib "C:\Users\Administrator\Desktop\每日站存\附件\DdddOcr.dll" (ByRef aa As Byte) As LongPtr
'Private Declare Function lstrlenA Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long
'Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As Long)
Sub GetimageFullPath()
    Dim fd As Office.fileDialog
    Dim strFile As String
    Set fd = Application.fileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.Clear
        .Filters.Add "Images", "*.png, *. jpg"
        .Title = "选择要识别的图片!"
        .AllowMultiSelect = False
        If .Show = True Then
            strFile = .SelectedItems(1)
            'GetimageFullPath = strFile
            MsgBox (GetStr(strFile))
        End If
    End With
End Sub
Sub main()
    Dim path As String, i As Long, pathbase As String
    pathbase = ThisWorkbook.path
    For i = 1 To 8
        path = ThisWorkbook.path & "\pics\" & i & ".png"
        Debug.Print (GetStr(path))
        Debug.Print
    Next
End Sub
'获取识别结果
Function GetStr(path As String)
    Dim address, str As String, bytearr() As Byte
    str = Pic2Base64(path)
    bytearr = StrConv(str, vbFromUnicode)
    Init
        address = Classification(bytearr(0))
        GetStr = StringFromPointerA(address)
    Shutdown
End Function

'图片(地址)转换为Base64
Public Function Pic2Base64(strPicPath As String) As String
    Const adTypeBinary = 1 ' Binary file is encoded
    Dim objXML, objDocElem, objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = adTypeBinary
    objStream.Open
    objStream.LoadFromFile (strPicPath)
    Set objXML = CreateObject("MSXml2.DOMDocument")
    Set objDocElem = objXML.createElement("Base64Data")
    objDocElem.DataType = "bin.base64"
    objDocElem.nodeTypedValue = objStream.Read()
    Pic2Base64 = objDocElem.text
    objStream.Close
    Set objXML = Nothing
    Set objDocElem = Nothing
    Set objStream = Nothing
End Function

'由字符串指针（内存地址）获取字符串
'Public Function StringFromPointerA(ByVal pointerToString As LongPtr) As String
'    Dim tmpBuffer()    As Byte
'    Dim byteCount      As Long
'    Dim retVal         As String
'    byteCount = lstrlenA(pointerToString)
'    If byteCount > 0 Then
'        ReDim tmpBuffer(0 To byteCount - 1) As Byte
'        Call CopyMemory(VarPtr(tmpBuffer(0)), pointerToString, byteCount)
'    End If
'    retVal = StrConv(tmpBuffer, vbUnicode)
'    StringFromPointerA = retVal
'End Function

Sub js()
Dim url As String
url = "http://10.190.168.62/nnjzj/user/exam/view.html?tpid=01J7RMGDZ8D563RXTND3MFPZ7S&egid=01J8V5PAR2CGSSZB22QHKM3W3S"
ThisWorkbook.FollowHyperlink url, , 1, 1, , , "Cookie: learning.session.id=3fcf55ae-626b-4d2b-b6e1-a6b5f70ca868"



End Sub
