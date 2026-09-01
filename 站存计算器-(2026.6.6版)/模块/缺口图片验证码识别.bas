Attribute VB_Name = "缺口图片验证码识别"
Option Explicit

'==================== 结构体 ====================
Private Type GDIPlusStartupInput
    GdiPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type rect
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

Private Type bitmapData
    width       As Long
    height      As Long
    Stride      As Long
    PixelFormat As Long
    scan0       As Long
    Reserved    As Long
End Type

'==================== API 声明（32/64 通用） ====================
Private Declare Function GdiplusStartup Lib "gdiplus" ( _
        ByRef token As Long, ByRef inputbuf As GDIPlusStartupInput, ByVal outputbuf As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)

Private Declare Function GdipCreateBitmapFromStream Lib "gdiplus" ( _
        ByVal stream As Long, ByRef bitmap As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" ( _
        ByVal Image As Long, ByRef width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" ( _
        ByVal Image As Long, ByRef height As Long) As Long
Private Declare Function GdipBitmapLockBits Lib "gdiplus" ( _
        ByVal bitmap As Long, ByRef rect As rect, ByVal flags As Long, _
        ByVal pixFmt As Long, ByRef lockedData As bitmapData) As Long
Private Declare Function GdipBitmapUnlockBits Lib "gdiplus" ( _
        ByVal bitmap As Long, ByRef lockedData As bitmapData) As Long

Private Declare Function CreateStreamOnHGlobal Lib "ole32" ( _
        ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ByRef ppstm As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" ( _
        ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" ( _
        ByVal hMem As Long) As Long
Private Declare Sub GlobalUnlock Lib "kernel32" (ByVal hMem As Long)
Private Declare Function GlobalFree Lib "kernel32" ( _
        ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByVal dst As Long, ByVal src As Long, ByVal length As Long)

'==================== 常量 ====================
Private Const PixelFormat32bppARGB As Long = &H26200A
Private Const ImageLockModeRead As Long = 1
Private Const GMEM_MOVEABLE As Long = &H2

'==========  传入 Base64，直接返回缺口 X ==========
Function PNG模式(ByVal base64Str As String) As Long
    '=== 1. Base64 → 字节数组 ===
    Dim xml As Object: Set xml = CreateObject("MSXML2.DOMDocument")
    Dim elem As Object: Set elem = xml.createElement("b64")
    elem.DataType = "bin.base64"
    elem.text = base64Str
    Dim byt() As Byte: byt = elem.nodeTypedValue

    '=== 2. GDI+ 初始化 ===
    Dim tok As Long, inp As GDIPlusStartupInput: inp.GdiPlusVersion = 1
    GdiplusStartup tok, inp, 0

    '=== 3. 流 → 位图 ===
    Dim hGlb As Long, pStr As Long, pImg As Long
    hGlb = GlobalAlloc(&H2, UBound(byt) + 1)
    CopyMemory GlobalLock(hGlb), VarPtr(byt(0)), UBound(byt) + 1
    GlobalUnlock hGlb
    CreateStreamOnHGlobal hGlb, 1, pStr
    GdipCreateBitmapFromStream pStr, pImg

    '=== 4. 取宽高 + 锁位图 ===
    Dim w As Long, h As Long
    GdipGetImageWidth pImg, w
    GdipGetImageHeight pImg, h
    Dim rct As rect
    rct.Left = 0
    rct.Top = 0
    rct.Right = w
    rct.Bottom = h
    Dim bd As bitmapData
    GdipBitmapLockBits pImg, rct, 1, &H26200A, bd
    ReDim result(0 To h - 1) As String

    '=== 5. 扫描像素 → 统计每行首个“1”的 x ===
    ' pScan  : 位图首地址
    ' Stride : 一行字节数（已取 Abs）
    ' w,h    : 图像宽高
    ' d      : Scripting.Dictionary 用于计数
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim x As Long, y As Long, first1Found As Boolean
    Dim Stride As Long: Stride = Abs(bd.Stride)
    Dim pScan As Long: pScan = bd.scan0
    
    Const refR As Long = 0, refG As Long = 0, refB As Long = 250
    Const TH As Long = 80
    Dim line() As Long
    For y = 0 To h - 1
        ReDim line(0 To w - 1)
        CopyMemory ByVal VarPtr(line(0)), ByVal pScan + y * Stride, w * 4

        For x = 0 To w - 1
            '蓝背景阈值：RGB(0,0,250) 距离 < 80 视为背景“1”
            If Sqr(((line(x) And &HFF&) - 250) ^ 2 + _
                   (((line(x) \ &H100) And &HFF&) - 0) ^ 2 + _
                   (((line(x) \ &H10000) And &HFF&) - 0) ^ 2) < 80 Then
                d(x) = d(x) + 1
                Exit For   '只要第一个1
            End If
        Next x
    Next y

    '=== 6. 取出现最多的 x ===
    Dim k As Variant, maxK As Variant
    maxK = d.keys()(0)
    For Each k In d
        If d(k) > d(maxK) Then maxK = k
    Next
    
    PNG模式 = maxK '导出

    '=== 7. 释放 ===
    GdipBitmapUnlockBits pImg, bd
    GdipDisposeImage pImg
    GdiplusShutdown tok
End Function

'==========  传入 Byte()，直接返回缺口 X ==========
Function JPG模式(ByVal byt) As Long

    '=== 1. GDI+ 初始化 ===
    Dim tok As Long, inp As GDIPlusStartupInput
    inp.GdiPlusVersion = 1
    GdiplusStartup tok, inp, 0

    '=== 2. 流 → 位图 ===
    Dim hGlb As Long, pStr As Long, pImg As Long
    hGlb = GlobalAlloc(GMEM_MOVEABLE, UBound(byt) + 1)
    CopyMemory GlobalLock(hGlb), VarPtr(byt(0)), UBound(byt) + 1
    GlobalUnlock hGlb
    CreateStreamOnHGlobal hGlb, 0, pStr
    GdipCreateBitmapFromStream pStr, pImg

    '=== 3. 取宽高 + 锁位图 ===
    Dim w As Long, h As Long
    GdipGetImageWidth pImg, w
    GdipGetImageHeight pImg, h
    Dim rct As rect
    rct.Left = 0
    rct.Top = 0
    rct.Right = w
    rct.Bottom = h
    Dim bd As bitmapData
    GdipBitmapLockBits pImg, rct, 1, &H26200A, bd
    ReDim result(0 To h - 1) As String

    '=== 4. 扫描像素 → 统计每行首个“1”的 x ===
    ' pScan  : 位图首地址
    ' Stride : 一行字节数（已取 Abs）
    ' w,h    : 图像宽高
    ' d      : Scripting.Dictionary 用于计数
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim x As Long, y As Long, first1Found As Boolean
    Dim Stride As Long: Stride = Abs(bd.Stride)
    Dim pScan As Long: pScan = bd.scan0
    
    Const refR As Long = 0, refG As Long = 0, refB As Long = 250
    Const TH As Long = 80
    Dim line() As Long
    For y = 0 To h - 1
        ReDim line(0 To w - 1)
        CopyMemory ByVal VarPtr(line(0)), ByVal pScan + y * Stride, w * 4

        For x = 0 To w - 1
            '蓝背景阈值：RGB(0,0,250) 距离 < 80 视为背景“1”
            If Sqr(((line(x) And &HFF&) - 250) ^ 2 + _
                   (((line(x) \ &H100) And &HFF&) - 0) ^ 2 + _
                   (((line(x) \ &H10000) And &HFF&) - 0) ^ 2) < 80 Then
                d(x) = d(x) + 1
                Exit For   '只要第一个1
            End If
        Next x
    Next y

    '=== 5. 取出现最多的 x ===
    Dim k As Variant, maxK As Variant
    maxK = d.keys()(0)
    For Each k In d
        If d(k) > d(maxK) Then maxK = k
    Next
    
    JPG模式 = maxK '导出

    '=== 7. 释放 ===
    GdipBitmapUnlockBits pImg, bd
    GdipDisposeImage pImg
    GdiplusShutdown tok
End Function
'Application.ExecuteExcel4Macro
