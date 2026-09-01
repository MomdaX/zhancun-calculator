Attribute VB_Name = "模拟键盘"
' ==============================================================================
' 功能：整合特殊按键模拟+中文输入（双方式）+ 非阻塞暂停
' 兼容：32/64位Office | 保留Sleep阻塞延迟 + NPause非阻塞延迟
' 说明：按键下标对应表见PressSpecialKey函数内注释
' ==============================================================================

' ========== API声明（32/64位兼容，保留Sleep）==========
#If VBA7 Then

    Private Declare PtrSafe Sub keybd_event Lib "user32" ( _
            ByVal bVk As Byte, _
            ByVal bScan As Byte, _
            ByVal dwFlags As LongPtr, _
            ByVal dwExtraInfo As LongPtr _
            )

    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
    ' 新增：剪贴板句柄操作API
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As LongPtr
#Else

    Private Declare Sub keybd_event Lib "user32" ( _
            ByVal bVk As Byte, _
            ByVal bScan As Byte, _
            ByVal dwFlags As Long, _
            ByVal dwExtraInfo As Long _
            )

    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    ' 新增：剪贴板句柄操作API
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
#End If

' ========== 常量定义（核心按键常量）==========
Private Const KEY_UP       As Byte = &H2    ' 按键释放标志
Private Const VK_SHIFT     As Byte = &H10   ' Shift键虚拟码
Private Const VK_CONTROL   As Byte = &H11   ' Ctrl键虚拟码

' 创建剪贴板对象
Private objData As Object

' ========== 1. 非阻塞暂停函数（界面可操作）==========
Sub NPause(seconds As Double)
    Dim startTime As Double
    startTime = Timer                      ' 记录开始时间戳（午夜起秒数）

    ' 循环等待，DoEvents释放CPU控制权
    Do While Timer < startTime + seconds
        DoEvents
    Loop
End Sub

' ========== 2. 模拟特殊按键（下标调用，精简易维护）==========
Sub PressSpecialKey( _
            keyIndex As Integer, _
            Optional delayMs As Long = 50 _
            )
    ' 按键配置数组：[虚拟键码, 是否需要Shift(0=否/1=是)]
    ' 下标对应关系：
    ' 1 → ENTER    回车键
    ' 2 → TAB      制表键
    ' 3 → ADD      主键盘+号（需Shift）
    ' 4 → ADD_NUMPAD 小键盘+号
    ' 5 → BACKSPACE 退格键
    ' 6 → ESC      退出键
    ' 7 → SPACE    空格键
    Dim keys As Variant
    keys = Array( _
            Array(&HD, 0), _
            Array(&H9, 0), _
            Array(&HBB, 1), _
            Array(&H6B, 0), _
            Array(&H8, 0), _
            Array(&H1B, 0), _
            Array(&H20, 0) _
            )

    ' 校验下标合法性
    If keyIndex < 1 Or keyIndex > UBound(keys) + 1 Then
        MsgBox "下标错误！仅支持1-" & UBound(keys) + 1 & "的下标", vbExclamation, "参数错误"
        Exit Sub
    End If
    delayMs = IIf(delayMs < 0, 0, delayMs)  ' 确保延迟非负

    ' 通用按键模拟逻辑
    Dim vk As Variant
    vk = keys(keyIndex - 1)

    ' 按下Shift（如需）
    If vk(1) = 1 Then
        keybd_event VK_SHIFT, 0, 0, 0
    End If

    ' 按下并释放目标键
    keybd_event vk(0), 0, 0, 0
    keybd_event vk(0), 0, KEY_UP, 0

    ' 释放Shift（如需）
    If vk(1) = 1 Then
        keybd_event VK_SHIFT, 0, KEY_UP, 0
    End If

    ' 阻塞延迟（兼容原有逻辑）
    If delayMs > 0 Then
        Sleep delayMs
    End If
End Sub

' ========== 3. 输入内容：方式1（SendKeys，极简）==========
Sub InputTextBySendKeys( _
            text As String, _
            Optional pauseSec As Double = 0.05 _
            )
    ' 参数校验
    If Len(text) = 0 Then
        MsgBox "输入文本不能为空！", vbExclamation, "参数错误"
        Exit Sub
    End If

    ' SendKeys发送文本（True=等待发送完成）
    SendKeys text, True

    ' 非阻塞暂停
    If pauseSec > 0 Then
        NPause (pauseSec)
    End If
End Sub

' ========== 4. 输入内容：方式2（剪贴板粘贴，稳定支持中文）==========
Sub InputTextByPaste( _
            text As String, _
            Optional pauseSec As Double = 0.05 _
            )
    ' 参数校验
    If Len(text) = 0 Then
        MsgBox "输入文本不能为空！", vbExclamation, "参数错误"
        Exit Sub
    End If

    ' 仅在对象未创建时初始化
    If objData Is Nothing Then
        Set objData = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    End If

    ' 写入目标文本到剪贴板（无需保存原有内容）
    objData.SetText text
    objData.PutInClipboard

    ' 模拟Ctrl+V粘贴
    keybd_event VK_CONTROL, 0, 0, 0
    keybd_event Asc("V"), 0, 0, 0
    keybd_event Asc("V"), 0, KEY_UP, 0
    keybd_event VK_CONTROL, 0, KEY_UP, 0

    ' 3. 显式关闭剪贴板（释放句柄，关键！）
    CloseClipboard
    
    ' 非阻塞暂停
    If pauseSec > 0 Then
        NPause (pauseSec)
    End If
End Sub

' ========== 测试示例：整合所有功能（可直接运行）==========
Sub TestAllFunctions()
    ' 示例1：粘贴方式输入中文（推荐，稳定）
    InputTextByPaste "姓名：张三", 0.1    ' 输入中文，非阻塞暂停0.1秒
    PressSpecialKey 2, 50                ' 按Tab键（下标2），阻塞延迟50ms
    InputTextByPaste "年龄：25", 0.1     ' 输入中文+数字

    ' 示例2：SendKeys方式输入简单内容
    PressSpecialKey 1, 50                ' 按回车键（下标1）
    InputTextBySendKeys "地址：北京市", 0.1 ' SendKeys输入中文（部分场景可能乱码）

    ' 示例3：混合按键+文本+非阻塞暂停
    PressSpecialKey 3, 50                ' 按主键盘+号（下标3）
    NPause 1                             ' 非阻塞暂停1秒（期间可操作Excel）
    InputTextByPaste "100000", 0.05      ' 输入数字
    PressSpecialKey 1, 0                 ' 按回车键（无延迟）
End Sub

Sub TestInput()
    ' 示例1：输入"Hello World!"，每个字符延迟100毫秒
    Dim rng As Range, dz As String, ch As String
    Set rng = ActiveSheet.Range("A5").CurrentRegion
    Debug.Print "开始！"
    NPause 3
    For i = 3 To rng.Rows.count
        dz = rng(i, 11).value
        ch = rng(i, 4).value
        InputTextByPaste ch, 0.2
        InputTextByPaste dz, 0.2
        PressSpecialKey 1, 50
    Next

    ' 释放对象
    Set objData = Nothing
End Sub


Function 监听()
    Do
        t = Application.OnKey("F6", 0)
        Debug.Print t
        If t = "F4" Then End Function
    Loop
End Function
