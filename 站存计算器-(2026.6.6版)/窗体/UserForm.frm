VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm 
   Caption         =   "钦州港运转-Momda"
   ClientHeight    =   2052
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10644
   OleObjectBlob   =   "UserForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "UserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ' 初始化日期选择窗体

    ' 获取当前日期
    Dim currentDate As Date
    currentDate = Date

    ' 设置年份选择框的范围
    Dim startYear As Integer
    Dim endYear As Integer
    startYear = Year(currentDate) - 1
    endYear = Year(currentDate) + 1
    For yearVal = startYear To endYear
        Me.ComboBoxYear.AddItem yearVal
    Next yearVal

    ' 设置月份选择框的范围
    For monthVal = 1 To 12
        Me.ComboBoxMonth.AddItem Format(monthVal, "00")
    Next monthVal

    ' 设置默认选中的日期为当前日期
    Me.ComboBoxYear.value = Year(currentDate)
    Me.ComboBoxMonth.value = Format(Month(currentDate), "00")
    ' 调用 UpdateDateComboBox 来更新日期选择框的范围
    UpdateDateComboBox
    Me.ComboBoxDate.value = Format(Day(Date) + 1, "00")
    
    '月度计划不显示日期
    If ActiveSheet.name = Sheet3.name Then
        Me.ComboBoxDate.Visible = False
        Me.ComboBoxYear.Left = Me.ComboBoxYear.Left + 60
        Me.ComboBoxMonth.Left = Me.ComboBoxMonth.Left + 80
    End If
End Sub

Private Sub ComboBoxYear_Change()
    ' 年份选择框内容变化时，更新日期选择框的范围
    UpdateDateComboBox
End Sub

Private Sub ComboBoxMonth_Change()
    ' 月份选择框内容变化时，更新日期选择框的范围
    UpdateDateComboBox
End Sub

Private Sub CommandButtonOK_Click()
    ' 确认选择日期并关闭窗体

    selectedDate = DateSerial(Me.ComboBoxYear.value, Me.ComboBoxMonth.value, Me.ComboBoxDate.value)

    'Debug.Print selectedDate
    ' 关闭窗体
    Unload Me
End Sub

Private Sub UpdateDateComboBox()
    ' 获取当前选择的年份和月份
    Dim selectedYear As Variant, selectedMonth As Variant, tepmDate As Variant
    
    selectedYear = CInt(Me.ComboBoxYear.value)
    selectedMonth = val(Me.ComboBoxMonth.value)
    tepmDate = val(Me.ComboBoxDate.value)
    ' 清空日期选择框的选项
    Me.ComboBoxDate.Clear

    ' 获取选定月份的最大天数
    Dim maxDay As Integer
    maxDay = Day(DateSerial(selectedYear, selectedMonth + 1, 0))

    ' 向日期选择框中添加选项
    For dayVal = 1 To maxDay
        Me.ComboBoxDate.AddItem Format(dayVal, "00")
    Next dayVal
    
    Me.ComboBoxDate.value = tepmDate
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        End '点击X关闭窗体
    End If
End Sub
