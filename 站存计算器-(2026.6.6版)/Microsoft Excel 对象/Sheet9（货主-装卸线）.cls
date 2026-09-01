VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Call 更新.加载
    Sheet9.Range("I1").value = Now()
End Sub

Private Sub CommandButton2_Click()
    Call 重置.Sheet9_货主清空
    Sheet9.Range("I1").value = Now()
End Sub


Private Sub CommandButton_删除整行_Click()
    Selection.EntireRow.Delete
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
CommandButton_删除整行.Visible = False
If Target.Column = 5 Then
     With CommandButton_删除整行
        .Font.Size = 10
        .width = 30
        .height = 20
        .Left = Target.Left + Target.width - .width
        .Top = Target.Top
        .Visible = True
     End With
End If

End Sub

Private Sub 货主载入_Click()
'    Call 货主与装卸线.货主信息载入(Sheet7)
End Sub

Function 排序()
If Sheet9.AutoFilterMode Then
    Sheet9.AutoFilterMode = False
Else
'Sheet9.Range("A1:H1").AutoFilter.Sort.SortFields.Add Key:=Range("A1"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    With Sheet9.Range("A1:H1").AutoFilter.Sort
'        .Header = xlYes
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'End If

End Function


