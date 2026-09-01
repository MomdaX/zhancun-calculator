VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
  Call 试卷成绩查询.导入题库_数组
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

    Dim rng As Range, rn As Range
    If Target.address = "$N$1" Then
        Cancel = True
        
        Excel转Word选择文件夹.Excel题库转到Word中

    End If
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Application.ScreenUpdating = False
    试卷成绩查询.选中行显示 Me, Me.Range("A1"), Target
    Application.ScreenUpdating = True
End Sub
