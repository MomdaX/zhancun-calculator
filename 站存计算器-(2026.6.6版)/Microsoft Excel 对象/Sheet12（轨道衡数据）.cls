VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

    If (Target.Column = 19 Or Target.Column = 20) And Target.Row > 2 Then
    
        Cancel = True
        请求地磅.查询地磅数据 Me.Cells(Target.Row, 19).value
    
    End If
End Sub



