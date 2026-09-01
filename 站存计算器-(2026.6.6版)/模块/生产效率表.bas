Attribute VB_Name = "生产效率表"
Sub 生产效率表()

Sheet3.Range("A1").value = "钦州港站 " & Year(Date) & "年" & Month(Date) & "月" & Day(Date) & "日 生产效率表"

End Sub


'Sub 时间()
'
'
'If Time < TimeValue("12:00:00") Then
'Sheet1.Range("G1").Value = "钦州港 " & Month(Date) & " 月 " & Day(Date) & " 日 18点 站存"
'Else
'Sheet1.Range("G1").Value = "钦州港 " & Month(Date) & " 月 " & Day(Date) & " 日 早6点 站存"
'End If
'End Sub
