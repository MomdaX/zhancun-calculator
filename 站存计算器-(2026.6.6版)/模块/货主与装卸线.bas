Attribute VB_Name = "货主与装卸线"
Sub 货主信息载入(sht As Worksheet)
    Dim d As New Dictionary, dic As New Dictionary, d_品名 As New Dictionary, d_到站 As New Dictionary, d_车型 As New Dictionary, d_记事 As New Dictionary
    Dim 专用线 As String, rng As Range
    
    xrr = Sheet9.Range("A1").CurrentRegion
    If UBound(xrr) > 1 Then
        For i = 2 To UBound(xrr)
            key = xrr(i, 1) & "$" & xrr(i, 4)
            Item = xrr(i, 2) & xrr(i, 3) & xrr(i, 5) & xrr(i, 6) & xrr(i, 7) & xrr(i, 8)
            dic(key) = Item
        Next
    End If
    
    xrr = sht.Range("A1").CurrentRegion
    
    For i = 4 To UBound(xrr)
        
        If Not IsNumeric(xrr(i, 1)) And xrr(i, 13) <> "" And Mid(xrr(i, 1), 1, 1) <> "Y" Then
            Select Case True
            Case Join(Array("H1", "H2", "H3", "H4", "H5"), "$") Like "*" & xrr(i, 1) & "*"  '货场
                专用线 = "货场"
            Case Join(Array("YX2", "YX3"), "$") Like "*" & xrr(i, 1) & "*"
                专用线 = "永鑫"
            Case Join(Array("TS1", "TS2"), "$") Like "*" & xrr(i, 1) & "*"
                专用线 = "天盛煤"
            Case Join(Array("SH1", "SH2", "SH3", "SH4"), "$") Like "*" & xrr(i, 1) & "*"
                专用线 = "石化"
            Case Join(Array("GT1", "GT2"), "$") Like "*" & xrr(i, 1) & "*"
                专用线 = "国投"
            Case Join(Array("CZ3", "CZ4"), "$") Like "*" & xrr(i, 1) & "*"
                专用线 = "超智"
            Case Join(Array("GM1", "GM2", "GM3", "GM4"), "$") Like "*" & xrr(i, 1) & "*"
                专用线 = "广明"
            Case Join(Array("ZL1", "ZL2", "ZL3"), "$") Like "*" & xrr(i, 1) & "*"
                专用线 = "中粮"
            Case xrr(i, 1) = "LZ"
                专用线 = "大洋"
            Case Join(Array("G1", "G2"), "$") Like "*" & xrr(i, 1) & "*"
                专用线 = "港务局"
            Case Else
                Debug.Print xrr(i, 1)
            End Select
            
            
            '品名
            If Not d.Exists(专用线 & "$" & xrr(i, 13) & xrr(i, 10)) Then
                d(专用线 & "$" & xrr(i, 13) & xrr(i, 10)) = ""
                If Not dic(专用线 & "$" & xrr(i, 13)) Like "*" & xrr(i, 10) & "*" Then
                    If d_品名.Exists(专用线 & "$" & xrr(i, 13)) Then
                        d_品名(专用线 & "$" & xrr(i, 13)) = d_品名(专用线 & "$" & xrr(i, 13)) & "," & xrr(i, 10)
                    Else
                        d_品名(专用线 & "$" & xrr(i, 13)) = xrr(i, 10)
                    End If
                End If
            End If
            
            '到站
            If Not d.Exists(专用线 & "$" & xrr(i, 13) & xrr(i, 8)) Then
                d(专用线 & "$" & xrr(i, 13) & xrr(i, 8)) = ""
                'd_到站(专用线 & "$" & xrr(i, 13)) = d_到站(专用线 & "$" & xrr(i, 13)) & "，" & xrr(i, 8)
                If Not dic(专用线 & "$" & xrr(i, 13)) Like "*" & xrr(i, 8) & "*" Then
                    If d_到站.Exists(专用线 & "$" & xrr(i, 13)) Then
                        d_到站(专用线 & "$" & xrr(i, 13)) = d_到站(专用线 & "$" & xrr(i, 13)) & "," & xrr(i, 8)
                    Else
                        d_到站(专用线 & "$" & xrr(i, 13)) = xrr(i, 8)
                    End If
                End If
            End If
            
            '车型
            车种 = 正则_车种(xrr(i, 3))
            If Not d.Exists(专用线 & "$" & xrr(i, 13) & 车种) Then
                d(专用线 & "$" & xrr(i, 13) & 车种) = ""
                'd_车型(专用线 & "$" & xrr(i, 13)) = d_车型(专用线 & "$" & xrr(i, 13)) & "，" & 车种
                If Not dic(专用线 & "$" & xrr(i, 13)) Like "*" & 车种 & "*" Then
                    If d_车型.Exists(专用线 & "$" & xrr(i, 13)) Then
                        d_车型(专用线 & "$" & xrr(i, 13)) = d_车型(专用线 & "$" & xrr(i, 13)) & "," & 车种
                    Else
                        d_车型(专用线 & "$" & xrr(i, 13)) = 车种
                    End If
                End If
            End If
            
            '记事
            If Not d.Exists(专用线 & "$" & xrr(i, 13) & xrr(i, 14)) And xrr(i, 14) <> "" Then
                d(专用线 & "$" & xrr(i, 13) & xrr(i, 14)) = ""
                'd_记事(专用线 & "$" & xrr(i, 13)) = d_记事(专用线 & "$" & xrr(i, 13)) & Chr(10) & xrr(i, 14)
                If Not dic(专用线 & "$" & xrr(i, 13)) Like "*" & xrr(i, 14) & "*" Then
                    If d_记事.Exists(专用线 & "$" & xrr(i, 13)) Then
                        d_记事(专用线 & "$" & xrr(i, 13)) = d_记事(专用线 & "$" & xrr(i, 13)) & Chr(10) & xrr(i, 14)
                    Else
                        d_记事(专用线 & "$" & xrr(i, 13)) = xrr(i, 14)
                    End If
                End If
            End If
            
            
        End If
        专用线 = ""
    Next
    
    '写入Sheet9表格中
    For Each key In d_品名.keys
    With Sheet9
        If dic(key) <> "" Then
            For Each rng In .Range("A1").CurrentRegion.Columns(1).Cells
                If Split(key, "$")(0) = rng.value Then
                    'For i = rng.Row To .Range("A1").CurrentRegion.Rows.count
                    i = rng.Row
                    Do
                        If .Cells(i, 1).value = rng.value Then
                            If .Cells(i, 4).value = Split(key, "$")(1) Then
                                .Cells(i, "B") = .Cells(i, "B") & "," & d_车型(key)      '车型
                                .Cells(i, "E") = .Cells(i, "E") & "," & d_品名(key)      '品名
                                .Cells(i, "H") = .Cells(i, "H") & "," & d_到站(key)      '到站
                                .Cells(i, "F") = .Cells(i, "F") & Chr(10) & d_记事(key)  '记事
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                        i = i + 1
                    Loop
                    'Next
                End If
            Next
        Else
            r = .Cells(.Rows.count, 1).End(3).Row + 1
        '    专用线 = Split(Key, "$")(0)
        '    货主 = Split(Key, "$")(1)
                .Cells(r, "A") = Split(key, "$")(0) '专用线
                .Cells(r, "D") = Split(key, "$")(1) '货主
                .Cells(r, "B") = d_车型(key)        '车型
                .Cells(r, "E") = d_品名(key)        '品名
                .Cells(r, "H") = d_到站(key)        '到站
                .Cells(r, "F") = d_记事(key)        '记事
        End If
      End With
    Next
    
End Sub

Function 正则_车种(车种)
Dim reg As New RegExp

With reg
    .Global = 1
    .MultiLine = 1
    .Pattern = "[^\d]+"
End With
Set 类型 = reg.Execute(车种)(0)
For Each ls In Array("C", "X", "P", "G", "YW", "K", "B")
    If 类型 Like "*" & ls & "*" Then
        正则_车种 = ls
        Exit For
    Else
        正则_车种 = 类型
        Exit For
    End If
Next
End Function
