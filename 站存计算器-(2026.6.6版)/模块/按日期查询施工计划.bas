Attribute VB_Name = "按日期查询施工计划"
Sub 按日期查询()
Dim xmlhttp As New MSXml2.ServerXMLHTTP
Dim link As String, path As String, 共用天窗 As String
Dim HTML As String, url As String, 车站 As String, 用户 As String
Dim match As MatchCollection
Dim d As New Dictionary, srr(2), crr()
Dim rng As Range, rngs As Range


'定位执行表
Sheet17.Select

'核对帐号
用户 = Sheet19.Range("E1").value
'Debug.Print "用户：" & 用户

'帐号密码 = 帐号校验.校对(用户)
帐号密码 = WorksheetFunction.Transpose(Sheet19.Range("E2:E3")) '启用时要改下标
If 帐号密码(1) = "" Or 帐号密码(2) = "" Then: MsgBox "请输入帐号或密码！": End
'Debug.Print "帐号：" & 帐号密码(1)
'Debug.Print "密码：" & 帐号密码(2)
'Debug.Print "加密密码：" & MD5加密.MD5Hash(帐号密码(2))

'打开窗体选择一个日期
UserForm.StartUpPosition = 0
UserForm.Top = Application.Top + (Application.height - UserForm.height) / 2
UserForm.Left = Application.Left + (Application.width - UserForm.width) / 2
UserForm.Show

Application.ScreenUpdating = False '取消屏幕刷新
Application.DisplayAlerts = False '取消警告界面

' 将选择的日期赋值给变量datek
If selectedDate = "0:00:00" Then
    Debug.Print "未选择时间"
    End
ElseIf IsDate(selectedDate) Then
    Debug.Print "我的日期是：" & Format(selectedDate, "yyyy年m月d日")
Else
    Debug.Print "日期错误！！！"
    End
End If
      
'设置标题
Title = Array("序号", "日计划" & Chr(10) & "编号", "月计划" & Chr(10) & "编号", "基本内容", "线路" & Chr(10) & "行别", "区间", "日期和时间", "施工内容" & Chr(10) & "及影响范围", "限速及" & Chr(10) & "行车方式变化", "路用列车" & Chr(10) & "信息", "施工单位" & Chr(10) & "及负责人", "配合单位" & Chr(10) & "及负责人", "计划" & Chr(10) & "提报单位", "登记站", "运输组织", "进出栅栏门" & Chr(10) & "编号", "运统-46", "备注", "状态", "主施工", "共用" & Chr(10) & "天窗号")
ActiveSheet.Cells.Clear
ActiveSheet.Range("A2").Resize(1, UBound(Title) + 1).value = Title
ActiveSheet.Range("A2").Resize(1, UBound(Title) + 1).Interior.ColorIndex = 8

'车站
Set rngs = Sheet19.Range("B1:B3")

'登录

    Set sdate = 提取(selectedDate, "\d+")
    If sdate.count <> 3 Then: MsgBox "时间格式不对！！！": End
    'ReDim srr(sdate.Count - 1)
    For sj = 0 To sdate.count - 1
        srr(sj) = sdate.Item(sj).value
    Next
    rq = Join(srr, ";")
    logurl = "http://10.190.128.96:8080/gaotie/login.do?empno=" & 帐号密码(1) & "&password=" & MD5加密.MD5Hash(帐号密码(2)) '密码经过加密
    'cookie = "JSESSIONID=9BCA38E2FA9613D55BD68803E84DE92C"
    xmlhttp.Open "get", logurl, False
    xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36"
    xmlhttp.setRequestHeader "Content-Type", "application/json;charset=UTF-8"
    'xmlhttp.setRequestHeader "Cookie", cookie
    xmlhttp.send
    
    '验证帐号和密码
    htm = xmlhttp.responseText
    'Debug.Print htm
    Set hrr = 提取(htm, """(\w+?)"":(""[^""]*""|[^""},]*)")
    登录状态 = hrr(2).SubMatches(1)
    If 登录状态 Like "*登录成功*" Then
        Debug.Print "登录状态" & 登录状态
    Else
        MsgBox 登录状态, 64, "钦州港运转-Momda"
        End
    End If
    
For Each rng In rngs

    车站 = URL编码.编码(rng.value)
    
    url = "http://10.190.128.96:8080/gaotie/management/queryDayPlan.do?queryStr=" & rq & ";allState;allGrade;allType;allDepartment;allLine;allStartStation;" & 车站 & ";allMonth;allMonth;all_do_unit;all_assist_unit;all_month_no;all_month_code;allMonth;allShift;allRailwayType&page=0" '&_=1689734604995"
    xmlhttp.Open "get", url, False
    xmlhttp.send
    
    HTML = xmlhttp.responseText
    'Debug.Print html
    
    HTML = 替换(HTML, "</lable>|<lable>|<a.*?>|</a>|vbCrLf", "")
    'ActiveSheet.UsedRange.Offset(1).ClearContents '清空区域
    
    Set arr = 提取(HTML, "(\{.*?\})")
    If arr.count = 0 Then
        Dim 提示 As String
        rqt = Replace(rq, ";", "/")
        提示 = 提示 & Chr(10) & rng & Format(rqt, "yyyy年m月d日") & ":" & "未找到计划"
        Debug.Print 提示
        ts = ts + 1
        If ts = 3 Then: MsgBox Format(rqt, "yyyy年m月d日") & "未找到施工计划": End
    Else
    
        '创建数组来存表，再导入，以加快速度！！！
        ReDim crr(1 To arr.count, 1 To 21)
        r = Cells(Rows.count, 1).End(3).Row + 1 '行号
        
        For i = 1 To arr.count
        text = arr(i - 1).SubMatches(0)
        'Debug.Print text
        
            'Set brr = 提取(text, """(\w+?)""\s*:\s*("".*?""|\d+)") '[""(.*?)""|(\d+?)]
            Set brr = 提取(text, """(\w+?)"":(""[^""]*""|[^""},]*)")
            'Debug.Print UBound(brr)

            For Each ca In brr
                n = n + 1
                If n = brr.count - 2 Then: Exit For
                'mmk = ca.SubMatches(0)
                    If ca.SubMatches(0) = "enable" And ca.SubMatches(1) <> "1" Then
                        m = 列号("memo")
                        共用天窗 = "共用天窗号：" & Chr(10) & 提取2(ca.SubMatches(1), """(.*?)""")
                        If crr(i, m) = "" Then
                            crr(i, m) = 共用天窗
                            GoTo 100:
                        Else
                            crr(i, m) = crr(i, m) & Chr(10) & Chr(10) & 共用天窗
                            GoTo 100:
                        End If
                    ElseIf ca.SubMatches(1) <> """""" And ca.SubMatches(1) <> "null" And ca.SubMatches(0) <> "enable" Then
                    
                        cellstr = 提取2(ca.SubMatches(1), """(.*?)""")
                        cellstr = Replace(cellstr, "<br>", Chr(10))
                        cellstr = Replace(cellstr, "<br />", Chr(10))
                        cellstr = Replace(cellstr, "</br>", Chr(10))
                        cellstr = Replace(cellstr, "：", "：" & Chr(10))
                        cellstr = Replace(cellstr, "至", "至" & Chr(10))
                        'cellstr = Replace(cellstr, "▲▲▲", "")
                        cellstr = Replace(cellstr, "\t", "")
                        cellstr = Replace(cellstr, "\r", "")
                        cellstr = Replace(cellstr, "\n", "")
                        
                    Else
                    
                        cellstr = ""
                    End If
                    
                '列号
                m = 列号(ca.SubMatches(0))
                
                If m <> "" Then: crr(i, m) = cellstr 'ActiveSheet.Cells(R, m).Value = cellstr
                
                'If cellstr = "无" Then: ActiveSheet.Cells(R + i - 1, m).HorizontalAlignment = xlCenter '有“无”的单元格居中
                  
            Next
            
100:
            n = 0
        Next
        
    '将crr中的数据扩展入表格中
    Sheet17.Range("A" & r).Resize(arr.count, 21) = crr
    
    End If
     
Next

'没有施工计划时，退出！！！
If ActiveSheet.Range("A3") = "" Then: MsgBox Format(rq, "yyyy年m月d日") & "未找到施工计划": End

'获取页面当前的宽度，计算权重
页面大小 = 190
ActiveSheet.Range("C:C, P:Q, S:S").Delete
Call 删除空列
列宽 = 总列宽
可调整的范围 = 页面大小 - 列宽

'///////////////////
'删除列后再计算列宽
colA = ActiveSheet.Cells(2, Columns.count).End(1).Column
Set rngs = ActiveSheet.Range(Cells(2, 1), Cells(2, colA))
'rngs.Select
For Each rng In rngs
     Select Case 替换(rng.value, "\s*", "")
        Case "基本内容"
        rng.ColumnWidth = rng.ColumnWidth + Int(可调整的范围 * 0.2)
        Case "区间"
        rng.ColumnWidth = rng.ColumnWidth + Int(可调整的范围 * 0.3)
        Case "施工内容及影响范围"
        rng.ColumnWidth = rng.ColumnWidth + Int(可调整的范围 * 0.5)
        Case "运输组织"
        
        '整列都要是"无"时，删除整列
        Wcount = WorksheetFunction.CountIf(Columns(rng.Column), "无")
            If Wcount = arr.count Then
                Columns(rng.Column).Delete '整列删除
            Else
                Columns(rng.Column).HorizontalAlignment = xlCenter '整列居中
            End If
        Case "登记站"
        Dim gdz As Range
            Do '不要港东的计划
            Set gdz = ActiveSheet.Columns(rng.Column).Find(Sheet19.Range("B4").value)
                If Not gdz Is Nothing Then
                    'gdz.Select
                    Rows(gdz.Row).Delete
                End If
            Loop While Not gdz Is Nothing
        Case Else
    End Select
    Newcw = 0
Next

'///////////////////
'设置边框和自动换行
With Rows(2) '标题格式化
    .AutoFit
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    '.Font.Size = 11
    .Font.Bold = True
    .Orientation = xlHorizontal
End With
ActiveSheet.Range("A2").CurrentRegion.Borders.LineStyle = xlContinuous
ActiveSheet.Cells.WrapText = True

Call 调整页面格式       '同时设置打印标题

'第一列合并居中
cols = ActiveSheet.Range("A2").CurrentRegion.Columns.count
ActiveSheet.Range(Cells(1, 1), Cells(1, cols)).Merge

'所有行自适应行高
ActiveSheet.Rows.EntireRow.AutoFit
ActiveSheet.Rows(2).RowHeight = 36 '标题行

'设置页头名称
rq = Replace(rq, ";", "/")
ActiveSheet.Range("A1").value = Format(rq, "yyyy年m月d日") & "施工计划"
ActiveSheet.Range("A1").Font.Size = 18
ActiveSheet.Rows(1).RowHeight = 29

'设置打印区域

Application.ScreenUpdating = True '开启屏幕刷新
Application.DisplayAlerts = True '开启警告界面

'行高+10
Call 设置行高.行高自动
Call 设置行高.行高自动

MsgBox Format(rq, "yyyy年m月d日") & "施工计划已全部显示" & 提示, 64, "钦州港运转-Momda"

'清空字典
arr = Array()
brr = Array()
Erase crr
Erase srr
hrr = Array()

End Sub
Function 提取(HTML, regex As String)
Dim reg As New RegExp
With reg
    .Pattern = regex
    .Global = True
    .MultiLine = True
End With

Set 提取 = reg.Execute(HTML)

End Function

Function 替换(HTML, regex As String, reptext As String)
Dim reg As New RegExp
With reg
    .Pattern = regex
    .Global = True
    .MultiLine = True
End With

替换 = reg.Replace(HTML, reptext)

End Function
Function 提取2(str As String, regex As String)
Dim reg As New RegExp
With reg
    .Pattern = regex
    .Global = True
    .MultiLine = True
End With
If reg.Test(str) Then
    提取2 = reg.Execute(str)(0).SubMatches(0)
Else
    提取2 = str
End If
End Function
Function 删除空列()
Dim rngs As Range, s As Range
Set rngs = Range("A2").CurrentRegion.Offset(1)
'rngs.Select
cols = UBound(Application.Transpose(rngs.Columns), 1)
    For i = 1 To cols
    Set s = rngs.Columns(i)
    's.Select

        If WorksheetFunction.CountA(s) = 0 Then
            'Debug.Print "空"
            s.EntireColumn.Delete
            i = i - 1
        End If
        
        Set rng = Range("A2").CurrentRegion
        If i = UBound(Application.Transpose(rng.Columns), 1) Then: Exit Function
        
    Next
End Function

Function 竖排文字(rng As Range)
    With rng
        '.AutoFit
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = xlVertical
        .WrapText = True
    End With
End Function

Function 调整页面格式()
With ActiveSheet.PageSetup 'Active
    .Orientation = xlLandscape '纸张横向
    '页边距
    .LeftMargin = 15
    .RightMargin = 15
    .TopMargin = 12
    .BottomMargin = 12
    .HeaderMargin = 10
    .FooterMargin = 10
    '页面居中
    .CenterHorizontally = True '页面居中
    .CenterVertically = False
    '页面缩放
    .Zoom = False
    .FitToPagesWide = 1
    .FitToPagesTall = False
    '设置打印标题
    .PrintTitleRows = "$2:$2"
    .PrintTitleColumns = ""
    .CenterHeader = ""
    '设置打印分辨率
'    .PrintQuality = 600
End With
    'Application.PrintCommunication = True
End Function

Function 总列宽()
cols = ActiveSheet.Range("A2").CurrentRegion.Columns.count
For i = 1 To cols
    cw = cw + Cells(1, i).ColumnWidth
Next
总列宽 = cw
End Function

Function 列号(key As String)

Select Case key
    Case "dayPlanNo"
    列号 = 1
    ActiveSheet.Columns(列号).ColumnWidth = 7 '序号
    ActiveSheet.Columns(列号).HorizontalAlignment = xlCenter
    Case "dayPlanCode"
    列号 = 2
    ActiveSheet.Columns(列号).ColumnWidth = 9 '日计划编号
    ActiveSheet.Columns(列号).HorizontalAlignment = xlCenter
    Case "monthPlanCode"
    列号 = 3
    ActiveSheet.Columns(列号).ColumnWidth = 7 '月计划编号
    Case "basicContent"
    列号 = 4
    ActiveSheet.Columns(列号).ColumnWidth = 13 '基本内容
    ActiveSheet.Columns(列号).HorizontalAlignment = xlCenter
    Case "line_Trend"
    列号 = 5
    ActiveSheet.Columns(列号).ColumnWidth = 5 '线路行别
    Call 竖排文字(ActiveSheet.Columns(列号))
    Case "section"
    列号 = 6
    ActiveSheet.Columns(列号).ColumnWidth = 9 '区间
    Case "date_Time"
    列号 = 7
    ActiveSheet.Columns(列号).ColumnWidth = 11 '日期和时间
    Case "content_Influence"
    列号 = 8
    ActiveSheet.Columns(列号).ColumnWidth = 30 '施工内容及影响范围
    Case "limit_speed"
    列号 = 9
    ActiveSheet.Columns(列号).ColumnWidth = 15 '限速及行车方式变化
    Case "train_message_type"
    列号 = 10
    ActiveSheet.Columns(列号).ColumnWidth = 15 '路用列车信息
    Case "unit_officer"
    列号 = 11
    ActiveSheet.Columns(列号).ColumnWidth = 11.5 '施工单位及负责人
    Case "assist_unit_and_officer"
    列号 = 12
    ActiveSheet.Columns(列号).ColumnWidth = 11.5 '配合单位及负责人
    Case "submit_unit"
    列号 = 13
    ActiveSheet.Columns(列号).ColumnWidth = 9 '计划提报单位
    Call 竖排文字(ActiveSheet.Columns(列号))
    Case "register_station"
    列号 = 14
    ActiveSheet.Columns(列号).ColumnWidth = 7 '登记站
    Call 竖排文字(ActiveSheet.Columns(列号))
    Case "transprot_organ"
    列号 = 15
    ActiveSheet.Columns(列号).ColumnWidth = 15 '运输组织
    Case "gate_no"
    列号 = 16
    ActiveSheet.Columns(列号).ColumnWidth = 5 '进出栅栏门编号
    Case "memo"
    列号 = 18
    ActiveSheet.Columns(列号).ColumnWidth = 12 '备注
    Case "state"
    列号 = 19
    ActiveSheet.Columns(列号).ColumnWidth = 10 '状态
    Call 竖排文字(ActiveSheet.Columns(列号))
    Case "run_statistics_46"
    列号 = 17
    ActiveSheet.Columns(列号).ColumnWidth = 10 '运统 -46
    ActiveSheet.Columns(列号).HorizontalAlignment = xlCenter
    Case "isread"
    列号 = 20
    ActiveSheet.Columns(列号).ColumnWidth = 5 '主施工
    Case "enable"
    列号 = 21
    ActiveSheet.Columns(列号).ColumnWidth = 6 '共用天窗号
    ActiveSheet.Columns(列号).HorizontalAlignment = xlCenter
    'Case "main_plan"
    '列号 = 22
    'Case "main_plan_id"
    '列号 = 23
    'Case "reset_main"
    '列号 = 24
End Select
End Function
