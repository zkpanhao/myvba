Sub 工作表拆分()
    Dim Splitcol As String, Headline As Byte, Indexsht As Byte, Rng As Range, Lastrow, Teparr
    Dim Keyword As New Collection
    Headline = CByte(InputBox("请输入[标题行]行数"))
    Splitcol = InputBox("请输入[拆分依据列]列号")                        '确定按哪一列进行拆分
    If Splitcol = "" Then Exit Sub
    Lastrow = ActiveSheet.UsedRange.Rows.Count
    Indexsht = ActiveSheet.Index
    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False            '如果活动表处于筛选模式则去除
    Teparr = Range(Splitcol & Headline + 1 & ":" & Splitcol & Lastrow).Value    '把筛选列的值写进临时数组
    On Error Resume Next
    For i = 1 To UBound(Teparr)                                '把数组里的值逐个写入集合，以达到去除重复值的目的
        If Len(CStr(Teparr(i, 1))) > 0 Then
            Keyword.Add CStr(Teparr(i, 1)), CStr(Teparr(i, 1))
        End If
    Next i
    Err = 0
    Application.ScreenUpdating = False
    For m = 1 To Keyword.Count                                 '按照集合的Item键值逐个创建新表
        Worksheets.Add after:=Worksheets(Worksheets.Count)
        Worksheets(Worksheets.Count).Name = Keyword(m)
        If Err <> 0 Then                                       '入股有同名表格先清空再复制标题
            Application.DisplayAlerts = False
            Worksheets(Worksheets.Count).Delete
            Worksheets(CStr(Keyword(m))).Cells.Clear
            Application.DisplayAlerts = True
        End If
        Err = 0
        Worksheets(Indexsht).Rows("1:" & Headline).Copy Worksheets(CStr(Keyword(m))).Rows("1")
    Next m
    Worksheets(Indexsht).Select
    Application.Calculation = xlCalculationManual
    For n = 1 To Keyword.Count                                 '逐表筛选复制数据
        Range(Cells(Headline, Splitcol), Cells(Lastrow, Splitcol)) _
                .AutoFilter Field:=1, Criteria1:=Keyword(n)
        Set Rng = Worksheets(Indexsht).Range(Cells(Headline + 1, Splitcol), Cells(Rows.Count, Splitcol).End(xlUp)). _
                SpecialCells(xlCellTypeVisible).EntireRow
        With Worksheets(CStr(Keyword(n))).Rows(Headline + 1)
            Rng.Copy .Cells(1)
            .Cells(1) = Rng.Value
            .Columns("A:P").AutoFit
        End With
    Next n
    Cells.AutoFilter
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub


