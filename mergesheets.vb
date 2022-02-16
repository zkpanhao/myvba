Sub 工作表合并()
    Dim Sht As Worksheet, Headrows As Byte, Cellrows As Long
    Dim Shtname
    Shtname = Worksheets("目录").Range("B2:B15")               '需要合并的表名放在指定单元格中，以实现精准合并
                                                                  '结合“为工作簿添加目录.vb”获取表名
    Application.ScreenUpdating = False                         '关闭屏幕刷新,提速措施
    Application.Calculation = xlCalculationManual              '自动计算改为手动,提速措施
    On Error Resume Next
    Set Sht = Worksheets("汇总")
    If Err <> 0 Then                                           '是否存在“汇总”表
        Worksheets.Add after:=Worksheets(Worksheets.Count)     '不存在则新建
        Worksheets(Worksheets.Count).Name = "汇总"
        Worksheets("汇总").Cells.Clear                           '清空汇总表
        Worksheets("汇总").Select                                '选中汇总表以便后续操作
    End If
    Headrows = 1                                               '设置表头行数
    Worksheets(CStr(Shtname(1, 1))).Rows("1:" & Headrows).Copy Worksheets("汇总").Cells(1, 1)  '复制表头到汇总表
    Columns("A:A").Insert                                      '插入一列用作存放被来源表的表名
    Cells(Headrows, 1) = "表名"
    For i = 1 To UBound(Shtname)
        If Shtname(1, 0) <> "汇总" Then
            With Worksheets(CStr(Shtname(i, 1)))
                With Intersect(.UsedRange.Offset(Headrows, 0), .UsedRange)    '向下偏移标题行同时与已使用区域取交集
                    Set Cell = Worksheets("汇总").Cells(Worksheets("汇总").UsedRange.Rows.Count + 1, 2)
                    .Copy Cell                                              '第一次复制格式
                    Cell.Resize(.Rows.Count, .Columns.Count) = .Value       '第二次复制数值
                    Cell.Offset(0, -1).Resize(.Rows.Count, 1).Merge         '合并第一列存放源表名
                End With
                Cell.Offset(0, -1) = .Name
            End With
        Else: MsgBox "不可以“汇总”作为源表"
        End If
    Next
    Application.ScreenUpdating = True                          '恢复屏幕刷新
    Application.Calculation = xlCalculationAutomatic           '恢复自动计算
End Sub
