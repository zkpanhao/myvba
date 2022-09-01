Sub 指定内容筛选()
    Dim arr2, arr1
    arr2 = Range("K2:K4")                                      '获取指定筛选条件
    ReDim arr1(1 To UBound(arr2))
    For i = 1 To UBound(arr2)                                  '转一维数组
        arr1(i) = CStr(arr2(i, 1))                             '注意数据类型
    Next
    Range("A:A").AutoFilter Field:=1, Criteria1:=arr1, Operator:=xlFilterValues
End Sub

Sub 多选表格()
    Dim arr
    arr = "数组来源"
    For i = 1 To UBound(arr)
        Worksheets(CStr(arr(i, 1))).Select Replace:=False
    Next
'多选后批量编辑
'Range("A4").Select
'Selection = "XXX"
End Sub

Sub 提取表名()
    Dim shname As String, arr1()
    ReDim arr1(Sheets.Count)
    For i = 1 To Sheets.Count
        shname = Worksheets(i).Name
        arr1(i) = shname
    Next
    Range("B1:B" & i - 1) = Application.Transpose(arr1)
End Sub


Sub 获取所有Environ变量参数()
    For i = 1 To 40
        Debug.Print i, Environ(i)
    Next
End Sub

Sub 自定义序列排序()
    Dim rng As Range, n As Long
    Set rng = Range("e2:e" & Cells(Rows.Count, "e").End(xlUp).Row) '自定义排序的规则
    Application.AddCustomList (rng) '增加一个自定义序列。
    n = Application.CustomListCount  '自定义序列的数目+1
    Range("a:c").Sort key1:=Range("a1"), _
                        order1:=xlAscending, _
                        Header:=xlYes, _
                        ordercustom:=n + 1 '按指定序列排序
    ActiveSheet.Sort.SortFields.Clear '清除排序痕迹，避免删除自定义规则后保存出错
    Application.DeleteCustomList n
End Sub
