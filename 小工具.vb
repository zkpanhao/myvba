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
'多选后批量编辑
'Range("A4").Select
'Selection = "XXX"
Next
End Sub
