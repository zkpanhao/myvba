Sub MergeBooks()
'学习重点
'FileDialog.Show/.SelectItem
'Dir的使用

    Dim PathStr As String, FileStr As String, File(), n As Integer
    Dim ActiveB As Workbook, Headline As Byte, Namess As String, Cell As Range,Tephead as String
    Dim vrtSelectedItem As Variant
'第一种写法
'    With Application.FileDialog(msoFileDialogFolderPicker)          '使用Application的Filedialog方法调出文件浏览对话框
'        If .Show = -1 Then                                          '.show=-1为用户点击“确定”
'            PathStr = .SelectedItems(1)
'        Else
'            Exit Sub
'        End If
'    End With
'    FileStr = Dir(PathStr & "\" & "*.csv")                          'Dir函数通过路径和使用通配符的文件名返回String值
'    While Len(FileStr) > 0
'        n = n + 1
'        ReDim Preserve File(1 To n)
'        File(n) = PathStr & "\" & FileStr                           '把文件夹下所有文件名写入数组
'        FileStr = Dir()                                             '多次使用Dir函数依次返回文件名
'    Wend
'另一种写法
    With Application.FileDialog(msoFileDialogFilePicker)       '选中多个文件
        If .Show = -1 Then
            For Each vrtSelectedItem In .SelectedItems
                n = n + 1
                ReDim Preserve File(1 To n)
                File(n) = vrtSelectedItem
            Next
        Else
            Exit Sub
        End If
    End With
'另一种写法结束
    Set ActiveB = ActiveWorkbook                               '把当前空表指定为ActiveB
    On Error Resume Next
    Tephead = InputBox("输入待合并的表的标题行数:")    '获取用户输入的标题行数
    If cbyte(Tephead) < 0 Then
        Exit Sub                                               '小于0退出程序
    ElseIf Tephead = "" Then                               '选择取消退出程序
        Exit Sub
    End If
    Headline=cbyte(Tephead)
    On Error GoTo 0
    Range("A1:B1") = Array("工作簿", "工作表")    '留两列放来源表名
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    For k = 1 To n                                             '第一层循环，遍历选中的所有文件
        Namess = Dir(File(k))                                  '获取文件名
        Workbooks.Open (File(k))                               '打开文件
        ActiveB.Activate
        If Headline = 0 Then GoTo Line0                        '标题行为0时跳过粘贴标题行的步骤
        If k = 1 Then Intersect(Workbooks(Namess).Worksheets(1).UsedRange, _
                Workbooks(Namess).Worksheets(1).Rows("1:" & Headline)).Copy Cells(1, 3)    '使用intersect的原因是为了copy的目标可以定位到单元格
Line0:
        For i = 1 To Workbooks(Namess).Worksheets.Count        '第二层循环，遍历每个工作簿里的工作表
            With Workbooks(Namess).Worksheets(i).UsedRange
                If .Rows.Count <= Headline Then GoTo Lines
                Set Cell = Cells(ActiveSheet.UsedRange.Rows.Count + 1, 3)
                Intersect(.Offset(Headline, 0), .Cells).Copy Cell    '第一次复制格式
                Cell.Resize(.Rows.Count - Headline, .Columns.Count) = Intersect(.Offset(Headline, 0), .Cells).Value    '第二次复制数值
                Cell.Offset(0, -2) = Namess                    '写入工作簿名
                Cell.Offset(0, -2).Resize(.Rows.Count - Headline, 1).Merge    '合并A列
                Cell.Offset(0, -1) = Workbooks(Namess).Worksheets(i).Name    '写入工作表名
                Cell.Offset(0, -1).Resize(.Rows.Count - Headline, 1).Merge    '合并B列
            End With
Lines:
        Next i
        Workbooks(Namess).Close False                          '关闭工作簿且不保存
    Next k
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
