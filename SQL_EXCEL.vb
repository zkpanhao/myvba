Sub SQL_EXCEL()
    Dim cnn As Object, rst As Object
    Dim StrFile As String, str_cnn As String
    Dim StrSQL As String
    '打开数据源文件
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls,*.xlsx"
        .Filters.Add "All Files", "*.*"
        If .Show = -1 Then
            '获取StrFile文件名
            StrFile = Right(.SelectedItems(1), Len(.SelectedItems(1)) - InStrRev(.SelectedItems(1), "\", , vbTextCompare))
            Workbooks.Open StrFile
        End If
    End With
    '创建数据库链接
    Set cnn = CreateObject("adodb.connection")
    If Application.Version < 12 Then
        str_cnn = "Provider=Microsoft.jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=" & StrFile
    Else
        str_cnn = "Provider=Microsoft.ACE.OLEDB.12.0;Extended Properties=Excel 12.0;Data Source=" & StrFile
    End If
    cnn.Open str_cnn
    '创建一个数据集
    Set rst = CreateObject("adodb.recordset")
    StrSQL = Trim(Application.InputBox("请输入SQL语句", "SQL语句输入框")) '获取查询语句
    Set rst = cnn.Execute(StrSQL)                              '执行查询
    '清理数据呈现区域
    With Workbooks("EXCEL使用SQL查询模板.xlsm")
        .Worksheets("查询结果").Cells.ClearContents
        '输出结果
        For i = 0 To rst.Fields.Count - 1
            .Worksheets("查询结果").Cells(1, i + 1) = rst.Fields(i).Name '输出表头
        Next
        .Worksheets("查询结果").Range("A2").CopyFromRecordset rst  '输出Recordset数据
    End With
    '重置数据集
    cnn.Close
    Set cnn = Nothing
    Workbooks("EXCEL使用SQL查询模板.xlsm").Activate
End Sub
