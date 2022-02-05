Sub MergeBooks()
    '学习重点
    'FileDialog.Show/.SelectItem
    'Dir的使用

    Dim PathStr As String, FileStr As String, File(), n As Integer
    With Application.FileDialog(msoFileDialogFolderPicker)          '使用Application的Filedialog方法调出文件浏览对话框
        If .Show = -1 Then                                          '.show=-1为用户点击“确定”
            PathStr = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    FileStr = Dir(PathStr & "\" & "*.csv")                          'Dir函数通过路径和使用通配符的文件名返回String值
    While Len(FileStr) > 0
        n = n + 1
        ReDim Preserve File(1 To n)
        File(n) = PathStr & "\" & FileStr                           '把文件夹下所有文件名写入数组
        FileStr = Dir()                                             '多次使用Dir函数依次返回文件名
    Wend
End Sub
