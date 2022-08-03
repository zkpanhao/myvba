Sub 批量获取文件名()
    Cells = ""
    Dim sfso
    Dim myPath As String
    Dim Sh As Object
    Dim Folder As Object
    Application.ScreenUpdating = False
    On Error Resume Next
    Set sfso = CreateObject("Scripting.FileSystemObject")
    Set objFD = Application.FileDialog(msoFileDialogFolderPicker)
    With objFD
        If .Show = -1 Then
            ' 如果单击了确定按钮，则将选取的路径保存在变量中
            myPath = .SelectedItems(1)
        End If
    End With
    If Not Folder Is Nothing Then
        myPath = Folder.Items.Item.Path
    End If
    Application.ScreenUpdating = True
    Cells(1, 1) = "旧版名称"
    Cells(1, 2) = "文件类型"
    Cells(1, 3) = "所在位置"
    Cells(1, 4) = "新版名称"
    Call 直接提取文件名(myPath & "\")
End Sub
  
    '获取选定文件夹下的所有表格名称
Sub 直接提取文件名(myPath As String)
    Dim i As Long
    Dim myTxt As String
    i = Range("A1048576").End(xlUp).Row
    myTxt = Dir(myPath, 31)
    Do While myTxt <> ""
        On Error Resume Next
        If myTxt <> ThisWorkbook.Name And myTxt <> "." And myTxt <> ".." And myTxt <> "081226" Then    '判断是否是隐藏文件或文件夹
            i = i + 1
            Cells(i, 1) = "'" & myTxt
            If (GetAttr(myPath & myTxt) And vbDirectory) = vbDirectory Then
                Cells(i, 2) = "文件夹"
            Else
                Cells(i, 2) = "文件"
            End If
            Cells(i, 3) = Left(myPath, Len(myPath) - 1)
        End If
        myTxt = Dir
    Loop
End Sub

    '批量重新命名表格名称
Sub 批量重命名()
    Dim y_name As String
    Dim x_name As String
    For i = 2 To Range("A1048576").End(xlUp).Row
        y_name = Cells(i, 3) & "\" & Cells(i, 1)
        x_name = Cells(i, 3) & "\" & Cells(i, 4)
        On Error Resume Next
        Name y_name As x_name
    Next
End Sub
