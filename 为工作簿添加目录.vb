Sub 为工作簿添加目录()
    On Error Resume Next
    Worksheets("目录").Activate
    If Err = 0 Then
        GoTo Add
    Else
        Worksheets.Add(before:=Worksheets(1)).Name = "目录"
    End If
    On Error GoTo 0
Add:
    With Worksheets("目录")
        .Range("A1:B1") = Array("序号", "表名")
        For Each Sht In Sheets
            If Sht.Name <> "目录" Then
            i = i + 1
            .Range("A" & i + 1) = i
            .Range("B" & i + 1) = Sht.Name
            .Hyperlinks.Add anchor:=.Range("B" & i + 1), Address:="#" & Sht.Name & "!A1", _
                TextToDisplay:=Sht.Name
            End If
        Next
    End With
End Sub

Sub 为每个工作表增加_返回目录_()
    For Each sht In Sheets
        With Worksheets(sht.Name)
            .Range("F9") = "返回目录"
            .Hyperlinks.Add anchor:=.Range("F9"), Address:="#目录!A1", TextToDisplay:="返回目录"
        End With
    Next
End Sub
    
