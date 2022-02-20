'在Excel表里使用正则表达式，提供“提取”和“替换”两种模式
Function Re(ByVal Tar As Range, Pt As String, Mold As Byte, Optional Rpstr As String) As String
    'Tar 为目标单元格
    'Pt 为正则表达式
    'Mold 为函数模式，1为提取匹配到的字符，2为替换
    'Rpstr 可选参数，当Mold为2时使用，用来替换掉匹配到的字符
    Dim Rex As Object
    Set Rex = CreateObject("VBScript.RegExp")
    With Rex
        .Global = True
        .IgnoreCase = True
        .Pattern = Pt
        If Mold = 1 Then                                       'Mold等于1为提取模式
            Re = .Execute(Tar)(0)
        ElseIf Mold = 2 Then                                   'Mold等于2为替换模式
            Re = .Replace(Tar, CStr(Rpstr))
        End If
    End With
End Function