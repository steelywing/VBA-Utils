Attribute VB_Name = "File"
Function exists(ByVal file As String) As Boolean
    exists = Dir(file) <> ""
End Function

Sub copy(ByVal src As String, ByVal dst As String)
    On Error GoTo Fail
    FileCopy src, dst
    copy = True
    Exit Function
Fail:
    copy = False
End Sub

Function remove(ByVal file As String)
    If Dir(file) = "" Then GoTo Fail
    
    On Error GoTo Fail
    SetAttr file, vbNormal
    Kill file
    remove = True
    Exit Function
Fail:
    remove = False
End Function

Sub rename(ByVal src As String, ByVal dst As String)
    On Error GoTo Fail
    Name src As dst
    rename = True
    Exit Function
Fail:
    rename = False
End Sub

