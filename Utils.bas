Attribute VB_Name = "Utils"
'union Range object
Function unionRange(ParamArray ranges() As Variant) As Range
    Dim result As Range
    
    For i = LBound(ranges) To UBound(ranges)
        If IsObject(ranges(i)) Then
            If Not ranges(i) Is Nothing Then
                If TypeOf ranges(i) Is Excel.Range Then
                    If result Is Nothing Then
                        Set result = ranges(i)
                    Else
                        Set result = Application.union(result, ranges(i))
                    End If
                End If
            End If
        End If
    Next
    
    Set unionRange = result
End Function

Function isWorksheetExist(ByVal name, Optional wb As Excel.Workbook) As Boolean
    Dim ws As Excel.Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    ' Method 1
    On Error Resume Next
    Set ws = wb.Sheets(name)
    ' disable error handler
    ' On Error GoTo 0
    isWorksheetExist = Not ws Is Nothing
    Exit Function

    ' Method 2
    On Error GoTo NotExist
    Set ws = wb.Worksheets(name)
    ' disable error handler
    ' On Error GoTo 0
    isWorksheetExist = True
    Exit Function
NotExist:
    isWorksheetExist = False
    Exit Function
    
    ' Method 3
    ' for case insensitive compare
    If VarType(name) = vbString Then name = LCase(name)
    
    For Each ws In Worksheets
        If name = ws.Index Or name = LCase(ws.name) Then
            isWorksheetExist = True
            Exit Function
        End If
    Next
    isWorksheetExist = False
End Function
