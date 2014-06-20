Attribute VB_Name = "Utils"
' for specialFolder()
Const SF_ALL_USERS_DESKTOP = "AllUsersDesktop"
Const SF_ALL_USERS_START_MENU = "AllUsersStartMenu"
Const SF_ALL_USERS_PROGRAMS = "AllUsersPrograms"
Const SF_ALL_USERS_STARTUP = "AllUsersStartup"
Const SF_DESKTOP = "Desktop"
Const SF_FAVORITES = "Favorites"
Const SF_FONTS = "Fonts"
Const SF_MY_DOCUMENTS = "MyDocuments"
Const SF_NET_HOOD = "NetHood"
Const SF_PRINT_HOOD = "PrintHood"
Const SF_PROGRAMS = "Programs"
Const SF_RECENT = "Recent"
Const SF_SEND_TO = "SendTo"
Const SF_START_MENU = "StartMenu"
Const SF_STARTUP = "Startup"
Const SF_TEMPLATES = "Templates"

' Union Range object, e.g.
' unionRange([A1:A8], [C1:C10])
Function unionRange(ParamArray ranges() As Variant) As Range
    Dim result As Range
    
    For Each r In ranges
        If IsObject(r) Then
            If Not r Is Nothing Then
                If TypeOf r Is Range Then
                    If result Is Nothing Then
                        Set result = r
                    Else
                        Set result = Application.union(result, r)
                    End If
                End If
            End If
        End If
    Next
    
    Set unionRange = result
End Function

' Sheets contain Worksheets and Charts
' Worksheets only contain Worksheets
' more detail http://blogs.msdn.com/b/frice/archive/2007/12/05/excel-s-worksheets-and-sheets-collection-what-s-the-difference.aspx

Function isSheetExist(ByVal name As Variant, Optional wb As Workbook) As Boolean
    Dim s As Variant
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    On Error Resume Next
    Set s = wb.Sheets(name)
    
    ' disable error handler, this can be omit because
    ' error handler only attach to current function
    ' On Error GoTo 0
    
    ' Variant default is Empty
    isSheetExist = Not s Is Empty
    Exit Function
End Function

Function isWorksheetExist(ByVal name As Variant, Optional wb As Workbook) As Boolean
    Dim ws As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    '----------
    ' Method 1
    '----------
    On Error Resume Next
    Set ws = wb.Worksheets(name)
    ' On Error GoTo 0
    isWorksheetExist = Not ws Is Nothing
    Exit Function

    '----------
    ' Method 2
    '----------
    On Error Resume Next
    ' because function default return value is False, so
    ' if this raise error, return False, otherwise return True
    isWorksheetExist = Not wb.Worksheets(name) Is Nothing
    ' On Error GoTo 0
    Exit Function

    '----------
    ' Method 3
    '----------
    On Error GoTo NotExist
    Set ws = wb.Worksheets(name)
    ' On Error GoTo 0
    isWorksheetExist = True
    Exit Function
NotExist:
    isWorksheetExist = False
    Exit Function
    
    '----------
    ' Method 4
    '----------
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

Function lastColumn(Optional ws As Worksheet) As Integer
    If ws Is Nothing Then Set ws = ActiveSheet
    lastColumn = ws.Cells.Find("*", [A1], _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious _
    ).column
End Function

Function lastRow(Optional ws As Worksheet) As Integer
    If ws Is Nothing Then Set ws = ActiveSheet
    '----------
    ' Method 1
    '----------
    lastRow = ws.Cells.Find("*", [A1], _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious _
    ).row
    Exit Function
    
    '----------
    ' Method 2
    '----------
    ' wrong if first row is empty
    lastRow = ws.UsedRange.Rows.Count
    
    '----------
    ' Method 3
    '----------
    ' wrong if last row deleted
    lastRow = ws.Cells.SpecialCells(xlCellTypeLastCell).row
    
    '----------
    ' Method 3
    '----------
    ' find only first column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp)
End Function

Function specialFolder(ByVal name As String) As String
    specialFolder = CreateObject("WScript.Shell").SpecialFolders(name)
End Function
