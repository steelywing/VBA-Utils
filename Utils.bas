Attribute VB_Name = "Utils"
'union Range object
Function unionRange(ParamArray ranges() As Variant) As Range
    Dim result As Range
    
    For i = LBound(ranges) To UBound(ranges)
        If IsObject(ranges(i)) Then
            If Not ranges(i) Is Nothing Then
                If TypeOf ranges(i) Is Range Then
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

Function isWorksheetExist(ByVal name, Optional wb As Workbook) As Boolean
    Dim ws As Worksheet
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

Function SpecialFolders(ByVal name As String)
    SpecialFolders = CreateObject("WScript.Shell").SpecialFolders(name)
End Function

Sub test()
MsgBox SpecialFolders(SF_MY_DOCUMENTS)
End Sub
