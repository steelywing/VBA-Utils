For each file in folder
```vb
Dim file As String
Dim folder As String
Dim wb As Workbook

folder = "Folder\"
file = Dir(folder & "*.xlsx")
While file <> ""
    Set wb = Workbooks.Open(folder & file)
    ' do something for the file
    wb.Close
    file = Dir()
Wend
```

Open Excel file in Word, remember enable [Tools] -> [References] -> [Microsoft Excel XX.X Object Library]
```vb
Dim xlApp As Excel.Application
Dim wb As Excel.Workbook
Dim ws As Excel.Worksheet
Dim name As ContentControl

Set xlApp = New Excel.Application

Set wb = xlApp.Workbooks.Open(ActiveDocument.Path & "\Book.xlsx")
Set ws = wb.Worksheets(1)
' xlApp.Visible = True
Set name = ActiveDocument.SelectContentControlsByTag("name")(1)
name.Range.Text = _
    ws.Range("A1").Value
wb.Close
```
