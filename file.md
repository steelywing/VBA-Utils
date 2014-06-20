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
