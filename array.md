Declare array
```vb
Dim a() as string
```
 
Init array
```vb
a = Array("A", "B", "C")
```

Resize array (preserve values)
```vb
ReDim Preserve a(20)
' Array("A", "B", "C", Nothing, Nothing, ...)
```

Init array size to 10 (remove values)
```vb
ReDim a(10)
' Array(Nothing, Nothing, ...)
```
