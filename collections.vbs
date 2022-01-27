' A Collection object is an ordered set of items that can be referred to as a unit.

Sub Proc()
    ' TODO: so once collection is created I can insert any type of object inside
    ' for e.g. collection of Workbook, Worksheet etc.
    Dim X As New Collection
    ' TODO: Add, Remove, Count api
    X.Add "X"
    X.Add "Y"
    X.Add "Z", "MyKey"  ' this collection value can be accessed using this key
    ' so collection can work like vector or map both interesting
    ' TODO: can we insert two different types of objects say string and integer
    Dim z, y As String
    y = X(1)
    z = X("MyKey")
    MsgBox z
    MsgBox X.Count

    X.Remove (2)
    X.Remove ("MyKey")
    Set X = Nothing
End Sub