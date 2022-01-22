Option Explicit
Type Human
    name As String
    age As Integer
End Type

Sub CustomTypes()
    Dim h As Human
    h.age = 20
    h.name = "Human"

    ' using With in custom types
    With h
        .age = 30
        .name = "Some new Name"
    End With
    MsgBox h.name

End Sub