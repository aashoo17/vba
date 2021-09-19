Option Explicit
Type Human
    name As String
    age As Integer
End Type

Sub CustomTypes()
    Dim h As Human
    h.age = 20
    h.name = "Human"
    MsgBox h.age
End Sub