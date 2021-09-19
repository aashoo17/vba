Option Explicit

Sub Arrays()
    Dim h(20) As Integer
    Dim i As Integer
    For i = 1 To 20
        h(i) = i
    Next
    MsgBox h(10)
End Sub