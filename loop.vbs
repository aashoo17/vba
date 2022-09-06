Option Explicit
Sub Loops()
    Dim i As Integer
    For i = 1 To 10
        MsgBox i
    Next

    Dim j As Integer
    ' step will define how many step to jump
    For j = 1 To 10 Step 2
        MsgBox j
    Next j      ' todo: what does j after Next means

    ' TODO: loop over arrays

    Dim k(4) As Integer
    ' Dim i As Variant
    For Each m In h
        MsgBox m
    Next

End Sub