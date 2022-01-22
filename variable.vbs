' [declaring variable](https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/declaring-variables)

Option Explicit     ' this statement does not allow implicit variable creation so Dim statement will be required in advance
Sub Variable()
' variables in vba
' integers
Dim a As Integer
Dim b As Long
' floats
Dim c As Single
Dim d As Double
' string
Dim e As String
' boolean
Dim f As Boolean
' date
Dim g As Date

a = 10
b = 10
c = 10.12
d = 11.123
e = "Hello World"
f = False
g = #1/1/2020#

MsgBox g

End Sub