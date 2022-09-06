' [declaring variable](https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/declaring-variables)

Option Explicit     ' this statement does not allow implicit variable creation so Dim statement will be required in advance
Sub Variable()
  ' variables in vba
  ' integers
  Dim int_a As Integer
  Dim long_b As Long
  ' floats
  Dim float_c As Single
  Dim float_d As Double
  ' string
  Dim string_e As String
  ' boolean
  Dim bool As Boolean
  ' date
  Dim date_g As Date

  int_a = 10
  long_b = 10
  float_c = 10.12
  float_d = 11.123
  string_e = "Hello World"
  bool = False
  date_g = #1/1/2020#

  ' object
  Dim obj as Object
  Set obj = New Workbooks

  ' Variant data type can contain string, date, time, Boolean, or numeric values, and can convert the values that they contain automatically
  ' [why to use variant](https://software-solutions-online.com/vba-variant/)
  ' TODO: once created and assigned to a type can variant be assigned to other type later
  Dim variant_a as Variant  
  Dim variant_b     'with no type it is created as variant
  theVar = "This is variant." 

  ' only c created as integer and a & b as variant - oops
  Dim a,b,c As Integer
  ' The shorthand for the types is: % -integer; & -long; @ -currency; # -double; ! -single; $ -string
  Dim d%,e%,f as Integer  ' using short form of declarion to avoid previous problem

End Sub