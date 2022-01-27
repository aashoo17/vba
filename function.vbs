Option Explicit
' Procedure/Function syntax
Function FunctionProcedure()
    MsgBox "inside a function"
End Function

' Procedure with arguments
Function FunctionWithArg(x As Integer, Optional y As Integer = 10)
    Dim z As Integer
    z = 10
    MsgBox x
    MsgBox y
End Function

' all args are passed by reference by default even primitives like Integer
' to pass by value ByVal keyword needs to be used
Function PassByValueFunction (ByVal MyVar As Integer) ' Function declaration. 

End Function 

' Procedure with return
' returning primitive types
Function procedureWithReturn(x As Integer) As Integer
    ' use the name of the function and assign it to the return value and then exit function
    procedureWithReturn = 10
    Exit Function
End Function
' returning object types
Function procedureReturningObjects() As Workbook
    ' use the name of the function and Set it to the object value and exit function
    Set procedureReturningObjects = Workbooks(1)
    Exit Function
End Function

' [calling sub and procedures](https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/calling-sub-and-function-procedures)
Sub SubProcedures()
    Call FunctionProcedure  ' using call
    FunctionProcedure       ' direct calling
    FunctionWithArg 20      ' arg passing and leaving optional args
    FunctionWithArg 20, 30  ' passing optional arg also
    ' assigning a variable to the return value of function - arguments need to be passed in parantheses
    Dim x As Integer
    x = procedureWithReturn(10)
    ' x = procedureWithReturn 10 => this does not work
    ' TODO: Pass named arguments
End Sub
