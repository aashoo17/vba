Option Explicit
Function FunctionProcedure()
    MsgBox "inside a function"
    ' TODO: return value in procedure
End Function
Function FunctionWithArg(x As Integer, Optional y As Integer = 10)
    MsgBox x
    MsgBox y
End Function

Sub SubProcedures()
    ' do something here
    Call FunctionProcedure  'using call
    FunctionProcedure       ' direct calling
    FunctionWithArg 20      'arg passing and leaving optional args
    FunctionWithArg 20, 30  ' passing optional arg also
End Sub
