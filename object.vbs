Sub ashu()
    ' declare object variable
    Dim book As Workbook
    ' assign object variable to an instance
    Set book = Workbooks(1)
    
    Dim MyObject As Object ' Declared as generic object.
    Dim MyVariantObject ' since no type is given it will be variant by default
    ' TODO: Variant vs Object 
    ' Variables declared as the Variant data type can contain string, date, time, Boolean, or numeric values,
    ' and can convert the values that they contain automatically
    
    Set SomeObject = New Workbook ' Create and Assign in a single line using New
    Set SomeObject = Nothing ' Discontinue association of variable to actual memory location or instance of object
    ' New vs CreateObject
    ' When you use one application to control another application's objects, you should set a reference to the other
    ' application's type library
    ' for e.g. I am calling Word from Excel
    ' New can be used but some application may not support it when called from another program
    ' To determine which syntax an application supports, see the application's documentation.
    Dim appAccess As Object 
    Set appAccess = CreateObject("Word.Application")    ' creating word from excel vba
End Sub
