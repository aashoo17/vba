# Excel

## Working with macro
1. open a workbook
2. alt + F11 an macro editor will open 
3. write click on VBAProject and insert a module where we can write our vba macro
4. macro code has to be written inside a sub procedure  
```vb
Sub MyMacro
    ' code goes here
End Sub
```
5. excel macro needs to be saved and this can only be done if workbook is saved as macro enabled workbook
having file extension .xlsm  
so go to save as and save as macro enabled workbook.  

## excel object hierarchy

Application => Workbook => Worksheet => Range and other objects all will generally lie in Worksheet  

when we have opened a Excel file following objects are automatically created for you  
Excel Application, 1 Workbook, 2 Worksheets  

and these objects are bound to following variables for user to access them  

Application => Excel Application  
Workbooks => All opened workbooks collection  
Worksheets => All worksheet collection of currently active workbook  


