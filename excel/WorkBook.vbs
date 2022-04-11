' Accesing already opened Workbook

Sub AccessAlreadyOpenedWorkbook()
    ' declare a Workbook variable
    Dim wb As Workbook
    ' set it to already opened Workbook 
    ' name of the excel file can be used as the key to access the Workbook
    ' though we can use index like 1,2,3... etc to access Workbook but key which 
    ' is Workbook name access them uniquely that is really nice
    ' this is also reason than same name excel file even if they are at different location
    ' can't be opened together (that is bad design though I must say)  
    Set wb = Workbooks("prac")  ' access prac.xlsx workbook
End Sub

' Saving a workbook

' creating a new workbook programmatically and saving it to some location
Sub CreateWorkbookAndSave()
    Dim wb As Workbook
    Dim sh As Worksheet
    
    Set wb = Workbooks.Add
    wb.SaveAs "sample"  ' by default files gets saved to Documents folder
End Sub

' creating a new workbook programmatically from existing excel file and saving it to some location
Sub CreateWorkbookFromExistingExcelAndSave()
    Dim wb As Workbook
    Dim sh As Worksheet
    ' use an existing excel file to create new file
    Set wb = Workbooks.Add("prac")
    wb.SaveAs "sample"  ' by default files gets saved to Documents folder
End Sub
