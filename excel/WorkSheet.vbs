' access worksheet, modify it and save 
Sub ModifyingWorksheetAndSave()
    Dim wb As Workbook
    Dim sh As Worksheet
    Set wb = Workbooks("prac")
    Set sh = wb.Worksheets("Sheet1")
    Dim rg As Range
    Set rg = sh.Range("a1:d10")
    rg.Value = 200
    wb.Save
End Sub

' Add a new worksheet programatically
Sub AddSheets()
    Application.DisplayAlerts = False
    Dim wb As Workbook
    Dim sh As Worksheet
    Set wb = Workbooks("prac")
    wb.Worksheets.Add , , 3
    wb.Save
End Sub

' delete a worksheet programmatically
Sub DeleteSheets()
    Application.DisplayAlerts = False
    Dim wb As Workbook
    Dim sh As Worksheet
    Set wb = Workbooks("prac")
    For Each sh In wb.Worksheets
        ' TODO: this does not delete the Sheet1 though
        sh.Delete
    Next
    wb.Save
End Sub



