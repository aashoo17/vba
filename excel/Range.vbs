' in general users perspective this is single cell or group of cell called as Range object
' selection 
Sub RangeSelection()
    Application.DisplayAlerts = False
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim rg As Range
    Set wb = Workbooks("prac")
    Set sh = wb.Worksheets("Sheet1")
    ' a1:c10 => this will select cell a1 to c10
    ' a:c => this will select row a to c
    ' 1:10 => select column 1 to 10
    ' sh.Cells => select entire worksheet
    ' a1:c10,f3:h9,a15 => this selects multiple cells (cell range)
    Set rg = sh.Range("a1:c10")
    rg.Select
    wb.Save
End Sub

' applying visual indicators like color, background, border, font (changing and modifying size) etc 

' Border
Sub BorderCells()
    Application.DisplayAlerts = False
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim rg As Range
    Set wb = Workbooks("prac")
    Set sh = wb.Worksheets("Sheet1")
    Set rg = sh.Range("c3:f13")
    ' TODO: putting border around all sides of a cells
    rg.BorderAround Weight:=xlThick, ColorIndex:=3
    wb.Save
End Sub

Sub BorderCells()
    Application.DisplayAlerts = False
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim rg As Range
    Set wb = Workbooks("prac")
    Set sh = wb.Worksheets("Sheet1")
    ' TODO: this accesses only one side of a cell/range how can we do that for all sides in one go 
    With sh.Range("b2:c8").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = 3
    End With
    wb.Save
End Sub

' BackgroundColor
Sub BackgroundColorCells()
    Application.DisplayAlerts = False
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim rg As Range
    Set wb = Workbooks("prac")
    Set sh = wb.Worksheets("Sheet1")
    Set rg = sh.Range("c3:f13")
    ' ColorIndex gives access to 56 colors (from excel decided colot pallette) 
    ' which we can access from 1 to 56 number
    rg.Interior.ColorIndex = 20
    wb.Save
End Sub

' Font
Sub FontCells()
    Application.DisplayAlerts = False
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim rg As Range
    Set wb = Workbooks("prac")
    Set sh = wb.Worksheets("Sheet1")
    With sh.Range("b2:c8").Font
        .name = "Arial"
        .Size = 12
        .ColorIndex = 10
        .FontStyle = "Bold Italic"
    End With
    wb.Save
End Sub



