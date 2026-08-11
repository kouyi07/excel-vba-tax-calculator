Attribute VB_Name = "Button_Expenses"
Option Explicit


Sub Top5_Click()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim myShape1 As Shape
    Dim myShape2 As Shape

    ' Set worksheet variable by name
    Set ws = ThisWorkbook.Sheets("Pivot Tables")

    ' Set pivot table variable by name
    Set pt = ws.PivotTables("Recipients")
    
    pt.PivotFields("Recipient").ClearAllFilters
    pt.PivotFields("Recipient").PivotFilters. _
    Add2 Type:=xlTopCount, DataField:=pt. _
    PivotFields("Sum of Total"), Value1:=5
    
    Set myShape1 = ThisWorkbook.Sheets("Dashboard (Expenses)").Shapes("BottomFiveButton")
    myShape1.Fill.ForeColor.RGB = RGB(48, 84, 150)
    
    Set myShape2 = ThisWorkbook.Sheets("Dashboard (Expenses)").Shapes("TopFiveButton")
    myShape2.Fill.ForeColor.RGB = RGB(179, 217, 255)
End Sub

Sub Bottom5_Click()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pf As PivotField
    Dim myShape1 As Shape
    Dim myShape2 As Shape

    ' Set worksheet variable by name
    Set ws = ThisWorkbook.Sheets("Pivot Tables")

    ' Set pivot table variable by name
    Set pt = ws.PivotTables("Recipients")
    
    pt.PivotFields("Recipient").ClearAllFilters
    pt.PivotFields("Recipient").PivotFilters. _
    Add2 Type:=xlBottomCount, DataField:=pt. _
    PivotFields("Sum of Total"), Value1:=5
    
    Set myShape1 = ThisWorkbook.Sheets("Dashboard (Expenses)").Shapes("TopFiveButton")
    myShape1.Fill.ForeColor.RGB = RGB(48, 84, 150)
    
    Set myShape2 = ThisWorkbook.Sheets("Dashboard (Expenses)").Shapes("BottomFiveButton")
    myShape2.Fill.ForeColor.RGB = RGB(179, 217, 255)
End Sub

Sub ShowDetails_Click()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Pivot Tables")
    ws.Visible = True
    ws.Activate
End Sub

Sub HideDetails_Click()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Set ws1 = ThisWorkbook.Sheets("Pivot Tables")
    Set ws2 = ThisWorkbook.Sheets("Dashboard (Expenses)")
    ws2.Activate
End Sub
