Attribute VB_Name = "Navigation"
Option Explicit
Sub GoHome()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Home")
    ws.Activate
End Sub

Sub GoDataset()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Expenses (2023)")
    ws.Activate
End Sub

Sub GoDashboard()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard (Expenses)")
    ws.Activate
End Sub

Sub GoCategories()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Categories (All)")
    ws.Activate
End Sub

Sub GoTaxSummary()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tax Summary")
    ws.Activate
End Sub

Sub GoTaxReport()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Tax Report Calculator")
    ws.Activate
End Sub

