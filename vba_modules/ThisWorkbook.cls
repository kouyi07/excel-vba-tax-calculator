Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_Open()
Application.ScreenUpdating = False
    Dim ws As Worksheet
    
    Worksheets("Tax Report Calculator").Activate
    Range("E10").Value = "NA"
    Range("J14:J16").Value = 0
    Range("J18").Value = 0
    Range("J21").Value = 0
    Range("J31").Value = 0
    Range("J32").Value = 0
    Range("G41").Value = 0
    Range("O41:O42").Value = 0
    Range("O47").Value = 0
    Range("M48:M50").Value = 0
    Range("O51:O55").Value = 0
    Range("G65").Value = 0
    Range("M60:M62").Value = 0
    Range("M64:M65").Value = 0
    Range("G65").Value = 0


    Range("E60").Value = "No"
    Range("E63").Value = "No"
 
    Range("E64").Value = "Yes"
    Range("E61").Value = "Single"
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",False)"
    Application.DisplayFormulaBar = False
    ActiveWindow.DisplayWorkbookTabs = False
    For Each ws In ThisWorkbook.Worksheets
        If Not ws.name = "Blank" Then
            ws.Activate
            With ActiveWindow
                .DisplayHeadings = False
            End With
        End If
    Next ws
    
    
    Worksheets("Home").Activate
    
     
Application.ScreenUpdating = True
End Sub
