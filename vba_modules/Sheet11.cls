Attribute VB_Name = "Sheet11"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "ComboBox1, 1, 0, MSForms, ComboBox"
Option Explicit


Sub DebugPrintShapeNames()
    Dim ws As Worksheet
    Dim shp As Shape
    
    ' Set worksheet variable to the active sheet
    Set ws = ActiveSheet ' Change to the sheet where your shapes are
    
    ' Loop through all shapes in the worksheet
    For Each shp In ws.Shapes
        Debug.Print shp.name
    Next shp
End Sub


Private Sub Worksheet_Activate()
     Sheets("Dashboard (Expenses)").ScrollArea = "A1:M65"
End Sub


