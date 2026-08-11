Attribute VB_Name = "Sheet7"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "TaxBracket, 6, 0, MSForms, ToggleButton"
Option Explicit
Sub SaveReportAsPdf()

Dim saveLocation As String
Dim ws As Worksheet
Dim rng As Range
Dim name As String
Application.ScreenUpdating = False

saveLocation = ActiveWorkbook.Path & "\"
Debug.Print saveLocation
Set ws = Sheets("Tax Report Calculator")
Set rng = ws.Range("B7:P95")

rng.ExportAsFixedFormat Type:=xlTypePDF, _
Filename:=saveLocation & Range("E10").Value & " Tax Calculator YA2023", _
OpenAfterPublish:=True

ws.DisplayPageBreaks = False

Application.ScreenUpdating = True

End Sub



Private Sub TaxBracket_Click()
    If TaxBracket.Value Then
        Rows("99:110").Hidden = False
    Else
        Rows("99:110").Hidden = True


    End If
End Sub


Private Sub Worksheet_Activate()
    'Limit scroll area
    
    Sheets("Tax Report Calculator").ScrollArea = "A1:R112"
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("E61")) Is Nothing Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        
        Dim married As String
        married = Range("E61").Value
        
         If married = "Single" Then
                Range("E63") = "No"
                Range("E64") = "Yes"
                Range("G65") = 0
                Range("C63:G65").NumberFormat = ";;;"
                
        ElseIf married = "Married" Then
                Range("E63") = "No"
                Range("E64") = "Yes"
                Range("G65") = 0
                Range("C63:G65").NumberFormat = "General"
                
         ElseIf married = "Divorced" Then
                Range("E63") = "No"
                Range("E64") = "Yes"
                Range("G65") = 0
                Range("C63:G64").NumberFormat = ";;;"
                Range("C65:G65").NumberFormat = "General"
         End If
         
         Application.ScreenUpdating = True
        Application.EnableEvents = True
         
    End If
    
End Sub
