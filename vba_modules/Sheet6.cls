Attribute VB_Name = "Sheet6"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "TaxBracket, 11, 0, MSForms, ToggleButton"
Option Explicit
Sub SaveRangeAsPDF()

Dim saveLocation As String
Dim ws As Worksheet
Dim rng As Range
Dim name As String


saveLocation = ActiveWorkbook.Path & "\"
Debug.Print saveLocation
Set ws = Sheets("Input Form")
Set rng = ws.Range("B3:L45")

rng.ExportAsFixedFormat Type:=xlTypePDF, _
Filename:=saveLocation & Range("D5").Value & " Tax Summary YA2023", _
OpenAfterPublish:=True




End Sub



Private Sub TaxBracket_Click()
    If TaxBracket.Value Then
        Columns("P:S").Hidden = False
    Else
        Columns("P:S").Hidden = True


    End If
End Sub


Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("D13", "D16")) Is Nothing Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        
         Dim married As String
            married = Range("D13").Value
            
            If married = "Single" Then
                Range("D15") = "No"
                Range("D16") = "Yes"
                Range("F17") = 0
                Range("B15:F17").NumberFormat = ";;;"
                Range("D15:F17").Interior.Color = xlNone
                Range("D15:F17").Borders.LineStyle = xlNone
                Range("D15:D16", "F17").Validation.Delete
                
                
            ElseIf married = "Married" Then
                'have to make sure no other validation before adding
                Range("D15:D16", "F17").Validation.Delete
                disabled_working_validation
                alimony_validation
                Range("B15:F17").NumberFormat = "General"
                Range("D15:D16").Interior.Color = RGB(255, 204, 153)
                Range("F15:F16").Interior.Color = RGB(226, 239, 218)
                Range("D15").BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(127, 127, 127)
                Range("D16").BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(127, 127, 127)
                    Dim WorkingOther As String
                    WorkingOther = Range("D16").Value
                    If WorkingOther = "Yes" Then
                    
                        Range("F17").Validation.Delete
                        alimony_validation
                        Range("B17:F17").NumberFormat = "General"
                        Range("F17").Interior.Color = RGB(255, 204, 153)
                        Range("F17").BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(127, 127, 127)
                        
                    Else
                        Range("F17").Validation.Delete
                        Range("B17:F17").NumberFormat = ";;;"
                        Range("F17").Interior.Color = xlNone
                        Range("F17").Borders.LineStyle = xlNone
                    End If
                    
            ElseIf married = "Divorced" Then
                Range("F17").Validation.Delete
                alimony_validation
                Range("D15:D16").Validation.Delete
                Range("D16") = "Yes"
                Range("B15:F16").NumberFormat = ";;;"
                Range("D15:F16").Interior.Color = xlNone
                Range("D15:F16").Borders.LineStyle = xlNone
                Range("F17").Interior.Color = RGB(255, 204, 153)
                Range("F17") = 0
                Range("B17:F17").NumberFormat = "General"
            End If
            

            Application.EnableEvents = True
        Application.ScreenUpdating = True
        

  End If
  
  End Sub
  
Sub disabled_working_validation()
    With Range("D15:D16").Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Yes,No"
        
    End With
        
        
End Sub



Sub alimony_validation()
    With Range("F17").Validation
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="0", Formula2:="4000"
        .InputTitle = "Input a whole number"
        .InputMessage = "Not greater than 4000"
        .ErrorTitle = "Error!"
        .ErrorMessage = "Value should be not greater than 4000!"
        
    End With
    
    
End Sub
