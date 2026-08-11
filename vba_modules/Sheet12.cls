Attribute VB_Name = "Sheet12"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "ShowRibbon, 6, 1, MSForms, ToggleButton"
Attribute VB_Control = "ShowRowCol, 5, 2, MSForms, ToggleButton"
Attribute VB_Control = "ShowFormula, 4, 3, MSForms, ToggleButton"
Attribute VB_Control = "BottomTabs, 2, 4, MSForms, ToggleButton"
Option Explicit



Private Sub BottomTabs_Click()
    If BottomTabs.Value Then
        ActiveWindow.DisplayWorkbookTabs = True
    Else
        ActiveWindow.DisplayWorkbookTabs = False
    End If
    

End Sub

Private Sub ShowFormula_Click()
If ShowFormula.Value Then

    Application.DisplayFormulaBar = True
    Else
    Application.DisplayFormulaBar = False
    End If
    
End Sub

Private Sub ShowRibbon_Click()
If ShowRibbon.Value Then
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",True)"
Else
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",False)"
End If

End Sub

Private Sub ShowRowCol_Click()
Dim ws As Worksheet
    If ShowRowCol.Value Then
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        If Not ws.name = "Blank" Then
            ws.Activate
            With ActiveWindow
                .DisplayHeadings = True
            End With
        End If
    Next ws
    Application.ScreenUpdating = True
    
    Else
    
    Application.ScreenUpdating = False
    For Each ws In ThisWorkbook.Worksheets
        If Not ws.name = "Blank" Then
            ws.Activate
            With ActiveWindow
                .DisplayHeadings = False
            End With
        End If
    Next ws
    Application.ScreenUpdating = True
    End If
    
    Worksheets("Home").Activate
    
    
End Sub



Private Sub Worksheet_Activate()
    'Limit scroll area
    
    Sheets("Home").ScrollArea = "A1:V30"
End Sub

 
Sub ProtectUnprotect()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    
    If ActiveSheet.ProtectContents = False Then
    
        For Each ws In Worksheets
            If ws.ProtectContents = False Then
                If ws.name = "Tax Report Calculator" Then
                    ws.Protect "1", userinterfaceonly:=True, AllowFormattingCells:=True, AllowFormattingRows:=True
                    
                    ws.EnableSelection = xlUnlockedCells
                ElseIf ws.name = "Pivot Tables" Then
                    ws.Protect "1", AllowUsingPivotTables:=True
                
                    
                    
                ElseIf ws.name = "Categories (All)" Then
                ElseIf ws.name = "Categories (expenses)" Then
                ElseIf ws.name = "Categories (Toggle)" Then
                ElseIf ws.name = "Expenses (2023)" Then
                
                Else
                      ws.Protect "1"
                      
                      
                End If
            End If
        Next
    Else
        Dim pw As String
        pw = InputBox("Please Enter the password")
        For Each ws In Worksheets
            On Error GoTo ErrMessage
            
            
            ws.Unprotect pw
            
            
        Next
      
        
     
        
              
    End If
        
   
    
    Worksheets("Home").Activate
    Application.ScreenUpdating = True
Done:
    Exit Sub
    
ErrMessage:
    MsgBox "Wrong Password!"
End Sub

