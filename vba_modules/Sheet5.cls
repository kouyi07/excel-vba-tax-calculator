Attribute VB_Name = "Sheet5"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Sub SaveSummaryAsPDF()

Dim saveLocation As String
Dim ws As Worksheet
Dim rng As Range
Dim name As String

Application.ScreenUpdating = False
saveLocation = ActiveWorkbook.Path & "\"
Debug.Print saveLocation
Set ws = Sheets("Tax Summary")

Set rng = ws.Range("D10:L27")

ws.Range("J26").Calculate
rng.ExportAsFixedFormat Type:=xlTypePDF, _
Filename:=saveLocation & Range("I12").Value & " Tax Summary YA2023", _
OpenAfterPublish:=True
ws.DisplayPageBreaks = False

Application.ScreenUpdating = True

End Sub

Private Sub Worksheet_Activate()
    Sheets("Tax Summary").ScrollArea = "A1:AB44"
End Sub


