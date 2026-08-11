Attribute VB_Name = "Sheet10"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit


Sub GetSlicerNames()
    Dim sc As slicerCache
    Dim si As slicerItem
    
    For Each sc In ThisWorkbook.SlicerCaches
        Debug.Print sc.name
    Next
End Sub
Private Sub Worksheet_PivotTableUpdate(ByVal Target As PivotTable)
    Dim slicerCache As slicerCache
    Dim slicerItem As slicerItem
    Dim boolFoundFirstSelect As Boolean

    ' Loop through all slicers in the worksheet
    For Each slicerCache In ThisWorkbook.SlicerCaches
        boolFoundFirstSelect = False

        Dim intCountNOTSelected As Integer
        intCountNOTSelected = 0

        For Each slicerItem In slicerCache.SlicerItems
            If slicerItem.Selected = False Then
                intCountNOTSelected = intCountNOTSelected + 1
            End If
        Next slicerItem

        If intCountNOTSelected > 0 Then
            For Each slicerItem In slicerCache.SlicerItems
                If slicerItem.Selected = True Then
                    If boolFoundFirstSelect = False Then
                        boolFoundFirstSelect = True
                    Else
                        slicerItem.Selected = False
                        MsgBox "Please select only one item"
                    End If
                End If
            Next slicerItem
        Else
            ' If all of them are selected, then we'll do nothing per the OP's request
        End If
    Next slicerCache
End Sub

