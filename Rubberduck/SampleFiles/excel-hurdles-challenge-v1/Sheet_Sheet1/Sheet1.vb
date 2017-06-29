Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Worksheet_BeforeDoubleClick(ByVal target As Range, Cancel As Boolean)
    addPenalty
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal target As Range, Cancel As Boolean)
    addPenalty
End Sub

Private Sub Worksheet_SelectionChange(ByVal target As Range)
    If Application.Intersect(Range("cellsInFirst"), target) Is Nothing Then
        'user selected a cell that is not on trial
        addPenalty
    ElseIf ActiveCell.Value = "" Then
        'user selected a blank cell
        addPenalty
    End If
    On Error GoTo Last
    If ActiveCell.address = "$C$32" And [isGameOn] Then
        Sheets("Second").Visible = True
    End If
Last:
    checkCell target
    showTime
End Sub
