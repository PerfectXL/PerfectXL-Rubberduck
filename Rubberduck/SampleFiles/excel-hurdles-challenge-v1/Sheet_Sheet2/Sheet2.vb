Attribute VB_Name = "Sheet2"
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
    On Error Resume Next
    If Application.Intersect(Range("cellsInSecond"), target) Is Nothing Then
        'user selected a cell that is not on trial
        addPenalty
    ElseIf ActiveCell.Value = "" And ActiveCell.address <> "$C$33" Then
        'user selected a blank cell
        addPenalty
    End If
    On Error GoTo Last
    If target.Name.Name = "hetch20" And [isGameOn] Then
            [endTime] = Now
            DoEvents
            [isGameOn] = False
            [valDuration] = [endTime] - [startTime] + ([penaltyTotal] / 24 / 60 / 60)
            MsgBox "You have reached the end in " & [timeElapsedF] & " seconds", vbOKOnly, "Congratulations"
            Sheets("Results").Visible = True
            Sheets("Results").Select
    End If
Last:
    checkCell target
    showTime
End Sub

