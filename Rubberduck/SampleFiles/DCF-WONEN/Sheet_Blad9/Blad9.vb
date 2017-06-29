Attribute VB_Name = "Blad9"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "Parkeerplaatsen, 1, 0, MSForms, CheckBox"
Attribute VB_Control = "MGW, 2, 1, MSForms, CheckBox"
Private Sub CheckBox1_Click()

Application.ScreenUpdating = False

If CheckBox1.Value = True And Parkeerplaatsen.Value = True Then
    kolombreedte1 = 5 'kolom AB
    kolombreedte2 = 51 'kolom AJ
    Columns("AB:AB").ColumnWidth = kolombreedte1
    Columns("AJ:AJ").ColumnWidth = kolombreedte2
End If
If CheckBox1.Value = True And Parkeerplaatsen.Value = False Then
    kolombreedte1 = 5 + 33
    kolombreedte2 = 51 + 48
    Columns("AB:AB").ColumnWidth = kolombreedte1
    Columns("AJ:AJ").ColumnWidth = kolombreedte2
End If
If CheckBox1.Value = False And Parkeerplaatsen.Value = True Then
    kolombreedte1 = 5 + 76
    kolombreedte2 = 51 + 64
    Columns("AB:AB").ColumnWidth = kolombreedte1
    Columns("AJ:AJ").ColumnWidth = kolombreedte2
End If
If CheckBox1.Value = False And Parkeerplaatsen.Value = False Then
    kolombreedte1 = 5 + 33 + 76
    kolombreedte2 = 51 + 48 + 64
    Columns("AB:AB").ColumnWidth = kolombreedte1
    Columns("AJ:AJ").ColumnWidth = kolombreedte2
End If

If CheckBox1.Value = True Then
    Columns("T:T").Hidden = False
    afdrukbereik = "zondertax"
Else
    Columns("T:T").Hidden = True
    afdrukbereik = "mettax"
End If
If CheckBox1.Value = True Then
    Columns("U:U").Hidden = False
Else
    Columns("U:U").Hidden = True
End If
If CheckBox1.Value = True Then
    Columns("V:V").Hidden = False
Else
    Columns("V:V").Hidden = True
End If
If CheckBox1.Value = True Then
    Columns("W:W").Hidden = False
Else
    Columns("W:W").Hidden = True
End If
If CheckBox1.Value = True Then
    Columns("X:X").Hidden = False
Else
    Columns("X:X").Hidden = True
End If
If CheckBox1.Value = True Then
    Columns("Y:Y").Hidden = False
Else
    Columns("Y:Y").Hidden = True
End If
If CheckBox1.Value = True Then
    Columns("Z:Z").Hidden = False
Else
    Columns("Z:Z").Hidden = True
End If
If CheckBox1.Value = True Then
    Columns("AA:AA").Hidden = False
Else
    Columns("AA:AA").Hidden = True
End If

If CheckBox1.Value = True Then
    Columns("AF:AF").Hidden = False
Else
    Columns("AF:AF").Hidden = True
End If
If CheckBox1.Value = True Then
    Columns("AH:AH").Hidden = False
Else
    Columns("AH:AH").Hidden = True
End If

If CheckBox1.Value = True And Parkeerplaatsen.Value = True Then
    Columns("AG:AG").Hidden = False
Else
    Columns("AG:AG").Hidden = True
End If
If CheckBox1.Value = True And Parkeerplaatsen.Value = True Then
    Columns("AI:AI").Hidden = False
Else
    Columns("AI:AI").Hidden = True
End If
       
    Application.ScreenUpdating = True

End Sub

Private Sub MGW_Click()

Application.ScreenUpdating = False

If MGW.Value = True Then
    Cells(14, 4) = "Verdieping"
    Cells(15, 4) = ""
    Cells(14, 7) = "Ori?ntatie"
    Cells(15, 7) = "welke zijde van het gebouw"
    Cells(15, 7).WrapText = True
    Cells(15, 7).ShrinkToFit = True
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ReadingOrder = xlContext
'        .MergeCells = False
    Cells(14, 8) = "Balkon"
    Cells(14, 9) = "Balkon-"
Else
    Cells(14, 4) = "Beukmaat"
    Cells(15, 4) = "in m1"
    Cells(14, 7) = "Berging (los)"
    Cells(15, 7) = "m2"
    Cells(15, 7).WrapText = False
    Cells(15, 7).ShrinkToFit = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ReadingOrder = xlContext
'        .MergeCells = False
    Cells(14, 8) = "Kavelopp"
    Cells(14, 9) = "Tuin-"

End If
    
Application.ScreenUpdating = True
    
End Sub

Private Sub Parkeerplaatsen_Click()

Application.ScreenUpdating = False

If Parkeerplaatsen.Value = True Then
    kolombreedte1 = 5 'kolom AH
    Columns("AH:AH").ColumnWidth = kolombreedte1
Else
    kolombreedte1 = 5 + 99
    Columns("AH:AH").ColumnWidth = kolombreedte1
End If
    
If Parkeerplaatsen.Value = True Then
    Columns("Q:Q").Hidden = False
Else
    Columns("Q:Q").Hidden = True
End If
If Parkeerplaatsen.Value = True Then
    Columns("R:R").Hidden = False
Else
    Columns("R:R").Hidden = True
End If
If Parkeerplaatsen.Value = True Then
    Columns("S:S").Hidden = False
Else
    Columns("S:S").Hidden = True
End If
If Parkeerplaatsen.Value = True Then
    Columns("X:X").Hidden = False
Else
    Columns("X:X").Hidden = True
End If
If Parkeerplaatsen.Value = True Then
    Columns("Y:Y").Hidden = False
Else
    Columns("Y:Y").Hidden = True
End If
If Parkeerplaatsen.Value = True Then
    Columns("Z:Z").Hidden = False
Else
    Columns("Z:Z").Hidden = True
End If
If Parkeerplaatsen.Value = True Then
    Columns("AE:AE").Hidden = False
Else
    Columns("AE:AE").Hidden = True
End If
If Parkeerplaatsen.Value = True Then
    Columns("AF:AF").Hidden = False
Else
    Columns("AF:AF").Hidden = True
End If
If Parkeerplaatsen.Value = True Then
    Columns("AG:AG").Hidden = False
Else
    Columns("AG:AG").Hidden = True
End If

If Parkeerplaatsen.Value = True Then
    kolombreedte1 = 117 'kolom AQ
    Columns("AQ:AQ").ColumnWidth = kolombreedte1
Else
    kolombreedte1 = 117 + 48
    Columns("AQ:AQ").ColumnWidth = kolombreedte1
End If

If Parkeerplaatsen.Value = True Then
    Columns("AL:AL").Hidden = False
Else
    Columns("AL:AL").Hidden = True
End If

If Parkeerplaatsen.Value = True Then
    Columns("AN:AN").Hidden = False
Else
    Columns("AN:AN").Hidden = True
End If
If Parkeerplaatsen.Value = True Then
    Columns("AP:AP").Hidden = False
Else
    Columns("AP:AP").Hidden = True
End If

    Application.ScreenUpdating = True

End Sub
