Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
'Macro's die worden uitgevoerd bij het opslaan van het document

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

If ActiveWindow.SelectedSheets.Count > 1 Then Exit Sub

Sheets("NRVT rapport (NL)").Protect "MVGM"
Sheets("NRVT rapport (ENG)").Protect "MVGM"
Sheets("VOORBLAD").Protect "MVGM"
Sheets("Samenvatting").Protect "MVGM"
Sheets("Summary Amvest").Protect "MVGM"
Sheets("Voorblad CBRE").Protect "MVGM"
Sheets("Summary").Protect "MVGM"
Sheets(" TYPERINGEN invoer").Protect "MVGM"
Sheets("UITPOND SCENARIO").Protect "MVGM"
Sheets("DOOREXPLOITATIE SCENARIO").Protect "MVGM"
Sheets("Engels (print)").Protect "MVGM"
Sheets("KENGETALLEN").Protect "MVGM"
Sheets("FACTSHEET Achmea").Protect "MVGM"
Sheets("UITPOND SCENARIO (20)").Protect "MVGM"
Sheets("DOOREXPLOITATIE SCENARIO (20)").Protect "MVGM"

End Sub

'Macro's die worden uitgevoerd bij het opstarten van het document

Private Sub Workbook_Open()
Application.Calculation = xlAutomatic

End Sub
