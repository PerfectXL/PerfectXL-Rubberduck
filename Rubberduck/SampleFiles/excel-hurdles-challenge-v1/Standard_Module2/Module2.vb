Attribute VB_Name = "Module2"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveSheet.Shapes.Range(Array("TextBox 5")).Select
    Selection.Formula = "=calc!G6"
    Range("E5").Select
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Sheets("Results").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("calc").Select
    Sheets("Results").Visible = True
End Sub
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    Sheets("calc").Select
    Range("G10").Select
    Selection.Copy
    Sheets("Results").Select
    Application.CutCopyMode = False
    Sheets("calc").Select
    Selection.Copy
    Application.CutCopyMode = False
    Sheets("Results").Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection.ShapeRange.Item(1), address:= _
        "https://twitter.com/intent/tweet?text=I+finished+Excel+Hurdle+challenge+in+19+seconds.+http%3A%2F%2Fchandoo.org%2Fwp%2F2012%2F08%2F10%2Fexcel-hurdles%2F+via+@r1c1"
End Sub
