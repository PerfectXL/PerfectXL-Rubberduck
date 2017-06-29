Attribute VB_Name = "Module6"
Sub DTVernieuwenHoofdblad()
Attribute DTVernieuwenHoofdblad.VB_ProcData.VB_Invoke_Func = " \n14"
'
' DTVernieuwenHoofdblad Macro
'

'
    Application.ScreenUpdating = False
    Sheets("DB - Forecast").Select
    Range("B40").Select
    ActiveSheet.PivotTables("Forecast").PivotCache.Refresh
    Sheets("DB - FC vs Actuals").Select
    Range("B109").Select
    ActiveSheet.PivotTables("FCvsActuals").PivotCache.Refresh
    Sheets("DB - Delta per straat").Select
    Range("Y32").Select
    ActiveSheet.PivotTables("Delta").PivotCache.Refresh
    Sheets("DB - Billed sales").Select
    Range("B59").Select
    ActiveSheet.PivotTables("Billed sales").PivotCache.Refresh
    Sheets("DB - Historie").Select
    Range("B15").Select
    ActiveSheet.PivotTables("Historie").PivotCache.Refresh
    Sheets("Rapportage").Select
    Range("A8").Select
    ActiveSheet.PivotTables("Rapportage1").PivotCache.Refresh
    Range("A52").Select
    ActiveSheet.PivotTables("Rapportage2").PivotCache.Refresh
    Range("D52").Select
    ActiveSheet.PivotTables("Rapportage3").PivotCache.Refresh
    Sheets("Hoofdblad").Select
    Application.ScreenUpdating = True
    
    End Sub
    Sub DTVernieuwenDBForecast()
'
' DTVernieuwenDBForecast Macro
'

'
    Application.ScreenUpdating = False
    Sheets("DB - Forecast").Select
    Range("B40").Select
    ActiveSheet.PivotTables("Forecast").PivotCache.Refresh
    Sheets("DB - FC vs Actuals").Select
    Range("B109").Select
    ActiveSheet.PivotTables("FCvsActuals").PivotCache.Refresh
    Sheets("DB - Delta per straat").Select
    Range("Y32").Select
    ActiveSheet.PivotTables("Delta").PivotCache.Refresh
    Sheets("DB - Billed sales").Select
    Range("B59").Select
    ActiveSheet.PivotTables("Billed sales").PivotCache.Refresh
    Sheets("DB - Historie").Select
    Range("B15").Select
    ActiveSheet.PivotTables("Historie").PivotCache.Refresh
    Sheets("Rapportage").Select
    Range("A8").Select
    ActiveSheet.PivotTables("Rapportage1").PivotCache.Refresh
    Range("A52").Select
    ActiveSheet.PivotTables("Rapportage2").PivotCache.Refresh
    Range("D52").Select
    ActiveSheet.PivotTables("Rapportage3").PivotCache.Refresh
    Sheets("DB - Forecast").Select
    Application.ScreenUpdating = True
    
    End Sub
    
    Sub DTVernieuwenDBForecastvsActuals()
'
' DTVernieuwenDBForecastvsActuals Macro
'

'
    Application.ScreenUpdating = False
    Sheets("DB - Forecast").Select
    Range("B40").Select
    ActiveSheet.PivotTables("Forecast").PivotCache.Refresh
    Sheets("DB - FC vs Actuals").Select
    Range("B109").Select
    ActiveSheet.PivotTables("FCvsActuals").PivotCache.Refresh
    Sheets("DB - Delta per straat").Select
    Range("Y32").Select
    ActiveSheet.PivotTables("Delta").PivotCache.Refresh
    Sheets("DB - Billed sales").Select
    Range("B59").Select
    ActiveSheet.PivotTables("Billed sales").PivotCache.Refresh
    Sheets("DB - Historie").Select
    Range("B15").Select
    ActiveSheet.PivotTables("Historie").PivotCache.Refresh
    Sheets("Rapportage").Select
    Range("A8").Select
    ActiveSheet.PivotTables("Rapportage1").PivotCache.Refresh
    Range("A52").Select
    ActiveSheet.PivotTables("Rapportage2").PivotCache.Refresh
    Range("D52").Select
    ActiveSheet.PivotTables("Rapportage3").PivotCache.Refresh
    Sheets("DB - FC vs Actuals").Select
    Application.ScreenUpdating = True
    
    End Sub
    
    
    Sub DTVernieuwenDBDeltaPerStraat()
'
' DTVernieuwenDBDeltaPerStraat Macro
'

'
    Application.ScreenUpdating = False
    Sheets("DB - Forecast").Select
    Range("B40").Select
    ActiveSheet.PivotTables("Forecast").PivotCache.Refresh
    Sheets("DB - FC vs Actuals").Select
    Range("B109").Select
    ActiveSheet.PivotTables("FCvsActuals").PivotCache.Refresh
    Sheets("DB - Delta per straat").Select
    Range("Y32").Select
    ActiveSheet.PivotTables("Delta").PivotCache.Refresh
    Sheets("DB - Billed sales").Select
    Range("B59").Select
    ActiveSheet.PivotTables("Billed sales").PivotCache.Refresh
    Sheets("DB - Historie").Select
    Range("B15").Select
    ActiveSheet.PivotTables("Historie").PivotCache.Refresh
    Sheets("Rapportage").Select
    Range("A8").Select
    ActiveSheet.PivotTables("Rapportage1").PivotCache.Refresh
    Range("A52").Select
    ActiveSheet.PivotTables("Rapportage2").PivotCache.Refresh
    Range("D52").Select
    ActiveSheet.PivotTables("Rapportage3").PivotCache.Refresh
    Sheets("DB - Delta per straat").Select
    Application.ScreenUpdating = True
    
    End Sub
    
    
    
    Sub DTVernieuwenDBBilledSales()
'
' DTVernieuwenDBBilledSales Macro
'

'
    Application.ScreenUpdating = False
    Sheets("DB - Forecast").Select
    Range("B40").Select
    ActiveSheet.PivotTables("Forecast").PivotCache.Refresh
    Sheets("DB - FC vs Actuals").Select
    Range("B109").Select
    ActiveSheet.PivotTables("FCvsActuals").PivotCache.Refresh
    Sheets("DB - Delta per straat").Select
    Range("Y32").Select
    ActiveSheet.PivotTables("Delta").PivotCache.Refresh
    Sheets("DB - Billed sales").Select
    Range("B59").Select
    ActiveSheet.PivotTables("Billed sales").PivotCache.Refresh
    Sheets("DB - Historie").Select
    Range("B15").Select
    ActiveSheet.PivotTables("Historie").PivotCache.Refresh
    Sheets("Rapportage").Select
    Range("A8").Select
    ActiveSheet.PivotTables("Rapportage1").PivotCache.Refresh
    Range("A52").Select
    ActiveSheet.PivotTables("Rapportage2").PivotCache.Refresh
    Range("D52").Select
    ActiveSheet.PivotTables("Rapportage3").PivotCache.Refresh
    Sheets("DB - Billed Sales").Select
    Application.ScreenUpdating = True
    
    End Sub
    
    
    
    Sub DTVernieuwenDBHistorie()
'
' DTVernieuwenDBHistorie Macro
'

'
    Application.ScreenUpdating = False
    Sheets("DB - Forecast").Select
    Range("B40").Select
    ActiveSheet.PivotTables("Forecast").PivotCache.Refresh
    Sheets("DB - FC vs Actuals").Select
    Range("B109").Select
    ActiveSheet.PivotTables("FCvsActuals").PivotCache.Refresh
    Sheets("DB - Delta per straat").Select
    Range("Y32").Select
    ActiveSheet.PivotTables("Delta").PivotCache.Refresh
    Sheets("DB - Billed sales").Select
    Range("B59").Select
    ActiveSheet.PivotTables("Billed sales").PivotCache.Refresh
    Sheets("DB - Historie").Select
    Range("B15").Select
    ActiveSheet.PivotTables("Historie").PivotCache.Refresh
    Sheets("Rapportage").Select
    Range("A8").Select
    ActiveSheet.PivotTables("Rapportage1").PivotCache.Refresh
    Range("A52").Select
    ActiveSheet.PivotTables("Rapportage2").PivotCache.Refresh
    Range("D52").Select
    ActiveSheet.PivotTables("Rapportage3").PivotCache.Refresh
    Sheets("DB - Historie").Select
    Application.ScreenUpdating = True
    
    End Sub
    
    
    
    
    Sub DTVernieuwenRapportage()
'
' DTVernieuwenRapportage Macro
'

'
    Application.ScreenUpdating = False
    Sheets("DB - Forecast").Select
    Range("B40").Select
    ActiveSheet.PivotTables("Forecast").PivotCache.Refresh
    Sheets("DB - FC vs Actuals").Select
    Range("B109").Select
    ActiveSheet.PivotTables("FCvsActuals").PivotCache.Refresh
    Sheets("DB - Delta per straat").Select
    Range("Y32").Select
    ActiveSheet.PivotTables("Delta").PivotCache.Refresh
    Sheets("DB - Billed sales").Select
    Range("B59").Select
    ActiveSheet.PivotTables("Billed sales").PivotCache.Refresh
    Sheets("DB - Historie").Select
    Range("B15").Select
    ActiveSheet.PivotTables("Historie").PivotCache.Refresh
    Sheets("Rapportage").Select
    Range("A8").Select
    ActiveSheet.PivotTables("Rapportage1").PivotCache.Refresh
    Range("A52").Select
    ActiveSheet.PivotTables("Rapportage2").PivotCache.Refresh
    Range("D52").Select
    ActiveSheet.PivotTables("Rapportage3").PivotCache.Refresh
    Sheets("Rapportage").Select
    Application.ScreenUpdating = True
    
    End Sub

Sub CashVerwijderenHoofdblad()
Attribute CashVerwijderenHoofdblad.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CashVerwijderenHoofdblad Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Cash").Select
    ActiveSheet.Range("$A:$Q").AutoFilter Field:=1, Criteria1:="="
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A:$Q").AutoFilter Field:=1
    Range("I5").Select
    Sheets("Hoofdblad").Select
    Application.ScreenUpdating = True

End Sub


Sub CashVerwijderenCash()
'
' CashVerwijderenCash Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Cash").Select
    ActiveSheet.Range("$A:$Q").AutoFilter Field:=1, Criteria1:="="
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A:$Q").AutoFilter Field:=1
    Range("I5").Select
    Sheets("Cash").Select
    Application.ScreenUpdating = True

End Sub

Sub BilledSalesVerwijderenHoofdblad()
'
' BilledSalesVerwijderenHoofdblad Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Billed Sales").Select
    ActiveSheet.Range("$A:$O").AutoFilter Field:=14, Criteria1:="Actual"
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A:$O").AutoFilter Field:=14
    Range("I5").Select
    Sheets("Hoofdblad").Select
    Application.ScreenUpdating = True

End Sub


Sub BilledSalesVerwijderenBilledSales()
'
' BilledSalesVerwijderenBilledSales Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Billed Sales").Select
    ActiveSheet.Range("$A:$O").AutoFilter Field:=14, Criteria1:="Actual"
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A:$O").AutoFilter Field:=14
    Range("I5").Select
    Sheets("Billed sales").Select
    Application.ScreenUpdating = True

End Sub
Sub CashVerwijderenActuals()
'
' CashVerwijderenActuals Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Cash").Select
    ActiveSheet.Range("$A:$Q").AutoFilter Field:=1, Criteria1:="="
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A:$Q").AutoFilter Field:=1
    Range("I5").Select
    Sheets("Actuals - Cash").Select
    Application.ScreenUpdating = True
End Sub
Sub BilledSalesVerwijderenActuals()
'
' BilledSalesVerwijderenActuals Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Billed sales").Select
    ActiveSheet.Range("$A:$O").AutoFilter Field:=14, Criteria1:="Actual"
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A:$O").AutoFilter Field:=14
    Range("I5").Select
    Sheets("Actuals - Billed sales").Select
    Application.ScreenUpdating = True
End Sub
Sub CashInlezenHoofdblad()
Attribute CashInlezenHoofdblad.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CashInlezenHoofdblad Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Cash").Select
    Sheets("Actuals - Cash").Select
    Range("B2:O2", Range("B2:O2").End(xlDown)).Copy
    Sheets("Cash").Select
    Range("D" & Rows.Count).End(xlUp).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Hoofdblad").Select
    Application.ScreenUpdating = True
    
End Sub

Sub CashInlezenCash()
'
' CashInlezenCash Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Cash").Select
    Sheets("Actuals - Cash").Select
    Range("B2:O2", Range("B2:O2").End(xlDown)).Copy
    Sheets("Cash").Select
    Range("D" & Rows.Count).End(xlUp).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Cash").Select
    Application.ScreenUpdating = True
    
End Sub

Sub BilledSalesInlezenHoofdblad()
'
' BilledSalesInlezenHoofdblad Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Billed sales").Select
    Sheets("Actuals - Billed sales").Select
    Range("B2:M2", Range("B2:M2").End(xlDown)).Copy
    Sheets("Billed sales").Select
    Range("D" & Rows.Count).End(xlUp).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Hoofdblad").Select
    Application.ScreenUpdating = True
    
End Sub
Sub CashInlezenActuals()
'
' CashInlezenActuals Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Cash").Select
    Sheets("Actuals - Cash").Select
    Range("B2:O2", Range("B2:O2").End(xlDown)).Copy
    Sheets("Cash").Select
    Range("D" & Rows.Count).End(xlUp).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Actuals - Cash").Select
    Application.ScreenUpdating = True
    
End Sub

Sub BilledSalesInlezenActuals()
'
' BilledSalesInlezenActuals Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Billed sales").Select
    Sheets("Actuals - Billed sales").Select
    Range("B2:M2", Range("B2:M2").End(xlDown)).Copy
    Sheets("Billed sales").Select
    Range("D" & Rows.Count).End(xlUp).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Actuals - Billed sales").Select
    Application.ScreenUpdating = True
    
End Sub


Sub BilledSalesInlezenBilledSales()
'
' BilledSalesInlezenBilledSales Macro
'

'
    Application.ScreenUpdating = False
    Sheets("Billed sales").Select
    Sheets("Actuals - Billed sales").Select
    Range("B2:M2", Range("B2:M2").End(xlDown)).Copy
    Sheets("Billed sales").Select
    Range("D" & Rows.Count).End(xlUp).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Billed sales").Select
    Application.ScreenUpdating = True
    
End Sub
