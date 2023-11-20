Attribute VB_Name = "Module3"
Sub tester()
    Workbooks("T1bbdl_ts_final.xlsm").Activate
    [b2].Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'D:\DOKUMENTACJA\Python\FIX_PROJECT\[Thomas Merz - bbcs_eur_stacked_manipulated.xlsx]bbcs_eur_stacked'!C1:C2,2,FALSE)"
    Range("B2").Select
    Selection.AutoFill Destination:=Range("B2:B47089"), Type:=xlFillDefault
    Range("B2:B47089").Select
End Sub


