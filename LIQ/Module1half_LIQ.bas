Attribute VB_Name = "Module1half_LIQ"
Sub proxy_rank_matrix()
Attribute proxy_rank_matrix.VB_ProcData.VB_Invoke_Func = " \n14"
' proxy_rank_matrix Macro
    Sheets("rank").Select
    ActiveCell.FormulaR1C1 = _
        "=RANK(Temp!RC,Temp!R1C:R20C,0)+COUNTIF(Temp!R1C:RC,Temp!RC[4])"
    Selection.AutoFill Destination:=Range("A1:A4361"), Type:=xlFillDefault
    Range("A1:A4361").Select
    Selection.AutoFill Destination:=Range("A1:BZ4361"), Type:=xlFillDefault
    Range("A1:BZ4361").Select
    Sheets("Sheet1").Select
End Sub
