Attribute VB_Name = "Module14"
Sub QUE_STIM()
Workbooks("T1bbdl_ts_final.xlsm").Activate
run1 = 2
Count = 1
Do Until IsEmpty(Cells(run1, 3))
If Count = 29 Then
Cells(run1, 3).Select

Cells(run1 - 3, 3).Select

    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(R[-16]C=""Y"",1,0)"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:BV1"), Type:= _
        xlFillDefault
    ActiveCell.Range("A1:BV1").Select
    
Cells(run1 - 2, 3).Select

    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=IF(R[-16]C=""Y"",1,0)"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:BV1"), Type:= _
        xlFillDefault
    ActiveCell.Range("A1:BV1").Select

Cells(run1 - 1, 3).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(R[-15]C=""Asset Management & Custody Banks"",R[-15]C=""Consumer Finance"",R[-15]C=""Diversified Financials"",R[-15]C=""Investment Banking & Brokerage"",R[-15]C=""Multi-line Insurance & Brokerage"",R[-15]C=""Banks""),1,0)"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:BV1"), Type:= _
        xlFillDefault

Cells(run1, 3).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(R[-16]C,'[Thomas Merz - GICS_sectors.xlsx]GICS Sectors'!C4:C8,5,FALSE)"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:BV1"), Type:= _
        xlFillDefault
        
        
Count = 0
End If
run1 = run1 + 1
Count = Count + 1
Loop
End Sub


