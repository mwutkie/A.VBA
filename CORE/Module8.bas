Attribute VB_Name = "Module8"
Sub ret_all()
    Workbooks("T1bbdl_ts_final.xlsm").Activate
    Sheets.Add.Name = "Sheet2"
    Sheets("Sheet2").Select
    Windows("T1TBill_ts.xlsx").Activate
    Range("Q4:CK5").Select
    Application.CutCopyMode = False
    Selection.Copy
    Workbooks("T1bbdl_ts_final.xlsm").Activate
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Sheet1").Activate
run1 = 2
run3 = 17
run2 = 5
Count = 1
Do Until IsEmpty(Cells(run1, 3))
If Count = 23 Then




Cells(run1 - 2, 3).Select
    ActiveCell.Offset(0, 2).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-19]C/R[-19]C[-1]"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:BU1"), Type:= _
        xlFillDefault
Cells(run1, 3).Select
    
    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=R[-2]C-Sheet2!R[-21]C[-4]"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:BU1"), Type:= _
        xlFillDefault
        
        
Cells(run1 - 1, 3).Select

    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C-Sheet2!R[-21]C[-4]"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:BU1"), Type:= _
        xlFillDefault


Count = 0
End If
Count = Count + 1
run1 = run1 + 1
Loop
End Sub


