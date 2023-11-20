Attribute VB_Name = "Module2_LIQ"
Sub m2()
'inserting values to dis
Workbooks("T1FMP_LIQ_ts.xlsm").Activate
Worksheets("Sheet1").Activate
run1 = 2
Count = 1
Do Until IsEmpty(Cells(run1, 3))
If Count = 7 Then
Cells(run1, 3).Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=STDEV(R[-6]C[-11]:R[-6]C)"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:BK1"), Type:= _
    xlFillDefault
    
Cells(run1, 3).Select
    ActiveCell.Offset(0, 12).Range("A1:BK1").Select
    Selection.Copy
    ActiveCell.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Cells(run1 + 1, 3).Select

    ActiveCell.Offset(0, 12).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=STDEV(R[-6]C[-11]:R[-6]C)"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:BK1"), Type:= _
    xlFillDefault
    
Cells(run1 + 1, 3).Select
    ActiveCell.Offset(0, 12).Range("A1:BK1").Select
    Selection.Copy
    ActiveCell.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
        
Cells(run1 + 2, 3).Select
    ActiveCell.Offset(0, 12).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=STDEV(R[-5]C[-11]:R[-5]C)"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:BK1"), Type:= _
        xlFillDefault
        
Cells(run1 + 2, 3).Select

    ActiveCell.Offset(0, 12).Range("A1:BK1").Select
    Selection.Copy
    ActiveCell.Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    run1 = run1 + 8
    Count = 1
    Cells(run1, 3).Select
    End If
    
run1 = run1 + 1
Count = Count + 1
Loop
End Sub

