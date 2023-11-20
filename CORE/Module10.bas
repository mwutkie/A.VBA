Attribute VB_Name = "Module10"
Sub test()
Workbooks("T1bbdl_ts_final.xlsm").Activate
run1 = 2
Count = 1
Do Until IsEmpty(Cells(run1, 3))
If Count = 24 Then
Cells(run1, 3).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=LN(R[-16]C)"
   ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:BV1"), Type:= _
        xlFillDefault
    ActiveCell.Range("A1:BV1").Select
    Count = 0
    End If
run1 = run1 + 1
Count = Count + 1
Loop
End Sub

