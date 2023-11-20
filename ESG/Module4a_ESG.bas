Attribute VB_Name = "Module4a"
Sub m4a()
' rank
Worksheets("Sheet1").Activate
Sheets("rank").Delete
run1 = 2
run2 = 15
run3 = 1
run4 = 1
Count = 1
Worksheets("Sheet1").Activate

Do Until IsEmpty(Cells(run1, 3))
    If Count = 16 Then
        Range(Cells(run1, 4), Cells(run1, 75)).Select
        Range(Cells(run1, 4), Cells(run1, 75)).Copy
        Worksheets("Temp").Cells(run3, 1).PasteSpecial xlPasteValues

        Count = 0
        run1 = run1 + 8
        run3 = run3 + 1
    End If
    
    
Count = Count + 1
run1 = run1 + 1
Loop

Sheets("Temp").Activate

Range("BT3076:a1").Select

Selection.SpecialCells(xlCellTypeBlanks).Select
Application.CutCopyMode = False
Selection.FormulaR1C1 = "-"
Range("AI22").Select
End Sub

