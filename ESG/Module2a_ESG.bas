Attribute VB_Name = "Module2a"
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
    If Count = 20 Then
        Range(Cells(run1, 4), Cells(run1, 75)).Select
        Range(Cells(run1, 4), Cells(run1, 75)).Copy
        Worksheets("Temp").Cells(run3, 1).PasteSpecial xlPasteValues

        Count = 0
        run1 = run1 + 4
        run3 = run3 + 1
    End If
    
    
Count = Count + 1
run1 = run1 + 1
Loop
Sheets("Temp").Activate
    Cells.Replace What:="#DIV/0!", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
End Sub

