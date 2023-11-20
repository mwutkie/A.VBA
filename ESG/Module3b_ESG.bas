Attribute VB_Name = "Module3b"
Sub m5b()
'ret
Sheets.Add.Name = "rank"

    Range("A1").Select
    ActiveCell.FormulaR1C1 = _
        "=RANK(Temp!RC,Temp!R1C:Temp!R3076C,0)+COUNTIF(Temp!R1C:Temp!RC,Temp!RC)"
    Range("A1").Select
    Selection.AutoFill Destination:=Range("A1:A3076"), Type:=xlFillDefault
    Range("A1:A3076").Select
    Range("A1:A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFill Destination:=Range("A1:BT3076"), Type:=xlFillDefault
    Range("A1:BT3076").Select
Worksheets("Sheet1").Activate
 run1 = 2
 Count = 1
 run2 = 15
 run3 = 1
 run4 = 1
Do Until IsEmpty(Cells(run1, 3))
    If Count = 23 Then
            Cells(run1, 3).Select
            Sheets("rank").Activate
            Range(Cells(run3, 1), Cells(run3, 75)).Copy
            Sheets("Sheet1").Activate
            Cells(run1, 4).PasteSpecial xlPasteValues
            Count = 0
        run3 = run3 + 1
        run4 = 1
        run1 = run1 + 1
        End If
run1 = run1 + 1
Count = Count + 1
Loop
    Cells.Replace What:="#N/A", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
            
            Cells.Replace What:="#VALUE!", Replacement:=" ", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        

End Sub




