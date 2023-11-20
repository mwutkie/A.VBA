Attribute VB_Name = "Module7_LIQ"
Sub ranking()
'rrnk age
run1 = 2
run2 = 4
run3 = 1
run4 = 1
Count = 1
Worksheets("Sheet1").Activate

Do Until IsEmpty(Cells(run1, 3))
    If Count = 6 Then
        Do Until IsEmpty(Cells(1, run2))
        'Cells(run1, run2) = WorksheetFunction.IfError(Cells(run1, run2), " ")
        Cells(run1, run2).Select
        Cells(run1, run2).Copy
        Worksheets("Temp").Cells(run3, run4).PasteSpecial xlPasteValues
        run2 = run2 + 1
        run4 = run4 + 1
        Loop
        Count = 0
        run1 = run1 + 8
        run3 = run3 + 1
        run2 = 4
        run4 = 1
    End If
Count = Count + 1
run1 = run1 + 1
Loop


Worksheets("Sheet1").Activate
 run1 = 2
 Count = 1
 run2 = 4
 run3 = 1
 run4 = 1
Do Until IsEmpty(Cells(run1, 3))
    If Count = 14 Then
    
        Do Until IsEmpty(Cells(1, run2))
            Cells(run1, run2).Select
            Selection.Value = Worksheets("rank").Cells(run3, run4)
            Selection.Value = WorksheetFunction.IfError(Selection.Value, " ")
            run4 = run4 + 1
            run2 = run2 + 1
            Count = 0
        Loop
        run3 = run3 + 1
        run4 = 1
        run1 = run1
        run2 = 4
        
        End If

run1 = run1 + 1
Count = Count + 1
Loop
End Sub

