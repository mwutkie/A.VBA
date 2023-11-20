Attribute VB_Name = "Module5"
'pasting values
Sub step3()
Workbooks("T1bbdl_ts_final.xlsm").Activate
run1 = 2
Do Until IsEmpty(Cells(run1, 2))
If Count = 12 Then
    Cells(run1 - 7, 4).Copy
    Range(Cells(run1 - 7, 5), Cells(run1 - 7, 77)).PasteSpecial
    
    Cells(run1 - 6, 4).Copy
    Range(Cells(run1 - 6, 5), Cells(run1 - 6, 77)).PasteSpecial
    
    Cells(run1 - 5, 4).Copy
    Range(Cells(run1 - 5, 5), Cells(run1 - 5, 77)).PasteSpecial
    
    Cells(run1 - 4, 4).Copy
    Range(Cells(run1 - 4, 5), Cells(run1 - 4, 77)).PasteSpecial
    
    Cells(run1 - 3, 4).Copy
    Range(Cells(run1 - 3, 5), Cells(run1 - 3, 77)).PasteSpecial
    
    Cells(run1 - 2, 4).Copy
    Range(Cells(run1 - 2, 5), Cells(run1 - 2, 77)).PasteSpecial
    
    Cells(run1 - 1, 4).Copy
    Range(Cells(run1 - 1, 5), Cells(run1 - 1, 77)).PasteSpecial
    Count = 0
    End If
    Count = Count + 1
    run1 = run1 + 1
    Loop
End Sub

