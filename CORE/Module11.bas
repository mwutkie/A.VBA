Attribute VB_Name = "Module11"
Sub age()
Workbooks("T1bbdl_ts_final.xlsm").Activate
run1 = 2
Count = 1
Do Until IsEmpty(Cells(run1, 3))
    If Count = 25 Then
    Rows(run1).Insert
    Cells(run1, 3).Value = "AGE"
    Count = 1
    run1 = run1 + 1
    End If
    
run1 = run1 + 1
Count = Count + 1
Loop
    
    

End Sub



