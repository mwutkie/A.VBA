Attribute VB_Name = "Module13"
Sub next_stim()
Workbooks("T1bbdl_ts_final.xlsm").Activate
run1 = 2
Count = 1
Do Until IsEmpty(Cells(run1, 3))
    If Count = 26 Then
    Rows(run1).Insert
    Cells(run1, 3).Value = "GICS_SECTOR"
    Rows(run1).Insert
    Cells(run1, 3).Value = "FIANCIALS"
    Rows(run1).Insert
    Cells(run1, 3).Value = "GREEN"
    Rows(run1).Insert
    Cells(run1, 3).Value = "QE_STIMULUS"
    
    Count = 1
    run1 = run1 + 4
    End If
    
run1 = run1 + 1
Count = Count + 1
Loop
    
    





End Sub


