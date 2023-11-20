Attribute VB_Name = "Module12"
Sub m12()
'age calcualted
Workbooks("T1bbdl_ts_final.xlsm").Activate
run1 = 2
run2 = 4
Count = 1
Do Until IsEmpty(Cells(run1, 3))
If Count = 25 Then
        
        Do Until IsEmpty(Cells(run1 - 1, run2))
        Cells(run1, run2).Select
        Cells(run1, run2) = DateDiff("d", Cells(run1 - 19, run2), Cells(1, run2))
        If Cells(run1, run2) < 0 Then
            Cells(run1, run2) = ""
            End If
        Count = 0
        run2 = run2 + 1
        Loop
End If
run2 = 4
run1 = run1 + 1
Count = Count + 1
Loop
End Sub

