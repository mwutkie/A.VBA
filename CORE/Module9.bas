Attribute VB_Name = "Module9"
Sub test()
Workbooks("T1bbdl_ts_final.xlsm").Activate
Application.ScreenUpdating = False
run1 = 2
Count = 1
Do Until IsEmpty(Cells(run1, 3))
    If Count = 24 Then
    Rows(run1).Insert
    Cells(run1, 3).Value = "LOG_SIZE"
    Count = 1
    run1 = run1 + 1
    End If
    
run1 = run1 + 1
Count = Count + 1
Loop
    
    
Application.ScreenUpdating = True
End Sub



