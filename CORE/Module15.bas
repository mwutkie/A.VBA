Attribute VB_Name = "Module15"
'fiiling isins again
Sub test2()
Workbooks("T1bbdl_ts_final.xlsm").Activate
run1 = 2
Count = 0
Do Until IsEmpty(Cells(run1, 3))
    If Count = 29 Then


    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 6, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 5, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 4, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 3, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 2, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 1, 1)
    
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 6, 2)
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 5, 2)
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 4, 2)
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 3, 2)
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 2, 2)
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 1, 2)
    
    
    Count = 0
    End If
    
run1 = run1 + 1
Count = Count + 1
Loop
    
    

End Sub
