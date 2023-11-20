Attribute VB_Name = "Module7"
'fiiling isins again
Sub test2()
Workbooks("T1bbdl_ts_final.xlsm").Activate
run1 = 2
Count = 0
Do Until IsEmpty(Cells(run1, 3))
    If Count = 23 Then
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 20, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 19, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 18, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 17, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 16, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 15, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 14, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 13, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 12, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 11, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 10, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 9, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 8, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 7, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 6, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 5, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 4, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 3, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 2, 1)
    Cells(run1 - 12, 1).Copy Destination:=Cells(run1 - 1, 1)
    
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 20, 2)
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 19, 2)
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 18, 2)
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 17, 2)
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 16, 2)
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 15, 2)
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 14, 2)
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 13, 2)
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 12, 2)
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 11, 2)
        Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 10, 2)
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 9, 2)
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 8, 2)
    Cells(run1 - 12, 2).Copy Destination:=Cells(run1 - 7, 2)
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
