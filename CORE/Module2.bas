Attribute VB_Name = "Module2"
'isins
Sub test2()
Workbooks("T1bbdl_ts_final.xlsm").Activate
run1 = 1
Count = 0
Do Until IsEmpty(Cells(run1, 3))
    If Count = 13 Then
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
    Count = 1
    End If
    
run1 = run1 + 1
Count = Count + 1
Loop
Range("A47079:A47089").Value = "XS1574671662"
    

End Sub


