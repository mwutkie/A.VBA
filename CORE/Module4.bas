Attribute VB_Name = "Module4"
'importing values
'changed form -2 to -3
Sub step3()
Workbooks("T1bbdl_ts_final.xlsm").Activate
run1 = 2
testr = "x"
Do Until IsEmpty(Cells(run1, 3))
    If Count = 12 Then
    Set myrange = Workbooks("T1bbdl_cs_final.xlsx").Worksheets("Sheet1").Columns("B:M")
    result = Application.VLookup(Cells(run1 - 5, 1), myrange, 7, False)
    result2 = Application.VLookup(Cells(run1 - 5, 1), myrange, 8, False)
    result3 = Application.VLookup(Cells(run1 - 5, 1), myrange, 9, False)
    result4 = Application.VLookup(Cells(run1 - 5, 1), myrange, 10, False)
    result5 = Application.VLookup(Cells(run1 - 5, 1), myrange, 11, False)
    result6 = Application.VLookup(Cells(run1 - 5, 1), myrange, 12, False)
    result7 = Application.VLookup(Cells(run1 - 5, 1), myrange, 7, False)
    Cells(run1 - 7, 4).NumberFormat = "@"
    Cells(run1 - 7, 4) = result
    Cells(run1 - 6, 4) = result2
    Cells(run1 - 5, 4) = result3
    Cells(run1 - 4, 4).NumberFormat = "@"
    Cells(run1 - 4, 4) = result4
    Cells(run1 - 3, 4) = result5
    Cells(run1 - 2, 4) = result6
    Cells(run1 - 1, 4) = Left(Cells(run1 - 1, 1), 2)
    
    Count = 0
    End If
    Count = Count + 1
    run1 = run1 + 1
    Loop
End Sub



