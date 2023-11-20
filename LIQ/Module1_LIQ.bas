Attribute VB_Name = "Module1_LIQ"
Sub create_LIQ_ts()
'adding rows + first values
ActiveSheet.Name = "Sheet1"
Sheets.Add.Name = "Temp"
Sheets.Add.Name = "rank"
Sheets("Sheet1").Activate
'T1bbdl_ts_final
'T1bbdl_ts_final
run1 = 2
run2 = 2
run3 = 2
run4 = 2
Count2 = 1
Count = 1
Workbooks("T1bbdl_ts_final.xlsm").Activate
Range("A1", "BY1").Copy
Workbooks("T1FMP_LIQ_ts.xlsm").Activate
Range("A1", "BY1").PasteSpecial
Workbooks("T1bbdl_ts_final.xlsm").Activate
Do Until IsEmpty(Cells(run1, 1))
    If Count = 3 Then
    Range(Cells(run1, 1), Cells(run1, 77)).Copy
    
    Workbooks("T1FMP_LIQ_ts.xlsm").Activate
    Range(Cells(run2, 1), Cells(run2, 77)).PasteSpecial xlPasteValues
    
    Workbooks("T1bbdl_ts_final.xlsm").Activate
    Range(Cells(run1 + 2, 1), Cells(run1 + 2, 77)).Copy
    
    Workbooks("T1FMP_LIQ_ts.xlsm").Activate
    Range(Cells(run2 + 1, 1), Cells(run2 + 1, 77)).PasteSpecial xlPasteValues
    
    Workbooks("T1bbdl_ts_final.xlsm").Activate
    Range(Cells(run1 + 5, 1), Cells(run1 + 5, 77)).Copy
    
    Workbooks("T1FMP_LIQ_ts.xlsm").Activate
    Range(Cells(run2 + 2, 1), Cells(run2 + 2, 77)).PasteSpecial xlPasteValues
    
    Workbooks("T1bbdl_ts_final.xlsm").Activate
    Range(Cells(run1 + 18, 1), Cells(run1 + 18, 77)).Copy
    
    Workbooks("T1FMP_LIQ_ts.xlsm").Activate
    Range(Cells(run2 + 3, 1), Cells(run2 + 3, 77)).PasteSpecial xlPasteValues
    
    Workbooks("T1bbdl_ts_final.xlsm").Activate
    Range(Cells(run1 + 21, 1), Cells(run1 + 21, 77)).Copy
    
    Workbooks("T1FMP_LIQ_ts.xlsm").Activate
    Range(Cells(run2 + 4, 1), Cells(run2 + 4, 77)).PasteSpecial xlPasteValues
    
    Workbooks("T1bbdl_ts_final.xlsm").Activate
    Range(Cells(run1 + 22, 1), Cells(run1 + 22, 77)).Copy
    
    Workbooks("T1FMP_LIQ_ts.xlsm").Activate
    Range(Cells(run2 + 5, 1), Cells(run2 + 5, 77)).PasteSpecial xlPasteValues
    Cells(run2 + 6, 3).Value = "dis_yield"
    Cells(run2 + 7, 3).Value = "dis_oas"
    Cells(run2 + 8, 3).Value = "dis_ret"
    Cells(run2 + 9, 3).Value = "rnk_dis_yield"
    Cells(run2 + 10, 3).Value = "rnk_dis_oas"
    Cells(run2 + 11, 3).Value = "rnk_dis_ret"
    Cells(run2 + 12, 3).Value = "rnk_log_size"
    Cells(run2 + 13, 3).Value = "rnk_age"
    run2 = run2 + 14
    run1 = run1 + 26
    

    Count = 0
End If
Workbooks("T1bbdl_ts_final.xlsm").Activate
Count = Count + 1
run1 = run1 + 1
Loop



Workbooks("T1FMP_LIQ_ts.xlsm").Activate
run3 = 2
Count2 = 1
Do Until IsEmpty(Cells(run3, 3))
    If Count2 = 14 Then
    Cells(run3, 1).Select
    Cells(run3 - 12, 1).Copy Destination:=Cells(run3 - 11, 1)
    Cells(run3 - 12, 1).Copy Destination:=Cells(run3 - 10, 1)
    Cells(run3 - 12, 1).Copy Destination:=Cells(run3 - 9, 1)
    Cells(run3 - 12, 1).Copy Destination:=Cells(run3 - 8, 1)
    Cells(run3 - 12, 1).Copy Destination:=Cells(run3 - 7, 1)
    Cells(run3 - 12, 1).Copy Destination:=Cells(run3 - 6, 1)
    Cells(run3 - 12, 1).Copy Destination:=Cells(run3 - 5, 1)
    Cells(run3 - 12, 1).Copy Destination:=Cells(run3 - 4, 1)
    Cells(run3 - 12, 1).Copy Destination:=Cells(run3 - 3, 1)
    Cells(run3 - 12, 1).Copy Destination:=Cells(run3 - 1, 1)
    Cells(run3 - 12, 1).Copy Destination:=Cells(run3 - 2, 1)
    Cells(run3 - 12, 1).Copy Destination:=Cells(run3, 1)
    Count2 = 0
    End If
Cells(run3, 3).Select
run3 = run3 + 1
Count2 = Count2 + 1
Loop



run3 = 2
Count2 = 1
Do Until IsEmpty(Cells(run3, 3))
    If Count2 = 14 Then
    Cells(run3, 2).Select
    Cells(run3 - 12, 2).Copy Destination:=Cells(run3 - 11, 2)
    Cells(run3 - 12, 2).Copy Destination:=Cells(run3 - 10, 2)
    Cells(run3 - 12, 2).Copy Destination:=Cells(run3 - 9, 2)
    Cells(run3 - 12, 2).Copy Destination:=Cells(run3 - 8, 2)
    Cells(run3 - 12, 2).Copy Destination:=Cells(run3 - 7, 2)
    Cells(run3 - 12, 2).Copy Destination:=Cells(run3 - 6, 2)
    Cells(run3 - 12, 2).Copy Destination:=Cells(run3 - 5, 2)
    Cells(run3 - 12, 2).Copy Destination:=Cells(run3 - 4, 2)
    Cells(run3 - 12, 2).Copy Destination:=Cells(run3 - 3, 2)
    Cells(run3 - 12, 2).Copy Destination:=Cells(run3 - 1, 2)
    Cells(run3 - 12, 2).Copy Destination:=Cells(run3 - 2, 2)
    Cells(run3 - 12, 2).Copy Destination:=Cells(run3, 2)
    Count2 = 0
    End If
Cells(run3, 3).Select
run3 = run3 + 1
Count2 = Count2 + 1
Loop
End Sub






