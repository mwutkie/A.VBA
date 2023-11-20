Attribute VB_Name = "Module1"
Sub create_LIQ_ts()
'adding rows + first values
'Sheets.Add.Name = "Temp"
'Sheets.Add.Name = "rank"
Workbooks("T1FMP_ESG_ts.xlsm").Activate
ActiveSheet.Name = "Sheet1"
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
Workbooks("T1FMP_ESG_ts.xlsm").Activate
Range("A1", "BY1").PasteSpecial
Workbooks("T1bbdl_ts_final.xlsm").Activate
Do Until IsEmpty(Cells(run1, 1))
    If Count = 1 Then
    Workbooks("T1bbdl_ts_final.xlsm").Activate
    
    Range(Cells(run1, 1), Cells(run1 + 20, 77)).Copy
    Workbooks("T1FMP_ESG_ts.xlsm").Activate
    Range(Cells(run2, 1), Cells(run2 + 20, 77)).PasteSpecial xlPasteValues
 
    
    Cells(run2 + 23, 3).Value = "rnk_weighted_score"
    Cells(run2 + 22, 3).Value = "rnk_adj_score"
    Cells(run2 + 21, 3).Value = "rnk_iva_comp_num"

    run2 = run2 + 24
    run1 = run1 + 28

    Count = 0
End If
Workbooks("T1bbdl_ts_final.xlsm").Activate
Count = Count + 1
run1 = run1 + 1
Loop



Workbooks("T1FMP_ESG_ts.xlsm").Activate
run3 = 2
Count2 = 1
Do Until IsEmpty(Cells(run3, 3))
    If Count2 = 24 Then
    Cells(run3, 1).Select

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
    If Count2 = 24 Then

    Cells(run3 - 12, 2).Copy Destination:=Cells(run3 - 1, 2)
    Cells(run3 - 12, 2).Copy Destination:=Cells(run3 - 2, 2)
    Cells(run3 - 12, 2).Copy Destination:=Cells(run3, 2)
    Count2 = 0
    End If
Cells(run3, 3).Select
run3 = run3 + 1
Count2 = Count2 + 1
Loop
Range("A73826:C73849").Delete


End Sub




