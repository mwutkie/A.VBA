Attribute VB_Name = "Module17"
Sub ins_ESG()
Workbooks("T1bbdl_ts_final.xlsm").Activate
run1 = 2
ccount = 0
run2 = 4
run3 = 2
Count = 1
Workbooks("T1bbdl_ts_final.xlsm").Activate
Do Until IsEmpty(Cells(run1, 2))
    If Count = 13 Then
    Cells(run1, 2).Select
    isin_check = Cells(run1, 2)
    Cells(run1, 4).Select
    date_check = [D1]
        
        
        Do Until IsEmpty(Workbooks("OPTIMIZED_ESG_FINAL.xlsx").Worksheets("Sheet1").Cells(run3, 3))
            esg_isin = Workbooks("OPTIMIZED_ESG_FINAL.xlsx").Worksheets("Sheet1").Cells(run3, 3)
            
            If esg_isin = isin_check Then
            
                esg_date = Workbooks("OPTIMIZED_ESG_FINAL.xlsx").Worksheets("Sheet1").Cells(run3, 13)

                If date_check = esg_date Then
                    Workbooks("OPTIMIZED_ESG_FINAL.xlsx").Sheets("Sheet1").Activate
                    Range(Cells(run3, 5), Cells(run3, 12)).Select
                    Selection.Copy
                    Workbooks("T1bbdl_ts_final.xlsm").Activate
                    Cells(run1, 4).Select
                    Selection.PasteSpecial Transpose:=True
                    
                End If
                
            End If
            run3 = run3 + 1
            Loop
            run3 = 2
        run2 = run2 + 1
        
        
    
    Count = -16
    End If
run1 = run1 + 1
Count = Count + 1
Loop

End Sub

