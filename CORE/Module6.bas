Attribute VB_Name = "Module6"
'adding empty "dates"
Sub test()
Workbooks("T1bbdl_ts_final.xlsm").Activate
run1 = 2
Count = 0
Do Until IsEmpty(Cells(run1, 3))
    If Count = 12 Then
    Rows(run1).Insert
    Cells(run1, 3).Value = "exc_ret_i26680eu"
    Rows(run1).Insert
    Cells(run1, 3).Value = "exc_ret_leb2treu"
    Rows(run1).Insert
    Cells(run1, 3).Value = "ret"
    Rows(run1).Insert
    Cells(run1, 3).Value = "IVA_COMPANY_RATING_NUM"
    Rows(run1).Insert
    Cells(run1, 3).Value = "GOVERNANCE_PILLAR_SCORE"
    Rows(run1).Insert
    Cells(run1, 3).Value = "SOCIAL_PILLAR_SCORE"
    Rows(run1).Insert
    Cells(run1, 3).Value = "ENVIRONMENTAL_PILLAR_SCORE"
    Rows(run1).Insert
    Cells(run1, 3).Value = "WEIGHTED_AVERAGE_SCORE"
    Rows(run1).Insert
    Cells(run1, 3).Value = "INDUSTRY_ADJUSTED_SCORE"
    Rows(run1).Insert
    Cells(run1, 3).Value = "IVA_COMPANY_RATING"
    Rows(run1).Insert
    Cells(run1, 3).Value = "IVA_INDUSTRY"
    Count = 0
    run1 = run1 + 11
    End If
    
run1 = run1 + 1
Count = Count + 1
Loop
    
    
Application.ScreenUpdating = True
End Sub

