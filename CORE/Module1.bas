Attribute VB_Name = "Module1"
Sub test()
'only first run
Workbooks("T1bbdl_ts_final.xlsm").Activate
Columns("A").Delete
Columns("BZ").Delete
Columns("BY").Delete
Columns("BX").Delete
[A1].Value = "ISIN"
[B1].Value = "PARENT_ISIN"


run1 = 2
Count = 0
Do Until IsEmpty(Cells(run1, 3))
    If Count = 5 Then
    Rows(run1).Insert
    Cells(run1, 3).Value = "ISSUER COUNTRY"
    Rows(run1).Insert
    Cells(run1, 3).Value = "GREEN_BOND_LOAN_INDICATOR"
    Rows(run1).Insert
    Cells(run1, 3).Value = "CURRENTLY_EUROPEAN_CENT_BK_ELIG"
    Rows(run1).Insert
    Cells(run1, 3).Value = "MATURITY"
    Rows(run1).Insert
    Cells(run1, 3).Value = "PAR_AMT"
    Rows(run1).Insert
    Cells(run1, 3).Value = "CPN"
    Rows(run1).Insert
    Cells(run1, 3).Value = "ISSUE_DT"
    Count = 0
    run1 = run1 + 7
    End If
    
run1 = run1 + 1
Count = Count + 1
Loop
[c47083].Value = "ISSUE_DT"
[c47084].Value = "CPN"
[c47085].Value = "PAR_AMT"
[c47086].Value = "MATURITY"
[c47087].Value = "CURRENTLY_EUROPEAN_CENT_BK_ELIG"
[c47088].Value = "GREEN_BOND_LOAN_INDICATOR"
[c47089].Value = "ISSUER COUNTRY"
[c47090].Value = "x"
[b47090].Value = "x"
[a47090].Value = "x"
End Sub


