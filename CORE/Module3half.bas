Attribute VB_Name = "Module3half"
Sub nadelete()
Attribute nadelete.VB_ProcData.VB_Invoke_Func = " \n14"
'
' nadelete Macro
'

'   Workbooks("T1bbdl_ts_final.xlsm").Activate
    Rows("1:1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$BW$47090").AutoFilter Field:=2, Criteria1:="#N/A"
    Rows("26:26").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=21
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range("$A$1:$BW$37658").AutoFilter Field:=2
End Sub
