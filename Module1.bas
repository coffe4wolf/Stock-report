Attribute VB_Name = "Module1"
Option Explicit

Sub ShtProtection()
Attribute ShtProtection.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ShtProtection Macro
'

'
    Sheets("Stock report").Select
    ActiveSheet.Unprotect
    Sheets("Stock report").Select
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        AllowFormattingRows:=True, AllowInsertingHyperlinks:=True, AllowSorting:= _
        True, AllowFiltering:=True, AllowUsingPivotTables:=True
End Sub
