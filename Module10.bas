Attribute VB_Name = "Module10"
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    Range("A4:G1701").Select
    ActiveSheet.Unprotect
    Selection.ClearContents
    Sheets("Request DB").Select
    Range("A4:X256").Select
    Range("X256").Activate
    Selection.ClearContents
End Sub
