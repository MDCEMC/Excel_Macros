Attribute VB_Name = "Module12"
Sub Autofit()
Attribute Autofit.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
On Error GoTo quitt
ThisWorkbook.Activate
    Rows("5:60").Select
    Selection.Rows.Autofit
    
    Sheets("Schedule").Select
    Range("6:8,11:13,16:18,21:23,26:28,31:33,36:38,41:43,46:48,51:53,56:58,61:63,66:68").Select
    Range("A6").Activate
    Selection.EntireRow.Hidden = True
quitt:
End Sub
Sub GoToToday()
'
' Macro1 Macro
'
ThisWorkbook.Activate
Sheets("Schedule").Select
   ' Worksheets("Schedule").Unprotect
   Cells(3, 3).Select
    Dattee = Range("B1")
 
 Cells.Find(What:=Dattee, After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=True).Activate

'Cells.Find(What:=Dattee, After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=True).Activate
    

End Sub

