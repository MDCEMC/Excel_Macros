Attribute VB_Name = "Module19"
Sub BackToReq()
Attribute BackToReq.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    ThisWorkbook.Activate
    Sheets("Request DB").Select
    Sheets("Version").Visible = False
    Sheets("Request DB").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
    
End Sub
Sub GoToVersionScreen()
'
' Macro3 Macro
'

'
    ThisWorkbook.Activate
    Sheets("Version").Visible = True
    Sheets("Version").Select
    
    
End Sub





