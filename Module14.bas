Attribute VB_Name = "Module14"
Sub ShowYearlyUpdateTabs()
Attribute ShowYearlyUpdateTabs.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
  
    Sheets("Older Requests").Visible = True
    Sheets("Older Requests").Unprotect
    Sheets("Older TestPlan DB").Visible = True
    Sheets("Older TestPlan DB").Unprotect
    Sheets("TestPlan DB").Visible = True
    Sheets("TestPlan DB").Unprotect
    Sheets("Yearly Update Directions").Select
  
End Sub
Sub HideYearlyUpdateTabs()
'
' Macro2 Macro
'

'
    ThisWorkbook.Activate
    Sheets("Older Requests").Protect
    Sheets("Older Requests").Visible = False
    Sheets("Older TestPlan DB").Protect
    Sheets("Older TestPlan DB").Visible = False
    Sheets("TestPlan DB").Protect
    Sheets("TestPlan DB").Visible = False
    Sheets("Yearly Update Directions").Select
  
End Sub
Sub YearlyUpdateInfo()
Attribute YearlyUpdateInfo.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    ThisWorkbook.Activate
    Sheets("Yearly Update Directions").Visible = True

    Sheets("Yearly Update Directions").Select

End Sub
Sub HideYearlyUpdateInfoTab()
'
' Macro2 Macro
'

'

    Sheets("Yearly Update Directions").Visible = False

    Sheets("Request DB").Select
  
End Sub
