Attribute VB_Name = "Module16"
Sub ShowOlderReqs()
Attribute ShowOlderReqs.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Sheets("Older Requests").Visible = True
    Sheets("Older Requests").Select
End Sub
Sub HideOlderReqs()
'
' Macro2 Macro
'

'
    Sheets("Older Requests").Visible = False
    Sheets("Request DB").Select
End Sub

