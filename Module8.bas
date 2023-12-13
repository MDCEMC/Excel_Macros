Attribute VB_Name = "Module8"
Option Explicit

Public LastActivityTime As Date
Public ActivityTime As Date


Sub Check_Inactivity()
    'Const Inactivity_Delay As Date = #12:02:00 AM#
     
     
    If LastActivityTime + TimeValue("00:09:00") < Now() Then
        ThisWorkbook.Application.StatusBar = "File Timer Stopped"
        ThisWorkbook.Close SaveChanges:=True
    Else
        ActivityTime = LastActivityTime
        ThisWorkbook.Application.OnTime ActivityTime + TimeValue("00:09:00"), "Check_Inactivity"
        End If
    End Sub

Sub StopTimer()
   ' Const Inactivity_Delay As Date = #12:02:00 AM#
   On Error GoTo Endd
   ThisWorkbook.Application.OnTime ActivityTime + TimeValue("00:09:00"), "Check_Inactivity", , False
Endd:
   ThisWorkbook.Application.StatusBar = "File Timer Stopped"
End Sub




