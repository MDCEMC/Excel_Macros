Attribute VB_Name = "Module15"

  



Sub OpenCommonTestPlan()

  Dim objWord

   Dim objDoc
Application.DisplayAlerts = False
   Set objWord = CreateObject("Word.Application")

   Set objDoc = objWord.Documents.Open("https://shiftup.sharepoint.com/sites/NAEMCEngineering/Shared Documents/General/Operations/Common Immunity Test Plan.docx")

   objWord.Visible = True
Application.DisplayAlerts = True
End Sub


Sub AddNewRequestToMaster()

'------------------------------------------------------------
  Dim wbe As Workbook ' engineers workbook
  
    Set wbe = ThisWorkbook
    
  
    'WbkRequest = ActiveWindow.Caption

    wbe.Sheets("Request DB").Unprotect
    wbe.Application.DisplayAlerts = False
    wbe.Application.Calculate
    Req = wbe.Sheets("Request DB").Range("e2") + 1
    RowNo = wbe.Sheets("Request DB").Range("C2") + 4
    wbe.Sheets("Request DB").Rows(RowNo).Select
    
    wbe.Sheets("Request DB").Cells(RowNo, 1).Value = Req
    wbe.Sheets("Request DB").Cells(RowNo, 2).Value = Req
    wbe.Application.Calculate
    
    Call SortRequestHiToLo
    Cells(4, 1).Select
    
    Application.WindowState = xlNormal
        
  End Sub

Sub GetScheduleFromMaster()
'
' Macro2 Macro
'

'
Dim wbe As Workbook ' engineers workbook
Dim wbm As Workbook ' master workbook

    Set wbe = ThisWorkbook
    wbe.Application.ScreenUpdating = False
    Set wbm = Workbooks.Open("https://shiftup.sharepoint.com/sites/NAEMCEngineering/Shared Documents/General/Operations/Master Vehicle Schedule.xlsm")
    wbm.Application.ScreenUpdating = False
     
    wbe.Sheets("Schedule").Unprotect
    wbe.Application.DisplayAlerts = False
        
    wbm.Sheets("Schedule").Unprotect
    wbm.Sheets("Schedule").Rows("3:35").Copy

     wbe.Sheets("Schedule").Rows("3:3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    'Range("B3").Select
    'ActiveWindow.ScrollRow = 5
    'Windows(WbkRequest).Activate
    'Rows("3:3").Select
    'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
      '  :=False, Transpose:=False
    
    'Windows(WbkSched).Activate
    wbm.Application.ScreenUpdating = True
    wbm.Close
    wbe.Application.DisplayAlerts = True
    'wbe.Sheets("Schedule").Select
    'Range("6:8,11:13,16:18,21:23,26:28,31:33,36:38,41:43,46:48,51:53,56:58,61:63,66:68").Select
    wbe.Sheets("Schedule").Range("A6").Activate
    'Selection.EntireRow.Hidden = True
    wbe.Sheets("Schedule").Protect
    wbe.Application.ScreenUpdating = True
    
   Call GoToToday
   
   
End Sub
