VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
Call ClearFilters
    LastActivityTime = Now()
   
    ThisWorkbook.Application.StatusBar = "File will automatically close at " & Now() + TimeValue("00:09:00")

    
   ' Call GoToLastRow
    Call Check_Inactivity
    
    ThisWorkbook.Sheets("Request DB").Range("A4").Select
    
    
    End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    LastActivityTime = Now()
    ThisWorkbook.Application.StatusBar = "File will automatically close at " & Now() + TimeValue("00:09:00")
   
    End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    LastActivityTime = Now()
    ThisWorkbook.Application.StatusBar = "File will automatically close at " & Now() + TimeValue("00:09:00")
   
    End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    LastActivityTime = Now()
    ThisWorkbook.Application.StatusBar = "File will automatically close at " & Now() + TimeValue("00:09:00")

    End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
         
        Call StopTimer
Endd:
        
End Sub



Private Sub Workbook_Open111()
 Call ClearFilters

 Application.WindowState = xlMaximized
 
 
 End Sub

