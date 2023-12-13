Attribute VB_Name = "Module9"
Sub PrintRNESheet()

ThisWorkbook.Activate
Calculate

Sheets("RNE Sheet").Visible = True
Sheets("RNE Sheet").Select
Calculate

' print Page setup
 
 aa = Sheets("RNE Sheet").Range("J9").Value
 Sheets("RNE Sheet").PageSetup.PrintArea = aa
 Range("A6").Select
  On Error Resume Next
  
      ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
        
On Error GoTo 0
Call CopySheetToClosedRNE

  Sheets("RNE Sheet").Visible = False
Sheets("Request DB").Select
End Sub
 
Sub PrintMFESheet()
ThisWorkbook.Activate
Calculate
Sheets("MFE2 Sheet").Visible = True
Sheets("MFE Sheet").Visible = True
Sheets("MFE Sheet").Select
 Calculate

 
' print Page setup
 

 Range("A6").Select
  On Error Resume Next
      ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False

  On Error GoTo 0
  
  Call CopySheetToClosedMFE
  Sheets("MFE2 Sheet").Visible = False
  Sheets("MFE Sheet").Visible = False
  
Sheets("Request DB").Select
End Sub
 




