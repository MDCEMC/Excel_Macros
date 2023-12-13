Attribute VB_Name = "Module11"

Sub CopySheetToClosedMFE()


ThisWorkbook.Activate
    Req = Sheets("MFE Sheet").Range("C2").Value
    MYModel = Sheets("MFE Sheet").Range("C4").Value
    CY = "20" & Left(Req, 2)
    
    ff = Req & " " & MYModel & " " & "MFE Data Sheet.xlsx"
    filen = "J:\5140_J Drive\Vehicle Testing\MFE Data Sheets\" & CY & "\" & ff
   
    On Error GoTo DIROK
    MkDir "J:\5140_J Drive\Vehicle Testing\MFE Data Sheets\" & CY
DIROK:
    
Application.ScreenUpdating = False
Dim strFileExists As String
 
  ThisWorkbook.Activate
    strFileExists = Dir(filen)
 
   If strFileExists = "" Then
        
        Sheets(Array("MFE Sheet", "MFE2 Sheet")).Select
        Sheets("MFE Sheet").Activate
        Sheets(Array("MFE Sheet", "MFE2 Sheet")).Copy
        
        Range("A1:L40").Select
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
        ActiveWorkbook.SaveAs Filename:=filen
        ActiveWorkbook.Close
    
    
    Else
       ' MsgBox "The selected file exists"
    End If


ThisWorkbook.Activate
    Application.ScreenUpdating = True


End Sub

Sub CopySheetToClosedRNE()

ThisWorkbook.Activate
    Req = Sheets("RNE Sheet").Range("B2").Value
    
    MYModel = Sheets("RNE Sheet").Range("B4").Value
    CY = "20" & Left(Req, 2)
    On Error GoTo DIROK
    MkDir "J:\5140_J Drive\Vehicle Testing\RNE Data Sheets\" & CY
DIROK:
    ff = Req & " " & MYModel & " " & "RNE Data Sheet.xlsx"
    filen = "J:\5140_J Drive\Vehicle Testing\RNE Data Sheets\" & CY & "\" & ff
     
    
Application.ScreenUpdating = False
Dim strFileExists As String
 ThisWorkbook.Activate
   
    strFileExists = Dir(filen)
 
   If strFileExists = "" Then
        
        Sheets("RNE Sheet").Select
        Sheets("RNE Sheet").Activate
        Sheets("RNE Sheet").Copy
          ActiveWindow.ScrollRow = 5
    
    Range("A1:J10").Select
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        ActiveWorkbook.SaveAs Filename:=filen
        ActiveWorkbook.Close
    
    
    Else
       ' MsgBox "The selected file exists"
    End If


ThisWorkbook.Activate
    Application.ScreenUpdating = True


End Sub




