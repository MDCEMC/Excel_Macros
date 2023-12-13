Attribute VB_Name = "Module6"

Sub ClearFilters()
'
'  Macro
'
    

    ThisWorkbook.Activate
    Worksheets("Request DB").Select
    Sheets("Request DB").Unprotect
    Rng = "C2"
    Indexx = Range(Rng)
    
    If Rng > 11 Then
       Indexx = Indexx - 10
       End If
       
       
    Rowss = Indexx & ":" & Indexx
    
    
    Range("A3:Y3").Select
    
    On Error Resume Next
    ActiveSheet.ShowAllData
    Sheets("Request DB").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
   
  
  Range(Rowss).Select
  

        
        
End Sub

Sub Test1()
Dim btn As Button, wks As Worksheet
For Each wks In Worksheets
For Each btn In wks.Buttons
MsgBox _
"Sheet name:" & vbTab & wks.Name & vbCrLf & _
"Button name:" & vbTab & btn.Name & vbCrLf & _
"Macro name:" & vbTab & btn.OnAction, , "Sheet button macros"
Next btn
Next wks
End Sub
