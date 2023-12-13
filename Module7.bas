Attribute VB_Name = "Module7"

Sub RoundedRectangle5_Click()
 If ActiveWorkbook.ReadOnly Then
         ActiveSheet.Shapes.Range(Array("Rounded Rectangle 4")).Visible = False
         ActiveSheet.Shapes.Range(Array("Rounded Rectangle 1")).Visible = False
         ActiveSheet.Shapes.Range(Array("Rounded Rectangle 2")).Visible = False
         Range("A2").Value = "File Checked out"
         
         Exit Sub
        End If


  Sheets("Request DB").Select
    Worksheets("Request DB").Unprotect
    maxno = Cells(2, 3).Value + 3
    RowNo = ActiveCell.Row
    ColNo = ActiveCell.Column
    If (RowNo < 4) Then
     MsgBox "Please Highlight an active row!"
     GoTo quit
     End If
         
    If (RowNo > maxno) Then
     MsgBox "Please Highlight an active row!"
     GoTo quit
     End If
    
    Emissions.Show
quit:
 
 
 


End Sub


