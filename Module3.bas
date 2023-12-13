Attribute VB_Name = "Module3"
Sub RoundedRectangle3_Click()
    
    ThisWorkbook.Activate
    Sheets("Request DB").Select
    Worksheets("Request DB").Unprotect

    Req = Range("e2") + 1
'Determine emptyRow

    RowNo = Range("C2") + 4
    Rows(RowNo).Select
    RowNo = ActiveCell.Row
    ColNo = ActiveCell.Column
    Cells(RowNo, 1).Value = Req
    
   EditForm.Show 0
quit:

End Sub

Sub RoundedRectangle4_Click()
 If ActiveWorkbook.ReadOnly Then
         ActiveSheet.Shapes.Range(Array("Rounded Rectangle 4")).Visible = False
         ActiveSheet.Shapes.Range(Array("Rounded Rectangle 1")).Visible = False
         ActiveSheet.Shapes.Range(Array("Rounded Rectangle 2")).Visible = False
         Range("A2").Value = "File Checked out"
         
         Exit Sub
        End If

    Calculate
    
    ThisWorkbook.Activate
    Sheets("Request DB").Select
    Worksheets("Request DB").Unprotect
    maxno = Cells(2, 3).Value + 3
    RowNo = ActiveCell.Row
    ColNo = ActiveCell.Column
    If (RowNo < 4) Then
     MsgBox "Please Select an active row!"
     GoTo quit
     End If
         
    If (RowNo > maxno) Then
     MsgBox "Please Select an active row!"
     GoTo quit
     End If
    
    Call SetWindowSize1(1)
      'ActiveWindow.WindowState = xlMinimized
    EditForm.Show 0
quit:
 
 
 


End Sub

