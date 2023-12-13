Attribute VB_Name = "Module2"
Sub RoundedRectangle1_Click()

        If ActiveWorkbook.ReadOnly Then
         ActiveSheet.Shapes.Range(Array("Rounded Rectangle 4")).Visible = False
         ActiveSheet.Shapes.Range(Array("Rounded Rectangle 1")).Visible = False
         ActiveSheet.Shapes.Range(Array("Rounded Rectangle 2")).Visible = False
         Range("A2").Value = "Checked out"
         Exit Sub
        End If
        Call SetWindowSize1(1)
        EnterRequest.Show 0
End Sub
