Attribute VB_Name = "Module4"

Sub AddHyperlinkToData()
'
' Macro
'
 If ActiveWorkbook.ReadOnly Then
         ActiveSheet.Shapes.Range(Array("Rounded Rectangle 4")).Visible = False
         ActiveSheet.Shapes.Range(Array("Rounded Rectangle 1")).Visible = False
         ActiveSheet.Shapes.Range(Array("Rounded Rectangle 2")).Visible = False
         Range("A2").Value = "Checked out"
         Exit Sub
        End If

ThisWorkbook.Activate
Worksheets("Request DB").Unprotect
Application.Dialogs(xlDialogInsertHyperlink).Show
 ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
'
   
End Sub
Sub Date_right()
'
' Macro6 Macro
'

'
    A = ActiveCell.Column + 7
    
    Cells(1, A).Select
    Cells(1, A).Activate

End Sub
Sub Date_Left()
'
' Macro6 Macro
'

'
    A = ActiveCell.Column - 7
    If A < 0 Then A = 1
    
    
    Cells(1, A).Select
    Cells(1, A).Activate

End Sub


Sub hidebutton()
    ActiveSheet.Shapes.Range(Array("Rounded Rectangle 4")).Visible = True
    Range("A2").Value = "Checked out"
End Sub

Sub Workbook_Open()
    
     If ActiveWorkbook.ReadOnly Then
     
          ActiveSheet.Shapes.Range(Array("Rounded Rectangle 1")).Visible = False
          ActiveSheet.Shapes.Range(Array("Rounded Rectangle 2")).Visible = False
          ActiveSheet.Shapes.Range(Array("Rounded Rectangle 3")).Visible = False
          ActiveSheet.Shapes.Range(Array("Rounded Rectangle 4")).Visible = False
          Range("A2").Value = "Checked out"
          
          Else
          ActiveSheet.Shapes.Range(Array("Rounded Rectangle 1")).Visible = True
          ActiveSheet.Shapes.Range(Array("Rounded Rectangle 2")).Visible = True
          ActiveSheet.Shapes.Range(Array("Rounded Rectangle 3")).Visible = True
          ActiveSheet.Shapes.Range(Array("Rounded Rectangle 4")).Visible = True
          
          
    End If
    
    
    Range("L1").Select
    ActiveCell.FormulaR1C1 = " "
    Range("A4").Select
End Sub
Sub CloseBook()
    
If ActiveWorkbook.ReadOnly Then
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = True
    
End If

End Sub

