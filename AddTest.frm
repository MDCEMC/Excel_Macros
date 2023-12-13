VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddTest 
   Caption         =   "Add test"
   ClientHeight    =   5520
   ClientLeft      =   0
   ClientTop       =   180
   ClientWidth     =   4680
   OleObjectBlob   =   "AddTest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CommandButton1_Click()
 Dim lItem As Long
 Dim lRow As Long
 
 
 Application.ScreenUpdating = False
 
   Sheets("TestPlan DB").Visible = True
   Sheets("TestPlan DB").Select
    Worksheets("TestPlan DB").Unprotect
    ActiveSheet.Calculate
    NoTests = Sheets("TestPlan DB").Range("L2")
    NextTP = NoTests + 1
    Req = Sheets("TestPlan DB").Range("I2")
    
    ActiveWorkbook.Worksheets("TestPlan DB").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("TestPlan DB").AutoFilter.Sort.SortFields.add Key:= _
        Range("A3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("TestPlan DB").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
'Determine emptyRow
 emptyRow = Range("F2") + 4
 
  For lItem = 0 To Me.TestList.ListCount - 1
    If Me.TestList.Selected(lItem) Then
           idd = Req * 100 + Left(Me.TestList.List(lItem), 2)
   
   
            lRow = 0 ' reset Irow value
            Sheets("TestPlan DB").Select
            On Error Resume Next
            lRow = Application.WorksheetFunction.Match(CStr(idd), Range("I:I"), 0)
            On Error GoTo 0
            'MsgBox (idd & lRow)
            
            If lRow > 0 Then
                 mmsg = "Test " & Left(Me.TestList.List(lItem), 2) & " Already Selected """
                 MsgBox (mmsg)
                 GoTo noadd
                 
            End If
             
       
add:
                Cells(emptyRow - 1, 1).Select
                 Application.CutCopyMode = False
                 Selection.Copy
                Cells(emptyRow, 1).Select
                 ActiveSheet.Paste
                Cells(emptyRow - 1, 3).Select
                 Application.CutCopyMode = False
                 Selection.Copy
                Cells(emptyRow, 3).Select
                 ActiveSheet.Paste
                 Cells(emptyRow - 1, 8).Select
                 Application.CutCopyMode = False
                 Selection.Copy
                Cells(emptyRow, 8).Select
                 ActiveSheet.Paste
                 
                 
                 Cells(emptyRow, 1).Value = Req
                 ff = Int(Left(Me.TestList.List(lItem), 2))
                 cc = 2
                 If ff > 9 Then cc = 3
                 BB = Len(Me.TestList.List(lItem))
                 Cells(emptyRow, 2).Value = Right(Me.TestList.List(lItem), BB - cc)
                 
                 Cells(emptyRow, 3).Value = NextTP
                 Cells(emptyRow, 4).Value = Int(Left(Me.TestList.List(lItem), 2))
                 emptyRow = emptyRow + 1
                 NextTP = NextTP + 1
noadd:
        End If
    Next lItem
    ActiveSheet.Calculate
      ActiveWorkbook.Worksheets("TestPlan DB").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("TestPlan DB").AutoFilter.Sort.SortFields.add Key:= _
        Range("K3:K4521"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("TestPlan DB").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
  
  ActiveSheet.Calculate
   
      ActiveWorkbook.Worksheets("TestPlan DB").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("TestPlan DB").AutoFilter.Sort.SortFields.add Key:= _
        Range("K3:K4521"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
        :=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("TestPlan DB").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
  
      
    

 ActiveSheet.Calculate
 

 
AddTest.Hide
Load EditForm
EditForm.Show 0
    
Application.ScreenUpdating = True






'Call UserForm_Initialize
    


End Sub

Private Sub UserForm_Activate()
TestTot = Sheets("Editor").Range("J5")
With TestList
     .RowSource = "=Editor!T1:T" & TestTot
 
End With
End Sub

Private Sub UserForm_Initialize()

TestTot = Sheets("Editor").Range("J5")
With TestList
     .RowSource = "=Editor!T1:T" & TestTot
 
End With

End Sub
