VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteTests 
   Caption         =   "UserForm1"
   ClientHeight    =   3180
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   4704
   OleObjectBlob   =   "DeleteTests.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DeleteTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
   ThisWorkbook.Activate
   Sheets("TestPlan DB").Visible = True
   Sheets("TestPlan DB").Select
   Worksheets("TestPlan DB").Unprotect
   ActiveSheet.Calculate
   'MsgBox "Please Wait"
   ID = RequestNo & PlanNo
   If PlanNo < 10 Then ID = RequestNo & "0" & PlanNo

            
    Columns("K:K").Select
    On Error GoTo quit
    aa = Selection.Find(What:=ID, After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=True).Activate
       
        RowNo = ActiveCell.Row
        Rows(ActiveCell.Row).EntireRow.Delete


quit:
          
  ActiveWorkbook.Worksheets("TestPlan DB").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("TestPlan DB").AutoFilter.Sort.SortFields.add Key:= _
        Range("K3"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("TestPlan DB").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
   mm = "Test Plan " & PlanNo & " from Request " & RequestNo & " was deleted"
   result = MsgBox(mm, vbInformation)
    


End Sub



Private Sub CommandButton2_Click()
DeleteTests.Hide
End Sub



Private Sub UserForm_Activate()
ThisWorkbook.Activate
Sheets("Request DB").Select
    Worksheets("Request DB").Unprotect
    ActiveSheet.Calculate
    RowNo = ActiveCell.Row
    ColNo = ActiveCell.Column
    If (RowNo < 4) Then GoTo quit
 
RequestNo = Cells(RowNo, 1).Value
ReqNo = Cells(RowNo, 1).Value
quit:
End Sub

Private Sub UserForm_Initialize()

ThisWorkbook.Activate
 Sheets("Request DB").Select
    Worksheets("Request DB").Unprotect
    ActiveSheet.Calculate
    RowNo = ActiveCell.Row
    ColNo = ActiveCell.Column
    If (RowNo < 4) Then GoTo quit
 
RequestNo = Cells(RowNo, 1).Value
ReqNo = Cells(RowNo, 1).Value

quit:

End Sub
