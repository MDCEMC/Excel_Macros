VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EnterRequest 
   Caption         =   "Enter Request"
   ClientHeight    =   3744
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   7728
   OleObjectBlob   =   "EnterRequest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EnterRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

If (Val(Request) < 16000 Or Val(Request) > 21000) Then
    MsgBox ("Please enter Valid Request No >16000")
   GoTo Endd
End If
 Unload EditForm

ThisWorkbook.Activate
Sheets("Request DB").Visible = True
Sheets("Request DB").Select
Worksheets("Request DB").Unprotect
 Columns("A:A").Select
    On Error GoTo quit
    Selection.Find(What:=Request, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=True).Activate
   
  On Error GoTo 0
 ThisWorkbook.Activate
  Sheets("Request DB").Cells(1, 15) = Request
  
  
  GoTo ok
quit:
  MsgBox ("Request Not Found Please enter new number.")
  GoTo Endd
  
  
  
  
ok:
  EnterRequest.Hide
  Load EditForm
  EditForm.Show 0
  
  
Endd:

End Sub

Private Sub CommandButton2_Click()

    ThisWorkbook.Activate
    Sheets("Request DB").Visible = True
    Sheets("Request DB").Select
    Worksheets("Request DB").Unprotect
    Call AddNewRequestToMaster
    EnterRequest.Hide
    Sheets("Request DB").Select
    Sheets("Request DB").Visible = True
    
    EditForm.Show 0
    
    
End Sub




Private Sub Request_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

If (Request < 13000) Then
    MsgBox ("Please enter Request >13000")
    End
    
End If
ThisWorkbook.Activate
Sheets("Request DB").Visible = True
Sheets("Request DB").Select
Worksheets("Request DB").Unprotect
 Columns("A:A").Select
    On Error GoTo quit
    Selection.Find(What:=Request, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=True).Activate
   
  On Error GoTo 0
  GoTo ok
quit:
  MsgBox ("Request Not Found Please enter new number.")
  End
  
  
  
  
ok:
  EnterRequest.Hide
  EditForm.Show
End Sub

Private Sub Request_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call CommandButton1_Click
    End If
End Sub

Private Sub UserForm_Terminate()
Call SetWindowSize1(2)
End Sub
