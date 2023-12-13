VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EditForm 
   Caption         =   "Request and Test List"
   ClientHeight    =   10788
   ClientLeft      =   -24
   ClientTop       =   84
   ClientWidth     =   13884
   OleObjectBlob   =   "EditForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EditForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EditForm_Activate()
    With Me
        'This will create a vertical scrollbar
       ' .ScrollBars = fmScrollBarsVertical
        
        'Change the values of 2 as Per your requirements
        '.ScrollHeight = .InsideHeight * 3
        '.ScrollWidth = .InsideWidth * 9
    End With
    Wait.Visible = False
End Sub


   
End Sub


Private Sub CommandButton3_Click()
DeleteTests.Show 0
Call UserForm_Initialize

End Sub
Private Sub CommandButton4_Click()
  ' Unload EditForm
   Call AddRequest_Click  ' save information from form
  
   Sheets("Request DB").Visible = True
   Sheets("Request DB").Select
   Sheets("Request DB").Unprotect
   Call AddNewRequestToMaster
   Sheets("Request DB").Select
  ' Sheets("Request DB").Visible = False
  ' Sheets("Request DB").Protect
'   EditForm.Show
 '  Load EditForm
   Call UserForm_Initialize
    
End Sub

Private Sub CommandButton6_Click()
  EditForm.Hide

  EnterRequest.Show 0
End Sub

Private Sub CommandButton5_Click()
    
    Call PrintTabSetup
    Sheets("Mechanic Check In-Out").Visible = True
    Sheets("Tests").Visible = True
    
     Sheets(Array("Mechanic Check In-Out", "Tests")).Select
 
    Sheets("Mechanic Check In-Out").Activate
    
     On Error Resume Next
  
    ActiveWindow.SelectedSheets.PrintOut from:=1, To:=2, Copies:=1, Collate _
        :=True, IgnorePrintAreas:=False
  On Error GoTo 0
   
    Sheets("Mechanic Check In-Out").Visible = False
    Sheets("Tests").Visible = False
    Sheets("Request DB").Select
    
End Sub


Private Sub CommandButton7_Click()
  Dim emptyRow As Long
  Call AddRequest_Click  ' make sure anything added to request is updated too


'Make Sheet1 active
    Sheets("Request DB").Select
    Worksheets("Request DB").Unprotect
    RowNo = ActiveCell.Row
    ColNo = ActiveCell.Column
    If (RowNo < 4) Then GoTo quit



'Transfer information
 
 Cells(RowNo, 3).Value = ReqDate.Value
 Cells(RowNo, 4).Value = ComplDate.Value
 Cells(RowNo, 5).Value = Requestor.Value
 Cells(RowNo, 6).Value = ReqPhone1.Value
 Cells(RowNo, 7).Value = MYData.Value
 Cells(RowNo, 8).Value = BodyData.Value
 Cells(RowNo, 9).Value = VIN.Value
 Cells(RowNo, 10).Value = BodyDesc1.Value
 Cells(RowNo, 11).Value = EngineData.Value
 Cells(RowNo, 12).Value = TransData.Value
 Cells(RowNo, 13).Value = Licence.Value
 Cells(RowNo, 14).Value = Comments.Value
 Cells(RowNo, 15).Value = Reason.Value
 Cells(RowNo, 16).Value = VehNoData.Value
 Cells(RowNo, 17).Value = Color.Value
 Cells(RowNo, 18).Value = Miles.Value
 Cells(RowNo, 19).Value = Phase.Value
 Cells(RowNo, 20).Value = Overnight.Value
 Cells(RowNo, 22).Value = Drive.Value
 Cells(RowNo, 23).Value = VehCommentsData.Value
 
 
quit:

Sheets("Request DB").Select
 ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
 
 Sheets("TestPlan DB").Visible = True
 Sheets("TestPlan DB").Select
 ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
 
Sheets("TestPlan DB").Visible = False

    Call UserForm_Initialize
End Sub

Private Sub ComplDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
ComplDate = Date

End Sub

Private Sub Label10_Click()

End Sub

Private Sub PrintMFE_Click()
Call PrintMFESheet
 
End Sub

Private Sub PrintRNE_Click()
Call PrintRNESheet
End Sub

 Private Sub Requestor_click()
        Dim Phone(0 To 9) As String
        II = Requestor.ListIndex
        For I = 1 To 9
           If Requestor = Worksheets("Editor").Cells(I + 1, 13) Then II = I
           Phone(I) = Worksheets("Editor").Cells(I + 1, 14)
        Next I
        ReqPhone1 = Phone(II)
End Sub







Private Sub VIN_Change()
 
VIN.BackColor = RGB(255, 0, 0)

If Len(VIN) = 17 Or Len(VIN) = 8 Then
    VIN.BackColor = RGB(255, 255, 255)
End If
 
End Sub

Private Sub ZoomIn_Click()
EditForm.Zoom = (EditForm.Zoom - 5)
End Sub
Private Sub ZoomOut_Click()
EditForm.Zoom = (EditForm.Zoom + 5)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
     If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use the exit button on the form.", vbCritical
    End If

    
    
   ' If CloseMode = 0 Then Cancel = True
End Sub



Private Sub CommandButton1_Click()
Dim emptyRow As Long

'Make Sheet1 active
    Sheets("Request DB").Select
    Worksheets("Request DB").Unprotect
    RowNo = ActiveCell.Row
    ColNo = ActiveCell.Column
    If (RowNo < 4) Then GoTo quit



'Transfer information
 
 Cells(RowNo, 3).Value = ReqDate.Value
 Cells(RowNo, 4).Value = ComplDate.Value
 Cells(RowNo, 5).Value = Requestor.Value
 Cells(RowNo, 6).Value = ReqPhone1.Value
 Cells(RowNo, 7).Value = MYData.Value
 Cells(RowNo, 8).Value = BodyData.Value
 Cells(RowNo, 9).Value = VIN.Value
 Cells(RowNo, 10).Value = BodyDesc1.Value
 Cells(RowNo, 11).Value = EngineData.Value
 Cells(RowNo, 12).Value = TransData.Value
 Cells(RowNo, 13).Value = Licence.Value
 Cells(RowNo, 14).Value = Comments.Value
 Cells(RowNo, 15).Value = Reason.Value
 Cells(RowNo, 16).Value = VehNoData.Value
 Cells(RowNo, 17).Value = Color.Value
 Cells(RowNo, 18).Value = Miles.Value
 Cells(RowNo, 19).Value = Phase.Value
 Cells(RowNo, 20).Value = Overnight.Value
 Cells(RowNo, 22).Value = Drive.Value
 Cells(RowNo, 23).Value = VehCommentsData.Value
 
 
quit:

Sheets("Request DB").Select
 ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
 
 Sheets("TestPlan DB").Visible = True
 Sheets("TestPlan DB").Select
 ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
 
Sheets("TestPlan DB").Visible = False


    AddTest.Show 0 ' modeless
    Call UserForm_Initialize


End Sub

Private Sub CommandButton2_Click()


   result = MsgBox("Do you want to save your Request/Plan changes?", vbYesNoCancel)
   If result = vbNo Then
      aa = 1
   End If
 
 If result = vbYes Then
   Wait.Visible = True
   Call AddRequest_Click
 End If
 
 If result = vbCancel Then
  GoTo quitt
 End If
   
ThisWorkbook.Activate
Sheets("Request DB").Select
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
 
On Error GoTo quit
    ThisWorkbook.Activate
    Sheets("TestPlan DB").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
    Sheets("TestPlan DB").Visible = False

quit:
Unload Me
Unload AddTest
Unload DeleteTests
Unload EnterRequest

quitt:
Call SetWindowSize1(2)
 
Sheets("Request DB").Select

    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
End Sub
Private Sub H1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  H1 = Date
End Sub


Private Sub H2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
H2 = Date
End Sub

Private Sub H3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
H3 = Date

End Sub


Private Sub H4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
H4 = Date
End Sub

Private Sub H5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
H5 = Date

End Sub


Private Sub H6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
H6 = Date
End Sub

Private Sub H7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
H7 = Date

End Sub


Private Sub H8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
D8 = Date
End Sub

Private Sub H9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
H9 = Date
End Sub

Private Sub H10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
H10 = Date
End Sub

Private Sub Days1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Days1 = Date

End Sub


Private Sub Days2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Days2 = Date
End Sub

Private Sub Days3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Days3 = Date

End Sub


Private Sub Days4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Days4 = Date
End Sub

Private Sub Days5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Days5 = Date

End Sub


Private Sub Days6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Days6 = Date
End Sub

Private Sub Days7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Days7 = Date

End Sub


Private Sub Days8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Days8 = Date
End Sub

Private Sub Days9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Days9 = Date
End Sub

Private Sub Days10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Days10 = Date
End Sub





Private Sub D1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
D1 = Date

End Sub


Private Sub D2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
D2 = Date
End Sub

Private Sub D3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
D3 = Date

End Sub


Private Sub D4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
D4 = Date
End Sub

Private Sub D5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
D5 = Date

End Sub


Private Sub D6_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
D6 = Date
End Sub

Private Sub D7_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
D7 = Date

End Sub


Private Sub D8_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
D8 = Date
End Sub

Private Sub D9_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
D9 = Date
End Sub

Private Sub D10_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
D10 = Date
End Sub

Private Sub ReqDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
ReqDate = Date
End Sub



Private Sub UserForm_Activate()
   'Name of the frame
   With Me.Frame1
        'This will create a vertical scrollbar
        '.ScrollBars = fmScrollBarsVertical
        
        'Change the values of 2 as Per your requirements
        '.ScrollHeight = .InsideHeight * 2
        '.ScrollWidth = .InsideWidth * 9
    End With
 
End Sub

Private Sub AddRequest_Click()
Dim emptyRow As Long


'Make Sheet1 active
ThisWorkbook.Activate
 Sheets("Request DB").Select
    Worksheets("Request DB").Unprotect
    RowNo = ActiveCell.Row
    ColNo = ActiveCell.Column
    ReqNo = Cells(RowNo, 1)
  
    
    If (RowNo < 4) Then GoTo quit

Wait.Visible = True

'Transfer information to request tab
 
 Cells(RowNo, 3).Value = ReqDate.Value
 Cells(RowNo, 4).Value = ComplDate.Value
 Cells(RowNo, 5).Value = Requestor.Value
 Cells(RowNo, 6).Value = ReqPhone1.Value
 Cells(RowNo, 7).Value = MYData.Value
 Cells(RowNo, 8).Value = BodyData.Value
 Cells(RowNo, 9).Value = VIN.Value
 Cells(RowNo, 10).Value = BodyDesc1.Value
 Cells(RowNo, 11).Value = EngineData.Value
 Cells(RowNo, 12).Value = TransData.Value
 Cells(RowNo, 13).Value = Licence.Value
 Cells(RowNo, 14).Value = Comments.Value
 Cells(RowNo, 15).Value = Reason.Value
 Cells(RowNo, 16).Value = VehNoData.Value
 Cells(RowNo, 17).Value = Color.Value
 Cells(RowNo, 18).Value = Miles.Value
 Cells(RowNo, 19).Value = Phase.Value
 Cells(RowNo, 20).Value = Overnight.Value
 Cells(RowNo, 22).Value = Drive.Value
 Cells(RowNo, 23).Value = VehCommentsData.Value
 

 
 
 'Transfer information to test plan tab
 


ThisWorkbook.Activate
 Sheets("TestPlan DB").Visible = True
 Sheets("TestPlan DB").Select
    Worksheets("TestPlan DB").Unprotect
    Sheets("TestPlan DB").Range("I2") = ReqNo
    Req = ReqNo
  
    Columns("A:A").Select
    On Error GoTo quit
    Selection.Find(What:=ReqNo, After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=True).Activate

  On Error GoTo 0
   
 'Plan 1
   
    RowNo = ActiveCell.Row
    
    
    If (D1 <> "") Then Cells(RowNo, 5).Value = FormatDateTime(D1) Else Cells(RowNo, 5).Value = D1
    If (H1 <> "") Then Cells(RowNo, 6).Value = FormatDateTime(H1) Else Cells(RowNo, 6).Value = H1
    If (Days1 <> "") Then Cells(RowNo, 7).Value = FormatDateTime(Days1) Else Cells(RowNo, 7).Value = Days1
    'Cells(RowNo, 7).Value = Val(Days1)
  
  'Plan 2
     
    RowNo = RowNo + 1
    req1 = Cells(RowNo, 1).Value
    If (Req <> req1) Then GoTo quit
        
        If (D2 <> "") Then Cells(RowNo, 5).Value = FormatDateTime(D2) Else Cells(RowNo, 5).Value = D2
    If (H2 <> "") Then Cells(RowNo, 6).Value = FormatDateTime(H2) Else Cells(RowNo, 6).Value = H2
    If (Days2 <> "") Then Cells(RowNo, 7).Value = FormatDateTime(Days2) Else Cells(RowNo, 7).Value = Days2
    
        
    'Plan 3
     
    RowNo = RowNo + 1
    req1 = Cells(RowNo, 1).Value
    If (Req <> req1) Then GoTo quit
         
          
        If (D3 <> "") Then Cells(RowNo, 5).Value = FormatDateTime(D3) Else Cells(RowNo, 5).Value = D3
    If (H3 <> "") Then Cells(RowNo, 6).Value = FormatDateTime(H3) Else Cells(RowNo, 6).Value = H3
    If (Days3 <> "") Then Cells(RowNo, 7).Value = FormatDateTime(Days3) Else Cells(RowNo, 7).Value = Days3
    
        
  'Plan 4
     
    RowNo = RowNo + 1
    req1 = Cells(RowNo, 1).Value
    If (Req <> req1) Then GoTo quit
        
          
        If (D4 <> "") Then Cells(RowNo, 5).Value = FormatDateTime(D4) Else Cells(RowNo, 5).Value = D4
    If (H4 <> "") Then Cells(RowNo, 6).Value = FormatDateTime(H4) Else Cells(RowNo, 6).Value = H4
    If (Days4 <> "") Then Cells(RowNo, 7).Value = FormatDateTime(Days4) Else Cells(RowNo, 7).Value = Days4
    
        
  'Plan 5
     
    RowNo = RowNo + 1
    req1 = Cells(RowNo, 1).Value
    If (Req <> req1) Then GoTo quit
        
          
        If (D5 <> "") Then Cells(RowNo, 5).Value = FormatDateTime(D5) Else Cells(RowNo, 5).Value = D5
        If (H5 <> "") Then Cells(RowNo, 6).Value = FormatDateTime(H5) Else Cells(RowNo, 6).Value = H5
        If (Days5 <> "") Then Cells(RowNo, 7).Value = FormatDateTime(Days5) Else Cells(RowNo, 7).Value = Days5
    

  'Plan 6
     
    RowNo = RowNo + 1
    req1 = Cells(RowNo, 1).Value
    If (Req <> req1) Then GoTo quit
        
        If (D6 <> "") Then Cells(RowNo, 5).Value = FormatDateTime(D6) Else Cells(RowNo, 5).Value = D6
        If (H6 <> "") Then Cells(RowNo, 6).Value = FormatDateTime(H6) Else Cells(RowNo, 6).Value = H6
        If (Days6 <> "") Then Cells(RowNo, 7).Value = FormatDateTime(Days6) Else Cells(RowNo, 7).Value = Days6
    
  
 'Plan 7
     
    RowNo = RowNo + 1
    req1 = Cells(RowNo, 1).Value
    If (Req <> req1) Then GoTo quit
       
        If (D7 <> "") Then Cells(RowNo, 5).Value = FormatDateTime(D7) Else Cells(RowNo, 5).Value = D7
        If (H7 <> "") Then Cells(RowNo, 6).Value = FormatDateTime(H7) Else Cells(RowNo, 6).Value = H7
        If (Days7 <> "") Then Cells(RowNo, 7).Value = FormatDateTime(Days7) Else Cells(RowNo, 7).Value = Days7
    
 
 'Plan 8
     
    RowNo = RowNo + 1
    req1 = Cells(RowNo, 1).Value
    If (Req <> req1) Then GoTo quit
        
        If (D8 <> "") Then Cells(RowNo, 5).Value = FormatDateTime(D8) Else Cells(RowNo, 5).Value = D8
        If (H8 <> "") Then Cells(RowNo, 6).Value = FormatDateTime(H8) Else Cells(RowNo, 6).Value = H8
        If (Days8 <> "") Then Cells(RowNo, 7).Value = FormatDateTime(Days8) Else Cells(RowNo, 7).Value = Days8
    
  
 'Plan 9
     
    RowNo = RowNo + 1
    req1 = Cells(RowNo, 1).Value
    If (Req <> req1) Then GoTo quit
         
        If (D9 <> "") Then Cells(RowNo, 5).Value = FormatDateTime(D9) Else Cells(RowNo, 5).Value = D9
        If (H9 <> "") Then Cells(RowNo, 6).Value = FormatDateTime(H9) Else Cells(RowNo, 6).Value = H9
        If (Days9 <> "") Then Cells(RowNo, 7).Value = FormatDateTime(Days9) Else Cells(RowNo, 7).Value = Days9
    
 
 
   'Plan 10
     
    RowNo = RowNo + 1
    req1 = Cells(RowNo, 1).Value
    If (Req <> req1) Then GoTo quit
         
        If (D10 <> "") Then Cells(RowNo, 5).Value = FormatDateTime(D10) Else Cells(RowNo, 5).Value = D10
        If (H10 <> "") Then Cells(RowNo, 6).Value = FormatDateTime(H10) Else Cells(RowNo, 6).Value = H10
        If (Days10 <> "") Then Cells(RowNo, 7).Value = FormatDateTime(Days10) Else Cells(RowNo, 7).Value = Days10
    
 
 
 
quit:


Sheets("Request DB").Select
 ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
 
 Sheets("TestPlan DB").Select
 ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True

Sheets("TestPlan DB").Visible = False
Sheets("Request DB").Select


Sheets("Request DB").Select
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
Wait.Visible = False


End Sub


Private Sub ScrollBar1_Change()

    Sheets("Request DB").Select
    RowNo = ActiveCell.Row - 1
    ColNo = ActiveCell.Column
   
    If (RowNo < 4) Then GoTo quit
    Rows(RowNo).Select
    Call UserForm_Initialize
    
quit:
 Sheets("Request DB").Select
End Sub


Private Sub SpinButton1_SpinDown()
Sheets("Request DB").Select
    RowNo = ActiveCell.Row + 1
    ColNo = ActiveCell.Column
    maxno = Cells(2, 3).Value + 3
   Call AddRequest_Click  ' save information from form
    If (RowNo < 4 Or RowNo > maxno) Then GoTo quit
    Rows(RowNo).Select
    Call UserForm_Initialize
quit:
Sheets("Request DB").Select
End Sub

Private Sub SpinButton1_SpinUp()
ThisWorkbook.Activate
Call AddRequest_Click  ' save information from form
Sheets("Request DB").Select
    RowNo = ActiveCell.Row - 1
    ColNo = ActiveCell.Column
   
    If (RowNo < 4) Then GoTo quit
    Rows(RowNo).Select
    Call UserForm_Initialize
    
quit:
 
 'Sheets("Request_DB").Select
End Sub

Public Sub UserForm_Initialize()


'EditForm.ScrollBars = fmScrollBarsVertical
'EditForm.KeepScrollBarsVisible = fmScrollBarsNone

'EditForm.Height = EditForm.Height / 1
'EditForm.ScrollHeight = 2 * EditForm.Height
'EditForm.ScrollWidth = 2 * EditForm.Width
 



ThisWorkbook.Activate
 Sheets("Request DB").Select
    Worksheets("Request DB").Unprotect
    RowNo = ActiveCell.Row
    ColNo = ActiveCell.Column
    If (RowNo < 4) Then GoTo quit
    

    
'TL = Cells(RowNo, 2).Value
'RR = Cells(RowNo, 6).Value
'RE = Cells(RowNo, 12).Value
RequestNo = Cells(RowNo, 1).Value
ReqNo = Cells(RowNo, 1).Value
RequestNo = "V" & Cells(RowNo, 1).Value
Cells(1, 15).Value = ReqNo

 ReqDate = Cells(RowNo, 3).Value
 ComplDate = Cells(RowNo, 4).Value
 VIN.Value = Cells(RowNo, 9).Value
 
 ReqPhone1.Value = Cells(RowNo, 6).Value
 VehNoData.Value = Cells(RowNo, 16).Value
 Miles.Value = Cells(RowNo, 18).Value
 Licence.Value = Cells(RowNo, 13).Value
 BodyDesc1.Value = Cells(RowNo, 10).Value
 
 Comments.Value = Cells(RowNo, 14).Value
 VehCommentsData.Value = Cells(RowNo, 23).Value
 'GoTo tt
  Call RemoveDuplicatesTL1
  MYData.Value = Cells(RowNo, 7).Value
  Call RemoveDuplicatesBody1
  BodyData.Value = Cells(RowNo, 8).Value
  Call RemoveDuplicatesRequestor
  Requestor.Value = Cells(RowNo, 5).Value
  Call RemoveDuplicatesTrans
  TransData.Value = Cells(RowNo, 12).Value
  Call RemoveDuplicatesEngine
  EngineData.Value = Cells(RowNo, 11).Value
  Call RemoveDuplicatesReason
  Reason.Value = Cells(RowNo, 15).Value
  Call RemoveDuplicatesPhase
  Phase.Value = Cells(RowNo, 19).Value
  Call RemoveDuplicatesOvernight
  Overnight.Value = Cells(RowNo, 20).Value
  Call RemoveDuplicatesDrive
  Drive.Value = Cells(RowNo, 22).Value
  Call RemoveDuplicatesColor
  Color.Value = Cells(RowNo, 17).Value
tt:
'testplan data
   T1 = "": T2 = "": T3 = "": T4 = "": T5 = "": T6 = "": T7 = "": T8 = ""
   D1 = "": D2 = "": D3 = "": D4 = "": D5 = "": D6 = "": D7 = "": D8 = ""
   H1 = "": H2 = "": H3 = "": H4 = "": H5 = "": H6 = "": H7 = "": H8 = ""
   P1 = "": P2 = "": P3 = "": P4 = "": P5 = "": P6 = "": P7 = "": P8 = ""
   n1 = "": n2 = "": n3 = "": n4 = "": n5 = "": n6 = "": n7 = "": n8 = ""
   T9 = "": T10 = ""
   D9 = "": D10 = ""
   H9 = "": H10 = ""
   P9 = "": P10 = ""
   n9 = "": n10 = ""
   Days1 = "": Days2 = "": Days3 = "": Days4 = "": Days5 = ""
   Days6 = "": Days7 = "": Days8 = "": Days9 = "": Days10 = ""
   
   
   


T1.Visible = False: D1.Visible = False: H1.Visible = False: P1.Visible = False: n1.Visible = False
T2.Visible = False: D2.Visible = False: H2.Visible = False: P2.Visible = False: n2.Visible = False
T3.Visible = False: D3.Visible = False: H3.Visible = False: P3.Visible = False: n3.Visible = False
T4.Visible = False: D4.Visible = False: H4.Visible = False: P4.Visible = False: n4.Visible = False
T5.Visible = False: D5.Visible = False: H5.Visible = False: P5.Visible = False: n5.Visible = False
T6.Visible = False: D6.Visible = False: H6.Visible = False: P6.Visible = False: n6.Visible = False
T7.Visible = False: D7.Visible = False: H7.Visible = False: P7.Visible = False: n7.Visible = False
T8.Visible = False: D8.Visible = False: H8.Visible = False: P8.Visible = False: n8.Visible = False
T9.Visible = False: D9.Visible = False: H9.Visible = False: P9.Visible = False: n9.Visible = False
T10.Visible = False: D10.Visible = False: H10.Visible = False: P10.Visible = False: n10.Visible = False



Days1.Visible = False: Days2.Visible = False: Days3.Visible = False: Days4.Visible = False: Days5.Visible = False
Days6.Visible = False: Days7.Visible = False: Days8.Visible = False: Days9.Visible = False: Days10.Visible = False



NoPlans.Visible = True
Wait.Visible = False



ThisWorkbook.Activate
Sheets("TestPlan DB").Visible = True
Sheets("TestPlan DB").Select
    Worksheets("TestPlan DB").Unprotect
    Sheets("TestPlan DB").Range("I2") = ReqNo
    
    
    Columns("A:A").Select
    On Error GoTo quit
    Selection.Find(What:=ReqNo, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=True).Activate
   
  On Error GoTo 0
  
 NoPlans.Visible = False
   
   
   
 'Plan 1
   
    RowNo = ActiveCell.Row
    Req = Cells(RowNo, 1).Value
    dd = Cells(RowNo, 3).Value
    aa = Cells(RowNo, 4).Value
    BB = Cells(RowNo, 5).Value
    cc = Cells(RowNo, 6).Value
    ee = Cells(RowNo, 2).Value
    ff = Cells(RowNo, 7).Value
    gg = Cells(RowNo, 12).Value
  
    T1.Visible = False: D1.Visible = True: H1.Visible = True: P1.Visible = True: n1.Visible = True
    T1 = aa: D1 = BB: H1 = cc: P1 = dd: n1 = ee
   
        Days1.Visible = True: Days1 = ff
        
  'Plan 2
    
    
    RowNo = RowNo + 1
    req1 = Cells(RowNo, 1).Value
    If (Req <> req1) Then GoTo quit
    
    Req = Cells(RowNo, 1).Value
    dd = Cells(RowNo, 3).Value
    aa = Cells(RowNo, 4).Value
    BB = Cells(RowNo, 5).Value
    cc = Cells(RowNo, 6).Value
    ee = Cells(RowNo, 2).Value
    ff = Cells(RowNo, 7).Value
    gg = Cells(RowNo, 12).Value
    
     T2.Visible = False: D2.Visible = True: H2.Visible = True: P2.Visible = True: n2.Visible = True
     Days2.Visible = True: Days2 = ff
     T2 = aa: D2 = BB: H2 = cc: P2 = dd: n2 = ee
    
  'Plan 3
    
    
    RowNo = RowNo + 1
    req1 = Cells(RowNo, 1).Value
    If (Req <> req1) Then GoTo quit
    Req = Cells(RowNo, 1).Value
    dd = Cells(RowNo, 3).Value
    aa = Cells(RowNo, 4).Value
    BB = Cells(RowNo, 5).Value
    cc = Cells(RowNo, 6).Value
    ee = Cells(RowNo, 2).Value
    ff = Cells(RowNo, 7).Value
    gg = Cells(RowNo, 12).Value
    
          
        T3.Visible = False: D3.Visible = True: H3.Visible = True: P3.Visible = True: n3.Visible = True
        Days3.Visible = True: Days3 = ff
        T3 = aa: D3 = BB: H3 = cc: P3 = dd: n3 = ee: day3 = ff
    
 'Plan 4
    
    
    RowNo = RowNo + 1
    req1 = Cells(RowNo, 1).Value
    If (Req <> req1) Then GoTo quit
    Req = Cells(RowNo, 1).Value
    dd = Cells(RowNo, 3).Value
    aa = Cells(RowNo, 4).Value
    BB = Cells(RowNo, 5).Value
    cc = Cells(RowNo, 6).Value
    ee = Cells(RowNo, 2).Value
    ff = Cells(RowNo, 7).Value
    gg = Cells(RowNo, 12).Value
    
          
        T4.Visible = False: D4.Visible = True: H4.Visible = True: P4.Visible = True: n4.Visible = True
        Days4.Visible = True: Days4 = ff
        T4 = aa: D4 = BB: H4 = cc: P4 = dd: n4 = ee
    
        
    'Plan 5
    
    
    RowNo = RowNo + 1
    req1 = Cells(RowNo, 1).Value
    If (Req <> req1) Then GoTo quit
    Req = Cells(RowNo, 1).Value
    dd = Cells(RowNo, 3).Value
    aa = Cells(RowNo, 4).Value
    BB = Cells(RowNo, 5).Value
    cc = Cells(RowNo, 6).Value
    ee = Cells(RowNo, 2).Value
    ff = Cells(RowNo, 7).Value
    gg = Cells(RowNo, 12).Value
    
          
        T5.Visible = False: D5.Visible = True: H5.Visible = True: P5.Visible = True: n5.Visible = True
        Days5.Visible = True: Days5 = ff
       T5 = aa: D5 = BB: H5 = cc: P5 = dd: n5 = ee
    
       
       
 'Plan 6
    
    
    RowNo = RowNo + 1
    req1 = Cells(RowNo, 1).Value
   If (Req <> req1) Then GoTo quit
    Req = Cells(RowNo, 1).Value
    dd = Cells(RowNo, 3).Value
    aa = Cells(RowNo, 4).Value
    BB = Cells(RowNo, 5).Value
    cc = Cells(RowNo, 6).Value
    ee = Cells(RowNo, 2).Value
    ff = Cells(RowNo, 7).Value
    gg = Cells(RowNo, 12).Value
    
          
        T6.Visible = False: D6.Visible = True: H6.Visible = True: P6.Visible = True: n6.Visible = True
        Days6.Visible = True: Days6 = ff
        T6 = aa: D6 = BB: H6 = cc: P6 = dd: n6 = ee
    
 'Plan 7
   
    RowNo = RowNo + 1
    req1 = Cells(RowNo, 1).Value
    If (Req <> req1) Then GoTo quit
    Req = Cells(RowNo, 1).Value
    dd = Cells(RowNo, 3).Value
    aa = Cells(RowNo, 4).Value
    BB = Cells(RowNo, 5).Value
    cc = Cells(RowNo, 6).Value
    ee = Cells(RowNo, 2).Value
    ff = Cells(RowNo, 7).Value
    gg = Cells(RowNo, 12).Value
    
          
        T7.Visible = False: D7.Visible = True: H7.Visible = True: P7.Visible = True: n7.Visible = True
        Days7.Visible = True: Days7 = ff
        T7 = aa: D7 = BB: H7 = cc: P7 = dd: n7 = ee
    
 'Plan 8
    
    
    RowNo = RowNo + 1
    req1 = Cells(RowNo, 1).Value
    If (Req <> req1) Then GoTo quit
    Req = Cells(RowNo, 1).Value
    dd = Cells(RowNo, 3).Value
    aa = Cells(RowNo, 4).Value
    BB = Cells(RowNo, 5).Value
    cc = Cells(RowNo, 6).Value
    ee = Cells(RowNo, 2).Value
    ff = Cells(RowNo, 7).Value
    gg = Cells(RowNo, 12).Value
    
          
        T8.Visible = False: D8.Visible = True: H8.Visible = True: P8.Visible = True: n8.Visible = True
        Days8.Visible = True: Days8 = ff
        T8 = aa: D8 = BB: H8 = cc: P8 = dd: n8 = ee
    
'Plan 9
    
    
    RowNo = RowNo + 1
    req1 = Cells(RowNo, 1).Value
    If (Req <> req1) Then GoTo quit
    Req = Cells(RowNo, 1).Value
    dd = Cells(RowNo, 3).Value
    aa = Cells(RowNo, 4).Value
    BB = Cells(RowNo, 5).Value
    cc = Cells(RowNo, 6).Value
    ee = Cells(RowNo, 2).Value
    ff = Cells(RowNo, 7).Value
    gg = Cells(RowNo, 12).Value
    
          
        T9.Visible = False: D9.Visible = True: H9.Visible = True: P9.Visible = True: n9.Visible = True
        Days9.Visible = True: Days9 = ff
        T9 = aa: D9 = BB: H9 = cc: P9 = dd: n9 = ee
    
   'Plan 10
    
    
    RowNo = RowNo + 1
    req1 = Cells(RowNo, 1).Value
    If (Req <> req1) Then GoTo quit
    Req = Cells(RowNo, 1).Value
    dd = Cells(RowNo, 3).Value
    aa = Cells(RowNo, 4).Value
    BB = Cells(RowNo, 5).Value
    cc = Cells(RowNo, 6).Value
    ee = Cells(RowNo, 2).Value
    ff = Cells(RowNo, 7).Value
     gg = Cells(RowNo, 12).Value
    
          
        T10.Visible = False: D10.Visible = True: H10.Visible = True: P10.Visible = True: n10.Visible = True
        Days10.Visible = True: Days10 = ff
        T10 = aa: D10 = BB: H10 = cc: P10 = dd: n10 = ee
   
  
                                                                               
                                                                                                                                                     
                                                                               
                                                                               
quit:
ThisWorkbook.Activate
 Sheets("Request DB").Select
End Sub
Sub RemoveDuplicatesTL1()
    Dim AllCells As Range, Cell As Range
    Dim NoDupes As New Collection
    Dim I As Integer, j As Integer
    Dim Swap1, Swap2, Item
    maxno = Cells(2, 3).Value + 3
'   The items are in F4:Fxxx
    Set AllCells = Range("G4:G" & maxno)
    

'   The next statement ignores the error caused
'   by attempting to add a duplicate key to the collection.
'   The duplicate is not added - which is just what we want!
    On Error Resume Next
    For Each Cell In AllCells
        NoDupes.add Cell.Value, CStr(Cell.Value)
'       Note: the 2nd argument (key) for the Add method must be a string
    Next Cell

'   Resume normal error handling
    On Error GoTo 0

'   Sort the collection (optional)
    For I = 1 To NoDupes.Count - 1
        For j = I + 1 To NoDupes.Count
            If NoDupes(I) > NoDupes(j) Then
                Swap1 = NoDupes(I)
                Swap2 = NoDupes(j)
                NoDupes.add Swap1, Before:=j
                NoDupes.add Swap2, Before:=I
                NoDupes.Remove I + 1
                NoDupes.Remove j + 1
            End If
        Next j
    Next I

'   Add the sorted, non-duplicated items to a ListBox
    MYData.Clear
    For Each Item In NoDupes
        EditForm.MYData.AddItem Item
    Next Item

End Sub
Sub RemoveDuplicatesBody1()
    Dim AllCells As Range, Cell As Range
    Dim NoDupes As New Collection
    Dim I As Integer, j As Integer
    Dim Swap1, Swap2, Item
    maxno = Cells(2, 3).Value + 3
'   The items are in F4:Fxxx
    Set AllCells = Range("H4:H" & maxno)
    

'   The next statement ignores the error caused
'   by attempting to add a duplicate key to the collection.
'   The duplicate is not added - which is just what we want!
    On Error Resume Next
    For Each Cell In AllCells
        NoDupes.add Cell.Value, CStr(Cell.Value)
'       Note: the 2nd argument (key) for the Add method must be a string
    Next Cell

'   Resume normal error handling
    On Error GoTo 0

'   Sort the collection (optional)
    For I = 1 To NoDupes.Count - 1
        For j = I + 1 To NoDupes.Count
            If NoDupes(I) > NoDupes(j) Then
                Swap1 = NoDupes(I)
                Swap2 = NoDupes(j)
                NoDupes.add Swap1, Before:=j
                NoDupes.add Swap2, Before:=I
                NoDupes.Remove I + 1
                NoDupes.Remove j + 1
            End If
        Next j
    Next I

'   Add the sorted, non-duplicated items to a ListBox
    BodyData.Clear
    For Each Item In NoDupes
        EditForm.BodyData.AddItem Item
    Next Item

End Sub

Sub RemoveDuplicatesRequestor()
    Dim AllCells As Range, Cell As Range
    Dim NoDupes As New Collection
    Dim I As Integer, j As Integer
    Dim Swap1, Swap2, Item
    maxno = Cells(2, 3).Value + 3
'   The items are in F4:Fxxx
    Set AllCells = Range("E4:E" & maxno)
    

'   The next statement ignores the error caused
'   by attempting to add a duplicate key to the collection.
'   The duplicate is not added - which is just what we want!
    On Error Resume Next
    For Each Cell In AllCells
        NoDupes.add Cell.Value, CStr(Cell.Value)
'       Note: the 2nd argument (key) for the Add method must be a string
    Next Cell

'   Resume normal error handling
    On Error GoTo 0

'   Sort the collection (optional)
    For I = 1 To NoDupes.Count - 1
        For j = I + 1 To NoDupes.Count
            If NoDupes(I) > NoDupes(j) Then
                Swap1 = NoDupes(I)
                Swap2 = NoDupes(j)
                NoDupes.add Swap1, Before:=j
                NoDupes.add Swap2, Before:=I
                NoDupes.Remove I + 1
                NoDupes.Remove j + 1
            End If
        Next j
    Next I

'   Add the sorted, non-duplicated items to a ListBox
    Requestor.Clear
    For Each Item In NoDupes
        EditForm.Requestor.AddItem Item
    Next Item

End Sub
Sub RemoveDuplicatesTrans()
    Dim AllCells As Range, Cell As Range
    Dim NoDupes As New Collection
    Dim I As Integer, j As Integer
    Dim Swap1, Swap2, Item
    maxno = Cells(2, 3).Value + 3
'   The items are in F4:Fxxx
    Set AllCells = Range("L4:L" & maxno)
    

'   The next statement ignores the error caused
'   by attempting to add a duplicate key to the collection.
'   The duplicate is not added - which is just what we want!
    On Error Resume Next
    For Each Cell In AllCells
        NoDupes.add Cell.Value, CStr(Cell.Value)
'       Note: the 2nd argument (key) for the Add method must be a string
    Next Cell

'   Resume normal error handling
    On Error GoTo 0

'   Sort the collection (optional)
    For I = 1 To NoDupes.Count - 1
        For j = I + 1 To NoDupes.Count
            If NoDupes(I) > NoDupes(j) Then
                Swap1 = NoDupes(I)
                Swap2 = NoDupes(j)
                NoDupes.add Swap1, Before:=j
                NoDupes.add Swap2, Before:=I
                NoDupes.Remove I + 1
                NoDupes.Remove j + 1
            End If
        Next j
    Next I

'   Add the sorted, non-duplicated items to a ListBox
    TransData.Clear
    For Each Item In NoDupes
        EditForm.TransData.AddItem Item
    Next Item

End Sub

Sub RemoveDuplicatesEngine()
    Dim AllCells As Range, Cell As Range
    Dim NoDupes As New Collection
    Dim I As Integer, j As Integer
    Dim Swap1, Swap2, Item
    maxno = Cells(2, 3).Value + 3
'   The items are in F4:Fxxx
    Set AllCells = Range("K4:K" & maxno)
    

'   The next statement ignores the error caused
'   by attempting to add a duplicate key to the collection.
'   The duplicate is not added - which is just what we want!
    On Error Resume Next
    For Each Cell In AllCells
        NoDupes.add Cell.Value, CStr(Cell.Value)
'       Note: the 2nd argument (key) for the Add method must be a string
    Next Cell

'   Resume normal error handling
    On Error GoTo 0

'   Sort the collection (optional)
    For I = 1 To NoDupes.Count - 1
        For j = I + 1 To NoDupes.Count
            If NoDupes(I) > NoDupes(j) Then
                Swap1 = NoDupes(I)
                Swap2 = NoDupes(j)
                NoDupes.add Swap1, Before:=j
                NoDupes.add Swap2, Before:=I
                NoDupes.Remove I + 1
                NoDupes.Remove j + 1
            End If
        Next j
    Next I

'   Add the sorted, non-duplicated items to a ListBox
   EngineData.Clear
    For Each Item In NoDupes
        EditForm.EngineData.AddItem Item
    Next Item

End Sub
Sub RemoveDuplicatesReason()
    Dim AllCells As Range, Cell As Range
    Dim NoDupes As New Collection
    Dim I As Integer, j As Integer
    Dim Swap1, Swap2, Item
    maxno = Cells(2, 3).Value + 3
'   The items are in F4:Fxxx
    Set AllCells = Range("O4:O" & maxno)
    

'   The next statement ignores the error caused
'   by attempting to add a duplicate key to the collection.
'   The duplicate is not added - which is just what we want!
    On Error Resume Next
    For Each Cell In AllCells
        NoDupes.add Cell.Value, CStr(Cell.Value)
'       Note: the 2nd argument (key) for the Add method must be a string
    Next Cell

'   Resume normal error handling
    On Error GoTo 0

'   Sort the collection (optional)
    For I = 1 To NoDupes.Count - 1
        For j = I + 1 To NoDupes.Count
            If NoDupes(I) > NoDupes(j) Then
                Swap1 = NoDupes(I)
                Swap2 = NoDupes(j)
                NoDupes.add Swap1, Before:=j
                NoDupes.add Swap2, Before:=I
                NoDupes.Remove I + 1
                NoDupes.Remove j + 1
            End If
        Next j
    Next I

'   Add the sorted, non-duplicated items to a ListBox
   Reason.Clear
    For Each Item In NoDupes
        EditForm.Reason.AddItem Item
    Next Item

End Sub
Sub RemoveDuplicatesPhase()
    Dim AllCells As Range, Cell As Range
    Dim NoDupes As New Collection
    Dim I As Integer, j As Integer
    Dim Swap1, Swap2, Item
    maxno = Cells(2, 3).Value + 3
'   The items are in F4:Fxxx
    Set AllCells = Range("S4:S" & maxno)
    

'   The next statement ignores the error caused
'   by attempting to add a duplicate key to the collection.
'   The duplicate is not added - which is just what we want!
    On Error Resume Next
    For Each Cell In AllCells
        NoDupes.add Cell.Value, CStr(Cell.Value)
'       Note: the 2nd argument (key) for the Add method must be a string
    Next Cell

'   Resume normal error handling
    On Error GoTo 0

'   Sort the collection (optional)
    For I = 1 To NoDupes.Count - 1
        For j = I + 1 To NoDupes.Count
            If NoDupes(I) > NoDupes(j) Then
                Swap1 = NoDupes(I)
                Swap2 = NoDupes(j)
                NoDupes.add Swap1, Before:=j
                NoDupes.add Swap2, Before:=I
                NoDupes.Remove I + 1
                NoDupes.Remove j + 1
            End If
        Next j
    Next I

'   Add the sorted, non-duplicated items to a ListBox
   Phase.Clear
    For Each Item In NoDupes
        EditForm.Phase.AddItem Item
    Next Item

End Sub

Sub RemoveDuplicatesOvernight()
    Dim AllCells As Range, Cell As Range
    Dim NoDupes As New Collection
    Dim I As Integer, j As Integer
    Dim Swap1, Swap2, Item
    maxno = Cells(2, 3).Value + 3
'   The items are in F4:Fxxx
    Set AllCells = Range("T4:T" & maxno)
    

'   The next statement ignores the error caused
'   by attempting to add a duplicate key to the collection.
'   The duplicate is not added - which is just what we want!
    On Error Resume Next
    For Each Cell In AllCells
        NoDupes.add Cell.Value, CStr(Cell.Value)
'       Note: the 2nd argument (key) for the Add method must be a string
    Next Cell

'   Resume normal error handling
    On Error GoTo 0

'   Sort the collection (optional)
    For I = 1 To NoDupes.Count - 1
        For j = I + 1 To NoDupes.Count
            If NoDupes(I) > NoDupes(j) Then
                Swap1 = NoDupes(I)
                Swap2 = NoDupes(j)
                NoDupes.add Swap1, Before:=j
                NoDupes.add Swap2, Before:=I
                NoDupes.Remove I + 1
                NoDupes.Remove j + 1
            End If
        Next j
    Next I

'   Add the sorted, non-duplicated items to a ListBox
   Overnight.Clear
    For Each Item In NoDupes
        EditForm.Overnight.AddItem Item
    Next Item

End Sub

Sub RemoveDuplicatesDrive()
    Dim AllCells As Range, Cell As Range
    Dim NoDupes As New Collection
    Dim I As Integer, j As Integer
    Dim Swap1, Swap2, Item
    maxno = Cells(2, 3).Value + 3
'   The items are in F4:Fxxx
    Set AllCells = Range("V4:V" & maxno)
    

'   The next statement ignores the error caused
'   by attempting to add a duplicate key to the collection.
'   The duplicate is not added - which is just what we want!
    On Error Resume Next
    For Each Cell In AllCells
        NoDupes.add Cell.Value, CStr(Cell.Value)
'       Note: the 2nd argument (key) for the Add method must be a string
    Next Cell

'   Resume normal error handling
    On Error GoTo 0

'   Sort the collection (optional)
    For I = 1 To NoDupes.Count - 1
        For j = I + 1 To NoDupes.Count
            If NoDupes(I) > NoDupes(j) Then
                Swap1 = NoDupes(I)
                Swap2 = NoDupes(j)
                NoDupes.add Swap1, Before:=j
                NoDupes.add Swap2, Before:=I
                NoDupes.Remove I + 1
                NoDupes.Remove j + 1
            End If
        Next j
    Next I

'   Add the sorted, non-duplicated items to a ListBox
   Drive.Clear
    For Each Item In NoDupes
        EditForm.Drive.AddItem Item
    Next Item

End Sub
Sub RemoveDuplicatesColor()
    Dim AllCells As Range, Cell As Range
    Dim NoDupes As New Collection
    Dim I As Integer, j As Integer
    Dim Swap1, Swap2, Item
    maxno = Cells(2, 3).Value + 3
'   The items are in F4:Fxxx
    Set AllCells = Range("Q4:Q" & maxno)
    

'   The next statement ignores the error caused
'   by attempting to add a duplicate key to the collection.
'   The duplicate is not added - which is just what we want!
    On Error Resume Next
    For Each Cell In AllCells
        NoDupes.add Cell.Value, CStr(Cell.Value)
'       Note: the 2nd argument (key) for the Add method must be a string
    Next Cell

'   Resume normal error handling
    On Error GoTo 0

'   Sort the collection (optional)
    For I = 1 To NoDupes.Count - 1
        For j = I + 1 To NoDupes.Count
            If NoDupes(I) > NoDupes(j) Then
                Swap1 = NoDupes(I)
                Swap2 = NoDupes(j)
                NoDupes.add Swap1, Before:=j
                NoDupes.add Swap2, Before:=I
                NoDupes.Remove I + 1
                NoDupes.Remove j + 1
            End If
        Next j
    Next I

'   Add the sorted, non-duplicated items to a ListBox
   Color.Clear
    For Each Item In NoDupes
        EditForm.Color.AddItem Item
    Next Item

End Sub





