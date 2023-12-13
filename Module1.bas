Attribute VB_Name = "Module1"
Sub PrintTabSetup()
'
' Macro2 Macro
' ActiveWorkbook.Unprotect ("Baker222")
' Clean up CR and LF - set pages for each plan
' Print setup for each tab

Dim objFSO As Object
Dim colDrives As Object
Dim strOut As String
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set colDrives = objFSO.Drives
On Error Resume Next
'File system errors for virtual drives - Find the J:drive path
For Each objDrive In colDrives
    strOut = "Drive letter: " & objDrive.DriveLetter & vbNewLine
    DriveL = objDrive.DriveLetter
    strOut = strOut & ("Drive type: " & Choose(objDrive.DriveType + 1, "Unknown", "Removable", "Fixed", "Network", "CD-ROM", "RAM Disk") & vbNewLine)
    strOut = strOut & ("File system: " & objDrive.FileSystem & vbNewLine)
    strOut = strOut & ("Path: " & objDrive.Path & vbNewLine)
    strOut = strOut & ("RootFolder: " & objDrive.RootFolder & vbNewLine)
    strOut = strOut & ("VolumeName: " & objDrive.VolumeName & vbNewLine)
    strOut = strOut & ("Path: " & objDrive.Path & vbNewLine)
    strOut = strOut & ("ShareName: " & objDrive.ShareName)
    SName = objDrive.ShareName
    If UCase(SName) = "\\SITMCTCR.SERVERS.CHRYSLER.COM\CTCGROUPS" Then Exit For
  '  MsgBox strOut
Next


Req = Sheets("Tests").Range("H1").Value
req1 = Sheets("Mechanic Check In-Out").Range("AE2").Value

'On Error Resume Next
Dr = DriveL & ":\5140_DTC logs Check-in Check-out\" & Sheets("Mechanic Check In-Out").Range("AB2").Value & " Vehicles\V" & Sheets("Tests").Range("H1").Value & " " & Sheets("Mechanic Check In-Out").Range("AE2").Value

MkDir Dr

BB = Dr & "\Check in-out"
MkDir Dr & "\Check in-out"

'On Error Resume Next
'Dr = "J:\5140_DTC logs Check-in Check-out\" & Sheets("Mechanic Check In-Out").Range("AB2").Value & " Vehicles\V" & Sheets("Tests").Range("H1").Value & " " & Sheets("Mechanic Check In-Out").Range("AE2").Value

'MkDir Dr
'MkDir Dr & "\Check in-out"



MkDir Dr & "\VSTR"
  
MkDir Dr & "\VATC"

MkDir Dr & "\VRTC"
 
MkDir Dr & "\VTEM"

MkDir Dr & "\VESD"

MkDir Dr & "\Transient"


End Sub
Sub MakeCopy()
Attribute MakeCopy.VB_ProcData.VB_Invoke_Func = " \n14"
'
'  Macro
'

'
    Sheets("Request DB").Visible = True
    Sheets("Request DB").Select
    RowNo = ActiveCell.Row
    ColNo = ActiveCell.Column
    Rng = "A" & RowNo & ": AA" & RowNo
    Range(Rng).Select
    
    Selection.Copy
     Sheets("Test Mod").Visible = True
    Sheets("Test Mod").Select
    
    
    Row1 = Worksheets("Test Mod").Range("B3")
    
    ActiveWindow.SmallScroll Down:=Row1 + 1
    
    Range("e" & Row1 + 1).Select
    ActiveSheet.Paste
    
    Range("D" & Row1 + 1) = Date
    ActiveSheet.Paste
    NewNo = Range("A" & Row1)
    Application.CutCopyMode = False
    Selection.Copy
    Range("A" & Row1 + 1) = NewNo + 1
    
End Sub

Sub AddReq()
'
'  Macro
'

'

    Sheets("Request DB").Select
    Worksheets("Request DB").Unprotect
    RowNo = Range("C2") + 4
    NewRequest = Range("e2") + 1
    Rows(RowNo).Select
    RowNo = ActiveCell.Row
    ColNo = ActiveCell.Column
    If (RowNo < 4) Then GoTo quit
    
    Rng = "A" & RowNo & ": M" & RowNo
    Range(Rng).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Editor").Visible = True
    
    Sheets("Editor").Select
    Sheets("Request DB").Visible = False
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
     Range("B1") = "add"
     Range("B2") = NewRequest
     Range("B4") = Date
     
        
quit:
End Sub

Sub Savedata()
Attribute Savedata.VB_ProcData.VB_Invoke_Func = " \n14"
'
'  Macro
'

'
    
    Sheets("Editor").Select
    If (Range("B1") = "edit") Then Call MakeCopy

    Sheets("Request DB").Visible = True
    Sheets("Editor").Select
    
    Range("B2:B28").Select
    Selection.Copy
    
    
    Sheets("Request DB").Visible = True
    Sheets("Request DB").Select
   
    
     
    RowNo = ActiveCell.Row
    ColNo = ActiveCell.Column
    Rng = "A" & RowNo
    Range(Rng).Select
   
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
   
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowSorting:=True, AllowFiltering:=True
   
   Sheets("Editor").Visible = False
   Sheets("Test Mod").Visible = False
End Sub
Sub CancelUpdate()
Attribute CancelUpdate.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'

'
    Sheets("Editor").Select
    Sheets("Request DB").Visible = True
    Sheets("Editor").Select
    ActiveWindow.SelectedSheets.Visible = False
End Sub

Private Sub AddRequestFromWorksheet()

Dim emptyRow As Long

'Make Sheet1 active
 Sheets("Request DB").Select
    Worksheets("Request DB").Unprotect
    RowNo = Range("C2") + 4
    NewRequest = Range("e2") + 1
'Determine emptyRow
 emptyRow = Range("C2") + 4

'Transfer information
 Cells(emptyRow, 1).Value = NewRequest
 Cells(emptyRow, 2).Value = TestLab.Value
 Cells(emptyRow, 3).Value = ReqDate.Value
 Cells(emptyRow, 5).Value = Component.Value





End Sub
