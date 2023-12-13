Attribute VB_Name = "Module23"

 
Sub ListDrives()
Dim objFSO As Object
Dim colDrives As Object
Dim strOut As String
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set colDrives = objFSO.Drives
On Error Resume Next
'File system errors for virtual drives
For Each objDrive In colDrives
    strOut = "Drive letter: " & objDrive.DriveLetter & vbNewLine
    strOut = strOut & ("Drive type: " & Choose(objDrive.DriveType + 1, "Unknown", "Removable", "Fixed", "Network", "CD-ROM", "RAM Disk") & vbNewLine)
    strOut = strOut & ("File system: " & objDrive.FileSystem & vbNewLine)
    strOut = strOut & ("Path: " & objDrive.Path & vbNewLine)
    strOut = strOut & ("RootFolder: " & objDrive.RootFolder & vbNewLine)
    strOut = strOut & ("VolumeName: " & objDrive.VolumeName & vbNewLine)
    strOut = strOut & ("Path: " & objDrive.Path & vbNewLine)
    strOut = strOut & ("ShareName: " & objDrive.ShareName)
    MsgBox strOut
Next
On Error GoTo 0
End Sub

