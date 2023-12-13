Attribute VB_Name = "Module5"

Sub SetWindowSize1(size)
If size = 1 Then
    Application.WindowState = xlNormal
    Application.Top = 1
    Application.Left = 1
    Application.Width = 30
    Application.Height = 20
    End If
If size = 2 Then
     ActiveWindow.WindowState = xlMaximized
    End If
    
    
End Sub

