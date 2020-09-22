Attribute VB_Name = "Module2"
'Read INI Module

Public Sub ReadIni()
On Error GoTo MakeIni
Open App.Path + "/" + "Settings.inf" For Input As #1
Do
Line Input #1, Data$
 If (Left$(LCase$(Data$), 9) = "privateip=") Then
  TrayForm.PrivateIP.Text = Right$(Data$, Len(Data$) - 9)
 End If
 If (Left$(LCase$(Data$), 8) = "publicip=") Then
  TrayForm.PrivateIP.Text = Right$(Data$, Len(Data$) - 8)
 End If
 If (Left$(LCase$(Data$), 5) = "port=") Then
  TrayForm.Port = Right$(Data$, Len(Data$) - 5)
 End If
Loop Until EOF(1)
Close #1
Exit Sub
MakeIni: Close #1: MakeIni
End Sub

Public Sub MakeIni()
Open App.Path + "/" + "Settings.inf" For Output As #1
Print #1, ""
Close #1
End Sub

Public Sub WriteIni()
On Error Resume Next
Open App.Path + "/" + "Settings.inf" For Output As #1
Print #1, "PrivateIP=" + TrayForm.PrivateIP.Text
Print #1, "PublicIP=" + TrayForm.PublicIP.Text
Print #1, "Port=" + TrayForm.Port.Text
Close #1
End Sub
