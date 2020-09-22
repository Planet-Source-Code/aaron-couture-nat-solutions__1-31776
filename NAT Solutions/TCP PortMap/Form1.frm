VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form TrayForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TCP Port Mapper"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3855
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Save Settings"
      Height          =   315
      Left            =   2640
      TabIndex        =   10
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop Server"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Server"
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Enable Traffic Logging"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   3735
   End
   Begin VB.TextBox TrafficLog 
      Height          =   2055
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox Port 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox PublicIP 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox PrivateIP 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Server 
      Left            =   960
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Client 
      Left            =   480
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Listen 
      Left            =   1560
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Port:"
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Public:"
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Private:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "TrayForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Code Was Written By Aaron Couture
'Email: AaronMCouture@hotmail.com

'Attention:

'I am looking for good VB programmers that are very good with
'Winsock to help me program a Multiplayer game server. This game
'server will host up to 3,000 users. 45% of the server is done.
'If you feal like helping on this project please send me an Email.

'Also, If you figure out how to make this program map more then
'one port, please let me know, becouse as far as i know it is
'imposible with the winsock control...

'Bugs:

'You cannot put DNS (ex: 'google.com') in for an IP, you must get the
'IP of the DNS. If you do put the DNS this program will NOT work!
'----------------------------------------------------------------

Public Logging As Boolean

Private Sub Check1_Click()
 If (Check1.Value = 1) Then Logging = True
 If (Check1.Value = 0) Then Logging = False
End Sub

Private Sub Client_Close()
Call Server_Close
End Sub

Private Sub Command1_Click()
'Check Items
If (Port.Text = "") Or (PrivateIP.Text = "") Or (PublicIP.Text = "") Then MsgBox "Winsock Error: Winsock Settings NOT Set. Please Configure Winsock Before Starting TCP Port Mapping Server!", vbCritical, "Winsock": Exit Sub
'Lock/Unlock Items
Command1.Enabled = False
Command2.Enabled = True
PrivateIP.Locked = True
PublicIP.Locked = True
Port.Locked = True
'Start Server
 Listen.LocalPort = Port.Text
 Listen.Protocol = sckTCPProtocol
 Listen.Listen
End Sub

Private Sub Command2_Click()
'Lock/Unlock Items
Command1.Enabled = True
Command2.Enabled = False
PrivateIP.Locked = False
PublicIP.Locked = False
Port.Locked = False
'Stop Server
 Listen.Close
 Client.Close
 Server.Close
End Sub

Private Sub Command3_Click()
WriteIni
End Sub

Private Sub Form_Load()
'Read Settings
ReadIni
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Listen.Close
 Client.Close
 Server.Close
 RemoveFromTray
End
End Sub

Private Sub Form_Resize()
If (Me.WindowState = 1) Then
 AddToTray Me.Icon, "TCP PortMapper", Me
 Me.WindowState = 0
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (RespondToTray(X) = 1) Then
 RemoveFromTray
 Me.Show
End If
End Sub

Private Sub Listen_ConnectionRequest(ByVal requestID As Long)

 If (Listen.RemoteHostIP = PublicIP.Text) Then
 'Connection Was Requested By A Public IP Address
 Server.Accept requestID
  'Now, Send Connection Request To Private IP
  Listen.Close
  Client.Protocol = sckTCPProtocol
  Client.RemoteHost = PrivateIP.Text
  Client.LocalPort = Port
  Client.RemotePort = Port
  Client.Connect
 ElseIf (Listen.RemoteHostIP = PrivateIP.Text) Then
 'Connection Was Requested By A Private IP Address
 Client.Accept requestID
  'Now, Send Connection Request To Private IP
  Listen.Close
  Server.Protocol = sckTCPProtocol
  Server.RemoteHost = PublicIP.Text
  Server.LocalPort = Port
  Server.RemotePort = Port
  Server.Connect
 End If
End Sub

Private Sub Server_Close()
On Error Resume Next
'Server Closed Connection; Restart Listing...
Client.Close
Server.Close
Listen.LocalPort = Port.Text
Listen.Protocol = sckTCPProtocol
Listen.Listen
End Sub

Private Sub Server_DataArrival(ByVal bytesTotal As Long)
Dim SocketData As String
Server.GetData SocketData
'Send Data
  Do
  DoEvents
  Loop Until Client.State = sckConnected
  Client.SendData SocketData
'Update Traffic Log
 If (Logging = True) Then TrafficLog.Text = TrafficLog.Text + Server.RemoteHostIP + ": " + SocketData + vbCrLf
End Sub

Private Sub Client_DataArrival(ByVal bytesTotal As Long)
Dim SocketData As String
Client.GetData SocketData
'Send Data
 Do
 DoEvents
 Loop Until Server.State = sckConnected
 Server.SendData SocketData
'Update Traffic Log
 If (Logging = True) Then TrafficLog.Text = TrafficLog.Text + Client.RemoteHostIP + ": " + SocketData + vbCrLf
End Sub
