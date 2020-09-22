VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relay"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   2700
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   2700
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Socket2 
      Index           =   0
      Left            =   1800
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   2280
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "IP:"
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Port:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "This program Takes all data recieved on a port and relays it to a specified IP"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code was written by Aaron Couture of
'Technoworld Software. All rights reserved.

'Bugs: If you drop connections the server will crash:
'      This can be caused by pressing the refresh
'      button during a download.

'This code SHOULD not cause ANY damage, but if it does occur I
'Aaron Couture NOR Technoworld Software will be liable for ANY
'damages that may have occured directly or indirectly from the
'use of this code, or any of my other codes.

'Saturday, January 26, 2002.

Dim PortsOpen As New Collection

Private Sub Command1_Click()
On Error GoTo BadAddress
'Check For Blanks
If (Text1.Text = "") Or _
(Text2.Text = "") Then MsgBox "Unable To Start Server: Please Suply All Needed Data(IP, Port)", vbCritical, "Winsock": Exit Sub
Load Socket(1)
'Set the main connection for OPEN
PortsOpen.Add "Open"
Socket(1).LocalPort = Text1.Text
Socket(1).Protocol = sckTCPProtocol
Socket(1).Listen
'Lock/Unlock Items
Command2.Enabled = True
Command1.Enabled = False
Text1.Locked = True
Text2.Locked = True
Exit Sub
'Error handleing
BadAddress:
Unload Socket(1)
MsgBox "Winsock Failed To Initilize! Please check all settings and try again.", vbCritical, "Winsock Error"
End Sub

Private Sub Command2_Click()
On Error Resume Next
'Unload Sockets
For a = 1 To PortsOpen.Count
Socket2(a).Close
Socket(a).Close
Unload Socket2(a)
Unload Socket(a)
Next a
'Set all connections as non-active
Set PortsOpen = Nothing
'Lock/Unlock Items
Command2.Enabled = False
Command1.Enabled = True
Text1.Locked = False
Text2.Locked = False
End Sub

Private Sub Form_Load()
'Lock Stop Button
Command2.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Command2_Click
End
End Sub

Private Sub Socket_Close(Index As Integer)
Dim TmpStack As New Collection

'Close socket
Socket(Index).Close
Socket2(Index).Close
'Set As Closed In Stack
 
For a = 1 To PortsOpen.Count
 If (a <> Index) Then
  TmpStack.Add PortsOpen.Item(a)
 Else
  TmpStack.Add "Closed"
 End If
Next a
Set PortsOpen = Nothing
For a = 1 To TmpStack.Count
 PortsOpen.Add TmpStack.Item(a)
Next a
End Sub

Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'Accept Connection
For a = 1 To PortsOpen.Count
If (PortsOpen.Item(a) = "Closed") Then Port% = a: Exit Sub
Next a
If (Port% = 0) Then
PortsOpen.Add ("Open")
Port% = PortsOpen.Count
 If (Port% = 256) Then PortsOpen.Remove (PortsOpen.Count): Exit Sub 'This server only supports 255 ports max!
Load Socket(Port%)
Load Socket2(Port%)
End If
'Open The Socket
Socket(Port%).Accept requestID

'Setup Socket
Socket2(Port%).Protocol = sckTCPProtocol
Socket2(Port%).RemotePort = Text1.Text
Socket2(Port%).RemoteHost = Text2.Text
Socket2(Port%).Connect
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
'Data recieved, and Relay!
Dim SocketData As String
Socket(Index).GetData SocketData
 'Wait for connection
 Do
 DoEvents
 If (Socket2(Index).State = sckError) Then Socket2(Index).Close: Exit Sub
 Loop Until (Socket2(Index).State = sckConnected)
 Socket2(Index).SendData SocketData
End Sub

Private Sub Socket2_Close(Index As Integer)
'Close socket
Socket(Index).Close
End Sub

Private Sub Socket2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
'Data recieved, and Relay!
Dim SocketData As String
Socket2(Index).GetData SocketData
 Do
 DoEvents
 If (Socket(Index).State = sckError) Then Socket(Index).Close: Exit Sub
 Loop Until (Socket(Index).State = sckConnected)
 Socket(Index).SendData SocketData
End Sub
