VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Main 
   Caption         =   "Client"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConnection 
      Caption         =   "Connect"
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4245
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2880
      Width           =   1500
   End
   Begin MSWinsockLib.Winsock tcpClient 
      Left            =   2475
      Top             =   180
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtData 
      Height          =   2535
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "main.frx":0000
      Top             =   270
      Width           =   4650
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const iPort = 544
Const sIP = "197.193.1.105" '// You'r own IP
Dim sName As String

Private Sub cmdConnection_Click()

   cmdConnection.Enabled = False '// You gotta look nice now a days, and so people don't hit it

   If cmdConnection.Caption = "Disconnect" Then
      tcpClient.Close '// We have to close it.
   Else
      tcpClient.Connect '// Connect to the server
   End If
   
End Sub

Private Sub Form_Load()

   '// Just setting the socket
   tcpClient.RemoteHost = sIP
   tcpClient.RemotePort = iPort
   
   sName = "Your Name" '// The name you wish to present your self with
   
End Sub

Private Sub Form_Resize()
'// Make it pretty when resizing.

   cmdConnection.Width = Main.ScaleWidth
   txtData.Width = Main.ScaleWidth
   txtData.Height = Main.ScaleHeight - 560
   txtSend.Width = Main.ScaleWidth
   txtSend.Top = txtData.Height + 285

End Sub

Private Sub tcpClient_Close()
   
   '// Make it SWEEET
   cmdConnection.Enabled = True
   cmdConnection.Caption = "Connect"
   
End Sub

Private Sub tcpClient_Connect()

   '// Again with the nice looking
   cmdConnection.Enabled = True
   cmdConnection.Caption = "Disconnect"

End Sub

Private Sub tcpClient_DataArrival(ByVal bytesTotal As Long)
Dim sTmp As String

   tcpClient.GetData sTmp '// Load into a temp string
   
   txtData.Text = txtData.Text & Chr(13) & Chr(10) & sTmp '// add it to the text box
   
   Debug.Print "test"
   
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer) '// Lets send it
   
   If KeyAscii = 13 Then '// If you press enter
      tcpClient.SendData sName & "|" & txtSend.Text '// send
      txtSend.Text = "" '// clear the send box
      txtSend.SetFocus '// set the focus
   End If
   
End Sub
