VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Chat Server"
   ClientHeight    =   3645
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5025
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock tcpSocket 
      Index           =   0
      Left            =   1080
      Top             =   1485
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   5010
   End
   Begin VB.Menu mnuServer 
      Caption         =   "Server"
      Visible         =   0   'False
      Begin VB.Menu mServer 
         Caption         =   "Start Server"
         Index           =   0
      End
      Begin VB.Menu mServer 
         Caption         =   "Stop Server"
         Index           =   1
      End
      Begin VB.Menu mServer 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mServer 
         Caption         =   "End Server"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ######################################################################
' #  The server part of a VB chat project.                             #
' #  I started this project to teach others, how to program network    #
' #  applications. And this is what I made.                            #
' #                                                                    #
' #  I have choosen to add most of the feature to "modMain".           #
' #                                                                    #
' #  ---INFO---                                                        #
' #    Port: 544                                                       #
' ######################################################################

Option Explicit


Private Sub Form_Unload(Cancel As Integer)

   Cancel = NoUnload '// See modMain

End Sub

Private Sub mServer_Click(Index As Integer)

   Select Case Index
      Case 0
         StartServer '// See modMain
      Case 1
         StopServer '// See modMain
      Case 3
         EndApplication '// See modMain
   End Select
   
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   TrayMove X '// See modMain, more used that then word "Sub" in this project

End Sub


Private Sub Form_Resize() '// So it scale as it should

   txtLog.Width = frmMain.ScaleWidth
   txtLog.Height = frmMain.ScaleHeight
   
End Sub

Private Sub tcpSocket_Close(Index As Integer)
Dim msgClose As String

   '// Show a close connection text
   msgClose = Chr(13) & Chr(10) & Chr(13) & Chr(10) & "-Connection Closed-"
   msgClose = msgClose & Chr(13) & Chr(10) & "    IP: " & tcpSocket(Index).RemoteHostIP
   msgClose = msgClose & Chr(13) & Chr(10) & "    Socket: " & Index
   
   txtLog.Text = txtLog.Text & msgClose

   '// Close the actual connection
   tcpSocket(Index).Close
   
   '// And unload the socket, just to help out that poor memory
   Unload tcpSocket(Index)
   
End Sub

Private Sub tcpSocket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim msgConn As String
   
   If Index = 0 Then '// why check, but just to be sure
   
      iSocket = iSocket + 1 '// add a little something to the socket count
      
      Load tcpSocket(iSocket) '// load new socket
      
      tcpSocket(iSocket).LocalPort = 544 '// set port
      tcpSocket(iSocket).Accept requestID '// and accept the request
      
      '// Like always we as a admin like to keep track of stuff, so add to log.
      '// hehe admin on a chat server (wow my dream job) :)
      msgConn = Chr(13) & Chr(10) & Chr(13) & Chr(10) & "-Connection Request-"
      msgConn = msgConn & Chr(13) & Chr(10) & "    IP: " & tcpSocket(Index).RemoteHostIP
      msgConn = msgConn & Chr(13) & Chr(10) & "    ID: " & requestID
      msgConn = msgConn & Chr(13) & Chr(10) & "    Connections: " & iSocket
      
      txtLog.Text = txtLog.Text & msgConn
      
   End If
End Sub

Private Sub tcpSocket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim sTmp As String

   tcpSocket(Index).GetData sTmp '// Loading the data into a temp string
   
   GetData sTmp '/* passing it on to our not needed sub (but we still have it,
                ' why is that, well read more about it in my next book, or
                ' just find it in modMain :)*/
   
End Sub
