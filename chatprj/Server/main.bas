Attribute VB_Name = "modMain"
Option Explicit

'// The tray icon needs the following to work, as I want it to:
Private Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const WM_MOUSEMOVE = &H200

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203

Private Const WM_RBUTTONUP = &H205
      
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'// that was that, it dosn't need anymore (declaring)


Private Const iPort = 544 '// the port number

Dim nTray As NOTIFYICONDATA '// ohh yes, this is also for the tray Icon
Dim I As Integer
Dim bEndApp As Boolean

Global iSocket As Integer '// The amount of open connections.


' -------------------------------------------------------------------
'   When the project starts, this will run. In the ChatServer.vbp
'   properties, you can see that instead of frmMain, in the startup
'   object it's called main
' -------------------------------------------------------------------
Sub Main()
   
   '/* Set the information about the icon
   ' There is a lot of good examples on using tray icon, so go get them
   ' to understand the following to it's fullest */
   With nTray
      .cbSize = Len(nTray) '// Can't remember, but in an early program I wrote this
      .hWnd = frmMain.hWnd '// The handle
      .uId = vbNull '// see comment at cbSize, :)
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE '// what it should support
      .uCallBackMessage = WM_MOUSEMOVE '// again support
      .hIcon = frmMain.Icon '// the icon to display
      .szTip = "Stopped" & vbNullChar '// The tooltip that appears
   End With
   
   Shell_NotifyIcon NIM_ADD, nTray '// add it to the tray
   
   '// Change the enabled property of the two menus
   frmMain.mServer(1).Enabled = False
   frmMain.mServer(0).Enabled = True
   
   frmMain.txtLog = "Chat Server"
   
   '// Start the server, see below some where, use definition.
   StartServer
   
End Sub


' -------------------------------------------------------------------
'   After receiving the data, you have to resend it.
' -------------------------------------------------------------------
Public Sub SendData(Data As String, sName As String)
On Error GoTo ErrHandler: '// If the socket is unloaded

   For I = 1 To iSocket '// Run through sockets
      frmMain.tcpSocket(I).SendData sName & ": " & Data
   Next

ErrHandler:
   If Err.Number = 340 Then '// array error
      Resume Next
   End If

End Sub


' -------------------------------------------------------------------
'   When you receive data from the client (DataArrival)
' -------------------------------------------------------------------
Public Sub GetData(Data As String)
Dim iTmp As Integer
Dim sName As String

   iTmp = InStr(1, Data, "|")
   sName = Left(Data, iTmp - 1)
   
   iTmp = Len(Data) - iTmp
   Data = Right(Data, iTmp)
   
   SendData Data, sName
   '/* I know that the above isn't needed, but I just keep it anyway
   ' it's now hard to adjust so it's not there, and it would improve
   ' the perfomance. But you could use it to create stats, over users.
   ' You could also use the name as a command selection. For example:
   ' The name is "WhoIs" then you will return data, only to that client
   ' and in that data, there will be some information about who is logged
   ' on. OR you could: "Whisper", this way two people could have a private
   ' chat, about life's important things. */
End Sub


' -------------------------------------------------------------------
'   To check if the user is in our tray icon.
' -------------------------------------------------------------------
Public Sub TrayMove(X As Single)
Dim lTmp As Long
   
   lTmp = X / Screen.TwipsPerPixelX
   
   '// If you press some mouse buttons
   Select Case lTmp
   
      Case WM_LBUTTONDBLCLK '// Left double click
         frmMain.Show
         frmMain.SetFocus
         
      Case WM_RBUTTONUP '// right click (just once)
         frmMain.PopupMenu frmMain.mnuServer
         
   End Select
   
End Sub


' -------------------------------------------------------------------
'   Just minimize when unload form, to tray.
' -------------------------------------------------------------------
Public Function NoUnload() As Integer
   
   frmMain.Hide
   
   If bEndApp = False Then
      NoUnload = -1
      
   ElseIf bEndApp = True Then
      NoUnload = 0
      
   End If
   
End Function


' -------------------------------------------------------------------
'   When server app is closed.
' -------------------------------------------------------------------
Public Sub EndApplication()

   '// Remove tray icon
   Shell_NotifyIcon NIM_DELETE, nTray
   
   bEndApp = True
   
   Unload frmMain
   
End Sub


' -------------------------------------------------------------------
'   Load the server options.
' -------------------------------------------------------------------
Public Sub LoadOptions()

   '/* I had some options in this at the start, it loaded, it from the
   ' registry, using GetSetting, but I don't want the server app
   ' to be a big and graphical needed app (although it wouldn't be).
   ' So I removed the options dialog, and some other fancy features.

   ' I have just reacently added a form to the server. To start with
   ' there should be no GUI, only an invisible running app. But then
   ' I made the log, so people could follow the information.*/
   
End Sub


' -------------------------------------------------------------------
'   Start the server.
' -------------------------------------------------------------------
Public Sub StartServer()
Dim msgStart As String

   '// Start listening
   frmMain.tcpSocket(0).LocalPort = iPort
   frmMain.tcpSocket(0).Listen
   
   '// Change the enable property of the two menus
   frmMain.mServer(0).Enabled = False
   frmMain.mServer(1).Enabled = True

   '// add a little to the log, so the GUI isn't all that bored
   msgStart = Chr(13) & Chr(10) & Chr(13) & Chr(10) & "-Server Start Up-"
   msgStart = msgStart & Chr(13) & Chr(10) & "    Date: " & Date
   msgStart = msgStart & Chr(13) & Chr(10) & "    Time: " & Time
   msgStart = msgStart & Chr(13) & Chr(10) & "    IP: " & frmMain.tcpSocket(0).LocalIP
   msgStart = msgStart & Chr(13) & Chr(10) & "    Port: " & iPort
   
   frmMain.txtLog.Text = frmMain.txtLog.Text & msgStart
   
   '// Update the icon and caption
   frmMain.Caption = "Running: " & frmMain.tcpSocket(0).LocalHostName & ":" & iPort & vbNullChar
   nTray.szTip = "Running: " & frmMain.tcpSocket(0).LocalHostName & ":" & iPort & vbNullChar
   Shell_NotifyIcon NIM_MODIFY, nTray

End Sub


' -------------------------------------------------------------------
'   Stop the server.
' -------------------------------------------------------------------
Public Sub StopServer()
On Error Resume Next '// Can't remember the error, and can't locate it again
'// but I still have this command, so you wont have to fix it.

Dim msgStop As String

   '// Close the first control
   frmMain.tcpSocket(0).Close
   
   '// Loop all the loaded sockets, and close them...ohh yes and unload them
   For I = 1 To iSocket
      frmMain.tcpSocket(I).Close
      Unload frmMain.tcpSocket(I)
   Next
   
   frmMain.mServer(1).Enabled = False
   frmMain.mServer(0).Enabled = True
   
   '// Make sure we start from fresh with loading sockets
   iSocket = 0
   
   '// add a stop text to the log
   msgStop = Chr(13) & Chr(10) & Chr(13) & Chr(10) & "-Server Stopped-"
   msgStop = msgStop & Chr(13) & Chr(10) & "    Date: " & Date
   msgStop = msgStop & Chr(13) & Chr(10) & "    Time: " & Time
   msgStop = msgStop & Chr(13) & Chr(10) & "    IP: " & frmMain.tcpSocket(0).LocalIP
   msgStop = msgStop & Chr(13) & Chr(10) & "    Port: " & iPort
   
   frmMain.txtLog.Text = frmMain.txtLog.Text & msgStop
   
   '// Update the icon and caption (most important part, of all time :))
   frmMain.Caption = "Stopped"
   nTray.szTip = "Stopped"
   Shell_NotifyIcon NIM_MODIFY, nTray
   
End Sub
