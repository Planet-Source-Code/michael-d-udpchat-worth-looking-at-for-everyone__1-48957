VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "UDP Chat"
   ClientHeight    =   5490
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8805
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   8805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   255
      Left            =   6960
      TabIndex        =   5
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox txtMsg 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   5160
      Width           =   6735
   End
   Begin MSComctlLib.ListView lvUsers 
      Height          =   4695
      Left            =   6960
      TabIndex        =   3
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Username"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSWinsockLib.Winsock sckServ 
      Left            =   8400
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "0"
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8281
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":1272
   End
   Begin UDPChat.UserHandler Users 
      Index           =   0
      Left            =   7920
      Top             =   0
      _extentx        =   873
      _extenty        =   873
   End
   Begin VB.Label lblStatus 
      Caption         =   "Offline"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Status:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFConnect 
         Caption         =   "&Connect"
      End
      Begin VB.Menu mnuFDisconnect 
         Caption         =   "&Disconnect"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBrk 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuSCU 
         Caption         =   "Change Username"
      End
   End
   Begin VB.Menu mnuCCM 
      Caption         =   "Chat Context Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuCCMGU 
         Caption         =   "&Uptime"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSend_Click()
On Error GoTo e
Dim EM As String
If txtMsg = "" Then Exit Sub
If Connected = False Then
Log "UDPChat", "Click File -> Connect First!", vbRed
Else
If LCase(Split(txtMsg, " ")(0)) = "/me" Then
If Len(txtMsg) < 5 Then Exit Sub
EM = Mid(txtMsg, 5)
If LCase(EM) = UCase(EM) Then txtMsg = "": Exit Sub 'Its either spaces, or dots or non-alpha chars...Need letters.. :)
Log "", Username & " " & EM, vbGold
SendRoomMessage "EMO " & EM
Else
Log Username, txtMsg, vbBlue
SendRoomMessage "MSG " & txtMsg
End If
txtMsg = ""
End If
Exit Sub
e:
Log "", "Error in Send Procedure", vbRed
End Sub

Private Sub Form_Load()
Log "UDPChat", "UDPChat Loaded", vbPurple
Log "UDPChat", "Click File -> Connect to go online", vbPurple
Username = VBA.GetSetting("UDPChat", "Settings", "Username", "")
If Len(Username) > 16 Then Username = Left(Username, 16)
Username = Replace(Username, " ", "_")
If Not Username = "" Then
Log "UDPChat", "Username set to " & Username, vbPurple
End If
StartTime = GetTickCount / 1000
If UCase(Command$) = "-CONNECT" Then
mnuFConnect_Click
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Connected = True Then
SendRoomMessage "DIS" 'Disconnect PROPERLY...Stops the remote chatters from having to ping timeout your connection
End If
End
End Sub

Private Sub Form_Resize()
'Gotta support resize!
If Me.WindowState = 1 Then Exit Sub 'Its minimized
On Error GoTo e
Me.lblStatus.Width = Me.Width - 990
Me.txtChat.Width = Me.Width - 2190
Me.lvUsers.Left = Me.Width - 1950
Me.txtMsg.Width = Me.Width - 2190
Me.cmdSend.Left = Me.Width - 1950

Me.txtChat.Height = Me.Height - 1485
Me.lvUsers.Height = Me.Height - 1485
Me.txtMsg.Top = Me.Height - 1020
Me.cmdSend.Top = Me.Height - 1020
Exit Sub
e:
End Sub

Private Sub lvUsers_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If lvUsers.ListItems.Count = 0 Then Exit Sub
If Button = 2 Then
PopupMenu mnuCCM, , lvUsers.Left + 50
End If
End Sub

Private Sub mnuCCMGU_Click()
If lvUsers.SelectedItem.Text = Username Then
Log Username, "Windows Uptime: " & SecToStr(GetTickCount / 1000) & ". UDP Chat Uptime: " & SecToStr((GetTickCount / 1000) - StartTime), vbBlack
Else
sckServ.RemoteHost = GetRIP(lvUsers.SelectedItem.Text)
sckServ.RemotePort = ChatPort
sckServ.SendData "UPT"
End If
End Sub

Private Sub mnuFConnect_Click()
On Error GoTo e
Dim C As Long
If Username = "" Then
Log "UDPChat", "Can't start chatting: No username set!", vbRed
Exit Sub
End If
C = Users.Count - 1
For i = 1 To C
Unload Users(i)
Next i
'Start the node
Log "UDPChat", "Starting Node", vbBlue
Connected = True
mnuFConnect.Enabled = False
mnuFDisconnect.Enabled = True
lblStatus = "Online"
With sckServ
    LocalIP = .LocalIP 'Ok...Why did I put this here? My 350 mhz Win98 box has a runtime if I dont...it still sends the packet! But it runtimes...and error handling is pointless...plus this saves the winsock control some work...Dunno why it happens, do you? Possibly the protocol switch, but i put THAT in to stop my 350 mhz from runtiming due to a 'blocking' or something...LOL
    .Close
    .Protocol = sckUDPProtocol
    .LocalPort = ChatPort
    .RemotePort = ChatPort
    .RemoteHost = "255.255.255.255"
    .Bind ChatPort, LocalIP
    .SendData "CHK " & Username
End With
lvUsers.ListItems.Clear
Set lItem = lvUsers.ListItems.Add(, , Username)
Log "UDPChat", "Node Online...Ready to chat", vbBlue
Exit Sub
e:
Log "UDPChat", "Error Starting Node: (" & Err.Number & ") " & Err.Description & ". Close any other copies of UDPChat.", vbRed
mnuFConnect.Enabled = True
mnuFDisconnect.Enabled = False
End Sub

Private Sub mnuFDisconnect_Click()
PDisconnect
End Sub

Private Sub mnuFExit_Click()
End
End Sub

Private Sub mnuHAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuSCU_Click()
Dim Blah As String
Blah = InputBox("Your new username?", "New Username", Username)
If Blah = "" Then Exit Sub
If Len(Blah) > 16 Then Blah = Left(Blah, 16)
Username = Blah
Username = Replace(Username, " ", "_")
Log "UDPChat", "Username changed to " & Username, vbRed
VBA.SaveSetting "UDPChat", "Settings", "Username", Username
If Connected = True Then
mnuFDisconnect_Click
mnuFConnect_Click
'Reconnect and Change nick
End If
End Sub

Private Sub sckServ_DataArrival(ByVal bytesTotal As Long)
On Error GoTo e
Dim Data As String
Dim S1 As String
Dim IP As String
Dim rUser As String
Dim tCount As Long
Dim TN As Long
Dim TS As String
sckServ.GetData Data
If Data = "" Then Log "Error", "Blank Packet that should have been " & bytesTotal & " bytes", vbRed: Exit Sub
IP = sckServ.RemoteHostIP
If IP = LocalIP Then Exit Sub 'Broadcast ALSO includes your own node...So ignore it :)
S1 = Split(Data, " ")(0)
If S1 = "CHK" Or S1 = "ACK" Then
tCount = Users.Count
TS = Mid(Data, 5)
For i = 1 To Users.Count - 1
If Users(i).Online = False Then
tCount = i 'Reuse the old entry
End If
Next i
End If
Select Case S1
Case "CHK"
If LCase(Mid(Data, 5)) = LCase(Username) Then 'Thats YOUR username
sckServ.RemoteHost = IP
sckServ.RemotePort = ChatPort
sckServ.SendData "NIU"
Log "", "Someone (" & IP & ") is trying to steal your nick", vbOrange
Exit Sub
End If
sckServ.RemoteHost = IP
sckServ.RemotePort = ChatPort
sckServ.SendData "ACK " & Username
Set lItem = lvUsers.ListItems.Add(, , Mid(Data, 5))
Log "", Mid(Data, 5) & " has joined [" & IP & "]", vbDGreen
PrepEntry tCount
Users(tCount).IP = IP
Users(tCount).Username = Mid(Data, 5)
Users(tCount).TimeOut = 0
Users(tCount).Online = True
Users(tCount).ReactivePing

Case "ACK"
Set lItem = lvUsers.ListItems.Add(, , Mid(Data, 5))
PrepEntry tCount
Users(tCount).IP = IP
Users(tCount).Username = Mid(Data, 5)
Users(tCount).TimeOut = 0
Users(tCount).Online = True
Users(tCount).ReactivePing

Case "NIU" 'Nickname in use
Log "", "That nickname is already in use!", vbRed
PDisconnect

Case "MSG"
rUser = GetRUser(IP)
If rUser = "UNKNOWN" Then
Log "", "Unknown nickname from IP " & IP & " sent a message. Message was dropped", vbRed
Exit Sub
End If
Log rUser, Mid(Data, 5), vbBlack

Case "EMO" 'Emote btw.
rUser = GetRUser(IP)
Log "", rUser & " " & Mid(Data, 5), vbGold

Case "UPT" 'Uptime btw.
rUser = GetRUser(IP)
Log "", rUser & " requested your uptime", vbPurple
TS = "Windows Uptime: " & SecToStr(GetTickCount / 1000) & ". UDP Chat Uptime: " & SecToStr((GetTickCount / 1000) - StartTime)
sckServ.RemoteHost = IP
sckServ.RemotePort = ChatPort
sckServ.SendData "MSG " & TS

Case "PING" 'Ping from other user
sckServ.RemoteHost = IP
sckServ.RemotePort = ChatPort
sckServ.SendData "PONG"
'Log "", "Ping? Pong!", vbYellow

Case "PONG" 'Reply from other user
'Log "", "Pong!", vbYellow
ResetTO IP

Case "PTO" 'Ping Timed Out
'NOTE: This will only be received by the remote end
'If it was just lagging badly
'If it is received the server will try to reestablish
'connection
rUser = GetRUser(IP)
Log "", "Connection to " & rUser & " timed out (Our Pong Packet Dropped Out) Re-establishing connection", vbOrange
sckServ.RemoteHost = IP
sckServ.RemotePort = ChatPort
sckServ.SendData "CHK " & Username
For i = 1 To lvUsers.ListItems.Count
If lvUsers.ListItems(i).Text = rUser Then
lvUsers.ListItems.Remove i
End If
Next i


Case "DIS"
rUser = GetRUser(IP)
Log "", rUser & " has left (Manual disconnect)", vbOrange
For i = 1 To Users.Count - 1
If Users(i).Online = True And Users(i).IP = IP Then
Users(i).Online = False
End If
Next i
For i = 1 To lvUsers.ListItems.Count
If lvUsers.ListItems(i).Text = rUser Then
lvUsers.ListItems.Remove i
Exit Sub
End If
Next i

Case Else
Log "Unknown Message", Data, vbRed

End Select
Exit Sub
e:
Log "Error", "Error parsing " & Data, vbRed
End Sub

Private Sub Users_PingTimedOut(Index As Integer, IP As String, rUser As String)
Users(Index).Online = False
Log "", rUser & " has left (Ping Timeout)", vbOrange
For i = 1 To lvUsers.ListItems.Count
If lvUsers.ListItems(i).Text = rUser Then
lvUsers.ListItems.Remove i
End If
Next i
End Sub

Private Sub Users_SendPing(Index As Integer, IP As String)
sckServ.RemoteHost = IP
sckServ.RemotePort = ChatPort
sckServ.SendData "PING"
End Sub
