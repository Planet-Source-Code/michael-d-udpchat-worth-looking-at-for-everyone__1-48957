Attribute VB_Name = "Functions"
Public Declare Function GetTickCount& Lib "kernel32" ()
Public Const ChatPort As Long = "1025"
Public Const vbGold = 49344
Public Const vbDGreen = 32768
Public Const vbOrange = 33023
Public Const vbPurple = 8388736
Public Const SendPingT As Long = 30
Public Const PingTimeout As Long = 10
Public Connected As Boolean
Public Username As String
Public lItem As ListItem
Public LocalIP As String
Public StartTime As Long

Sub Log(User As String, Data As String, Colour As ColorConstants)
'Data = "[" & Time & "] " & Data
With frmMain.txtChat
    If Not User = "" Then
    .SelStart = Len(.Text)
    .SelColor = vbBlack
    If .Text = "" Then
    .SelText = "<"
    Else
    .SelText = vbCrLf & "<"
    End If
    .SelStart = Len(.Text)
    .SelColor = vbBlue
    .SelText = User

    .SelStart = Len(.Text)
    .SelColor = vbBlack
    .SelText = "> "
    Else
    If Not .Text = "" Then
    .SelStart = Len(.Text)
    .SelText = vbCrLf
    End If
    End If
    .SelStart = Len(.Text)
    .SelColor = Colour
    .SelText = Data
    .SelStart = Len(.Text)
    DoEvents
End With
End Sub

Sub SendRoomMessage(Msg As String)
With frmMain.sckServ
For i = 1 To frmMain.Users.Count - 1
    If frmMain.Users(i).Online = True Then
    .RemoteHost = frmMain.Users(i).IP
    .RemotePort = ChatPort
    .SendData Msg
    End If
Next i
End With
End Sub

Function GetRUser(IP As String)
For i = 1 To frmMain.Users.Count - 1
If IP = frmMain.Users(i).IP And frmMain.Users(i).Online Then
GetRUser = frmMain.Users(i).Username
End If
Next i
If GetRUser = "" Then GetRUser = "UNKNOWN"
End Function

Function GetRIP(User As String)
For i = 1 To frmMain.Users.Count - 1
If User = frmMain.Users(i).Username And frmMain.Users(i).Online Then
GetRIP = frmMain.Users(i).IP
End If
Next i
If GetRIP = "" Then GetRIP = "0.0.0.0"
End Function

Public Function SecToStr(TimeInSec As Long) As String
Dim ActualTime As String
Dim Sec As Single
Dim Min As Single
Dim Hour As Single
On Error Resume Next
TimeInSec = Fix(TimeInSec)
Hour = IIf(TimeInSec >= 3600, Int(TimeInSec / 3600), 0)
Min = IIf((Hour > 0 And Not TimeInSec Mod 3600 = 0) Or TimeInSec < 3600, Fix(TimeInSec / 60) - Hour * 60, 0)
Sec = IIf(Min >= 0, TimeInSec - ((3600 * Hour) + (60 * Min)), 0)
SecToStr = Hour & " hours, " & Min & " minutes, " & Sec & " seconds"
End Function

Sub ResetTO(IP As String)
For i = 1 To frmMain.Users.Count - 1
If IP = frmMain.Users(i).IP And frmMain.Users(i).Online Then
frmMain.Users(i).ReactivePing
End If
Next i
End Sub

Sub PrepEntry(Index)
On Error Resume Next 'Reusing Object :)
Load frmMain.Users(Index)
End Sub

Sub PDisconnect()
'Disconnect
Dim C As Long
SendRoomMessage "DIS" 'Disconnect PROPERLY...Stops the remote chatters from having to ping timeout your connection
C = frmMain.Users.Count - 1
For i = 1 To C
Unload frmMain.Users(i)
Next i
frmMain.sckServ.Close
Log "UDPChat", "Disconnected", vbPurple
frmMain.lvUsers.ListItems.Clear
frmMain.lblStatus = "Offline"
Connected = False
frmMain.mnuFConnect.Enabled = True
frmMain.mnuFDisconnect.Enabled = False
End Sub
