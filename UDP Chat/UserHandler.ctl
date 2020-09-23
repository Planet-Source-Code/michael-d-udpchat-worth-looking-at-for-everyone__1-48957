VERSION 5.00
Begin VB.UserControl UserHandler 
   BackColor       =   &H000000FF&
   CanGetFocus     =   0   'False
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   495
   ScaleMode       =   0  'User
   ScaleWidth      =   206.25
   Begin VB.Timer tmrSendPing 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   -120
      Top             =   0
   End
   Begin VB.Timer tmrPTO 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   0
   End
End
Attribute VB_Name = "UserHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Why did I put this here? I couldnt get collections working
'So I whipped up this simple control
'Which is split into arrays :)
Public Username As String
Public IP As String
Public TimeOut As Long
Public Online As Boolean
Private PSP As Long
Event PingTimedOut(IP As String, rUser As String)
Event SendPing(IP As String)

Private Sub tmrPTO_Timer()
TimeOut = TimeOut + 1
If TimeOut >= PingTimeout Then
RaiseEvent PingTimedOut(IP, Username)
tmrPTO = False
End If
End Sub

Private Sub tmrSendPing_Timer()
If Online = True Then
PSP = PSP + 1
If PSP >= SendPingT Then
RaiseEvent SendPing(IP)
PSP = PSP + 1
tmrSendPing = False
tmrPTO = True
End If
End If
End Sub

Private Sub UserControl_Resize()
Width = 500
Height = 500
End Sub

Sub ReactivePing()
tmrSendPing = True
tmrPTO = False
TimeOut = 0
PSP = 0
End Sub
