VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deny Buddy Example"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3315
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Sock 
      Left            =   3240
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete From Group"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "Group Name"
      Top             =   2640
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Deny"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "You Have Just been Removed !!"
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Their ID"
      Top             =   1440
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Logout"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "TestPass"
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Yahoo! ID"
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/Lets start
Private Sub Command1_Click()
ID = Text1.Text
pass = Text2.Text
Sock.Close 'Stop overlapping scK errors
Sock.Connect "scs.msg.yahoo.com", 5050 'server name and port number
End Sub

Private Sub Command3_Click()
Sock.SendData Deny(Text1, Text3, Text4) 'sends packet to deny to server
Pause (0.5) '- pausing to stop false positives
Label1.Caption = "Buddy Removal Successful!" 'sucess message in label
End Sub

Private Sub Command4_Click()
Sock.SendData DeleteFriend(Text1, Text3, Text5) '- sending group delete packet to yahoo
Pause (0.5)
Label1.Caption = "Group Removal Successful!" 'success group delete message
End Sub

'/* Winsock connection
Private Sub Sock_Connect()
Label1.Caption = "Connecting.." 'telling this text to show in label
Sock.SendData Data(ID) 'sending data to yahoo
End Sub

'/* winsock data arrival from yahoo!
Private Sub Sock_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
Sock.GetData Data 'incoming data back from yahoo
Debug.Print Data
  If Mid(Data, 12, 1) = "W" Then 'time to split cookie
    Sessionkey = Mid(Data, 17, 4)
    ChallengeString = Mid(Data, 30 + Len(ID), Len(Data) - 29)
    ChallengeString = Replace(ChallengeString, "À€13À€1À€", "")
    Call GetStrings(ID, pass, ChallengeString, Crypt(0), Crypt(1), 1)
    Sock.SendData Login(ID)
    ElseIf Mid(Data, 12, 1) = "T" Then 'T in cookie is bad
    Label1.Caption = "Wrong ID/Pass."
    Sock.Close
    ElseIf Mid(Data, 12, 1) = "U" Then 'U is good
    Sessionkey = Mid(Data, 17, 4)
        Label1.Caption = "OnLine." 'Thank Fuck we're OnLine
End If
End Sub

'/* Error recieved thru winsock control
Private Sub Sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Label1.Caption = "Error!!." 'notify of errors
Sock.Close 'close connection
End Sub

'/* logout
Private Sub Command2_Click()
Sock.Close
Sock.Close '3 offline commands to disconnect socks properly...stops hanging
Sock.Close
Label1.Caption = "Offline." 'offline notification
End Sub

'/* force close of form and winsock
Private Sub Form_Unload(cancel As Integer)
Sock.Close 'close connection
Unload Form1 'close form
End Sub
