Attribute VB_Name = "YMSG"
'/* New YMSG Login
'/* Dermot
Const name As String = "YMSG" '- YMSG10 YMSG11 YMSG12 is the three types
Const Ver As Integer = 11
Public Sessionkey As String, ID As String, pass As String, Buffer As String, Crypt(1) As String, ChallengeString As String
Private Declare Function YMSG12_ScriptedMind_Encrypt Lib "YMSG.dll" (ByVal username As String, ByVal Password As String, ByVal Seed As String, ByVal result_6 As String, ByVal result_96 As String, intt As Long) As Boolean

'/*LOogin key strings split from the DLL
Public Function GetStrings(YahooID As String, YahooPass As String, Seed As String, Str1 As String, Str2 As String, Mode As Long) As Boolean
Dim A(1) As String, B As Long
On Error GoTo err
A(0) = String(100, vbNullChar)
A(1) = String(100, vbNullChar)
GetStrings = YMSG12_ScriptedMind_Encrypt(YahooID, YahooPass, Seed, A(0), A(1), Mode)
B = InStr(1, A(0), vbNullChar)
Str1 = Left$(A(0), B - 1)
B = InStr(1, A(1), vbNullChar)
Str2 = Left$(A(1), B - 1)
Exit Function
err:
GetStrings = False
End Function

'/* each packet has a header...in this case its YMSG
Public Function Header(ByVal PacketType As String, ByVal Pck As String) As String
Dim i As Integer
Dim X As Integer
X = 0
i = Len(Pck)
Do While i > 255
i = i - 256
X = X + 1
Loop
Header = name & Chr(0) & Chr(Ver) & String(2, 0) & Chr(X) & Chr(i) & Chr(0) & _
Chr("&H" & PacketType) & String(8, 0) & Pck
Debug.Print Header
End Function

'/* login info for send to yahoo *id*
Public Function Login(YahooID As String) As String
Dim Pck As String
Pck = "6��" & Crypt(0) & "��96��" & Crypt(1) & "��0��" & YahooID & "��2��" & YahooID & "��192��-1��2��1��1��" & YahooID & "��99��beta��135��6,0,0,1555��148��300��59��B04um3lh08ql2q&b=2��59����"
Login = Header("54", Pck)
End Function

'/* Login data for authentication
Public Function Data(YahooID As String) As String
Dim Pck As String
Pck = "1��" & YahooID & "��"
Data = Header("57", Pck)
End Function

'/* pause timer for many functions in Visual Basics
Sub Pause(ByVal Sec As Long)
Sec = Timer & Sec
Do Until Timer > Sec
    DoEvents
Loop
End Sub

'/* add friend packet for YMSG yahoo! protocol
Public Function AddMyFriend(from As String, whoto As String, Group As String, message As String) As String
Dim Packet As String
Packet = "1��" & from & "��7��" & whoto & "��14����65��" & Group & "s��97��1��216����"
AddMyFriend = Header("D0", Packet)
End Function

'/* Delete friend packet...requires group name
Public Function DeleteFriend(from As String, FriendToDelete As String, Group As String) As String
Dim Packet As String
Packet = "1��" & from & "��7��" & FriendToDelete & "��65��" & Group & "��"
DeleteFriend = Header("84", Packet)
End Function

'/* the status packet send for cam...idle...bust....etc
Public Function Status(message As String, busy As Boolean) As String
Dim Packet As String
If busy = True Then
Packet = "10��99��19��" & message & "��47��1��187��0��"
Else
Packet = "10��99��19��" & message & "��47��0��187��0��"
End If
Status = Header("C6", Packet)
End Function

'/* the infamous Buddy denial packet that removes u from their list
Public Function Deny(from As String, whoto As String, message As String) As String
Dim Packet As String
Packet = "1��" & from & "��7��" & whoto & "��14��" & message & "��"
Deny = Header("86", Packet)
End Function

'/* leave room packet for YMSG...not sure why i left it here..lol
Public Function LeaveRoom(user As String) As String
Dim Packet As String
Packet = "1��" & user & "��1005��322" & "85272��"
LeaveRoom = Header("A0", Packet)
End Function

'/* will make your logged in id into ivisible
Public Function Invisible() As String
'This will make you Invisible
Dim Packet As String
Packet = "13��2��"
Invisible = Header("C5", Packet)
End Function

