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
Pck = "6¿Ä" & Crypt(0) & "¿Ä96¿Ä" & Crypt(1) & "¿Ä0¿Ä" & YahooID & "¿Ä2¿Ä" & YahooID & "¿Ä192¿Ä-1¿Ä2¿Ä1¿Ä1¿Ä" & YahooID & "¿Ä99¿Äbeta¿Ä135¿Ä6,0,0,1555¿Ä148¿Ä300¿Ä59¿ÄB04um3lh08ql2q&b=2¿Ä59¿Ä¿Ä"
Login = Header("54", Pck)
End Function

'/* Login data for authentication
Public Function Data(YahooID As String) As String
Dim Pck As String
Pck = "1¿Ä" & YahooID & "¿Ä"
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
Packet = "1¿Ä" & from & "¿Ä7¿Ä" & whoto & "¿Ä14¿Ä¿Ä65¿Ä" & Group & "s¿Ä97¿Ä1¿Ä216¿Ä¿Ä"
AddMyFriend = Header("D0", Packet)
End Function

'/* Delete friend packet...requires group name
Public Function DeleteFriend(from As String, FriendToDelete As String, Group As String) As String
Dim Packet As String
Packet = "1¿Ä" & from & "¿Ä7¿Ä" & FriendToDelete & "¿Ä65¿Ä" & Group & "¿Ä"
DeleteFriend = Header("84", Packet)
End Function

'/* the status packet send for cam...idle...bust....etc
Public Function Status(message As String, busy As Boolean) As String
Dim Packet As String
If busy = True Then
Packet = "10¿Ä99¿Ä19¿Ä" & message & "¿Ä47¿Ä1¿Ä187¿Ä0¿Ä"
Else
Packet = "10¿Ä99¿Ä19¿Ä" & message & "¿Ä47¿Ä0¿Ä187¿Ä0¿Ä"
End If
Status = Header("C6", Packet)
End Function

'/* the infamous Buddy denial packet that removes u from their list
Public Function Deny(from As String, whoto As String, message As String) As String
Dim Packet As String
Packet = "1¿Ä" & from & "¿Ä7¿Ä" & whoto & "¿Ä14¿Ä" & message & "¿Ä"
Deny = Header("86", Packet)
End Function

'/* leave room packet for YMSG...not sure why i left it here..lol
Public Function LeaveRoom(user As String) As String
Dim Packet As String
Packet = "1¿Ä" & user & "¿Ä1005¿Ä322" & "85272¿Ä"
LeaveRoom = Header("A0", Packet)
End Function

'/* will make your logged in id into ivisible
Public Function Invisible() As String
'This will make you Invisible
Dim Packet As String
Packet = "13¿Ä2¿Ä"
Invisible = Header("C5", Packet)
End Function

