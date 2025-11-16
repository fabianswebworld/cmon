Attribute VB_Name = "modTCP"
Private Declare Function gethostbyname Lib "wsock32.dll" ( _
  ByVal name As String) As Long
Private Declare Function socket Lib "wsock32.dll" ( _
  ByVal af As Long, _
  ByVal prototype As Long, _
  ByVal protocol As Long) As Long
Private Declare Function closesocket Lib "wsock32.dll" (ByVal s As Long) As Long
Private Declare Function connect Lib "wsock32.dll" ( _
  ByVal s As Long, _
  name As SOCKADDR, _
  ByVal namelen As Long) As Long
Private Declare Function send Lib "wsock32.dll" ( _
  ByVal s As Long, _
  buf As Any, _
  ByVal length As Long, _
  ByVal flags As Long) As Long
Private Declare Function recv Lib "wsock32.dll" ( _
  ByVal s As Long, _
  buf As Any, _
  ByVal length As Long, _
  ByVal flags As Long) As Long
Private Declare Function ioctlsocket Lib "wsock32.dll" ( _
  ByVal s As Long, _
   ByVal cmd As Long, _
  argp As Long) As Long
Private Declare Function inet_addr Lib "wsock32.dll" ( _
  ByVal cp As String) As Long
Private Declare Function htons Lib "wsock32.dll" ( _
  ByVal hostshort As Integer) As Integer
Private Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
Private Declare Sub MoveMemory Lib "kernel32" _
  Alias "RtlMoveMemory" ( _
  Destination As Any, _
  Source As Any, _
  ByVal length As Long)
  
Declare Function WSAStartup Lib "wsock32.dll" ( _
  ByVal wVersionRequested As Integer, _
  lpWSAData As WSAData) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
 
Type WSAData
  wVersion As Integer
  wHighVersion As Integer
  szDescription As String * 257
  szSystemStatus As String * 129
  iMaxSockets As Long
  iMaxUdpDg As Long
  lpVendorInfo As Long
End Type
 
Private Type HOSTENT
  hname As Long
  haliases As Long
  haddrtype As Integer
  hlength As Integer
  haddrlist As Long
End Type
 
Private Type SOCKADDR
  sin_family As Integer
  sin_port As Integer
  sin_addr As Long
  sin_zero As String * 8
End Type
 
Private Const AF_INET = 2
 
Private Const SOCK_STREAM = 1

Private Const SOCK_DGRAM = 2
 
Private Const MSG_PEEK = &H2
 
Private Const FIONBIO = &H8004667E

Public Function GetIP(ByVal HostName As String) As String
  Dim pHost As Long, HostInfo As HOSTENT
  Dim pIP As Long, IPArray(3) As Byte
 
  pHost = gethostbyname(HostName)
  If pHost = 0 Then Exit Function
 
  MoveMemory HostInfo, ByVal pHost, Len(HostInfo)
 
  ReDim IpAddress(HostInfo.hlength - 1)
  MoveMemory pIP, ByVal HostInfo.haddrlist, 4
  MoveMemory IPArray(0), ByVal pIP, 4
 
  GetIP = IPArray(0) & "." & IPArray(1) & "." & IPArray(2) & "." & IPArray(3)
End Function

Public Function ConnectToServer(ByVal ServerIP As String, ByVal ServerPort _
As Long) As Long
  Dim hSock As Long, Retval As Long, ServerAddr As SOCKADDR
 
  hSock = socket(AF_INET, SOCK_STREAM, 0&)
  If hSock = -1 Then
    ConnectToServer = -1
    Exit Function
  End If
 
  With ServerAddr
    .sin_addr = inet_addr(ServerIP)
    .sin_port = htons(ServerPort)
    .sin_family = AF_INET
  End With
  Retval = connect(hSock, ServerAddr, Len(ServerAddr))
  If Retval <> 0 Then
    Call closesocket(hSock)
    ConnectToServer = -1
    Exit Function
  End If
 
  Retval = ioctlsocket(hSock, FIONBIO, 1&)
 
  ConnectToServer = hSock
End Function

Public Function Disconnect(ByRef Sock As Long)
  Call closesocket(hSock)
  Sock = 0
End Function

Public Function SendData(ByVal Sock As Long, ByVal Data As String) As Long
  SendData = send(Sock, ByVal Data, Len(Data), 0&)
End Function

Public Function DataComeIn(ByVal Sock As Long) As Long
  Dim Tmpstr As String * 1
 
  DataComeIn = recv(Sock, ByVal Tmpstr, Len(Tmpstr), MSG_PEEK)
  If DataComeIn = -1 Then
    DataComeIn = WSAGetLastError()
  End If
End Function

Public Function GetData(ByVal Sock As Long) As String
  Dim Tmpstr As String * 4096, Retval As Long
 
  Retval = recv(Sock, ByVal Tmpstr, Len(Tmpstr), 0&)
  If Retval > 1 Then GetData = Left$(Tmpstr, Retval)
End Function
 
