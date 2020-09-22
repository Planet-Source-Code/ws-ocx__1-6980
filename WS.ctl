VERSION 5.00
Begin VB.UserControl WS 
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   675
   Picture         =   "WS.ctx":0000
   ScaleHeight     =   570
   ScaleWidth      =   675
End
Attribute VB_Name = "WS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () _
        As Long
        
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal _
        wVersionRequired&, lpWSAData As WinSocketDataType) _
        As Long
        
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () _
        As Long
        
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal _
        HostName$, ByVal HostLen%) As Long
        
Private Declare Function gethostbyname Lib "WSOCK32.DLL" _
        (ByVal HostName$) As Long
        
Private Declare Function gethostbyaddr Lib "WSOCK32.DLL" _
        (ByVal addr$, ByVal laenge%, ByVal typ%) As Long
        
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As _
        Any, ByVal hpvSource&, ByVal cbCopy&)

Const WS_VERSION_REQD = &H101
Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&

Const MIN_SOCKETS_REQD = 1
Const SOCKET_ERROR = -1
Const WSADescription_Len = 256
Const WSASYS_Status_Len = 128


Private Type HostDeType
  hName As Long
  hAliases As Long
  hAddrType As Integer
  hLength As Integer
  hAddrList As Long
End Type

Private Type WinSocketDataType
   wversion As Integer
   wHighVersion As Integer
   szDescription(0 To WSADescription_Len) As Byte
   szSystemStatus(0 To WSASYS_Status_Len) As Byte
   iMaxSockets As Integer
   iMaxUdpDg As Integer
   lpszVendorInfo As Long
End Type
Dim IP As String
Dim Domain As String


Public Function getIPof(Domain$)
 Dim aa$
  InitSockets
  aa = HostByName$(Domain$)
  If aa = "" Then aa = "Nicht gefunden"
  CleanSockets
  getIPof = aa
End Function

Public Function getHOSTNAMEof(IP$)
Dim aa$
  InitSockets
  aa = HostByAddress(IP$)
  If aa = "" Then aa = "Nicht gefunden"
  CleanSockets
  getHOSTNAMEof = aa
End Function

Private Function HostByAddress(ByVal Addresse$) As String
  Dim X%
  Dim HostDeAddress&
  Dim aa$, BB As String * 5
  Dim HOST As HostDeType
  
    aa = Chr$(Val(NextChar(Addresse, ".")))
    aa = aa + Chr$(Val(NextChar(Addresse, ".")))
    aa = aa + Chr$(Val(NextChar(Addresse, ".")))
    aa = aa + Chr$(Val(Addresse))

    HostDeAddress = gethostbyaddr(aa, Len(aa), 2)
    If HostDeAddress = 0 Then
      HostByAddress = ""
      Exit Function
    End If
    
    Call RtlMoveMemory(HOST, HostDeAddress, LenB(HOST))
 
    aa = ""
    X = 0
    Do
       Call RtlMoveMemory(ByVal BB, HOST.hName + X, 1)
       If Left$(BB, 1) = Chr$(0) Then Exit Do
       aa = aa + Left$(BB, 1)
       X = X + 1
    Loop
    
    HostByAddress = aa
End Function

Private Function HostByName(Name$, Optional X% = 0) As String
  Dim MemIp() As Byte
  Dim Y%
  Dim HostDeAddress&, HostIp&
  Dim IpAddress$
  Dim HOST As HostDeType
  
    HostDeAddress = gethostbyname(Name)
    If HostDeAddress = 0 Then
      HostByName = ""
      Exit Function
    End If
    
    Call RtlMoveMemory(HOST, HostDeAddress, LenB(HOST))
    
    For Y = 0 To X
      Call RtlMoveMemory(HostIp, HOST.hAddrList + 4 * Y, 4)
      If HostIp = 0 Then
        HostByName = ""
        Exit Function
      End If
    Next Y
    
    ReDim MemIp(1 To HOST.hLength)
    Call RtlMoveMemory(MemIp(1), HostIp, HOST.hLength)
    
    IpAddress = ""
    
    For Y = 1 To HOST.hLength
      IpAddress = IpAddress & MemIp(Y) & "."
    Next Y
    
    IpAddress = Left$(IpAddress, Len(IpAddress) - 1)
    HostByName = IpAddress
End Function

Private Function MyHostName() As String
  Dim HostName As String * 256
  
    If gethostname(HostName, 256) = SOCKET_ERROR Then
      Exit Function
    Else
      MyHostName = NextChar(Trim$(HostName), Chr$(0))
    End If
End Function

Private Sub InitSockets()
  Dim Result%
  Dim LoBy%, HiBy%
  Dim SocketData As WinSocketDataType
  
    Result = WSAStartup(WS_VERSION_REQD, SocketData)
    If Result <> 0 Then
      
    End If
    
    LoBy = SocketData.wversion And &HFF&
    HiBy = SocketData.wversion \ &H100 And &HFF&
    
    If LoBy < WS_VERSION_MAJOR Or LoBy = WS_VERSION_MAJOR And _
       HiBy < WS_VERSION_MINOR Then
        
        
    End If
    
    If SocketData.iMaxSockets < MIN_SOCKETS_REQD Then
      
      
    End If
End Sub

Private Sub CleanSockets()
  Dim Result&
  
    Result = WSACleanup()
    If Result <> 0 Then
      
      
    End If
End Sub

Private Function NextChar(Text$, Char$) As String
  Dim POS%
    POS = InStr(1, Text, Char)
    If POS = 0 Then
      NextChar = Text
      Text = ""
    Else
      NextChar = Left$(Text, POS - 1)
      Text = Mid$(Text, POS + Len(Char))
    End If

End Function


