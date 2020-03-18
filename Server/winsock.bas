Attribute VB_Name = "modWinsock"
'    Seyerdin Online - A MMO RPG based on Odyssey Online Classic - In memory of Clay Rance
'    Copyright (C) 2020  Samuel Cook and Eric Robinson
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.

Option Explicit

'windows declares here
'Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Src As Any, ByVal cb&)
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

'WINSOCK DEFINES START HERE

Global Const IPPROTO_TCP = 6
Global Const IPPROTO_UDP = 17

Global Const INADDR_NONE = &HFFFFFFFF
Global Const INADDR_ANY = &H0

Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type
Public Const sockaddr_size = 16
Dim saZero As sockaddr

Global Const WSA_DESCRIPTIONLEN = 256
Global Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1

Global Const WSA_SYS_STATUS_LEN = 128
Global Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1

Type WSADataType
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSA_DescriptionSize
    szSystemStatus As String * WSA_SysStatusSize
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Global Const INVALID_SOCKET = -1
Global Const SOCKET_ERROR = -1

Global Const SOCK_STREAM = 1
Global Const SOCK_DGRAM = 2

'Global Const MAXGETHOSTSTRUCT = 1024

Global Const AF_INET = 2
Global Const PF_INET = 2

Type LingerType
    l_onoff As Integer
    l_linger As Integer
End Type

Global Const SOL_SOCKET = &HFFFF&
'Global Const SO_DEBUG = &H1&         ' 0x0001 turn on debugging info recording
'Global Const SO_ACCEPTCONN = &H2&    ' 0x0002 socket has had listen()
'Global Const SO_REUSEADDR = &H4&     ' 0x0004 allow local address reuse
'Global Const SO_KEEPALIVE = &H8&     ' 0x0008 keep connections alive
'Global Const SO_DONTROUTE = &H10&    ' 0x0010 just use interface addresses
'Global Const SO_BROADCAST = &H20&    ' 0x0020 permit sending of broadcast messages
'Global Const SO_USELOOPBACK = &H40&  ' 0x0040 bypass hardware when possible
Global Const SO_LINGER = &H80&       ' 0x0080 linger on close if data present
'Global Const SO_OOBINLINE = &H100&   ' 0x0100 leave received OOB data in line
'Global Const SO_DONTLINGER = Not SO_LINGER
'' Additional options
'Global Const SO_SNDBUF = &H1001&   ' 0x1001 send buffer size
'Global Const SO_RCVBUF = &H1002&   ' 0z1002 receive buffer size
'Global Const SO_SNDLOWAT = &H1003& ' 0x1003 send low-water mark
'Global Const SO_RCVLOWAT = &H1004& ' 0x1004 receive low-water mark
'Global Const SO_SNDTIMEO = &H1005& ' 0x1005 send timeout
'Global Const SO_RCVTIMEO = &H1006& ' 0x1006 receive timeout
'Global Const SO_ERROR = &H1007&    ' 0x1007 get error status and clear
'Global Const SO_TYPE = &H1008&     ' 0x1008 get socket type
'' TCP options
Global Const TCP_NODELAY = &H1& ' 0x0001

Global Const FD_READ = &H1&
Global Const FD_WRITE = &H2&
'Global Const FD_OOB = &H4&
Global Const FD_ACCEPT = &H8&
Global Const FD_CONNECT = &H10&
Global Const FD_CLOSE = &H20&

'SOCKET FUNCTIONS
Declare Function accept Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, addrlen As Long) As Long
Declare Function bind Lib "wsock32.dll" (ByVal s As Long, addr As sockaddr, ByVal namelen As Long) As Long
Declare Function closesocket Lib "wsock32.dll" (ByVal s As Long) As Long
Declare Function getpeername Lib "wsock32.dll" (ByVal s As Long, sname As sockaddr, namelen As Long) As Long
Declare Function htonl Lib "wsock32.dll" (ByVal hostlong As Long) As Long
Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
Declare Function inet_ntoa Lib "wsock32.dll" (ByVal Inn As Long) As Long
Declare Function listen Lib "wsock32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
Declare Function recv Lib "wsock32.dll" (ByVal s As Long, Buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Declare Function send Lib "wsock32.dll" (ByVal s As Long, Buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Declare Function setsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Declare Function Socket Lib "wsock32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
'WINDOWS EXTENSIONS
Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
Declare Function WSACleanup Lib "wsock32.dll" () As Long
Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long
Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal s As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long


'SOME STUFF I ADDED
Global Const INVALID_PORT = -1  'added by me
Global Const INVALID_PROTO = -1 'added by me

Global WSAStartedUp%    'Flag to keep track of whether winsock WSAStartup wascalled

Public Function SendData(ByVal s&, St As String) As Long
    Dim TheMsg() As Byte
    TheMsg = ""
    TheMsg = StrConv(St, vbFromUnicode)
    If UBound(TheMsg) > -1 Then
        SendData = send(s, TheMsg(0), UBound(TheMsg) + 1, 0)
    End If
End Function
Function Receive(s As Long) As String
    Dim Buf() As Byte, L As Long, St As String
    ReDim Buf(2049) As Byte
    L = recv(s, Buf(0), 2048, 0)
    If L > 0 Then
        ReDim Preserve Buf(L - 1)
        St = StrConv(Buf, vbUnicode)
        Receive = St
    End If
End Function

Sub EndWinsock()
    Dim ret&
    If WSAIsBlocking() Then
        ret = WSACancelBlockingCall()
    End If
    ret = WSACleanup()
    WSAStartedUp = False
End Sub

Function getascip(ByVal Inn As Long) As String
    On Error Resume Next
    Dim lpStr&
    Dim nStr&
    Dim retString$
    
    retString = String(32, 0)
    lpStr = inet_ntoa(Inn)
    If lpStr = 0 Then
        getascip = "255.255.255.255"
        Exit Function
    End If
    nStr = lstrlen(lpStr)
    If nStr > 32 Then nStr = 32
    MemCopy ByVal retString, ByVal lpStr, nStr
    retString = Left(retString, nStr)
    getascip = retString
    If Err Then getascip = "255.255.255.255"
End Function


Function GetPeerAddress(ByVal s&) As String
    Dim ret&
    
    On Error Resume Next
    Dim sa As sockaddr
    ret = getpeername(s, sa, sockaddr_size)
    If ret = 0 Then
        GetPeerAddress = SockaddressToString(sa)
    Else
        GetPeerAddress = ""
    End If
    If Err Then GetPeerAddress = ""
End Function

Function ListenForConnect(ByVal Port&, ByVal HWndToMsg&, Msg As Long) As Long
    Dim s&, dummy&
    Dim SelectOps&
    Dim sockin As sockaddr
    
    sockin = saZero     'zero out the structure
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)
    If sockin.sin_port = INVALID_PORT Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    sockin.sin_addr = htonl(INADDR_ANY)
    If sockin.sin_addr = INADDR_NONE Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    s = Socket(PF_INET, SOCK_STREAM, 0)
    If s < 0 Then
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    If bind(s, sockin, sockaddr_size) Then
        If s > 0 Then
            dummy = closesocket(s)
        End If
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    SelectOps = FD_READ Or FD_WRITE Or FD_CLOSE Or FD_ACCEPT
    If WSAAsyncSelect(s, HWndToMsg, ByVal Msg, ByVal SelectOps) Then
        If s > 0 Then
            dummy = closesocket(s)
        End If
        ListenForConnect = SOCKET_ERROR
        Exit Function
    End If
    
    If listen(s, 1) Then
        If s > 0 Then
            dummy = closesocket(s)
        End If
        ListenForConnect = INVALID_SOCKET
        Exit Function
    End If
    ListenForConnect = s
End Function
'this function should work on 16 and 32 bit systems
Function SockaddressToString(sa As sockaddr) As String
    On Error Resume Next
    SockaddressToString = getascip(sa.sin_addr) '& ":" & ntohs(sa.sin_port)
    If Err Then SockaddressToString = ""
End Function

Function StartWinsock(desc$) As Long
    Dim ret&
    Dim WinsockVers&
    
    Dim wsadStartupData As WSADataType
    'WinsockVers = &H101   'Vers 1.1
    WinsockVers = MAKEWORD(2, 2)
    If WSAStartedUp = False Then
        ret = 1
        ret = WSAStartup(WinsockVers, wsadStartupData)
        If ret = 0 Then
            WSAStartedUp = True
            desc = wsadStartupData.szDescription
        Else
            WSAStartedUp = False
        End If
    End If
    StartWinsock = WSAStartedUp
End Function

Public Function MAKEWORD(ByVal bLow As Byte, ByVal bHigh As Byte) As Integer
    MAKEWORD = Val("&H" & Right("00" & Hex(bHigh), 2) & Right("00" & Hex(bLow), 2))
End Function

