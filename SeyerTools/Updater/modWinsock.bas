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

'Winsock Declarations
Declare Function WSAStartup Lib "WS2_32.DLL" (ByVal wVR As Long, lpWSAD As WSADataType) As Long
Declare Function socket Lib "WS2_32.DLL" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
Declare Function htons Lib "WS2_32.DLL" (ByVal hostshort As Long) As Integer
Declare Function ntohs Lib "WS2_32.DLL" (ByVal netshort As Long) As Integer
Declare Function connect Lib "WS2_32.DLL" (ByVal S As Long, addr As sockaddr, ByVal namelen As Long) As Long
Declare Function closesocket Lib "WS2_32.DLL" (ByVal S As Long) As Long
Declare Function WSAAsyncSelect Lib "WS2_32.DLL" (ByVal S As Long, ByVal hwnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Declare Function inet_addr Lib "WS2_32.DLL" (ByVal cp As String) As Long
Declare Function gethostbyname Lib "WS2_32.DLL" (ByVal host_name As String) As Long
Declare Function inet_ntoa Lib "WS2_32.DLL" (ByVal inn As Long) As Long
Declare Function WSAIsBlocking Lib "WS2_32.DLL" () As Long
Declare Function WSACleanup Lib "WS2_32.DLL" () As Long
Declare Function WSACancelBlockingCall Lib "WS2_32.DLL" () As Long
Declare Function recv Lib "WS2_32.DLL" (ByVal S As Long, Buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Declare Function send Lib "WS2_32.DLL" (ByVal S As Long, Buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long

'Windows Declarations
Declare Sub MemCopy Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

Global Const WSA_DESCRIPTIONLEN = 256
Global Const WSA_DescriptionSize = WSA_DESCRIPTIONLEN + 1
Global Const WSA_SYS_STATUS_LEN = 128
Global Const WSA_SysStatusSize = WSA_SYS_STATUS_LEN + 1

Const sockaddr_size = 16
Type sockaddr
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type
Dim saZero As sockaddr

Type HostEnt
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Const hostent_size = 16

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
Global Const INVALID_PORT = -1  'added by me
Global Const SOCKET_ERROR = -1
Global Const INADDR_NONE = &HFFFFFFFF
Global Const INADDR_ANY = &H0


Global Const FD_READ = &H1&
Global Const FD_WRITE = &H2&
Global Const FD_OOB = &H4&
Global Const FD_ACCEPT = &H8&
Global Const FD_CONNECT = &H10&
Global Const FD_CLOSE = &H20


Global Const AF_INET = 2
Global Const PF_INET = 2

Global Const IPPROTO_TCP = 6
Global Const IPPROTO_UDP = 17

Global Const SOCK_STREAM = 1

Global WSAStartedUp%    'Flag to keep track of whether winsock WSAStartup wascalled
Global SockReadBuffer$

Function StartWinsock(desc$) As Long
    Dim ret&
    Dim WinsockVers&
    
    Dim wsadStartupData As WSADataType
    WinsockVers = &H101   'Vers 1.1
    
    If WSAStartedUp = False Then
        ret = 1
        ret = WSAStartup(WinsockVers, wsadStartupData)
        If ret = 0 Then
            WSAStartedUp = True
            'Debug.Print "wVersion="; VBntoaVers(wsadStartupData.wVersion), "wHighVersion="; VBntoaVers(wsadStartupData.wHighVersion)
            'Debug.Print "szDescription="; wsadStartupData.szDescription
            'Debug.Print "szSystemStatus="; wsadStartupData.szSystemStatus
            'Debug.Print "iMaxSockets="; wsadStartupData.iMaxSockets, "iMaxUdpDg="; wsadStartupData.iMaxUdpDg
            desc = wsadStartupData.szDescription
        Else
            WSAStartedUp = False
        End If
    End If
    StartWinsock = WSAStartedUp
End Function

Function ConnectSock(ByVal host$, ByVal Port&, retIpPort$, ByVal HWndToMsg&, ByVal Async%) As Long
    Dim S&, SelectOps&, dummy&
    Dim sockin As sockaddr
    
    SockReadBuffer$ = ""
    sockin = saZero
    sockin.sin_family = AF_INET
    sockin.sin_port = htons(Port)
    If sockin.sin_port = INVALID_PORT Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If

    sockin.sin_addr = GetHostByNameAlias(host$)
    If sockin.sin_addr = INADDR_NONE Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    retIpPort$ = getascip$(sockin.sin_addr) & ":" & ntohs(sockin.sin_port)

    S = socket(PF_INET, SOCK_STREAM, IPPROTO_TCP)
    If S < 0 Then
        ConnectSock = INVALID_SOCKET
        Exit Function
    End If
    If Not Async Then
        If connect(S, sockin, sockaddr_size) <> 0 Then
            If S > 0 Then
                dummy = closesocket(S)
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
        SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
        If WSAAsyncSelect(S, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
            If S > 0 Then
                dummy = closesocket(S)
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
    Else
        SelectOps = FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE
        If WSAAsyncSelect(S, HWndToMsg, ByVal 1025, ByVal SelectOps) Then
            If S > 0 Then
                dummy = closesocket(S)
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
        If connect(S, sockin, sockaddr_size) <> -1 Then
            If S > 0 Then
                dummy = closesocket(S)
            End If
            ConnectSock = INVALID_SOCKET
            Exit Function
        End If
    End If
    ConnectSock = S
End Function

Function GetHostByNameAlias(ByVal hostname$) As Long
    On Error Resume Next
    'Return IP address as a long, in network byte order

    Dim phe&    ' pointer to host information entry
    Dim heDestHost As HostEnt 'hostent structure
    Dim addrList&
    Dim retIP&
    'first check to see if what we have been passed is a valid IP
    retIP = inet_addr(hostname)
    If retIP = INADDR_NONE Then
        'it wasn't an IP, so do a DNS lookup
        phe = gethostbyname(hostname)
        If phe <> 0 Then
            'Pointer is non-null, so copy in hostent structure
            MemCopy heDestHost, ByVal phe, hostent_size
            'Now get first pointer in address list
            MemCopy addrList, ByVal heDestHost.h_addr_list, 4
            MemCopy retIP, ByVal addrList, heDestHost.h_length
        Else
            'its not a valid address
            retIP = INADDR_NONE
        End If
    End If
    GetHostByNameAlias = retIP
    If Err Then GetHostByNameAlias = INADDR_NONE
End Function

Function getascip(ByVal inn As Long) As String
    On Error Resume Next
    Dim lpStr&
    Dim nStr&
    Dim retString$
    
    retString = String(32, 0)
    lpStr = inet_ntoa(inn)
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

Sub EndWinsock()
    Dim ret&
    If WSAIsBlocking() Then
        ret = WSACancelBlockingCall()
    End If
    ret = WSACleanup()
    WSAStartedUp = False
End Sub

Public Function SendData(ByVal S&, St As String) As Long
    Dim TheMsg() As Byte
    TheMsg = ""
    TheMsg = StrConv(St, vbFromUnicode)
    If UBound(TheMsg) > -1 Then
        SendData = send(S, TheMsg(0), UBound(TheMsg) + 1, 0)
    End If
End Function

Function Receive(S As Long) As String
    Dim Buf() As Byte, L As Long
    ReDim Buf(2049) As Byte
    L = recv(S, Buf(0), 2048, 0)
    If L > 0 Then
        ReDim Preserve Buf(L - 1)
        Receive = StrConv(Buf, vbUnicode)
    End If
End Function
