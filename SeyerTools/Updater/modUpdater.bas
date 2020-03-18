Attribute VB_Name = "modUpdater"
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

Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Hook
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = -4
Public lpPrevWndProc As Long
Public gHW As Long

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public Const CS_ERROR_CONNECTING = 0
Public Const CS_SERVER_CLOSED = 1
Public Const CS_FINISHED = 2

Public ClientState As Long
    Public Const STATE_INVALID = 0
    Public Const STATE_WAITING_FOR_VERSIONS = 1
    Public Const STATE_DOWNLOADED_VERSIONS = 2
    Public Const STATE_PARSING_VERSIONS = 3
    Public Const STATE_COMPARE_FILES = 4
    Public Const STATE_START_DOWNLOAD_NEXT_FILE = 5
    Public Const STATE_RECEIVE_NEXT_FILE = 6
    Public Const STATE_DOWNLOADING = 7
    'Public Const STATE_CONNECTED = 4
    'Public Const STATE_DOWNLOADING = 5
    'Public Const STATE_DOWNLOADNEXTFILE = 4

Public CurrentFile As Long
Public Type FileType
    Name As String
    Version As Long
    Size As Long
End Type
Dim RemoteFiles() As FileType
Dim LocalFiles() As FileType
Public VersionType As Long
    Public Const LOCALVERSION = 1
    Public Const REMOTEVERSION = 2

Public ClientSocket As Long, SocketData As String, RemoteVersionData As String
Public ServerIP As String
Public ServerPath As String
Public ServerPort As Long
Public ReceivingFile As Boolean

Public TotalMaxDownloadSize As Long
Public CurrentMaxDownloadSize As Long
Public TotalFileDownloadSize As Long
Public CurrentFileDownloadSize As Long

Public blnEnd As Boolean

Sub Main()
Dim St As String, TempArray() As String, TempArray2() As String
Dim A As Long, B As Long, C As Long, D As Long

On Error Resume Next
MkDir "Data"
MkDir "Data\Graphics"
MkDir "Data\Graphics\Portraits"
MkDir "Data\Graphics\Misc"
MkDir "Data\Graphics\Misc\Dungeons"
MkDir "Data\Graphics\Misc\Fog"
MkDir "Data\Graphics\Interface"
MkDir "Data\Graphics\Particles"
MkDir "Data\Sound"
MkDir "Data\Music"
MkDir "Data\Cache"
On Error GoTo SkipUpdate
ChDir (App.Path)

Load frmMain
frmMain.Show

LoadLocalVersions

gHW = frmMain.hwnd
'Hook form to receive callbacks
Hook

'Start Winsock
StartWinsock St

ServerIP = ReadStr("Updater", "Server")
ServerPort = ReadStr("Updater", "Port")
ServerPath = ReadStr("Updater", "Path")
'PrintChat "Connecting to Update Server.", 14

Sleep (100)
DoEvents

frmMain.lblStatus = "Connecting to Update Server."
CurrentFile = -1
ConnectClient

Dim connectionTimeout As Long
connectionTimeout = 0
Do
BreakCase:
    DoEvents
    Sleep 15
        
    Select Case ClientState
        Case 0
            connectionTimeout = connectionTimeout + 1
            If connectionTimeout = 2000 Then
                MsgBox ("The updater has failed to connect to the update server, please check your connection to the internet.  This failure could also be due to firewall permissions.  The update process will be skipped, your files may be outdated.")
                blnEnd = True
                Shell "Seyerdin.exe -update", vbNormalFocus
            End If
        Case STATE_DOWNLOADED_VERSIONS
            'PrintChat "File List Downloaded", 15
            'PrintChat "Checking for updated files", 15
            frmMain.lblStatus = "Checking for updated files"
            A = InStr(1, SocketData, vbCrLf & vbCrLf)
            If A Then
                St = Mid$(SocketData, A + 4, Len(SocketData) - A)
                SocketData = St
            Else
                'PrintChat "Invalid File Versions!", 14
                frmMain.lblStatus = "Invalid File Version!"
                closesocket ClientSocket
                blnEnd = True
            End If
            ClientState = STATE_PARSING_VERSIONS
        Case STATE_PARSING_VERSIONS
            RemoteVersionData = CStr(SocketData)
            
            SocketData = Replace(SocketData, vbCr, "")
            SocketData = Replace(SocketData, vbLf, "")
            
            
            
            TempArray = Split(SocketData, "|")
            ReDim RemoteFiles(UBound(TempArray()))
            For C = 0 To UBound(TempArray())
                TempArray2 = Split(TempArray(C), ",")
                RemoteFiles(C).Name = TempArray2(0)
                RemoteFiles(C).Version = Val(TempArray2(1))
                RemoteFiles(C).Size = Val(TempArray2(2))
            Next C
            CurrentFile = 0
            'Start out by going through and getting the max size of all the files that need updated
            For A = 0 To UBound(RemoteFiles())
                C = 1
                For B = 0 To UBound(LocalFiles())
                    If RemoteFiles(A).Name = LocalFiles(B).Name Then
                        C = 0
                        If RemoteFiles(A).Version <> LocalFiles(B).Version Then
                            C = 1
                            Exit For
                        End If
                        If Exists(RemoteFiles(A).Name) Then
                            D = FreeFile
                            Open RemoteFiles(A).Name For Binary As #D
                                If RemoteFiles(A).Size <> LOF(D) Then
                                    C = 1
                                    Close #1
                                    Exit For
                                End If
                            Close #1
                        Else
                            C = 1
                        End If
                        Exit For
                    End If
                Next B
                If C = 1 Then
                    'No files match
                    TotalMaxDownloadSize = TotalMaxDownloadSize + RemoteFiles(A).Size
                End If
                If Exists(RemoteFiles(A).Name & ".tmp") Then
                    Kill RemoteFiles(A).Name & ".tmp"
                End If
            Next A
            ClientState = STATE_COMPARE_FILES
        Case STATE_COMPARE_FILES
            If CurrentFile = -1 Then CurrentFile = 0
            For A = CurrentFile To UBound(RemoteFiles())
                C = 1
                If Exists(RemoteFiles(A).Name) Then
                    For B = 0 To UBound(LocalFiles())
                        If RemoteFiles(A).Name = LocalFiles(B).Name Then
                            C = 0
                            If RemoteFiles(A).Version <> LocalFiles(B).Version Then
                                C = 1
                            End If
                            D = FreeFile
                            Open RemoteFiles(A).Name For Binary As #D
                                If RemoteFiles(A).Size <> LOF(D) Then
                                    C = 1
                                End If
                            Close #D
                            Exit For
                        End If
                    Next B
                Else
                    C = 1
                End If
                If C = 1 Then
                    'PrintChat "Updating " & RemoteFiles(A).Name, 15
                    frmMain.lblStatus = "Updating " & RemoteFiles(A).Name
                    CurrentFile = A
                    ClientState = STATE_START_DOWNLOAD_NEXT_FILE
                    GoTo BreakCase
                End If
            Next A
            ClientState = STATE_INVALID
            'PrintChat "All files updated.", 15
            frmMain.lblStatus = "All files up to date."
            frmMain.picProgress.Width = frmMain.picCurrent.Width
            blnEnd = True
            Shell "Seyerdin.exe -update", vbNormalFocus
            On Error GoTo SkipUpdate
        Case STATE_START_DOWNLOAD_NEXT_FILE
            ConnectClient
            ClientState = STATE_INVALID
        Case STATE_RECEIVE_NEXT_FILE
            B = FreeFile
            
            Open RemoteFiles(CurrentFile).Name & ".tmp" For Binary As #B
                D = LOF(B)
                If RemoteFiles(CurrentFile).Size <> D Then
                    'PrintChat RemoteFiles(CurrentFile).Name & " was not updated. (Invalid Size)", 15
                    frmMain.lblStatus = "Invalid file size for " & RemoteFiles(CurrentFile).Name
                    ClientState = STATE_INVALID
                End If
            Close #B
            On Error Resume Next
            
            If ClientState <> STATE_INVALID Then
                A = 0
                Do
                    DoEvents
                    A = A + 1
                    Err.Clear
                    FileCopy RemoteFiles(CurrentFile).Name & ".tmp", RemoteFiles(CurrentFile).Name

                Loop While Err.Number > 0 And A < 1000
                
                If UCase$(Right$(RemoteFiles(CurrentFile).Name, 3)) = "DLL" Then
                    'Shell "regsvr32 /s " & RemoteFiles(CurrentFile).Name
                End If
                On Error GoTo SkipUpdate
                UpdateLocalFiles CurrentFile
                'If CurrentFile <= UBound(LocalFiles) Then
                '    LocalFiles(CurrentFile).Version = RemoteFiles(CurrentFile).Version
                '    LocalFiles(CurrentFile).Size = RemoteFiles(CurrentFile).Size
                'End If
                'If ClientState <> STATE_INVALID Then PrintChat RemoteFiles(CurrentFile).Name & " has been updated!", 14
            End If
            On Error Resume Next
            Kill RemoteFiles(CurrentFile).Name & ".tmp"
            On Error GoTo SkipUpdate
            CurrentFile = CurrentFile + 1
            ClientState = STATE_COMPARE_FILES
        Case STATE_DOWNLOADING
            DoEvents
    End Select
Loop While blnEnd = False

ENDIT:
Unhook
EndWinsock
End

SkipUpdate:
MsgBox ("There was an error in the update process, potentially resulting in files not being updated.  The update process will be skipped, some files may be outdated or missing.")
blnEnd = True
Shell "Seyerdin.exe -update", vbNormalFocus
Unhook
EndWinsock
End

End Sub

Function ReadStr(lpAppName, lpKeyName As String, Optional Filename As String = "Updater") As String
    Dim lpReturnedString As String, Valid As Long
    lpReturnedString = Space$(256)
    Valid = GetPrivateProfileString&(lpAppName, lpKeyName, "", lpReturnedString, 256, App.Path + "\Data\Cache\" + Filename + ".ini")
    ReadStr = Left$(lpReturnedString, Valid)
End Function

Public Sub Hook()
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub Unhook()
    SetWindowLong gHW, GWL_WNDPROC, lpPrevWndProc
End Sub

Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim St1 As String, A As Long, B As Byte, C As Long, D As Long
    If uMsg = 1025 Then
        'Client Socket
        Select Case lParam And 255
            Case FD_CLOSE
'                CloseClientSocket CS_SERVER_CLOSED
                Select Case ClientState
                    Case STATE_WAITING_FOR_VERSIONS
                        ClientState = STATE_DOWNLOADED_VERSIONS
                End Select
                If CurrentFile >= 0 Then
                    ClientState = STATE_RECEIVE_NEXT_FILE
                End If
            Case FD_CONNECT
                SocketData = ""
                If lParam = FD_CONNECT Then
                    If CurrentFile = -1 Then
                        'PrintChat "Connected to server. . .", 14
                        frmMain.lblStatus = "Connected to server."
                        ClientState = STATE_WAITING_FOR_VERSIONS
                        
                        
                        SendSocket "GET " & ServerPath & "/FileVersions.txt HTTP/1.0" & vbCrLf & _
                        "Accept: */*" & vbCrLf & _
                        "Accept-Language: en -us" & vbCrLf & _
                        "User-Agent: Mozilla/4.0" & vbCrLf & _
                        "Host: " & ServerIP & vbCrLf & vbCrLf
                    ElseIf CurrentFile >= 0 Then
                        CurrentFileDownloadSize = 0
                        TotalFileDownloadSize = RemoteFiles(CurrentFile).Size
                        DrawUpdateBars
                        ClientState = STATE_DOWNLOADING
                        ReceivingFile = False
                        
                        'MsgBox "GET " & ServerPath & "/" & RemoteFiles(CurrentFile).Name & " HTTP/1.0" & vbCrLf & _
                        "Accept: */*" & vbCrLf & _
                        "Accept-Language: en -us" & vbCrLf & _
                        "User-Agent: Mozilla/4.0" & vbCrLf & _
                        "Host: " & ServerIP & vbCrLf & vbCrLf
                        
                        
                        SendSocket "GET " & ServerPath & "/" & RemoteFiles(CurrentFile).Name & " HTTP/1.0" & vbCrLf & _
                        "Accept: */*" & vbCrLf & _
                        "Accept-Language: en -us" & vbCrLf & _
                        "User-Agent: Mozilla/4.0" & vbCrLf & _
                        "Host: " & ServerIP & vbCrLf & vbCrLf
                    End If
                Else
                    CloseClientSocket CS_ERROR_CONNECTING
                    WaitForConnect "Error Connecting - Waiting"
                End If
            Case FD_READ
                If lParam = FD_READ Then ReceiveData
        End Select
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function

Sub ConnectClient()
    ClientSocket = ConnectSock(ServerIP, ServerPort, "", gHW, True)
End Sub

Sub CloseClientSocket(Action As Byte)
    closesocket ClientSocket
    ClientSocket = INVALID_SOCKET
    'ClientState = STATE_INVALID

    Select Case Action
        Case CS_ERROR_CONNECTING 'Error connecting to server
            'PrintChat "There has been an error connecting to the server!", 14
            frmMain.lblStatus = "There has been an error connecting to the server!"
        Case CS_SERVER_CLOSED
            'PrintChat "Update server has closed the connection!", 14
            frmMain.lblStatus = "Update server has closed the connection!"
        Case CS_FINISHED
            'Show status and run game
            'PrintChat "Seyerdin Online has been successfully updated!", 14
            frmMain.lblStatus = "Seyerdin Online has been successfully updated!"
        Case Else
            frmMain.Show
    End Select
End Sub

Sub WaitForConnect(Message As String)
    'PrintChat Message, 14
    frmMain.lblStatus = Message
End Sub


'Public Sub PrintChat(ByVal St As String, Color As Byte)
'    Dim A As Long, B As Long, FoundLine As Boolean
'    Dim Text As String, TextHeight As Long, TextWidth As Long
'
'    With frmMain.picStatusText
'        .ForeColor = QBColor(Color)
'        TextHeight = .TextHeight("A")
'        MoveUp
'        While St <> ""
'            B = 0
'            FoundLine = False
'            For A = 1 To Len(St)
'                If .TextWidth(Left$(St, A)) > .ScaleWidth - .CurrentX Then
'                    FoundLine = True
'                    If B = 0 Then
'                        B = A - 1
'                    End If
'                    If B > 0 Then
'                        Text = Left$(St, B)
'                        St = Mid$(St, B + 1)
'                    Else
'                        Text = ""
'                    End If
'                    Exit For
'                End If
'                If Mid$(St, A, 1) = " " Then B = A
'            Next A
'            If FoundLine = False Then
'                Text = St
'               St = ""
'            End If
'            If Text <> "" Then
'                TextWidth = .TextWidth(Text)
'                TextOut .hdc, .CurrentX, .ScaleHeight - TextHeight, Text, Len(Text)
'                If FoundLine = True Then
'                    MoveUp
'                Else
'                    .CurrentX = .CurrentX + TextWidth
'                End If
'            Else
'                If St <> "" Then
'                    MoveUp
'                End If
'            End If
'        Wend
'    End With
'    frmMain.picStatusText.Refresh
'End Sub
'Sub MoveUp()
'    Dim TextHeight As Long
'    Dim A As Long
'    With frmMain.picStatusText
'        A = .TextHeight("A")
'        .CurrentX = 0
'        BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight - A, .hdc, 0, A, vbSrcCopy
'        BitBlt .hdc, 0, .ScaleHeight - A, .ScaleWidth, A, 0, 0, 0, 0
'    End With
'End Sub

Sub SendSocket(ByVal St As String)
    If SendData(ClientSocket, St) = SOCKET_ERROR Then
        CloseClientSocket 0
    End If
End Sub

Sub ReceiveData()
Dim St As String
Dim A As Long, B As Long, C As Long
St = Receive(ClientSocket)
SocketData = SocketData & St

'If CurrentFile > 0 Then
'    A = TotalFileDownloadSize - CurrentFileDownloadSize
'    CurrentFileDownloadSize = CurrentFileDownloadSize + Len(St)
'    If CurrentFileDownloadSize >= TotalFileDownloadSize Then
'        CurrentMaxDownloadSize = CurrentMaxDownloadSize + A
'        CurrentFileDownloadSize = TotalFileDownloadSize
'    Else
'        CurrentMaxDownloadSize = CurrentMaxDownloadSize + Len(St)
'    End If
'    DrawUpdateBars
'End If

If ClientState = STATE_DOWNLOADING Or ClientState = STATE_RECEIVE_NEXT_FILE Then
    If ReceivingFile Then
    
badcodingpractice:
        B = FreeFile
        'MsgBox SocketData
        Open RemoteFiles(CurrentFile).Name & ".tmp" For Binary As #B
            C = LOF(B) + 1
            Put #1, C, SocketData
        Close #1
        CurrentFileDownloadSize = CurrentFileDownloadSize + Len(SocketData)
        CurrentMaxDownloadSize = CurrentMaxDownloadSize + Len(SocketData)
        SocketData = ""
        DrawUpdateBars
    Else
        A = InStr(1, SocketData, vbCrLf + vbCrLf)
        If A Then
            ReceivingFile = True
            St = Mid$(SocketData, A + 4, Len(SocketData) - A)
            SocketData = St
            GoTo badcodingpractice
        End If
    End If
End If

End Sub

Public Function AppPath() As String
Dim Path As String
Path = App.Path
If Right(Path, 1) = "/" Or Right(Path, 1) = "\" Then
    AppPath = Path
Else
    AppPath = Path & "\"
End If
End Function

Sub LoadLocalVersions()
Dim St As String, FileNum As Long, A As Long, B As Long, C As Long
Dim TempArray() As String, TempArray2() As String
    FileNum = FreeFile
    If Exists(AppPath & "Data/Cache/FileVersions.txt") Then
        Open AppPath & "Data/Cache/FileVersions.txt" For Binary As #FileNum
            St = Space(LOF(FileNum))
            Get #FileNum, , St
        Close #FileNum
        St = Replace(St, vbCr, "")
        St = Replace(St, vbLf, "")
        TempArray = Split(St, "|")
        ReDim LocalFiles(UBound(TempArray()))
        For C = 0 To UBound(TempArray())
            If Len(TempArray(C)) > 0 Then
                TempArray2 = Split(TempArray(C), ",")
                LocalFiles(C).Name = TempArray2(0)
                LocalFiles(C).Version = Val(TempArray2(1))
                LocalFiles(C).Size = Val(TempArray2(2))
            End If
        Next C
    End If
End Sub

Sub DrawUpdateBars()
DoEvents
Dim A As Double, B As Double
'If TotalFileDownloadSize > 0 Then
'    A = CDbl(CurrentFileDownloadSize) / CDbl(TotalFileDownloadSize)
'    If A > 100 Then A = 100
'End If
'If TotalMaxDownloadSize > 0 Then
'    B = CDbl(CurrentMaxDownloadSize) / CDbl(TotalMaxDownloadSize)
'    If B > 100 Then B = 100
'End If
'    With frmMain
'        .picCurrent.Width = 324 * A
'        .picTotal.Width = 324 * B
'    End With
If TotalFileDownloadSize > 0 Then
    A = CDbl(CurrentFileDownloadSize) / CDbl(TotalFileDownloadSize)
    If A > 100 Then A = 100
    frmMain.picProgress.Width = (315 * A) + 1
End If
End Sub


Public Function Exists(Filename As String) As Boolean
On Error Resume Next
Open Filename For Input As #1
Close #1
If Err.Number <> 0 Then
   Exists = False
Else
   Exists = True
End If
End Function

Public Function UpdateLocalFiles(CurrentFile As Long)
    Dim St As String
    Dim A As Long, Found As Long
    
    St = RemoteFiles(CurrentFile).Name
    For A = 0 To UBound(LocalFiles())
        If LocalFiles(A).Name = St Then
            LocalFiles(A).Version = RemoteFiles(CurrentFile).Version
            LocalFiles(A).Size = RemoteFiles(CurrentFile).Size
            Found = 1
            Exit For
        End If
    Next A
    If Found = 0 Then
        ReDim Preserve LocalFiles(0 To UBound(LocalFiles) + 1)
        LocalFiles(UBound(LocalFiles)).Name = RemoteFiles(CurrentFile).Name
        LocalFiles(UBound(LocalFiles)).Size = RemoteFiles(CurrentFile).Size
        LocalFiles(UBound(LocalFiles)).Version = RemoteFiles(CurrentFile).Version
    End If
    
    St = ""
    For A = 0 To UBound(LocalFiles())
        If LocalFiles(A).Name <> "" Then
            St = St & LocalFiles(A).Name & "," & LocalFiles(A).Version & "," & LocalFiles(A).Size & IIf(A = UBound(LocalFiles()), "", "|") & vbCrLf
        End If
    Next A
    
    On Error Resume Next
    Kill "/Data/Cache/FileVersions.txt"

    A = FreeFile
    Open AppPath & "/Data/Cache/FileVersions.txt" For Binary As #A
        Put #A, , St
    Close #A

End Function
