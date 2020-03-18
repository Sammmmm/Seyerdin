Attribute VB_Name = "modMapper"
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

Public Const TitleString = "Seyerdin Online Mapping Utility"

Public Const ClientVer = "56"
Public Const MaxMapWidth = 500
Public Const MaxMapHeight = 500

#Const DEBUGWINDOWPROC = 0
#Const USEGETPROP = 0

#If DEBUGWINDOWPROC Then
Private m_SCHook As WindowProcHook
#End If

'INI File Related
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long

'SetBkMode Constants
Public Const TRANSPARENT = 1

'Stretch Modes
Public Const HALFTONE = 4

'BitBlt Constants
Public Const BLACKNESS = &H42
Public Const WHITENESS = &HFF0062
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const SRCAND = &H8800C6
Public Const NOTSRCCOPY = &H330008
Public Const SRCINVERT = &H660046
Public Const DSTINVERT = &H550009

'DrawText Constants
Public Const DT_TOP = &H0
Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_BOTTOM = &H8
Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20
Public Const DT_EXPANDTABS = &H40
Public Const DT_TABSTOP = &H80
Public Const DT_NOCLIP = &H100
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_CALCRECT = &H400
Public Const DT_NOPREFIX = &H800
Public Const DT_INTERNAL = &H1000

'Hook
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = -4
Public lpPrevWndProc As Long
Public gHW As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Type MapStartLocationData
    Map As Integer
    X As Byte
    Y As Byte
End Type

Type TileData
    Ground As Integer
    Ground2 As Integer
    BGTile1 As Integer
    BGTile2 As Integer
    FGTile As Integer
    Att As Byte
    AttData(0 To 3) As Byte
End Type

Type MapMonsterSpawnData
    Monster As Byte
    Rate As Byte
End Type

Type MapData
    Name As String
    ExitUp As Integer
    ExitDown As Integer
    ExitLeft As Integer
    ExitRight As Integer
    Tile(0 To 11, 0 To 11) As TileData
    MonsterSpawn(0 To 2) As MapMonsterSpawnData
    BootLocation As MapStartLocationData
    NPC As Byte
    MIDI As Byte
    flags As Byte
    Version As Long
End Type

Type MapGridData
    MapNum As Integer
    Received As Boolean
End Type

Public MapGrid(1 To MaxMapWidth, 1 To MaxMapHeight) As MapGridData

Public Map(1 To 5000) As MapData

Public ClientSocket As Long
Public CPacketsSent As Long

Public frmWait_Loaded As Boolean
Public frmLogin_Loaded As Boolean
Public frmMain_Loaded As Boolean

Public hdcMap As Long

'Public hdcTiles As Long, hbmpTiles As Long, obmpTiles As Long
'Public hdcTilesMask As Long, hbmpTilesMask As Long, obmpTilesMask As Long
'
'Public hbmpBuffer As Long, obmpBuffer As Long
'Public hdcBackup As Long, hbmpBackup As Long, obmpBackup As Long

Public blnEnd As Boolean

Public User As String, Pass As String

Public ServerIP As String

Public SocketData As String

Public MapNum As Long, MapX As Long, MapY As Long
Public MapData As String * 2374

Public blnDone As Boolean, blnLowQuality
Public blnStop As Boolean

Public SelectedX As Long, SelectedY As Long

Public MapWidth As Long, MapHeight As Long
Public FoundError As Boolean

Public SentPacketCount As Long
Sub DrawMap()
    Dim A As Long, B As Long, X As Long, Y As Long
    Dim R1 As RECT, R2 As RECT
    R1.Left = 0: R1.Top = 0
    R1.Bottom = 384: R1.Right = 384
    sfcBuffer.Surface.BltColorFill R1, 0
    
    For X = 0 To 11
        For Y = 0 To 11
            With Map(MapNum).Tile(X, Y)
                R1.Left = X * 32: R1.Right = R1.Left + 32: R1.Top = Y * 32: R1.Bottom = R1.Top + 32
                If .Ground > 0 Then
                    'BitBlt hdcBuffer, X * 32, Y * 32, 32, 32, hdcTiles, ((.Ground - 1) Mod 7) * 32, Int((.Ground - 1) / 7) * 32, SRCCOPY
                    R2.Left = ((.Ground - 1) Mod 7) * 32:          R2.Right = R2.Left + 32
                    R2.Top = Int((.Ground - 1) / 7) * 32:          R2.Bottom = R2.Top + 32
                    sfcBuffer.Surface.BltFast X * 32, Y * 32, sfcTiles.Surface, R2, DDBLTFAST_SRCCOLORKEY
                End If
                If .Ground2 > 0 Then
                    'TransparentBlt hdcBuffer, X * 32, Y * 32, 32, 32, hdcTiles, ((.Ground2 - 1) Mod 7) * 32, Int((.Ground2 - 1) / 7) * 32, hdcTilesMask
                    R2.Left = ((.Ground2 - 1) Mod 7) * 32:         R2.Right = R2.Left + 32
                    R2.Top = Int((.Ground2 - 1) / 7) * 32:         R2.Bottom = R2.Top + 32
                    sfcBuffer.Surface.BltFast X * 32, Y * 32, sfcTiles.Surface, R2, DDBLTFAST_SRCCOLORKEY
                End If
                If .BGTile1 > 0 Then
                    'TransparentBlt hdcBuffer, X * 32, Y * 32, 32, 32, hdcTiles, ((.BGTile1 - 1) Mod 7) * 32, Int((.BGTile1 - 1) / 7) * 32, hdcTilesMask
                    R2.Left = ((.BGTile1 - 1) Mod 7) * 32:         R2.Right = R2.Left + 32
                    R2.Top = Int((.BGTile1 - 1) / 7) * 32:         R2.Bottom = R2.Top + 32
                    sfcBuffer.Surface.BltFast X * 32, Y * 32, sfcTiles.Surface, R2, DDBLTFAST_SRCCOLORKEY
                End If
                If .FGTile > 0 Then
                    'TransparentBlt hdcBuffer, X * 32, Y * 32, 32, 32, hdcTiles, ((.FGTile - 1) Mod 7) * 32, Int((.FGTile - 1) / 7) * 32, hdcTilesMask
                    R2.Left = ((.FGTile - 1) Mod 7) * 32:          R2.Right = R2.Left + 32
                    R2.Top = Int((.FGTile - 1) / 7) * 32:          R2.Bottom = R2.Top + 32
                    sfcBuffer.Surface.BltFast X * 32, Y * 32, sfcTiles.Surface, R2, DDBLTFAST_SRCCOLORKEY
                End If
            End With
        Next Y
    Next X
End Sub

Sub LoadMap(MapData As String)
    Dim A As Long, X As Long, Y As Long
    If Len(MapData) = 2374 Then
        With Map(MapNum)
            .Name = ClipString$(Mid$(MapData, 1, 30))
            .Version = Asc(Mid$(MapData, 31, 1)) * 16777216 + Asc(Mid$(MapData, 32, 1)) * 65536 + Asc(Mid$(MapData, 33, 1)) * 256& + Asc(Mid$(MapData, 34, 1))
            '.NPC = Asc(Mid$(MapData, 35, 1))
            '.MIDI = Asc(Mid$(MapData, 36, 1))
            .ExitUp = Asc(Mid$(MapData, 37, 1)) * 256 + Asc(Mid$(MapData, 38, 1))
            .ExitDown = Asc(Mid$(MapData, 39, 1)) * 256 + Asc(Mid$(MapData, 40, 1))
            .ExitLeft = Asc(Mid$(MapData, 41, 1)) * 256 + Asc(Mid$(MapData, 42, 1))
            .ExitRight = Asc(Mid$(MapData, 43, 1)) * 256 + Asc(Mid$(MapData, 44, 1))
            '.BootLocation.Map = Asc(Mid$(MapData, 45, 1)) * 256 + Asc(Mid$(MapData, 46, 1))
            '.BootLocation.X = Asc(Mid$(MapData, 47, 1))
            '.BootLocation.Y = Asc(Mid$(MapData, 48, 1))
            '.flags = Asc(Mid$(MapData, 49, 1))
            'For A = 0 To 4 '50 - 59
            '    .MonsterSpawn(A).Monster = Asc(Mid$(MapData, 61 + A * 2))
            '    .MonsterSpawn(A).Rate = Asc(Mid$(MapData, 62 + A * 2))
            'Next A
            For Y = 0 To 11
                For X = 0 To 11
                    With .Tile(X, Y)
                        A = 71 + Y * 192 + X * 16
                        .Ground = Asc(Mid$(MapData, A, 1)) * 256 + Asc(Mid$(MapData, A + 1, 1))
                        .Ground2 = Asc(Mid$(MapData, A + 2, 1)) * 256 + Asc(Mid$(MapData, A + 3, 1))
                        .BGTile1 = Asc(Mid$(MapData, A + 4, 1)) * 256 + Asc(Mid$(MapData, A + 5, 1))
                        .BGTile2 = Asc(Mid$(MapData, A + 6, 1)) * 256 + Asc(Mid$(MapData, A + 7, 1))
                        .FGTile = Asc(Mid$(MapData, A + 8, 1)) * 256 + Asc(Mid$(MapData, A + 9, 1))
                        '.Att = Asc(Mid$(MapData, A + 10, 1))
                        '.AttData(0) = Asc(Mid$(MapData, A + 11, 1))
                        '.AttData(1) = Asc(Mid$(MapData, A + 12, 1))
                        '.AttData(2) = Asc(Mid$(MapData, A + 13, 1))
                        '.AttData(3) = Asc(Mid$(MapData, A + 14, 1))
                        '.WallTile = Asc(Mid$(MapData, A + 15, 1))
                    End With
                Next X
            Next Y
        End With
    End If
End Sub
Function ClipString(St As String) As String
    Dim A As Long
    For A = Len(St) To 1 Step -1
        If Mid$(St, A, 1) <> Chr$(32) Then
            ClipString = Mid$(St, 1, A)
            Exit Function
        End If
    Next A
End Function
Sub MoveGridRight()
    Dim X As Long
    Dim Y As Long
    
    WidenMap
    
    For Y = 1 To MaxMapHeight
        For X = MaxMapWidth To 2 Step -1
            With MapGrid(X, Y)
                .MapNum = MapGrid(X - 1, Y).MapNum
                .Received = MapGrid(X - 1, Y).Received
            End With
        Next X
    Next Y
    For Y = 1 To MaxMapHeight
        With MapGrid(1, Y)
            .MapNum = 0
            .Received = False
        End With
    Next Y
    With frmMain.picMap
        BitBlt hdcMap, MapWidth, 0, .ScaleWidth - MapWidth, .ScaleHeight, hdcMap, 0, 0, SRCCOPY
        BitBlt hdcMap, 0, 0, MapWidth, .ScaleHeight, 0, 0, 0, BLACKNESS
        .Refresh
    End With
    MapX = MapX + 1
End Sub
Sub MoveGridDown()
    Dim X As Long
    Dim Y As Long
    
    HeightenMap
    
    For X = 1 To MaxMapWidth
        For Y = MaxMapHeight To 2 Step -1
            With MapGrid(X, Y)
                .MapNum = MapGrid(X, Y - 1).MapNum
                .Received = MapGrid(X, Y - 1).Received
            End With
        Next Y
    Next X
    For X = 1 To MaxMapWidth
        With MapGrid(X, 1)
            .MapNum = 0
            .Received = False
        End With
    Next X
    With frmMain.picMap
        BitBlt hdcMap, 0, MapHeight, .ScaleWidth, .ScaleHeight - MapHeight, hdcMap, 0, 0, SRCCOPY
        BitBlt hdcMap, 0, 0, .ScaleWidth, MapHeight, 0, 0, 0, BLACKNESS
        .Refresh
    End With
    MapY = MapY + 1
End Sub
Sub RefreshMap()
    Dim A As Long, X As Long, Y As Long
    Dim St As String
    Dim R As RECT
    
    'BitBlt hdcMap, 0, 0, frmMain.picMap.ScaleWidth, frmMain.picMap.ScaleHeight, hdcBackup, 0, 0, SRCCOPY
    If frmMain.mnuMapBorders.Checked = True Then
        For X = 0 To frmMain.picMap.ScaleWidth - MapWidth Step MapWidth
            frmMain.picMap.Line (X, 0)-(X, frmMain.picMap.ScaleHeight), QBColor(15)
            frmMain.picMap.Line (X + MapWidth - 1, 0)-(X + MapWidth - 1, 432), QBColor(15)
        Next X
        For Y = 0 To frmMain.picMap.ScaleHeight - MapHeight Step MapHeight
            frmMain.picMap.Line (0, Y)-(frmMain.picMap.ScaleWidth, Y), QBColor(15)
            frmMain.picMap.Line (0, Y + MapHeight - 1)-(432, Y + MapHeight - 1), QBColor(15)
        Next Y
    End If
    If frmMain.mnuMapNumbers.Checked = True Then
        For X = 1 To Int(frmMain.picMap.Width / MapWidth)
            For Y = 1 To Int(frmMain.picMap.Height / MapHeight)
                A = MapGrid(X, Y).MapNum
                If A > 0 Then
                    St = CStr(A)
                    With R
                        .Left = (X - 1) * MapWidth + 1
                        .Top = (Y - 1) * MapHeight + 1
                        .Right = X * MapWidth
                        .Bottom = Y * MapHeight
                    End With
                    SetTextColor hdcMap, QBColor(0)
                    DrawText hdcMap, St, Len(St), R, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
                    With R
                        .Left = (X - 1) * MapWidth
                        .Top = (Y - 1) * MapHeight
                        .Right = X * MapWidth - 1
                        .Bottom = Y * MapHeight - 1
                    End With
                    SetTextColor hdcMap, QBColor(15)
                    DrawText hdcMap, St, Len(St), R, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
                End If
            Next Y
        Next X
    End If
    frmMain.picMap.Refresh
End Sub
Sub ShrinkMap()
Dim R1 As RECT, R2 As RECT

    'If MapWidth = 384 And MapHeight = 384 Then
        'BitBlt hdcMap, (MapX - 1) * MapWidth, (MapY - 1) * MapHeight, 384, 384, hdcBuffer, 0, 0, SRCCOPY
        R1.Top = 0: R1.Left = 0: R1.Right = 384: R1.Bottom = 384
        R2.Left = (MapX - 1) * MapWidth: R2.Right = R2.Left + MapWidth
        R2.Top = (MapY - 1) * MapHeight: R2.Bottom = R2.Top + MapHeight
        If MapWidth = 384 Then
            sfcBuffer.Surface.BltToDC hdcMap, R1, R2
        Else
            R2.Left = 0: R2.Top = 0
            R2.Right = MapWidth: R2.Bottom = MapHeight
            sfcShrinkBuffer.Surface.Blt R2, sfcBuffer.Surface, R1, DDBLT_WAIT
            R1.Left = 0: R1.Top = 0
            R1.Right = MapWidth: R1.Bottom = MapHeight
            R2.Left = (MapX - 1) * MapWidth: R2.Right = R2.Left + MapWidth
            R2.Top = (MapY - 1) * MapHeight: R2.Bottom = R2.Top + MapHeight
            sfcShrinkBuffer.Surface.BltToDC hdcMap, R1, R2
        End If
'    Else
'        If blnLowQuality = True Then
'            SetStretchBltMode hdcMap, HALFTONE
'            StretchBlt hdcMap, (MapX - 1) * MapWidth, (MapY - 1) * MapHeight, MapWidth, MapHeight, hdcBuffer, 0, 0, 384, 384, SRCCOPY
'        Else
'            Dim X As Long, Y As Long
'            Dim X1 As Long, Y1 As Long
'            Dim C As Long
'            Dim TR As Long, TG As Long, TB As Long
'            Dim PixelWidth As Long, PixelHeight As Long
'            Dim Product As Long
'
'            PixelWidth = 384 / MapWidth
'            PixelHeight = 384 / MapHeight
'            Product = PixelWidth * PixelHeight
'
'            For X = 0 To 384 - PixelWidth Step PixelWidth
'                For Y = 0 To 384 - PixelHeight Step PixelHeight
'                    TR = 0
'                    TG = 0
'                    TB = 0
'                    For X1 = 0 To PixelWidth - 1
'                        For Y1 = 0 To PixelHeight - 1
'                            C = GetPixel(hdcBuffer, X + X1, Y + Y1)
'                            TR = TR + C Mod 256
'                            TG = TG + Int(C / 256) Mod 256
'                            TB = TB + Int(C / 65536) Mod 256
'                        Next Y1
'                    Next X1
'                    SetPixel hdcMap, (MapX - 1) * MapWidth + Int(X / PixelWidth), (MapY - 1) * MapHeight + Int(Y / PixelHeight), RGB(Int(TR / Product), Int(TG / Product), Int(TB / Product))
'                Next Y
'                DoEvents
'                If blnEnd = True Then Exit For
'            Next X
'        End If
'    End If
End Sub
Sub WidenMap()
    With frmMain.picMap
        .Width = .Width + MapWidth
        If .Width > frmMain.picMapContainer.ScaleWidth Then
            frmMain.HScroll.Max = .Width - frmMain.picMapContainer.ScaleWidth
        Else
            frmMain.HScroll.Max = 0
        End If
        hdcMap = .hdc
    End With
End Sub
Sub HeightenMap()
    With frmMain.picMap
        .Height = .Height + MapHeight
        If .Height > frmMain.picMapContainer.ScaleHeight Then
            frmMain.VScroll.Max = .Height - frmMain.picMapContainer.ScaleHeight
        Else
            frmMain.VScroll.Max = 0
        End If
        hdcMap = .hdc
    End With
End Sub

Sub WriteString(lpAppName, lpKeyName As String, A)
    Dim lpString As String, Valid As Long
    lpString = A
    Valid = WritePrivateProfileString&(lpAppName, lpKeyName, lpString, App.Path + "\odyssey.ini")
End Sub
Function ReadString(lpAppName, lpKeyName As String) As String
    Dim lpReturnedString As String, Valid As Long
    lpReturnedString = Space$(256)
    Valid = GetPrivateProfileString&(lpAppName, lpKeyName, "", lpReturnedString, 256, App.Path + "\odyssey.ini")
    ReadString = Left$(lpReturnedString, Valid)
End Function
Function ReadInt(lpAppName, lpKeyName$) As Integer
    ReadInt = GetPrivateProfileInt&(lpAppName, lpKeyName$, 0, App.Path + "\odyssey.ini")
End Function

Sub CloseClientSocket(Action As Byte)
    closesocket ClientSocket
    ClientSocket = INVALID_SOCKET
    
    If Action <> 2 Then
        frmMain.Hide
    End If
    
    If frmWait_Loaded = True Then Unload frmWait
    
    Select Case Action
        Case 0
            DeInitialize
        Case 1
            frmLogin.Show
    End Select
End Sub
Sub Main()
    Dim Pic As StdPicture, St As String, Numbers As String
        
    'If Command$ <> "" Then
        'ServerIP = Command$
        'ServerIP = "24.51.21.217" '- pk server
        'ServerIP = "64.126.86.227"
        'ServerIP = "144.118.210.220" '-temp server
        Numbers = "1234567890."
        ServerIP = Mid$(Numbers, 6, 1) & Mid$(Numbers, 9, 1) & Mid$(Numbers, 11, 1) & Mid$(Numbers, 8, 1) & Mid$(Numbers, 9, 1) & Mid$(Numbers, 11, 1) & Mid$(Numbers, 2, 1) & Mid$(Numbers, 2, 1) & Mid$(Numbers, 6, 1) & Mid$(Numbers, 11, 1) & Mid$(Numbers, 9, 1) & Mid$(Numbers, 2, 1)
        'ServerIP = "127.0.0.1"
        'ServerIP = "64.22.98.110"
        ServerIP = "Seyerdin.servegame.com"
    'Else
    '    End
    'End If

    CheckFile "tiles.rsc"
    
    Init_DX
    
    blnStop = False
    
    frmWait.Show
    frmWait.Refresh
    
    frmWait.lblStatus = "Creating Buffers .."
    frmWait.Refresh
    
    'hbmpBuffer = CreateCompatibleBitmap(frmWait.hdc, 384, 384)
    'obmpBuffer = SelectObject(hdcBuffer, hbmpBuffer)
    
    frmWait.lblStatus = "Loading Map Tiles .."
    frmWait.Refresh
    
'    hdcTiles = CreateCompatibleDC(0&)
'    Set Pic = LoadPicture("tiles.rsc")
'    hbmpTiles = Pic.Handle
'    obmpTiles = SelectObject(hdcTiles, hbmpTiles)
'
'    hdcTilesMask = CreateCompatibleDC(0&)
'    Set Pic = LoadPicture("tilesm.rsc")
'    hbmpTilesMask = Pic.Handle
'    obmpTilesMask = SelectObject(hdcTilesMask, hbmpTilesMask)

    Unload frmWait
    
    Load frmLogin
    
    'Hook Form
    Hook
    
    'Load Winsock
    StartWinsock (St)
    
    frmLogin.Show
    
    While blnEnd = False
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
        DoEvents
    Wend
    
    DeInitialize
End Sub
Sub TransparentBlt(hdc As Long, ByVal destX As Long, ByVal destY As Long, destWidth As Long, destHeight As Long, srcDC As Long, srcX As Long, srcY As Long, maskDC As Long)
    BitBlt hdc, destX, destY, destWidth, destHeight, maskDC, srcX, srcY, SRCAND
    BitBlt hdc, destX, destY, destWidth, destHeight, srcDC, srcX, srcY, SRCPAINT
End Sub

Public Sub Hook()
#If DEBUGWINDOWPROC Then
    On Error Resume Next
    Set m_SCHook = CreateWindowProcHook
    If Err Then
        MsgBox Err.Description
        Err.Clear
        Unhook
        Exit Sub
    End If
    On Error GoTo 0
    With m_SCHook
        .SetMainProc AddressOf WindowProc
        lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, .ProcAddress)
        .SetDebugProc lpPrevWndProc
    End With
#Else
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
#End If
End Sub
Sub DeInitialize()
    If ClientSocket <> INVALID_SOCKET Then
        closesocket ClientSocket
    End If
    
    'Unload Graphics
    
    Set sfcTiles.Surface = Nothing
    Set DD = Nothing
    Set Dx7 = Nothing
    
    'Unload Winsock
    EndWinsock
    
    'Unhook Form
    Unhook
    
    End
End Sub
Public Sub Unhook()
    SetWindowLong gHW, GWL_WNDPROC, lpPrevWndProc
End Sub
Sub CheckFile(Filename As String)
    If Exists(Filename) = False Then
        MsgBox "Error: File " + Chr$(34) + Filename + Chr$(34) + " not found!", vbOKOnly + vbExclamation, TitleString
        End
    End If
End Sub
Function Exists(Filename As String) As Boolean
     Exists = (Dir(Filename) <> "")
End Function

Sub SendSocket(ByVal St As String)
Dim A As Long, B As Long, C As Byte
If CPacketsSent = 255 Then CPacketsSent = 0
CPacketsSent = CPacketsSent + 1
For A = 1 To Len(St)
    B = B + Asc(Mid$(St, A, 1)) + 7
Next A
    B = B Mod 256
    C = B Xor CPacketsSent
    C = Not C
    If SendData(ClientSocket, DoubleChar(Len(St) + 1) + St + Chr$(C)) = SOCKET_ERROR Then
        CloseClientSocket 0
    End If
End Sub
Function EncryptString(St As String) As String
Dim TempStr As String, TempStr2 As String
Dim A As Integer, TmpNum As Integer

TempStr = ""
TempStr2 = ""

For A = 1 To Len(St)
    TempStr = Mid$(St, A, 1)
    TmpNum = Asc(TempStr)
    TempStr2 = TempStr2 + Chr$(TmpNum - 7)
Next A

EncryptString = Trim$(TempStr2)
End Function

Function DoubleChar(Num As Long) As String
    DoubleChar = Chr$(Int(Num / 256)) + Chr$(Num Mod 256)
End Function
Public Function Registry_Read(Key_Path, Key_Name) As Variant
    On Error Resume Next
    Dim Registry As Object
    Set Registry = CreateObject("WScript.Shell")
    Registry_Read = Registry.RegRead(Key_Path & Key_Name)
End Function
Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = 1025 Then
        'Client Socket
        Select Case lParam
            Case FD_CLOSE
                CloseClientSocket 0
                MsgBox "Disconnected from server!", vbOKOnly, TitleString
            Case FD_CONNECT
                SentPacketCount = 0
                frmWait.lblStatus = "Sending Login Information ..."
                                        SendSocket Chr$(92) + Registry_Read("HKEY_LOCAL_MACHINE\Software\Classes\OddKeys\", "UID2")
                        SendSocket Chr$(93) + Registry_Read("HKEY_LOCAL_MACHINE\Software\Classes\OddKeys\", "UID2")
                SendSocket Chr$(61) + Chr$(ClientVer) + QuadChar(GetTickCount)
                'SendSocket Chr$(1) + User + Chr$(0) + EncryptString(Pass)
                SendSocket Chr$(1) + User + Chr$(0) + Pass
            Case FD_READ
                ReceiveData
        End Select
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function

Function QuadChar(Num As Long) As String
    QuadChar = Chr$(Int(Num / 16777216) Mod 256) + Chr$(Int(Num / 65536) Mod 256) + Chr$(Int(Num / 256) Mod 256) + Chr$(Num Mod 256)
End Function

Sub ReceiveData()
'On Error GoTo errhandler
    Dim PacketLength As Integer, PacketID As Integer
    Dim St As String, St1 As String
    Dim A As Long, B As Long, C As Long, D As Long, X As Long, Y As Long
    
    SocketData = SocketData + Receive(ClientSocket)
LoopRead:
     If Len(SocketData) >= 3 Then
        PacketLength = GetInt(Mid$(SocketData, 1, 2))
        If Len(SocketData) - 2 >= PacketLength Then
            St = Mid$(SocketData, 3, PacketLength)
            SocketData = Mid$(SocketData, PacketLength + 3)
            If PacketLength > 0 Then
                PacketID = Asc(Mid$(St, 1, 1))
                If Len(St) > 1 Then
                    St = Mid$(St, 2)
                Else
                    St = ""
                End If
                Select Case PacketID
                    Case 0 'Error Logging On
                        If Len(St) >= 1 Then
                            Select Case Asc(Mid$(St, 1, 1))
                                Case 0 'Custom Message
                                    If Len(St) >= 2 Then
                                        MsgBox Mid$(St, 2), vbOKOnly + vbExclamation, TitleString
                                    End If
                                Case 1 'Invalid User/Pass
                                    MsgBox "Invalid user name/password!"
                                Case 2 'Account already in use
                                    MsgBox "Someone is already using that account!", TitleString, vbOKOnly + vbExclamation
                                Case 3 'Banned
                                    If Len(St) >= 5 Then
                                        A = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                                        If Len(St) > 5 Then
                                            MsgBox "You are banned from The Odyssey Online Classic until " + CStr(CDate(A)) + " (" + Mid$(St, 6) + ")!", TitleString, vbOKOnly
                                        Else
                                            MsgBox "You are banned from The Odyssey Online Classic until " + CStr(CDate(A)) + "!", TitleString, vbOKOnly
                                        End If
                                    End If
                                Case 4 'Server Full
                                    MsgBox "The server is full, please try again in a few minutes!", TitleString, vbOKOnly + vbExclamation
                                Case 5 'Logging in with a god account
                                    MsgBox "You are not permitted to use this god account. Your access is now set at 0!", TitleString, vbOKOnly + vbExclamation
                                    CloseClientSocket 0
                                Case 6 'Out of date
                                    MsgBox "Your version is outdated!  Download the latest version at ", TitleString, vbOKOnly + vbExclamation
                            End Select
                        End If
                        CloseClientSocket 0


                    
                    Case 3 'Logged On / Character Data
                        If frmWait_Loaded = True Then Unload frmWait
                        With frmMain
                            .Show
                            .lblStatus = "Receiving Map #" + CStr(MapNum) + " ..."
                            .Refresh
                            With .picMap
                                .Width = MapWidth
                                .Height = MapHeight
                                .Refresh
                                frmMain.HScroll.SmallChange = MapWidth
                                frmMain.HScroll.LargeChange = frmMain.picMapContainer.Width
                                frmMain.VScroll.SmallChange = MapHeight
                                frmMain.VScroll.LargeChange = frmMain.picMapContainer.Height
                                hdcMap = .hdc
                            End With
                        End With
                        MapX = 1
                        MapY = 1
                        With MapGrid(1, 1)
                            .MapNum = MapNum
                            .Received = False
                        End With
                        SendSocket Chr$(45) + DoubleChar(MapNum)
                        
                    Case 21 'Map Data
                        If Len(St) = 2379 Then
                            MapData = St
                            If MapNum > 0 Then
                            Open "mapcache.dat" For Random As #1 Len = 2379
                            Put #1, MapNum, MapData
                            Close #1
                            LoadMap MapData
                            
                            MapGrid(MapX, MapY).Received = True
                            
                            With Map(MapNum)
                                If .ExitLeft > 0 Then
                                    If MapX = 1 Then
                                        MoveGridRight
                                    End If
                                    If MapGrid(MapX - 1, MapY).MapNum > 0 Then
                                        If MapGrid(MapX - 1, MapY).MapNum <> .ExitLeft Then
                                            MsgBox "These maps do not link correctly!  Map " + CStr(MapGrid(MapX - 1, MapY).MapNum) + " and map " + CStr(.ExitLeft) + " exist in the same physical space!", vbOKOnly + vbExclamation, TitleString
                                            ''founderror = true
                                            frmMain.btnMenu.Enabled = True
                                        End If
                                    Else
                                        MapGrid(MapX - 1, MapY).MapNum = .ExitLeft
                                        MapGrid(MapX - 1, MapY).Received = False
                                    End If
                                End If
                                If .ExitUp > 0 Then
                                    If MapY = 1 Then
                                        MoveGridDown
                                    End If
                                    If MapGrid(MapX, MapY - 1).MapNum > 0 Then
                                        If MapGrid(MapX, MapY - 1).MapNum <> .ExitUp Then
                                            MsgBox "These maps do not link correctly!  Map " + CStr(MapGrid(MapX, MapY - 1).MapNum) + " and map " + CStr(.ExitUp) + " exist in the same physical space!", vbOKOnly + vbExclamation, TitleString
                                            ''founderror = true
                                            frmMain.btnMenu.Enabled = True
                                        End If
                                    Else
                                        MapGrid(MapX, MapY - 1).MapNum = .ExitUp
                                        MapGrid(MapX, MapY - 1).Received = False
                                    End If
                                End If
                                If .ExitRight > 0 Then
                                    If MapX < MaxMapWidth Then
                                        If (MapX + 1) * MapWidth > frmMain.picMap.Width Then
                                            WidenMap
                                        End If
                                        If MapGrid(MapX + 1, MapY).MapNum > 0 Then
                                            If MapGrid(MapX + 1, MapY).MapNum <> .ExitRight Then
                                                MsgBox "These maps do not link correctly!  Map " + CStr(MapGrid(MapX + 1, MapY).MapNum) + " and map " + CStr(.ExitRight) + " exist in the same physical space!", vbOKOnly + vbExclamation, TitleString
                                                ''founderror = true
                                                frmMain.btnMenu.Enabled = True
                                            End If
                                        Else
                                            MapGrid(MapX + 1, MapY).MapNum = .ExitRight
                                            MapGrid(MapX + 1, MapY).Received = False
                                        End If
                                    Else
                                        MsgBox "The map is too large to display!", vbOKOnly + vbExclamation, TitleString
                                        ''founderror = true
                                        frmMain.btnMenu.Enabled = True
                                    End If
                                End If
                                If .ExitDown > 0 Then
                                    If MapY < MaxMapHeight Then
                                        If (MapY + 1) * MapHeight > frmMain.picMap.Height Then
                                            HeightenMap
                                        End If
                                        If MapGrid(MapX, MapY + 1).MapNum > 0 Then
                                            If MapGrid(MapX, MapY + 1).MapNum <> .ExitDown Then
                                                MsgBox "These maps do not link correctly!  Map " + CStr(MapGrid(MapX, MapY + 1).MapNum) + " and map " + CStr(.ExitDown) + " exist in the same physical space!", vbOKOnly + vbExclamation, TitleString
                                                ''founderror = true
                                                frmMain.btnMenu.Enabled = True
                                            End If
                                        Else
                                            MapGrid(MapX, MapY + 1).MapNum = .ExitDown
                                            MapGrid(MapX, MapY + 1).Received = False
                                        End If
                                    Else
                                        MsgBox "The map is too large to display!", vbOKOnly + vbExclamation, TitleString
                                        frmMain.Caption = "Error: too large to display!"
                                        Exit Sub
                                    End If
                                End If
                            End With
                            
                            frmMain.lblStatus = "Drawing Map #" + CStr(MapNum) + " ..."
                            frmMain.lblStatus.Refresh
                            
                            DrawMap
                            
                            frmMain.lblStatus = "Shrinking Map #" + CStr(MapNum) + " ..."
                            frmMain.lblStatus.Refresh
                            
                            ShrinkMap
                            
                            frmMain.picMap.Refresh
                            
                            If FoundError = False Then
                                For X = 1 To MaxMapWidth
                                    For Y = 1 To MaxMapHeight
                                        With MapGrid(X, Y)
                                            If .MapNum > 0 And .Received = False And blnStop = False Then
                                                MapNum = .MapNum
                                                MapX = X
                                                MapY = Y
                                                'A = GetTickCount + 250
                                                Do While GetTickCount < A
                                                    DoEvents
                                                Loop
                                                SendSocket Chr$(45) + DoubleChar(MapNum)
                                                frmMain.lblStatus = "Receiving Map #" + CStr(MapNum) + " ..."
                                                Exit Sub
                                            End If
                                        End With
                                    Next Y
                                Next X
                            End If
                            CloseClientSocket 2
                            frmMain.lblStatus = "Done"
                            
'                            hdcBackup = CreateCompatibleDC(0&)
'                            hbmpBackup = CreateCompatibleBitmap(frmWait.hdc, frmMain.picMap.Width, frmMain.picMap.Height)
'                            obmpBackup = SelectObject(hdcBackup, hbmpBackup)
'
'                            BitBlt hdcBackup, 0, 0, frmMain.picMap.Width, frmMain.picMap.Height, hdcMap, 0, 0, SRCCOPY
                            blnDone = True
                            frmMain.cmdSTOP.Visible = False
                            frmMain.btnMenu.Enabled = True
                            End If
                        End If
                        
                    Case 58 'Ping
                        SendSocket Chr$(29) 'Pong
                    Case 67 'Booted
                        If Len(St) >= 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                If Len(St) > 1 Then
                                    MsgBox "You have been booted from The Odyssey by " + "PLAYER" + ": " + Mid$(St, 2), TitleString, vbOKOnly + vbExclamation
                                Else
                                    MsgBox "You have been booted from The Odyssey by " + "PLAYER" + "!", TitleString, vbOKOnly + vbExclamation
                                End If
                            Else
                                If Len(St) > 1 Then
                                    MsgBox "You have been booted from The Odyssey: " + Mid$(St, 2), TitleString, vbOKOnly + vbExclamation
                                Else
                                    MsgBox "You have been booted from The Odyssey!", TitleString, vbOKOnly + vbExclamation
                                End If
                            End If
                            CloseClientSocket 0
                        End If
                    Case Else
                        frmWait.Caption = PacketID
                End Select
            GoTo LoopRead
            End If
        End If
    End If
    
    Exit Sub
errhandler:
    Open App.Path & "/log.txt" For Output As #1
        Print #1, PacketID, St
    Close #1
End Sub
Function GetInt(Chars As String) As Long
    GetInt = Asc(Mid$(Chars, 1, 1)) * 256 + Asc(Mid$(Chars, 2, 1))
End Function
