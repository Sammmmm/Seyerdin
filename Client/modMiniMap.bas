Attribute VB_Name = "modMiniMap"
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

'Made minimap have its own module for the sake of keeping
'the code from being randomly spewed over the entire project
Option Explicit

Type MiniTileData
    Ground As Integer
    Ground2 As Integer
    BGTile1 As Integer
    Anim(1 To 2) As Byte
    FGTile As Integer
End Type

Type MiniMapData
    'Name As String
    map As Integer
    ExitUp As Integer
    ExitDown As Integer
    ExitLeft As Integer
    ExitRight As Integer
    Tile(0 To 11, 0 To 11) As MiniTileData
    Intensity As Byte
    Version As Long
    Loaded As Boolean
End Type

Public Const MiniMapWidth = 5
Public Const MiniMapHeight = 5
Public doMiniMapDraw As Boolean
'Other Variables
Public CMapString(625) As Byte 'This holds a bit array of how many

Dim MiniMaps(0 To MiniMapWidth - 1, 0 To MiniMapHeight - 1) As MiniMapData

Public texMiniMap As TextureType

Public LastCX As Long
Public LastCY As Long

Sub InitMiniMap()
        doMiniMapDraw = True
        Set texMiniMap.Texture = D3DX.CreateTexture(D3DDevice, 256, 256, 0, D3DUSAGE_RENDERTARGET, D3DFMT_X8R8G8B8, D3DPOOL_DEFAULT)
        'set texminimap(a).Texture = d3dx.CreateCubeTextureFromFileEx(d3ddevice,
        '                                CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Atts.rsc", 256, 256, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texAtts.TexInfo, ByVal 0)
        texMiniMap.TexInfo.Width = 256
        texMiniMap.TexInfo.Height = 256
End Sub

Sub UnloadMiniMap()
        Set texMiniMap.Texture = Nothing
End Sub

Sub CreateMiniMap()
    Dim CurX As Long, CurY As Long
    Dim blnLoading As Boolean

    For CurX = 0 To MiniMapWidth - 1
        For CurY = 0 To MiniMapHeight - 1
            MiniMaps(CurX, CurY).Loaded = False
            MiniMaps(CurX, CurY).Version = 0
            MiniMaps(CurX, CurY).map = 0
        Next CurY
    Next CurX

    If CMap > 0 Then
        LoadMiniMap CMap, (MiniMapWidth \ 2), (MiniMapHeight \ 2)
        Do
            blnLoading = False
            For CurX = 0 To MiniMapWidth - 1
                For CurY = 0 To MiniMapHeight - 1
                    With MiniMaps(CurX, CurY)
                        If .Version > 0 Then
                            If .Loaded = True Then
                                If .ExitUp > 0 Then
                                    If CurY > 0 Then
                                        If MiniMaps(CurX, CurY - 1).Loaded = False Then
                                            If HasBeenThere(.ExitUp) Then
                                                LoadMiniMap .ExitUp, CurX, CurY - 1
                                                blnLoading = True
                                            Else
                                                MiniMaps(CurX, CurY - 1).Loaded = True
                                            End If
                                        End If
                                    End If
                                End If
                                If .ExitDown > 0 Then
                                    If CurY < MiniMapHeight - 1 Then
                                        If MiniMaps(CurX, CurY + 1).Loaded = False Then
                                            If HasBeenThere(.ExitDown) Then
                                                LoadMiniMap .ExitDown, CurX, CurY + 1
                                                blnLoading = True
                                            Else
                                                MiniMaps(CurX, CurY + 1).Loaded = True
                                            End If
                                        End If
                                    End If
                                End If
                                If .ExitLeft > 0 Then
                                    If CurX > 0 Then
                                        If MiniMaps(CurX - 1, CurY).Loaded = False Then
                                            If HasBeenThere(.ExitLeft) Then
                                                LoadMiniMap .ExitLeft, CurX - 1, CurY
                                                blnLoading = True
                                            Else
                                                MiniMaps(CurX - 1, CurY).Loaded = True
                                            End If
                                        End If
                                    End If
                                End If
                                If .ExitRight > 0 Then
                                    If CurX < MiniMapWidth - 1 Then
                                        If MiniMaps(CurX + 1, CurY).Loaded = False Then
                                            If HasBeenThere(.ExitRight) Then
                                                LoadMiniMap .ExitRight, CurX + 1, CurY
                                                blnLoading = True
                                            Else
                                                MiniMaps(CurX + 1, CurY).Loaded = True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End With
                Next CurY
            Next CurX
        Loop While blnLoading = True
    End If
    
        ReDrawMiniMap
        DrawMiniMapSection
End Sub

Function LoadMiniMap(ByVal MapNum As Integer, MapX As Long, MapY As Long) As Boolean
    Dim A As Long, x As Long, y As Long
    Dim MapData As String * 2379
    
    Open MapCacheFile For Random As #1 Len = 2379
    Get #1, MapNum, MapData
    Close #1
    MapData = RC4_DecryptString(MapData, MapKey)
    
    If Len(MapData) = 2379 Or Len(MapData) = 2374 Then
        With MiniMaps(MapX, MapY)
            '.Name = ClipString$(Mid$(MapData, 1, 30))
            .map = MapNum
            .Version = Asc(Mid$(MapData, 31, 1)) * 16777216 + Asc(Mid$(MapData, 32, 1)) * 65536 + Asc(Mid$(MapData, 33, 1)) * 256& + Asc(Mid$(MapData, 34, 1))
            If .Version = 0 Then
                .Loaded = True
                LoadMiniMap = False
                Exit Function
            End If
            
            .ExitUp = Asc(Mid$(MapData, 37, 1)) * 256& + Asc(Mid$(MapData, 38, 1))
            .ExitDown = Asc(Mid$(MapData, 39, 1)) * 256& + Asc(Mid$(MapData, 40, 1))
            .ExitLeft = Asc(Mid$(MapData, 41, 1)) * 256& + Asc(Mid$(MapData, 42, 1))
            .ExitRight = Asc(Mid$(MapData, 43, 1)) * 256& + Asc(Mid$(MapData, 44, 1))
            .Intensity = 255 - Asc(Mid$(MapData, 50, 1))
            For y = 0 To 11
                For x = 0 To 11
                    With .Tile(x, y)
                        A = 71 + y * 192 + x * 16
                        .Ground = Asc(Mid$(MapData, A, 1)) * 256& + Asc(Mid$(MapData, A + 1, 1))
                        .Ground2 = Asc(Mid$(MapData, A + 2, 1)) * 256& + Asc(Mid$(MapData, A + 3, 1))
                        .BGTile1 = Asc(Mid$(MapData, A + 4, 1)) * 256& + Asc(Mid$(MapData, A + 5, 1))
                        .Anim(1) = Asc(Mid$(MapData, A + 6, 1))
                        .Anim(2) = Asc(Mid$(MapData, A + 7, 1))
                        .FGTile = Asc(Mid$(MapData, A + 8, 1)) * 256& + Asc(Mid$(MapData, A + 9, 1))
                    End With
                Next x
            Next y
            .Loaded = True
        End With
    End If
    
    LoadMiniMap = True
End Function

Public Sub ReDrawMiniMap()
'Exit Sub
    If D3DDevice.TestCooperativeLevel = D3D_OK Then

        Dim CurX As Long, CurY As Long
        Dim x As Long, y As Long
        'Dim tSurface As Direct3DSurface8
    
        'Set tSurface = texMiniMap.Texture.GetSurfaceLevel(0)
        'D3DDevice.SetRenderTarget texMiniMap.Texture.GetSurfaceLevel(0), Nothing, 0
        'D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, ByVal &HFF000000, 1, 0
        'D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, ByVal &H0&, 1#, 0
        If (Character.StatusEffect And (2 ^ SE_BLIND)) = 0 Then
            D3DDevice.SetRenderTarget texMiniMap.Texture.GetSurfaceLevel(0), Nothing, 0
            D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, ByVal &H0&, 1#, 0
            D3DDevice.BeginScene
            For CurX = 0 To MiniMapWidth - 1
                For CurY = 0 To MiniMapHeight - 1
                    If MiniMaps(CurX, CurY).Loaded Then
                        If MiniMaps(CurX, CurY).Version > 0 Then
                            For x = 0 To 11
                                For y = 0 To 11
                                    With MiniMaps(CurX, CurY).Tile(x, y)
                                        If .Ground > 0 Then
                                            DrawTileEX CurX * 48 + x * 4, CurY * 48 + y * 4, 4, 4, .Ground, MiniMaps(CurX, CurY).Intensity
                                        End If
                                    End With
                                Next y
                            Next x
                        End If
                    End If
                Next CurY
            Next CurX
            
            'D3DDevice.SetRenderTarget texMiniMap(2).Texture.GetSurfaceLevel(0), Nothing, 0
            'D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, ByVal &H0&, 1#, 0
            For CurX = 0 To MiniMapWidth - 1
                For CurY = 0 To MiniMapHeight - 1
                    If MiniMaps(CurX, CurY).Loaded Then
                        If MiniMaps(CurX, CurY).Version > 0 Then
                            For x = 0 To 11
                                For y = 0 To 11
                                    With MiniMaps(CurX, CurY).Tile(x, y)
                                        If .Ground2 > 0 Then
                                            DrawTileEX CurX * 48 + x * 4, CurY * 48 + y * 4, 4, 4, .Ground2, MiniMaps(CurX, CurY).Intensity
                                        End If
                                    End With
                                Next y
                            Next x
                        End If
                    End If
                Next CurY
            Next CurX
            
            'D3DDevice.SetRenderTarget texMiniMap(2).Texture.GetSurfaceLevel(0), Nothing, 0
            'D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, ByVal &H0&, 1#, 0
            For CurX = 0 To MiniMapWidth - 1
                For CurY = 0 To MiniMapHeight - 1
                    If MiniMaps(CurX, CurY).Loaded Then
                        If MiniMaps(CurX, CurY).Version > 0 Then
                            For x = 0 To 11
                                For y = 0 To 11
                                    With MiniMaps(CurX, CurY).Tile(x, y)
                                        If .BGTile1 > 0 Then
                                            DrawTileEX CurX * 48 + x * 4, CurY * 48 + y * 4, 4, 4, .BGTile1, MiniMaps(CurX, CurY).Intensity
                                        End If
                                    End With
                                Next y
                            Next x
                        End If
                    End If
                Next CurY
            Next CurX
            
            'D3DDevice.SetRenderTarget texMiniMap(3).Texture.GetSurfaceLevel(0), Nothing, 0
            'D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, ByVal &H0&, 1#, 0
            For CurX = 0 To MiniMapWidth - 1
                For CurY = 0 To MiniMapHeight - 1
                    If MiniMaps(CurX, CurY).Loaded Then
                        If MiniMaps(CurX, CurY).Version > 0 Then
                            For x = 0 To 11
                                For y = 0 To 11
                                    With MiniMaps(CurX, CurY).Tile(x, y)
                                        If .FGTile > 0 Then
                                            DrawTileEX CurX * 48 + x * 4, CurY * 48 + y * 4, 4, 4, .FGTile, MiniMaps(CurX, CurY).Intensity
                                        End If
                                    End With
                                Next y
                            Next x
                        End If
                    End If
                    
           '         For X = 1 To MAXUSERS
           '             If (player(X).Party = Character.Party And player(X).Map = MiniMaps(CurX, CurY).Map) Then
           '                 DrawBmpString3D Create3DString("P"), 48 * CurX + 22, 48 * CurY + 12, True, StatusColors(23)
           '                 'DrawRect CurX * 48 + 23, CurX * 48 + 25, CurY * 48 + 23, CurY * 48 + 25, BS_SOLID, StatusColors(23), 0, 0
           '             End If
           '         Next X
                Next CurY
            Next CurX
            
            
            
            D3DDevice.EndScene
            End If
        D3DDevice.SetRenderTarget RenderSurface(0), Nothing, 0
    End If
End Sub

Sub DrawMiniMapSection()
If D3DDevice.TestCooperativeLevel = D3D_OK Then
Dim A As Long, x As Long, y As Long
    If MiniMapTab = tsMap Then
        D3DDevice.SetRenderTarget RenderSurface(1), Nothing, 0
        D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, ByVal &H0&, 1#, 0
        D3DDevice.BeginScene
            D3DDevice.SetTexture 0, texMiniMap.Texture
            Draw3D 0, 0, 174, 122, 33 - (24 - (cX * 4)), 59 - (24 - (cY * 4)), 0, -1, 256, 256
            'D3DDevice.SetTexture 0, texMiniMap(2).Texture
            'Draw3D 0, 0, 174, 122, 33 - (24 - (cX * 4)), 59 - (24 - (cY * 4)), 0, -1, 256, 256
            'D3DDevice.SetTexture 0, texMiniMap(3).Texture
            'Draw3D 0, 0, 174, 122, 33 - (24 - (cX * 4)), 59 - (24 - (cY * 4)), 0, -1, 256, 256
            'D3DDevice.SetTexture 0, texMiniMap(4).Texture
            'Draw3D 0, 0, 174, 122, 33 - (24 - (cX * 4)), 59 - (24 - (cY * 4)), 0, -1, 256, 256
            
            
                                
                    For A = 1 To MAXUSERS
                        If (player(A).Guild = Character.Guild And Character.Guild > 0) Then
                            With (player(A))
                                For x = 0 To 4
                                For y = 0 To 4
                                    If MiniMaps(x, y).map = .map And .map <> 0 Then
                                        'DrawBmpString3D Create3DString(Mid$(.Name, 1, 1)), -(55 - (24 - (cX * 4))) + (48 * X + 22) + .X * 4, -(76 - (24 - (cY * 4))) + 48 * Y + 12 + .Y * 4, True, StatusColors(26), True, 0.8
                                        DrawBmpString3D Create3DString(Mid$(.Name, 1, 1)), -9 - (cX * 4) + (48 * x) + .x * 4, -40 - (cY * 4) + 48 * y + .y * 4, StatusColors(26), True, 0.8
                                        'DrawRect CurX * 48 + 23, CurX * 48 + 25, CurY * 48 + 23, CurY * 48 + 25, BS_SOLID, StatusColors(23), 0, 0
                                    End If
                                Next y
                                Next x
                            End With
                        ElseIf (player(A).Party = Character.Party And Character.Party > 0) Then
                            With (player(A))
                                For x = 0 To 4
                                For y = 0 To 4
                                    If MiniMaps(x, y).map = .map And .map <> 0 Then
                                        DrawBmpString3D Create3DString(Mid$(.Name, 1, 1)), -9 - (cX * 4) + (48 * x) + .x * 4, -40 - (cY * 4) + 48 * y + .y * 4, StatusColors(7), True, 0.8
                                        'DrawBmpString3D Create3DString(Mid$(.Name, 1, 1)), -(55 - (24 - (cX * 4))) + (48 * X + 22) + .X * 4, -(76 - (24 - (cY * 4))) + 48 * Y + 12 + .Y * 4, True, StatusColors(7), True, 0.8
                                        'DrawRect CurX * 48 + 23, CurX * 48 + 25, CurY * 48 + 23, CurY * 48 + 25, BS_SOLID, StatusColors(23), 0, 0
                                    End If
                                Next y
                                Next x
                            End With
                        End If
                    Next A
            
            
            DrawRect 86, 90, 60, 64, BS_SOLID, &HFFFFCF5F, 0, 0
        D3DDevice.EndScene
        SwapChain(1).Present ByVal 0, ByVal 0, 0, ByVal 0
        D3DDevice.SetRenderTarget RenderSurface(0), Nothing, 0
        LastTexture = 0
    Else
        If Options.VsyncEnabled Then
            SwapChain(1).Present ByVal 0, ByVal 0, 0, ByVal 0
        End If
    End If
    LastCX = cX: LastCY = cY
End If
End Sub

Sub LoadMapStringData(Account As String)
Dim St1 As String * 1024
Dim DatPath As String
Dim CheckSum As Long, Check As Long, ByteArray() As Byte
Dim A As Long, b As Long

MyKey = Account & MapKey

DatPath = "Data\Cache\" & Account & ".dat" & ServerId
    If Exists(DatPath) Then
        Open DatPath For Random As #1 Len = 1024
            Get #1, , St1
        Close #1
        St1 = RC4_DecryptString(St1, MyKey)
    Else
        CreateMapStringData Account
        ClearMapString
    End If

    ByteArray() = StrConv(St1, vbFromUnicode)

    
    CheckSum = Asc(Mid$(St1, 1021, 1))
    If CheckSum > 127 Then GoTo notworking
    CheckSum = (CheckSum * 16777216) + Asc(Mid$(St1, 1022, 1)) * 65536 + Asc(Mid$(St1, 1023, 1)) * 256& + Asc(Mid$(St1, 1024, 1))
    'Loop through bits
    For A = 0 To 1019
        For b = 0 To 7
            If (ByteArray(A) And 2 ^ b) Then
                Check = Check + 1
            End If
        Next b
    Next A
    
    If Check = CheckSum Then
        MemCopy CMapString(0), ByteArray(0), 625
    Else
notworking:
        ClearMapString
    End If
    
End Sub

Sub CreateMapStringData(Account As String)
Dim St1 As String * 1024
Dim DatPath As String
DatPath = "Data\Cache\" & Account & ".dat" & ServerId
MyKey = Account & MapKey
If Exists(DatPath) = False Then
    St1 = String$(1024, 0)
    St1 = RC4_EncryptString(St1, MyKey)
    Open DatPath For Random As #1 Len = 1024
        Put #1, , St1
    Close #1
End If
End Sub

Function HasBeenThere(ByVal MapNum As Long)
    If MapNum > 0 And MapNum <= 5000 Then
        If (CMapString(MapNum \ 8) And 2 ^ (MapNum Mod 8)) Then
            HasBeenThere = True
        Else
            HasBeenThere = False
        End If
    Else
        HasBeenThere = False
    End If
End Function

Sub SetMapBit(MapNum As Long, Account As String)
    Dim byt As Byte
    
    byt = CMapString(MapNum \ 8)
    byt = byt Or (2 ^ (MapNum Mod 8))
    CMapString(MapNum \ 8) = byt
    
    SaveMapString Account
End Sub

Sub ClearMapString()
    Dim A As Long
    For A = 0 To 625
        CMapString(A) = 0
    Next A
End Sub

Sub SaveMapString(Account As String)
Dim St1 As String * 625
Dim St2 As String * 1024
Dim DatPath As String
Dim ByteArray() As Byte, A As Long, b As Long, Check As Long
DatPath = "Data\Cache\" & Account & ".dat" & ServerId
MyKey = Account & MapKey

St1 = StrConv(CMapString, vbUnicode)

If Exists(DatPath) Then
    Open DatPath For Random As #1 Len = 1024
        Get #1, , St2
    Close #1
    St2 = RC4_DecryptString(St2, MyKey)
    St2 = Mid$(St1, 1, 625) & Mid$(St2, 626, 399)
    
    ByteArray = StrConv(St2, vbFromUnicode)
    
    For A = 0 To 1019
        For b = 0 To 7
            If (ByteArray(A) And 2 ^ b) Then
                Check = Check + 1
            End If
        Next b
    Next A
    
    St2 = Mid$(St2, 1, 1020) & QuadChar(Check)
    'MsgBox Check
    
    St2 = RC4_EncryptString(St2, MyKey)
    
    Open DatPath For Random As #1 Len = 1024
        Put #1, , St2
    Close #1
End If
End Sub

Sub DrawTileEX(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Tile As Long, ByVal Brightness As Long)
Dim VertexArray(0 To 3) As TLVERTEX
    Dim tColor As Long
    tColor = D3DColorARGB(255, Brightness, Brightness, Brightness)
    
    'Static LastTex As Long
    Dim CurTex As Long
    CurTex = (((Tile - 1) \ 64) + 1)
    'If LastTex <> CurTex Then
        'D3DDevice.SetTexture 0, texTiles(CurTex).Texture
        ReadyTexture TexTile0 + CurTex
    '    LastTex = CurTex
    'End If
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
    VertexArray(0).Color = tColor
    VertexArray(1).Color = tColor
    VertexArray(2).Color = tColor
    VertexArray(3).Color = tColor

    VertexArray(0).x = x - 0.5
    VertexArray(0).tu = ((((Tile - 1) Mod 8) * 32) / 256)
    VertexArray(0).y = y - 0.5
    VertexArray(0).tv = ((((Tile - 1) \ 8) * 32) / 256)
    VertexArray(1).x = x + Width - 0.5
    VertexArray(1).tu = (((Tile - 1) Mod 8) * 32 + 32) / 256
    VertexArray(2).x = VertexArray(0).x
    VertexArray(3).x = VertexArray(1).x
    VertexArray(2).y = y + Height - 0.5
    VertexArray(2).tv = (((Tile - 1) \ 8) * 32 + 32) / 256

    VertexArray(1).y = VertexArray(0).y
    VertexArray(1).tv = VertexArray(0).tv
    VertexArray(2).tu = VertexArray(0).tu
    VertexArray(3).y = VertexArray(2).y
    VertexArray(3).tu = VertexArray(1).tu
    VertexArray(3).tv = VertexArray(2).tv

    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
End Sub
