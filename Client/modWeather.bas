Attribute VB_Name = "modWeather"
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

Public CurrentWeather As Long


Public Const RAIN_SPEED As Single = 0.7

Type RainDropData
    Life As Long
    Direction As Long
    y As Single
    x As Single
End Type

Public NumRainDrops As Long
Private WindSpeed As Single
Private WindSpeedGoal As Single
Private WindSpeedGoalOld As Single
Private WindCounter As Long
Public RainDrop() As RainDropData
Public RainDropVertex() As TLVERTEX

Private LastRainUpdate As Long
Private LastSnowUpdate As Long

Type SnowFlakeData
    Life As Long
    Melt As Long
End Type

Public NumSnowFlakes As Long
Public SnowFlakes() As SnowFlakeData
Public SnowFlakeVertex() As TLVERTEX

'Fog
Public texFog As Direct3DTexture8
Public CurFog As Long
Public FogX As Single
Public FogA As Single
Public FogFade As Boolean
Public FogFadeStart As Long

Public LastRainColor As Byte
Public LastSnowColor As Byte

'Lightning Bugs
Type LightningBug
    x As Single
    y As Single
End Type

'Lightning
Public LightningFlash As Long 'if > 0 then all white


Sub InitRain(ByVal Intensity As Long)
    If World.Rain > 0 Then LastRainColor = map.RainCOlor
    If map.Raining > 0 And ((map.Flags(1) And MAP_RAINING)) Then LastRainColor = map.RainCOlor
    If Intensity > 0 Then
        NumRainDrops = Intensity
        ReDim Preserve RainDrop(0 To NumRainDrops)
        ReDim Preserve RainDropVertex(0 To NumRainDrops)
        
        Dim A As Long
        
        For A = 0 To NumRainDrops
            With RainDrop(A)
                If .Life = 0 Or .Life = -1 Then
                    .Life = Int(Rnd * 30)
                    RainDropVertex(A).rhw = 1
                    If LastRainColor = 0 Then
                        RainDropVertex(A).Color = D3DColorARGB(200 - (4 * RainDrop(A).Life), 0, 0, 100)
                    Else
                        RainDropVertex(A).Color = D3DColorARGB(Lights(LastRainColor).Intensity - (4 * RainDrop(A).Life) * Lights(LastRainColor).Intensity / 200, Lights(LastRainColor).Red, Lights(LastRainColor).Green, Lights(LastRainColor).Blue)
                    End If
                    RainDropVertex(A).x = -16 + Int(Rnd * 416)
                    RainDrop(A).x = RainDropVertex(A).x
                    RainDropVertex(A).y = Int(Rnd * 420) - 36
                End If
            End With
        Next A
        WindSpeedGoal = -0.75 + (Rnd * 1.5)
        WindCounter = 500
    End If
End Sub

Sub UpdateRain(worldrain As Boolean)
    Dim A As Long, TileFlags As Long, x As Long, y As Long
    
    A = GetTickCount / 25
    If A = LastRainUpdate Then Exit Sub ' until we redo this so rain position is based on time, this is a stop gap
    LastRainUpdate = A
    
    
    TileFlags = (map.Flags(1) And MAP_TILEFLAGS)
    If WindSpeed <> WindSpeedGoal Then
        If WindSpeed < WindSpeedGoal Then
            WindSpeed = WindSpeed + (Abs(WindSpeedGoal - WindSpeedGoalOld) / 100)
        ElseIf WindSpeed > WindSpeedGoal Then
            WindSpeed = WindSpeed - (Abs(WindSpeedGoal - WindSpeedGoalOld) / 100)
        End If
    End If
    If NumRainDrops > 0 Then
        Dim raindone As Boolean
        raindone = True
        For A = 0 To NumRainDrops
            If RainDrop(A).Life > 0 Then
                RainDrop(A).Life = RainDrop(A).Life - 1
                'RainDrop(A).Y = RainDrop(A).Y + RAIN_SPEED
                RainDropVertex(A).y = RainDropVertex(A).y + 2.2
                RainDrop(A).x = RainDrop(A).x + WindSpeed
                RainDropVertex(A).x = Int(RainDrop(A).x)
                If LastRainColor = 0 Then
                    RainDropVertex(A).Color = D3DColorARGB(200 - (6.5 * RainDrop(A).Life), 100, 100, 100)
                Else
                    RainDropVertex(A).Color = D3DColorARGB(Lights(LastRainColor).Intensity - (6.5 * RainDrop(A).Life) * Lights(LastRainColor).Intensity / 200, Lights(LastRainColor).Red, Lights(LastRainColor).Green, Lights(LastRainColor).Blue)
                End If
                raindone = False
            Else
                If World.Rain = 0 And worldrain Then
                    If (map.Flags(0) And MAP_INDOORS) Then NumRainDrops = 0
                    If RainDrop(A).Life > -1 Then
                        raindone = False
                        If Int(Rnd * 2) = 0 Then
                            RainDrop(A).Life = -1
                            RainDropVertex(A).Color = &H0
                        Else
                            GoTo rainmore
                        End If
                    End If
                ElseIf (map.Raining = 0 Or (map.Flags(1) And MAP_RAINING) = False) And Not worldrain Then
                    If (map.Flags(0) And MAP_INDOORS) Then NumRainDrops = 0
                    If RainDrop(A).Life > -1 Then
                        raindone = False
                        If Int(Rnd * 2) = 0 Then
                            RainDrop(A).Life = -1
                            RainDropVertex(A).Color = &H0
                        Else
                            GoTo rainmore
                        End If
                    End If
                Else
rainmore:
                    RainDrop(A).Life = 20 + Int(Rnd * 10)
                
                    RainDropVertex(A).x = -16 + Int(Rnd * 416)
                    RainDrop(A).x = RainDropVertex(A).x
                    RainDropVertex(A).y = Int(Rnd * 420) - 36
                    
                    If LastRainColor = 0 Then
                        RainDropVertex(A).Color = D3DColorARGB(200 - (6.5 * RainDrop(A).Life), 100, 100, 100)
                    Else
                        RainDropVertex(A).Color = D3DColorARGB(Lights(LastRainColor).Intensity - (6.5 * RainDrop(A).Life) * Lights(LastRainColor).Intensity / 200, Lights(LastRainColor).Red, Lights(LastRainColor).Green, Lights(LastRainColor).Blue)
                    End If
                End If
            End If
            If TileFlags Then
                x = RainDropVertex(A).x \ 32
                y = RainDropVertex(A).y \ 32
                If x >= 0 And x <= 11 Then
                    If y >= 0 And y <= 11 Then
                        If (map.Tile(x, y).Anim(2) And 64) Then
                            RainDropVertex(A).Color = &H0
                        End If
                    End If
                End If
            End If
        Next A
        
        If raindone Then
            NumRainDrops = 0
        End If
        
        WindCounter = WindCounter - 1
        If WindCounter = 0 Then
            WindCounter = 500
            If Rnd < 0.25 Then
                WindSpeedGoalOld = WindSpeedGoal
                WindSpeedGoal = -0.75 + (Rnd * 1.5)
            End If
        End If
'        If LightningFlash = 0 Then
'            If Int(Rnd * 5000) = 1 Then
'                LightningFlash = 17
'                AmbientAlpha = 255
'                AmbientRed = 255
'                AmbientGreen = 255
'                AmbientBlue = 255
'            End If
'        Else
'            LightningFlash = LightningFlash - 1
'            If LightningFlash = 7 Then
'                CalculateAmbientAlpha
'                AmbientAlpha = AmbientAlpha - 40
'                If AmbientAlpha < 0 Then AmbientAlpha = 0
'            End If
'            If LightningFlash = 0 Then CalculateAmbientAlpha
'        End If
    End If
End Sub

Sub DrawRain()
    
    Dim Buf As D3DXBuffer, A As Long, b As Single, rad As Single, tex As Byte
    
    If NumRainDrops > 0 Then
        If LastRainColor = 0 Then
            rad = 5
            tex = 1
        Else
            rad = Lights(LastRainColor).Radius
            tex = Lights(LastRainColor).MaxFlicker
            If tex > 10 Then tex = 1
            If tex = 0 Then tex = 1
            If rad > 50 Then rad = 50
        End If
        'Set Buf = D3DX.CreateBuffer(4)
        'B = 7
        'D3DX.BufferSetData Buf, 0, 4, 1, B
        'D3DX.BufferGetData Buf, 0, 4, 1, A
        
        D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
        'D3DDevice.SetRenderState D3DRS_POINTSIZE, A
        D3DDevice.SetRenderState D3DRS_POINTSIZE, FtoDW(rad)
        'D3DDevice.SetRenderState D3DRS_POINTSIZE_MIN, A
        
        D3DDevice.SetTexture 0, texParticles(tex)
        'D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        
        D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, NumRainDrops + 1, RainDropVertex(0), Len(RainDropVertex(0))
        
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
End Sub
Sub DrawSnow()
    Dim Buf As D3DXBuffer, A As Long, b As Single, rad As Single, tex As Byte
    
    A = 0
    If NumSnowFlakes > 0 Then
        If LastSnowColor = 0 Then
            rad = 1
            tex = 2
        Else
            rad = Lights(LastSnowColor).Radius
            tex = Lights(LastSnowColor).MaxFlicker
            If tex > 10 Then tex = 2
            If tex = 0 Then tex = 2
            If rad > 50 Then rad = 50
        End If
        'Set Buf = D3DX.CreateBuffer(4)
        'B = 1
        'D3DX.BufferSetData Buf, 0, 4, 1, B
        'D3DX.BufferGetData Buf, 0, 4, 1, A
        
        D3DDevice.SetRenderState D3DRS_POINTSIZE, FtoDW(rad)
        D3DDevice.SetTexture 0, texParticles(tex)
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        
        D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, NumSnowFlakes + 1, SnowFlakeVertex(0), Len(SnowFlakeVertex(0))
        
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
    
    

    
End Sub


Sub InitSnow(ByVal Intensity As Long)
    Dim A As Long
    Dim C As Long
    
    If World.Snow > 0 Then LastSnowColor = map.SnowColor
    If map.Snowing > 0 And ((map.Flags(1) And MAP_SNOWING)) Then LastSnowColor = map.SnowColor
    If Intensity > 0 Then
        NumSnowFlakes = Intensity
        If NumSnowFlakes > 0 Then
            ReDim SnowFlakes(0 To NumSnowFlakes)
            ReDim SnowFlakeVertex(0 To NumSnowFlakes)
        End If
    
            If LastSnowColor = 0 Then
                C = D3DColorARGB(255, 255, 255, 255)
            Else
                C = D3DColorARGB(Lights(LastSnowColor).Intensity, Lights(LastSnowColor).Red, Lights(LastSnowColor).Green, Lights(LastSnowColor).Blue)
            End If
            For A = 0 To NumSnowFlakes
            
            With SnowFlakes(A)
                .Life = Int(Rnd * 95)
            End With
            With SnowFlakeVertex(A)
                .rhw = 1
                .x = Int(Rnd * 384)
                .y = Int(Rnd * 420) - 36
                .Color = C
            End With
        Next A
    End If
End Sub

Sub UpdateSnow(worldsnow As Boolean)
    Dim TileFlags As Long, x As Long, y As Long, A As Long
    
    A = GetTickCount / 25
    If A = LastSnowUpdate Then Exit Sub ' until we redo this so snow position is based on time, this is a stop gap
    LastSnowUpdate = A
     
    TileFlags = (map.Flags(1) And MAP_TILEFLAGS)
    If NumSnowFlakes > 0 Then
        Dim snowdone As Boolean
        snowdone = True
        For A = 0 To NumSnowFlakes
            If SnowFlakes(A).Life > 0 Then
                SnowFlakes(A).Life = SnowFlakes(A).Life - 1
                SnowFlakeVertex(A).y = SnowFlakeVertex(A).y + Int(Rnd * 2)
                SnowFlakeVertex(A).x = SnowFlakeVertex(A).x + Int(Rnd * 3) - 1
                
                If LastSnowColor = 0 Then
                    SnowFlakeVertex(A).Color = D3DColorARGB(255 * ((95 - SnowFlakes(A).Life) / 95), 255, 255, 255)
                Else
                     SnowFlakeVertex(A).Color = D3DColorARGB(Lights(LastSnowColor).Intensity * ((95 - SnowFlakes(A).Life) / 95), Lights(LastSnowColor).Red, Lights(LastSnowColor).Green, Lights(LastSnowColor).Blue)
                End If
                snowdone = False
            Else
                If World.Snow = 0 And worldsnow Then
                    If (map.Flags(0) And MAP_INDOORS) Then NumSnowFlakes = 0
                    If SnowFlakes(A).Life > -1 Then
                        snowdone = False
                        'If Int(Rnd * 2) = 0 Then
                            SnowFlakes(A).Life = -1
                            SnowFlakeVertex(A).Color = &H0
                        'Else
                        '    GoTo snowmore
                        'End If
                    End If
                ElseIf (map.Snowing = 0 Or (map.Flags(1) And MAP_SNOWING) = False) And Not worldsnow Then
                    If (map.Flags(0) And MAP_INDOORS) Then NumSnowFlakes = 0
                    If SnowFlakes(A).Life > -1 Then
                        snowdone = False
                        'If Int(Rnd * 2) = 0 Then
                            SnowFlakes(A).Life = -1
                            SnowFlakeVertex(A).Color = &H0
                        'Else
                        '    GoTo snowmore
                        'End If
                    End If
                        
                Else
'snowmore:
                    If SnowFlakes(A).Melt > 0 Then
                        SnowFlakes(A).Melt = SnowFlakes(A).Melt - 1
                        If LastSnowColor = 0 Then
                            SnowFlakeVertex(A).Color = D3DColorARGB(255 * ((95 - SnowFlakes(A).Life) / 95), 255, 255, 255)
                        Else
                            SnowFlakeVertex(A).Color = D3DColorARGB(Lights(LastSnowColor).Intensity * ((SnowFlakes(A).Melt) / 30), Lights(LastSnowColor).Red, Lights(LastSnowColor).Green, Lights(LastSnowColor).Blue)
                        End If
                    Else
                        SnowFlakes(A).Life = 65 + Int(Rnd * 30)
                        SnowFlakes(A).Melt = 30
                        With SnowFlakeVertex(A)
                            .x = Int(Rnd * 384)
                            .y = Int(Rnd * 420) - 36
                            .Color = &HFFFFFF
                        End With
                    End If
                End If
            End If
            If TileFlags Then
                x = SnowFlakeVertex(A).x \ 32
                y = SnowFlakeVertex(A).y \ 32
                If x >= 0 And x <= 11 Then
                    If y >= 0 And y <= 11 Then
                        If (map.Tile(x, y).Anim(2) And 64) Then
                            SnowFlakeVertex(A).Color = &H0
                        End If
                    End If
                End If
            End If
        Next A
            If snowdone Then
                NumSnowFlakes = 0
            End If
    End If
End Sub



Public Sub InitFog(ByVal Fog As Long, FadeIn As Boolean)
    Dim TexInfo As D3DXIMAGE_INFO_A
    
    If Fog > 0 And Fog <= 31 Then
        If CurFog <> Fog Then
            FogFade = False
            FogFadeStart = GetTickCount
            If Exists("Data/Graphics/Misc/Fog/Fog" & Fog & ".rsc") Then
                If CurFog > 0 Then
                    Set texFog = Nothing
                Else
                    If FadeIn Then FogA = 0 Else FogA = 127
                End If
                Set texFog = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Misc/Fog/Fog" & Fog & ".rsc", 512, 256, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_NONE, &HFF000000, TexInfo, ByVal 0)
                CurFog = Fog
                FogX = 0#
            Else
                ClearFog False
            End If
        End If
    End If
End Sub

Public Sub ClearFog(ByVal HardFade As Boolean)
    If HardFade Or FogA = 0 Then
        FogFade = False
        FogFadeStart = GetTickCount
        FogA = 0
        Set texFog = Nothing
        CurFog = 0
    Else
        FogFade = True
        FogFadeStart = GetTickCount
    End If
End Sub

Sub UpdateFog()
    FogX = GetTickCount / 84 Mod 512 'FogX + 0.3 per 40 fps frame
    If FogX > 512 Then FogX = 0
    If FogFade Then
        If FogA > 0 Then
            FogA = 127 - (GetTickCount - FogFadeStart) / 84 'FogX + 0.3
        Else
            ClearFog True
        End If
    Else
        If FogA < 127 Then FogA = (GetTickCount - FogFadeStart) / 84 'FogX - 0.3
    End If
End Sub


Sub DrawFog()
    Dim VertexArray(0 To 3) As TLVERTEX
    
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
    VertexArray(0).Color = D3DColorARGB(FogA, 255, 255, 255)
    VertexArray(1).Color = D3DColorARGB(FogA, 255, 255, 255)
    VertexArray(2).Color = D3DColorARGB(FogA, 255, 255, 255)
    VertexArray(3).Color = D3DColorARGB(FogA, 255, 255, 255)

    If CurFog > 0 Then
            D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
            D3DDevice.SetTexture 0, texFog
            
            VertexArray(0).x = 0
            VertexArray(0).tu = FogX / 512
            VertexArray(0).y = 0
            VertexArray(0).tv = 0
            VertexArray(1).x = 384
            VertexArray(1).tu = (FogX + 513) / 512
            VertexArray(2).x = VertexArray(0).x
            VertexArray(3).x = VertexArray(1).x
            VertexArray(2).y = 384
            VertexArray(2).tv = 257 / 256
            VertexArray(1).y = VertexArray(0).y
            VertexArray(1).tv = VertexArray(0).tv
            VertexArray(2).tu = VertexArray(0).tu
            VertexArray(3).y = VertexArray(2).y
            VertexArray(3).tu = VertexArray(1).tu
            VertexArray(3).tv = VertexArray(2).tv
            
            

            D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
            D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
            D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
    End If
End Sub



