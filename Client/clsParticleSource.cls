VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParticleSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Public SourceType As Long
Public x As Long
Public y As Long
Private XOffset As Single
Private YOffset As Single
Public Red As Long
Public Green As Long
Public Blue As Long
Public NumParticles As Long
Public ParticleSize As Long
Public lng0 As Long
Public lng1 As Long
Public ParticleNum As Long
Public Key As String
Public Size1 As Long
Public VelocityX As Single
Public VelocityY As Single
Public Life As Long, LifeLeft As Long
Public Direction As Long
Public BlendMode As Long


Public TargType As Long
Public Target As Long

Private Type ParticleData
    Birth As Long
    Life As Single
    Direction As Single
    VelocityX As Single
    VelocityY As Single
    R As Integer
    G As Integer
    b As Integer
    x As Single
    y As Single
End Type

Private LastParticleUpdate As Long
Private Particles() As ParticleData
Private ParticlesVertex() As TLVERTEX


Public Sub Init()
    Dim A As Long, Angle As Double, dA As Double, Amt As Long, b As Long

    If NumParticles > 0 Then
        Select Case SourceType
            Case 1
                Amt = NumParticles * 0.3
            Case 2
                Amt = NumParticles * 0.4
            Case 3
                Amt = NumParticles * 0.4
        End Select
        ReDim Particles(0 To NumParticles)
        ReDim ParticlesVertex(0 To NumParticles)
        
        YOffset = 0
        XOffset = 0
        
        'Default
        For A = 0 To NumParticles
            ParticlesVertex(A).rhw = 1
        Next A
        
        BlendMode = D3DBLEND_ONE
        
        Select Case SourceType
            Case 1 'Round Circle, fade in, fall down while disappearing
                For A = 0 To NumParticles - Amt
                    Angle = Rnd * (3.14159265358979 * 2)
                    dA = Rnd * Size1
                    ParticlesVertex(A).x = x + (Cos(Angle) * dA)
                    ParticlesVertex(A).y = y + (Sin(Angle) * dA) - 16
                    ParticlesVertex(A).Color = D3DColorARGB(255, Red, Green, Blue)
                    Particles(A).Life = Int(Life * 0.75)
                    Particles(A).Birth = Int(Life * 0.25)
                Next A
                For A = NumParticles - Amt To NumParticles
                    Angle = Rnd * (3.14159265358979 * 2)
                    ParticlesVertex(A).x = x + (Cos(Angle) * Size1)
                    ParticlesVertex(A).y = y + (Sin(Angle) * Size1) - 16
                    ParticlesVertex(A).Color = D3DColorARGB(255, Red, Green, Blue)
                    Particles(A).Life = Int(Life * 0.75)
                    Particles(A).Birth = Int(Life * 0.25)
                Next A
            Case 2 'Start at bottom, rise up, dropping particles
                Life = 50
                For A = 0 To NumParticles - Amt 'Falling Particles
                    ParticlesVertex(A).x = x + XOffset + Int(Rnd * 32)
                    ParticlesVertex(A).y = y + 32
                    Particles(A).Life = Int(Rnd * 10) + 1
                    ParticlesVertex(A).Color = D3DColorARGB(255, Red, Green, Blue)
                Next A
                For A = NumParticles - Amt To NumParticles 'Line
                    ParticlesVertex(A).x = x + XOffset + ((A - (NumParticles - Amt)) * 32 \ Amt)
                    ParticlesVertex(A).y = y + 32
                    Particles(A).Life = Int(Rnd * 10) + 1
                    ParticlesVertex(A).Color = D3DColorARGB(255, Red, Green, Blue)
                Next A
            Case 3 'Explosion
                b = 0
                For A = 0 To NumParticles
                    Particles(A).y = y
                    Particles(A).x = x - 8 + Int(Rnd * 16)
                    Particles(A).G = Int(Rnd * 256)
                    ParticlesVertex(A).Color = D3DColorARGB(255, Particles(A).G, 0, 127)
                    Particles(A).Direction = 0.785398163397448 + (Rnd * 1.5707963267949)
                    Particles(A).VelocityX = VelocityX * (Rnd / 2)
                Next A
            Case 4 'heal type effect
                XOffset = x
                YOffset = y
                For A = 0 To NumParticles
                    Particles(A).x = Int(Rnd * 32)
                    Particles(A).y = Int(Rnd * 32)
                    Particles(A).Life = Life - Particles(A).y
                    ParticlesVertex(A).y = y + Particles(A).y
                    ParticlesVertex(A).x = x + Particles(A).x
                    ParticlesVertex(A).Color = D3DColorARGB(200, Red, Green, Blue)
                    Particles(A).VelocityY = VelocityY + Rnd
                Next A
            Case 5 'Swirling circles
                b = 0
                For A = 0 To NumParticles
                    Particles(A).Direction = Rnd * (3.14159265358979 * 2)
                    b = 1 - b
                    Particles(A).VelocityX = (Rnd / 2 + 0.5) * VelocityX
                    If b Then Particles(A).VelocityX = -Particles(A).VelocityX
                    Particles(A).VelocityY = (Rnd + 0.1) * Size1
                    ParticlesVertex(A).x = x + (Cos(Particles(A).Direction * Particles(A).VelocityX) * Particles(A).VelocityY)
                    ParticlesVertex(A).y = y + (Sin(Particles(A).Direction * Particles(A).VelocityX) * Particles(A).VelocityY)
                    ParticlesVertex(A).Color = D3DColorARGB(150, Red, Green, Blue)
                    Particles(A).Life = Life
                Next A
            Case 6 'Sucking Effect
                For A = 0 To NumParticles
                    Angle = Rnd * (3.14159265358979 * 2)
                    Particles(A).Direction = 0
                    Particles(A).Life = Rnd * Size1
                    Particles(A).VelocityX = Cos(Angle)
                    Particles(A).VelocityY = Sin(Angle)
                    ParticlesVertex(A).x = x + (Particles(A).VelocityX * Particles(A).Life)
                    ParticlesVertex(A).y = y + (Particles(A).VelocityY * Particles(A).Life)
                    ParticlesVertex(A).Color = D3DColorARGB(255, Red, Green, Blue)
                Next A
            Case 7 'Smoke
                BlendMode = D3DBLEND_INVSRCALPHA
                LifeLeft = Life
                Life = 60
                For A = 0 To NumParticles
                    Particles(A).Life = Int(Rnd * Life)
                    ParticlesVertex(A).x = x - (Size1 / 2) + (Rnd * Size1)
                    ParticlesVertex(A).y = y - (Particles(A).Life / 2)
                    Particles(A).VelocityY = ((Rnd / 2 + 0.5) * VelocityY)
                    ParticlesVertex(A).Color = D3DColorARGB(80, Red, Green, Blue)
                Next A
            Case 8 'Firesmoke
                LifeLeft = Life
                Life = 100
                For A = 0 To NumParticles
                    Particles(A).Life = Int(Rnd * Life)
                    ParticlesVertex(A).x = x - (Size1 / 2) + (Rnd * Size1)
                    ParticlesVertex(A).y = y - (Particles(A).Life / 2)
                    Particles(A).G = Int(Rnd * 256)
                    ParticlesVertex(A).Color = D3DColorARGB(127, 255, Particles(A).G, 0)
                    Particles(A).VelocityY = ((Rnd / 2 + 0.5) * VelocityY)
                    ParticlesVertex(A).Color = D3DColorARGB(80, Red, Green, Blue)
                Next A
            Case 9 'Quick Line
                For A = 0 To NumParticles
                    Particles(A).Life = Life + (Life * A) / NumParticles

                    ParticlesVertex(A).x = x + ((VelocityX - x) * A) / NumParticles
                    ParticlesVertex(A).y = y + ((VelocityY - y) * A) / NumParticles


                        ParticlesVertex(A).Color = D3DColorARGB(255, Red, Green, Blue)
                Next A
            Case 10 ' lotus cloud
                For A = 0 To NumParticles - Amt
                    Angle = Rnd * (3.14159265358979 * 2)
                    dA = Rnd * Size1
                    ParticlesVertex(A).x = x + (Cos(Angle) * dA)
                    ParticlesVertex(A).y = y + (Sin(Angle) * dA) - 16
                    ParticlesVertex(A).Color = D3DColorARGB(255, 0, 0, 0)
                    Particles(A).Life = Int(Life * 0.75)
                    Particles(A).Birth = Int(Life * 0.25)
                Next A
                For A = NumParticles - Amt To NumParticles
                    Angle = Rnd * (3.14159265358979 * 2)
                    ParticlesVertex(A).x = x + (Cos(Angle) * Size1)
                    ParticlesVertex(A).y = y + (Sin(Angle) * Size1) - 16
                    ParticlesVertex(A).Color = D3DColorARGB(255, 0, 0, 0)
                    Particles(A).Life = Int(Life * 0.75)
                    Particles(A).Birth = Int(Life * 0.25)
                Next A
            Case 11 'outward burst
                For A = 0 To NumParticles
                    Angle = Rnd * (3.14159265358979 * 2)
                    Particles(A).Direction = 0
                    Particles(A).Life = 0
                    Particles(A).VelocityX = Cos(Angle)
                    Particles(A).VelocityY = Sin(Angle)
                    ParticlesVertex(A).x = x + (Particles(A).VelocityX * Particles(A).Life)
                    ParticlesVertex(A).y = y + (Particles(A).VelocityY * Particles(A).Life)
                    ParticlesVertex(A).Color = D3DColorARGB(255, Red, Green, Blue)
                Next A
            Case 12 'fish
                For A = 0 To NumParticles
                    Angle = Rnd * (3.14159265358979 * 2)
                    Particles(A).b = 255
                    Particles(A).Direction = 0
                    If A Mod 2 = 0 Then
                        Particles(A).Life = -1 * Life / 2
                    Else
                        Particles(A).Life = 1
                    End If
                    Particles(A).VelocityX = Cos(Angle)
                    Particles(A).VelocityY = Sin(Angle)
                    ParticlesVertex(A).x = 0 'X + (Particles(A).VelocityX * Particles(A).Life)
                    ParticlesVertex(A).y = 0 'Y + (Particles(A).VelocityY * Particles(A).Life)
                    ParticlesVertex(A).Color = D3DColorARGB(Particles(A).b, Red, Green, Blue)
                Next A
        End Select
        
    End If
End Sub

Public Function Update() As Boolean
    
    Dim A As Long, RetVal As Boolean, Amt As Double
    
    Update = True
    A = GetTickCount / 25
    If A = LastParticleUpdate Then Exit Function ' until we redo this so rain position is based on time, this is a stop gap
    LastParticleUpdate = A

    
    
    RetVal = True
    
    If NumParticles > 0 Then
        Select Case SourceType
            Case 1 'Round Circle, fade in, fall down while fading out
                For A = 0 To NumParticles
                    If Particles(A).Birth > 0 Then
                        Particles(A).Birth = Particles(A).Birth - 1
                        ParticlesVertex(A).Color = D3DColorARGB((255 * (15 - Particles(A).Birth)) \ 15, Red, Green, Blue)
                    ElseIf Particles(A).Life > 0 Then
                        ParticlesVertex(A).x = ParticlesVertex(A).x + Int(Rnd * 3) - 1
                        ParticlesVertex(A).y = ParticlesVertex(A).y + VelocityY
                        Particles(A).Life = Particles(A).Life - 1
                        ParticlesVertex(A).Color = D3DColorARGB((255 * (Particles(A).Life)) \ 50, Red, Green, Blue)
                    End If
                Next A
            Case 2 ' Line, going up, dropping particles
                Amt = NumParticles * 0.4
                YOffset = YOffset - 0.64
                For A = 0 To NumParticles - Amt 'Falling Particles
                    Particles(A).Life = Particles(A).Life - 1
                    ParticlesVertex(A).Color = D3DColorARGB(127 + 2 * (50 - Particles(A).Life), Red, Green, Blue)
                    If Particles(A).Life = 0 Then
                        ParticlesVertex(A).y = y + 32 + YOffset
                        Particles(A).Life = Int(Rnd * 50) + 1
                        ParticlesVertex(A).x = x + Int(Rnd * 32)
                    End If
                Next A
                For A = NumParticles - Amt To NumParticles 'Line
                    ParticlesVertex(A).y = y + 32 + YOffset
                    Particles(A).Life = Particles(A).Life - Int(Rnd * 1)
                    ParticlesVertex(A).Color = D3DColorARGB(127 + 2 * (10 - Particles(A).Life), Red, Green, Blue)
                    If Particles(A).Life = 0 Then
                        ParticlesVertex(A).y = y + 32 + YOffset
                        Particles(A).Life = Int(Rnd * 10) + 1
                        ParticlesVertex(A).x = x + Int(Rnd * 32)
                    End If
                Next A
                Life = Life - 1
                If Life = 0 Then RetVal = False
            Case 3 'Explosion
                For A = 0 To NumParticles
                    If Particles(A).Life < Life Then
                        ParticlesVertex(A).y = y - (Sin(Particles(A).Direction) * Particles(A).VelocityX * Particles(A).Life)
                        ParticlesVertex(A).x = Particles(A).x + (Cos(Particles(A).Direction) * Particles(A).VelocityX * Particles(A).Life)
                        ParticlesVertex(A).Color = D3DColorARGB(127 - 127 * (Particles(A).Life / Life), 255, Particles(A).G, 0)
                        Particles(A).Life = Particles(A).Life + 1
                    Else
                        Life = 0
                        Exit For
                    End If
                Next A
                If Life = 0 Then RetVal = False
                
            Case 4 'Heal type effect, only works for "up"
                Life = Life - 1
                If Life = 0 Then
                    RetVal = False
                Else
                    Select Case TargType
                        Case TT_CHARACTER
                            XOffset = Cxo
                            YOffset = CYO - 16 + y
                        Case TT_PLAYER
                            If Target > 0 Then
                                XOffset = player(Target).XO
                                YOffset = player(Target).YO - 16 + y
                            End If
                    End Select
                    For A = 0 To NumParticles
                        Particles(A).y = Particles(A).y + Particles(A).VelocityY
                        Particles(A).x = Particles(A).x + Particles(A).VelocityX
                        ParticlesVertex(A).y = YOffset + Particles(A).y
                        ParticlesVertex(A).x = XOffset + Particles(A).x
                        Particles(A).Life = Particles(A).Life - Particles(A).VelocityY
                        If Particles(A).Life <= 0 Then
                            Particles(A).Life = Life
                            Particles(A).y = 0
                            'ParticlesVertex(A).Color = D3DColorARGB(0, 0, 0, 0)
                        Else
                            ParticlesVertex(A).Color = D3DColorARGB(200 * (Particles(A).Life / 32), Red, Green, Blue)
                        End If
                    Next A
                End If
            Case 5 'Swirling circles
                'Life = Life - 1
                If Life = 0 Then
                    RetVal = False
                Else
                    For A = 0 To NumParticles
                        Particles(A).Direction = Particles(A).Direction + Particles(A).VelocityX
                        ParticlesVertex(A).x = x + (Cos(Particles(A).Direction) * Particles(A).VelocityY)
                        ParticlesVertex(A).y = y + (Sin(Particles(A).Direction) * Particles(A).VelocityY)
                        ParticlesVertex(A).Color = D3DColorARGB(150 * (Particles(A).Life / Life), Red, Green, Blue)
                        Particles(A).Life = Particles(A).Life - 1
                        If Particles(A).Life = 0 Then
                            Life = 0
                            Exit For
                        End If
                    Next A
                End If
            Case 6 'Sucking thing (Black Hole)
                If Life > 1 Then Life = Life - 1
                If Life = 1 Then
                    RetVal = False
                Else
                    For A = 0 To NumParticles
                        Particles(A).Direction = Particles(A).Direction + VelocityX
                        Particles(A).Life = Particles(A).Life - Particles(A).Direction
                        If Particles(A).Life <= 0 Then
                            Particles(A).Life = Rnd * Size1
                            Particles(A).Direction = 0
                        End If
                        ParticlesVertex(A).x = x + (Particles(A).VelocityX * Particles(A).Life)
                        ParticlesVertex(A).y = y + (Particles(A).VelocityY * Particles(A).Life)
                    Next A
                End If
            Case 7 'Smoke
                If LifeLeft > 1 Then LifeLeft = LifeLeft - 1
                If LifeLeft = 1 Then RetVal = False
                
                    For A = 0 To NumParticles
                        ParticlesVertex(A).y = ParticlesVertex(A).y + Particles(A).VelocityY
                        Particles(A).Life = Particles(A).Life + Particles(A).VelocityY
                        If Particles(A).Life >= 0 Then
                            RetVal = True
                        Else
                            If LifeLeft <> 1 Then
                                Particles(A).Life = Life
                                Particles(A).VelocityY = ((Rnd + 1) / 2) * VelocityY
                                ParticlesVertex(A).y = y
                            Else
                                Particles(A).Life = 0
                            End If
                        End If
                        ParticlesVertex(A).Color = D3DColorARGB(80 * (Particles(A).Life / Life), Red, Green, Blue)
                    Next A
            Case 8 'Smoke Fire
                If LifeLeft > 1 Then LifeLeft = LifeLeft - 1
                If LifeLeft = 1 Then RetVal = False
                
                    For A = 0 To NumParticles
                        ParticlesVertex(A).y = ParticlesVertex(A).y + Particles(A).VelocityY
                        Particles(A).Life = Particles(A).Life + Particles(A).VelocityY * 2
                        If Particles(A).Life >= 0 Then
                            RetVal = True
                        Else
                            If LifeLeft <> 1 Then
                                Particles(A).Life = Life
                                Particles(A).VelocityY = ((Rnd + 1) / 2) * VelocityY
                                ParticlesVertex(A).y = y
                            Else
                                Particles(A).Life = 0
                            End If
                        End If
                        ParticlesVertex(A).Color = D3DColorARGB(80 * (Particles(A).Life / Life), 255, Particles(A).G, Blue)
                    Next A
            Case 9 '...
                RetVal = True
                For A = 0 To NumParticles
                    
                    
                    If (Particles(NumParticles).Life = 0) Then
                        RetVal = False
                    Else
                        ParticlesVertex(A).Color = D3DColorARGB(255, Red, Green, Blue)
                    End If
                Next A
            Case 10 ' lotus cloud
                For A = 0 To NumParticles
                    If Particles(A).Birth > 0 Then
                        Particles(A).Birth = Particles(A).Birth - 1
                        ParticlesVertex(A).Color = D3DColorARGB((255 * (15 - Particles(A).Birth)) \ 15, Red, Green, Blue)
                    ElseIf Particles(A).Life > 0 Then
                        ParticlesVertex(A).x = ParticlesVertex(A).x + Int(Rnd * 3) - 1
                        ParticlesVertex(A).y = ParticlesVertex(A).y + VelocityY
                        Particles(A).Life = Particles(A).Life - 1
                        ParticlesVertex(A).Color = D3DColorARGB((255 * (Particles(A).Life)) \ 50, Red, Green, Blue)
                    End If
                Next A
            Case 11 'outward burst
                If Life > 1 Then Life = Life - 1
                    If Life = 1 Then
                        RetVal = False
                    Else
                        For A = 0 To NumParticles
                            Particles(A).Direction = Particles(A).Direction + VelocityX
                            Particles(A).Life = Particles(A).Life + Particles(A).Direction
                            If Particles(A).Life >= Size1 Then
                                Particles(A).Life = 0
                                Particles(A).Direction = 0
                            End If
                            ParticlesVertex(A).x = x + (Particles(A).VelocityX * Particles(A).Life)
                            ParticlesVertex(A).y = y + (Particles(A).VelocityY * Particles(A).Life)
                        Next A
                    End If
            Case 12 'fish
                If Life > 1 Then Life = Life - 1
                    If Life = 1 Then
                        RetVal = False
                    Else
                        For A = 0 To NumParticles
                            If Particles(A).Life >= 0 Then
                                If Particles(A).Life <> 0 Then
                                    Particles(A).Direction = Particles(A).Direction + VelocityX
                                    Particles(A).Life = Particles(A).Life + Particles(A).Direction
                                    If Particles(A).Life >= Size1 Then
                                        Particles(A).b = Particles(A).b - 25
                                        ParticlesVertex(A).Color = D3DColorARGB(Particles(A).b, Red, Green, Blue)
                                        If Particles(A).b <= 0 Then
                                            Particles(A).b = Particles(A).b = 0
                                            Particles(A).Life = 0
                                            Particles(A).Direction = 0
                                        End If
                                    End If
                                    ParticlesVertex(A).x = x + (Particles(A).VelocityX * Particles(A).Life)
                                    ParticlesVertex(A).y = y + (Particles(A).VelocityY * Particles(A).Life) * 0.8
                                End If
                            Else
                                Particles(A).Life = Particles(A).Life + 1
                                If Particles(A).Life = 0 Then Particles(A).Life = 1
                            End If
                        Next A
                    End If
        End Select
    End If
    
    Update = RetVal
End Function


Public Sub Render()
    If NumParticles > 0 Then
    With D3DDevice
        

        .SetRenderState D3DRS_POINTSIZE, ParticleSize
        .SetVertexShader D3DFVF_TLVERTEX

                LastTexture = -1
        .SetTexture 0, texParticles(ParticleNum)
        
        .SetRenderState D3DRS_DESTBLEND, BlendMode
        
        
        .DrawPrimitiveUP D3DPT_POINTLIST, NumParticles + 1, ParticlesVertex(0), Len(ParticlesVertex(0))
    
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        
        End With
    End If
End Sub
