Attribute VB_Name = "modDirectX"
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

#Const UseFlags = False
'Bitmap Information Stuff
Public Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Public Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
'End Bitmap Shit

Public Type SurfaceType
    Surface As DirectDrawSurface7
    desc As DDSURFACEDESC2
    Width As Long
    Height As Long
End Type

Dim frameOffsetsX(0 To 11) As Integer
Dim frameOffsetsY(0 To 11) As Integer
Dim frameOffsetsX2(0 To 11) As Integer
Dim frameOffsetsY2(0 To 11) As Integer

Private Trans As Long

Public Dx7 As New DirectX7
Public DD As DirectDraw7
Public sfcObjects(1 To 8) As SurfaceType
Public sfcSymbols(1 To 3) As SurfaceType
Public sfcSprites As SurfaceType
Public sfcSprites2 As SurfaceType
Public sfcObjectFrames As SurfaceType
Public sfcLSprites As SurfaceType

Public sfcInventory(1) As SurfaceType
Public sfcInventory2 As SurfaceType
Public sfcChatTabs As SurfaceType
Public sfcTimers As SurfaceType
Public sfcTimers2 As SurfaceType
Public sfcCursor As SurfaceType



Public Clipper As DirectDrawClipper

'Transformable lit vertex
Public Const D3DFVF_TLVERTEX = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
Public Type TLVERTEX
    x As Single
    y As Single
    z As Single
    rhw As Single
    Color As Long
    tu As Single
    tv As Single
End Type

'The size of a FVF vertex
Public Const FVF_Size As Long = 28

'Describes the return from a texture init
Public Type D3DXIMAGE_INFO_A
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type


'********** Direct X ***********
Public SurfaceTimer() As Long                   'How long until the surface unloads
Public LastTexture As Long                      'The last texture used
Public SwapChain(1) As Direct3DSwapChain8
Public D3DWindow(1) As D3DPRESENT_PARAMETERS       'Describes the viewport and used to restore when in fullscreen
Public UsedCreateFlags As CONST_D3DCREATEFLAGS  'The flags we used to create the device when it first succeeded
Public DispMode As D3DDISPLAYMODE               'Describes the display mode
Public RenderSurface(1) As Direct3DSurface8

'DirectX 8 Objects
Public DX8 As DirectX8
Public D3D As Direct3D8
Public D3DX As D3DX8
Public D3DDevice As Direct3DDevice8

Public mapTexture(1 To 2) As Direct3DTexture8

Public Type TextureType
    Height As Long
    Width As Long
    TexInfo As D3DXIMAGE_INFO_A
    Texture As Direct3DTexture8
    Loaded As Boolean
    LastUsed As Long
    Filename As String
    Encrypted As Boolean
    Filter As Long
End Type

Public Const TextureWaitTime As Long = 300000

Public Const NumTiles As Long = 40
Public Const NumSprites As Long = 16 * 2 'max = 16x2
Public Const NumLSprites As Long = 8 'max = 64
Public Const NumProjectiles As Long = 5
Public Const NumEffects As Long = 13
Public Const NumObjects As Long = 8
Public Const NumObjectFrames = 16
Public Const NumLights As Long = 3
Public Const NumSymbols As Long = 3
Public Const NumControls As Long = 1
Public Const NumPortraits As Long = 50
Public Const AnimationFps = 25


Public CurTexture As Long
Public Const TexTile0 As Long = 0
Public Const TexSprite0 As Long = NumTiles * 2
Public Const TexProjectile0 As Long = TexSprite0 + NumSprites
Public Const TexEffect0 As Long = TexProjectile0 + NumProjectiles
Public Const TexObject0 As Long = TexEffect0 + NumEffects
Public Const TexObjectFrames0 As Long = TexObject0 + NumObjects
Public Const TexLight0 As Long = TexObjectFrames0 + NumObjectFrames
Public Const TexPortrait0 As Long = TexLight0 + NumLights
Public Const TexSymbol0 = TexPortrait0 + NumPortraits
Public Const TexLSprite0 As Long = TexSymbol0 + NumSymbols
Public Const TexControl1 As Long = TexLSprite0 + NumLSprites + 1

Public Const TexINVALID As Long = 1000


Public Const NumTextures = TexControl1

Public DynamicTextures(1 To NumTextures) As TextureType

Public TexFont As TextureType
Public texAtts As TextureType
Public texShade As TextureType
Public texShadeEX As TextureType
Public texLightsEX As TextureType

Public Const NumParticleTextures As Long = 10
Public texParticles(1 To NumParticleTextures) As Direct3DTexture8

'Border Style Constants
Public Const BS_NONE = 0
Public Const BS_SOLID = 1

Public Sub CheckFiles()
End Sub



Public Function Init_Directx()
    Set DD = Dx7.DirectDrawCreate("")
    Call DD.SetCooperativeLevel(frmMain.hwnd, DDSCL_NORMAL)
    InitializeSurfaces
End Function

Public Sub InitializeSurfaces()
    Dim CK As DDCOLORKEY
    CK.high = 0
    CK.low = 0
    Dim A As Long
    
    frameOffsetsX(0) = 17
    frameOffsetsY(0) = -9
    frameOffsetsX(1) = 17
    frameOffsetsY(1) = -7
    frameOffsetsX(2) = 13
    frameOffsetsY(2) = -12

    frameOffsetsX(3) = -8
    frameOffsetsY(3) = -7
    frameOffsetsX(4) = -8
    frameOffsetsY(4) = -9
    frameOffsetsX(5) = 6
    frameOffsetsY(5) = -10

    frameOffsetsX(6) = -5
    frameOffsetsY(6) = -2
    frameOffsetsX(7) = -12
    frameOffsetsY(7) = -3
    frameOffsetsX(8) = -14
    frameOffsetsY(8) = -3
    
    frameOffsetsX(9) = 3
    frameOffsetsY(9) = -2
    frameOffsetsX(10) = 7
    frameOffsetsY(10) = -1
    frameOffsetsX(11) = 13
    frameOffsetsY(11) = -3
    

    frameOffsetsX2(0) = 20
    frameOffsetsY2(0) = -8
    frameOffsetsX2(1) = 20
    frameOffsetsY2(1) = -10
    frameOffsetsX2(2) = 24
    frameOffsetsY2(2) = -12

    frameOffsetsX2(3) = 42
    frameOffsetsY2(3) = -9
    frameOffsetsX2(4) = 41
    frameOffsetsY2(4) = -7
    frameOffsetsX2(5) = 30
    frameOffsetsY2(5) = -12

    frameOffsetsX2(6) = 35
    frameOffsetsY2(6) = -1
    frameOffsetsX2(7) = 43
    frameOffsetsY2(7) = -3
    frameOffsetsX2(8) = 45
    frameOffsetsY2(8) = -2
    
    frameOffsetsX2(9) = 28
    frameOffsetsY2(9) = -2
    frameOffsetsX2(10) = 23
    frameOffsetsY2(10) = -2
    frameOffsetsX2(11) = 17
    frameOffsetsY2(11) = -2
    



    For A = 0 To 1
        sfcInventory(A).desc.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        sfcInventory(A).desc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Next A
    sfcInventory(1).desc.lWidth = 143
    sfcInventory(1).desc.lHeight = 179
    Set sfcInventory(0).Surface = DD.CreateSurfaceFromFile("Data/Graphics/Interface/Inventory.rsc", sfcInventory(0).desc)
    Set sfcInventory(1).Surface = DD.CreateSurface(sfcInventory(1).desc)



    sfcInventory(0).Surface.SetColorKey DDCKEY_SRCBLT, CK
    sfcInventory(1).Surface.SetColorKey DDCKEY_SRCBLT, CK
    sfcInventory2.desc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    sfcInventory2.desc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    sfcInventory2.desc.lWidth = 190
    sfcInventory2.desc.lHeight = 370
    Set sfcInventory2.Surface = DD.CreateSurface(sfcInventory2.desc)
    sfcInventory2.Surface.SetColorKey DDCKEY_SRCBLT, CK

    
    sfcChatTabs.desc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    sfcChatTabs.desc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    sfcChatTabs.desc.lWidth = 54
    sfcChatTabs.desc.lHeight = 155
    Set sfcChatTabs.Surface = DD.CreateSurface(sfcChatTabs.desc)
    sfcChatTabs.Surface.SetColorKey DDCKEY_SRCBLT, CK
    
        sfcTimers.desc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    sfcTimers.desc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    sfcTimers.desc.lWidth = 78
    sfcTimers.desc.lHeight = 36
    Set sfcTimers.Surface = DD.CreateSurface(sfcTimers.desc)
    sfcTimers.Surface.SetColorKey DDCKEY_SRCBLT, CK

        sfcTimers2.desc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    sfcTimers2.desc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    sfcTimers2.desc.lWidth = 64
    sfcTimers2.desc.lHeight = 32
    Set sfcTimers2.Surface = DD.CreateSurface(sfcTimers2.desc)
    sfcTimers2.Surface.SetColorKey DDCKEY_SRCBLT, CK
        
    For A = 1 To NumObjects
        sfcObjects(A).desc.lFlags = DDSD_CAPS
        sfcObjects(A).desc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        EncryptFile "Data/Graphics/Object" & Format(A, "000") & ".rsc", "Data/Cache/Temp.ts"
        DoEvents
        Set sfcObjects(A).Surface = DD.CreateSurfaceFromFile("Data/Cache/Temp.ts", sfcObjects(A).desc)
        'Set sfcObjects(A).Surface = DD.CreateSurfaceFromFile("Data/Graphics/Object" & Format(A, "000") & ".rsc", sfcObjects(A).desc)
        sfcObjects(A).Surface.SetColorKey DDCKEY_SRCBLT, CK
    Next A
    For A = 1 To 3
        sfcSymbols(A).desc.lFlags = DDSD_CAPS
        sfcSymbols(A).desc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        'EncryptFile "Data/Graphics/Symbol" & A & ".rsc", "Data/Cache/Temp.ts"
        'DoEvents
        sfcSymbols(A).Width = 320
        sfcSymbols(A).Height = 320
        Set sfcSymbols(A).Surface = DD.CreateSurfaceFromFile("Data/Graphics/Symbol" & A & ".rsc", sfcSymbols(A).desc)
        'Set sfcObjects(A).Surface = DD.CreateSurfaceFromFile("Data/Graphics/Object" & Format(A, "000") & ".rsc", sfcObjects(A).desc)
        sfcSymbols(A).Surface.SetColorKey DDCKEY_SRCBLT, CK
    Next A
    Kill "Data/Cache/Temp.ts"

    'Load Sprites
    sfcSprites.desc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    sfcSprites.desc.lWidth = 512
    sfcSprites.desc.lHeight = 512
    sfcSprites.desc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set sfcSprites.Surface = DD.CreateSurfaceFromFile("Data/Graphics/Sprites.rsc", sfcSprites.desc)
    sfcSprites.Surface.SetColorKey DDCKEY_SRCBLT, CK
    
    If Exists("Data/Graphics/Sprites2.rsc") Then
        sfcSprites2.desc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        sfcSprites2.desc.lWidth = 512
        sfcSprites2.desc.lHeight = 512
        sfcSprites2.desc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set sfcSprites2.Surface = DD.CreateSurfaceFromFile("Data/Graphics/Sprites2.rsc", sfcSprites2.desc)
        sfcSprites2.Surface.SetColorKey DDCKEY_SRCBLT, CK
    Else
        Set sfcSprites2.Surface = Nothing
    End If
    
    If Exists("Data/Graphics/ObjectFramePreview.rsc") Then
        sfcObjectFrames.desc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        sfcObjectFrames.desc.lWidth = 512
        sfcObjectFrames.desc.lHeight = 512
        sfcObjectFrames.desc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set sfcObjectFrames.Surface = DD.CreateSurfaceFromFile("Data/Graphics/ObjectFramePreview.rsc", sfcObjectFrames.desc)
        sfcObjectFrames.Surface.SetColorKey DDCKEY_SRCBLT, CK
    End If
    
    
    If Exists("Data/Graphics/LSprites.rsc") Then
        sfcLSprites.desc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        sfcLSprites.desc.lWidth = 1024
        sfcLSprites.desc.lHeight = 1024
        sfcLSprites.desc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set sfcLSprites.Surface = DD.CreateSurfaceFromFile("Data/Graphics/LSprites.rsc", sfcLSprites.desc)
        sfcLSprites.Surface.SetColorKey DDCKEY_SRCBLT, CK
    Else
        Set sfcLSprites.Surface = Nothing
    End If
        
    
    sfcCursor.desc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    sfcCursor.desc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    sfcCursor.desc.lWidth = 64
    sfcCursor.desc.lHeight = 32
    Set sfcCursor.Surface = DD.CreateSurface(sfcCursor.desc)
    sfcCursor.Surface.SetColorKey DDCKEY_SRCBLT, CK
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

Public Sub Draw(Dest As SurfaceType, DestX As Long, DestY As Long, nWidth As Long, nHeight As Long, Source As SurfaceType, srcX As Long, srcY As Long, Optional Trans As Boolean = False, Optional TransColor As Long = 0)
    On Local Error GoTo errOut
    
    Dim SrcRect As RECT, DestRect As RECT
    
    ' Update RECTs
    With SrcRect
        .Top = srcY
        .Bottom = srcY + nHeight
        .Left = srcX
        .Right = srcX + nWidth
    End With
    With DestRect
        .Top = DestY
        .Bottom = DestY + nHeight
        .Left = DestX
        .Right = DestX + nWidth
    End With
    
    ' Check Boundaries
    If DestX < 0 Then
        DestRect.Left = 0
        SrcRect.Left = SrcRect.Left - DestX
    End If
    If DestY < 0 Then
        DestRect.Top = 0
        SrcRect.Top = SrcRect.Top - DestY
    End If
    
    If DestRect.Right > Dest.desc.lWidth Then
        SrcRect.Right = SrcRect.Right - (DestRect.Right - Dest.desc.lWidth)
        DestRect.Right = Dest.desc.lWidth
    End If
    If DestRect.Bottom > Dest.desc.lHeight Then
        SrcRect.Bottom = SrcRect.Bottom - (DestRect.Bottom - Dest.desc.lHeight)
        DestRect.Bottom = Dest.desc.lHeight
    End If

    If SrcRect.Right > Source.desc.lWidth Then
        DestRect.Right = DestRect.Right - (SrcRect.Right - SrcRect.Left - Source.desc.lWidth)
        SrcRect.Right = Source.desc.lWidth
    End If
    If SrcRect.Bottom > Source.desc.lHeight Then
        DestRect.Bottom = DestRect.Bottom - (SrcRect.Bottom - SrcRect.Top - Source.desc.lHeight)
        SrcRect.Bottom = Source.desc.lHeight
    End If
    
    
    If Not Trans Then
        Call Dest.Surface.BltFast(DestRect.Left, DestRect.Top, Source.Surface, SrcRect, DDBLTFAST_WAIT)
    Else
        If TransColor > 0 Then
            Dim CK As DDCOLORKEY
            CK.high = TransColor
            CK.low = TransColor
            Source.Surface.SetColorKey DDCKEY_SRCBLT, CK
        End If
        'Call Dest.Surface.Blt(DestRect, Source.Surface, SrcRect, DDBLT_KEYSRC)
        'Dim fx As DDBLTFX
        'fx.lDDFX = DDBLTFX_MIRRORLEFTRIGHT
        
        'Call Dest.Surface.bltfx(DestRect, Source.Surface, SrcRect, DDBLT_KEYSRC, fx)
        Call Dest.Surface.BltFast(DestRect.Left, DestRect.Top, Source.Surface, SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        
        If TransColor > 0 Then
            CK.high = 0
            CK.low = 0
            Source.Surface.SetColorKey DDCKEY_SRCBLT, CK
        End If
    End If

errOut:
End Sub

Public Sub LoadFontData()
    Dim St As String * 192, A As Long
    Dim CurXPos As Long, CurYPos As Long
    
        Open "Data\Cache\Font1.dat" For Binary Access Read As #1
            If LOF(1) <> 192 Then Exit Sub
            Get #1, , St
        Close #1
    
        For A = 1 To 96
            With FontChar(1, 31 + A)
                .Width = Val(Mid$(St, (A - 1) * 2 + 1, 2))
                If CurXPos + .Width > 512 Then
                    CurXPos = 0
                    CurYPos = CurYPos + 48
                End If
                .srcX = CurXPos
                CurXPos = CurXPos + .Width
                .srcY = CurYPos
            End With
        Next A

        Open "Data\Cache\Font2.dat" For Binary Access Read As #1
            If LOF(1) <> 192 Then
                Exit Sub
            End If
            Get #1, , St
        Close #1
        
        Dim SpaceOffsetX As Long
        Dim SpaceOffsetY As Long
        CurXPos = 0
        CurYPos = 256
        SpaceOffsetX = 1
        SpaceOffsetY = 1
        For A = 1 To 96
            With FontChar(2, 31 + A)
                .Width = Val(Mid$(St, (A - 1) * 2 + 1, 2))
                If CurXPos + .Width > 256 Then
                    CurXPos = 0
                    SpaceOffsetX = 0
                    CurYPos = CurYPos + 18
                End If
                .srcX = CurXPos + SpaceOffsetX
                SpaceOffsetX = SpaceOffsetX + 2
                CurXPos = CurXPos + .Width
                .srcY = CurYPos
            End With
        Next A
End Sub

Public Sub DrawUnzStrings()
    Dim A As Long, CurXPos As Long, b As Long
    Dim VertexArray(0 To 3) As TLVERTEX

    If NumUnzText > 0 Then
        D3DDevice.SetTexture 0, TexFont.Texture
        VertexArray(0).rhw = 1
        VertexArray(1).rhw = 1
        VertexArray(2).rhw = 1
        VertexArray(3).rhw = 1
        For b = 0 To NumUnzText - 1
            With UnzText(b)
                CurXPos = .x
                For A = 1 To Len(.Text)
                    If Asc(Mid$(.Text, A, 1)) = 32 Then
                        CurXPos = CurXPos + 20
                    Else
                        With FontChar(1, Asc(Mid$(.Text, A, 1)))
                             VertexArray(0).tu = .srcX / 512
                             VertexArray(1).tu = (.srcX + .Width) / 512
                             VertexArray(0).tv = (.srcY) / 512
                             VertexArray(2).tv = (.srcY + FontHeight) / 512
                             VertexArray(1).tv = VertexArray(0).tv
                             VertexArray(2).tu = VertexArray(0).tu
                             VertexArray(3).tu = VertexArray(1).tu
                             VertexArray(3).tv = VertexArray(2).tv
                            
                             'Draw Shadow Letter
                             VertexArray(0).Color = D3DColorARGB(255 - (UnzText(b).Fade * 5), 127, 0, 0)
                             VertexArray(1).Color = D3DColorARGB(255 - (UnzText(b).Fade * 5), 127, 0, 0)
                             VertexArray(2).Color = D3DColorARGB(255 - (UnzText(b).Fade * 5), 127, 0, 0)
                             VertexArray(3).Color = D3DColorARGB(255 - (UnzText(b).Fade * 5), 127, 0, 0)
                             VertexArray(0).x = CurXPos + 1.5
                             VertexArray(0).y = UnzText(b).y + 1.5
                             VertexArray(1).x = CurXPos + .Width + 1.5
                             VertexArray(2).x = VertexArray(0).x
                             VertexArray(3).x = VertexArray(1).x
                             VertexArray(2).y = UnzText(b).y + FontHeight + 1.5
                             VertexArray(1).y = VertexArray(0).y
                             VertexArray(3).y = VertexArray(2).y
                             D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
                             
                             'Draw Colored Text
                             VertexArray(0).Color = D3DColorARGB(255 - (UnzText(b).Fade * 5), 255, 255, 255)
                             VertexArray(1).Color = D3DColorARGB(255 - (UnzText(b).Fade * 5), 255, 255, 255)
                             VertexArray(2).Color = D3DColorARGB(255 - (UnzText(b).Fade * 5), 255, 255, 255)
                             VertexArray(3).Color = D3DColorARGB(255 - (UnzText(b).Fade * 5), 255, 255, 255)
                             VertexArray(0).x = CurXPos - 0.5
                             VertexArray(0).y = UnzText(b).y - 0.5
                             VertexArray(1).x = CurXPos + .Width - 0.5
                             VertexArray(2).x = VertexArray(0).x
                             VertexArray(3).x = VertexArray(1).x
                             VertexArray(2).y = UnzText(b).y + FontHeight - 0.5
                             VertexArray(1).y = VertexArray(0).y
                             VertexArray(3).y = VertexArray(2).y
                             D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
                            CurXPos = CurXPos + .Width - 5
                        End With
                    End If
                Next A
                
                On Error Resume Next
                If .Lifetime = 0 Then
                    If .Lifetime = 0 Then
                        .Fade = .Fade + 1
                        If .Fade = 51 Then 'DESTROY
                            If b < NumUnzText - 1 Then
                                .Text = UnzText(b + 1).Text
                                .Fade = UnzText(b + 1).Fade
                                .Lifetime = UnzText(b + 1).Lifetime
                                .x = UnzText(b + 1).x
                                .y = UnzText(b + 1).y
                                ReDim Preserve UnzText(b)
                                b = b - 1
                            Else
                                If b > 0 Then ReDim Preserve UnzText(b - 1)
                                NumUnzText = 0
                            End If
                        End If
                    End If
                Else
                    .Lifetime = .Lifetime - 1
                End If
            End With
        Next b
    End If
End Sub

Public Function Engine_Init_D3DDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS) As Boolean
    On Error GoTo errOut
    
    
    D3DWindow(0).EnableAutoDepthStencil = 1
    D3DWindow(0).AutoDepthStencilFormat = D3DFMT_D16
    D3DWindow(0).hDeviceWindow = frmMain.picViewport.hwnd
    
    
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    D3DWindow(0).Windowed = 1
    D3DWindow(0).SwapEffect = D3DSWAPEFFECT_DISCARD
    D3DWindow(0).BackBufferFormat = DispMode.Format
    D3DWindow(0).BackBufferHeight = 384
    D3DWindow(0).BackBufferWidth = 384

    If (Options.DisableMultiSampling = False And Options.ResolutionIndex <> 1) Then
    Dim A As Long, b As Long
    For A = 0 To 20
    If D3D.CheckDeviceMultiSampleType(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DispMode.Format, 1, A) = 0 Then b = A
    Next A
    D3DWindow(0).MultiSampleType = b
    Else
        D3DWindow(0).MultiSampleType = D3DMULTISAMPLE_NONE
    End If

    
    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.picViewport.hwnd, D3DCREATEFLAGS, D3DWindow(0))
    UsedCreateFlags = D3DCREATEFLAGS
    
    D3DWindow(1).Windowed = 1
    
    D3DWindow(1).SwapEffect = D3DSWAPEFFECT_COPY
    If (Options.VsyncEnabled) Then
        D3DWindow(1).SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    End If
    
    D3DWindow(1).BackBufferFormat = DispMode.Format
    D3DWindow(1).hDeviceWindow = frmMain.picMiniMap.hwnd
    Set SwapChain(1) = D3DDevice.CreateAdditionalSwapChain(D3DWindow(1))
    
    

  
    Set RenderSurface(0) = D3DDevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)
    Set RenderSurface(1) = SwapChain(1).GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)
       
        
    
    Engine_Init_D3DDevice = True

Exit Function

errOut:
'    MsgBox Err.Description
    
    Set D3DDevice = Nothing
    Engine_Init_D3DDevice = False
End Function


Function InitD3D(Initial As Boolean)
Trans = -1
If Initial Then
    Set DX8 = New DirectX8
    Set D3D = DX8.Direct3DCreate()
    Set D3DX = New D3DX8
End If
    
    If Not Engine_Init_D3DDevice(D3DCREATE_PUREDEVICE) Then
        If Not Engine_Init_D3DDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
            If Not Engine_Init_D3DDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                If Not Engine_Init_D3DDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
                    If (Initial) Then
                        MsgBox "Could not Create D3D Device"
                        InitD3D = False
                        Exit Function
                    End If
                End If
            End If
        End If
    End If

    
    With D3DDevice
        Call .SetVertexShader(D3DFVF_TLVERTEX)
        Call .SetRenderState(D3DRS_LIGHTING, False)
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        Call .SetRenderState(D3DRS_ZENABLE, False)
        Call .SetRenderState(D3DRS_ZWRITEENABLE, False)
        Call .SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
        Call .SetRenderState(D3DRS_FILLMODE, CONST_D3DFILLMODE.D3DFILL_SOLID)
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, 1)
        
'        Call .SetRenderState(D3DRS_MULTISAMPLE_ANTIALIAS, 1)
'        Call .SetRenderState(D3DRS_MULTISAMPLE_MASK, 1)
'
'        Call .SetRenderState(D3DRS_EDGEANTIALIAS, 1)


        
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1 'True
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0 'True
        .SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
        .SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
        .SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE

        Call .SetTextureStageState(0, D3DTSS_MINFILTER, D3DTEXF_POINT)
        Call .SetTextureStageState(0, D3DTSS_MAGFILTER, D3DTEXF_POINT)
    End With
    InitializeTextures
    InitD3D = True
End Function

Sub InitializeTextures()
    Dim A As Long, tex() As Byte, TexLen As Long
    
    TexLen = DecryptFile("Data/Graphics/Font.rsc", tex())
    Set TexFont.Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, tex(0), TexLen, D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, TexFont.TexInfo, ByVal 0)
    TexFont.Loaded = True
    
    For A = 1 To NumTextures
        DynamicTextures(A).Filter = D3DX_FILTER_POINT
    Next A
    
    For A = 1 To NumTiles
        DynamicTextures(TexTile0 + A).Filename = "Data/Graphics/Tile" & Format(A, "000") & ".rsc"
        DynamicTextures(TexTile0 + A).Encrypted = True
        'TexLen = DecryptFile("Data/Graphics/Tile" & Format(A, "000") & ".rsc", Tex)
        'Set texTiles(A).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, Tex(0), TexLen, 256, 256, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texTiles(A).TexInfo, ByVal 0)
    Next A
    For A = 1 To NumSprites
        DynamicTextures(TexSprite0 + A).Filename = "Data/Graphics/Sprite" & Format(A, "000") & ".rsc"
        DynamicTextures(TexSprite0 + A).Encrypted = True
        'TexLen = DecryptFile("Data/Graphics/Sprite" & Format(A, "000") & ".rsc", Tex)
        'Set texSprites(A).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, Tex(0), TexLen, 512, 512, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texSprites(A).TexInfo, ByVal 0)
    Next A
    For A = 1 To 5
        DynamicTextures(TexProjectile0 + A).Filename = "Data/Graphics/Projectile" & Format(A, "000") & ".rsc"
        DynamicTextures(TexProjectile0 + A).Encrypted = True
        'Set texProjectiles(A).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Projectile" & Format(A, "000") & ".rsc", 128, 128, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texProjectiles(A).TexInfo, ByVal 0)
    Next A
    For A = 1 To 3
        DynamicTextures(TexSymbol0 + A).Filename = "Data/Graphics/Symbol" & A & ".rsc"
        DynamicTextures(TexSymbol0 + A).Encrypted = False
        DynamicTextures(TexLight0 + A).Width = 320
        DynamicTextures(TexLight0 + A).Height = 320
        'DynamicTextures(TexLight0 + A).Filter = D3DX_FILTER_LINEAR
        'Set texProjectiles(A).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Projectile" & Format(A, "000") & ".rsc", 128, 128, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texProjectiles(A).TexInfo, ByVal 0)
    Next A
    For A = 1 To 13
        DynamicTextures(TexEffect0 + A).Filename = "Data/Graphics/Effect" & Format(A, "000") & ".rsc"
        DynamicTextures(TexEffect0 + A).Encrypted = True
        'Set TexEffects(A).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Effect" & Format(A, "000") & ".rsc", 256, 256, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, TexEffects(A).TexInfo, ByVal 0)
    Next A
    For A = 1 To NumObjects
        DynamicTextures(TexObject0 + A).Filename = "Data/Graphics/Object" & Format(A, "000") & ".rsc"
        DynamicTextures(TexObject0 + A).Encrypted = True
        'Set TexObjects(A).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Object" & Format(A, "000") & ".rsc", 256, 256, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, TexObjects(A).TexInfo, ByVal 0)
    Next A
    For A = 1 To NumObjectFrames
        DynamicTextures(TexObjectFrames0 + A).Filename = "Data/Graphics/ObjectFrames" & Format(A, "000") & ".rsc"
        DynamicTextures(TexObjectFrames0 + A).Encrypted = True
        'TexLen = DecryptFile("Data/Graphics/Sprite" & Format(A, "000") & ".rsc", Tex)
        'Set texSprites(A).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, Tex(0), TexLen, 512, 512, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texSprites(A).TexInfo, ByVal 0)
    Next A
    For A = 1 To 3
        DynamicTextures(TexLight0 + A).Filename = "Data/Graphics/Light" & Format(A, "000") & ".rsc"
        DynamicTextures(TexLight0 + A).Encrypted = True
        DynamicTextures(TexLight0 + A).Width = 256
        DynamicTextures(TexLight0 + A).Height = 256
        DynamicTextures(TexLight0 + A).Filter = D3DX_FILTER_LINEAR
        'TexLen = DecryptFile("Data/Graphics/Light" & Format(A, "000") & ".rsc", Tex())
        'Set texLights(A).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, Tex(0), TexLen, 256, 256, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texLights(A).TexInfo, ByVal 0)
    Next A
    DynamicTextures(TexControl1).Filename = "Data/Graphics/Interface/Controls001.rsc"
    For A = 1 To NumPortraits
        DynamicTextures(TexPortrait0 + A).Filename = "Data/Graphics/Portraits/" & A & ".rsc"
    Next A
    For A = 1 To NumLSprites
        DynamicTextures(TexLSprite0 + A).Filename = "Data/Graphics/LSprite" & Format(A, "000") & ".rsc"
        DynamicTextures(TexLSprite0 + A).Encrypted = True
        'TexLen = DecryptFile("Data/Graphics/Sprite" & Format(A, "000") & ".rsc", Tex)
        'Set texSprites(A).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, Tex(0), TexLen, 512, 512, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texSprites(A).TexInfo, ByVal 0)
    Next A
    For A = 1 To NumTextures
        DynamicTextures(A).Loaded = False
    Next A
    
    Set texAtts.Texture = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Atts.rsc", 256, 256, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texAtts.TexInfo, ByVal 0)
    texAtts.Loaded = True
    
    Set texParticles(1) = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Particles/Particle001.rsc", 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, &HFF000000, ByVal 0, ByVal 0)
    'texParticles(1).Loaded = True
    Set texParticles(2) = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Particles/Particle002.rsc", 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, &HFF000000, ByVal 0, ByVal 0)
   ' texParticles(2).Loaded = True
    Set texParticles(3) = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Particles/Particle003.rsc", 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, &HFF000000, ByVal 0, ByVal 0)
   ' texParticles(3).Loaded = True
    Set texParticles(4) = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Particles/Particle004.rsc", 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, &HFF000000, ByVal 0, ByVal 0)
   ' texParticles(4).Loaded = True
    Set texParticles(5) = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Particles/Particle005.rsc", 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, &HFF000000, ByVal 0, ByVal 0)
   ' texParticles(5).Loaded = True
    Set texParticles(6) = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Particles/Particle006.rsc", 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, &HFF000000, ByVal 0, ByVal 0)
   ' texParticles(6).Loaded = True
    Set texParticles(7) = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Particles/Particle007.rsc", 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, &HFF000000, ByVal 0, ByVal 0)
   ' texParticles(7).Loaded = True
    Set texParticles(8) = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Particles/Particle008.rsc", 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, &HFF000000, ByVal 0, ByVal 0)
   ' texParticles(8).Loaded = True
    Set texParticles(9) = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Particles/Particle009.rsc", 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, &HFF000000, ByVal 0, ByVal 0)
   ' texParticles(9).Loaded = True
    Set texParticles(10) = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Particles/Particle010.rsc", 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, &HFF000000, ByVal 0, ByVal 0)
   ' texParticles(10).Loaded = True
    
    'Set myTexture = g_d3dx.CreateTextureFromFileEx(g_dev, App.Path & "\particle.bmp", 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, D3DColorARGB(255, 0, 0, 0), ByVal 0, ByVal 0)
    
    
    
    'TexLen = DecryptFile("Data/Graphics/Particles/Particle001.rsc", Tex())
    'Set texParticles(1).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, Tex(0), TexLen, 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, &HFF000000, ByVal 0, ByVal 0)
    ' TexLen = DecryptFile("Data/Graphics/Particles/Particle002.rsc", Tex())
    'Set texParticles(2).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, Tex(0), TexLen, 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, &HFF000000, ByVal 0, ByVal 0)
    '  TexLen = DecryptFile("Data/Graphics/Particles/Particle003.rsc", Tex())
    'Set texParticles(3).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, Tex(0), TexLen, 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, &HFF000000, texParticles(3).TexInfo, ByVal 0)
    '  TexLen = DecryptFile("Data/Graphics/Particles/Particle004.rsc", Tex())
    'Set texParticles(4).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, Tex(0), TexLen, 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, &HFF000000, texParticles(4).TexInfo, ByVal 0)
    '  TexLen = DecryptFile("Data/Graphics/Particles/Particle005.rsc", Tex())
    'Set texParticles(5).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, Tex(0), TexLen, 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, &HFF000000, texParticles(5).TexInfo, ByVal 0)
    '  TexLen = DecryptFile("Data/Graphics/Particles/Particle006.rsc", Tex())
    'Set texParticles(6).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, Tex(0), TexLen, 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, &HFF000000, texParticles(6).TexInfo, ByVal 0)
    '  TexLen = DecryptFile("Data/Graphics/Particles/Particle007.rsc", Tex())
    'Set texParticles(7).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, Tex(0), TexLen, 0, 0, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_BOX, &HFF000000, texParticles(7).TexInfo, ByVal 0)
    '  TexLen = DecryptFile("Data/Graphics/Particles/Particle008.rsc", Tex())
    'Set texParticles(8).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, Tex(0), TexLen, 8, 8, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texParticles(8).TexInfo, ByVal 0)
    '  TexLen = DecryptFile("Data/Graphics/Particles/Particle009.rsc", Tex())
    'Set texParticles(9).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, Tex(0), TexLen, 8, 8, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texParticles(9).TexInfo, ByVal 0)
    '  TexLen = DecryptFile("Data/Graphics/Particles/Particle010.rsc", Tex())
    'Set texParticles(10).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, Tex(0), TexLen, 8, 8, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texParticles(10).TexInfo, ByVal 0)
       
    
    
    
    
    
    
    
    Set texShade.Texture = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Shade.rsc", 32, 32, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texShade.TexInfo, ByVal 0)
    'Set texShade.Texture = D3DX.CreateTexture(D3DDevice, 32, 32, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED)
    texShade.Loaded = True
    texShade.LastUsed = 0
    
    'Set texShadeEX.Texture = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Shade.rsc", 256, 256, 0, D3DUSAGE_RENDERTARGET, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texShadeEX.TexInfo, ByVal 0)
    Set texShadeEX.Texture = D3DX.CreateTexture(D3DDevice, 256, 256, 0, D3DUSAGE_RENDERTARGET, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT)
    texShadeEX.Loaded = True
    texShadeEX.LastUsed = 0
    texShadeEX.TexInfo.Width = 256
    texShadeEX.TexInfo.Height = 256
    
    Set texLightsEX.Texture = D3DX.CreateTexture(D3DDevice, 256, 256, 0, D3DUSAGE_RENDERTARGET, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT)
    texLightsEX.Loaded = True
    texLightsEX.LastUsed = 0
    texLightsEX.TexInfo.Width = 256
    texLightsEX.TexInfo.Height = 256
    
    If Not Options.fullredraws Then
        Set mapTexture(1) = D3DX.CreateTexture(D3DDevice, 512, 512, 0, D3DUSAGE_RENDERTARGET, D3DFMT_X8R8G8B8, D3DPOOL_DEFAULT)
        Set mapTexture(2) = D3DX.CreateTexture(D3DDevice, 512, 512, 0, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
        
        mapChanged = True
        For A = 0 To 11
            For TexLen = 0 To 11
                mapChangedBg(A, TexLen) = True
            Next TexLen
        Next A
    Else
        Set mapTexture(1) = Nothing
        Set mapTexture(2) = Nothing
    End If
    
    
End Sub

Function InitTexture(ByVal TextureNum As Integer) As Boolean
Dim FilePath As String, TexLen As Long

    If TextureNum < 1 Then Exit Function
    If DynamicTextures(TextureNum).LastUsed > GetTickCount Then Exit Function
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Function
    If blnEnd Then Exit Function
    
    DynamicTextures(TextureNum).LastUsed = GetTickCount + TextureWaitTime
    DynamicTextures(TextureNum).Loaded = True
    FilePath = DynamicTextures(TextureNum).Filename
    If Exists(FilePath) Then
        If DynamicTextures(TextureNum).Encrypted Then
            Dim tex() As Byte
            TexLen = DecryptFile(FilePath, tex())
            Set DynamicTextures(TextureNum).Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, tex(0), TexLen, D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, DynamicTextures(TextureNum).Filter, D3DX_FILTER_NONE, &HFF000000, DynamicTextures(TextureNum).TexInfo, ByVal 0)
            
        Else
            Set DynamicTextures(TextureNum).Texture = D3DX.CreateTextureFromFileEx(D3DDevice, FilePath, D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, DynamicTextures(TextureNum).Filter, D3DX_FILTER_NONE, &HFF000000, DynamicTextures(TextureNum).TexInfo, ByVal 0)
        End If
        InitTexture = True
        'PrintChat FilePath, 15, Options.FontSize
    Else
        DynamicTextures(TextureNum).Loaded = False
        D3DDevice.SetTexture 0, Nothing
        LastTexture = 0
        InitTexture = False
    End If
End Function

Public Function ReadyTexture(ByVal TextureNum As Long) As Boolean

    If TextureNum > 0 And TextureNum <= NumTextures Then
        If DynamicTextures(TextureNum).Loaded = False Then
            If InitTexture(TextureNum) = False Then
                ReadyTexture = False
                Exit Function
            End If
        End If

        If LastTexture <> TextureNum Then
            D3DDevice.SetTexture 0, DynamicTextures(TextureNum).Texture
            DynamicTextures(TextureNum).LastUsed = GetTickCount + TextureWaitTime
            LastTexture = TextureNum
            ReadyTexture = True
        Else
            ReadyTexture = True
        End If
    Else
        D3DDevice.SetTexture 0, Nothing
        LastTexture = 0
        ReadyTexture = False
    End If
End Function

Public Sub FixD3DError()
    Dim A As Long
    If CurFog > 0 Then
        Dim TexInfo As D3DXIMAGE_INFO_A
        Set texFog = D3DX.CreateTextureFromFileEx(D3DDevice, "Data/Graphics/Misc/Fog/Fog" & CurFog & ".rsc", 512, 256, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_NONE, &HFF000000, TexInfo, ByVal 0)
        D3DDevice.SetTexture 0, texFog
    End If
    CMap2 = CMap * CMap + 5: CX2 = cX ^ 2 + 5: CY2 = cY ^ 2 + 5
    For A = 1 To NumTextures
        Set DynamicTextures(A).Texture = Nothing
        DynamicTextures(A).Loaded = False
        DynamicTextures(A).LastUsed = 0
    Next A
    BOOLD3DERROR = False
End Sub

Sub DrawTile(ByVal x As Long, ByVal y As Long, ByVal Tile As Long, Optional XOffset As Long = 0, Optional YOffset As Long = 0, Optional Height As Long = 32, Optional Width As Long = 32)
Dim VertexArray(0 To 3) As TLVERTEX
    Dim tColor As Long
    tColor = D3DColorARGB(255, 255, 255, 255)
    
    'Static LastTex As Long
    Dim CurTex As Long
    CurTex = (((Tile - 1) \ 64) + 1)
    'If LastTex <> CurTex Then
        'D3DDevice.SetTexture 0, texTiles(CurTex).Texture
        ReadyTexture TexTile0 + CurTex
    '    LastTex = CurTex
    'End If
    'Set the RHWs (must always be 1)
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1

    'Apply the colors
    VertexArray(0).Color = tColor
    VertexArray(1).Color = tColor
    VertexArray(2).Color = tColor
    VertexArray(3).Color = tColor

    'Find the left side of the rectangle
    VertexArray(0).x = x + XOffset - 0.5
    VertexArray(0).tu = ((((Tile - 1) Mod 8) * 32 + XOffset) / 256)

    'Find the top side of the rectangle
    VertexArray(0).y = y + YOffset - 0.5
    VertexArray(0).tv = ((((Tile - 1) \ 8) * 32 + YOffset) / 256)

    'Find the right side of the rectangle
    VertexArray(1).x = x + Width + XOffset - 0.5
    VertexArray(1).tu = (((Tile - 1) Mod 8) * 32 + Width + XOffset) / 256

    'These values will only equal each other when not a shadow
    VertexArray(2).x = VertexArray(0).x
    VertexArray(3).x = VertexArray(1).x
    
    'Find the bottom of the rectangle
    VertexArray(2).y = y + Height + YOffset - 0.5
    VertexArray(2).tv = (((Tile - 1) \ 8) * 32 + Height + YOffset) / 256

    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(1).y = VertexArray(0).y
    VertexArray(1).tv = VertexArray(0).tv
    VertexArray(2).tu = VertexArray(0).tu
    VertexArray(3).y = VertexArray(2).y
    VertexArray(3).tu = VertexArray(1).tu
    VertexArray(3).tv = VertexArray(2).tv

    'Render the texture to the device
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size

End Sub
Sub DrawAtts(ByVal x As Long, ByVal y As Long, ByVal Att As Long, Optional Wall As Long = 0)
Dim VertexArray(0 To 3) As TLVERTEX
    Dim tColor As Long
    tColor = D3DColorARGB(255, 255, 255, 255)
    
    LastTexture = 0
    D3DDevice.SetTexture 0, texAtts.Texture
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
    VertexArray(0).Color = tColor
    VertexArray(1).Color = tColor
    VertexArray(2).Color = tColor
    VertexArray(3).Color = tColor

    'Find the left side of the rectangle
    VertexArray(0).x = x - 0.5
    VertexArray(0).tu = ((((Att - 1) Mod 7) * 32) / 256)

    'Find the top side of the rectangle
    VertexArray(0).y = y - 0.5
    VertexArray(0).tv = ((((Att - 1) \ 7) * 32 + (Wall * 160)) / 256)

    'Find the right side of the rectangle
    VertexArray(1).x = x + 32 - 0.5
    VertexArray(1).tu = (((Att - 1) Mod 7) * 32 + 32) / 256

    VertexArray(2).x = VertexArray(0).x
    VertexArray(3).x = VertexArray(1).x
    
    'Find the bottom of the rectangle
    VertexArray(2).y = y + 32 - 0.5
    VertexArray(2).tv = (((Att - 1) \ 7) * 32 + 32 + (Wall * 160)) / 256

    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(1).y = VertexArray(0).y
    VertexArray(1).tv = VertexArray(0).tv
    VertexArray(2).tu = VertexArray(0).tu
    VertexArray(3).y = VertexArray(2).y
    VertexArray(3).tu = VertexArray(1).tu
    VertexArray(3).tv = VertexArray(2).tv
    'Render the texture to the device
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size

End Sub
Sub DrawTile2(ByVal x As Long, ByVal y As Long, ByVal Tile As Long, ByVal tColor As Long, Optional XOffset As Long = 0, Optional YOffset As Long = 0, Optional Height As Long = 32, Optional Width As Long = 32)
Dim VertexArray(0 To 3) As TLVERTEX
    
    'Static LastTex As Long
    Dim CurTex As Long
    CurTex = (((Tile - 1) \ 64) + 1)
    'If LastTex <> CurTex Then
        'D3DDevice.SetTexture 0, texTiles(CurTex).Texture
        ReadyTexture TexTile0 + CurTex
    '    LastTex = CurTex
    'End If
    'Set the RHWs (must always be 1)
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1

    'Apply the colors
    VertexArray(0).Color = tColor
    VertexArray(1).Color = tColor
    VertexArray(2).Color = tColor
    VertexArray(3).Color = tColor

    'Find the left side of the rectangle
    VertexArray(0).x = x + XOffset - 0.5
    VertexArray(0).tu = ((((Tile - 1) Mod 8) * 32 + XOffset) / 256)

    'Find the top side of the rectangle
    VertexArray(0).y = y + YOffset - 0.5
    VertexArray(0).tv = ((((Tile - 1) \ 8) * 32 + YOffset) / 256)

    'Find the right side of the rectangle
    VertexArray(1).x = x + Width + XOffset - 0.5
    VertexArray(1).tu = (((Tile - 1) Mod 8) * 32 + Width + XOffset) / 256

    'These values will only equal each other when not a shadow
    VertexArray(2).x = VertexArray(0).x
    VertexArray(3).x = VertexArray(1).x
    
    'Find the bottom of the rectangle
    VertexArray(2).y = y + Height + YOffset - 0.5
    VertexArray(2).tv = (((Tile - 1) \ 8) * 32 + Height + YOffset) / 256

    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(1).y = VertexArray(0).y
    VertexArray(1).tv = VertexArray(0).tv
    VertexArray(2).tu = VertexArray(0).tu
    VertexArray(3).y = VertexArray(2).y
    VertexArray(3).tu = VertexArray(1).tu
    VertexArray(3).tv = VertexArray(2).tv


    VertexArray(0).x = (VertexArray(0).x / 384) * 256
    VertexArray(0).y = (VertexArray(0).y / 384) * 256
    VertexArray(1).x = (VertexArray(1).x / 384) * 256
    VertexArray(1).y = (VertexArray(1).y / 384) * 256
    VertexArray(2).x = (VertexArray(2).x / 384) * 256
    VertexArray(2).y = (VertexArray(2).y / 384) * 256
    VertexArray(3).x = (VertexArray(3).x / 384) * 256
    VertexArray(3).y = (VertexArray(3).y / 384) * 256
    
    
    
    'Render the texture to the device
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size

End Sub

Sub DrawObject(ByVal x As Long, ByVal y As Long, ByVal Object As Long, Optional ByVal alpha As Long = 255)
    
    Dim VertexArray(0 To 3) As TLVERTEX
    Dim tColor As Long
    tColor = D3DColorARGB(alpha, 255, 255, 255)
    Dim CurTex As Long
    CurTex = (((Object - 1) \ 64) + 1)
    If CurTex > 0 And CurTex <= NumObjects Then
        ReadyTexture TexObject0 + CurTex
    
        VertexArray(0).rhw = 1
        VertexArray(1).rhw = 1
        VertexArray(2).rhw = 1
        VertexArray(3).rhw = 1
    
        'Apply the colors
        VertexArray(0).Color = tColor
        VertexArray(1).Color = tColor
        VertexArray(2).Color = tColor
        VertexArray(3).Color = tColor
    
        'Find the left side of the rectangle
        VertexArray(0).x = x - 0.5
        VertexArray(0).tu = (((Object - 1) Mod 8) * 32) / 256
    
        'Find the top side of the rectangle
        VertexArray(0).y = y - 0.5
        VertexArray(0).tv = (((Object - 1) \ 8) * 32) / 256
    
        'Find the right side of the rectangle
        VertexArray(1).x = x + 32 - 0.5
        VertexArray(1).tu = (((Object - 1) Mod 8) * 32 + 32) / 256
    
        VertexArray(2).x = VertexArray(0).x
        VertexArray(3).x = VertexArray(1).x
        VertexArray(2).y = y + 32 - 0.5
        VertexArray(2).tv = (((Object - 1) \ 8) * 32 + 32) / 256
        VertexArray(1).y = VertexArray(0).y
        VertexArray(1).tv = VertexArray(0).tv
        VertexArray(2).tu = VertexArray(0).tu
        VertexArray(3).y = VertexArray(2).y
        VertexArray(3).tu = VertexArray(1).tu
        VertexArray(3).tv = VertexArray(2).tv

        'Render the texture to the device
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
    End If
End Sub

Sub DrawObjectGlow(ByVal x As Long, ByVal y As Long, ByVal Object As Long, ByVal gColor As Long)
    Dim VertexArray(0 To 3) As TLVERTEX
    Dim x1 As Long, y1 As Long

    Dim CurTex As Long
    CurTex = (((Object - 1) \ 64) + 1)
    If CurTex > 0 And CurTex <= NumObjects Then
        ReadyTexture TexObject0 + CurTex

        VertexArray(0).rhw = 1
        VertexArray(1).rhw = 1
        VertexArray(2).rhw = 1
        VertexArray(3).rhw = 1
        
        'Apply the colors
        VertexArray(0).Color = gColor
        VertexArray(1).Color = gColor
        VertexArray(2).Color = gColor
        VertexArray(3).Color = gColor
    
        For x1 = -1 To 1
            For y1 = -1 To 1
                If Not ((x1 = 0) Or (y1 = 0)) Then
                    'Find the left side of the rectangle
                    VertexArray(0).x = x - 0.5 + x1
                    VertexArray(0).tu = (((Object - 1) Mod 8) * 32) / 256
                 
                    'Find the top side of the rectangle
                    VertexArray(0).y = y - 0.5 + y1
                    VertexArray(0).tv = (((Object - 1) \ 8) * 32) / 256
                
                    'Find the right side of the rectangle
                    VertexArray(1).x = x + 31.5 + x1
                    VertexArray(1).tu = (((Object - 1) Mod 8) * 32 + 32) / 256
                    VertexArray(2).x = VertexArray(0).x
                    VertexArray(3).x = VertexArray(1).x
                    VertexArray(2).y = y + 31.5 + y1
                    VertexArray(2).tv = (((Object - 1) \ 8) * 32 + 32) / 256
                
                    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
                    VertexArray(1).y = VertexArray(0).y
                    VertexArray(1).tv = VertexArray(0).tv
                    VertexArray(2).tu = VertexArray(0).tu
                    VertexArray(3).y = VertexArray(2).y
                    VertexArray(3).tu = VertexArray(1).tu
                    VertexArray(3).tv = VertexArray(2).tv
                
                    'Render the texture to the device
                    D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_DIFFUSE
                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
                    D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
                End If
            Next y1
        Next x1
    End If
End Sub

Sub DrawProjectile(ByVal x As Long, ByVal y As Long, ByVal Projectile As Long, Direction As Long, ByVal Frame As Long)
    Dim VertexArray(0 To 3) As TLVERTEX
    Dim tColor As Long
    tColor = D3DColorARGB(255, 255, 255, 255)
        'D3DDevice.SetTexture 0, texProjectiles(Projectile).Texture
        ReadyTexture TexProjectile0 + Projectile
    '    LastTex = CurTex
    'End If
    'Set the RHWs (must always be 1)
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
    
    'Apply the colors
    VertexArray(0).Color = tColor
    VertexArray(1).Color = tColor
    VertexArray(2).Color = tColor
    VertexArray(3).Color = tColor

    'Find the left side of the rectangle
    VertexArray(0).x = x - 0.5
    VertexArray(0).tu = (Frame * 32 / 128)

    'Find the top side of the rectangle
    VertexArray(0).y = y - 0.5
    VertexArray(0).tv = (Direction * 32 / 128)

    'Find the right side of the rectangle
    VertexArray(1).x = x + 32 - 0.5
    VertexArray(1).tu = (Frame * 32 + 32) / 128
    VertexArray(2).x = VertexArray(0).x
    VertexArray(3).x = VertexArray(1).x
    
    'Find the bottom of the rectangle
    VertexArray(2).y = y + 32 - 0.5
    VertexArray(2).tv = (Direction * 32 + 32) / 128

    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(1).y = VertexArray(0).y
    VertexArray(1).tv = VertexArray(0).tv
    VertexArray(2).tu = VertexArray(0).tu
    VertexArray(3).y = VertexArray(2).y
    VertexArray(3).tu = VertexArray(1).tu
    VertexArray(3).tv = VertexArray(2).tv
    
    'Render the texture to the device
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size

End Sub
Sub DrawEffect(ByVal x As Long, ByVal y As Long, ByVal Effect As Long, ByVal Frame As Long)
    
    Dim VertexArray(0 To 3) As TLVERTEX
    Dim tColor As Long
    tColor = D3DColorARGB(255, 255, 255, 255)
    
    'Static LastTex As Long
    Dim CurTex As Long
    CurTex = ((Effect - 1) \ 8) + 1
    'If LastTex <> CurTex Then
        ReadyTexture TexEffect0 + CurTex
        'D3DDevice.SetTexture 0, TexEffects(CurTex).Texture
    '    LastTex = CurTex
    'End If
    'Set the RHWs (must always be 1)
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
    
    'Apply the colors
    VertexArray(0).Color = tColor
    VertexArray(1).Color = tColor
    VertexArray(2).Color = tColor
    VertexArray(3).Color = tColor

    'Find the left side of the rectangle
    VertexArray(0).x = x - 0.5
    VertexArray(0).tu = (Frame * 32 / 256)

    'Find the top side of the rectangle
    VertexArray(0).y = y - 0.5
    VertexArray(0).tv = ((Effect - 1) * 32 / 256)

    'Find the right side of the rectangle
    VertexArray(1).x = x + 32 - 0.5
    VertexArray(1).tu = (Frame * 32 + 32) / 256
    VertexArray(2).x = VertexArray(0).x
    VertexArray(3).x = VertexArray(1).x
    
    'Find the bottom of the rectangle
    VertexArray(2).y = y + 32 - 0.5
    VertexArray(2).tv = ((Effect - 1) * 32 + 32) / 256

    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(1).y = VertexArray(0).y
    VertexArray(1).tv = VertexArray(0).tv
    VertexArray(2).tu = VertexArray(0).tu
    VertexArray(3).y = VertexArray(2).y
    VertexArray(3).tu = VertexArray(1).tu
    VertexArray(3).tv = VertexArray(2).tv

    'Render the texture to the device
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size

End Sub

Sub DrawShadow(ByVal x As Long, ByVal y As Long, ByVal Sprite As Long, ByVal Frame As Long, VCO As Long)
Dim VertexArray(0 To 3) As TLVERTEX
    Dim tColor As Long, Fade As Single, xOff As Single, yOff As Single
    
    xOff = 15.5
    yOff = 7.5
    
    
    'Fade = (AmbientAlpha) / 200
    Fade = 1
    'If Fade > 1 Then Fade = 1
    
    Select Case World.Hour
    
    Case 7 'day
        xOff = -18
    Case 8
        xOff = -16
    Case 9
        xOff = -14
    Case 10
        xOff = -12
    Case 11
        xOff = -10
    Case 12
        xOff = -8
    Case 13
        xOff = -6
    Case 14
        xOff = -4
    Case 15
        xOff = 4
    Case 16
        xOff = 6
    Case 17
        xOff = 8
    Case 18
        xOff = 10
    Case 19
        xOff = 12
    Case 20
        xOff = 14
    Case 21
        xOff = 16
    Case 22
        xOff = 18
    Case 23
        xOff = 22
        Fade = Fade - 0.2
    Case 24
        xOff = 22
        Fade = Fade - 0.4
    Case 1
        xOff = 15
        Fade = Fade - 0.6
    Case 2
        xOff = 8
        Fade = Fade - 0.8
    Case 3
        xOff = -8
        Fade = Fade - 0.8
    Case 4
        xOff = -15
        Fade = Fade - 0.6
    Case 5
        xOff = -22
        Fade = Fade - 0.4
    Case 6
        xOff = -22
        Fade = Fade - 0.2
        

    If Fade < 0 Then Fade = 0
    End Select
    
    
    tColor = D3DColorARGB(127 * Fade, 0, 0, 0)
    'Static LastTex As Long
    Dim CurTex As Long
    CurTex = (((Sprite - 1) \ 16) + 1)
    'If LastTex <> CurTex Then
        'D3DDevice.SetTexture 0, texSprites(CurTex).Texture
        ReadyTexture TexSprite0 + CurTex
    '    LastTex = CurTex
    'End If
    'Set the RHWs (must always be 1)
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
    
    'Apply the colors
    VertexArray(0).Color = tColor
    VertexArray(1).Color = tColor
    VertexArray(2).Color = tColor
    VertexArray(3).Color = tColor

    'Find the left side of the rectangle
    VertexArray(0).x = x + xOff
    VertexArray(0).tu = (Frame * 32 / 512)

    'Find the top side of the rectangle
    VertexArray(0).y = y + 7.5
    VertexArray(0).tv = ((Sprite - 1) * 32 / 512)

    'Find the right side of the rectangle

    VertexArray(1).x = x + xOff + 32
    
    
    VertexArray(1).tu = (Frame * 32 + 32) / 512
    VertexArray(2).x = x - 0.5
    VertexArray(3).x = x + 31.5
    
    'Find the bottom of the rectangle
    VertexArray(2).y = y + 7.5 + 24 - VCO
    VertexArray(2).tv = ((Sprite - 1) * 32 + 32 - VCO) / 512

    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(1).y = VertexArray(0).y
    VertexArray(1).tv = VertexArray(0).tv
    VertexArray(2).tu = VertexArray(0).tu
    VertexArray(3).y = VertexArray(2).y
    VertexArray(3).tu = VertexArray(1).tu
    VertexArray(3).tv = VertexArray(2).tv

    'Render the texture to the device
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size

End Sub
Sub DrawLargeShadow(ByVal x As Long, ByVal y As Long, ByVal Sprite As Long, ByVal Frame As Long, VCO As Long)
Dim VertexArray(0 To 3) As TLVERTEX
    Dim tColor As Long, Fade As Single, xOff As Single, yOff As Single
    
    xOff = 15.5
    yOff = 7.5
    

    Fade = 1
    'Fade = (AmbientAlpha) / 200
    'If Fade > 1 Then Fade = 1
    
    Select Case World.Hour
    
    Case 7 'day
        xOff = -18
    Case 8
        xOff = -16
    Case 9
        xOff = -14
    Case 10
        xOff = -12
    Case 11
        xOff = -10
    Case 12
        xOff = -8
    Case 13
        xOff = -6
    Case 14
        xOff = -4
    Case 15
        xOff = 4
    Case 16
        xOff = 6
    Case 17
        xOff = 8
    Case 18
        xOff = 10
    Case 19
        xOff = 12
    Case 20
        xOff = 14
    Case 21
        xOff = 16
    Case 22
        xOff = 18
    Case 23
        xOff = 22
        Fade = Fade - 0.2
    Case 24
        xOff = 22
        Fade = Fade - 0.4
    Case 1
        xOff = 15
        Fade = Fade - 0.6
    Case 2
        xOff = 8
        Fade = Fade - 0.8
    Case 3
        xOff = -8
        Fade = Fade - 0.8
    Case 4
        xOff = -15
        Fade = Fade - 0.6
    Case 5
        xOff = -22
        Fade = Fade - 0.4
    Case 6
        xOff = -22
        Fade = Fade - 0.2
        

    If Fade < 0 Then Fade = 0
    End Select
    
    
    tColor = D3DColorARGB(127 * Fade, 0, 0, 0)
    'Static LastTex As Long
    Dim CurTex As Long

    CurTex = (((Sprite - 1) \ 16) + 1)

    'If LastTex <> CurTex Then
        'D3DDevice.SetTexture 0, texSprites(CurTex).Texture
        ReadyTexture TexLSprite0 + CurTex
    '    LastTex = CurTex
    'End If
    'Set the RHWs (must always be 1)
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
    
    'Apply the colors
    VertexArray(0).Color = tColor
    VertexArray(1).Color = tColor
    VertexArray(2).Color = tColor
    VertexArray(3).Color = tColor

    'Find the left side of the rectangle
    VertexArray(0).x = x + xOff
    VertexArray(0).tu = (Frame * 64 / 512)

    'Find the top side of the rectangle
    VertexArray(0).y = y + 7.5
    VertexArray(0).tv = ((Sprite - 1) * 32 / 512)

    'Find the right side of the rectangle

    VertexArray(1).x = x + xOff + 64
    
    
    VertexArray(1).tu = (Frame * 64 + 64) / 512
    VertexArray(2).x = x - 0.5
    VertexArray(3).x = x + 63.5
    
    'Find the bottom of the rectangle
    VertexArray(2).y = y + 7.5 + 24 + 32 - VCO
    VertexArray(2).tv = ((Sprite - 1) * 32 + 64 - VCO) / 512

    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(1).y = VertexArray(0).y
    VertexArray(1).tv = VertexArray(0).tv
    VertexArray(2).tu = VertexArray(0).tu
    VertexArray(3).y = VertexArray(2).y
    VertexArray(3).tu = VertexArray(1).tu
    VertexArray(3).tv = VertexArray(2).tv

    'Render the texture to the device
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size

End Sub
Sub DrawSpriteGlow(ByVal x As Long, ByVal y As Long, ByVal Sprite As Long, ByVal Frame As Long, ByVal gColor As Long, Optional VCO As Long = 0)
Dim VertexArray(0 To 3) As TLVERTEX
    Dim x1 As Long, y1 As Long
    
    'Static LastTex As Long
    Dim CurTex As Long
    CurTex = (((Sprite - 1) \ 16) + 1)
    
    'If LastTex <> CurTex Then
        'D3DDevice.SetTexture 0, texSprites(CurTex).Texture
        ReadyTexture TexSprite0 + CurTex
    '    LastTex = CurTex
    'End If
    'Set the RHWs (must always be 1)
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
    
    'Apply the colors
    VertexArray(0).Color = gColor
    VertexArray(1).Color = gColor
    VertexArray(2).Color = gColor
    VertexArray(3).Color = gColor

    For x1 = -1 To 1
        For y1 = -1 To 1
        If Not ((x1 = 0) Or (y1 = 0)) Then
        'Find the left side of the rectangle
        VertexArray(0).x = x - 0.5 + x1
        VertexArray(0).tu = (Frame * 32 / 512)
    
        'Find the top side of the rectangle
        VertexArray(0).y = y - 0.5 + y1
        VertexArray(0).tv = ((Sprite - 1) * 32 / 512)
    
        'Find the right side of the rectangle
        VertexArray(1).x = x + 31.5 + x1
        'VertexArray(1).tu = (Frame * 32 + 32) / 512
      
        VertexArray(1).tu = (Frame * 32 + 32) / 512
        
        VertexArray(2).x = VertexArray(0).x
        VertexArray(3).x = VertexArray(1).x
        
        'Find the bottom of the rectangle
        VertexArray(2).y = y + 31.5 + y1 - VCO
        VertexArray(2).tv = ((Sprite - 1) * 32 + 32 - VCO) / 512
    
        'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
        VertexArray(1).y = VertexArray(0).y
        VertexArray(1).tv = VertexArray(0).tv
        VertexArray(2).tu = VertexArray(0).tu
        VertexArray(3).y = VertexArray(2).y
        VertexArray(3).tu = VertexArray(1).tu
        VertexArray(3).tv = VertexArray(2).tv
        
        'Render the texture to the device
        D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_DIFFUSE
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
        D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
        End If
        Next y1
    Next x1
    
End Sub
Sub DrawLargeSpriteGlow(ByVal x As Long, ByVal y As Long, ByVal Sprite As Long, ByVal Frame As Long, ByVal gColor As Long, Optional VCO As Long = 0)
Dim VertexArray(0 To 3) As TLVERTEX
    Dim x1 As Long, y1 As Long
    
    Sprite = (Sprite * 4) - 3
    
    'Static LastTex As Long
    Dim CurTex As Long
              If Frame >= 8 Then
            Sprite = Sprite + 2
            Frame = Frame - 8
        End If
    CurTex = (((Sprite - 1) \ 16) + 1)

    
    'If LastTex <> CurTex Then
        'D3DDevice.SetTexture 0, texSprites(CurTex).Texture
        ReadyTexture TexLSprite0 + CurTex
    '    LastTex = CurTex
    'End If
    'Set the RHWs (must always be 1)
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
    
    'Apply the colors
    VertexArray(0).Color = gColor
    VertexArray(1).Color = gColor
    VertexArray(2).Color = gColor
    VertexArray(3).Color = gColor

    For x1 = -1 To 1
        For y1 = -1 To 1
        If Not ((x1 = 0) Or (y1 = 0)) Then
        'Find the left side of the rectangle
        VertexArray(0).x = x - 0.5 + x1
        VertexArray(0).tu = (Frame * 64 / 512)
    
        'Find the top side of the rectangle
        VertexArray(0).y = y - 0.5 + y1
        VertexArray(0).tv = ((Sprite - 1) * 32 / 512)
    
        'Find the right side of the rectangle
        VertexArray(1).x = x + 63.5 + x1
        VertexArray(1).tu = (Frame * 64 + 64) / 512
        VertexArray(2).x = VertexArray(0).x
        VertexArray(3).x = VertexArray(1).x
        
        'Find the bottom of the rectangle
        VertexArray(2).y = y + 63.5 + y1 - VCO
        VertexArray(2).tv = ((Sprite - 1) * 32 + 64 - VCO) / 512
    
        'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
        VertexArray(1).y = VertexArray(0).y
        VertexArray(1).tv = VertexArray(0).tv
        VertexArray(2).tu = VertexArray(0).tu
        VertexArray(3).y = VertexArray(2).y
        VertexArray(3).tu = VertexArray(1).tu
        VertexArray(3).tv = VertexArray(2).tv
        
        'Render the texture to the device
        D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_DIFFUSE
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
        D3DDevice.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
        End If
        Next y1
    Next x1
    
End Sub
Sub DrawSprite(ByVal x As Long, ByVal y As Long, ByVal Sprite As Long, ByVal Frame As Long, ByVal Color As Long, Optional Shadow As Boolean = False, Optional VCO As Long = 0)
Dim VertexArray(0 To 3) As TLVERTEX

    Dim CurTex As Long
    CurTex = (((Sprite - 1) \ 16) + 1)
    If CurTex >= 1 And CurTex <= 16 Then
        If Shadow Then DrawShadow x, y, Sprite, Frame, VCO
        ReadyTexture TexSprite0 + CurTex
        VertexArray(0).rhw = 1
        VertexArray(1).rhw = 1
        VertexArray(2).rhw = 1
        VertexArray(3).rhw = 1
        VertexArray(0).Color = Color
        VertexArray(1).Color = Color
        VertexArray(2).Color = Color
        VertexArray(3).Color = Color
        VertexArray(0).x = x - 0.5
        VertexArray(0).tu = (Frame * 32 / 512)
        VertexArray(0).y = y - 0.5
        VertexArray(0).tv = ((Sprite - 1) * 32 / 512)
        VertexArray(1).x = x + 31.5
        VertexArray(1).tu = (Frame * 32 + 32) / 512
        VertexArray(2).x = VertexArray(0).x
        VertexArray(3).x = VertexArray(1).x
        VertexArray(2).y = y + 31.5 - VCO
        VertexArray(2).tv = ((Sprite - 1) * 32 + 32 - VCO) / 512
        VertexArray(1).y = VertexArray(0).y
        VertexArray(1).tv = VertexArray(0).tv
        VertexArray(2).tu = VertexArray(0).tu
        VertexArray(3).y = VertexArray(2).y
        VertexArray(3).tu = VertexArray(1).tu
        VertexArray(3).tv = VertexArray(2).tv
    
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
        
    End If
End Sub
Sub DrawSymbol(ByVal x As Long, ByVal y As Long, ByVal GuildNum As Long, Optional VCO As Long = 0)
Dim VertexArray(0 To 3) As TLVERTEX

Dim symbol As Long



    Dim CurTex As Long

symbol = Guild(GuildNum).Symbol3
If symbol > 0 Then
    CurTex = TexSymbol0 + 3
        ReadyTexture CurTex
        
        VertexArray(0).rhw = 1
        VertexArray(1).rhw = 1
        VertexArray(2).rhw = 1
        VertexArray(3).rhw = 1
        VertexArray(0).Color = D3DColorARGB(255, 255, 255, 255)
        VertexArray(1).Color = VertexArray(0).Color
        VertexArray(2).Color = VertexArray(0).Color
        VertexArray(3).Color = VertexArray(0).Color
        VertexArray(0).x = x - 0.5
        VertexArray(0).tu = (((symbol Mod 16) * 20) - 20) / 320
        VertexArray(0).y = y - 0.5
        VertexArray(0).tv = (Int(symbol / 16) * 20) / 320
        VertexArray(1).x = x + 19.5
        VertexArray(1).tu = ((symbol Mod 16) * 20) / 320
        VertexArray(2).x = VertexArray(0).x
        VertexArray(3).x = VertexArray(1).x
        VertexArray(2).y = y + 19.5 - VCO
        VertexArray(2).tv = (Int(symbol / 16) * 20 + 19 - VCO) / 320
        VertexArray(1).y = VertexArray(0).y
        VertexArray(1).tv = VertexArray(0).tv
        VertexArray(2).tu = VertexArray(0).tu
        VertexArray(3).y = VertexArray(2).y
        VertexArray(3).tu = VertexArray(1).tu
        VertexArray(3).tv = VertexArray(2).tv
    
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
End If


symbol = Guild(GuildNum).Symbol2
If symbol > 0 Then
    CurTex = TexSymbol0 + 2
        ReadyTexture CurTex
        
        VertexArray(0).rhw = 1
        VertexArray(1).rhw = 1
        VertexArray(2).rhw = 1
        VertexArray(3).rhw = 1
        VertexArray(0).Color = D3DColorARGB(240, 255, 255, 255)
        VertexArray(1).Color = VertexArray(0).Color
        VertexArray(2).Color = VertexArray(0).Color
        VertexArray(3).Color = VertexArray(0).Color
        VertexArray(0).x = x - 0.5
        VertexArray(0).tu = (((symbol Mod 16) * 20) - 20) / 320
        VertexArray(0).y = y - 0.5
        VertexArray(0).tv = (Int(symbol / 16) * 20) / 320
        VertexArray(1).x = x + 19.5
        VertexArray(1).tu = ((symbol Mod 16) * 20) / 320
        VertexArray(2).x = VertexArray(0).x
        VertexArray(3).x = VertexArray(1).x
        VertexArray(2).y = y + 19.5 - VCO
        VertexArray(2).tv = (Int(symbol / 16) * 20 + 19 - VCO) / 320
        VertexArray(1).y = VertexArray(0).y
        VertexArray(1).tv = VertexArray(0).tv
        VertexArray(2).tu = VertexArray(0).tu
        VertexArray(3).y = VertexArray(2).y
        VertexArray(3).tu = VertexArray(1).tu
        VertexArray(3).tv = VertexArray(2).tv
    
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
End If


symbol = Guild(GuildNum).Symbol1
If symbol > 0 Then
    CurTex = TexSymbol0 + 1
        ReadyTexture CurTex
        
        VertexArray(0).rhw = 1
        VertexArray(1).rhw = 1
        VertexArray(2).rhw = 1
        VertexArray(3).rhw = 1
        VertexArray(0).Color = D3DColorARGB(240, 255, 255, 255)
        VertexArray(1).Color = VertexArray(0).Color
        VertexArray(2).Color = VertexArray(0).Color
        VertexArray(3).Color = VertexArray(0).Color
        VertexArray(0).x = x - 0.5
        VertexArray(0).tu = (((symbol Mod 16) * 20) - 20) / 320
        VertexArray(0).y = y - 0.5
        VertexArray(0).tv = (Int(symbol / 16) * 20) / 320
        VertexArray(1).x = x + 19.5
        VertexArray(1).tu = ((symbol Mod 16) * 20) / 320
        VertexArray(2).x = VertexArray(0).x
        VertexArray(3).x = VertexArray(1).x
        VertexArray(2).y = y + 19.5 - VCO
        VertexArray(2).tv = (Int(symbol / 16) * 20 + 19 - VCO) / 320
        VertexArray(1).y = VertexArray(0).y
        VertexArray(1).tv = VertexArray(0).tv
        VertexArray(2).tu = VertexArray(0).tu
        VertexArray(3).y = VertexArray(2).y
        VertexArray(3).tu = VertexArray(1).tu
        VertexArray(3).tv = VertexArray(2).tv
    
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
End If

End Sub
Sub DrawEquipmentShadowAlt(ByVal x As Long, ByVal y As Long, ByVal Sprite As Long, ByVal Frame As Long, VCO As Long, eqType As Byte)
Dim VertexArray(0 To 3) As TLVERTEX
    Dim tColor As Long, Fade As Single, xOff As Single, yOff As Single
    
    
        If Frame >= 6 And Frame <= 8 Then
        Frame = Frame + 3
    ElseIf Frame >= 9 And Frame <= 11 Then Frame = Frame - 3
End If
    
    If eqType = 2 Then
        x = x + frameOffsetsX2(Frame)
        y = y + frameOffsetsY2(Frame)
    End If

   

    xOff = 15.5
    yOff = 7.5
    
    
    'Fade = (AmbientAlpha) / 200
    Fade = 1
    'If Fade > 1 Then Fade = 1
    
    Select Case World.Hour
    
    Case 7 'day
        xOff = -18
    Case 8
        xOff = -16
    Case 9
        xOff = -14
    Case 10
        xOff = -12
    Case 11
        xOff = -10
    Case 12
        xOff = -8
    Case 13
        xOff = -6
    Case 14
        xOff = -4
    Case 15
        xOff = 4
    Case 16
        xOff = 6
    Case 17
        xOff = 8
    Case 18
        xOff = 10
    Case 19
        xOff = 12
    Case 20
        xOff = 14
    Case 21
        xOff = 16
    Case 22
        xOff = 18
    Case 23
        xOff = 22
        Fade = Fade - 0.2
    Case 24
        xOff = 22
        Fade = Fade - 0.4
    Case 1
        xOff = 15
        Fade = Fade - 0.6
    Case 2
        xOff = 8
        Fade = Fade - 0.8
    Case 3
        xOff = -8
        Fade = Fade - 0.8
    Case 4
        xOff = -15
        Fade = Fade - 0.6
    Case 5
        xOff = -22
        Fade = Fade - 0.4
    Case 6
        xOff = -22
        Fade = Fade - 0.2
        

    If Fade < 0 Then Fade = 0
    End Select
    
    
    tColor = D3DColorARGB(127 * Fade, 0, 0, 0)
    'Static LastTex As Long
    Dim CurTex As Long
    CurTex = (((Sprite - 1) \ 16) + 1)
    'If LastTex <> CurTex Then
        'D3DDevice.SetTexture 0, texSprites(CurTex).Texture
        ReadyTexture TexObjectFrames0 + CurTex
    '    LastTex = CurTex
    'End If
    'Set the RHWs (must always be 1)
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
    
    'Apply the colors
    VertexArray(0).Color = tColor
    VertexArray(1).Color = tColor
    VertexArray(2).Color = tColor
    VertexArray(3).Color = tColor

    'Find the left side of the rectangle
    VertexArray(0).x = x - xOff
    VertexArray(0).tu = (Frame * 32 / 512)

    'Find the top side of the rectangle
    VertexArray(0).y = y + 7.5
    VertexArray(0).tv = ((Sprite - 1) * 32 / 512)

    'Find the right side of the rectangle

    VertexArray(1).x = x - xOff - 32
    
    
    VertexArray(1).tu = (Frame * 32 + 32) / 512
    VertexArray(2).x = x + 0.5
    VertexArray(3).x = x - 31.5
    
    'Find the bottom of the rectangle
    VertexArray(2).y = y + 7.5 + 24 - VCO
    VertexArray(2).tv = ((Sprite - 1) * 32 + 32 - VCO) / 512

    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(1).y = VertexArray(0).y
    VertexArray(1).tv = VertexArray(0).tv
    VertexArray(2).tu = VertexArray(0).tu
    VertexArray(3).y = VertexArray(2).y
    VertexArray(3).tu = VertexArray(1).tu
    VertexArray(3).tv = VertexArray(2).tv

    'Render the texture to the device
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size

End Sub

Sub DrawEquipmentAlt(ByVal x As Long, ByVal y As Long, ByVal Sprite As Long, ByVal Frame As Long, ByVal Color As Long, VCO As Long, eqType As Byte)
Dim VertexArray(0 To 3) As TLVERTEX


    
    If Frame >= 6 And Frame <= 8 Then
        Frame = Frame + 3
    ElseIf Frame >= 9 And Frame <= 11 Then Frame = Frame - 3
End If

If eqType = 2 Then
    x = x + frameOffsetsX2(Frame)
    y = y + frameOffsetsY2(Frame)
End If


    Dim CurTex As Long
    CurTex = (((Sprite - 1) \ 16) + 1)
    If CurTex >= 1 And CurTex <= 16 Then
        ReadyTexture TexObjectFrames0 + CurTex
        VertexArray(0).rhw = 1
        VertexArray(1).rhw = 1
        VertexArray(2).rhw = 1
        VertexArray(3).rhw = 1
        VertexArray(0).Color = Color
        VertexArray(1).Color = Color
        VertexArray(2).Color = Color
        VertexArray(3).Color = Color
        VertexArray(0).x = x + 0.5
        VertexArray(0).tu = (Frame * 32 / 512)
        VertexArray(0).y = y - 0.5
        VertexArray(0).tv = ((Sprite - 1) * 32 / 512)
        VertexArray(1).x = x - 31.5
        VertexArray(1).tu = (Frame * 32 + 32) / 512
        VertexArray(2).x = VertexArray(0).x
        VertexArray(3).x = VertexArray(1).x
        VertexArray(2).y = y + 31.5 - VCO
        VertexArray(2).tv = ((Sprite - 1) * 32 + 32 - VCO) / 512
        VertexArray(1).y = VertexArray(0).y
        VertexArray(1).tv = VertexArray(0).tv
        VertexArray(2).tu = VertexArray(0).tu
        VertexArray(3).y = VertexArray(2).y
        VertexArray(3).tu = VertexArray(1).tu
        VertexArray(3).tv = VertexArray(2).tv
    
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
        
    End If
End Sub
Sub DrawEquipmentShadow(ByVal x As Long, ByVal y As Long, ByVal Sprite As Long, ByVal Frame As Long, VCO As Long, eqType As Byte)
Dim VertexArray(0 To 3) As TLVERTEX
    Dim tColor As Long, Fade As Single, xOff As Single, yOff As Single
    
    If eqType = 1 Then
        If Frame = 6 Then
            Frame = 7
        ElseIf Frame = 7 Then
            Frame = 6
        End If
        If Frame = 9 Then
            Frame = 10
        ElseIf Frame = 10 Then
            Frame = 9
        End If
        x = x + frameOffsetsX(Frame)
        y = y + frameOffsetsY(Frame)
    End If
    
    xOff = 15.5
    yOff = 7.5
    
    
    'Fade = (AmbientAlpha) / 200
    Fade = 1
    'If Fade > 1 Then Fade = 1
    
    Select Case World.Hour
    
    Case 7 'day
        xOff = -18
    Case 8
        xOff = -16
    Case 9
        xOff = -14
    Case 10
        xOff = -12
    Case 11
        xOff = -10
    Case 12
        xOff = -8
    Case 13
        xOff = -6
    Case 14
        xOff = -4
    Case 15
        xOff = 4
    Case 16
        xOff = 6
    Case 17
        xOff = 8
    Case 18
        xOff = 10
    Case 19
        xOff = 12
    Case 20
        xOff = 14
    Case 21
        xOff = 16
    Case 22
        xOff = 18
    Case 23
        xOff = 22
        Fade = Fade - 0.2
    Case 24
        xOff = 22
        Fade = Fade - 0.4
    Case 1
        xOff = 15
        Fade = Fade - 0.6
    Case 2
        xOff = 8
        Fade = Fade - 0.8
    Case 3
        xOff = -8
        Fade = Fade - 0.8
    Case 4
        xOff = -15
        Fade = Fade - 0.6
    Case 5
        xOff = -22
        Fade = Fade - 0.4
    Case 6
        xOff = -22
        Fade = Fade - 0.2
        

    If Fade < 0 Then Fade = 0
    End Select
    
    
    tColor = D3DColorARGB(127 * Fade, 0, 0, 0)
    'Static LastTex As Long
    Dim CurTex As Long
    CurTex = (((Sprite - 1) \ 16) + 1)
    'If LastTex <> CurTex Then
        'D3DDevice.SetTexture 0, texSprites(CurTex).Texture
        ReadyTexture TexObjectFrames0 + CurTex
    '    LastTex = CurTex
    'End If
    'Set the RHWs (must always be 1)
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
    
    'Apply the colors
    VertexArray(0).Color = tColor
    VertexArray(1).Color = tColor
    VertexArray(2).Color = tColor
    VertexArray(3).Color = tColor

    'Find the left side of the rectangle
    VertexArray(0).x = x + xOff
    VertexArray(0).tu = (Frame * 32 / 512)

    'Find the top side of the rectangle
    VertexArray(0).y = y + 7.5
    VertexArray(0).tv = ((Sprite - 1) * 32 / 512)

    'Find the right side of the rectangle

    VertexArray(1).x = x + xOff + 32
    
    
    VertexArray(1).tu = (Frame * 32 + 32) / 512
    VertexArray(2).x = x - 0.5
    VertexArray(3).x = x + 31.5
    
    'Find the bottom of the rectangle
    VertexArray(2).y = y + 7.5 + 24 - VCO
    VertexArray(2).tv = ((Sprite - 1) * 32 + 32 - VCO) / 512

    'Because this is a perfect rectangle, all of the values below will equal one of the values we already got
    VertexArray(1).y = VertexArray(0).y
    VertexArray(1).tv = VertexArray(0).tv
    VertexArray(2).tu = VertexArray(0).tu
    VertexArray(3).y = VertexArray(2).y
    VertexArray(3).tu = VertexArray(1).tu
    VertexArray(3).tv = VertexArray(2).tv

    'Render the texture to the device
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size

End Sub

Sub DrawEquipment(ByVal x As Long, ByVal y As Long, ByVal Sprite As Long, ByVal Frame As Long, ByVal Color As Long, VCO As Long, eqType As Byte)
Dim VertexArray(0 To 3) As TLVERTEX


If eqType = 1 Then
    If Frame = 6 Then
        Frame = 7
    ElseIf Frame = 7 Then
        Frame = 6
    End If
    If Frame = 9 Then
        Frame = 10
    ElseIf Frame = 10 Then
        Frame = 9
    End If
    x = x + frameOffsetsX(Frame)
    y = y + frameOffsetsY(Frame)
End If
    Dim CurTex As Long
    CurTex = (((Sprite - 1) \ 16) + 1)
    If CurTex >= 1 And CurTex <= 16 Then
        ReadyTexture TexObjectFrames0 + CurTex
        VertexArray(0).rhw = 1
        VertexArray(1).rhw = 1
        VertexArray(2).rhw = 1
        VertexArray(3).rhw = 1
        VertexArray(0).Color = Color
        VertexArray(1).Color = Color
        VertexArray(2).Color = Color
        VertexArray(3).Color = Color
        VertexArray(0).x = x - 0.5
        VertexArray(0).tu = (Frame * 32 / 512)
        VertexArray(0).y = y - 0.5
        VertexArray(0).tv = ((Sprite - 1) * 32 / 512)
        VertexArray(1).x = x + 31.5
        VertexArray(1).tu = (Frame * 32 + 32) / 512
        VertexArray(2).x = VertexArray(0).x
        VertexArray(3).x = VertexArray(1).x
        VertexArray(2).y = y + 31.5 - VCO
        VertexArray(2).tv = ((Sprite - 1) * 32 + 32 - VCO) / 512
        VertexArray(1).y = VertexArray(0).y
        VertexArray(1).tv = VertexArray(0).tv
        VertexArray(2).tu = VertexArray(0).tu
        VertexArray(3).y = VertexArray(2).y
        VertexArray(3).tu = VertexArray(1).tu
        VertexArray(3).tv = VertexArray(2).tv
    
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
        
    End If
End Sub

Sub DrawLargeSprite(ByVal x As Long, ByVal y As Long, ByVal Sprite As Long, ByVal Frame As Long, ByVal Color As Long, Optional Shadow As Boolean = False, Optional VCO As Long = 0)
Dim VertexArray(0 To 3) As TLVERTEX
        Sprite = (Sprite * 4) - 3
    Dim CurTex As Long
              If Frame >= 8 Then
            Sprite = Sprite + 2
            Frame = Frame - 8
        End If
    CurTex = (((Sprite - 1) \ 16) + 1)
    If CurTex >= 1 And CurTex <= 20 Then

        If Shadow Then DrawLargeShadow x, y, Sprite, Frame, VCO
        ReadyTexture TexLSprite0 + CurTex
        VertexArray(0).rhw = 1
        VertexArray(1).rhw = 1
        VertexArray(2).rhw = 1
        VertexArray(3).rhw = 1
        VertexArray(0).Color = Color
        VertexArray(1).Color = Color
        VertexArray(2).Color = Color
        VertexArray(3).Color = Color
        VertexArray(0).x = x - 0.5
        VertexArray(0).tu = ((Frame * 64) / 512)
        VertexArray(0).y = y + 31.5
        VertexArray(0).tv = (((Sprite - 1) * 32 + 32) / 512)
        VertexArray(1).x = x + 63.5
        VertexArray(1).tu = (Frame * 64 + 64) / 512
        VertexArray(2).x = VertexArray(0).x
        VertexArray(3).x = VertexArray(1).x
        VertexArray(2).y = y + 63.5 - VCO
        VertexArray(2).tv = ((Sprite - 1) * 32 + 64 - VCO) / 512
        VertexArray(1).y = VertexArray(0).y
        VertexArray(1).tv = VertexArray(0).tv
        VertexArray(2).tu = VertexArray(0).tu
        VertexArray(3).y = VertexArray(2).y
        VertexArray(3).tu = VertexArray(1).tu
        VertexArray(3).tv = VertexArray(2).tv
    
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
        
    End If
End Sub
Sub DrawLargeSpriteTop(ByVal x As Long, ByVal y As Long, ByVal Sprite As Long, ByVal Frame As Long, ByVal Color As Long, Optional VCO As Long = 0)
Dim VertexArray(0 To 3) As TLVERTEX
        Sprite = (Sprite * 4) - 3
    Dim CurTex As Long
              If Frame >= 8 Then
            Sprite = Sprite + 2
            Frame = Frame - 8
        End If
    CurTex = (((Sprite - 1) \ 16) + 1)
    If CurTex >= 1 And CurTex <= 20 Then

        ReadyTexture TexLSprite0 + CurTex
        VertexArray(0).rhw = 1
        VertexArray(1).rhw = 1
        VertexArray(2).rhw = 1
        VertexArray(3).rhw = 1
        VertexArray(0).Color = Color
        VertexArray(1).Color = Color
        VertexArray(2).Color = Color
        VertexArray(3).Color = Color
        VertexArray(0).x = x - 0.5
        VertexArray(0).tu = (Frame * 64 / 512)
        VertexArray(0).y = y - 0.5
        VertexArray(0).tv = ((Sprite - 1) * 32 / 512)
        VertexArray(1).x = x + 63.5
        VertexArray(1).tu = (Frame * 64 + 64) / 512
        VertexArray(2).x = VertexArray(0).x
        VertexArray(3).x = VertexArray(1).x
        VertexArray(2).y = y + 31.5 - VCO
        VertexArray(2).tv = ((Sprite - 1) * 32 + 32 - VCO) / 512
        VertexArray(1).y = VertexArray(0).y
        VertexArray(1).tv = VertexArray(0).tv
        VertexArray(2).tu = VertexArray(0).tu
        VertexArray(3).y = VertexArray(2).y
        VertexArray(3).tu = VertexArray(1).tu
        VertexArray(3).tv = VertexArray(2).tv
    
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
        
    End If
End Sub


Public Sub DrawBmpString3D(St As String3D, ByVal x As Long, ByVal y As Long, Color As Long, Optional Centered As Boolean = True, Optional Mult As Single = 1, Optional bgColor As Long = &HFF000000)
Dim VertexArray(0 To 3) As TLVERTEX

    If Centered Then x = x - (St.Width / 2) * Mult

    LastTexture = 0
    D3DDevice.SetTexture 0, TexFont.Texture
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1


    Dim i As Long, CurrentChar As Long
    For i = 1 To Len(St.Text)
        CurrentChar = Asc(Mid$(St.Text, i, 1))
        If CurrentChar >= 32 And CurrentChar <= 128 Then
            With FontChar(2, CurrentChar)
                VertexArray(0).tu = .srcX / 512
                VertexArray(1).tu = (.srcX + .Width + 2) / 512
                VertexArray(0).tv = .srcY / 512
                VertexArray(2).tv = (.srcY + 18) / 512
                VertexArray(1).tv = VertexArray(0).tv
                VertexArray(2).tu = VertexArray(0).tu
                VertexArray(3).tu = VertexArray(1).tu
                VertexArray(3).tv = VertexArray(2).tv
        
'                'Draw Shadow Letter
                VertexArray(0).Color = bgColor
                VertexArray(1).Color = VertexArray(0).Color
                VertexArray(2).Color = VertexArray(0).Color
                VertexArray(3).Color = VertexArray(0).Color
                VertexArray(0).x = x + 0.5
                VertexArray(0).y = y + 1.5
                VertexArray(1).x = x + .Width * Mult + 2.5
                VertexArray(2).x = VertexArray(0).x
                VertexArray(3).x = VertexArray(1).x
                VertexArray(2).y = y + 19.5 * Mult
                VertexArray(1).y = VertexArray(0).y
                VertexArray(3).y = VertexArray(2).y
                
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
                
                'Draw Colored Text
                VertexArray(0).Color = Color
                VertexArray(1).Color = Color
                VertexArray(2).Color = Color
                VertexArray(3).Color = Color
                VertexArray(0).x = x - 1.5
                VertexArray(0).y = y - 0.5
                If Mult = 1 Then
                    VertexArray(1).x = x + .Width + 0.5
                Else
                    VertexArray(1).x = x + .Width * Mult
                End If
                VertexArray(2).x = VertexArray(0).x
                VertexArray(3).x = VertexArray(1).x
                VertexArray(2).y = y + 17.5 * Mult
                VertexArray(1).y = VertexArray(0).y
                VertexArray(3).y = VertexArray(2).y
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
                
               x = x + .Width * Mult
            End With
        End If
    Next i
End Sub

Public Sub DrawMultilineString3D(St As String, x As Long, y As Long, Width As Long, Optional Style As Long)
    Dim A As Long, b As Long, FoundLine As Boolean
    Dim Text As String, textHeight As Long, TextWidth As Long
    Dim curLine As Long
    Dim NextChar As Long

    textHeight = 16
    While St <> ""
        b = 0
        FoundLine = False
        For A = 1 To Len(St)
            NextChar = Asc(Mid$(St, A, 1))
            
            If NextChar <> Asc(vbCr) Then
                TextWidth = TextWidth + FontChar(2, NextChar).Width
            Else
                A = A + 1
            End If
            If TextWidth > Width Or NextChar = Asc(vbCr) Then
                FoundLine = True
                If b = 0 Then
                    b = A - 1
                End If
                If b > 0 Then
                    Text = Left$(St, b)
                    St = Mid$(St, b + 1)
                Else
                    Text = ""
                End If
                Exit For
            End If
            If Mid$(St, A, 1) = " " Then b = A
        Next A
        If FoundLine = False Then
            Text = St
            St = ""
        End If
        If Text <> "" Then
            If (Style And STYLE_CENTERED) Then
                DrawBmpString3D Create3DString(Text), x + ((Width - TextWidth) / 2), y + curLine * 17, &HFFFFFFFF, False
                TextWidth = 0
            Else
                DrawBmpString3D Create3DString(Text), x, y + curLine * 17, &HFFFFFFFF, False
                TextWidth = 0
            End If
            If FoundLine = True Then
                curLine = curLine + 1
            End If
        Else
            If St <> "" Then
                curLine = curLine + 1
            End If
        End If
    Wend
End Sub

Public Sub CreateLightMap()
    Dim VertexArray(0 To 3) As TLVERTEX
    Dim tx As Long, ty As Long
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1

   VertexArray(0).Color = -1: VertexArray(1).Color = -1: VertexArray(2).Color = -1: VertexArray(3).Color = -1
   VertexArray(0).tu = 0: VertexArray(0).tv = 0: VertexArray(1).tu = 1#: VertexArray(2).tv = 1#: VertexArray(1).tv = VertexArray(0).tv: VertexArray(2).tu = VertexArray(0).tu: VertexArray(3).tu = VertexArray(1).tu: VertexArray(3).tv = VertexArray(2).tv
    Dim tColor As Long
    Dim tSurface As Direct3DSurface8
    
    tColor = D3DColorARGB(255, 0, 0, 0)
       
    Set tSurface = texLightsEX.Texture.GetSurfaceLevel(0)
    D3DDevice.SetRenderTarget tSurface, Nothing, 0
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, ByVal tColor, 1, 0
    D3DDevice.BeginScene

    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    Dim A As Long, b As Long, C As Long
    For A = 0 To 29
        With LightSource(A)
            If .Type > LT_NONE Then
                If .Red > 0 Or .Blue > 0 Or .Green > 0 Then
                    

                    If AmbientAlpha <= 105 Then
                        b = .Intensity
                    Else
                        b = .Intensity - ((.Intensity * (AmbientAlpha - 105)) \ 200)
                    End If
                    VertexArray(0).Color = D3DColorARGB(b, .Red, .Green, .Blue)
                    VertexArray(1).Color = VertexArray(0).Color: VertexArray(2).Color = VertexArray(0).Color: VertexArray(3).Color = VertexArray(0).Color
                    b = .Radius + .Flicker
                    
                    tx = (.x / 384) * 256
                    ty = (.y / 384) * 256

                    If .Type = LT_PROJECTILE Then
                        ReadyTexture TexLight0 + 2
                        Select Case .D
                            Case 0 'Up
                                C = (b * 0.25) + 16
                                VertexArray(0).x = tx - b - 0.5
                                VertexArray(0).y = ty + b - 0.5 + C
                                VertexArray(1).x = tx + b - 0.5
                                VertexArray(2).y = ty - b - 0.5 + C
                                VertexArray(2).x = VertexArray(0).x: VertexArray(3).x = VertexArray(1).x: VertexArray(1).y = VertexArray(0).y: VertexArray(3).y = VertexArray(2).y
                            Case 1 'Down
                                C = (b * 0.25)
                                VertexArray(0).x = tx - b - 0.5
                                VertexArray(0).y = ty - b - 0.5 - C
                                VertexArray(1).x = tx + b - 0.5
                                VertexArray(2).y = ty + b - 0.5 - C
                                VertexArray(2).x = VertexArray(0).x: VertexArray(3).x = VertexArray(1).x: VertexArray(1).y = VertexArray(0).y: VertexArray(3).y = VertexArray(2).y
                            Case 2 'Left
                                C = (b * 0.5)
                                VertexArray(0).x = tx + b - 0.5 + C
                                VertexArray(0).y = ty + b - 0.5
                                VertexArray(1).y = ty - b - 0.5
                                VertexArray(2).x = tx - b - 0.5 + C
                                VertexArray(3).x = VertexArray(2).x: VertexArray(3).y = VertexArray(1).y: VertexArray(1).x = VertexArray(0).x: VertexArray(2).y = VertexArray(0).y
                            Case 3 'Right
                                C = (b * 0.5)
                                VertexArray(0).x = tx - b - 0.5 - C
                                VertexArray(0).y = ty + b - 0.5
                                VertexArray(1).y = ty - b - 0.5
                                VertexArray(2).x = tx + b - 0.5 - C
                                VertexArray(1).x = VertexArray(0).x: VertexArray(2).y = VertexArray(0).y: VertexArray(3).x = VertexArray(2).x: VertexArray(3).y = VertexArray(1).y
                        End Select
                    Else
                        ReadyTexture TexLight0 + 1
                        VertexArray(0).x = tx - b - 0.5
                        VertexArray(0).y = ty - b - 0.5
                        VertexArray(1).x = tx + b - 0.5
                        VertexArray(2).y = ty + b - 0.5
                        VertexArray(2).x = VertexArray(0).x: VertexArray(3).x = VertexArray(1).x: VertexArray(1).y = VertexArray(0).y: VertexArray(3).y = VertexArray(2).y
                    End If
                    
                  
                    
                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
                End If
            End If
        End With
    Next A
    
    
    
 '  tColor = D3DColorARGB(255, 0, 0, 0)
 '   VertexArray(0).Color = tColor: VertexArray(1).Color = tColor: VertexArray(2).Color = tColor: VertexArray(3).Color = tColor
 '  Dim x As Byte, y As Byte
 '   For x = 0 To 11
 '       For y = 0 To 11
 '           If map.Tile(x, y).Ground = 0 Then
 '               D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
 '               VertexArray(0).x = ((x * 32) / 384) * 256 - 0.5
 '               VertexArray(0).y = ((y * 32) / 384) * 256 - 1
 '               VertexArray(1).x = ((32 + x * 32) / 384) * 256 - 0.5
 '               VertexArray(2).y = ((32 + y * 32) / 384) * 256 - 0.5
 '               VertexArray(2).x = VertexArray(0).x: VertexArray(3).x = VertexArray(1).x: VertexArray(1).y = VertexArray(0).y: VertexArray(3).y = VertexArray(2).y
 '               D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
 '   '
 '   '            'D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
'    '            'tColor = D3DColorARGB(255, 255, 255, 255)
'    '            'VertexArray(0).Color = tColor: VertexArray(1).Color = tColor: VertexArray(2).Color = tColor: VertexArray(3).Color = tColor
'    '            'If map.Tile(x, y).Ground2 > 0 Then DrawTile x, y, map.Tile(x, y).Ground2
'    '            'If map.Tile(x, y).BGTile1 > 0 Then DrawTile x, y, map.Tile(x, y).BGTile1
'    '            'If map.Tile(x, y).FGTile > 0 Then DrawTile x, y, map.Tile(x, y).FGTile
'    '
'            End If
'        Next y
'    Next x
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.EndScene
    D3DDevice.SetRenderTarget RenderSurface(0), Nothing, 0
End Sub

Public Sub CreateShadeMap()
If AmbientAlpha > 245 And World.FlickerDark = 0 Then Exit Sub
    Dim VertexArray(0 To 3) As TLVERTEX
    Dim tx As Long, ty As Long
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
    
    VertexArray(0).Color = -1: VertexArray(1).Color = -1: VertexArray(2).Color = -1: VertexArray(3).Color = -1
    VertexArray(0).tu = 0: VertexArray(0).tv = 0: VertexArray(1).tu = 1#: VertexArray(2).tv = 1#: VertexArray(1).tv = VertexArray(0).tv: VertexArray(2).tu = VertexArray(0).tu: VertexArray(3).tu = VertexArray(1).tu: VertexArray(3).tv = VertexArray(2).tv
    
    Dim tColor As Long
    Dim tSurface As Direct3DSurface8
    Set tSurface = texShadeEX.Texture.GetSurfaceLevel(0)
    
    D3DDevice.SetRenderTarget tSurface, Nothing, 0

    tColor = D3DColorARGB(AmbientAlpha, AmbientAlpha, AmbientAlpha, AmbientAlpha)
    If World.FlickerDark > 0 Or FlickerCount Then
        If (255 * Rnd) <= World.FlickerDark Then
            FlickerCount = World.FlickerLength
        End If
        If FlickerCount Then
            FlickerCount = FlickerCount - 1
            tColor = D3DColorARGB(255, 0, 0, 0)
        End If
    End If
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, ByVal tColor, 1, 0
    D3DDevice.BeginScene
    
    Dim A As Long, b As Long
   
   
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    For A = 0 To 29
        With LightSource(A)
            If .Type > LT_NONE Then
                'D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
                'D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
                
                
                If .Type = LT_TILE And .MaxFlicker Then
                    If .FlickerCount = 0 Then
                        .Flicker = Int(Rnd * .MaxFlicker) \ 2
                        .FlickerCount = .FlickerRate
                    Else
                        .FlickerCount = .FlickerCount - 1
                    End If
                End If
                
                tColor = D3DColorARGB(255, .Intensity, .Intensity, .Intensity)
                VertexArray(0).Color = tColor: VertexArray(1).Color = tColor: VertexArray(2).Color = tColor: VertexArray(3).Color = tColor
                
                b = .Radius

                
                'D3DDevice.SetTexture 0, DynamicTextures(TexLight0 + 1).Texture
                
                tx = (.x / 384) * 256
                ty = (.y / 384) * 256

                 
                If .Type = LT_PROJECTILE Then
                    ReadyTexture TexLight0 + 2
                    Select Case .D
                        Case 0 'Up
                            ty = ty + (b * 0.25) + 11 'Take yOffset into consideration
                            VertexArray(0).x = tx - b - 0.5
                            VertexArray(0).y = ty + b - 0.5
                            VertexArray(1).x = tx + b - 0.5
                            VertexArray(2).y = ty - b - 0.5
                            VertexArray(2).x = VertexArray(0).x: VertexArray(3).x = VertexArray(1).x: VertexArray(1).y = VertexArray(0).y: VertexArray(3).y = VertexArray(2).y
                        Case 1 'Down
                            ty = ty - (b * 0.25)
                            VertexArray(0).x = tx - b - 0.5
                            VertexArray(0).y = ty - b - 0.5
                            VertexArray(1).x = tx + b - 0.5
                            VertexArray(2).y = ty + b - 0.5
                            VertexArray(2).x = VertexArray(0).x: VertexArray(3).x = VertexArray(1).x: VertexArray(1).y = VertexArray(0).y: VertexArray(3).y = VertexArray(2).y
                        Case 2 'Left
                            tx = tx + (b * 0.5)
                            VertexArray(0).x = tx + b - 0.5
                            VertexArray(0).y = ty + b - 0.5
                            VertexArray(1).y = ty - b - 0.5
                            VertexArray(2).x = tx - b - 0.5
                            VertexArray(3).x = VertexArray(2).x: VertexArray(3).y = VertexArray(1).y: VertexArray(1).x = VertexArray(0).x: VertexArray(2).y = VertexArray(0).y
                        Case 3 'Right
                            tx = tx - (b * 0.5)
                            VertexArray(0).x = tx - b - 0.5
                            VertexArray(0).y = ty + b - 0.5
                            VertexArray(1).y = ty - b - 0.5
                            VertexArray(2).x = tx + b - 0.5
                            VertexArray(1).x = VertexArray(0).x: VertexArray(2).y = VertexArray(0).y: VertexArray(3).x = VertexArray(2).x: VertexArray(3).y = VertexArray(1).y
                    End Select
                Else
                    ReadyTexture TexLight0 + 1
                    VertexArray(0).x = tx - b - 0.5
                    VertexArray(0).y = ty - b - 0.5
                    VertexArray(1).x = tx + b - 0.5
                    VertexArray(2).y = ty + b - 0.5
                    VertexArray(2).x = VertexArray(0).x: VertexArray(3).x = VertexArray(1).x: VertexArray(1).y = VertexArray(0).y: VertexArray(3).y = VertexArray(2).y
                End If
                
                D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
                
                '
                'GoTo tempskip
            End If
        End With
'tempskip:
    Next A
    
    'D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    
    'tColor = D3DColorARGB(255 - AmbientAlpha, 0, 0, 0)
    'VertexArray(0).Color = tColor: VertexArray(1).Color = tColor: VertexArray(2).Color = tColor: VertexArray(3).Color = tColor
    'Dim x As Byte, y As Byte
    'For x = 0 To 11
    '    For y = 0 To 11
    '        If map.Tile(x, y).Ground = 0 And map.Tile(x, y).Ground2 = 0 And map.Tile(x, y).BGTile1 = 0 And map.Tile(x, y).FGTile = 0 Then
    '            ReadyTexture TexLight0 + 3
    '            D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    '            VertexArray(0).x = ((x * 32) / 384) * 256 - 0.5
    '            VertexArray(0).y = ((y * 32) / 384) * 256 - 0.5
    '            VertexArray(1).x = ((32 + x * 32) / 384) * 256 - 0.5
    '            VertexArray(2).y = ((32 + y * 32) / 384) * 256 - 0.5
    '            VertexArray(2).x = VertexArray(0).x: VertexArray(3).x = VertexArray(1).x: VertexArray(1).y = VertexArray(0).y: VertexArray(3).y = VertexArray(2).y
    '            D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
    '
    '            'D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    '            'tColor = D3DColorARGB(AmbientAlpha, 255, 255, 255)
    '            'If map.Tile(x, y).Ground2 > 0 Then DrawTile2 x * 32, y * 32, map.Tile(x, y).Ground2, tColor
    '            'If map.Tile(x, y).BGTile1 > 0 Then DrawTile2 x * 32, y * 32, map.Tile(x, y).BGTile1, tColor
    '            'If map.Tile(x, y).FGTile > 0 Then DrawTile2 x * 32, y * 32, map.Tile(x, y).FGTile, tColor
    '
  '
  '          End If
   '     Next y
   ' Next x

    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.EndScene
    D3DDevice.SetRenderTarget RenderSurface(0), Nothing, 0
  




End Sub

Public Sub DrawShadeMap()
If AmbientAlpha > 245 Then Exit Sub
    Dim VertexArray(0 To 3) As TLVERTEX
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
    
    VertexArray(0).tu = 0: VertexArray(0).tv = 0: VertexArray(1).tu = 1#: VertexArray(2).tv = 1#: VertexArray(1).tv = VertexArray(0).tv: VertexArray(2).tu = VertexArray(0).tu: VertexArray(3).tu = VertexArray(1).tu: VertexArray(3).tv = VertexArray(2).tv:
    VertexArray(0).Color = -1: VertexArray(1).Color = -1: VertexArray(2).Color = -1: VertexArray(3).Color = -1
    
    LastTexture = 0
    D3DDevice.SetTexture 0, texShadeEX.Texture
    VertexArray(0).x = -0.5
    VertexArray(0).y = -0.5
    VertexArray(1).x = 383.5
    VertexArray(2).x = VertexArray(0).x
    VertexArray(3).x = VertexArray(1).x
    VertexArray(2).y = 383.5
    VertexArray(1).y = VertexArray(0).y
    VertexArray(3).y = VertexArray(2).y
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTCOLOR
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ZERO
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub

Public Sub DrawLightMap()
    Dim VertexArray(0 To 3) As TLVERTEX
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
    
    VertexArray(0).tu = 0: VertexArray(0).tv = 0: VertexArray(1).tu = 1#: VertexArray(2).tv = 1#: VertexArray(1).tv = VertexArray(0).tv: VertexArray(2).tu = VertexArray(0).tu: VertexArray(3).tu = VertexArray(1).tu: VertexArray(3).tv = VertexArray(2).tv:
    VertexArray(0).Color = -1: VertexArray(1).Color = -1: VertexArray(2).Color = -1: VertexArray(3).Color = -1
    
    LastTexture = 0
    D3DDevice.SetTexture 0, texLightsEX.Texture
    VertexArray(0).x = -0.5
    VertexArray(0).y = -0.5
    VertexArray(1).x = 383.5
    VertexArray(2).x = VertexArray(0).x
    VertexArray(3).x = VertexArray(1).x
    VertexArray(2).y = 383.5
    VertexArray(1).y = VertexArray(0).y
    VertexArray(3).y = VertexArray(2).y
    
    Dim tColor As Long
    tColor = D3DColorARGB(255, 255, 255, 255)
    VertexArray(0).Color = tColor: VertexArray(1).Color = tColor: VertexArray(2).Color = tColor: VertexArray(3).Color = tColor
    
   ' D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_DESTCOLOR
    
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
End Sub

Public Sub ShadeMap(A As Byte, R As Byte, G As Byte, b As Byte)
Dim VertexArray(0 To 3) As TLVERTEX
    Dim tColor As Long
    tColor = D3DColorARGB(A, R, G, b)
    LastTexture = 0
    D3DDevice.SetTexture 0, texShade.Texture
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1

    'Apply the colors
    VertexArray(0).Color = tColor
    VertexArray(1).Color = tColor
    VertexArray(2).Color = tColor
    VertexArray(3).Color = tColor
    
    VertexArray(0).x = -0.5
    VertexArray(0).tu = 0
    VertexArray(0).y = -0.5
    VertexArray(0).tv = 0
    VertexArray(1).x = 383.5
    VertexArray(1).tu = 1#
    VertexArray(2).x = VertexArray(0).x
    VertexArray(3).x = VertexArray(1).x
    VertexArray(2).y = 383.5
    VertexArray(2).tv = 1#
    VertexArray(1).y = VertexArray(0).y
    VertexArray(1).tv = VertexArray(0).tv
    VertexArray(2).tu = VertexArray(0).tu
    VertexArray(3).y = VertexArray(2).y
    VertexArray(3).tu = VertexArray(1).tu
    VertexArray(3).tv = VertexArray(2).tv

    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
End Sub

Public Sub Draw3D(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal tx As Long, ByVal ty As Long, TextureNum As Long, Optional ByVal Color As Long = -1, Optional texWidth As Long = 0, Optional texHeight As Long = 0)

    Dim VertexArray(0 To 3) As TLVERTEX
    If TextureNum > 0 And TextureNum <= NumTextures Then
        'D3DDevice.SetTexture 0, Texture.Texture
        If ReadyTexture(TextureNum) = False Then Exit Sub
        texWidth = DynamicTextures(TextureNum).TexInfo.Width
        texHeight = DynamicTextures(TextureNum).TexInfo.Height
    Else
        LastTexture = 0
    End If
    
    
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
    
    'Apply the colors
    VertexArray(0).Color = Color
    VertexArray(1).Color = Color
    VertexArray(2).Color = Color
    VertexArray(3).Color = Color
    
    VertexArray(0).x = x - 0.5
    VertexArray(0).tu = tx / texWidth
    VertexArray(0).y = y - 0.5
    VertexArray(0).tv = ty / texHeight
    VertexArray(1).x = x + Width - 0.5
    VertexArray(1).tu = (tx + Width) / texWidth
    VertexArray(2).x = VertexArray(0).x
    VertexArray(3).x = VertexArray(1).x
    VertexArray(2).y = y + Height - 0.5
    VertexArray(2).tv = (ty + Height) / texHeight
    VertexArray(1).y = VertexArray(0).y
    VertexArray(1).tv = VertexArray(0).tv
    VertexArray(2).tu = VertexArray(0).tu
    VertexArray(3).y = VertexArray(2).y
    VertexArray(3).tu = VertexArray(1).tu
    VertexArray(3).tv = VertexArray(2).tv

    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
End Sub

Sub forceDraw(ByVal A As Byte, ByVal R As Byte, ByVal G As Byte, ByVal b As Byte)
    D3DDevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, ByVal 0, 1, 0
    If (Character.StatusEffect And (2 ^ SE_BLIND)) = 0 Then
        D3DDevice.BeginScene
            DrawNextFrame3D
            If MapEdit = False Then DrawShadeMap
            ShadeMap A, R, G, b
        D3DDevice.EndScene
    End If
    If D3DDevice.TestCooperativeLevel = D3D_OK Then
        D3DDevice.Present ByVal 0&, ByVal 0&, 0, ByVal 0&
         
        If cX <> LastCX Or cY <> LastCY Then
            DrawMiniMapSection
        End If
    End If
End Sub
Sub Transition(ByRef TransitionType As Long, ByVal R As Byte, ByVal G As Byte, ByVal b As Byte, ByVal Delay As Long)
    Dim A As Long, start As Long

Delay = 10
'If Not Options.fullredraws Then
'    setPartialRedraw = True
'    Options.fullredraws = True
'End If

    DisableScreen = False

    start = GetTickCount
    If TransitionType = 1 Then CreateShadeMap
    For Trans = 1 To 16
    
    A = (255 * (GetTickCount - start))
    A = A / (150)
    If A > 255 Then A = 255
    If A < 0 Then A = 0
    If TransitionType = 1 Then A = 255 - A
    forceDraw A, R, b, G
    While GetTickCount - start < Delay * Trans
    Wend
    If GetTickCount - start > Delay * 17 Then Exit For
   Next Trans
   Trans = -1
   If TransitionType = 0 Then
        DisableScreen = True
   End If
   If TransitionType = 2 Then
        Freeze = False
   End If
'If setPartialRedraw Then
'    Options.fullredraws = False
'End If

End Sub

Function CheckForDrawingErrors()
    Dim A As Long
        CheckForDrawingErrors = False
                If DD.TestCooperativeLevel <> DD_OK Then
                    'Do While DD.TestCooperativeLevel <> DD_OK
                        Select Case DD.TestCooperativeLevel
                            Case DDERR_WRONGMODE
                                Debug.Print "wrong mode "
                                
                            Case Else
                                Debug.Print "Other error: " & DD.TestCooperativeLevel
                        End Select
                    'Loop
                    
                    Init_Directx
                    InitializeTextures
                    BOOLD3DERROR = True
                    STOPMAPCHECK = True
                    InitMiniMap
                    If CurrentTab = tsMap Then
                        SetMiniMapTab tsMap
                    End If
                    Draw sfcTimers2, 0, 0, 64, 32, sfcInventory(0), 148, 393, True, 0
                End If
                If D3DDevice.TestCooperativeLevel <> D3D_OK Then
                    Do Until D3DDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET
                        'Sleep 1
                        DoEvents
                    Loop
        
                    For A = 1 To NumTextures
                        Set DynamicTextures(A).Texture = Nothing
                        DynamicTextures(A).Loaded = False
                        DynamicTextures(A).LastUsed = 0
                    Next A
                    
                    A = CurFog
                    ClearFog (True)
                    If Not InitD3D(False) Then
                       CheckForDrawingErrors = True
                    End If
                    If A > 0 And A <= 31 Then
                        InitFog A, False
                    End If
                    


                    InitMiniMap
                    If CurrentTab = tsMap Then
                        SetMiniMapTab tsMap
                    End If
                End If
                           
End Function


Public Sub Draw3DEx(ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal tx As Long, ByVal ty As Long, Texture As TextureType, Optional ByVal Color As Long = -1)
    
    Dim VertexArray(0 To 3) As TLVERTEX
    D3DDevice.SetTexture 0, Texture.Texture
    VertexArray(0).rhw = 1
    VertexArray(1).rhw = 1
    VertexArray(2).rhw = 1
    VertexArray(3).rhw = 1
    
    'Apply the colors
    VertexArray(0).Color = Color
    VertexArray(1).Color = Color
    VertexArray(2).Color = Color
    VertexArray(3).Color = Color
    VertexArray(0).x = x - 0.5
    VertexArray(0).tu = tx / Texture.TexInfo.Width
    VertexArray(0).y = y - 0.5
    VertexArray(0).tv = ty / Texture.TexInfo.Height
    VertexArray(1).x = x + Width - 0.5
    VertexArray(1).tu = (tx + Width) / Texture.TexInfo.Width
    VertexArray(2).x = VertexArray(0).x
    VertexArray(3).x = VertexArray(1).x
    VertexArray(2).y = y + Height - 0.5
    VertexArray(2).tv = (ty + Height) / Texture.TexInfo.Height
    VertexArray(1).y = VertexArray(0).y
    VertexArray(1).tv = VertexArray(0).tv
    VertexArray(2).tu = VertexArray(0).tu
    VertexArray(3).y = VertexArray(2).y
    VertexArray(3).tu = VertexArray(1).tu
    VertexArray(3).tv = VertexArray(2).tv

    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VertexArray(0), FVF_Size
End Sub
Public Sub DrawRect(x1 As Long, x2 As Long, y1 As Long, y2 As Long, FillStyle As Long, FillColor As Long, BorderSize As Integer, BorderColor As Long)
    Dim R1(0 To 3) As D3DRECT
    
    If x1 >= x2 Then Exit Sub
    If y1 >= y2 Then Exit Sub
    
    R1(0).x1 = x1: R1(0).x2 = x2: R1(0).y1 = y1: R1(0).y2 = y2
    
    Select Case FillStyle
        Case BS_SOLID
            D3DDevice.Clear 1, R1(0), D3DCLEAR_TARGET, FillColor, 1#, 0
    End Select
    
    If BorderSize > 0 Then
        R1(0).x1 = x1: R1(0).x2 = x2: R1(0).y1 = y1: R1(0).y2 = y1 + 1 'Top
        R1(1).x1 = x1: R1(1).x2 = x1 + 1: R1(1).y1 = y1: R1(1).y2 = y2 'Left
        R1(2).x1 = x2: R1(2).x2 = x2 + 1: R1(2).y1 = y1: R1(2).y2 = y2 'Right
        R1(3).x1 = x1: R1(3).x2 = x2 + 1: R1(3).y1 = y2: R1(3).y2 = y2 + 1 'Bottom
        D3DDevice.Clear 4, R1(0), D3DCLEAR_TARGET, BorderColor, 1#, 0
    End If
End Sub

Public Sub DrawNextFrame3D()
    Dim x As Long, y As Long, VCO As Long
    Dim X32 As Long, Y32 As Long, XOffset As Long, YOffset As Long
    Dim A As Long, b As Long, C As Long, D As Long, VertexArray(0 To 3) As TLVERTEX
    
    
    If Not Options.fullredraws Then
        If mapChanged Then
            D3DDevice.SetRenderTarget mapTexture(2).GetSurfaceLevel(0), Nothing, 0
            D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, ByVal &H0&, 1#, 0
            D3DDevice.SetRenderTarget mapTexture(1).GetSurfaceLevel(0), Nothing, 0
            D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, ByVal &H0&, 1#, 0
        End If
        
        If mapFGChanged Then
            D3DDevice.SetRenderTarget mapTexture(2).GetSurfaceLevel(0), Nothing, 0
            D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, ByVal &H0&, 1#, 0
        End If
        
        D3DDevice.SetRenderTarget mapTexture(1).GetSurfaceLevel(0), Nothing, 0
        If MapEdit Then D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, ByVal &H0&, 1#, 0
    End If
    
    For x = 0 To 11
        For y = 0 To 11
            b = 0
            A = 0
            If MapEdit = False Then
                With map.Tile(x, y)
                    'If (.Anim(1) \ 16) > 0 Then
                        If (.Anim(1) And 15) > 0 Then
                        A = TileAnim(x, y).Frame2
                        TileAnim(x, y).Frame2 = (GetTickCount / (AnimationFps * (.Anim(1) And 15) * 4) + TileAnim(x, y).Frame) Mod ((.Anim(1) \ 16) + 1)
                        End If
                    'End If
                   
                    If Options.fullredraws Or mapChangedBg(x, y) Or A <> TileAnim(x, y).Frame2 Then
                        DrawRect x * 32, x * 32 + 32, y * 32, y * 32 + 32, BS_SOLID, D3DColorARGB(255, 0, 0, 0), 0, 0
                        mapChangedBg(x, y) = False
                        
                        X32 = x * 32
                        Y32 = y * 32
                        
                        'Draw Lower Portion Of Map
                        If map.Tile(x, y).Ground > 0 Then
                            If .Att = 23 Then
                                If ExamineBit(.AttData(2), 0) Then
                                    If ExamineBit(.AttData(3), 0) Then XOffset = -.AttData(0) Else XOffset = .AttData(0)
                                    If ExamineBit(.AttData(3), 1) Then YOffset = -.AttData(1) Else YOffset = .AttData(1)
                                    
                                    If Not Options.fullredraws Then
                                        If (.Anim(1) \ 16) = 0 Then mapChangedBg(x, y) = True
                                        C = 0
                                        D = 0
                                        If XOffset <> 0 Then C = XOffset / 32 + IIf(XOffset Mod 32 <> 0, XOffset / Abs(XOffset), 0)
                                        If YOffset <> 0 Then D = YOffset / 32 + IIf(YOffset Mod 32 <> 0, YOffset / Abs(YOffset), 0)
                                        If (C <> 0 Or D <> 0) Then
                                            If (x + C >= 0 And x + C <= 11) Then
                                                mapChangedBg(x + C, y) = True
                                                If y + D >= 0 And y + D <= 11 Then mapChangedBg(x + C, y + D) = True
                                            End If
                                            If (y + D >= 0 And y + D <= 11) Then mapChangedBg(x, y + D) = True
                                            If (x + C - 1 >= 0 And x + C - 1 <= 11) And (y + D - 1 >= 0 And y + D - 1 <= 11) Then mapChangedBg(x + C - 1, y + D - 1) = True
                                        End If
                                    End If
                                    
                                Else
                                    XOffset = 0: YOffset = 0
                                End If
                            Else
                                XOffset = 0: YOffset = 0
                            End If
                            b = 0
                            If (.Anim(2) And 3) = 0 Then b = TileAnim(x, y).Frame2
                            DrawTile X32 + XOffset, Y32 + YOffset, map.Tile(x, y).Ground + b
                        End If
                        If map.Tile(x, y).Ground2 > 0 Then
                            If .Att = 23 Then
                                If ExamineBit(.AttData(2), 1) Then
                                    If ExamineBit(.AttData(3), 0) Then XOffset = -.AttData(0) Else XOffset = .AttData(0)
                                    If ExamineBit(.AttData(3), 1) Then YOffset = -.AttData(1) Else YOffset = .AttData(1)
                                    
                                    If Not Options.fullredraws Then ' And (.Anim(1) \ 16) > 0 Then
                                        If (.Anim(1) \ 16) = 0 Then mapChangedBg(x, y) = True
                                        C = 0
                                        D = 0
                                    If XOffset <> 0 Then C = XOffset / 32 + IIf(XOffset Mod 32 <> 0, XOffset / Abs(XOffset), 0)
                                    If YOffset <> 0 Then D = YOffset / 32 + IIf(YOffset Mod 32 <> 0, YOffset / Abs(YOffset), 0)
                                        If (C <> 0 Or D <> 0) Then
                                        If (x + C >= 0 And x + C <= 11) Then
                                            mapChangedBg(x + C, y) = True
                                            If y + D >= 0 And y + D <= 11 Then mapChangedBg(x + C, y + D) = True
                                        End If
                                        If (y + D >= 0 And y + D <= 11) Then mapChangedBg(x, y + D) = True
                                        If (x + C - 1 >= 0 And x + C - 1 <= 11) And (y + D - 1 >= 0 And y + D - 1 <= 11) Then mapChangedBg(x + C - 1, y + D - 1) = True
                                    End If
                                    End If
                                    
                                Else
                                    XOffset = 0: YOffset = 0
                                End If
                            Else
                                XOffset = 0: YOffset = 0
                            End If
                            b = 0
                            If (.Anim(2) And 3) = 1 Then b = TileAnim(x, y).Frame2
                            DrawTile X32 + XOffset, Y32 + YOffset, map.Tile(x, y).Ground2 + b
                        End If
                        If .Att = 19 Then
                            If .AttData(0) > 0 Then
                                Select Case .AttData(1)
                                    Case 0
                                        If .AttData(2) = 1 Then
                                            DrawObject X32, Y32, .AttData(0) + 255
                                        Else
                                            DrawObject X32, Y32, .AttData(0)
                                        End If
                                    Case 1
                                        DrawSprite X32, Y32, .AttData(0), .AttData(2), -1
                                End Select
                            End If
                        ElseIf .Att = 25 Then
                            If NPC(.AttData(0)).Sprite > 0 Then
                                If CurX = x And CurY = y Then
                                    DrawSpriteGlow X32, Y32, NPC(.AttData(0)).Sprite, NPC(.AttData(0)).Direction * 3, &HAF00FF00
                                End If
                                DrawSprite X32, Y32, NPC(.AttData(0)).Sprite, NPC(.AttData(0)).Direction * 3, -1
    
                            End If
                        End If
                        
                        If .BGTile1 > 0 Then
                            If .Att = 23 Then
                                If ExamineBit(.AttData(2), 2) Then
                                    If ExamineBit(.AttData(3), 0) Then XOffset = -.AttData(0) Else XOffset = .AttData(0)
                                    If ExamineBit(.AttData(3), 1) Then YOffset = -.AttData(1) Else YOffset = .AttData(1)
                                    
                                    If Not Options.fullredraws Then ' And (.Anim(1) \ 16) > 0 Then
                                        If (.Anim(1) \ 16) = 0 Then mapChangedBg(x, y) = True
                                        C = 0
                                        D = 0
                                    If XOffset <> 0 Then C = XOffset / 32 + IIf(XOffset Mod 32 <> 0, XOffset / Abs(XOffset), 0)
                                    If YOffset <> 0 Then D = YOffset / 32 + IIf(YOffset Mod 32 <> 0, YOffset / Abs(YOffset), 0)
                                        If (C <> 0 Or D <> 0) Then
                                        If (x + C >= 0 And x + C <= 11) Then
                                            mapChangedBg(x + C, y) = True
                                            If y + D >= 0 And y + D <= 11 Then mapChangedBg(x + C, y + D) = True
                                        End If
                                        If (y + D >= 0 And y + D <= 11) Then mapChangedBg(x, y + D) = True
                                        If (x + C - 1 >= 0 And x + C - 1 <= 11) And (y + D - 1 >= 0 And y + D - 1 <= 11) Then mapChangedBg(x + C - 1, y + D - 1) = True
                                    End If
                                    End If
                                    
                                Else
                                    XOffset = 0: YOffset = 0
                                End If
                            Else
                                XOffset = 0: YOffset = 0
                            End If
                            b = 0
                            If (.Anim(2) And 3) = 2 Then b = TileAnim(x, y).Frame2
                            If .Att <> 24 Or .AttData(3) <> 0 Then DrawTile X32 + XOffset, Y32 + YOffset, map.Tile(x, y).BGTile1 + b
                        End If
                        If .FGTile > 0 Then
                            If .Att = 18 Then
                                If .AttData(0) = 0 And .AttData(1) = 0 And .AttData(2) = 0 And .AttData(3) = 0 Then
                                    DrawTile X32, Y32, .FGTile, 0, 16, 16, 32
                                Else
                                    If .AttData(0) And 1 Then
                                        If .AttData(2) = 0 Then
                                            DrawTile X32, Y32, .FGTile, 16, 0, 16, 16
                                        End If
                                    End If
                                    If .AttData(0) And 2 Then
                                        If .AttData(2) = 0 Then
                                            DrawTile X32, Y32, .FGTile, 0, 0, 16, 16
                                        End If
                                    End If
                                    If .AttData(1) And 1 Then
                                        If .AttData(2) = 0 Then
                                            DrawTile X32, Y32, .FGTile, 0, 16, 16, 16
                                        End If
                                    End If
                                    If .AttData(1) And 2 Then
                                        If .AttData(2) = 0 Then
                                            DrawTile X32, Y32, .FGTile, 16, 16, 16, 16
                                        End If
                                    End If
                                End If
                            ElseIf .Att = 4 Then
                                DrawTile X32, Y32, .FGTile, 0, 16, 16, 32
                            End If
                        End If
                    End If
                End With
            Else
                With EditMap.Tile(x, y)
                    'If (.Anim(1) \ 16) > 0 Then
                    If (.Anim(1) And 15) Then
                        If TileAnim(x, y).AnimDelay = 0 Then
                            TileAnim(x, y).Frame = TileAnim(x, y).Frame + 1
                            
                            If TileAnim(x, y).Frame > (.Anim(1) \ 16) Then TileAnim(x, y).Frame = 0
                            TileAnim(x, y).AnimDelay = ((.Anim(1) And 15) * 4 - 1)
                        Else
                            TileAnim(x, y).AnimDelay = TileAnim(x, y).AnimDelay - 1
                        End If
                        A = TileAnim(x, y).Frame
                    End If
                    X32 = x * 32
                    Y32 = y * 32
                    
                    'Draw Lower Portion Of Map
                    If .Ground > 0 Then
                        If (.Anim(2) And 3) = 0 Then b = A Else b = 0
                        DrawTile X32 + XOffset, Y32 + YOffset, .Ground + b
                    End If
                    If .Ground2 > 0 Then
                        If (.Anim(2) And 3) = 1 Then b = A Else b = 0
                        DrawTile X32, Y32, .Ground2 + b
                    End If
                    If .Att = 19 Then
                        If .AttData(0) > 0 Then
                            Select Case .AttData(1)
                                Case 0
                                    If .AttData(2) = 1 Then
                                        DrawObject X32, Y32, .AttData(0) + 255
                                    Else
                                        DrawObject X32, Y32, .AttData(0)
                                    End If
                                Case 1
                                    DrawSprite X32, Y32, .AttData(0), .AttData(2), -1
                            End Select
                        End If
                    ElseIf .Att = 25 Then
                        If NPC(.AttData(0)).Sprite > 0 Then
                            DrawSprite X32, Y32, NPC(.AttData(0)).Sprite, NPC(.AttData(0)).Direction * 3, -1
                        End If
                    End If
                    If .BGTile1 > 0 Then
                        If (.Anim(2) And 3) = 2 Then b = A Else b = 0
                        DrawTile X32, Y32, .BGTile1 + b
                    End If
                End With
            End If
        Next y
    Next x
    
    If Not Options.fullredraws Then
        D3DDevice.SetRenderTarget RenderSurface(0), Nothing, 0
        D3DDevice.SetTexture 0, mapTexture(1)
        LastTexture = 0
        Draw3D 0, 0, 384, 384, 0, 0, 0, -1, 512, 512
    End If
    
    If CastingSpell Then
        Dim tx As Long, ty As Long
                    tx = CurX * 32 + CurSubX - 16 * WindowScaleX
                    ty = CurY * 32 + CurSubY - 16 * WindowScaleY
                    
                    tx = tx / TileSizeX
                    ty = ty / TileSizeY
                    
                    If tx > 11 Then tx = 11
                    If ty > 11 Then ty = 11
                    If tx < 0 Then tx = 0
                    If ty < 0 Then ty = 0
                    
        If Sqr((tx - cX) ^ 2 + (ty - cY) ^ 2) > Skills(Character.MacroSkill).Range Then
            Draw3D tx * 32, ty * 32, 32, 32, 64, 224, TexControl1 'red target circle
        ElseIf Not LOS(cX, cY, tx, ty, 1) Then
            Draw3D tx * 32, ty * 32, 32, 32, 32, 224, TexControl1 'red target circle
        Else
            Draw3D tx * 32, ty * 32, 32, 32, 96, 224, TexControl1 'blue target circle
        End If
    End If
    
    For A = 0 To MaxTraps
        With map.Trap(A)
            If .Created > 0 Then
                If .Created > 0 And .Counter = 0 Then
                    .Created = 0
                    FloatingText.Add .x * 32, .y * 32, "Trap Set!", 0
                End If
                If .Created <> 0 Then
                    DrawObject .x * 32, .y * 32, 106
                End If
            End If
        End With
    Next A
    
    For A = 0 To 49
        With map.Object(A)
            If .Object > 0 Then
                b = Object(.Object).Picture
                If b > 0 Then
                    If .TimeStamp = 0 Or map.Tile(.x, .y).Att = 5 Then
                        D = 255
                    Else
                        If .TimeStamp > 12000 Or (GetTickCount Mod 1000) / 500 > 1 Then
                            If .DeathObj Then
                                D = 90 + 165 * CDec(.TimeStamp) / 3600000
                            Else
                                D = 90 + 165 * CDec(.TimeStamp) / 180000
                            End If
                        Else
                            D = 0
                        End If
                    End If
                        
                    If D > 0 Then
                        If ExamineBit(Object(.Object).Flags, 6) Then b = b + 255
                        If .Prefix > 0 Or .Suffix > 0 Or .Affix > 0 Then
                            C = ((.PrefixVal \ 64)) ' And 7)
                            If ((.SuffixVal \ 64) And 3) > C Then C = ((.SuffixVal \ 64)) ' And 7)
                            If ((.AffixVal \ 64) And 3) > C Then C = ((.AffixVal \ 64)) 'And 7)
                            
                            Select Case C
                                Case 3
                                    DrawObjectGlow .x * 32 + .XOffset, .y * 32 + .YOffset, b, D3DColorARGB(TargetPulse * D / 255, &HD3, &H1C, &HFB)
                                Case 2
                                    DrawObjectGlow .x * 32 + .XOffset, .y * 32 + .YOffset, b, D3DColorARGB(TargetPulse * D / 255, &HFF, &HFF, &H7F)
                                Case 1
                                    DrawObjectGlow .x * 32 + .XOffset, .y * 32 + .YOffset, b, D3DColorARGB(TargetPulse * D / 255, &H66, &HF4, &H7E)
                                Case 0
                                    DrawObjectGlow .x * 32 + .XOffset, .y * 32 + .YOffset, b, D3DColorARGB(TargetPulse * D / 255, &H95, &HDB, &HD8)
                            End Select
                        End If
                        
                        DrawObject .x * 32 + .XOffset, .y * 32 + .YOffset, b, D
                    End If
                End If
            End If
        End With
    Next A

    Dim tmpPS As clsParticleSource
    For Each tmpPS In ParticleEngineB
        With tmpPS
            .Render
        End With
    Next

    For A = 0 To 9
        With map.DeadBody(A)
            If .Sprite > 0 Then
                If .Frame = 13 Then
                    b = 255 - (255 * ((450 - .Counter) / 450))
                    If b < 0 Then b = 0
                    If b > 255 Then b = 255
                    If Monster(.MonNum).Flags2 And MONSTER_LARGE Then
                        DrawLargeSprite .x * 32, .y * 32, .Sprite, 12, D3DColorARGB(b, 255, 255, 255), False
                    Else
                        If Monster(.MonNum).Flags2 And MONSTER_SPRITE255 Then
                            DrawSprite .x * 32, .y * 32, .Sprite + 255, 12, D3DColorARGB(b, 255, 255, 255), False
                        Else
                            DrawSprite .x * 32, .y * 32, .Sprite, 12, D3DColorARGB(b, 255, 255, 255), False
                        End If
                    End If
                End If
            End If
        End With
    Next A

    For A = 0 To 9
        With map.Monster(A)
            If .Monster > 0 Then
                C = Monster(.Monster).Sprite
                If Monster(.Monster).Flags2 And MONSTER_LARGE Then
                    If C > 0 And C < 255 Then
                        'Draw LARGE Monster
                        If .A > 0 Then
                            b = .D * 3 + 2
                        Else
                            b = .D * 3 + .W
                        End If
                        If map.Tile(.x, .y).Att = 20 Then
                            VCO = map.Tile(.x, .y).AttData(0)
                        ElseIf map.Tile(.x, .y).Att = 26 Then
                            VCO = map.Tile(.x, .y).AttData(3)
                        Else
                            VCO = 0
                        End If
                        If CurrentTarget.TargetType = TT_MONSTER Then
                            If CurrentTarget.Target = A Then
                                'DrawSpriteGlow .XO, .YO - 16, C, B, D3DColorARGB(TargetPulse, 255, 0, 0), VCO
                                DrawLargeSpriteGlow .XO, .YO - 16, C, b, D3DColorARGB(TargetPulse * (Monster(.Monster).alpha - .alpha) / 255, 255, 0, 0), VCO
                                
                            End If
                        End If
                        DrawLargeSprite .XO, .YO - 16, C, b, D3DColorARGB(Monster(.Monster).alpha - .alpha, Monster(.Monster).Red - .R, Monster(.Monster).Green - .G, Monster(.Monster).Blue) - .b, IIf(Monster(.Monster).Flags And MONSTER_NO_SHADOW, False, True), VCO
                        'draw hp bar
                        If A = TargetMonster Then
                            If (Monster(.Monster).alpha - .alpha) >= 100 Then
                                DrawRect .XO + 30 + 32, .XO + 32 + 32, .YO - 16, .YO + 48, BS_SOLID, D3DColorARGB(255, 0, 0, 0), 0, 0
                                If Monster(.Monster).MaxHP > 0 Then
                                    b = (64 * CLng(.HP) \ Monster(.Monster).MaxHP)
                                    b = 64 - b
                                    If b < 0 Then b = 0
                                    If b > 64 Then b = 64
                                    DrawRect .XO + 30 + 32, .XO + 32 + 32, .YO - 16 + b, .YO + 48, BS_SOLID, D3DColorARGB(255, 255, 0, 0), 0, 0
                                End If
                            End If
                        End If
                    End If
                Else
                    If Monster(.Monster).Flags2 And MONSTER_MEDIUM Then
                        If C > 0 And C < 255 Then
                            'Draw LARGE Monster
                            If .A > 0 Then
                                b = .D * 3 + 2
                            Else
                                b = .D * 3 + .W
                            End If
                            If map.Tile(.x, .y).Att = 20 Then
                                VCO = map.Tile(.x, .y).AttData(0)
                            ElseIf map.Tile(.x, .y).Att = 60 Then
                                VCO = map.Tile(.x, .y).AttData(3)
                            Else
                                VCO = 0
                            End If
                            If CurrentTarget.TargetType = TT_MONSTER Then
                                If CurrentTarget.Target = A Then
                                    'DrawSpriteGlow .XO, .YO - 16, C, B, D3DColorARGB(TargetPulse, 255, 0, 0), VCO
                                    DrawLargeSpriteGlow .XO - 16, .YO - 48, C, b, D3DColorARGB(TargetPulse * (Monster(.Monster).alpha - .alpha) / 255, 255, 0, 0), VCO
                                    
                                End If
                            End If
                            DrawLargeSprite .XO - 16, .YO - 48, C, b, D3DColorARGB(Monster(.Monster).alpha - .alpha, Monster(.Monster).Red - .R, Monster(.Monster).Green - .G, Monster(.Monster).Blue) - .b, IIf(Monster(.Monster).Flags And MONSTER_NO_SHADOW, False, True), VCO
                            'draw hp bar
                            If A = TargetMonster Then
                                If (Monster(.Monster).alpha - .alpha) >= 100 Then
                                    DrawRect .XO + 30 + 16, .XO + 32 + 16, .YO - 48, .YO + 16, BS_SOLID, D3DColorARGB(255, 0, 0, 0), 0, 0
                                    If Monster(.Monster).MaxHP > 0 Then
                                        b = (64 * CLng(.HP) \ Monster(.Monster).MaxHP)
                                        b = 64 - b
                                        If b < 0 Then b = 0
                                        If b > 64 Then b = 64
                                        DrawRect .XO + 30 + 16, .XO + 32 + 16, .YO - 48 + b, .YO + 16, BS_SOLID, D3DColorARGB(255, 255, 0, 0), 0, 0
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If C > 0 And C < 255 Then
                            If Monster(.Monster).Flags2 And MONSTER_SPRITE255 Then C = C + 255
                            'Draw Monster
                            If .A > 0 Then
                                b = .D * 3 + 2
                            Else
                                b = .D * 3 + .W
                            End If
                            If map.Tile(.x, .y).Att = 20 Then
                                VCO = map.Tile(.x, .y).AttData(0)
                            ElseIf map.Tile(.x, .y).Att = 26 Then
                                VCO = map.Tile(.x, .y).AttData(3)
                            Else
                                VCO = 0
                            End If
                            If CurrentTarget.TargetType = TT_MONSTER Then
                                If CurrentTarget.Target = A Then
                                    'DrawSpriteGlow .XO, .YO - 16, C, B, D3DColorARGB(TargetPulse, 255, 0, 0), VCO
                                    Monster(.Monster).Flags2 = Monster(.Monster).Flags2
                                    
                                    DrawSpriteGlow .XO, .YO - 16, C, b, D3DColorARGB(TargetPulse * (Monster(.Monster).alpha - .alpha) / 255, 255, 0, 0), VCO
                                    
                                End If
                            End If
                            DrawSprite .XO, .YO - 16, C, b, D3DColorARGB(Monster(.Monster).alpha - .alpha, Monster(.Monster).Red - .R, Monster(.Monster).Green - .G, Monster(.Monster).Blue) - .b, IIf(Monster(.Monster).Flags And MONSTER_NO_SHADOW, False, True), VCO
                            'draw hp bar
                            If A = TargetMonster Then
                                If (Monster(.Monster).alpha - .alpha) >= 100 Then
                                    DrawRect .XO + 30, .XO + 32, .YO - 16, .YO + 16, BS_SOLID, D3DColorARGB(255, 0, 0, 0), 0, 0
                                    If Monster(.Monster).MaxHP > 0 Then
                                        b = (32 * CLng(.HP) \ Monster(.Monster).MaxHP)
                                        b = 32 - b
                                        If b < 0 Then b = 0
                                        If b > 32 Then b = 32
                                        DrawRect .XO + 30, .XO + 32, .YO - 16 + b, .YO + 16, BS_SOLID, D3DColorARGB(255, 255, 0, 0), 0, 0
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next A

    For A = 1 To MAXUSERS
        With player(A)
            If .map = CMap And (.Status <> 9 Or Character.Access >= 9) Then
                If GetStatusEffect(A, SE_INVISIBLE) = 0 Or (Character.SkillLevels(SKILL_PERCEPTION) > 0 And Abs(cX - player(A).x) < 4 And Abs(cY - player(A).y) < 4) Or (player(A).Guild = Character.Guild And Character.Guild > 0) Or Character.Access >= 9 Then
                If GetStatusEffect(A, SE_INVISIBLE) = 1 Or .Status = 9 Then
                    D = (255 - .alpha) / 3
                Else
                    D = 255 - .alpha
                End If
                    
                
                'Draw Player
                If .A > 0 Then
                    b = .D * 3 + 2
                Else
                    b = .D * 3 + .W
                End If
                If map.Tile(.x, .y).Att = 20 Then
                    VCO = map.Tile(.x, .y).AttData(0)
                ElseIf map.Tile(.x, .y).Att = 26 Then
                    VCO = map.Tile(.x, .y).AttData(3)
                Else
                    VCO = 0
                End If
                If CurrentTarget.TargetType = TT_PLAYER Then
                    If CurrentTarget.Target = A Then
                        DrawSpriteGlow .XO, .YO - 16, .Sprite, b, D3DColorARGB(TargetPulse * (255 - .alpha) / 255, 255, 255, 0), VCO
                    End If
                End If
                If (.Sprite) > 0 And (.Sprite) < 255 Then
                    If GetStatusEffect(A, SE_BERSERK) Then
                        C = D3DColorARGB(D, 255 - .Red, 255 - 150, 255 - 150)
                    ElseIf GetStatusEffect(A, SE_MASTERDEFENSE) Then
                        C = D3DColorARGB(D, 255 - 150, 255 - 150, 255 - .Blue)
                    ElseIf GetStatusEffect(A, SE_INVULNERABILITY) Then
                        C = D3DColorARGB(D, 255, 255, 0)
                    ElseIf GetStatusEffect(A, SE_ETHEREALITY) Then
                        D = D / 2
                        C = D3DColorARGB(D, 255 - .Red, 255 - .Green, 255 - .Blue)
                    Else
                        C = D3DColorARGB(D, 255 - .Red, 255 - .Green, 255 - .Blue)
                    End If
                    
                    
                    
                        If .equippedPicture(1) > 0 Then
                            If Object(.equippedPicture(1)).EquipmentPicture > 0 Then DrawEquipmentShadow .XO, .YO - 16, Object(.equippedPicture(1)).EquipmentPicture, b, VCO, 1
                        End If
                        If .equippedPicture(2) > 0 Then
                            If Object(.equippedPicture(2)).EquipmentPicture > 0 Then
                                If ExamineBit(Object(.equippedPicture(2)).Flags, 5) Then
                                    DrawEquipmentShadowAlt .XO, .YO - 16, Object(.equippedPicture(2)).EquipmentPicture, b, VCO, 2
                                Else
                                    DrawEquipmentShadow .XO, .YO - 14, Object(.equippedPicture(2)).EquipmentPicture, b, VCO, 2
                                End If
                            End If
                        End If
                        
                        If .equippedPicture(3) > 0 Then
                            If Object(.equippedPicture(3)).EquipmentPicture > 0 Then DrawEquipmentShadow .XO, .YO - 14, Object(.equippedPicture(3)).EquipmentPicture, b, VCO, 3
                        End If
                    
                    
                        
                        If b <= 2 Or (b >= 6 And b <= 8) Then 'draw weap behind player
                            If .equippedPicture(1) > 0 Then
                                If Object(.equippedPicture(1)).EquipmentPicture > 0 Then DrawEquipment .XO, .YO - 16, Object(.equippedPicture(1)).EquipmentPicture, b, C, VCO, 1
                            End If
                        End If
                        

                        If .equippedPicture(2) > 0 Then
                            If Object(.equippedPicture(2)).EquipmentPicture > 0 Then
                                If Not ExamineBit(Object(.equippedPicture(2)).Flags, 5) Then
                                    If (b >= 9 And b <= 11) Then
                                        DrawEquipment .XO, .YO - 14, Object(.equippedPicture(2)).EquipmentPicture, b, C, VCO, 2
                                    End If
                                Else
                                    If b <= 2 Or (b >= 9 And b <= 11) Then
                                        DrawEquipmentAlt .XO, .YO - 16, Object(.equippedPicture(2)).EquipmentPicture, b, C, VCO, 2
                                    End If
                                End If
                            End If
                        End If
                        
                        
                        DrawSprite .XO, .YO - 16, .Sprite, b, C, True, VCO
                        If .equippedPicture(3) > 0 Then
                            If Object(.equippedPicture(3)).EquipmentPicture > 0 Then DrawEquipment .XO, .YO - 14, Object(.equippedPicture(3)).EquipmentPicture, b, C, VCO, 3
                        End If
                        If Options.dontDisplayHelms = False Then
                            If .equippedPicture(4) > 0 Then
                                If Object(.equippedPicture(4)).EquipmentPicture > 0 Then DrawEquipment .XO, .YO - 16 - 8, Object(.equippedPicture(4)).EquipmentPicture, b, C, VCO, 3
                            End If
                        End If
                        If (b > 2 And b < 6) Or b > 8 Then
                            If .equippedPicture(1) > 0 Then
                                If Object(.equippedPicture(1)).EquipmentPicture > 0 Then DrawEquipment .XO, .YO - 16, Object(.equippedPicture(1)).EquipmentPicture, b, C, VCO, 1
                            End If
                        End If
                        

                        If .equippedPicture(2) > 0 Then
                            If Object(.equippedPicture(2)).EquipmentPicture > 0 Then
                                If Not ExamineBit(Object(.equippedPicture(2)).Flags, 5) Then
                                    If Not (b >= 9 And b <= 11) Then
                                        DrawEquipment .XO, .YO - 14, Object(.equippedPicture(2)).EquipmentPicture, b, C, VCO, 2
                                    End If
                                Else
                                    If Not (b <= 2 Or (b >= 9 And b <= 11)) Then
                                        DrawEquipmentAlt .XO, .YO - 16, Object(.equippedPicture(2)).EquipmentPicture, b, C, VCO, 2
                                    End If
                                End If
                            End If
                        End If

                        If .Guild > 0 Then DrawSymbol .XO - 2, .YO - 4, .Guild, VCO
         
                    
                      
                    If .StatusEffect Then
                        If GetStatusEffect(A, SE_POISON) Then
                            DrawSpriteGlow .XO, .YO - 16, .Sprite, b, D3DColorARGB((TargetPulse \ 3) * D \ 255, 0, 127, 0), VCO
                        End If
                        If GetStatusEffect(A, SE_CRIPPLE) Then
                            DrawSpriteGlow .XO, .YO - 16, .Sprite, b, D3DColorARGB((TargetPulse \ 3) * D \ 255, 200, 200, 0), VCO
                        End If
                        
                        If GetStatusEffect(A, SE_MUTE) Then DrawEffect .XO, .YO - 16, 64, FrameCounter
                        If GetStatusEffect(A, SE_DEADLYCLARITY) Then DrawEffect .XO, .YO - 16, 85, FrameCounter
                        If GetStatusEffect(A, SE_REGENERATION) Then DrawEffect .XO, .YO - 16, 41, FrameCounter
                        If GetStatusEffect(A, SE_EXHAUST) Then DrawEffect .XO, .YO - 16, 87, FrameCounter
                        If GetStatusEffect(A, SE_SHATTERSHIELD) Then DrawEffect .XO, .YO - 16, 96, 0
                        If GetStatusEffect(A, SE_EVANESCENCE) Then DrawEffect .XO, .YO - 16, 12, FrameCounter
                        If GetStatusEffect(A, SE_FIERYESSENCE) Then DrawEffect .XO, .YO - 16, 53, FrameCounter
                        
                        If GetStatusEffect(A, SE_RETRIBUTION) Then DrawEffect .XO, .YO - 16, 75, FrameCounter
                    End If
                    Select Case .buff
                        Case BUFF_EMPOWER:    DrawEffect .XO, .YO - 16, 95, 0
                        Case BUFF_WILLPOWER:  DrawEffect .XO, .YO - 16, 95, 1
                        Case BUFF_ZEAL:       DrawEffect .XO, .YO - 16, 95, 2
                        Case BUFF_VITALITY:   DrawEffect .XO, .YO - 16, 95, 3
                        Case BUFF_HOLYARMOR:  DrawEffect .XO, .YO - 16, 95, 4
                        Case BUFF_NECROMANCY: DrawEffect .XO, .YO - 16, 95, 5
                        Case BUFF_EVOCATION:  DrawEffect .XO, .YO - 16, 95, 6
                        Case BUFF_MANASHIELD: DrawEffect .XO, .YO - 16, 96, 1
                    End Select
                End If
                If (.Guild = Character.Guild And .Guild > 0) Or Character.Access > 7 Or (.Party = Character.Party And .Party > 0) Then
                    If .HP <> 100 Or .Mana <> 100 Then
                        DrawRect .XO + 30, .XO + 32, .YO - 16, .YO + 16, BS_SOLID, &HFF000000, 0, 0
                        b = (32 * .HP) \ 100
                        b = 32 - b
                        If b < 0 Then b = 0
                        If b > 32 Then b = 32
                        DrawRect .XO + 30, .XO + 32, .YO - 16 + b, .YO + 16, BS_SOLID, &HFFFF0000, 0, 0
                        
                        DrawRect .XO + 32, .XO + 34, .YO - 16, .YO + 16, BS_SOLID, &HFF000000, 0, 0
                        b = (32 * .HP) \ 100
                        b = 32 - b
                        If b < 0 Then b = 0
                        If b > 32 Then b = 32
                        DrawRect .XO + 32, .XO + 34, .YO - 16 + b, .YO + 16, BS_SOLID, StatusColors(3), 0, 0
                    End If
                End If
            End If
            End If
        End With
    Next A

    'Draw You
    If CAttack > 0 Then
        b = CDir * 3 + 2
    Else
        b = CDir * 3 + CWalk
    End If
    If map.Tile(cX, cY).Att = 20 Then
        VCO = map.Tile(cX, cY).AttData(0)
    ElseIf map.Tile(cX, cY).Att = 26 Then
        VCO = map.Tile(cX, cY).AttData(3)
    Else
        VCO = 0
    End If
    If CurrentTarget.TargetType = TT_CHARACTER Then
        DrawSpriteGlow Cxo, CYO - 16, Character.Sprite, b, D3DColorARGB(TargetPulse * (255 - Character.alpha) / 255, 0, 255, 0), VCO
    End If
    
    If Character.Equipped(1).Object > 0 Then
        If Object(Character.Equipped(1).Object).EquipmentPicture > 0 Then DrawEquipmentShadow Cxo, CYO - 16, Object(Character.Equipped(1).Object).EquipmentPicture, b, VCO, 1
    End If
    If Character.Equipped(2).Object > 0 Then
        If Object(Character.Equipped(2).Object).EquipmentPicture > 0 Then
            If ExamineBit(Object(Character.Equipped(2).Object).Flags, 5) Then
                DrawEquipmentShadowAlt Cxo, CYO - 16, Object(Character.Equipped(2).Object).EquipmentPicture, b, VCO, 2
            Else
                DrawEquipmentShadow Cxo, CYO - 14, Object(Character.Equipped(2).Object).EquipmentPicture, b, VCO, 2
            End If
        End If
    End If
    If Character.Equipped(3).Object > 0 Then
        If Object(Character.Equipped(3).Object).EquipmentPicture > 0 Then DrawEquipmentShadow Cxo, CYO - 14, Object(Character.Equipped(3).Object).EquipmentPicture, b, VCO, 3
    End If

    
    If GetStatusEffect(Character.Index, SE_INVISIBLE) Then
        C = D3DColorARGB(127, 255 - Character.Red, 255 - Character.Green, 255 - Character.Blue)
    Else
        If GetStatusEffect(Character.Index, SE_BERSERK) Then
            C = D3DColorARGB(255 - Character.alpha, 255 - Character.Red, 255 - 150, 255 - 150)
        Else
            If GetStatusEffect(Character.Index, SE_MASTERDEFENSE) Then
                C = D3DColorARGB(255 - Character.alpha, 255 - 160, 255 - 160, 255 - Character.Blue)
            Else
                If GetStatusEffect(Character.Index, SE_INVULNERABILITY) Then
                    C = D3DColorARGB(255 - Character.alpha, 255, 255, 0)
                Else
                    If GetStatusEffect(Character.Index, SE_ETHEREALITY) Then
                        C = D3DColorARGB(255 / 2 - Character.alpha / 2, 255 - Character.Red, 255 - Character.Green, 255 - Character.Blue)
                    Else
                        C = D3DColorARGB(255 - Character.alpha, 255 - Character.Red, 255 - Character.Green, 255 - Character.Blue)
                    End If
                End If
            End If
        End If
    End If
    
    If b <= 2 Or (b >= 6 And b <= 8) Then 'draw weap behind player
        If Character.Equipped(1).Object > 0 Then
            If Object(Character.Equipped(1).Object).EquipmentPicture > 0 Then DrawEquipment Cxo, CYO - 16, Object(Character.Equipped(1).Object).EquipmentPicture, b, C, VCO, 1
        End If
    End If
    
    
    If Character.Equipped(2).Object > 0 Then
        If Object(Character.Equipped(2).Object).EquipmentPicture > 0 Then
            If Not ExamineBit(Object(Character.Equipped(2).Object).Flags, 5) Then
                If (b >= 9 And b <= 11) Then
                    DrawEquipment Cxo, CYO - 14, Object(Character.Equipped(2).Object).EquipmentPicture, b, C, VCO, 2
                End If
            Else
                If b <= 2 Or (b >= 9 And b <= 11) Then
                    DrawEquipmentAlt Cxo, CYO - 16, Object(Character.Equipped(2).Object).EquipmentPicture, b, C, VCO, 2
                End If
            End If
        End If
    End If
    
    DrawSprite Cxo, CYO - 16, Character.Sprite, b, C, True, VCO
    If Character.Equipped(3).Object > 0 Then
        If Object(Character.Equipped(3).Object).EquipmentPicture > 0 Then DrawEquipment Cxo, CYO - 14, Object(Character.Equipped(3).Object).EquipmentPicture, b, C, VCO, 3
    End If
    If Options.dontDisplayHelms = False Then
        If Character.Equipped(4).Object > 0 Then
            If Object(Character.Equipped(4).Object).EquipmentPicture > 0 Then DrawEquipment Cxo, CYO - 16 - 8, Object(Character.Equipped(4).Object).EquipmentPicture, b, C, VCO, 3
        End If
    End If
    If (b > 2 And b < 6) Or b > 8 Then
        If Character.Equipped(1).Object > 0 Then
            If Object(Character.Equipped(1).Object).EquipmentPicture > 0 Then DrawEquipment Cxo, CYO - 16, Object(Character.Equipped(1).Object).EquipmentPicture, b, C, VCO, 1
        End If
    End If
    
    If Character.Equipped(2).Object > 0 Then
        If Object(Character.Equipped(2).Object).EquipmentPicture > 0 Then
            If Not ExamineBit(Object(Character.Equipped(2).Object).Flags, 5) Then
                If Not (b >= 9 And b <= 11) Then
                    DrawEquipment Cxo, CYO - 14, Object(Character.Equipped(2).Object).EquipmentPicture, b, C, VCO, 2
                End If
            Else
                If Not (b <= 2 Or (b >= 9 And b <= 11)) Then
                    DrawEquipmentAlt Cxo, CYO - 16, Object(Character.Equipped(2).Object).EquipmentPicture, b, C, VCO, 2
                End If
            End If
        End If
    End If

    
    If Character.Guild > 0 Then DrawSymbol Cxo - 2, CYO - 4, Character.Guild, VCO
         
    If Character.StatusEffect Then
        If GetStatusEffect(Character.Index, SE_POISON) Then DrawSpriteGlow Cxo, CYO - 16, Character.Sprite, b, D3DColorARGB(TargetPulse / 3, 0, 127, 0), VCO
        If GetStatusEffect(Character.Index, SE_CRIPPLE) Then DrawSpriteGlow Cxo, CYO - 16, Character.Sprite, b, D3DColorARGB(TargetPulse / 3, 200, 200, 0), VCO
        
        If GetStatusEffect(Character.Index, SE_MUTE) Then DrawEffect Cxo, CYO - 16, 64, FrameCounter
        If GetStatusEffect(Character.Index, SE_DEADLYCLARITY) Then DrawEffect Cxo, CYO - 16, 85, FrameCounter
        If GetStatusEffect(Character.Index, SE_REGENERATION) Then DrawEffect Cxo, CYO - 16, 41, FrameCounter
        If GetStatusEffect(Character.Index, SE_EXHAUST) Then DrawEffect Cxo, CYO - 16, 87, FrameCounter
        'If GetStatusEffect(Character.Index, SE_BERSERK) Then DrawEffect Cxo, CYO - 16, 77, FrameCounter
        If GetStatusEffect(Character.Index, SE_SHATTERSHIELD) Then DrawEffect Cxo, CYO - 16, 96, 0
        If GetStatusEffect(Character.Index, SE_EVANESCENCE) Then DrawEffect Cxo, CYO - 16, 12, FrameCounter
        If GetStatusEffect(Character.Index, SE_FIERYESSENCE) Then DrawEffect Cxo, CYO - 16, 53, FrameCounter
        If GetStatusEffect(Character.Index, SE_RETRIBUTION) Then DrawEffect Cxo, CYO - 16, 75, FrameCounter
    End If
    If Options.ShowHP Then
        If Character.HP < Character.MaxHP Then
            DrawRect Cxo + 30, Cxo + 32, CYO - 16, CYO + 16, BS_SOLID, D3DColorARGB(255, 0, 0, 0), 0, 0
            b = (32 * CLng(Character.HP) \ Character.MaxHP)
            b = 32 - b
            If b < 0 Then b = 0
            If b > 32 Then b = 32
            DrawRect Cxo + 30, Cxo + 32, CYO - 16 + b, CYO + 16, BS_SOLID, D3DColorARGB(255, 255, 0, 0), 0, 0
        End If
    End If
    Select Case Character.buff
        Case BUFF_EMPOWER:    DrawEffect Cxo, CYO - 16, 95, 0
        Case BUFF_WILLPOWER:  DrawEffect Cxo, CYO - 16, 95, 1
        Case BUFF_ZEAL:       DrawEffect Cxo, CYO - 16, 95, 2
        Case BUFF_VITALITY:   DrawEffect Cxo, CYO - 16, 95, 3
        Case BUFF_HOLYARMOR:  DrawEffect Cxo, CYO - 16, 95, 4
        Case BUFF_NECROMANCY: DrawEffect Cxo, CYO - 16, 95, 5
        Case BUFF_EVOCATION:  DrawEffect Cxo, CYO - 16, 95, 6
        Case BUFF_MANASHIELD: DrawEffect Cxo, CYO - 16, 96, 1
    End Select
    
    For Each tmpPS In ParticleEngineF
        With tmpPS
            .Render
        End With
    Next
    
        For A = 0 To 9
        With map.Monster(A)
            If .Monster > 0 Then
                If (Monster(.Monster).Flags2 And MONSTER_LARGE) Or (Monster(.Monster).Flags2 And MONSTER_MEDIUM) Then
                    If .A > 0 Then
                        b = .D * 3 + 2
                    Else
                        b = .D * 3 + .W
                    End If
                    If map.Tile(.x, .y).Att = 20 Then
                        VCO = map.Tile(.x, .y).AttData(0)
                    ElseIf map.Tile(.x, .y).Att = 26 Then
                        VCO = map.Tile(.x, .y).AttData(3)
                    Else
                        VCO = 0
                    End If
                    C = Monster(.Monster).Sprite
                    If C > 0 And C < 255 Then
                        If (Monster(.Monster).Flags2 And MONSTER_LARGE) Then
                            DrawLargeSpriteTop .XO, .YO - 16, C, b, D3DColorARGB(Monster(.Monster).alpha - .alpha, Monster(.Monster).Red - .R, Monster(.Monster).Green - .G, Monster(.Monster).Blue) - .b, VCO
                        Else
                            DrawLargeSpriteTop .XO - 16, .YO - 48, C, b, D3DColorARGB(Monster(.Monster).alpha - .alpha, Monster(.Monster).Red - .R, Monster(.Monster).Green - .G, Monster(.Monster).Blue) - .b, VCO
                        End If
                    End If
                End If
            End If
        End With
        Next A
    
    
    Dim tmpEffect As clsEffect
    For Each tmpEffect In Effects
        With tmpEffect
            If .Sprite > 0 Then
                ''Draw BackBuffer, .X, .Y - 16, 32, 32 + VCO, sfcEffects, .Frame * 32, (.Sprite - 1) * 32, True, 0
                DrawEffect .x, .y - 16, .Sprite, .Frame
                
            End If
        End With
    Next

    Dim PJ As clsProjectile
    For Each PJ In Projectiles
        With PJ
            If .Sprite < 256 Then 'Normal Effects.rsc
                DrawEffect .XO, .YO + .YOffset, .Sprite, 0
            Else
                DrawProjectile .XO, .YO + .YOffset, .Sprite - 255, .D, .Frame \ 2
            End If
        End With
    Next

    If Not Options.fullredraws Then
        D3DDevice.SetRenderTarget mapTexture(2).GetSurfaceLevel(0), Nothing, 0
        If MapEdit Then D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, ByVal &H0&, 1#, 0
    End If
    If Options.fullredraws Or mapChanged Or MapEdit Or mapFGChanged Then
        mapChanged = False
        mapFGChanged = False
        For x = 0 To 11
            For y = 0 To 11
                X32 = x * 32
                Y32 = y * 32
                'Now Draw Upper Portion (Above Players)
                If MapEdit = False Then
                        With map.Tile(x, y)
                            If .FGTile > 0 Then
                                If .Att = 18 Then
                                    If .AttData(0) = 0 And .AttData(1) = 0 And .AttData(2) = 0 And .AttData(3) = 0 Then
                                        'Draw sfcFront, X * 32, Y * 32, 32, 16, sfcTiles, ((.FGTile - 1) Mod 7) * 32, Int((.FGTile - 1) / 7) * 32, True, 0
                                        DrawTile X32, Y32, .FGTile, 0, 0, 16, 32
                                    Else
                                        If (.AttData(0) And 1) = 0 Then
                                            DrawTile X32, Y32, .FGTile, 16, 0, 16, 16
                                        End If
                                        If (.AttData(0) And 2) = 0 Then
                                            DrawTile X32, Y32, .FGTile, 0, 0, 16, 16
                                        End If
                                        If (.AttData(1) And 1) = 0 Then
                                            DrawTile X32, Y32, .FGTile, 0, 16, 16, 16
                                        End If
                                        If (.AttData(1) And 2) = 0 Then
                                            DrawTile X32, Y32, .FGTile, 16, 16, 16, 16
                                        End If
                                    End If
                                ElseIf .Att = 4 Then
                                    DrawTile X32, Y32, .FGTile, 0, 0, 16, 32
                                Else
                                    If .Att = 23 Then
                                        If ExamineBit(.AttData(2), 4) Then
                                            If ExamineBit(.AttData(3), 0) Then XOffset = -.AttData(0) Else XOffset = .AttData(0)
                                            If ExamineBit(.AttData(3), 1) Then YOffset = -.AttData(1) Else YOffset = .AttData(1)
                                        Else
                                            XOffset = 0: YOffset = 0
                                        End If
                                    Else
                                        XOffset = 0: YOffset = 0
                                    End If
                                    A = 0
                                    
                                        If y + 1 <= 11 Then
                                            If map.Tile(x, y + 1).Att = 24 And map.Tile(x, y + 1).AttData(3) = 0 And ExamineBit(map.Tile(x, y + 1).AttData(2), 1) Then
                                                A = 1
                                            End If
                                        End If
                                        If y + 2 <= 11 Then
                                            If map.Tile(x, y + 2).Att = 24 And map.Tile(x, y + 2).AttData(3) = 0 And ExamineBit(map.Tile(x, y + 2).AttData(2), 2) Then
                                                A = 1
                                            End If
                                        End If
                                        If y + 2 <= 11 And x + 1 <= 11 Then
                                            If map.Tile(x + 1, y + 2).Att = 24 And map.Tile(x + 1, y + 2).AttData(3) = 0 And ExamineBit(map.Tile(x + 1, y + 2).AttData(2), 5) Then
                                                A = 1
                                            End If
                                        End If
                                        If y + 1 <= 11 And x + 1 <= 11 Then
                                            If map.Tile(x + 1, y + 1).Att = 24 And map.Tile(x + 1, y + 1).AttData(3) = 0 And ExamineBit(map.Tile(x + 1, y + 1).AttData(2), 5) Then
                                                A = 1
                                            End If
                                        End If
                                        If y + 2 <= 11 And x - 1 >= 0 Then
                                            If map.Tile(x - 1, y + 2).Att = 24 And map.Tile(x - 1, y + 2).AttData(3) = 0 And ExamineBit(map.Tile(x - 1, y + 2).AttData(2), 6) Then
                                                A = 1
                                            End If
                                        End If
                                        If y + 1 <= 11 And x - 1 >= 0 Then
                                            If map.Tile(x - 1, y + 1).Att = 24 And map.Tile(x - 1, y + 1).AttData(3) = 0 And ExamineBit(map.Tile(x - 1, y + 1).AttData(2), 6) Then
                                                A = 1
                                            End If
                                        End If
                                        If x - 1 >= 0 Then
                                            If map.Tile(x - 1, y).Att = 24 And map.Tile(x - 1, y).AttData(3) = 0 And ExamineBit(map.Tile(x - 1, y).AttData(2), 7) And ExamineBit(map.Tile(x - 1, y).AttData(2), 5) Then
                                                A = 1
                                            End If
                                        End If
                                        If x >= 0 Then
                                            If map.Tile(x, y).Att = 24 And map.Tile(x, y).AttData(3) = 0 And ExamineBit(map.Tile(x, y).AttData(2), 7) And ExamineBit(map.Tile(x, y).AttData(2), 5) Then
                                                A = 1
                                            End If
                                        End If
                                        If x + 1 <= 11 Then
                                            If map.Tile(x + 1, y).Att = 24 And map.Tile(x + 1, y).AttData(3) = 0 And ExamineBit(map.Tile(x + 1, y).AttData(2), 7) And ExamineBit(map.Tile(x + 1, y).AttData(2), 6) Then
                                                A = 1
                                            End If
                                        End If
                                    
                                    If A = 0 Then DrawTile X32 + XOffset, Y32 + YOffset, map.Tile(x, y).FGTile
                                
                                End If
                            End If
                        End With
                Else
                    With EditMap.Tile(x, y)
                        If .FGTile > 0 Then
                            DrawTile X32, Y32, .FGTile
                        End If
                        If EditMode = 2 Then
                            DrawRect X32, X32 + 32, Y32, Y32 + 32, BS_NONE, 0, 1, &HFFFFFFFF
                            If keyAlt = True Then
                                DrawBmpString3D Create3DString((.Anim(2) \ 4 And 15)), X32 + 6, Y32, &HFFFFFFFF
                                DrawBmpString3D Create3DString(Format((.Anim(2) \ 128) And 1, "0") & Format((.Anim(2) \ 64) And 1, "0")), X32 + 24, Y32, &HFFFFFF00
                                DrawBmpString3D Create3DString("+"), X32 + 6, Y32 + 16, &HFFFFFFFF
                                DrawBmpString3D Create3DString("-"), X32 + 22, Y32 + 16, &HFFFFFFFF
                            ElseIf ((.Anim(1) \ 16) > 0) Then
                                DrawBmpString3D Create3DString((.Anim(1) \ 16) + 1), X32 + 6, Y32, &HFFFFFFFF
                                DrawBmpString3D Create3DString((.Anim(1) And 15) * 4), X32 + 10, Y32 + 16, &HFFFFFFFF
                            End If
                        End If
                        If EditMode = 4 Then
                            If .Att > 0 Then
                                DrawAtts X32, Y32, .Att, 0
                            End If
                        End If
                        If EditMode = 7 Then
                            If .WallTile > 0 Then
                                For A = 0 To 7
                                    If ExamineBit(.WallTile, A) Then
                                        DrawAtts X32, Y32, A + 2, 1
                                    End If
                                Next A
                            End If
                        End If
                    End With
                End If
            Next y
        Next x
    End If
    
    If Not Options.fullredraws Then
        D3DDevice.SetRenderTarget RenderSurface(0), Nothing, ByVal 0
        D3DDevice.SetTexture 0, mapTexture(2)
        Draw3D 0, 0, 384, 384, 0, 0, 0, -1, 512, 512
        LastTexture = 0
    End If
    
    DrawLightMap

    A = Cxo + 16
    If Character.Guild > 0 Then
        If Character.Status = 1 And CurFrame = 0 Then
            DrawBmpString3D Create3DString(Character.Name), A, CYO - 34, &HFF800000
        Else
            DrawBmpString3D Create3DString(Character.Name), A, CYO - 34, &HFF00FFFF
        End If
    Else
        If GetStatusEffect(Character.Index, SE_INVISIBLE) Then
            DrawBmpString3D Create3DString(Character.Name), A, CYO - 34, &H7F7F7F7F
        Else
            If Character.Status = 2 Then
                DrawBmpString3D Create3DString(Character.Name), A, CYO - 34, &HFFFFFF00
            ElseIf Character.Status = 3 Then
                DrawBmpString3D Create3DString(Character.Name), A, CYO - 34, &HFF0000FF
            ElseIf Character.Status = 1 And CurFrame = 0 Then
                DrawBmpString3D Create3DString(Character.Name), A, CYO - 34, &HFF800000
            Else
                If Character.Status = 21 Then 'Rainbow
                    DrawBmpString3D Create3DString(Character.Name), A, CYO - 34, StatusColors((Int(Rnd * 20)))
                Else
                    If Character.Status = 25 And CurFrame = 0 Then
                        DrawBmpString3D Create3DString(Character.Name), A, CYO - 34, StatusColors(10)
                    ElseIf Character.Status = 25 And CurFrame = 1 Then
                        DrawBmpString3D Create3DString(Character.Name), A, CYO - 34, StatusColors(10)
                    Else
                        DrawBmpString3D Create3DString(Character.Name), A, CYO - 34, StatusColors(Character.Status)
                    End If
                End If
            End If
        End If
    End If
    For A = 1 To MAXUSERS
        With player(A)
            If .map = CMap Then
                If (.Status <> 9 And GetStatusEffect(A, SE_INVISIBLE) = 0) Or (.Guild = Character.Guild And Character.Guild > 0) Or (Character.Access >= 9) Then  'Invisible Name
                    If .Guild > 0 And CurFrame = 1 Then
                        b = .Color
                    Else
                        If .Status = 1 And CurFrame = 0 Then
                            b = &HFF800000
                        ElseIf .Status = 1 And CurFrame = 1 Then
                            b = .Color
                        ElseIf .Status = 21 Then  'Rainbow
                            b = StatusColors(Int(Rnd * 20))
                        Else
                            If .Status = 25 And CurFrame = 0 Then
                                b = StatusColors(10)
                            ElseIf .Status = 25 And CurFrame = 1 Then
                                b = StatusColors(16)
                            Else
                                If .Status = 2 Then
                                    b = &HFFFFFF00
                                Else
                                    b = StatusColors(.Status)
                                End If
                            End If
                        End If
                    End If
                    
                  '  DrawBmpString3D Create3DString(.Name), .XO + 16, .YO - 34, .Color, True, 1, B
                    
                    
                    DrawBmpString3D Create3DString(.Name), .XO + 16, .YO - 34, b
                End If
            End If
        End With
    Next A
    If Options.MName Then
        For A = 0 To 9
            With map.Monster(A)
                If .Monster > 0 Then
                    C = Monster(.Monster).Sprite
                    If C > 0 Then
                        If (Monster(.Monster).Flags And MONSTER_FRIENDLY) Then
                            b = &HFF00FFFF
                        ElseIf (Monster(.Monster).Flags And MONSTER_GUARD) Then
                            b = &HFFFFFFFF
                        Else
                            b = StatusColors(17) '&HFFC00000
                        End If
                        If (Monster(.Monster).Flags2 And MONSTER_LARGE) Then
                            If Monster(.Monster).alpha - .alpha > 50 Then DrawBmpString3D Create3DString(Monster(.Monster).Name), .XO + 32, .YO - 32, b
                        ElseIf Monster(.Monster).Flags2 And MONSTER_MEDIUM Then
                            If Monster(.Monster).alpha - .alpha > 50 Then DrawBmpString3D Create3DString(Monster(.Monster).Name), .XO + 16, .YO - 64, b
                        Else
                            If Monster(.Monster).alpha - .alpha > 50 Then DrawBmpString3D Create3DString(Monster(.Monster).Name), .XO + 16, .YO - 32, b
                        End If
                    End If
                    
                End If
            End With
        Next A
    End If
    Dim ft As clsFloatText
    For Each ft In FloatingText
        With ft
            If .Big = False Then
                If .Life > 0 Then
                    DrawBmpString3D Create3DString(.Text), .x, .y - 6, .Color
                Else
                    DrawBmpString3D Create3DString(.Text), .x + 16, .y - 16 - (.Step * 1), .Color
                End If
            Else
                If .Life > 0 Then
                    DrawBmpString3D Create3DString(.Text), .x, .y - 6, .Color, True, .Mult
                Else
                    DrawBmpString3D Create3DString(.Text), .x + 16, .y - 16 - (.Step * 1), .Color, True, .Mult
                End If
            End If
        End With
    Next

    If (map.Raining > 0 And (map.Flags(1) And MAP_RAINING)) Then
        UpdateRain False
        DrawRain
    ElseIf (World.Rain > 0 Or NumRainDrops > 0) Then
        UpdateRain True
        DrawRain
    End If
    If (map.Snowing > 0 And (map.Flags(1) And MAP_SNOWING)) Then
        UpdateSnow False
        DrawSnow
    ElseIf (World.Snow > 0 Or NumSnowFlakes > 0) Then
        UpdateSnow True
        DrawSnow
    End If
    
    If (map.Fog > 0 And map.Fog <= 31) Or World.Fog Or CurFog Then
        If ExamineBit(map.Flags(0), 1) Or Options.ShowFog Then
            UpdateFog
            DrawFog
        End If
    End If
    



    DrawUnzStrings
    
End Sub
