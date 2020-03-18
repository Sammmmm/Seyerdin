Attribute VB_Name = "modDirectx"
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

Public Type SurfaceType
    Surface As DirectDrawSurface7
    desc As DDSURFACEDESC2
End Type

Public Dx7 As New DirectX7
Public DD As DirectDraw7
Public sfcTiles As SurfaceType
Public sfcBuffer As SurfaceType
Public sfcShrinkBuffer As SurfaceType


Public Function Init_DX()
    Set DD = Dx7.DirectDrawCreate("")
    Call DD.SetCooperativeLevel(frmMain.hWnd, DDSCL_NORMAL)

    sfcBuffer.desc.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    sfcBuffer.desc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    sfcBuffer.desc.lWidth = 384
    sfcBuffer.desc.lHeight = 384
    Set sfcBuffer.Surface = DD.CreateSurface(sfcBuffer.desc)
    
    sfcShrinkBuffer.desc.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    sfcShrinkBuffer.desc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    sfcShrinkBuffer.desc.lWidth = 192
    sfcShrinkBuffer.desc.lHeight = 192
    Set sfcShrinkBuffer.Surface = DD.CreateSurface(sfcShrinkBuffer.desc)

    sfcTiles.desc.lFlags = DDSD_CAPS
    sfcTiles.desc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set sfcTiles.Surface = DD.CreateSurfaceFromFile(App.Path & "\tiles.rsc", sfcTiles.desc)
    
    Dim CK As DDCOLORKEY
    CK.low = 0
    CK.high = 0
    sfcTiles.Surface.SetColorKey DDCKEY_SRCBLT, CK
    sfcBuffer.Surface.SetColorKey DDCKEY_SRCBLT, CK
End Function
