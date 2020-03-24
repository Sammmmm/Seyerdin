VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
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

Public Key As String
Public WidgetType As Long
Public Style As Long
Public x As Long
Public y As Long
Public Width As Long
Public Height As Long
Public Visible As Boolean
Public Caption As String
Public Highlighted As Boolean
Public Selected As Boolean
Public Data0 As Long
Public Data1 As Long
Public Remote As Boolean
Private wTexture As TextureType
Private ImageLoaded As Boolean
Public BackColor As Long
Public ForeColor As Long

Public Children As New clsWidgets


Public Function GetWidgetRect() As RECT
    Dim R1 As RECT
    R1.Left = x
    R1.Top = y
    R1.Right = x + Width
    R1.Bottom = y + Height
    GetWidgetRect = R1
End Function

Public Sub InitTexture()
    Dim CK As DDCOLORKEY
    CK.high = 0
    CK.low = 0
    
    Select Case Caption
        Case "sprites"
            
        Case "objects"
            
        Case "effects"
        
        Case "projectiles"
        
        Case "tiles"
        
        Case "tilesex"
        
        Case Else
            If Exists("Data\Graphics\" & Caption & ".rsc") Then
                Set wTexture.Texture = D3DX.CreateTextureFromFileEx(D3DDevice, "Data\Graphics\" & Caption & ".rsc", Width, Height, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, wTexture.TexInfo, ByVal 0)
                ImageLoaded = True
            Else
                ImageLoaded = False
            End If
    End Select
End Sub

Public Sub DrawImage()
On Error Resume Next
    Select Case LCase(Caption)
        Case "sprites"
            ''Draw BackBuffer, X, Y, Width, Height, sfcSprites, Data0, Data1, True
            DrawSprite x, y, Data0, Data1, -1
        Case "objects"
            ''Draw BackBuffer, X, Y, Width, Height, sfcObjects, Data0, Data1, True
            DrawObject x, y, Data0
        Case "effects"
            ''Draw BackBuffer, X, Y, Width, Height, sfcEffects, Data0, Data1, True
        Case "tiles"
            ''Draw BackBuffer, X, Y, Width, Height, sfcTiles, Data0, Data1, True
            DrawTile x, y, Data0, 0, 0, Height, Width
            
        Case "tilesex"
            ''Draw BackBuffer, X, Y, Width, Height, sfcTiles, Data0, Data1, True
            DrawTileEX x, y, Width, Height, Data0, Data1
        Case "projectiles"
            ''Draw BackBuffer, X, Y, Width, Height, sfcProjectiles, Data0, Data1, True
        Case Else
            If ImageLoaded Then
                ''Draw BackBuffer, X, Y, Width, Height, wSurface, Data0, Data1, True
                
                D3DDevice.SetTexture 0, wTexture.Texture
                Draw3D x, y, Width, Height, Data0, Data1, 0, -1, wTexture.TexInfo.Width, wTexture.TexInfo.Height
                'Draw3DEx x, y, Width, Height, Data0, Data1, wTexture, -1
                LastTexture = 0
            End If
    End Select
End Sub

Private Sub Class_Initialize()
    ImageLoaded = False
End Sub

Private Sub Class_Terminate()
    Set wTexture.Texture = Nothing
End Sub
