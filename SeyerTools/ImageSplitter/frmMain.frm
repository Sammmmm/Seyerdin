VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   10725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16065
   LinkTopic       =   "Form1"
   ScaleHeight     =   715
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1071
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Command1"
      Height          =   615
      Left            =   8040
      TabIndex        =   2
      Top             =   7920
      Width           =   1815
   End
   Begin VB.PictureBox picDest 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   8040
      ScaleHeight     =   3840
      ScaleWidth      =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   3840
   End
   Begin VB.PictureBox picSrc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1.80000e5
      Left            =   120
      ScaleHeight     =   16000
      ScaleMode       =   0  'User
      ScaleWidth      =   512
      TabIndex        =   0
      Top             =   120
      Width           =   7680
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Sub cmdConvert_Click()
Dim A As Long, B As Long
For A = 1 To 40 '"tilesets"
    For B = 1 To 64
        BitBlt picDest.hDC, ((B - 1) Mod 8) * 32, ((B - 1) \ 8) * 32, 32, 32, picSrc.hDC, ((((A - 1) * 64) + (B - 1)) Mod 7) * 32, ((((A - 1) * 64) + (B - 1)) \ 7) * 32, &HCC0020
    Next B
    SavePicture picDest.Image, App.Path & "/Tile" & Format$(A, "000") & ".rsc"
Next A

'For A = 1 To 20 '"tilesets"
'    'For B = 1 To 64
'        'BitBlt picDest.hDC, ((B - 1) Mod 8) * 32, ((B - 1) \ 8) * 32, 32, 32, picSrc.hDC, ((((A - 1) * 64) + (B - 1)) Mod 7) * 32, ((((A - 1) * 64) + (B - 1)) \ 7) * 32, &HCC0020
'        BitBlt picDest.hDC, 0, 0, 512, 512, picSrc.hDC, 0, (A - 1) * 512, &HCC0020
'        'SavePicture picDest.Image, App.Path & "/Object" & Format$(B, "000") & ".rsc"
'        'BitBlt picDest.hDC, ((B - 1) Mod 8) * 32, ((B - 1) \ 8) * 32, 32, 32, picSrc.hDC, 0, (((A - 1) * 64) + (B - 1)) * 32, &HCC0020
'        'BitBlt picDest.hDC, ((B - 1) Mod 16) * 32, ((B - 1) \ 16) * 32, 32, 32, picSrc.hDC, 96, (B - 1) * 32, &HCC0020
'    'Next B
'    'SavePicture picDest.Image, App.Path & "/Tile" & Format$(A, "000") & ".rsc"
'    SavePicture picDest.Image, App.Path & "/Sprite" & Format$(A, "000") & ".rsc"
'Next A

'    For B = 1 To 256
'        BitBlt picDest.hDC, ((B - 1) Mod 16) * 32, ((B - 1) \ 16) * 32, 32, 32, picSrc.hDC, 96, (B - 1) * 32, &HCC0020
'    Next B
'SavePicture picDest.Image, App.Path & "/Sprites.rsc"
picDest.Refresh
End Sub

Private Sub Form_Load()
    On Error Resume Next
    picSrc.Picture = LoadPicture(App.Path & "/tiles.rsc")
End Sub
