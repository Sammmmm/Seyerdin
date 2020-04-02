VERSION 5.00
Begin VB.Form frmMapEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seyerdin Online [Editing Map]"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3720
   Icon            =   "frmMapEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   575
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   248
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMapEdit 
      Height          =   8580
      Left            =   0
      ScaleHeight     =   8520
      ScaleWidth      =   3615
      TabIndex        =   0
      Top             =   0
      Width           =   3675
      Begin VB.PictureBox picAnim 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   195
         ScaleHeight     =   287
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   223
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   3375
         Begin VB.CheckBox chkTileFlag 
            Caption         =   "No Weather"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   39
            Top             =   3480
            Width           =   1935
         End
         Begin VB.HScrollBar sclLayer 
            Height          =   255
            Left            =   360
            Max             =   2
            TabIndex        =   35
            Top             =   2640
            Value           =   2
            Width           =   1935
         End
         Begin VB.HScrollBar sclNumFrames 
            Height          =   255
            Left            =   360
            Max             =   15
            TabIndex        =   27
            Top             =   840
            Width           =   2535
         End
         Begin VB.HScrollBar sclFrameDelay 
            Height          =   255
            Left            =   360
            Max             =   15
            TabIndex        =   26
            Top             =   1440
            Width           =   2535
         End
         Begin VB.Label Label5 
            Caption         =   "Tile Flags"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   38
            Top             =   3240
            Width           =   2055
         End
         Begin VB.Label lblLayer 
            Alignment       =   2  'Center
            Caption         =   "BGTile"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   37
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Layer:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   36
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label lblAnimLengthInfo 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   34
            Top             =   2040
            Width           =   2535
         End
         Begin VB.Label lblFrameDelayInfo 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   33
            Top             =   1800
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Animation Layer:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   120
            Width           =   3015
         End
         Begin VB.Label Label2 
            Caption         =   "Number of Frames:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   31
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label3 
            Caption         =   "Frame Delay:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   30
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblNumFrames 
            Alignment       =   2  'Center
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   29
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblFrameDelay 
            Alignment       =   2  'Center
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   28
            Top             =   1440
            Width           =   495
         End
      End
      Begin VB.PictureBox picTiles 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   4320
         Left            =   195
         ScaleHeight     =   288
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   224
         TabIndex        =   4
         Top             =   480
         Width           =   3360
      End
      Begin VB.PictureBox picTile 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   510
         Left            =   120
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   3
         Top             =   5640
         Width           =   510
      End
      Begin VB.VScrollBar MapScroll 
         Height          =   4290
         Left            =   35
         TabIndex        =   2
         Top             =   450
         Width           =   135
      End
      Begin VB.PictureBox picRecent 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   120
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   224
         TabIndex        =   1
         Top             =   8040
         Width           =   3360
      End
      Begin VB.Label lblMapOptions 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Removed"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   24
         Top             =   7080
         Width           =   735
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Up"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   3480
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Down"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   4920
         Width           =   3480
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Upload"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   4
         Left            =   2400
         TabIndex        =   21
         Top             =   5460
         Width           =   1155
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ground"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   6240
         Width           =   1095
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BGTile1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   19
         Top             =   6240
         Width           =   1095
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Anim"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   6660
         Width           =   1095
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FGTile"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   1320
         TabIndex        =   17
         Top             =   6660
         Width           =   1095
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Attribute"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   2520
         TabIndex        =   16
         Top             =   6660
         Width           =   1095
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   5
         Left            =   720
         TabIndex        =   15
         Top             =   5880
         Width           =   2835
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   54
         Left            =   720
         TabIndex        =   14
         Top             =   5460
         Width           =   795
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Paste"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   55
         Left            =   1560
         TabIndex        =   13
         Top             =   5460
         Width           =   795
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ground2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   1320
         TabIndex        =   12
         Top             =   6240
         Width           =   1095
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Map Properties"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   7680
         Width           =   3495
      End
      Begin VB.Label lblMapTile 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   5430
         Width           =   495
      End
      Begin VB.Label lblMapOptions 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Select"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   9
         Top             =   7380
         Width           =   735
      End
      Begin VB.Label lblEditMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Wall"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   8
         Top             =   7080
         Width           =   1095
      End
      Begin VB.Label lblMapOptions 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Convert"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   7
         Top             =   7080
         Width           =   735
      End
      Begin VB.Label lblMapOptions 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Trees"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   6
         Top             =   7380
         Width           =   735
      End
      Begin VB.Label lblMapOptions 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Undo"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   5
         Top             =   7080
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmMapEdit"
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

Option Explicit

Private sfcTiles As SurfaceType
Private sfcAtts As SurfaceType

Private LastAtt As Byte

Private Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER '40 bytes
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

Private Sub chkTileFlag_Click(Index As Integer)
    UpdateAnimInfo
End Sub

Private Sub Form_Activate()
    SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / Screen.twipsPerpixelX, Me.RightToLeft / Screen.twipsPerpixelY, Me.Width / Screen.twipsPerpixelX, Me.Height / Screen.twipsPerpixelY, &HA
End Sub

Private Sub Form_Load()
    Dim CK As DDCOLORKEY

    sfcTiles.desc.lFlags = DDSD_CAPS Or DDSD_HEIGHT
    sfcTiles.desc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY

    Dim FileHeader As BITMAPFILEHEADER
    Dim InfoHeader As BITMAPINFOHEADER
    Open AppPath & "Data/Graphics/Tiles.rsc" For Binary As #1
        Get #1, , FileHeader
        Get #1, , InfoHeader
    Close #1

    sfcTiles.desc.lHeight = InfoHeader.biHeight
    frmMapEdit.MapScroll.max = (sfcTiles.desc.lHeight - 192) / 32

    Set sfcTiles.Surface = DD.CreateSurfaceFromFile("Data/Graphics/Tiles.rsc", sfcTiles.desc)
    sfcTiles.Surface.SetColorKey DDCKEY_SRCBLT, CK

    sfcAtts.desc.lFlags = DDSD_CAPS
    sfcAtts.desc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set sfcAtts.Surface = DD.CreateSurfaceFromFile("Data/Graphics/Atts.rsc", sfcAtts.desc)
    sfcAtts.Surface.SetColorKey DDCKEY_SRCBLT, CK

    
    frmMapEdit_Loaded = True
    lblEditMode(0).BackColor = QBColor(8)
    lblEditMode(EditMode).BackColor = QBColor(15)
    'EditMode = 0

    RedrawRecentTiles

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set sfcTiles.Surface = Nothing
    Set sfcAtts.Surface = Nothing
    frmMapEdit_Loaded = False
    Dim x As Long, y As Long
    For x = 0 To 11
        For y = 0 To 11
            mapChangedBg(x, y) = True
        Next y
    Next x
    mapChanged = True
End Sub

Private Sub lblEditMode_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    lblEditMode(Index).BackColor = QBColor(15)
End Sub


Private Sub lblEditMode_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim A As Long
    If x >= 0 And x <= lblEditMode(Index).Width And y >= 0 And y <= lblEditMode(Index).Height Then
        If Index = 2 Then
            picAnim.Visible = True
        Else
            picAnim.Visible = False
        End If
        If Index <= 5 Or Index = 7 Then
            EditMode = Index
            For A = 0 To 5
                If A <> Index Then
                    lblEditMode(A).BackColor = QBColor(8)
                End If
            Next A
            If 7 <> Index Then
                lblEditMode(7).BackColor = QBColor(8)
            End If
        End If
        Select Case Index
            Case 2 'Animation Layer
                picAnim.Visible = True
            Case 6 'Properties
                If frmMapProperties_Loaded = False Then Load frmMapProperties
                With EditMap
                    frmMapProperties.Caption = "Seyerdin Online [Map " + CStr(CMap) + " Properties]"
                    frmMapProperties.txtName = EditMap.Name
                    frmMapProperties.sclMIDI = .MIDI
                    frmMapProperties.txtUp = CStr(.ExitUp)
                    frmMapProperties.txtDown = CStr(.ExitDown)
                    frmMapProperties.txtLeft = CStr(.ExitLeft)
                    frmMapProperties.txtRight = CStr(.ExitRight)
                    frmMapProperties.txtBootMap = CStr(.BootLocation.map)
                    frmMapProperties.txtBootX = CStr(.BootLocation.x)
                    frmMapProperties.txtBootY = CStr(.BootLocation.y)
                    frmMapProperties.sclIntensity = .Intensity
                    frmMapProperties.sclSnowColor = .SnowColor
                    frmMapProperties.sclRainColor = .RainCOlor

                    For A = 0 To 4
                        frmMapProperties.cmbMonster(A).ListIndex = .MonsterSpawn(A).Monster
                        frmMapProperties.sclRate(A) = .MonsterSpawn(A).Rate
                    Next A
                    For A = 0 To 7
                        If ExamineBit(.Flags(0), CByte(A)) Then
                            frmMapProperties.chkFlag(A) = 1
                        Else
                            frmMapProperties.chkFlag(A) = 0
                        End If
                    Next A
                    For A = 0 To 5
                        If ExamineBit(.Flags(1), CByte(A)) Then
                            frmMapProperties.chkFlag(A + 8) = 1
                        Else
                            frmMapProperties.chkFlag(A + 8) = 0
                        End If
                    Next A
                    frmMapProperties.sclRaining = IIf(.Raining > 3000, 3000, .Raining)
                    frmMapProperties.sclSnowing = IIf(.Snowing > 3000, 3000, .Snowing)
                    frmMapProperties.sclZone = .Zone
                    frmMapProperties.sclFog = IIf(.Fog <= 31, .Fog, 0)
                End With
                Me.Visible = False
                frmMapProperties.Show 1
                lblEditMode(Index).BackColor = QBColor(8)
        End Select
        RedrawTiles
        RedrawTile
        ReDoLightSources
    Else
        lblEditMode(Index).BackColor = QBColor(8)
    End If
End Sub

Private Sub lblMapOptions_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If lblMapOptions(Index).BackColor = QBColor(15) Then
    lblMapOptions(Index).BackColor = QBColor(8)
Else
    lblMapOptions(Index).BackColor = QBColor(15)
End If
End Sub

Private Sub lblMapOptions_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim A As Long, b As Long
    If x >= 0 And x <= lblMapOptions(Index).Width * Screen.twipsPerpixelX And y >= 0 And y <= lblMapOptions(Index).Height * Screen.twipsPerpixelY Then
        For A = 0 To 1
            If A <> Index Then
                lblMapOptions(A).BackColor = QBColor(8)
            End If
        Next A
        Select Case Index
            Case 2
                For x = 0 To 11
                    For y = 0 To 11
                        If EditMap.Tile(x, y).Ground = 40 Or (EditMap.Tile(x, y).Ground >= 988 And EditMap.Tile(x, y).Ground <= 1015) Then
                            If Rnd < 0.8 Then
                                EditMap.Tile(x, y).Ground = 988 + (Rnd * 18)
                            Else
                                EditMap.Tile(x, y).Ground = 1005 + (Rnd * 10)
                            End If
                        End If
                    Next y
                Next x
            Case 3
                Dim Trees(1 To 7, 1 To 2) As Long
                Trees(1, 1) = 169:      Trees(1, 2) = 176
                Trees(2, 1) = 1240:     Trees(2, 2) = 1247
                Trees(3, 1) = 1241:     Trees(3, 2) = 1248
                Trees(4, 1) = 1242:     Trees(4, 2) = 1249
                Trees(5, 1) = 1243:     Trees(5, 2) = 1250
                Trees(6, 1) = 1244:     Trees(6, 2) = 1251
                Trees(7, 1) = 1254:     Trees(7, 2) = 1261
                For x = 0 To 11
                    For y = 0 To 11
                        For A = 1 To 7
                            If Trees(A, 2) = EditMap.Tile(x, y).BGTile1 Then
                                b = Int(7 * Rnd) + 1
                                EditMap.Tile(x, y).BGTile1 = Trees(b, 2)
                                If y > 0 Then EditMap.Tile(x, y - 1).FGTile = Trees(b, 1)
                                Exit For
                            End If
                            If y = 11 Then
                                If Trees(A, 1) = EditMap.Tile(x, y).FGTile Then
                                    b = Int(7 * Rnd) + 1
                                    EditMap.Tile(x, y).FGTile = Trees(b, 1)
                                    Exit For
                                End If
                            End If
                        Next A
                    Next y
                Next x
                lblMapOptions(3).BackColor = QBColor(8)
            Case 4
                UndoMapEdit
                lblMapOptions(4).BackColor = QBColor(8)
        End Select
    End If
End Sub

Private Sub lblMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  lblMenu(Index).BackColor = QBColor(15)
End Sub


Private Sub lblMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    lblMenu(Index).BackColor = QBColor(8)
    If x >= 0 And x <= lblMenu(Index).Width * Screen.twipsPerpixelX And y >= 0 And y <= lblMenu(Index).Height * Screen.twipsPerpixelY Then
        Select Case Index
            Case 2 'MapEdit/Up
                If TopY > 0 Then
                    TopY = TopY - 32
                    RedrawTiles
                    MapScroll.Value = Int(TopY / 32)
                End If
            Case 3 'MapEdit/Down
                If TopY + 32 < sfcTiles.desc.lHeight - 192 Then
                    TopY = TopY + 32
                    RedrawTiles
                    MapScroll.Value = Int(TopY / 32)
                End If
            Case 4 'MapEdit/Upload
                UploadMap
                CloseMapEdit
            Case 5 'MapEdit/Cancel
                If MsgBox("Changes will be lost, continue?", vbYesNo + vbQuestion, TitleString) = vbYes Then
                    CloseMapEdit
                End If
            Case 54 'MapEdit/Copy
                CopyMap ClipboardMap, EditMap
            Case 55 'MapEdit/Paste
                If MsgBox("This will overwrite your current map-- are you sure you wish to paste?", vbYesNo + vbQuestion, TitleString) = vbYes Then
                    CopyMap EditMap, ClipboardMap
                    ReDoLightSources
                End If
        End Select
    End If
End Sub

Private Sub MapScroll_Change()
TopY = MapScroll.Value * 32
RedrawTiles
End Sub

Private Sub MapScroll_Scroll()
MapScroll_Change
End Sub

Private Sub picTile_Click()
    CurTile = 0
    BitBlt picTile.hdc, 0, 0, 32, 32, 0, 0, 0, BLACKNESS
    picTile.Refresh
End Sub

Private Sub picTiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim A As Long, TempByte As Byte
    If Button = 1 Then
        If EditMode < 4 Or EditMode = 5 Then
            CurTile = Int((y + TopY) / 32) * 7 + Int(x / 32) + 1
            If RecentTile(0) <> CurTile And RecentTile(1) <> CurTile And RecentTile(2) <> CurTile And RecentTile(3) <> CurTile And RecentTile(4) <> CurTile And RecentTile(5) <> CurTile And RecentTile(6) <> CurTile Then
                For A = 1 To 6
                    RecentTile(A - 1) = RecentTile(A)
                Next A
                RecentTile(6) = CurTile
                RedrawRecentTiles
            End If
            RedrawTile
        ElseIf EditMode = 7 Then
            NewAtt = Int(y / 32) * 7 + Int(x / 32) + 1
            'If NewAtt <= 11 Then
                If NewAtt = 1 Then
                    frmMapAtt.Show 1
                ElseIf NewAtt = 10 Or NewAtt = 11 Then
                    If NewAtt = 10 Then
                        If ExamineBit(CurWall, 0) Then
                            For A = 0 To 3
                                ClearBit CurWall, A
                            Next A
                        Else
                            For A = 0 To 3
                                SetBit CurWall, A
                            Next A
                        End If
                    End If
                    If NewAtt = 11 Then
                        If ExamineBit(CurWall, 4) Then
                            For A = 4 To 7
                                ClearBit CurWall, A
                            Next A
                        Else
                            For A = 4 To 7
                                SetBit CurWall, A
                            Next A
                        End If
                    End If
                Else
                    If NewAtt > 11 Then
                        CurWall = 0
                    Else
                        TempByte = CurWall
                        If ExamineBit(CurWall, NewAtt - 2) Then
                            ClearBit CurWall, NewAtt - 2
                        Else
                            SetBit CurWall, NewAtt - 2
                        End If
                    End If
                End If
                RedrawTile
            'End If
        Else
            NewAtt = Int(y / 32) * 7 + Int(x / 32) + 1
            If NewAtt <= 28 Then
                Select Case NewAtt
                    Case 2, 3, 6, 7, 8, 9, 14, 15, 17, 18, 19, 20, 22, 23, 24, 25, 26 'Warp, Key,Door, News, Object, Touch Plate, Damage, ClickTile, Wall2, Mining, NPC, 'movement
                        If LastAtt = NewAtt Then NewAtt = NewAtt + 50
                        'If CurAttData(0) <> 0 Or CurAttData(1) <> 0 Or CurAttData(2) <> 0 Or CurAttData(3) <> 0 Then NewAtt = NewAtt + 50
                        frmMapAtt.Show 1
                    Case 4
                        CurAttData(3) = 3
                        CurAtt = NewAtt
                    Case Else
                        CurAttData(0) = 0
                        CurAttData(1) = 0
                        CurAttData(2) = 0
                        CurAttData(3) = 0
                        CurAtt = NewAtt
                End Select
                RedrawTile
            End If
            LastAtt = CurAtt
        End If
    End If
    lblMapTile = CurTile
End Sub

Private Sub picRecent_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim A As Long
A = Int(x / 32)
If A < 0 Then A = 0
If A > 6 Then A = 6
Select Case EditMode
    Case 0, 1, 2, 3, 5 'Ground
        CurTile = RecentTile(A)
End Select
RedrawTile
End Sub

Sub RedrawTiles()
    Dim tDC As Long
    BitBlt frmMapEdit.picTiles.hdc, 0, 0, 224, 288, 0, 0, 0, BLACKNESS
    If EditMode < 4 Or EditMode = 5 Then
        tDC = sfcTiles.Surface.GetDC
        BitBlt frmMapEdit.picTiles.hdc, 0, 0, 224, 288, tDC, 0, TopY, SRCCOPY
        sfcTiles.Surface.ReleaseDC tDC
    ElseIf EditMode = 7 Then
        tDC = sfcAtts.Surface.GetDC
        BitBlt frmMapEdit.picTiles.hdc, 0, 0, 224, 192, tDC, 0, 160, SRCCOPY
        sfcAtts.Surface.ReleaseDC tDC
    Else
        tDC = sfcAtts.Surface.GetDC
        BitBlt frmMapEdit.picTiles.hdc, 0, 0, 224, 192, tDC, 0, 0, SRCCOPY
        sfcAtts.Surface.ReleaseDC tDC
    End If
    frmMapEdit.picTiles.Refresh
End Sub


Sub RedrawRecentTiles()
    Dim hdcTiles As Long, A As Long
    hdcTiles = sfcTiles.Surface.GetDC
    BitBlt frmMapEdit.picRecent.hdc, 0, 0, 224, 32, 0, 0, 0, BLACKNESS
    For A = 0 To 6
        If RecentTile(A) > 0 Then
            BitBlt frmMapEdit.picRecent.hdc, A * 32, 0, 32, 32, hdcTiles, ((RecentTile(A) - 1) Mod 7) * 32, Int((RecentTile(A) - 1) / 7) * 32, SRCCOPY
        End If
    Next A
    sfcTiles.Surface.ReleaseDC hdcTiles
    frmMapEdit.picRecent.Refresh
End Sub

Sub RedrawTile()
    Dim tDC As Long, A As Long
    BitBlt frmMapEdit.picTile.hdc, 0, 0, 32, 32, 0, 0, 0, BLACKNESS
    If EditMode < 4 Or EditMode = 5 Then
        If CurTile > 0 Then
            tDC = sfcTiles.Surface.GetDC
            BitBlt frmMapEdit.picTile.hdc, 0, 0, 32, 32, tDC, ((CurTile - 1) Mod 7) * 32, Int((CurTile - 1) / 7) * 32, SRCCOPY
            sfcTiles.Surface.ReleaseDC tDC
        End If
    ElseIf EditMode = 7 Then
        tDC = sfcAtts.Surface.GetDC
        For A = 0 To 7
            If ExamineBit(CurWall, A) Then TransparentBlt frmMapEdit.picTile.hdc, 0, 0, 32, 32, tDC, ((A + 1) Mod 7) * 32, 160 + Int((A + 1) / 7) * 32, SRCCOPY
        Next A
        sfcAtts.Surface.ReleaseDC tDC
    Else
        If CurAtt > 0 Then
            tDC = sfcAtts.Surface.GetDC
            BitBlt frmMapEdit.picTile.hdc, 0, 0, 32, 32, tDC, ((CurAtt - 1) Mod 7) * 32, Int((CurAtt - 1) / 7) * 32, SRCCOPY
            sfcAtts.Surface.ReleaseDC tDC
        End If
    End If
    frmMapEdit.picTile.Refresh
End Sub

Private Sub sclFrameDelay_Change()
    lblFrameDelay = (sclFrameDelay * 4)
    UpdateAnimInfo
    lblFrameDelayInfo = "Delay in seconds: " & ((1 / 40) * (sclFrameDelay * 4))
    lblAnimLengthInfo = "Total animation length: " & (((1 / 40) * (sclFrameDelay * 4)) * sclNumFrames)
End Sub

Private Sub sclFrameDelay_Scroll()
    sclFrameDelay_Change
End Sub

Private Sub sclLayer_Change()
    Select Case sclLayer
        Case 0
            lblLayer = "Ground"
        Case 1
            lblLayer = "Ground2"
        Case 2
            lblLayer = "BGTile"
    End Select
    UpdateAnimInfo
End Sub

Private Sub sclNumFrames_Change()
    lblNumFrames = sclNumFrames + 1
    UpdateAnimInfo
    lblAnimLengthInfo = "Total animation length: " & (((1 / 40) * sclFrameDelay) * sclNumFrames)
End Sub

Private Sub sclNumFrames_Scroll()
    sclNumFrames_Change
End Sub

Sub UpdateAnimInfo()
    CurAnim(1) = sclNumFrames * 16 + sclFrameDelay
    CurAnim(2) = (chkTileFlag(0).Value * 64) Or sclLayer
End Sub
