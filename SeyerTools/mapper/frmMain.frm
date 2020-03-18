VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "The Odyssey Online Classic Mapper Utility [Map]"
   ClientHeight    =   7950
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8385
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   530
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   559
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSTOP 
      Caption         =   "STOP"
      Height          =   375
      Left            =   7200
      TabIndex        =   7
      Top             =   120
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CMDialog 
      Left            =   6120
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.VScrollBar VScroll 
      Height          =   7125
      Left            =   8115
      Max             =   0
      TabIndex        =   6
      Top             =   540
      Width           =   255
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   0
      Max             =   0
      TabIndex        =   5
      Top             =   7680
      Width           =   8115
   End
   Begin VB.PictureBox picTop 
      Height          =   495
      Left            =   15
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   432
      TabIndex        =   1
      Top             =   15
      Width           =   6540
      Begin VB.CommandButton btnMenu 
         Caption         =   "Menu"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   45
         TabIndex        =   3
         Top             =   75
         Width           =   735
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   120
         Width           =   5415
      End
   End
   Begin VB.PictureBox picMapContainer 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7140
      Left            =   0
      ScaleHeight     =   472
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   537
      TabIndex        =   0
      Top             =   540
      Width           =   8115
      Begin VB.PictureBox picMap 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   0
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   4
         Top             =   0
         Width           =   480
      End
   End
   Begin VB.Menu mnuMap 
      Caption         =   "&Map"
      Visible         =   0   'False
      Begin VB.Menu mnuMapSave 
         Caption         =   "&Save Map"
      End
      Begin VB.Menu mnuMapDivider1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMapNumbers 
         Caption         =   "Draw Map Numbers"
      End
      Begin VB.Menu mnuMapBorders 
         Caption         =   "Draw Map Borders"
      End
      Begin VB.Menu mnuMapDivider2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMapExit 
         Caption         =   "E&xit"
      End
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

Private Sub btnMenu_Click()
    PopupMenu mnuMap
End Sub

Private Sub cmdSTOP_Click()
    blnStop = True
    cmdSTOP.Visible = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        blnEnd = True
    End If
End Sub

Private Sub Form_Load()
    frmMain_Loaded = True
End Sub

Private Sub Form_Resize()
        VScroll.Left = Me.ScaleWidth - 17
        VScroll.Height = Me.ScaleHeight - 55
        HScroll.Width = Me.ScaleWidth - 18
        HScroll.Top = Me.ScaleHeight - 17
        picMapContainer.Width = Me.ScaleWidth - 22
        picMapContainer.Height = Me.ScaleHeight - 58
        cmdSTOP.Left = 436 + (Me.ScaleWidth - 436) / 2 - 28.5
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain_Loaded = False
End Sub


Private Sub HScroll_Change()
    picMap.Left = 0 - HScroll
End Sub

Private Sub HScroll_Scroll()
    HScroll_Change
End Sub

Private Sub lblStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Me.WindowState = 0 Then
       Dim ReturnVal As Long
       ReleaseCapture
       ReturnVal = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub

Private Sub lblTitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Me.WindowState = 0 Then
       Dim ReturnVal As Long
       ReleaseCapture
       ReturnVal = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub

Private Sub mnuMapBorders_Click()
    If mnuMapBorders.Checked = True Then
        mnuMapBorders.Checked = False
    Else
        mnuMapBorders.Checked = True
    End If
    RefreshMap
End Sub

Private Sub mnuMapExit_Click()
    blnEnd = True
End Sub

Private Sub mnuMapNumbers_Click()
    If mnuMapNumbers.Checked = True Then
        mnuMapNumbers.Checked = False
    Else
        mnuMapNumbers.Checked = True
    End If
    RefreshMap
End Sub

Private Sub mnuMapSave_Click()
    With CMDialog
        .Filename = ""
        .DefaultExt = "BMP"
        .Filter = "Bitmap Files (*.BMP)|*.BMP"
        .flags = &H4& + &H2& + &H800&
        .Action = 2
        If Not .Filename = "" Then
            SavePicture picMap.Image, .Filename
        End If
    End With
End Sub

Private Sub picMap_DblClick()
    MapNum = MapGrid(SelectedX, SelectedY).MapNum
    If MapNum > 0 And blnDone = True Then
        Load frmWait
        With frmWait
            .lblStatus = "Drawing Map ..."
            .Show
            .Refresh
        End With
        DrawMap
        Load frmMap
        With frmMap
            BitBlt .picMap.hdc, 0, 0, 384, 384, hdcBuffer, 0, 0, SRCCOPY
            If Map(MapNum).ExitUp = 0 Then
                frmMap.lnExit(0).BorderColor = &HFF&
            Else
                If Map(Map(MapNum).ExitUp).ExitDown > 0 Then
                    frmMap.lnExit(0).BorderColor = &HFFFFFF
                Else
                    frmMap.lnExit(0).BorderColor = &HFFFF&
                End If
            End If
            If Map(MapNum).ExitDown = 0 Then
                frmMap.lnExit(1).BorderColor = &HFF&
            Else
                If Map(Map(MapNum).ExitDown).ExitUp > 0 Then
                    frmMap.lnExit(1).BorderColor = &HFFFFFF
                Else
                    frmMap.lnExit(1).BorderColor = &HFFFF&
                End If
            End If
            If Map(MapNum).ExitLeft = 0 Then
                frmMap.lnExit(2).BorderColor = &HFF&
            Else
                If Map(Map(MapNum).ExitLeft).ExitRight > 0 Then
                    frmMap.lnExit(2).BorderColor = &HFFFFFF
                Else
                    frmMap.lnExit(2).BorderColor = &HFFFF&
                End If
            End If
            If Map(MapNum).ExitRight = 0 Then
                frmMap.lnExit(3).BorderColor = &HFF&
            Else
                If Map(Map(MapNum).ExitRight).ExitLeft > 0 Then
                    frmMap.lnExit(3).BorderColor = &HFFFFFF
                Else
                    frmMap.lnExit(3).BorderColor = &HFFFF&
                End If
            End If
            .Caption = "The Odyssey Online Classic Mapper Utility [Map #" + CStr(MapNum) + ": " + Map(MapNum).Name + "]"
            .Show
        End With
        Unload frmWait
    End If
End Sub
Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelectedX = Int(X / MapWidth) + 1
    SelectedY = Int(Y / MapHeight) + 1
End Sub

Private Sub picTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Me.WindowState = 0 Then
       Dim ReturnVal As Long
       ReleaseCapture
       ReturnVal = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub

Private Sub VScroll_Change()
    picMap.Top = 0 - VScroll
End Sub

Private Sub VScroll_Scroll()
    VScroll_Change
End Sub

