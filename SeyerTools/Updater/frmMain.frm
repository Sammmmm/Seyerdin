VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Seyerdin Online Updater"
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   326
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   386
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4890
      Left            =   0
      ScaleHeight     =   326
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   386
      TabIndex        =   0
      Top             =   0
      Width           =   5790
      Begin VB.PictureBox picProgress 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         FillColor       =   &H00C0FFC0&
         ForeColor       =   &H00C0FFC0&
         Height          =   405
         Left            =   555
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1
         TabIndex        =   3
         Top             =   3675
         Width           =   15
      End
      Begin VB.PictureBox picCurrent 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         DrawMode        =   11  'Not Xor Pen
         DrawStyle       =   5  'Transparent
         FillColor       =   &H00C0FFC0&
         ForeColor       =   &H00C0FFC0&
         Height          =   405
         Left            =   555
         ScaleHeight     =   27
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   300
         TabIndex        =   1
         Top             =   3675
         Visible         =   0   'False
         Width           =   4500
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   600
         TabIndex        =   2
         Top             =   3360
         Width           =   4695
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

Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub cmdCancel_Click()
    blnEnd = True
End Sub

Private Sub Form_Load()
  picProgress.Picture = LoadPicture("Data\Graphics\Interface\UpdaterBar.rsc")
  Picture1.Picture = LoadPicture("Data\Graphics\Interface\NSUpdater.rsc")
End Sub


Private Sub Form_Activate()
                          'HWND_TOPMOST
    SetWindowPos Me.hwnd, 0, Me.Left / Screen.TwipsPerPixelX, Me.RightToLeft / Screen.TwipsPerPixelY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, &HA
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Me.WindowState = 0 Then
       Dim ReturnVal As Long
       ReleaseCapture
       ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Me.WindowState = 0 Then
       Dim ReturnVal As Long
       ReleaseCapture
       ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub

Private Sub lblStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Me.WindowState = 0 Then
       Dim ReturnVal As Long
       ReleaseCapture
       ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X >= 194 And Y >= 279 And X <= 271 And Y <= 306 Then 'play
        Dim ReturnVal As Long
        ReleaseCapture
        ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
    If X >= 279 And Y >= 279 And X <= 356 And Y <= 306 Then 'cancel
        blnEnd = True
    End If
    
    If X >= 360 And Y >= 7 And X <= 376 And Y <= 21 Then
        blnEnd = True
    End If
    
    
    If X >= 340 And Y >= 7 And X <= 355 And Y <= 21 Then
        frmMain.WindowState = vbMinimized

    End If

End Sub
