VERSION 5.00
Begin VB.Form frmMacros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seyerdin Online [Macros]"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   Icon            =   "frmMacros.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   372
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLineFeed 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   9
      Left            =   7440
      TabIndex        =   19
      Top             =   4530
      Width           =   195
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   9
      Left            =   720
      MaxLength       =   255
      TabIndex        =   18
      Top             =   4440
      Width           =   6615
   End
   Begin VB.CheckBox chkLineFeed 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   8
      Left            =   7440
      TabIndex        =   17
      Top             =   4050
      Width           =   195
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   8
      Left            =   720
      MaxLength       =   255
      TabIndex        =   16
      Top             =   3960
      Width           =   6615
   End
   Begin VB.CheckBox chkLineFeed 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   7
      Left            =   7440
      TabIndex        =   15
      Top             =   3570
      Width           =   195
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   7
      Left            =   720
      MaxLength       =   255
      TabIndex        =   14
      Top             =   3480
      Width           =   6615
   End
   Begin VB.CheckBox chkLineFeed 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   6
      Left            =   7440
      TabIndex        =   13
      Top             =   3090
      Width           =   195
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   6
      Left            =   720
      MaxLength       =   255
      TabIndex        =   12
      Top             =   3000
      Width           =   6615
   End
   Begin VB.CheckBox chkLineFeed 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   5
      Left            =   7440
      TabIndex        =   11
      Top             =   2610
      Width           =   195
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   5
      Left            =   720
      MaxLength       =   255
      TabIndex        =   10
      Top             =   2520
      Width           =   6615
   End
   Begin VB.CheckBox chkLineFeed 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   4
      Left            =   7440
      TabIndex        =   9
      Top             =   2130
      Width           =   195
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   720
      MaxLength       =   255
      TabIndex        =   8
      Top             =   2040
      Width           =   6615
   End
   Begin VB.CheckBox chkLineFeed 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   7440
      TabIndex        =   7
      Top             =   1650
      Width           =   195
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   720
      MaxLength       =   255
      TabIndex        =   6
      Top             =   1560
      Width           =   6615
   End
   Begin VB.CheckBox chkLineFeed 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   7440
      TabIndex        =   5
      Top             =   1170
      Width           =   195
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   720
      MaxLength       =   255
      TabIndex        =   4
      Top             =   1080
      Width           =   6615
   End
   Begin VB.CheckBox chkLineFeed 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   7440
      TabIndex        =   3
      Top             =   690
      Width           =   195
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   720
      MaxLength       =   255
      TabIndex        =   2
      Top             =   600
      Width           =   6615
   End
   Begin VB.CheckBox chkLineFeed 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   7440
      TabIndex        =   1
      Top             =   210
      Width           =   195
   End
   Begin VB.TextBox txtMacro 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   720
      MaxLength       =   255
      TabIndex        =   0
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label btnCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4320
      TabIndex        =   21
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label btnOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5880
      TabIndex        =   20
      Top             =   5040
      Width           =   1455
   End
End
Attribute VB_Name = "frmMacros"
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

Private Sub btnCancel_Click()
    Unload Me
End Sub


Private Sub btnOk_Click()
    Dim A As Long
    For A = 0 To 9
        With Macro(A)
            .Text = txtMacro(A)
            .LineFeed = Choose(chkLineFeed(A) + 1, False, True)
        End With
        WriteString "Macros", "Text" + CStr(A + 1), txtMacro(A)
        WriteString "Macros", "LineFeed" + CStr(A + 1), CStr(chkLineFeed(A))
    Next A
    Unload Me
End Sub

Private Sub Form_Load()
    Dim A As Long
    For A = 0 To 9
        With Macro(A)
            txtMacro(A) = .Text
            If .LineFeed = True Then
                chkLineFeed(A) = 1
            Else
                chkLineFeed(A) = 0
            End If
        End With
    Next A
    frmMacros_Loaded = True
    Set Me.Picture = LoadPicture("Data/Graphics/Interface/Macros.rsc")
    Set Me.Icon = frmMenu.Icon
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMacros_Loaded = False
    Set Me.Picture = Nothing
End Sub


Private Sub txtMacro_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 8 Or (KeyAscii >= 32 And KeyAscii <= 127) Then
        'Valid Key
    Else
        KeyAscii = 0
        'Beep
    End If
End Sub

