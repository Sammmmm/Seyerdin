VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Seyerdin Light Editor"
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   10380
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picColor 
      Height          =   255
      Left            =   360
      ScaleHeight     =   195
      ScaleWidth      =   3435
      TabIndex        =   10
      Top             =   1680
      Width           =   3495
   End
   Begin VB.HScrollBar sclBlue 
      Height          =   255
      Left            =   1080
      Max             =   255
      TabIndex        =   9
      Top             =   1320
      Width           =   2775
   End
   Begin VB.HScrollBar sclGreen 
      Height          =   255
      Left            =   1080
      Max             =   255
      TabIndex        =   8
      Top             =   1080
      Width           =   2775
   End
   Begin VB.HScrollBar sclRed 
      Height          =   255
      Left            =   1080
      Max             =   255
      TabIndex        =   7
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   2760
      Width           =   1815
   End
   Begin VB.ListBox lstLights 
      Height          =   10200
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label lblBlue 
      Caption         =   "0"
      Height          =   255
      Left            =   3960
      TabIndex        =   13
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblGreen 
      Caption         =   "0"
      Height          =   255
      Left            =   3960
      TabIndex        =   12
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblRed 
      Caption         =   "0"
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Blue:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Green:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Red:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1815
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

Private Sub cmdSave_Click()
    SaveLights
End Sub

Private Sub Command1_Click()
UpdateCurrentLight
End Sub

Private Sub Form_Load()
    LoadLights
End Sub

Private Sub lstLights_Click()
    If lstLights.ListIndex >= 0 Then
        If CurrentLight > 0 Then
            If CurrentLight <> lstLights.ListIndex + 1 Then
                UpdateCurrentLight
            End If
        End If
        CurrentLight = lstLights.ListIndex + 1
        With Lights(CurrentLight)
            txtName = ClipString$(.Name)
            sclRed.Value = .Red
            sclGreen.Value = .Green
            sclBlue.Value = .Blue
        End With
    End If
End Sub

Private Sub sclBlue_Change()
    DoColor
End Sub

Private Sub sclBlue_LostFocus()
    UpdateCurrentLight
End Sub

Private Sub sclBlue_Scroll()
    DoColor
End Sub

Private Sub sclGreen_Change()
    DoColor
End Sub

Private Sub sclGreen_LostFocus()
    UpdateCurrentLight
End Sub

Private Sub sclGreen_Scroll()
    DoColor
End Sub

Private Sub sclRed_Change()
    DoColor
End Sub

Private Sub sclRed_LostFocus()
    UpdateCurrentLight
End Sub

Private Sub sclRed_Scroll()
    DoColor
End Sub

Private Sub DoColor()
    lblRed = sclRed
    lblGreen = sclGreen
    lblBlue = sclBlue
    picColor.BackColor = RGB(sclRed, sclGreen, sclBlue)
End Sub

Public Sub UpdateCurrentLight()
    If CurrentLight > 0 Then
        With Lights(CurrentLight)
            .Name = txtName
            lstLights.List(CurrentLight - 1) = "[" & (CurrentLight) & "] " & .Name
            .Red = sclRed.Value
            .Green = sclGreen.Value
            .Blue = sclBlue.Value
        End With
    End If
End Sub

Private Sub txtName_KeyUp(KeyCode As Integer, Shift As Integer)
    UpdateCurrentLight
End Sub
