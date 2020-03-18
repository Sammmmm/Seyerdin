VERSION 5.00
Begin VB.Form frmMapProperties 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seyerdin Online [Map Properties]"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   ControlBox      =   0   'False
   Icon            =   "frmMapProperties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   523
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar sclRainColor 
      Height          =   255
      Left            =   3000
      Max             =   255
      TabIndex        =   71
      Top             =   3960
      Width           =   2655
   End
   Begin VB.HScrollBar sclSnowColor 
      Height          =   255
      Left            =   3000
      Max             =   255
      TabIndex        =   69
      Top             =   4800
      Width           =   2655
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Projectiles Pass Through Green Walls"
      Height          =   435
      Index           =   13
      Left            =   3960
      TabIndex        =   68
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3600
      TabIndex        =   21
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Check Tile Flags"
      Height          =   195
      Index           =   11
      Left            =   2040
      TabIndex        =   66
      Top             =   7080
      Width           =   1935
   End
   Begin VB.HScrollBar sclFog 
      Height          =   255
      Left            =   3000
      Max             =   31
      TabIndex        =   64
      Top             =   6000
      Width           =   3135
   End
   Begin VB.HScrollBar sclZone 
      Height          =   255
      Left            =   3000
      Max             =   255
      TabIndex        =   61
      Top             =   5400
      Width           =   3135
   End
   Begin VB.HScrollBar sclSnowing 
      Height          =   255
      Left            =   3000
      Max             =   3000
      TabIndex        =   57
      Top             =   4560
      Value           =   1
      Width           =   3135
   End
   Begin VB.HScrollBar sclRaining 
      Height          =   255
      LargeChange     =   100
      Left            =   3000
      Max             =   3000
      TabIndex        =   56
      Top             =   3720
      Width           =   3135
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Always Start Music"
      Height          =   195
      Index           =   10
      Left            =   2040
      TabIndex        =   55
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Snowing"
      Height          =   195
      Index           =   9
      Left            =   3000
      TabIndex        =   54
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Raining"
      Height          =   195
      Index           =   8
      Left            =   3000
      TabIndex        =   53
      Top             =   3480
      Width           =   975
   End
   Begin VB.HScrollBar sclIntensity 
      Height          =   255
      Left            =   3000
      Max             =   255
      TabIndex        =   50
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Reset Items on Exit"
      Height          =   195
      Index           =   7
      Left            =   2040
      TabIndex        =   49
      Top             =   6360
      Width           =   1935
   End
   Begin VB.HScrollBar sclRate 
      Height          =   255
      Index           =   4
      Left            =   120
      Max             =   255
      TabIndex        =   47
      Top             =   3240
      Width           =   2175
   End
   Begin VB.ComboBox cmbMonster 
      Height          =   315
      Index           =   4
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   2880
      Width           =   2175
   End
   Begin VB.HScrollBar sclRate 
      Height          =   255
      Index           =   3
      Left            =   120
      Max             =   255
      TabIndex        =   44
      Top             =   4680
      Width           =   2175
   End
   Begin VB.ComboBox cmbMonster 
      Height          =   315
      Index           =   3
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox txtBootY 
      Height          =   285
      Left            =   5280
      MaxLength       =   2
      TabIndex        =   15
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtBootX 
      Height          =   285
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   14
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtBootMap 
      Height          =   285
      Left            =   3600
      MaxLength       =   4
      TabIndex        =   13
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtRight 
      Height          =   285
      Left            =   5280
      MaxLength       =   4
      TabIndex        =   12
      Top             =   1530
      Width           =   975
   End
   Begin VB.TextBox txtLeft 
      Height          =   285
      Left            =   3600
      MaxLength       =   4
      TabIndex        =   11
      Top             =   1530
      Width           =   975
   End
   Begin VB.TextBox txtDown 
      Height          =   285
      Left            =   5280
      MaxLength       =   4
      TabIndex        =   10
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtUp 
      Height          =   285
      Left            =   3600
      MaxLength       =   4
      TabIndex        =   9
      Top             =   1200
      Width           =   975
   End
   Begin VB.ComboBox cmbNPC 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   5400
      Width           =   2175
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Anyone can fight"
      Height          =   195
      Index           =   6
      Left            =   2040
      TabIndex        =   37
      Top             =   6840
      Width           =   1935
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Can't Attack Monsters"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   36
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Double Monsters"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   7200
      Width           =   1695
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Monsters Start on Map"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   19
      Top             =   6960
      Width           =   1935
   End
   Begin VB.ComboBox cmbMonster 
      Height          =   315
      Index           =   2
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2160
      Width           =   2175
   End
   Begin VB.ComboBox cmbMonster 
      Height          =   315
      Index           =   1
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3600
      Width           =   2175
   End
   Begin VB.ComboBox cmbMonster 
      Height          =   315
      Index           =   0
      ItemData        =   "frmMapProperties.frx":000C
      Left            =   120
      List            =   "frmMapProperties.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.HScrollBar sclMIDI 
      Height          =   255
      Left            =   1080
      Max             =   255
      TabIndex        =   1
      Top             =   600
      Width           =   3855
   End
   Begin VB.HScrollBar sclRate 
      Height          =   255
      Index           =   2
      Left            =   120
      Max             =   255
      TabIndex        =   5
      Top             =   2520
      Width           =   2175
   End
   Begin VB.HScrollBar sclRate 
      Height          =   255
      Index           =   1
      Left            =   120
      Max             =   255
      TabIndex        =   7
      Top             =   3960
      Width           =   2175
   End
   Begin VB.HScrollBar sclRate 
      Height          =   255
      Index           =   0
      Left            =   120
      Max             =   255
      TabIndex        =   3
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Arena"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Indoors"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   6480
      Width           =   975
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Friendly"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1080
      MaxLength       =   30
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   5160
      TabIndex        =   22
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Fish in Water"
      Height          =   195
      Index           =   12
      Left            =   2040
      TabIndex        =   67
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label lblRainColor 
      Alignment       =   2  'Center
      Caption         =   "Default"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   72
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblsnowcolor 
      Alignment       =   2  'Center
      Caption         =   "Default"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   70
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label lblFog 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   65
      Top             =   6000
      Width           =   615
   End
   Begin VB.Label Label17 
      Caption         =   "Fog:"
      Height          =   255
      Left            =   3000
      TabIndex        =   63
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label lblZone 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   62
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label Label15 
      Caption         =   "Zone:"
      Height          =   255
      Left            =   3000
      TabIndex        =   60
      Top             =   5160
      Width           =   735
   End
   Begin VB.Label lblSnowing 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   59
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label lblRaining 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   58
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lblIntensity 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   6120
      TabIndex        =   52
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label16 
      Caption         =   "Darkness"
      Height          =   255
      Left            =   3000
      TabIndex        =   51
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblRate 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   48
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblRate 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   45
      Top             =   4680
      Width           =   375
   End
   Begin VB.Label Label14 
      Caption         =   "Y:"
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
      Left            =   4680
      TabIndex        =   42
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label13 
      Caption         =   "X:"
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
      Left            =   3120
      TabIndex        =   41
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label12 
      Caption         =   "Map:"
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
      Left            =   3120
      TabIndex        =   40
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "NPC:"
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
      Left            =   120
      TabIndex        =   39
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Boot Location:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   38
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Exits:"
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
      Left            =   3120
      TabIndex        =   35
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Up:"
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
      Left            =   3120
      TabIndex        =   34
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "Down:"
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
      Left            =   4680
      TabIndex        =   33
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Left:"
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
      Left            =   3120
      TabIndex        =   32
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Right:"
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
      Left            =   4680
      TabIndex        =   31
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblMidi 
      Alignment       =   2  'Center
      Caption         =   "<None>"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   30
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Midi:"
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
      Left            =   120
      TabIndex        =   29
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblRate 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   28
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label lblRate 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   27
      Top             =   3960
      Width           =   375
   End
   Begin VB.Label lblRate 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   26
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Monsters:"
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
      Left            =   120
      TabIndex        =   25
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Flags:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Name:"
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
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmMapProperties"
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
Private Sub btnCancel_Click()
    Me.Hide
    frmMapEdit.Visible = True
End Sub

Private Sub btnOk_Click()
    Dim A As Long
    With EditMap
        .Name = txtName
        .MIDI = sclMIDI
        .ExitUp = Int(Val(txtUp))
        .ExitDown = Int(Val(txtDown))
        .ExitLeft = Int(Val(txtLeft))
        .ExitRight = Int(Val(txtRight))
        .BootLocation.map = Int(Val(txtBootMap))
        .BootLocation.x = Int(Val(txtBootX))
        .BootLocation.y = Int(Val(txtBootY))
        .Intensity = Abs(sclIntensity)
        .Raining = sclRaining
        .Snowing = sclSnowing
        .Zone = sclZone
        .Fog = sclFog
        .SnowColor = sclSnowColor
        .RainCOlor = sclRainColor
        If sclIntensity < 0 Then
            SetBit .Intensity, 7
        End If
        For A = 0 To 4
            .MonsterSpawn(A).Monster = cmbMonster(A).ListIndex
            .MonsterSpawn(A).Rate = sclRate(A)
        Next A
        For A = 0 To 7
            If chkFlag(A) = 1 Then
                SetBit .Flags(0), CByte(A)
            Else
                ClearBit .Flags(0), CByte(A)
            End If
        Next A
        For A = 0 To 5
            If chkFlag(A + 8) = 1 Then
                SetBit .Flags(1), CByte(A)
            Else
                ClearBit .Flags(1), CByte(A)
            End If
        Next A
    End With
    Me.Hide
    frmMapEdit.Visible = True
End Sub

Private Sub Form_Load()
    Dim A As Long

    cmbMonster(0).AddItem "<None>"
    cmbMonster(1).AddItem "<None>"
    cmbMonster(2).AddItem "<None>"
    cmbMonster(3).AddItem "<None>"
    cmbMonster(4).AddItem "<None>"
    cmbNPC.AddItem "<None>"
    
    For A = 1 To MAXITEMS
        cmbMonster(0).AddItem CStr(A) + ": " + Monster(A).Name
        cmbMonster(1).AddItem CStr(A) + ": " + Monster(A).Name
        cmbMonster(2).AddItem CStr(A) + ": " + Monster(A).Name
        cmbMonster(3).AddItem CStr(A) + ": " + Monster(A).Name
        cmbMonster(4).AddItem CStr(A) + ": " + Monster(A).Name
        
    Next A
    For A = 1 To 255
        cmbNPC.AddItem CStr(A) + ": " + NPC(A).Name
    Next A
    frmMapProperties_Loaded = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmMapProperties_Loaded = False
End Sub

Private Sub Label18_Click()

End Sub

Private Sub sclFog_Change()
    lblFog = sclFog
End Sub

Private Sub sclFog_Scroll()
    sclFog_Change
End Sub

Private Sub sclIntensity_Change()
    lblIntensity.Caption = sclIntensity
End Sub

Private Sub sclMIDI_Change()
    If sclMIDI = 0 Then
        lblMidi = "<None>"
    Else
        lblMidi = sclMIDI
    End If
End Sub
Private Sub sclMIDI_Scroll()
    sclMIDI_Change
End Sub

Private Sub sclRainColor_Change()
If sclRainColor.Value = 0 Then
        lblRainColor.Caption = "Default"
    Else
        lblRainColor.Caption = Lights(sclRainColor).Name
    End If
End Sub

Private Sub sclRaining_Change()
    lblRaining = sclRaining
End Sub

Private Sub sclRaining_Scroll()
    sclRaining_Change
End Sub

Private Sub sclRate_Change(Index As Integer)
    lblRate(Index) = sclRate(Index)
End Sub

Private Sub sclRate_Scroll(Index As Integer)
    sclRate_Change (Index)
End Sub

Private Sub sclSnowColor_Change()
    If sclSnowColor.Value = 0 Then
        lblsnowcolor.Caption = "Default"
    Else
        lblsnowcolor.Caption = Lights(sclSnowColor).Name
    End If
End Sub

Private Sub sclSnowing_Change()
    lblSnowing = sclSnowing
End Sub

Private Sub sclSnowing_Scroll()
    sclSnowing_Change
End Sub

Private Sub sclZone_Change()
    lblZone = sclZone
End Sub

Private Sub sclZone_Scroll()
    sclZone_Change
End Sub

Private Sub txtBootMap_LostFocus()
    Dim A As Double
    A = Int(Val(txtBootMap))
    If A > 5000 Then A = 5000
    If A < 0 Then A = 0
    txtBootMap = CStr(A)
End Sub

Private Sub txtBootX_LostFocus()
    Dim A As Double
    A = Int(Val(txtBootX))
    If A > 11 Then A = 11
    If A < 0 Then A = 0
    txtBootX = CStr(A)
End Sub

Private Sub txtBootY_LostFocus()
    Dim A As Double
    A = Int(Val(txtBootY))
    If A > 11 Then A = 11
    If A < 0 Then A = 0
    txtBootY = CStr(A)
End Sub

Private Sub txtDown_LostFocus()
    Dim A As Double
    A = Int(Val(txtDown))
    If A > 5000 Then A = 5000
    If A < 0 Then A = 0
    txtDown = CStr(A)
End Sub

Private Sub txtLeft_LostFocus()
    Dim A As Double
    A = Int(Val(txtLeft))
    If A > 5000 Then A = 5000
    If A < 0 Then A = 0
    txtLeft = CStr(A)
End Sub

Private Sub txtRight_LostFocus()
    Dim A As Double
    A = Int(Val(txtRight))
    If A > 5000 Then A = 5000
    If A < 0 Then A = 0
    txtRight = CStr(A)
End Sub

Private Sub txtUp_LostFocus()
    Dim A As Double
    A = Int(Val(txtUp))
    If A > 5000 Then A = 5000
    If A < 0 Then A = 0
    txtUp = CStr(A)
End Sub
