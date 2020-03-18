VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H006BA6CE&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seyerdin Online Skill Editor"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmMain.frx":0000
   ScaleHeight     =   485
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   613
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar sclBlue 
      Height          =   135
      Left            =   4200
      Max             =   255
      TabIndex        =   53
      Top             =   5520
      Width           =   1815
   End
   Begin VB.HScrollBar sclGreen 
      Height          =   135
      Left            =   4200
      Max             =   255
      TabIndex        =   52
      Top             =   5280
      Width           =   1815
   End
   Begin VB.HScrollBar sclRed 
      Height          =   135
      Left            =   4200
      Max             =   255
      TabIndex        =   51
      Top             =   5040
      Width           =   1815
   End
   Begin VB.PictureBox picColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4800
      ScaleHeight     =   225
      ScaleWidth      =   1545
      TabIndex        =   50
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox txtCoolDown 
      Height          =   285
      Index           =   1
      Left            =   5640
      TabIndex        =   49
      Text            =   "1"
      Top             =   4320
      Width           =   735
   End
   Begin VB.TextBox txtCoolDown 
      Height          =   285
      Index           =   0
      Left            =   5640
      TabIndex        =   48
      Text            =   "1"
      Top             =   4680
      Width           =   735
   End
   Begin VB.TextBox txtMana 
      Height          =   285
      Index           =   2
      Left            =   5400
      TabIndex        =   36
      Text            =   "1"
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox txtMana 
      Height          =   285
      Index           =   1
      Left            =   4680
      TabIndex        =   34
      Text            =   "1"
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox txtMana 
      Height          =   285
      Index           =   0
      Left            =   3120
      TabIndex        =   32
      Text            =   "1"
      Top             =   3480
      Width           =   615
   End
   Begin VB.HScrollBar sclIcon 
      Height          =   255
      Left            =   4200
      Max             =   100
      TabIndex        =   28
      Top             =   3960
      Width           =   1935
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6180
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   27
      Top             =   3960
      Width           =   240
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H006BA6CE&
      Caption         =   "Requirements"
      Height          =   1695
      Left            =   1920
      TabIndex        =   19
      Top             =   3840
      Width           =   2175
      Begin VB.TextBox txtRequirementLevel 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   24
         Text            =   "0"
         Top             =   1320
         Width           =   495
      End
      Begin VB.ComboBox cmbRequirementSkill 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         ItemData        =   "frmMain.frx":B6F22
         Left            =   120
         List            =   "frmMain.frx":B6F29
         TabIndex        =   23
         Text            =   "No Skill"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtRequirementLevel 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   22
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.ComboBox cmbRequirementSkill 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         ItemData        =   "frmMain.frx":B6F37
         Left            =   120
         List            =   "frmMain.frx":B6F3E
         TabIndex        =   20
         Text            =   "No Skill"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label13 
         BackColor       =   &H006BA6CE&
         Caption         =   "Level:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label12 
         BackColor       =   &H006BA6CE&
         Caption         =   "Level:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H006BA6CE&
      Caption         =   "Target Type"
      Height          =   1455
      Left            =   120
      TabIndex        =   17
      Top             =   3840
      Width           =   1695
      Begin VB.ListBox lstTargetType 
         Appearance      =   0  'Flat
         Height          =   1155
         ItemData        =   "frmMain.frx":B6F4C
         Left            =   120
         List            =   "frmMain.frx":B6F5F
         Style           =   1  'Checkbox
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H006BA6CE&
      Caption         =   "Flags"
      Height          =   2415
      Left            =   2520
      TabIndex        =   15
      Top             =   720
      Width           =   1935
      Begin VB.ListBox lstFlags 
         Appearance      =   0  'Flat
         Height          =   2055
         ItemData        =   "frmMain.frx":B6F90
         Left            =   120
         List            =   "frmMain.frx":B6FAF
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Skills"
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update Skill"
      Height          =   375
      Left            =   1560
      TabIndex        =   13
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add New Skill"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   6840
      Width           =   1335
   End
   Begin VB.ListBox lstSkills 
      Appearance      =   0  'Flat
      Height          =   7050
      Left            =   6480
      TabIndex        =   11
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox txtMaxLevel 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5400
      MaxLength       =   2
      TabIndex        =   10
      Text            =   "1"
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtRange 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5400
      MaxLength       =   2
      TabIndex        =   8
      Text            =   "1"
      Top             =   2400
      Width           =   495
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H006BA6CE&
      Caption         =   "Skill Flags"
      Height          =   1455
      Left            =   4560
      TabIndex        =   5
      Top             =   840
      Width           =   1695
      Begin VB.ListBox lstSkillType 
         Appearance      =   0  'Flat
         Height          =   1155
         ItemData        =   "frmMain.frx":B701C
         Left            =   120
         List            =   "frmMain.frx":B702F
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H006BA6CE&
      Caption         =   "Class Requirements"
      Height          =   3090
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2295
      Begin VB.TextBox txtMinLevel 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   57
         Text            =   "0"
         Top             =   2640
         Width           =   495
      End
      Begin VB.TextBox txtMinLevel 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   39
         Text            =   "0"
         Top             =   780
         Width           =   495
      End
      Begin VB.TextBox txtMinLevel 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   38
         Text            =   "0"
         Top             =   510
         Width           =   495
      End
      Begin VB.TextBox txtMinLevel 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   45
         Text            =   "0"
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtMinLevel 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   44
         Text            =   "0"
         Top             =   2130
         Width           =   495
      End
      Begin VB.TextBox txtMinLevel 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   43
         Text            =   "0"
         Top             =   1860
         Width           =   495
      End
      Begin VB.TextBox txtMinLevel 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   42
         Text            =   "0"
         Top             =   1590
         Width           =   495
      End
      Begin VB.TextBox txtMinLevel 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   41
         Text            =   "0"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtMinLevel 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   40
         Text            =   "0"
         Top             =   1050
         Width           =   495
      End
      Begin VB.TextBox txtMinLevel 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   37
         Text            =   "0"
         Top             =   240
         Width           =   495
      End
      Begin VB.ListBox lstClass 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2730
         ItemData        =   "frmMain.frx":B705B
         Left            =   120
         List            =   "frmMain.frx":B707D
         Style           =   1  'Checkbox
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   120
      MaxLength       =   256
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   5760
      Width           =   4575
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1200
      MaxLength       =   32
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.PictureBox picIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   7800
      ScaleHeight     =   3840
      ScaleWidth      =   3840
      TabIndex        =   29
      Top             =   5280
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.Label lblBlue 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   6000
      TabIndex        =   56
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label lblGreen 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   6000
      TabIndex        =   55
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label lblRed 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   6000
      TabIndex        =   54
      Top             =   5040
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Global Cooldown"
      Height          =   255
      Left            =   4200
      TabIndex        =   47
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Local Cooldown"
      Height          =   255
      Left            =   4200
      TabIndex        =   46
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   35
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "+ SLVL * "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   33
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "LVL *"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   31
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mana Requirements:"
      Height          =   255
      Left            =   2520
      TabIndex        =   30
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label Label11 
      BackColor       =   &H006BA6CE&
      Caption         =   "MaxLevel:"
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label10 
      BackColor       =   &H006BA6CE&
      Caption         =   "Range:"
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H006BA6CE&
      Caption         =   "Description:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H006BA6CE&
      Caption         =   "Name:"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmmain"
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

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Sub cmdAdd_Click()
    Dim A As Long
    A = UBound(Skills())
    ReDim Preserve Skills(0 To (A + 1))
    
    UpdateSkillList True
End Sub

Private Sub cmdSave_Click()
    Dim St As String, St1 As String * 32, St2 As String * 256, St3 As String
    Dim A As Long, B As Long

    Open App.Path & "/skilldata.dat" For Binary As #1
        If UBound(Skills()) > 0 Then
            For A = 1 To UBound(Skills())
                With Skills(A)
                    St1 = .Name
                    St2 = .Description
                    St = St1 & St2 & DoubleChar(.Class) & Chr$(.Icon) & Chr$(.Red) & _
                    Chr$(.Green) & Chr$(.Blue) & Chr$(.TargetType) & QuadChar(.Type) & QuadChar(.Flags) & _
                    Chr$(.Range) & Chr$(.MaxLevel) & Chr$(.EXPTable) & Chr$(.Requirements(0).Skill) & _
                    Chr$(.Requirements(0).Level) & Chr$(.Requirements(1).Skill) & Chr$(.Requirements(1).Level) & _
                    Chr$(.ManaCost1(1)) & Chr$(.ManaCost1(2)) & Chr$(.ManaCost1(3)) & Chr$(.ManaCost1(4)) & Chr$(.ManaCost(2)) & Chr$(.ManaCost(3)) & _
                    Chr$(.Level(1)) & Chr$(.Level(2)) & Chr$(.Level(3)) & Chr$(.Level(4)) & Chr$(.Level(5)) & Chr$(.Level(6)) & Chr$(.Level(7)) & Chr$(.Level(8)) & Chr$(.Level(9)) & Chr$(.Level(10)) & _
                    DoubleChar$(.GlobalTick) + DoubleChar$(.LocalTick) + Chr$(0)
                End With
                Put #1, , St
            Next A
        End If
    Close #1
End Sub

Private Sub cmdUpdate_Click()
    Dim A As Long, tmpFloat As Single
    
    If CurrentSkill > 0 Then
        With Skills(CurrentSkill)
            .Name = txtName
            .Description = txtDescription
            
            .Class = 0
            For A = 0 To lstClass.ListCount - 1
                If lstClass.Selected(A) Then
                    .Class = .Class Or (2 ^ A)
                End If
            Next A
            
            tmpFloat = Val(txtMana(0))
            CopyMem .ManaCost1(1), tmpFloat, 4
            .ManaCost(2) = Val(txtMana(1))
            .ManaCost(3) = Val(txtMana(2))
            
            .Type = 0
            For A = 0 To (lstSkillType.ListCount - 1)
                If lstSkillType.Selected(A) Then
                    .Type = .Type + (2 ^ A)
                End If
            Next A
            
            .Flags = 0
            For A = 0 To (lstFlags.ListCount - 1)
                If lstFlags.Selected(A) Then
                    .Flags = .Flags + (2 ^ A)
                End If
            Next A
            
            .TargetType = 0
            For A = 0 To (lstTargetType.ListCount - 1)
                If lstTargetType.Selected(A) Then
                    .TargetType = .TargetType + (2 ^ A)
                End If
            Next A

            If cmbRequirementSkill(0).ListIndex > 0 Then
                .Requirements(0).Skill = cmbRequirementSkill(0).ListIndex
                .Requirements(0).Level = txtRequirementLevel(0)
            Else
                .Requirements(0).Skill = 0
                .Requirements(0).Level = 0
            End If
            If cmbRequirementSkill(1).ListIndex > 0 Then
                .Requirements(1).Skill = cmbRequirementSkill(1).ListIndex
                .Requirements(1).Level = txtRequirementLevel(1)
            Else
                .Requirements(1).Skill = 0
                .Requirements(1).Level = 0
            End If
            
            .Range = txtRange
            .MaxLevel = txtMaxLevel
            For A = 1 To 10
                .Level(A) = txtMinLevel(A - 1)
            Next A
            
            .LocalTick = Val(txtCoolDown(1))
            .GlobalTick = Val(txtCoolDown(0))
                        
            .Icon = sclIcon
            .Red = sclRed
            .Green = sclGreen
            .Blue = sclBlue
        End With
    End If
    
    UpdateSkillList False
End Sub

Public Sub UpdateCurrentSkill()
Dim A As Long, tmpFloat As Single
    With Skills(CurrentSkill)
        txtName = .Name
        txtDescription = .Description
    
        For A = 0 To 9
            lstClass.Selected(A) = IIf(.Class And (2 ^ A), True, False)
        Next A
        
        CopyMem tmpFloat, .ManaCost1(1), 4
        txtMana(0) = tmpFloat
        txtMana(1) = .ManaCost(2)
        txtMana(2) = .ManaCost(3)
        
        For A = 0 To (lstSkillType.ListCount - 1)
            If (2 ^ A) And .Type Then
                lstSkillType.Selected(A) = True
            Else
                lstSkillType.Selected(A) = False
            End If
        Next A
        
        For A = 0 To (lstFlags.ListCount - 1)
            If (2 ^ A And .Flags) Then
                lstFlags.Selected(A) = True
            Else
                lstFlags.Selected(A) = False
            End If
        Next A
        
        For A = 0 To (lstTargetType.ListCount - 1)
            If (2 ^ A And .TargetType) Then
                lstTargetType.Selected(A) = True
            Else
                lstTargetType.Selected(A) = False
            End If
        Next A
        
        cmbRequirementSkill(0).ListIndex = .Requirements(0).Skill
        txtRequirementLevel(0) = .Requirements(0).Level
        cmbRequirementSkill(1).ListIndex = .Requirements(1).Skill
        txtRequirementLevel(1) = .Requirements(1).Level
        
        txtRange = .Range
        txtMaxLevel = .MaxLevel
        
        For A = 1 To 10
            txtMinLevel(A - 1) = .Level(A)
        Next A
        
        
        txtCoolDown(0) = .GlobalTick
        txtCoolDown(1) = .LocalTick
        
        sclRed = .Red
        sclGreen = .Green
        sclBlue = .Blue
        
        sclIcon = .Icon
    End With
End Sub

Private Sub Form_Load()
    LoadSkillData
    
    'On Error Resume Next
    picIcons.Picture = LoadPicture(App.Path & "/Controls001.bmp")
End Sub

Public Sub LoadSkillData()
    Dim St As String * 331, A As Long, B As Long, C As Long
    Dim CurrentSkill As SkillType
    ReDim Skills(0)
    Open App.Path & "/skilldata.dat" For Binary As #1
        If LOF(1) Mod 331 = 0 And LOF(1) > 0 Then
            A = LOF(1) / 331
            If A > 120 Then A = 120
            ReDim Skills(0 To A)
            For B = 1 To A
                Get #1, , St
                
                With Skills(B)
                    .Name = ClipString$(Mid$(St, 1, 32))
                    .Description = ClipString$(Mid$(St, 33, 256))

                    .Class = Asc(Mid$(St, 289, 1)) * 256& + Asc(Mid$(St, 290, 1))
                    .Icon = Asc(Mid$(St, 291, 1))
                    .Red = Asc(Mid$(St, 292, 1))
                    .Green = Asc(Mid$(St, 293, 1))
                    .Blue = Asc(Mid$(St, 294, 1))
                    .TargetType = Asc(Mid$(St, 295, 1))
                    .Type = Asc(Mid$(St, 296, 1)) * 16777216 + Asc(Mid$(St, 297, 1)) * 65536 + Asc(Mid$(St, 298, 1)) * 256& + Asc(Mid$(St, 299, 1))
                    .Flags = Asc(Mid$(St, 300, 1)) * 16777216 + Asc(Mid$(St, 301, 1)) * 65536 + Asc(Mid$(St, 302, 1)) * 256& + Asc(Mid$(St, 303, 1))
                    .Range = Asc(Mid$(St, 304, 1))
                    .MaxLevel = Asc(Mid$(St, 305, 1))
                    .EXPTable = Asc(Mid$(St, 306, 1))
                    .Requirements(0).Skill = Asc(Mid$(St, 307, 1))
                    .Requirements(0).Level = Asc(Mid$(St, 308, 1))
                    .Requirements(1).Skill = Asc(Mid$(St, 309, 1))
                    .Requirements(1).Level = Asc(Mid$(St, 310, 1))
                    .ManaCost1(1) = Asc(Mid$(St, 311, 1))
                    .ManaCost1(2) = Asc(Mid$(St, 312, 1))
                    .ManaCost1(3) = Asc(Mid$(St, 313, 1))
                    .ManaCost1(4) = Asc(Mid$(St, 314, 1))
                    .ManaCost(2) = Asc(Mid$(St, 315, 1))
                    .ManaCost(3) = Asc(Mid$(St, 316, 1))
                    .Level(1) = Asc(Mid$(St, 317, 1))
                    .Level(2) = Asc(Mid$(St, 318, 1))
                    .Level(3) = Asc(Mid$(St, 319, 1))
                    .Level(4) = Asc(Mid$(St, 320, 1))
                    .Level(5) = Asc(Mid$(St, 321, 1))
                    .Level(6) = Asc(Mid$(St, 322, 1))
                    .Level(7) = Asc(Mid$(St, 323, 1))
                    .Level(8) = Asc(Mid$(St, 324, 1))
                    .Level(9) = Asc(Mid$(St, 325, 1))
                    .Level(10) = Asc(Mid$(St, 326, 1))
                    .GlobalTick = GetInt(Mid$(St, 327, 2))
                    .LocalTick = GetInt(Mid$(St, 329, 2))
                End With
            Next B
        End If
    Close #1

    UpdateSkillList True
End Sub

Public Sub UpdateSkillList(AddNew As Boolean)
    Dim A As Long, St As String

    If AddNew Then
        If UBound(Skills()) > 0 Then
            lstSkills.Clear
            cmbRequirementSkill(0).Clear
            cmbRequirementSkill(1).Clear
            cmbRequirementSkill(0).AddItem "No Skill"
            cmbRequirementSkill(1).AddItem "No Skill"

            For A = 1 To UBound(Skills())
                If Len(Skills(A).Name) > 0 Then
                    lstSkills.AddItem "[" & A & "] " & (Skills(A).Name)
                    cmbRequirementSkill(0).AddItem Skills(A).Name
                    cmbRequirementSkill(1).AddItem Skills(A).Name
                Else
                    lstSkills.AddItem "<Unnamed>"
                    cmbRequirementSkill(0).AddItem "<Unnamed>"
                    cmbRequirementSkill(1).AddItem "<Unnamed>"
                End If
            Next A
        End If
    Else
            cmbRequirementSkill(0).List(CurrentSkill) = Skills(CurrentSkill).Name
            cmbRequirementSkill(1).List(CurrentSkill) = Skills(CurrentSkill).Name
            lstSkills.List(CurrentSkill - 1) = "[" & CurrentSkill & "] " & Skills(CurrentSkill).Name
    End If

    UpdateCurrentSkill
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ReturnVal As Long
    ReleaseCapture
    ReturnVal = SendMessage(Me.hwnd, &HA1, 2, 0)
End Sub

Private Sub lstSkills_Click()
    CurrentSkill = lstSkills.ListIndex + 1
    UpdateCurrentSkill
End Sub

Private Sub sclBlue_Change()
    lblBlue = sclBlue
    picColor.BackColor = RGB(sclRed, sclGreen, sclBlue)
End Sub

Private Sub sclBlue_Scroll()
    sclBlue_Change
End Sub

Private Sub sclGreen_Change()
    lblGreen = sclGreen
    picColor.BackColor = RGB(sclRed, sclGreen, sclBlue)
End Sub

Private Sub sclGreen_Scroll()
    sclGreen_Change
End Sub

Private Sub sclIcon_Change()
    BitBlt picIcon.hDC, 0, 0, 16, 16, picIcons.hDC, 128 + (((sclIcon - 1) Mod 8) * 16), 32 + ((sclIcon - 1) \ 8) * 32, vbSrcCopy
    picIcon.Refresh
End Sub

Private Sub sclRed_Change()
    lblRed = sclRed
    picColor.BackColor = RGB(sclRed, sclGreen, sclBlue)
End Sub

Private Sub sclRed_Scroll()
    sclRed_Change
End Sub

Private Sub txtCoolDown_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        If KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtCoolDown_LostFocus(Index As Integer)
    txtCoolDown(Index) = Val(txtCoolDown(Index))
End Sub

Private Sub txtMaxLevel_KeyPress(KeyAscii As Integer)
Dim A As Long
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    If KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End If

End Sub

Private Sub txtMaxLevel_LostFocus()
    txtMaxLevel = Val(txtMaxLevel)
End Sub

Private Sub txtMinLevel_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    If KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub txtMinLevel_LostFocus(Index As Integer)
    txtMinLevel(Index) = Val(txtMinLevel(Index))
End Sub

Private Sub txtRange_KeyPress(KeyAscii As Integer)
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    If KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub txtRange_LostFocus()
    txtRange = Val(txtRange)
End Sub

Private Sub txtRequirementLevel_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    If KeyAscii <> vbKeyBack Then
        KeyAscii = 0
    End If
End If
End Sub

Private Sub txtRequirementLevel_LostFocus(Index As Integer)
    txtRequirementLevel(Index) = Val(txtRequirementLevel(Index))
End Sub
