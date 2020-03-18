VERSION 5.00
Begin VB.Form frmMonster 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seyerdin Online [Edit Monster]"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   ControlBox      =   0   'False
   Icon            =   "frmMonster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   472
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFlag2 
      Caption         =   "Nothing"
      Height          =   255
      Index           =   7
      Left            =   6480
      TabIndex        =   104
      Top             =   6720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox chkFlag2 
      Caption         =   "Nothing"
      Height          =   255
      Index           =   6
      Left            =   5280
      TabIndex        =   103
      Top             =   6600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox chkFlag2 
      Caption         =   "Nothing"
      Height          =   255
      Index           =   5
      Left            =   6600
      TabIndex        =   102
      Top             =   6600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox chkFlag2 
      Caption         =   "Nothing"
      Height          =   255
      Index           =   4
      Left            =   7440
      TabIndex        =   101
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox chkFlag2 
      Caption         =   "Nothing"
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   100
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox chkFlag2 
      Caption         =   "Medium Monster"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   99
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CheckBox chkFlag2 
      Caption         =   "Sprite>255"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   98
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CheckBox chkFlag2 
      Caption         =   "Large Monster"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   97
      Top             =   6480
      Width           =   1575
   End
   Begin VB.HScrollBar sclLight 
      Height          =   255
      LargeChange     =   5
      Left            =   5400
      Max             =   50
      TabIndex        =   92
      Top             =   3960
      Width           =   2895
   End
   Begin VB.HScrollBar sclRed 
      Height          =   255
      Left            =   5400
      Max             =   255
      TabIndex        =   84
      Top             =   2760
      Value           =   255
      Width           =   2895
   End
   Begin VB.HScrollBar sclGreen 
      Height          =   255
      Left            =   5400
      Max             =   255
      TabIndex        =   83
      Top             =   3120
      Value           =   255
      Width           =   2895
   End
   Begin VB.HScrollBar sclBlue 
      Height          =   255
      Left            =   5400
      Max             =   255
      TabIndex        =   82
      Top             =   3480
      Value           =   255
      Width           =   2895
   End
   Begin VB.HScrollBar sclAlpha 
      Height          =   255
      Left            =   5400
      Max             =   255
      TabIndex        =   81
      Top             =   2400
      Value           =   255
      Width           =   2895
   End
   Begin VB.HScrollBar sclWander 
      Height          =   255
      LargeChange     =   5
      Left            =   5520
      Max             =   99
      TabIndex        =   78
      Top             =   1200
      Value           =   4
      Width           =   2895
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calc"
      Height          =   375
      Left            =   3840
      TabIndex        =   77
      Top             =   600
      Width           =   615
   End
   Begin VB.HScrollBar sclMoveSpeed 
      Height          =   255
      LargeChange     =   5
      Left            =   5880
      Max             =   16
      TabIndex        =   72
      Top             =   1560
      Value           =   2
      Width           =   2535
   End
   Begin VB.HScrollBar sclAttackSpeed 
      Height          =   255
      LargeChange     =   5
      Left            =   5880
      Max             =   16
      TabIndex        =   71
      Top             =   1920
      Value           =   4
      Width           =   2535
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Monster Tick"
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   70
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "See Invisible"
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   69
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdTestAttack 
      Caption         =   "Test"
      Height          =   255
      Left            =   8520
      TabIndex        =   67
      Top             =   480
      Width           =   495
   End
   Begin VB.HScrollBar sclAttackSound 
      Height          =   255
      Left            =   5880
      Max             =   255
      TabIndex        =   65
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton cmdTestDeath 
      Caption         =   "Test"
      Height          =   255
      Left            =   8520
      TabIndex        =   63
      Top             =   120
      Width           =   495
   End
   Begin VB.HScrollBar sclDeathSound 
      Height          =   255
      Left            =   5880
      Max             =   255
      TabIndex        =   61
      Top             =   120
      Width           =   2175
   End
   Begin VB.HScrollBar sclArmor 
      Height          =   255
      Left            =   960
      Max             =   255
      TabIndex        =   56
      Top             =   2520
      Width           =   2895
   End
   Begin VB.ListBox lstStatusEffect 
      Appearance      =   0  'Flat
      Height          =   1155
      ItemData        =   "frmMonster.frx":000C
      Left            =   5160
      List            =   "frmMonster.frx":0019
      Style           =   1  'Checkbox
      TabIndex        =   54
      Top             =   4800
      Width           =   2055
   End
   Begin VB.HScrollBar sclMagicResist 
      Height          =   255
      LargeChange     =   5
      Left            =   5520
      Max             =   100
      TabIndex        =   51
      Top             =   840
      Width           =   2895
   End
   Begin VB.HScrollBar sclLevel 
      Height          =   255
      LargeChange     =   5
      Left            =   960
      Max             =   255
      TabIndex        =   48
      Top             =   3960
      Value           =   1
      Width           =   2895
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Attack Monsters"
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   47
      Top             =   6120
      Width           =   1575
   End
   Begin VB.TextBox txtChance 
      Height          =   335
      Index           =   2
      Left            =   4440
      MaxLength       =   4
      TabIndex        =   46
      Top             =   5160
      Width           =   450
   End
   Begin VB.TextBox txtChance 
      Height          =   335
      Index           =   1
      Left            =   4440
      MaxLength       =   4
      TabIndex        =   45
      Top             =   4800
      Width           =   450
   End
   Begin VB.TextBox txtChance 
      Height          =   335
      Index           =   0
      Left            =   4440
      MaxLength       =   4
      TabIndex        =   44
      Top             =   4440
      Width           =   450
   End
   Begin VB.HScrollBar sclEXP 
      Height          =   255
      Left            =   960
      Max             =   10000
      TabIndex        =   41
      Top             =   3600
      Value           =   1
      Width           =   2895
   End
   Begin VB.HScrollBar sclMax 
      Height          =   255
      Left            =   960
      Max             =   10000
      Min             =   1
      TabIndex        =   38
      Top             =   2160
      Value           =   1
      Width           =   2895
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Player Friendly"
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   36
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Only comes out at night"
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   35
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Only comes out in day"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   34
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Guard"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   33
      Top             =   5640
      Width           =   1095
   End
   Begin VB.HScrollBar sclAgility 
      Height          =   255
      Left            =   960
      Max             =   100
      TabIndex        =   5
      Top             =   3240
      Width           =   2895
   End
   Begin VB.HScrollBar sclSight 
      Height          =   255
      Left            =   960
      Max             =   255
      Min             =   1
      TabIndex        =   4
      Top             =   2880
      Value           =   1
      Width           =   2895
   End
   Begin VB.TextBox txtValue 
      Height          =   335
      Index           =   2
      Left            =   3480
      MaxLength       =   4
      TabIndex        =   11
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox txtValue 
      Height          =   335
      Index           =   1
      Left            =   3480
      MaxLength       =   4
      TabIndex        =   9
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtValue 
      Height          =   335
      Index           =   0
      Left            =   3480
      MaxLength       =   4
      TabIndex        =   7
      Top             =   4440
      Width           =   855
   End
   Begin VB.ComboBox cmbObject 
      Height          =   315
      Index           =   1
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4800
      Width           =   2295
   End
   Begin VB.ComboBox cmbObject 
      Height          =   315
      Index           =   2
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   5160
      Width           =   2295
   End
   Begin VB.ComboBox cmbObject 
      Height          =   315
      Index           =   0
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4440
      Width           =   2295
   End
   Begin VB.HScrollBar sclMin 
      Height          =   255
      Left            =   960
      Max             =   255
      Min             =   1
      TabIndex        =   3
      Top             =   1800
      Value           =   1
      Width           =   2895
   End
   Begin VB.HScrollBar sclHP 
      Height          =   255
      Left            =   960
      Max             =   32000
      Min             =   1
      TabIndex        =   2
      Top             =   1440
      Value           =   1
      Width           =   2895
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      Height          =   1020
      Left            =   3840
      ScaleHeight     =   128
      ScaleMode       =   0  'User
      ScaleWidth      =   128
      TabIndex        =   16
      Top             =   0
      Width           =   1020
   End
   Begin VB.HScrollBar sclSprite 
      Height          =   255
      Left            =   960
      Max             =   510
      Min             =   1
      TabIndex        =   1
      Top             =   1080
      Value           =   1
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      MaxLength       =   15
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7440
      TabIndex        =   12
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   7440
      TabIndex        =   13
      Top             =   5760
      Width           =   1455
   End
   Begin VB.ListBox lstMonsterType 
      Appearance      =   0  'Flat
      Height          =   1155
      ItemData        =   "frmMonster.frx":0034
      Left            =   5160
      List            =   "frmMonster.frx":003E
      Style           =   1  'Checkbox
      TabIndex        =   59
      Top             =   4800
      Width           =   2055
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "No Shadow"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   68
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Label Label23 
      Caption         =   "Flags:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   96
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label lblRed 
      Alignment       =   2  'Center
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   95
      Top             =   2760
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   296
      X2              =   296
      Y1              =   288
      Y2              =   80
   End
   Begin VB.Label lblLight 
      Alignment       =   2  'Center
      Caption         =   "No"
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
      Left            =   8400
      TabIndex        =   94
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label31 
      Caption         =   "Light?"
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
      Left            =   4560
      TabIndex        =   93
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label30 
      Caption         =   "Red:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   91
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label29 
      Caption         =   "Green:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   90
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label28 
      Caption         =   "Blue:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   89
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label lblGreen 
      Alignment       =   2  'Center
      Caption         =   "255"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   88
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label lblBlue 
      Alignment       =   2  'Center
      Caption         =   "255"
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
      Left            =   8400
      TabIndex        =   87
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label24 
      Caption         =   "Alpha:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   86
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblAlpha 
      Alignment       =   2  'Center
      Caption         =   "255"
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
      Left            =   8400
      TabIndex        =   85
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lblWander 
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
      Left            =   8520
      TabIndex        =   80
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label20 
      Caption         =   "Wander:"
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
      Left            =   4560
      TabIndex        =   79
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblMoveSpeed 
      Alignment       =   2  'Center
      Caption         =   "1"
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
      Left            =   8520
      TabIndex        =   76
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label22 
      Caption         =   "MoveSpeed:"
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
      Left            =   4560
      TabIndex        =   75
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label21 
      Caption         =   "AttackSpeed:"
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
      Left            =   4560
      TabIndex        =   74
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblAttackSpeed 
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
      Left            =   8520
      TabIndex        =   73
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblAttackSound 
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
      Left            =   8040
      TabIndex        =   66
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label15 
      Caption         =   "Attack Sound:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   64
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblDeathSound 
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
      Left            =   8040
      TabIndex        =   62
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Death Sound:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   60
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblMonsterType 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Type"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6360
      TabIndex        =   58
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label lblStatusEffects 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status Effects"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   5160
      TabIndex        =   57
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label19 
      Caption         =   "Cause Status Effects:"
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
      TabIndex        =   55
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label lblMagicResist 
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
      Left            =   8520
      TabIndex        =   53
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label17 
      Caption         =   "MResist:"
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
      Left            =   4560
      TabIndex        =   52
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label18 
      Caption         =   "Level:"
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
      Left            =   120
      TabIndex        =   50
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      Caption         =   "1"
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
      Left            =   3960
      TabIndex        =   49
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label lblEXP 
      Alignment       =   2  'Center
      Caption         =   "1"
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
      Left            =   3960
      TabIndex        =   43
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label16 
      Caption         =   "Exp:"
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
      Left            =   120
      TabIndex        =   42
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label lblMax 
      Alignment       =   2  'Center
      Caption         =   "1"
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
      Left            =   3960
      TabIndex        =   40
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label14 
      Caption         =   "Max:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblSprite 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Left            =   3960
      TabIndex        =   37
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label10 
      Caption         =   "Flags 2:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "Agility"
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
      Left            =   120
      TabIndex        =   31
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label lblAgility 
      Alignment       =   2  'Center
      Caption         =   "1"
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
      Left            =   3960
      TabIndex        =   30
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "Sight:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblSight 
      Alignment       =   2  'Center
      Caption         =   "1"
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
      Left            =   3960
      TabIndex        =   28
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "Armor:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   27
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      Caption         =   "1"
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
      Left            =   3960
      TabIndex        =   26
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label9 
      Caption         =   "Object 3:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Object 2:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Object 1:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label lblMin 
      Alignment       =   2  'Center
      Caption         =   "1"
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
      Left            =   3960
      TabIndex        =   22
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblHP 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   21
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Min:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "HP:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblNumber 
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
      Height          =   375
      Left            =   1080
      TabIndex        =   18
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "Number:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Sprite:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmMonster"
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
End Sub

Private Sub btnOk_Click()
    Dim V1 As Long, V2 As Long, V3 As Long
    Dim C1 As Byte, C2 As Byte, C3 As Byte
    Dim A As Long, Flags As Byte, Flags2 As Byte, StatusEffect As Long, MonsterType As Long
        
    For A = 0 To 7
        If chkFlag(A) = 1 Then
            SetBit Flags, CByte(A)
        Else
            ClearBit Flags, CByte(A)
        End If
    Next A
    For A = 0 To 7
        If chkFlag2(A) = 1 Then
            SetBit Flags2, CByte(A)
        Else
            ClearBit Flags2, CByte(A)
        End If
    Next A
    
    
    
    For A = 0 To 2
        If lstStatusEffect.Selected(A) = True Then
            SetBit StatusEffect, A
        Else
            ClearBit StatusEffect, A
        End If
    Next A
    
    For A = 0 To 1
        If lstMonsterType.Selected(A) = True Then
            SetBit MonsterType, A
        Else
            ClearBit MonsterType, A
        End If
    Next A
    
    V1 = Val(txtValue(0))
    If V1 < 0 Then V1 = 0
    V2 = Val(txtValue(1))
    If V2 < 0 Then V2 = 0
    V3 = Val(txtValue(2))
    If V3 < 0 Then V3 = 0
    For A = 0 To 2
        If Val(txtChance(A)) > 255 Then txtChance(A) = 255
        If Val(txtChance(A)) < 0 Then txtChance(A) = 0
    Next A
    C1 = CByte(Val(txtChance(0)))
    C2 = CByte(Val(txtChance(1)))
    C3 = CByte(Val(txtChance(2)))
    
    If chkFlag2(1) Then
        SendSocket Chr$(22) + DoubleChar(lblNumber) + Chr$(sclSprite - 255) + DoubleChar(sclHP) + Chr$(sclMin) + DoubleChar(sclMax) + Chr$(sclArmor) + Chr$(sclSight) + Chr$(sclAgility) + Chr$(Flags) + DoubleChar(cmbObject(0).ListIndex) + Chr$(V1) + DoubleChar(cmbObject(1).ListIndex) + Chr$(V2) + DoubleChar(cmbObject(2).ListIndex) + Chr$(V3) + DoubleChar(sclEXP) + Chr$(C1) + Chr$(C2) + Chr$(C3) + Chr$(sclLevel) + Chr$(sclMagicResist) + QuadChar(StatusEffect) + QuadChar(MonsterType) + Chr$(sclDeathSound) + Chr$(sclAttackSound) + Chr$(sclMoveSpeed) + Chr$(sclAttackSpeed) + Chr$(sclWander) + Chr$(sclAlpha) + Chr$(sclRed) + Chr$(sclGreen) + Chr$(sclBlue) + Chr$(sclLight) + Chr$(Flags2) + txtName
    Else
        SendSocket Chr$(22) + DoubleChar(lblNumber) + Chr$(sclSprite) + DoubleChar(sclHP) + Chr$(sclMin) + DoubleChar(sclMax) + Chr$(sclArmor) + Chr$(sclSight) + Chr$(sclAgility) + Chr$(Flags) + DoubleChar(cmbObject(0).ListIndex) + Chr$(V1) + DoubleChar(cmbObject(1).ListIndex) + Chr$(V2) + DoubleChar(cmbObject(2).ListIndex) + Chr$(V3) + DoubleChar(sclEXP) + Chr$(C1) + Chr$(C2) + Chr$(C3) + Chr$(sclLevel) + Chr$(sclMagicResist) + QuadChar(StatusEffect) + QuadChar(MonsterType) + Chr$(sclDeathSound) + Chr$(sclAttackSound) + Chr$(sclMoveSpeed) + Chr$(sclAttackSpeed) + Chr$(sclWander) + Chr$(sclAlpha) + Chr$(sclRed) + Chr$(sclGreen) + Chr$(sclBlue) + Chr$(sclLight) + Chr$(Flags2) + txtName
    End If
    Me.Hide
End Sub

Private Sub chkFlag2_Click(Index As Integer)
If chkFlag2(0) Then sclSprite.max = 100
If Not chkFlag2(0) Then sclSprite.max = 510


End Sub

Private Sub cmdCalculate_Click()
    sclHP = sclLevel * 8
    sclMin = sclLevel * 1.3
    sclMax = sclLevel * 1.3
    sclEXP = sclLevel * 25
    sclArmor = sclLevel * 0.4
End Sub

Private Sub cmdTestAttack_Click()
    PlayWav sclAttackSound
End Sub

Private Sub cmdTestDeath_Click()
    PlayWav sclDeathSound
End Sub

Private Sub Form_Load()
    Dim A As Long
    cmbObject(0).AddItem "<None>"
    cmbObject(1).AddItem "<None>"
    cmbObject(2).AddItem "<None>"
    For A = 1 To MAXITEMS
        cmbObject(0).AddItem CStr(A) + ": " + Object(A).Name
        cmbObject(1).AddItem CStr(A) + ": " + Object(A).Name
        cmbObject(2).AddItem CStr(A) + ": " + Object(A).Name
    Next A
    frmMonster_Loaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMonster_Loaded = False
End Sub

Private Sub lblMonsterType_Click()
lstStatusEffect.Visible = False
lstMonsterType.Visible = True
End Sub

Private Sub lblStatusEffects_Click()
lstStatusEffect.Visible = True
lstMonsterType.Visible = False
End Sub

Private Sub sclAgility_Change()
    lblAgility = sclAgility
End Sub

Private Sub sclAgility_Scroll()
    sclAgility_Change
End Sub


Private Sub sclAlpha_Change()
    lblAlpha = sclAlpha
End Sub

Private Sub sclArmor_Change()
    lblArmor = sclArmor
End Sub

Private Sub sclArmor_Scroll()
    sclArmor_Change
End Sub

Private Sub sclAttackSound_Change()
    lblAttackSound = sclAttackSound
End Sub

Private Sub sclAttackSound_Scroll()
    sclAttackSound_Change
End Sub

Private Sub sclAttackSpeed_Change()
    lblAttackSpeed = sclAttackSpeed
End Sub

Private Sub sclAttackSpeed_Scroll()
    sclAttackSpeed_Change
End Sub

Private Sub sclBlue_Change()
    lblBlue = sclBlue
End Sub

Private Sub sclBlue_Scroll()
    sclBlue_Change
End Sub

Private Sub sclDeathSound_Change()
    lblDeathSound = sclDeathSound
End Sub

Private Sub sclDeathSound_Scroll()
    sclDeathSound_Change
End Sub

Private Sub sclGreen_Change()
    lblGreen = sclGreen
End Sub

Private Sub sclGreen_Scroll()
    sclGreen_Change
End Sub

Private Sub sclLight_Change()
    If sclLight = 0 Then
        lblLight = "No"
    Else
        lblLight = sclLight
    End If
End Sub

Private Sub sclLight_Scroll()
    sclLight_Change
End Sub

Private Sub sclMin_Change()
    lblMin = sclMin
End Sub

Private Sub sclMin_Scroll()
    lblMin = sclMin
End Sub

Private Sub sclEXP_Change()
    lblEXP = sclEXP
End Sub

Private Sub sclEXP_Scroll()
    lblEXP = sclEXP
End Sub

Private Sub sclHP_Change()
    lblHP = sclHP
End Sub

Private Sub sclHP_Scroll()
    sclHP_Change
End Sub

Private Sub sclLevel_Change()
lblLevel = sclLevel
End Sub

Private Sub sclLevel_Scroll()
sclLevel_Change
End Sub

Private Sub sclMagicResist_Change()
lblMagicResist = sclMagicResist
End Sub

Private Sub sclMagicResist_Scroll()
sclMagicResist_Change
End Sub

Private Sub sclMax_Change()
lblMax = sclMax
End Sub

Private Sub sclMax_Scroll()
lblMax = sclMax
End Sub

Private Sub sclMoveSpeed_Change()
    lblMoveSpeed = sclMoveSpeed
End Sub

Private Sub sclMoveSpeed_Scroll()
    sclMoveSpeed_Change
End Sub

Private Sub sclRed_Change()
    lblRed = sclRed
End Sub

Private Sub sclRed_Scroll()
    sclRed_Change
End Sub

Private Sub sclSight_Change()
    lblSight = sclSight
End Sub

Private Sub sclSight_Scroll()
    sclSight_Change
End Sub

Private Sub sclSprite_Change()
       
        lblSprite.Caption = sclSprite
        picSprite.Height = 36
        picSprite.Width = 36
        Dim R1 As RECT, r2 As RECT, A As Long
        If chkFlag2(0) Or chkFlag2(2) Then
                picSprite.Height = 68
                picSprite.Width = 68
                A = sclSprite - 1
                R1.Top = 0: R1.Bottom = 64: R1.Left = 0: R1.Right = 64
                r2.Top = (A \ 16) * 64: r2.Left = (A Mod 16) * 64: r2.Right = r2.Left + 64: r2.Bottom = r2.Top + 64
                On Error Resume Next
                    sfcLSprites.Surface.BltToDC picSprite.hdc, r2, R1
                On Error GoTo 0
        Else
            If sclSprite > 255 Then
                A = sclSprite - 255 - 1
                R1.Top = 0: R1.Bottom = 32: R1.Left = 0: R1.Right = 32
                r2.Top = (A \ 16) * 32: r2.Left = (A Mod 16) * 32: r2.Right = r2.Left + 32: r2.Bottom = r2.Top + 32
                On Error Resume Next
                    sfcSprites2.Surface.BltToDC picSprite.hdc, r2, R1
                On Error GoTo 0
            Else
                A = sclSprite - 1
                R1.Top = 0: R1.Bottom = 32: R1.Left = 0: R1.Right = 32
                r2.Top = (A \ 16) * 32: r2.Left = (A Mod 16) * 32: r2.Right = r2.Left + 32: r2.Bottom = r2.Top + 32
                On Error Resume Next
                    sfcSprites.Surface.BltToDC picSprite.hdc, r2, R1
                On Error GoTo 0
            End If
        End If
    picSprite.Refresh
    
        If chkFlag2(0) Or chkFlag2(2) Then
            chkFlag2(1) = 0
            sclSprite.max = 100
        Else
        If sclSprite > 255 Then
            chkFlag2(1) = 1
        Else
            chkFlag2(1) = 0
        End If
        End If
    
End Sub

Private Sub sclSprite_Scroll()
    sclSprite_Change
End Sub

Private Sub sclWander_Change()
    lblWander = sclWander
End Sub

Private Sub sclWander_Scroll()
    sclWander_Change
End Sub
