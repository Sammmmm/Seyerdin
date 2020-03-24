VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seyerdin Options"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar sclTargetFps 
      Height          =   255
      Left            =   1320
      Max             =   200
      TabIndex        =   90
      Top             =   6600
      Value           =   40
      Width           =   2175
   End
   Begin VB.CommandButton btnDefaults 
      Caption         =   "Defaults"
      Height          =   495
      Left            =   5760
      TabIndex        =   88
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CheckBox chkVsync 
      Caption         =   "Vsync Enabled"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   86
      Top             =   5880
      Width           =   2655
   End
   Begin VB.CheckBox chkDisableMulti 
      Caption         =   "Disable MultiSampling"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   85
      Top             =   5520
      Width           =   2535
   End
   Begin VB.HScrollBar sclRes 
      Height          =   255
      Left            =   1320
      Max             =   20
      Min             =   1
      TabIndex        =   83
      Top             =   3480
      Value           =   1
      Width           =   1935
   End
   Begin VB.ComboBox ddlChat 
      Height          =   315
      ItemData        =   "frmOptions.frx":000C
      Left            =   5400
      List            =   "frmOptions.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   74
      Top             =   4320
      Width           =   1335
   End
   Begin VB.ComboBox ddlBroadCast 
      Height          =   315
      ItemData        =   "frmOptions.frx":0010
      Left            =   5400
      List            =   "frmOptions.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   73
      Top             =   4680
      Width           =   1335
   End
   Begin VB.ComboBox ddlTell 
      Height          =   315
      ItemData        =   "frmOptions.frx":0014
      Left            =   5400
      List            =   "frmOptions.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   72
      Top             =   5040
      Width           =   1335
   End
   Begin VB.ComboBox ddlSay 
      Height          =   315
      ItemData        =   "frmOptions.frx":0018
      Left            =   5400
      List            =   "frmOptions.frx":001A
      Style           =   2  'Dropdown List
      TabIndex        =   71
      Top             =   5400
      Width           =   1335
   End
   Begin VB.ComboBox ddlGuild 
      Height          =   315
      ItemData        =   "frmOptions.frx":001C
      Left            =   5400
      List            =   "frmOptions.frx":001E
      Style           =   2  'Dropdown List
      TabIndex        =   70
      Top             =   5760
      Width           =   1335
   End
   Begin VB.ComboBox ddlParty 
      Height          =   315
      ItemData        =   "frmOptions.frx":0020
      Left            =   5400
      List            =   "frmOptions.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   69
      Top             =   6120
      Width           =   1335
   End
   Begin VB.ComboBox ddlCycle 
      Height          =   315
      ItemData        =   "frmOptions.frx":0024
      Left            =   5400
      List            =   "frmOptions.frx":0026
      Style           =   2  'Dropdown List
      TabIndex        =   68
      Top             =   3240
      Width           =   1335
   End
   Begin VB.ComboBox ddlSpell 
      Height          =   315
      Index           =   9
      ItemData        =   "frmOptions.frx":0028
      Left            =   8280
      List            =   "frmOptions.frx":002A
      Style           =   2  'Dropdown List
      TabIndex        =   66
      Top             =   3600
      Width           =   1335
   End
   Begin VB.ComboBox ddlSpell 
      Height          =   315
      Index           =   8
      ItemData        =   "frmOptions.frx":002C
      Left            =   8280
      List            =   "frmOptions.frx":002E
      Style           =   2  'Dropdown List
      TabIndex        =   65
      Top             =   3240
      Width           =   1335
   End
   Begin VB.ComboBox ddlSpell 
      Height          =   315
      Index           =   7
      ItemData        =   "frmOptions.frx":0030
      Left            =   8280
      List            =   "frmOptions.frx":0032
      Style           =   2  'Dropdown List
      TabIndex        =   64
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ComboBox ddlSpell 
      Height          =   315
      Index           =   6
      ItemData        =   "frmOptions.frx":0034
      Left            =   8280
      List            =   "frmOptions.frx":0036
      Style           =   2  'Dropdown List
      TabIndex        =   63
      Top             =   2520
      Width           =   1335
   End
   Begin VB.ComboBox ddlSpell 
      Height          =   315
      Index           =   5
      ItemData        =   "frmOptions.frx":0038
      Left            =   8280
      List            =   "frmOptions.frx":003A
      Style           =   2  'Dropdown List
      TabIndex        =   62
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ComboBox ddlSpell 
      Height          =   315
      Index           =   4
      ItemData        =   "frmOptions.frx":003C
      Left            =   8280
      List            =   "frmOptions.frx":003E
      Style           =   2  'Dropdown List
      TabIndex        =   61
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox ddlSpell 
      Height          =   315
      Index           =   3
      ItemData        =   "frmOptions.frx":0040
      Left            =   8280
      List            =   "frmOptions.frx":0042
      Style           =   2  'Dropdown List
      TabIndex        =   60
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox ddlSpell 
      Height          =   315
      Index           =   2
      ItemData        =   "frmOptions.frx":0044
      Left            =   8280
      List            =   "frmOptions.frx":0046
      Style           =   2  'Dropdown List
      TabIndex        =   59
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ComboBox ddlSpell 
      Height          =   315
      Index           =   1
      ItemData        =   "frmOptions.frx":0048
      Left            =   8280
      List            =   "frmOptions.frx":004A
      Style           =   2  'Dropdown List
      TabIndex        =   58
      Top             =   720
      Width           =   1335
   End
   Begin VB.ComboBox ddlSpell 
      Height          =   315
      Index           =   0
      ItemData        =   "frmOptions.frx":004C
      Left            =   8280
      List            =   "frmOptions.frx":004E
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Top             =   360
      Width           =   1335
   End
   Begin VB.ComboBox ddlStrafe 
      Height          =   315
      ItemData        =   "frmOptions.frx":0050
      Left            =   5400
      List            =   "frmOptions.frx":0052
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Top             =   2880
      Width           =   1335
   End
   Begin VB.ComboBox ddlPickup 
      Height          =   315
      ItemData        =   "frmOptions.frx":0054
      Left            =   5400
      List            =   "frmOptions.frx":0056
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   2520
      Width           =   1335
   End
   Begin VB.ComboBox ddlRun 
      Height          =   315
      ItemData        =   "frmOptions.frx":0058
      Left            =   5400
      List            =   "frmOptions.frx":005A
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ComboBox ddlAttack 
      Height          =   315
      ItemData        =   "frmOptions.frx":005C
      Left            =   5400
      List            =   "frmOptions.frx":005E
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ComboBox ddlRight 
      Height          =   315
      ItemData        =   "frmOptions.frx":0060
      Left            =   5400
      List            =   "frmOptions.frx":0062
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox ddlLeft 
      Height          =   315
      ItemData        =   "frmOptions.frx":0064
      Left            =   5400
      List            =   "frmOptions.frx":0066
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ComboBox ddlDown 
      Height          =   315
      ItemData        =   "frmOptions.frx":0068
      Left            =   5400
      List            =   "frmOptions.frx":006A
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   720
      Width           =   1335
   End
   Begin VB.ComboBox ddlUp 
      Height          =   315
      ItemData        =   "frmOptions.frx":006C
      Left            =   5400
      List            =   "frmOptions.frx":006E
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   360
      Width           =   1335
   End
   Begin VB.CheckBox chkDisplayHelms 
      Caption         =   "Don't Display Helms"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   28
      Top             =   2880
      Width           =   2415
   End
   Begin VB.CheckBox chkFog 
      Caption         =   "Draw Fog"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   20
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CheckBox chkYells 
      Caption         =   "Display &Yells"
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
      Left            =   8280
      TabIndex        =   25
      Top             =   5400
      Width           =   2295
   End
   Begin VB.CheckBox chkEmotes 
      Caption         =   "Display &Emotes"
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
      Left            =   8280
      TabIndex        =   24
      Top             =   5160
      Width           =   2295
   End
   Begin VB.CheckBox chkSays 
      Caption         =   "Display &Says"
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
      Left            =   8280
      TabIndex        =   22
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CheckBox chkTells 
      Caption         =   "Display &Tells"
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
      Left            =   8280
      TabIndex        =   21
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CheckBox chkMidi 
      Caption         =   "Enable Music"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.CheckBox chkHighTask 
      Caption         =   "Windows High Priority"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   6240
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.HScrollBar sclPauseTime 
      Height          =   255
      LargeChange     =   10
      Left            =   2760
      Max             =   1000
      TabIndex        =   14
      Top             =   2400
      Value           =   1
      Width           =   1455
   End
   Begin VB.HScrollBar sclSoundVolume 
      Height          =   255
      Left            =   1800
      Max             =   255
      Min             =   1
      TabIndex        =   12
      Top             =   2040
      Value           =   1
      Width           =   3135
   End
   Begin VB.CheckBox chkWalkSound 
      Caption         =   "Walking Sound"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   480
      Width           =   2415
   End
   Begin VB.HScrollBar sclMusicVolume 
      Height          =   255
      Left            =   1800
      Max             =   255
      Min             =   1
      TabIndex        =   9
      Top             =   1680
      Value           =   1
      Width           =   3135
   End
   Begin VB.CheckBox chkMName 
      Caption         =   "Show Monster Names"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox cmbFontSize 
      Height          =   315
      ItemData        =   "frmOptions.frx":0070
      Left            =   2520
      List            =   "frmOptions.frx":007F
      TabIndex        =   6
      Text            =   "Font Size"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CheckBox chkAway 
      Caption         =   "Away"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.CheckBox chkBroadcasts 
      Caption         =   "Display &Broadcasts"
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
      Left            =   8280
      TabIndex        =   4
      Top             =   4440
      Width           =   2295
   End
   Begin VB.CheckBox chkWAV 
      Caption         =   "Play Sound Effects"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   8880
      TabIndex        =   3
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   6720
      Width           =   1455
   End
   Begin VB.CheckBox chkPriority 
      Caption         =   "Higher CPU usage"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   5160
      Width           =   5055
   End
   Begin VB.CheckBox chkHP 
      Caption         =   "Show your Hp Bar"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   840
      Width           =   5055
   End
   Begin VB.CheckBox chkFullRedraws 
      Caption         =   "Do Full Redraws"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   5520
      Width           =   5055
   End
   Begin VB.CheckBox chkAutoRun 
      Caption         =   "Autorun"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "* Square resolutions offer the best experience.     1200x900 and 1600x1200 are recommended"
      Height          =   375
      Left            =   120
      TabIndex        =   93
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Label Label38 
      Caption         =   "Target FPS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   92
      Top             =   6600
      Width           =   2655
   End
   Begin VB.Label lblTargetFps 
      Caption         =   "40"
      Height          =   375
      Left            =   3600
      TabIndex        =   91
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   5280
      X2              =   5280
      Y1              =   0
      Y2              =   7320
   End
   Begin VB.Label Label5 
      Caption         =   "* Default performance options usually offer the best experience"
      Height          =   375
      Left            =   120
      TabIndex        =   89
      Top             =   7080
      Width           =   5055
   End
   Begin VB.Label Label37 
      Caption         =   "Font Size:"
      Height          =   255
      Left            =   1800
      TabIndex        =   87
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblRes 
      Caption         =   "800x600"
      Height          =   375
      Left            =   3360
      TabIndex        =   84
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label36 
      Caption         =   "Resolution"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   82
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label Label35 
      Caption         =   "Chat"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   81
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label34 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Chat Toggle"
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
      Height          =   375
      Left            =   6720
      TabIndex        =   80
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Label Label33 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "/Broadcast"
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
      Height          =   375
      Left            =   6720
      TabIndex        =   79
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label32 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "/Tell"
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
      Height          =   375
      Left            =   6720
      TabIndex        =   78
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label31 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "/Say"
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
      Height          =   375
      Left            =   6720
      TabIndex        =   77
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label30 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "/Guild Chat"
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
      Height          =   375
      Left            =   6720
      TabIndex        =   76
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "/Party Chat"
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
      Height          =   375
      Left            =   6720
      TabIndex        =   75
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cycle Target"
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
      Height          =   375
      Left            =   6720
      TabIndex        =   67
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label28 
      Caption         =   "Spell Hotkeys"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   57
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label27 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Spell 9"
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
      Height          =   375
      Left            =   9600
      TabIndex        =   55
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Spell 10"
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
      Height          =   375
      Left            =   9600
      TabIndex        =   54
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Spell 1"
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
      Height          =   375
      Left            =   9600
      TabIndex        =   53
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Spell 2"
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
      Height          =   375
      Left            =   9600
      TabIndex        =   52
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label23 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Spell 3"
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
      Height          =   375
      Left            =   9600
      TabIndex        =   51
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Spell 4"
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
      Height          =   375
      Left            =   9600
      TabIndex        =   50
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label21 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Spell 5"
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
      Height          =   375
      Left            =   9600
      TabIndex        =   49
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Spell 6"
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
      Height          =   375
      Left            =   9600
      TabIndex        =   48
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Spell 7"
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
      Height          =   375
      Left            =   9600
      TabIndex        =   47
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Spell 8"
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
      Height          =   375
      Left            =   9600
      TabIndex        =   46
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Strafe"
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
      Height          =   375
      Left            =   6720
      TabIndex        =   44
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Pick Up item"
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
      Height          =   375
      Left            =   6720
      TabIndex        =   43
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Run"
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
      Height          =   375
      Left            =   6720
      TabIndex        =   42
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Attack"
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
      Height          =   375
      Left            =   6720
      TabIndex        =   41
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Right"
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
      Height          =   375
      Left            =   6720
      TabIndex        =   40
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Left"
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
      Height          =   375
      Left            =   6720
      TabIndex        =   39
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   375
      Left            =   6720
      TabIndex        =   38
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   375
      Left            =   6720
      TabIndex        =   37
      Top             =   360
      Width           =   1455
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   5280
      X2              =   0
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label8 
      Caption         =   "Performance Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   27
      Top             =   4800
      Width           =   2715
   End
   Begin VB.Label Label1 
      Caption         =   "Chat Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8280
      TabIndex        =   23
      Top             =   4080
      Width           =   2715
   End
   Begin VB.Label Label4 
      Caption         =   "Pause time on map switch:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label lblPauseTime 
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Sound Volume"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Music Volume"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   56
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frmOptions"
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

Dim currentRes As Long

Private Sub btnCancel_Click()
    LoadOptions
    If blnPlaying = False Then frmMenu.Show
    Unload Me
End Sub


Private Sub btnDefaults_Click()
    LoadOptions (True)
    Form_Load
End Sub

Private Sub btnOk_Click()
Dim tbool As Boolean, tbool2 As Boolean, tbool3 As Boolean, A As Long, b As Long
    With Options
        If chkMidi = 1 Then
            .MIDI = True
        Else
            If .MIDI = True Then
                Sound_StopStream
            End If
            .MIDI = False
        End If
        If chkWAV = 1 Then
            .Wav = True
        Else
            .Wav = False
        End If
        If chkBroadcasts = 1 Then
            .Broadcasts = True
        Else
            .Broadcasts = False
        End If
        If chkEmotes = 1 Then
            .Emotes = True
        Else
            .Emotes = False
        End If
        If chkSays = 1 Then
            .Says = True
        Else
            .Says = False
        End If
        If chkYells = 1 Then
            .Yells = True
        Else
            .Yells = False
        End If
        If chkTells = 1 Then
            .Tells = True
        Else
            .Tells = False
        End If
        If chkDisplayHelms = 0 Then
            .dontDisplayHelms = False
        Else
            .dontDisplayHelms = True
        End If
        
        tbool = .fullredraws
        If chkFullRedraws = 1 Then
            .fullredraws = True
        Else
            .fullredraws = False
        End If

        
        If chkHP = 1 Then
            .ShowHP = True
        Else
            .ShowHP = False
        End If
        If chkAway = 1 Then
            If .ForwardUser <> "" Then
                .ForwardUser = ""
                .Away = True
                MsgBox "Since forwarding was on it is now set to off, away mode is active!", vbCritical + vbOKOnly, "Changing Status"
            Else
                .Away = True
            End If
            PrintChat "Away mode is active, please use /AMSG <message> to set your away message, it is set to default now.", 14, Options.FontSize
        Else
            .Away = False
        End If
        'If chkMName = 1 Then
        '    .MName = True
        'Else
            .MName = False
        'End If
        If chkWalkSound = 1 Then
            .WalkSound = True
        Else
            .WalkSound = False
        End If
        If chkAutoRun = 1 Then
            .AutoRun = True
        Else
            .AutoRun = False
        End If
        If cmbFontSize.ListIndex <> -1 Then
            .FontSize = cmbFontSize.ItemData(cmbFontSize.ListIndex)
        End If
        
   
        If chkPriority = 1 Then
            .highpriority = True
        Else
            .highpriority = False
        End If
        If chkHighTask = 1 Then
            .hightask = True
        Else
            .hightask = False
        End If
        If chkFog = 1 Then
            .ShowFog = True
        Else
            .ShowFog = False
        End If
        
        tbool2 = .DisableMultiSampling
        If chkDisableMulti = 1 Then
            .DisableMultiSampling = True
        Else
            .DisableMultiSampling = False
        End If
        
        tbool3 = .VsyncEnabled
        If chkVsync = 1 Then
            .VsyncEnabled = True
            If frmMain.picMiniMap.Visible = False Then SetMiniMapTab tsButtons
        Else
            .VsyncEnabled = False
        End If

        If .fullredraws <> tbool Or tbool2 <> .DisableMultiSampling Or .ResolutionIndex <> sclRes.Value Or .VsyncEnabled <> tbool3 Then
            Select Case sclRes.Value
                Case 1:
                    currentRes = 200
                    Width = currentRes * 4
                    Height = currentRes * 3
                Case 2:
                    currentRes = 240
                    Width = currentRes * 4
                    Height = currentRes * 3
                Case 3:
                    currentRes = 256
                    Width = currentRes * 4
                    Height = currentRes * 3
                Case 4:
                    currentRes = 300
                    Width = currentRes * 4
                    Height = currentRes * 3
                Case 5:
                    currentRes = 320
                    Width = currentRes * 4
                    Height = currentRes * 3
                Case 6:
                    currentRes = 350
                    Width = currentRes * 4
                    Height = currentRes * 3
                Case 7:
                    currentRes = 360
                    Width = currentRes * 4
                    Height = currentRes * 3
                Case 8:
                    currentRes = 400
                    Width = currentRes * 4
                    Height = currentRes * 3
                Case 9:
                    Width = 1152
                    Height = 864
                Case 10:
                    Width = 1280
                    Height = 720
                Case 11:
                    Width = 1280
                    Height = 768
                Case 12:
                    Width = 1280
                    Height = 800
                Case 13:
                    Width = 1280
                    Height = 1024
                Case 14:
                    Width = 1360
                    Height = 768
                Case 15:
                    Width = 1440
                    Height = 900
                Case 16:
                    Width = 1600
                    Height = 900
                Case 17:
                    Width = 1600
                    Height = 1024
                Case 18:
                    Width = 1680
                    Height = 1050
                Case 19:
                    Width = 1768
                    Height = 992
                Case 20:
                    Width = 1920
                    Height = 1080
            End Select
            
            Dim scry As Long, scrx As Long
            Dim twipsPerpixelX As Long
            Dim twipsPerpixelY As Long
            Dim resX As Long
            Dim resY As Long
        
            scrx = Screen.Width
            scry = Screen.Height
        
            twipsPerpixelX = Screen.twipsPerpixelX
            twipsPerpixelY = Screen.twipsPerpixelY
        
            resX = scrx / twipsPerpixelX
            resY = scry / twipsPerpixelY
            
            If Width > resX Or Height > resY Then
                MsgBox "You have chosen a resolution that is larger than your desktop resolution (including windows scaling factor), please choose another."
            Else
                .ResolutionIndex = sclRes.Value
            End If

            Set mapTexture(1) = Nothing
            Set mapTexture(2) = Nothing
            For A = 1 To NumTextures
                Set DynamicTextures(A).Texture = Nothing
                DynamicTextures(A).Loaded = False
                DynamicTextures(A).LastUsed = 0
            Next A
            A = CurFog
            ClearFog (True)
            If Not InitD3D(False) Then
               
            End If
            If A > 0 And A <= 31 Then
                InitFog A, False
            End If
            InitMiniMap
            ResizeGameWindow
        End If
        
       ' If chkBindingsEnabled = 1 Then
       '     .AltKeysEnabled = True
       '     Chat.Enabled = False
       ' Else
       '     .AltKeysEnabled = False
       '     Chat.Enabled = True
       ' End If
        
        For A = 0 To MAXKEYCODES
            If ddlUp.ListIndex = A Then .UpKey = A
            If ddlDown.ListIndex = A Then .DownKey = A
            If ddlLeft.ListIndex = A Then .LeftKey = A
            If ddlRight.ListIndex = A Then .RightKey = A
            If ddlAttack.ListIndex = A Then .AttackKey = A
            If ddlRun.ListIndex = A Then .RunKey = A
            If ddlStrafe.ListIndex = A Then .StrafeKey = A
            If ddlPickup.ListIndex = A Then .PickupKey = A
            If ddlCycle.ListIndex = A Then .CycleKey = A
            
            If ddlChat.ListIndex = A Then .ChatKey = A
            If ddlBroadCast.ListIndex = A Then .BroadcastKey = A
            If ddlTell.ListIndex = A Then .TellKey = A
            If ddlSay.ListIndex = A Then .SayKey = A
            If ddlGuild.ListIndex = A Then .GuildKey = A
            If ddlParty.ListIndex = A Then .PartyKey = A
            
            For b = 0 To 9
                If ddlSpell(b).ListIndex = A Then .SpellKey(b) = A
            Next b
        Next A
        
        .MusicVolume = sclMusicVolume
        .SoundVolume = sclSoundVolume
        .pausetime = sclPauseTime
        .TargetFps = sclTargetFps
        

        DrawChat
        
        frmMain.Refresh
    
        End With
    SaveOptions
    Unload Me
    If blnPlaying = False Then frmMenu.Show
End Sub


Private Sub chkVsync_Click()
    If chkVsync Then
        sclTargetFps.Enabled = False
        lblTargetFps.Enabled = False
    Else
        sclTargetFps.Enabled = True
        lblTargetFps.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Dim A As Long
Dim b As Long
        
    For A = 0 To MAXKEYCODES
        ddlUp.AddItem KeyCodeList(A).Text
        ddlDown.AddItem KeyCodeList(A).Text
        ddlLeft.AddItem KeyCodeList(A).Text
        ddlRight.AddItem KeyCodeList(A).Text
        ddlAttack.AddItem KeyCodeList(A).Text
        ddlRun.AddItem KeyCodeList(A).Text
        ddlPickup.AddItem KeyCodeList(A).Text
        ddlStrafe.AddItem KeyCodeList(A).Text
        ddlCycle.AddItem KeyCodeList(A).Text
        
        ddlChat.AddItem KeyCodeList(A).Text
        ddlBroadCast.AddItem KeyCodeList(A).Text
        ddlTell.AddItem KeyCodeList(A).Text
        ddlSay.AddItem KeyCodeList(A).Text
        ddlGuild.AddItem KeyCodeList(A).Text
        ddlParty.AddItem KeyCodeList(A).Text
        
        For b = 0 To 9
            ddlSpell(b).AddItem KeyCodeList(A).Text
        Next b
        
        If Options.UpKey = A Then ddlUp.ListIndex = A
        If Options.DownKey = A Then ddlDown.ListIndex = A
        If Options.LeftKey = A Then ddlLeft.ListIndex = A
        If Options.RightKey = A Then ddlRight.ListIndex = A
        
        If Options.AttackKey = A Then ddlAttack.ListIndex = A
        If Options.RunKey = A Then ddlRun.ListIndex = A
        If Options.StrafeKey = A Then ddlStrafe.ListIndex = A
        If Options.PickupKey = A Then ddlPickup.ListIndex = A
        If Options.CycleKey = A Then ddlCycle.ListIndex = A
        
        If Options.ChatKey = A Then ddlChat.ListIndex = A
        If Options.BroadcastKey = A Then ddlBroadCast.ListIndex = A
        If Options.TellKey = A Then ddlTell.ListIndex = A
        If Options.SayKey = A Then ddlSay.ListIndex = A
        If Options.GuildKey = A Then ddlGuild.ListIndex = A
        If Options.PartyKey = A Then ddlParty.ListIndex = A
        
        For b = 0 To 9
            If Options.SpellKey(b) = A Then ddlSpell(b).ListIndex = A
        Next b
        
    Next A

    With Options
        If .MIDI = True Then
            chkMidi = 1
        Else
            chkMidi = 0
        End If
        If .Wav = True Then
            chkWAV = 1
        Else
            chkWAV = 0
        End If
        If .ShowHP Then
            chkHP = 1
        Else
            chkHP = 0
        End If
        
        'If .AltKeysEnabled = True Then
        '    chkBindingsEnabled = 1
        'Else
        '    chkBindingsEnabled = 0
        'End If
        
        If .Broadcasts = True Then
            chkBroadcasts = 1
        Else
            chkBroadcasts = 0
        End If
        If .Tells = True Then
            chkTells = 1
        Else
            chkTells = 0
        End If
        If .Says = True Then
            chkSays = 1
        Else
            chkSays = 0
        End If
        If .Yells = True Then
            chkYells = 1
        Else
            chkYells = 0
        End If
        If .Emotes = True Then
            chkEmotes = 1
        Else
            chkEmotes = 0
        End If
        If .dontDisplayHelms = True Then
            chkDisplayHelms = 1
        Else
            chkDisplayHelms = 0
        End If
        If .Away = True Then
            chkAway = 1
        Else
            chkAway = 0
        End If
        If .MName = True Then
            chkMName = 1
        Else
            chkMName = 0
        End If
        If .WalkSound = True Then
            chkWalkSound = 1
        Else
            chkWalkSound = 0
        End If
        If .AutoRun = True Then
            chkAutoRun = 1
        Else
            chkAutoRun = 0
        End If
        If .fullredraws = True Then
            chkFullRedraws = 1
        Else
            chkFullRedraws = 0
        End If
        If .DisableMultiSampling = True Then
            chkDisableMulti = 1
        Else
            chkDisableMulti = 0
        End If
        
        If .VsyncEnabled = True Then
            chkVsync = 1
        Else
            chkVsync = 0
        End If
        
        'If .PostTells = True Then
        '    chKPostTells = 1
        'Else
        '    chKPostTells = 0
        'End If
        'If .PostSystem = True Then
        '    chkPostSystem = 1
        'Else
        '    chkPostSystem = 0
        'End If
        'If .PostSays = True Then
        '    chkpostSays = 1
        'Else
        '    chkpostSays = 0
        'End If
        'If .DefaultSay = True Then
        '    chkDefaultSay = 1
        'Else
        '    chkDefaultSay = 0
        'End If
        'If .ChannelTags = True Then
        '    chkTags = 1
        'Else
        '    chkTags = 0
        'End If
        If .highpriority Then
            chkPriority = 1
        Else
            chkPriority = 0
        End If
        If .hightask Then
            chkHighTask = 1
        Else
            chkHighTask = 0
        End If
        If .ShowFog Then
            chkFog = 1
        Else
            chkFog = 0
        End If
        If .FontSize < 8 And .FontSize > 12 Then
            cmbFontSize.ListIndex = 0
        Else
            Select Case .FontSize
                Case 8
                    cmbFontSize.ListIndex = 0
                Case 10
                    cmbFontSize.ListIndex = 1
                Case 12
                    cmbFontSize.ListIndex = 2
            End Select
        End If
        If .MusicVolume > 0 Then
            sclMusicVolume = .MusicVolume
        Else
            sclMusicVolume = 64
        End If
        If .SoundVolume > 0 Then
            sclSoundVolume = .SoundVolume
        Else
            sclSoundVolume = 64
        End If

        sclRes.Value = .ResolutionIndex
        sclTargetFps.Value = .TargetFps

        sclPauseTime = .pausetime
        

    End With
    frmOptions_Loaded = True
    Set Me.Icon = frmMenu.Icon
End Sub
Private Sub Form_LostFocus()
    Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmOptions_Loaded = False
End Sub


Private Sub sclMusicVolume_Change()
    Sound_SetStreamVolume sclMusicVolume
End Sub

Private Sub sclMusicVolume_Scroll()
    sclMusicVolume_Change
End Sub

Private Sub sclPauseTime_Change()
lblPauseTime = sclPauseTime
End Sub

Private Sub sclRes_Change()
If sclRes.Value <= 8 Then
    Select Case sclRes.Value
    Case 1:
    currentRes = 200
    Case 2:
    currentRes = 240
    Case 3:
    currentRes = 256
    Case 4:
    currentRes = 300
    Case 5:
    currentRes = 320
    Case 6:
    currentRes = 350
    Case 7:
    currentRes = 360
    Case 8:
    currentRes = 400
    End Select
    lblRes.Caption = currentRes * 4 & "x" & currentRes * 3 & " (4:3 - square) "
Else
    Dim Width As Long, Height As Long

    Select Case sclRes.Value
        Case 9:
            Width = 1152
            Height = 864
        Case 10:
            Width = 1280
            Height = 720
        Case 11:
            Width = 1280
            Height = 768
        Case 12:
            Width = 1280
            Height = 800
        Case 13:
            Width = 1280
            Height = 1024
        Case 14:
            Width = 1360
            Height = 768
        Case 15:
            Width = 1440
            Height = 900
        Case 16:
            Width = 1600
            Height = 900
        Case 17:
            Width = 1600
            Height = 1024
        Case 18:
            Width = 1680
            Height = 1050
        Case 19:
            Width = 1768
            Height = 992
        Case 20:
            Width = 1920
            Height = 1080
    End Select
    lblRes.Caption = Width & "x" & Height
End If
End Sub

Private Sub sclTargetFps_Change()
    If sclTargetFps.Value = 0 Then
        lblTargetFps.Caption = "unlimited"
    Else
        lblTargetFps.Caption = sclTargetFps.Value
    End If
End Sub
