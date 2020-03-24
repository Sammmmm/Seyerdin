VERSION 5.00
Begin VB.Form frmMapAtt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seyerdin Online [Map Attribute]"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   ControlBox      =   0   'False
   Icon            =   "frmMapAtt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox picAtt8 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chkAtt8 
         Caption         =   "Opens Att Layer"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   83
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkAtt8 
         Caption         =   "Opens Wall Layer"
         Height          =   375
         Index           =   0
         Left            =   2160
         TabIndex        =   82
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.HScrollBar sclAtt8Hall 
         Height          =   255
         LargeChange     =   10
         Left            =   720
         Max             =   255
         TabIndex        =   35
         Top             =   840
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt8X 
         Height          =   255
         Left            =   720
         Max             =   11
         TabIndex        =   30
         Top             =   120
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt8Y 
         Height          =   255
         Left            =   720
         Max             =   11
         TabIndex        =   29
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label lblAtt8Hall 
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
         Left            =   3240
         TabIndex        =   37
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "Hall:"
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
         TabIndex        =   36
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblAtt8X 
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
         Left            =   3240
         TabIndex        =   34
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "X:"
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
         TabIndex        =   33
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblAtt8Y 
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
         Left            =   3240
         TabIndex        =   32
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Y:"
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
         TabIndex        =   31
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   2760
      Width           =   1455
   End
   Begin VB.PictureBox picAtt19 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   85
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chkAtt19Clickable 
         Caption         =   "Clickable"
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
         TabIndex        =   122
         Top             =   1125
         Width           =   1215
      End
      Begin VB.HScrollBar sclAtt19Frame 
         Height          =   255
         Left            =   840
         Max             =   11
         TabIndex        =   121
         Top             =   840
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt19Source 
         Height          =   255
         Left            =   1080
         Max             =   2
         TabIndex        =   118
         Top             =   120
         Width           =   1815
      End
      Begin VB.PictureBox picAtt19Object 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3480
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   112
         Top             =   840
         Width           =   480
      End
      Begin VB.HScrollBar sclAtt19 
         Height          =   255
         Left            =   840
         Max             =   512
         Min             =   1
         TabIndex        =   86
         Top             =   480
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label Label19 
         Caption         =   "Frame:"
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
         TabIndex        =   120
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblAtt19Source 
         Alignment       =   2  'Center
         Caption         =   "Objects"
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
         Left            =   3000
         TabIndex        =   119
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Source:"
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
         TabIndex        =   117
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblAtt19 
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
         Left            =   3360
         TabIndex        =   88
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label20 
         Caption         =   "Num:"
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
         TabIndex        =   87
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.PictureBox picAtt18 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   67
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chkAtt18 
         Caption         =   "Background checked boxes become invisible"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   72
         Top             =   1200
         Width           =   3855
      End
      Begin VB.CheckBox chkAtt18 
         Caption         =   "Background"
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   71
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkAtt18 
         Caption         =   "Background"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   70
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkAtt18 
         Caption         =   "Background"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   69
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkAtt18 
         Caption         =   "Background"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   68
         Top             =   240
         Width           =   1215
      End
      Begin VB.Line Line6 
         X1              =   3360
         X2              =   3360
         Y1              =   120
         Y2              =   1080
      End
      Begin VB.Line Line5 
         X1              =   480
         X2              =   480
         Y1              =   120
         Y2              =   1080
      End
      Begin VB.Line Line4 
         X1              =   480
         X2              =   3360
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line3 
         X1              =   480
         X2              =   3360
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line2 
         X1              =   480
         X2              =   3360
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line1 
         X1              =   1920
         X2              =   1920
         Y1              =   120
         Y2              =   1080
      End
   End
   Begin VB.PictureBox picAtt15 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   49
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclAtt15Chest 
         Height          =   255
         LargeChange     =   25
         Left            =   840
         Max             =   255
         Min             =   1
         TabIndex        =   50
         Top             =   720
         Value           =   1
         Width           =   2415
      End
      Begin VB.CheckBox chkAtt15 
         Caption         =   "Personal Storage"
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
         TabIndex        =   53
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblAtt15Chest 
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
         Left            =   3360
         TabIndex        =   52
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Chest:"
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
         TabIndex        =   51
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.PictureBox picAtt14 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   63
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chkAtt14Flag 
         Caption         =   "Attackable"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   66
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chkAtt14Flag 
         Caption         =   "Clickable"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   65
         Top             =   360
         Width           =   1575
      End
      Begin VB.CheckBox chkAtt14Flag 
         Caption         =   "Blocked"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   64
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.PictureBox picAtt7 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   21
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclAtt7Mag 
         Height          =   255
         Left            =   720
         Max             =   25
         TabIndex        =   123
         Top             =   1080
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt7Val 
         Height          =   255
         Left            =   720
         TabIndex        =   25
         Top             =   480
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt7Obj 
         Height          =   255
         Left            =   720
         Max             =   1000
         Min             =   1
         TabIndex        =   22
         Top             =   120
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblAtt7Mag 
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
         Left            =   3240
         TabIndex        =   125
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Mag:"
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
         TabIndex        =   124
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label6 
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
         TabIndex        =   43
         Top             =   780
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Val:"
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
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblAtt7Val 
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
         Left            =   3240
         TabIndex        =   26
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Obj:"
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
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblAtt7Obj 
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
         Left            =   3240
         TabIndex        =   23
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblAtt7Name 
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
         Left            =   840
         TabIndex        =   44
         Top             =   780
         Width           =   3135
      End
   End
   Begin VB.PictureBox picAtt1 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   73
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chkAtt1 
         Caption         =   "Right Out"
         Height          =   195
         Index           =   7
         Left            =   2880
         TabIndex        =   81
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "Left Out"
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   80
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "Down Out"
         Height          =   195
         Index           =   5
         Left            =   1440
         TabIndex        =   79
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "Up Out"
         Height          =   195
         Index           =   4
         Left            =   1440
         TabIndex        =   78
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "Right In"
         Height          =   195
         Index           =   3
         Left            =   2880
         TabIndex        =   77
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "Left In"
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   76
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "Down In"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   75
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox chkAtt1 
         Caption         =   "Up In"
         Height          =   195
         Index           =   0
         Left            =   1440
         TabIndex        =   74
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox picAtt9 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   38
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclAtt9Damage 
         Height          =   255
         Left            =   1080
         Max             =   50
         Min             =   1
         TabIndex        =   39
         Top             =   120
         Value           =   1
         Width           =   2295
      End
      Begin VB.Label Label14 
         Caption         =   "Damage:"
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
         TabIndex        =   41
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblAtt9Damage 
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
         Left            =   3480
         TabIndex        =   40
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox picAtt3 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclAtt3Str 
         Height          =   255
         Left            =   960
         Max             =   100
         Min             =   1
         TabIndex        =   46
         Top             =   840
         Value           =   1
         Width           =   2415
      End
      Begin VB.CheckBox optAtt3Pick 
         Caption         =   "Pickable"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   480
         Width           =   1455
      End
      Begin VB.HScrollBar sclAtt3Key 
         Height          =   255
         Left            =   720
         Max             =   400
         Min             =   1
         TabIndex        =   14
         Top             =   120
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label lblAtt3Str 
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
         Left            =   3360
         TabIndex        =   48
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Strength:"
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
         TabIndex        =   47
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Key:"
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
         TabIndex        =   16
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblAtt3Key 
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
         Left            =   3240
         TabIndex        =   15
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox picAtt4 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   84
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.PictureBox picAtt2 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclAtt2Y 
         Height          =   255
         Left            =   720
         Max             =   11
         TabIndex        =   9
         Top             =   840
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt2X 
         Height          =   255
         Left            =   720
         Max             =   11
         TabIndex        =   8
         Top             =   480
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt2Map 
         Height          =   255
         LargeChange     =   25
         Left            =   720
         Max             =   5000
         Min             =   1
         TabIndex        =   7
         Top             =   120
         Value           =   1
         Width           =   2415
      End
      Begin VB.TextBox txtAtt2Map 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         MaxLength       =   4
         TabIndex        =   42
         Text            =   "0"
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label lblAtt2Map 
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
         Left            =   3240
         TabIndex        =   10
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblAtt2Y 
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
         Left            =   3240
         TabIndex        =   12
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblAtt2X 
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
         Left            =   3240
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Y:"
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
         TabIndex        =   6
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "X:"
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
         TabIndex        =   5
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Map:"
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
         TabIndex        =   4
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox picAtt6 
      Height          =   2055
      Left            =   120
      ScaleHeight     =   1995
      ScaleWidth      =   4035
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chkClickable 
         Caption         =   "Clickable"
         Height          =   375
         Left            =   480
         TabIndex        =   168
         Top             =   1200
         Width           =   1335
      End
      Begin VB.HScrollBar sclAtt6Data 
         Height          =   255
         Index           =   1
         Left            =   720
         Max             =   255
         TabIndex        =   141
         Top             =   790
         Value           =   1
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt6Data 
         Height          =   255
         Index           =   0
         Left            =   720
         Max             =   255
         TabIndex        =   138
         Top             =   435
         Value           =   1
         Width           =   2415
      End
      Begin VB.HScrollBar sclAtt6Num 
         Height          =   255
         Left            =   720
         Max             =   255
         Min             =   1
         TabIndex        =   18
         Top             =   75
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label Label25 
         Caption         =   "Data2:"
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
         TabIndex        =   143
         Top             =   795
         Width           =   615
      End
      Begin VB.Label lblAtt6Data 
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
         Index           =   1
         Left            =   3240
         TabIndex        =   142
         Top             =   790
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "Data1:"
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
         TabIndex        =   140
         Top             =   435
         Width           =   615
      End
      Begin VB.Label lblAtt6Data 
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
         Index           =   0
         Left            =   3240
         TabIndex        =   139
         Top             =   435
         Width           =   735
      End
      Begin VB.Label lblAtt6Num 
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
         Left            =   3240
         TabIndex        =   20
         Top             =   75
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Num:"
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
         Top             =   75
         Width           =   495
      End
   End
   Begin VB.PictureBox picAtt26 
      Height          =   2175
      Left            =   120
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   155
      Top             =   480
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclHalfPlayer 
         Height          =   255
         Left            =   1320
         Max             =   32
         TabIndex        =   166
         Top             =   1680
         Width           =   2055
      End
      Begin VB.HScrollBar sclRunSpeed 
         Height          =   255
         Left            =   1320
         Max             =   31
         Min             =   1
         TabIndex        =   164
         Top             =   480
         Value           =   16
         Width           =   2055
      End
      Begin VB.HScrollBar sclWalkSpeed 
         Height          =   255
         Left            =   1320
         Max             =   31
         Min             =   1
         TabIndex        =   163
         Top             =   120
         Value           =   8
         Width           =   2055
      End
      Begin VB.CheckBox chkNoRunEnergy 
         Caption         =   "-1 Run Energy"
         Height          =   255
         Left            =   120
         TabIndex        =   158
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox chk1Energy 
         Caption         =   "Use +1 Energy"
         Height          =   255
         Left            =   120
         TabIndex        =   157
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox chk2Energy 
         Caption         =   "Use +2 Energy"
         Height          =   255
         Left            =   120
         TabIndex        =   156
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label lblAtt24 
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
         Index           =   4
         Left            =   3240
         TabIndex        =   167
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label33 
         Caption         =   "Half Player:"
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
         TabIndex        =   165
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label lblAtt24 
         Alignment       =   2  'Center
         Caption         =   "8"
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
         Index           =   3
         Left            =   3360
         TabIndex        =   162
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label32 
         Caption         =   "Walk Speed:"
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
         TabIndex        =   161
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label lblAtt24 
         Alignment       =   2  'Center
         Caption         =   "16"
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
         Index           =   2
         Left            =   3360
         TabIndex        =   160
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label30 
         Caption         =   "Run Speed:"
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
         TabIndex        =   159
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.PictureBox picAtt24 
      Height          =   2175
      Left            =   120
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   113
      Top             =   480
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chkZeroTall 
         Caption         =   "Delete Y FG Tile"
         Height          =   255
         Left            =   2280
         TabIndex        =   154
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CheckBox chkBranchRight 
         Caption         =   "Branches also occupy right"
         Height          =   255
         Left            =   120
         TabIndex        =   153
         Top             =   1800
         Width           =   2415
      End
      Begin VB.CheckBox chkBranchLeft 
         Caption         =   "branches also occupy left"
         Height          =   255
         Left            =   120
         TabIndex        =   152
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CheckBox chkTwoTall 
         Caption         =   "Delete Y-2 FG tile"
         Height          =   255
         Left            =   2280
         TabIndex        =   148
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CheckBox chkOneTall 
         Caption         =   "Delete Y-1 FG Tile"
         Height          =   375
         Left            =   2280
         TabIndex        =   151
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox chkTreeRight 
         Caption         =   "stump also occupies right"
         Height          =   255
         Left            =   120
         TabIndex        =   150
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CheckBox chkTreeLeft 
         Caption         =   "stump also occupies left"
         Height          =   255
         Left            =   120
         TabIndex        =   149
         Top             =   1080
         Width           =   2055
      End
      Begin VB.CheckBox chkDisposable 
         Caption         =   "Permanant Dispose"
         Height          =   255
         Left            =   120
         TabIndex        =   147
         Top             =   840
         Width           =   2055
      End
      Begin VB.HScrollBar sclAtt24 
         Height          =   255
         Index           =   1
         Left            =   1320
         Max             =   255
         TabIndex        =   145
         Top             =   480
         Value           =   1
         Width           =   2175
      End
      Begin VB.HScrollBar sclAtt24 
         Height          =   255
         Index           =   0
         Left            =   1320
         Max             =   255
         Min             =   1
         TabIndex        =   114
         Top             =   120
         Value           =   1
         Width           =   2175
      End
      Begin VB.Label Label31 
         Caption         =   "Max Ore:"
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
         TabIndex        =   146
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label lblAtt24 
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
         Index           =   1
         Left            =   3360
         TabIndex        =   144
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "Mine Type:"
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
         TabIndex        =   116
         Top             =   120
         Width           =   2655
      End
      Begin VB.Label lblAtt24 
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
         Index           =   0
         Left            =   3360
         TabIndex        =   115
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox picAtt22 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   93
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclAtt22Light 
         Height          =   255
         Left            =   840
         Max             =   50
         TabIndex        =   135
         Top             =   1110
         Width           =   1455
      End
      Begin VB.HScrollBar sclAtt22 
         Height          =   255
         Index           =   2
         Left            =   840
         Max             =   31
         TabIndex        =   126
         Top             =   750
         Width           =   2655
      End
      Begin VB.HScrollBar sclAtt22 
         Height          =   255
         Index           =   1
         Left            =   840
         Max             =   255
         Min             =   4
         TabIndex        =   97
         Top             =   390
         Value           =   4
         Width           =   2655
      End
      Begin VB.HScrollBar sclAtt22 
         Height          =   255
         Index           =   0
         Left            =   840
         Max             =   255
         TabIndex        =   94
         Top             =   30
         Width           =   2655
      End
      Begin VB.Label lblAtt22Light 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   2400
         TabIndex        =   137
         Top             =   1110
         Width           =   1575
      End
      Begin VB.Label Label24 
         Caption         =   "Light:"
         Height          =   255
         Left            =   120
         TabIndex        =   136
         Top             =   1110
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "Flicker:"
         Height          =   255
         Left            =   120
         TabIndex        =   128
         Top             =   750
         Width           =   615
      End
      Begin VB.Label lblAtt22 
         Caption         =   "0"
         Height          =   255
         Index           =   2
         Left            =   3600
         TabIndex        =   127
         Top             =   750
         Width           =   375
      End
      Begin VB.Label lblAtt22 
         Caption         =   "0"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   99
         Top             =   390
         Width           =   375
      End
      Begin VB.Label Label22 
         Caption         =   "Radius:"
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   390
         Width           =   615
      End
      Begin VB.Label lblAtt22 
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   96
         Top             =   30
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "Intensity:"
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   30
         Width           =   615
      End
   End
   Begin VB.PictureBox picAtt25 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   129
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclAtt25NPC 
         Height          =   255
         Left            =   840
         Max             =   255
         Min             =   1
         TabIndex        =   130
         Top             =   360
         Value           =   1
         Width           =   2415
      End
      Begin VB.Label LblAtt25Name 
         Alignment       =   2  'Center
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
         Left            =   840
         TabIndex        =   134
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label lblAtt25NPC 
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
         Left            =   3240
         TabIndex        =   133
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label29 
         Caption         =   "NPC:"
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
         TabIndex        =   132
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label26 
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
         TabIndex        =   131
         Top             =   780
         Width           =   735
      End
   End
   Begin VB.PictureBox picAtt17 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   54
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chkAtt17Blocked 
         Caption         =   "Right Out"
         Height          =   195
         Index           =   7
         Left            =   3000
         TabIndex        =   62
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox chkAtt17Blocked 
         Caption         =   "Right In"
         Height          =   195
         Index           =   6
         Left            =   3000
         TabIndex        =   61
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox chkAtt17Blocked 
         Caption         =   "Left Out"
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   60
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox chkAtt17Blocked 
         Caption         =   "Left In"
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   59
         Top             =   480
         Width           =   1815
      End
      Begin VB.CheckBox chkAtt17Blocked 
         Caption         =   "Down In"
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   57
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox chkAtt17Blocked 
         Caption         =   "Up Out"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   56
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chkAtt17Blocked 
         Caption         =   "Up In"
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   55
         Top             =   0
         Width           =   1815
      End
      Begin VB.CheckBox chkAtt17Blocked 
         Caption         =   "Down Out"
         Height          =   195
         Index           =   3
         Left            =   1680
         TabIndex        =   58
         Top             =   1200
         Width           =   1815
      End
   End
   Begin VB.PictureBox picAtt23 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   100
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CheckBox chkAtt23 
         Caption         =   "FGTile"
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   111
         Top             =   120
         Width           =   975
      End
      Begin VB.CheckBox chkAtt23 
         Caption         =   "BGTile2"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   110
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox chkAtt23 
         Caption         =   "BGTile1"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   109
         Top             =   120
         Width           =   975
      End
      Begin VB.CheckBox chkAtt23 
         Caption         =   "Ground2"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   108
         Top             =   360
         Width           =   1095
      End
      Begin VB.CheckBox chkAtt23 
         Caption         =   "Ground"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   107
         Top             =   120
         Width           =   1095
      End
      Begin VB.HScrollBar sclAtt23 
         Height          =   255
         Index           =   0
         Left            =   840
         Max             =   255
         Min             =   -255
         TabIndex        =   106
         Top             =   720
         Width           =   2655
      End
      Begin VB.HScrollBar sclAtt23 
         Height          =   255
         Index           =   1
         Left            =   840
         Max             =   255
         Min             =   -255
         TabIndex        =   101
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label lblAtt23 
         Caption         =   "0"
         Height          =   255
         Index           =   0
         Left            =   3600
         TabIndex        =   105
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblAtt23 
         Caption         =   "X:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   104
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblAtt23 
         Caption         =   "Y:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   103
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblAtt23 
         Caption         =   "0"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   102
         Top             =   1080
         Width           =   375
      End
   End
   Begin VB.PictureBox picAtt20 
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4035
      TabIndex        =   89
      Top             =   600
      Visible         =   0   'False
      Width           =   4095
      Begin VB.HScrollBar sclAtt20Height 
         Height          =   255
         Left            =   840
         Max             =   31
         Min             =   1
         TabIndex        =   90
         Top             =   600
         Value           =   16
         Width           =   2415
      End
      Begin VB.Label Label16 
         Caption         =   "Height:"
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
         TabIndex        =   92
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblAtt20Height 
         Alignment       =   2  'Center
         Caption         =   "16"
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
         Left            =   3360
         TabIndex        =   91
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Label lblAtt 
      Alignment       =   2  'Center
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
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmMapAtt"
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
    Unload Me
End Sub

Private Sub btnOk_Click()
    Dim bytByte As Byte, A As Long
    CurAtt = NewAtt
    Select Case CurAtt
        Case 1 'WallTile
            bytByte = 0
            For A = 0 To 7
                If chkAtt1(A).Value Then
                    SetBit bytByte, A
                End If
            Next A
            CurWall = bytByte
            frmMapEdit.RedrawTile
        Case 2 'Warp
            CurAttData(0) = Int(sclAtt2Map \ 256)
            CurAttData(1) = sclAtt2Map Mod 256
            CurAttData(2) = sclAtt2X
            CurAttData(3) = sclAtt2Y
        Case 3 'Key
            CurAttData(0) = sclAtt3Key Mod 256
            CurAttData(1) = optAtt3Pick.Value
            CurAttData(2) = sclAtt3Str
            CurAttData(3) = sclAtt3Key \ 256
        Case 6 'News
            bytByte = 0
            If chkClickable Then
                SetBit bytByte, 0
            End If
            CurAttData(0) = sclAtt6Num
            CurAttData(1) = sclAtt6Data(0)
            CurAttData(2) = sclAtt6Data(1)
            CurAttData(3) = bytByte
        Case 7 'Object
            CurAttData(0) = (((sclAtt7Obj \ 256) And 1) Or (sclAtt7Mag * 2))
            CurAttData(1) = sclAtt7Val.Value \ 256
            CurAttData(2) = sclAtt7Val Mod 256
            CurAttData(3) = sclAtt7Obj Mod 256
        Case 8 'Touch Plate
            CurAttData(0) = sclAtt8X
            CurAttData(1) = sclAtt8Y
            CurAttData(2) = sclAtt8Hall
            'CurAttData(3) = 0
            'If chkAtt8(0).Value = 1 Then
            '    SetBit CurAttData(3), 0
            'End If
            'If chkAtt8(1).Value Then
            '    SetBit CurAttData(3), 1
            'End If
            CurAttData(3) = chkAtt8(0) + chkAtt8(1) * 2
        Case 9 'Damage
            CurAttData(0) = sclAtt9Damage
            CurAttData(1) = 0
            CurAttData(2) = 0
            CurAttData(3) = 0
        Case 14 'ClickTile
            bytByte = 0
            For A = 0 To 2
                If chkAtt14Flag(A).Value Then
                    SetBit bytByte, A
                End If
            Next A
            CurAttData(0) = bytByte
            CurAttData(1) = 0
            CurAttData(2) = 0
            CurAttData(3) = 0
        Case 15 'Chest
            CurAttData(0) = chkAtt15
            CurAttData(1) = sclAtt15Chest
            CurAttData(2) = 0
            CurAttData(3) = 0
        Case 17 'Wall2
            bytByte = 0
            For A = 0 To 7
                If chkAtt17Blocked(A).Value Then
                    SetBit bytByte, A
                End If
            Next A
            CurAttData(0) = bytByte
            CurAttData(1) = 0
            CurAttData(2) = 0
            CurAttData(3) = 0
        Case 18 'Half Tile
            CurAttData(0) = chkAtt18(0) + chkAtt18(1) * 2
            CurAttData(1) = chkAtt18(2) + chkAtt18(3) * 2
            CurAttData(2) = chkAtt18(4)
            CurAttData(3) = 0
        Case 19 'Object Graphic
            If sclAtt19Source <> 0 Then
                CurAttData(0) = sclAtt19
                CurAttData(1) = sclAtt19Source
                CurAttData(2) = sclAtt19Frame
                CurAttData(3) = chkAtt19Clickable
            Else
                CurAttData(0) = IIf(sclAtt19 > 255, sclAtt19 - 255, sclAtt19)
                CurAttData(1) = sclAtt19Source
                CurAttData(2) = IIf(sclAtt19 > 255, 1, 0)
                CurAttData(3) = chkAtt19Clickable
            End If
        Case 20 'Half Player
            CurAttData(0) = sclAtt20Height
            CurAttData(1) = 0
            CurAttData(2) = 0
            CurAttData(3) = 0
        Case 22 'Light Source
            CurAttData(0) = sclAtt22(0)
            CurAttData(1) = sclAtt22(1)
            CurAttData(2) = sclAtt22(2)
            CurAttData(3) = sclAtt22Light
        Case 23 'Shift Tile
            CurAttData(0) = Abs(sclAtt23(0))
            CurAttData(1) = Abs(sclAtt23(1))
            bytByte = 0
            For A = 0 To 4
                If chkAtt23(A) = 1 Then SetBit bytByte, A
            Next A
            CurAttData(2) = bytByte
            bytByte = 0
            If sclAtt23(0) < 0 Then SetBit bytByte, 0
            If sclAtt23(1) < 0 Then SetBit bytByte, 1
            CurAttData(3) = bytByte
        Case 24 'Mining
            CurAttData(0) = sclAtt24(0) 'mine type
            CurAttData(1) = sclAtt24(1) 'max ore
            bytByte = 0
            If chkDisposable = 1 Then SetBit bytByte, 0
            If chkOneTall = 1 Then SetBit bytByte, 1
            If chkTwoTall = 1 Then SetBit bytByte, 2
            If chkTreeLeft = 1 Then SetBit bytByte, 3
            If chkTreeRight = 1 Then SetBit bytByte, 4
            If chkBranchLeft = 1 Then SetBit bytByte, 5
            If chkBranchRight = 1 Then SetBit bytByte, 6
            If chkZeroTall = 1 Then SetBit bytByte, 7
            CurAttData(2) = bytByte 'flags
            CurAttData(3) = sclAtt24(1) 'current ore
        Case 25 'NPC
            CurAttData(0) = sclAtt25NPC
        Case 26 'Movement
            bytByte = 0
            CurAttData(0) = sclWalkSpeed
            CurAttData(1) = sclRunSpeed
            If chkNoRunEnergy = 1 Then SetBit bytByte, 0
            If chk1Energy = 1 Then SetBit bytByte, 1
            If chk2Energy = 1 Then SetBit bytByte, 2
            CurAttData(2) = bytByte
            CurAttData(3) = sclHalfPlayer
        Case 27 'Target Tile blocker
            'blocks target from finihsing on this square
        Case 28 'Mirror
    End Select
    Unload Me
End Sub

Private Sub Form_Load()
Dim A As Long, b As Long
Dim R1 As RECT, r2 As RECT
    If NewAtt < 50 Then
        Select Case NewAtt
            Case 1 'Wall
                lblAtt = "Wall Tile"
                picAtt1.Visible = True
            Case 2 'Warp
                lblAtt = "2 - Warp"
                picAtt2.Visible = True
            Case 3 'Key
                lblAtt = "3 - Key"
                picAtt3.Visible = True
            Case 6 'News
                lblAtt = "6 - News"
                picAtt6.Visible = True
            Case 7 'Obj
                lblAtt = "7 - Object"
                picAtt7.Visible = True
            Case 8 'Touch Plate
                lblAtt = "8 - Touch Plate"
                picAtt8.Visible = True
            Case 9 'Damage
                lblAtt = "9 - Damage"
                picAtt9.Visible = True
            Case 14 'ClickTile
                picAtt14.Visible = True
                lblAtt = "14 - Click Tile"
            Case 15 'Chest
                picAtt15.Visible = True
                lblAtt = "15 - Chest"
            Case 17 'Directional Wall
                For A = 0 To 7
                    chkAtt17Blocked(A).Value = 0
                Next A
                picAtt17.Visible = True
                lblAtt = "17 - Wall"
            Case 18 'Half
                picAtt18.Visible = True
                lblAtt = "18 - Half"
            Case 19 'Graphic Att
                picAtt19.Visible = True
                lblAtt = "19 - ObjectGFX"
                A = (sclAtt19 - 1)
                b = A Mod 64
                A = A \ 64
                If A > 0 And A <= 4 Then
                    R1.Top = 0: R1.Bottom = 32: R1.Left = 0: R1.Right = 32
                    r2.Top = (b \ 8) * 32: r2.Left = (b Mod 8) * 32: r2.Right = r2.Left + 32: r2.Bottom = r2.Top + 32
                    sfcObjects(A).Surface.BltToDC picAtt19Object.hdc, r2, R1
                End If
                picAtt19Object.Refresh
            Case 20 'Half Player
                picAtt20.Visible = True
                lblAtt = "20 - Half Player"
            Case 22
                picAtt22.Visible = True
                lblAtt = "22 - Light Source"
            Case 23
                picAtt23.Visible = True
                lblAtt = "23 - Shift Tile"
            Case 24 'Mining
                picAtt24.Visible = True
                lblAtt = "24 - Mining"
            Case 25 'NPC
                picAtt25.Visible = True
                lblAtt = "25 - NPC"
            Case 26 'Movement
                picAtt26.Visible = True
                lblAtt = "26 - Movement"
        End Select
    Else
        NewAtt = NewAtt - 50
        Select Case NewAtt
            Case 1 'Wall
                lblAtt = "Wall Tile"
                picAtt1.Visible = True
            Case 2 'Warp
                lblAtt = "2 - Warp"
                sclAtt2Map.Value = CurAttData(0) * 256 + CurAttData(1)
                sclAtt2X = CurAttData(2)
                sclAtt2Y = CurAttData(3)
                picAtt2.Visible = True
            Case 3 'Key
                lblAtt = "3 - Key"
                sclAtt3Key = CurAttData(3) * 256 + CurAttData(0)
                optAtt3Pick.Value = CurAttData(1)
                sclAtt3Str = CurAttData(2)
                picAtt3.Visible = True
            Case 6 'News
                lblAtt = "6 - News"
                sclAtt6Num = CurAttData(0)
                sclAtt6Data(0) = CurAttData(1)
                sclAtt6Data(1) = CurAttData(2)
                'sclAtt6Data(2) = CurAttData(3)
                picAtt6.Visible = True
            Case 7 'Obj
                lblAtt = "7 - Object"
                sclAtt7Obj.Value = (CurAttData(0) And 1) * 256 + CurAttData(3)
                sclAtt7Val.Value = CurAttData(1) * 256 + CurAttData(2)
                sclAtt7Mag.Value = CurAttData(0) \ 2
                picAtt7.Visible = True
            Case 8 'Touch Plate
                lblAtt = "8 - Touch Plate"
                sclAtt8X = CurAttData(0)
                sclAtt8Y = CurAttData(1)
                sclAtt8Hall = CurAttData(2)
                chkAtt8(0).Value = (CurAttData(3) And 1)
                chkAtt8(1).Value = CurAttData(3) \ 2
                picAtt8.Visible = True
            Case 9 'Damage
                lblAtt = "9 - Damage"
                sclAtt9Damage = CurAttData(0)
                picAtt9.Visible = True
            Case 14 'ClickTile
                picAtt14.Visible = True
                For A = 0 To 2
                    chkAtt14Flag(A).Value = ExamineBit(CurAttData(0), A)
                Next A
                lblAtt = "14 - Click Tile"
            Case 15 'Chest
                picAtt15.Visible = True
                chkAtt15 = CurAttData(0)
                sclAtt15Chest = CurAttData(1)
                lblAtt = "15 - Chest"
            Case 17 'Directional Wall
                For A = 0 To 7
                    chkAtt17Blocked(A).Value = 0
                Next A
                picAtt17.Visible = True
                lblAtt = "17 - Wall"
            Case 18 'Half
                picAtt18.Visible = True
                chkAtt18(0) = CurAttData(0) And 1
                chkAtt18(1) = CurAttData(0) \ 2
                chkAtt18(2) = CurAttData(1) And 1
                chkAtt18(3) = CurAttData(1) \ 2
                chkAtt18(4) = CurAttData(2)
                lblAtt = "18 - Half"
            Case 19 'Graphic Att
                picAtt19.Visible = True
                lblAtt = "19 - ObjectGFX"
                sclAtt19 = CurAttData(0)
                sclAtt19Source = CurAttData(1)
                sclAtt19Frame = CurAttData(2)
                chkAtt19Clickable = CurAttData(3)
                A = (sclAtt19 - 1)
                b = A Mod 64
                A = A \ 64
                If A > 0 And A <= 4 Then
                    R1.Top = 0: R1.Bottom = 32: R1.Left = 0: R1.Right = 32
                    r2.Top = (b \ 8) * 32: r2.Left = (b Mod 8) * 32: r2.Right = r2.Left + 32: r2.Bottom = r2.Top + 32
                    sfcObjects(A).Surface.BltToDC picAtt19Object.hdc, r2, R1
                End If
                picAtt19Object.Refresh
            Case 20 'Half Player
                picAtt20.Visible = True
                sclAtt20Height = CurAttData(0)
                lblAtt = "20 - Half Player"
            Case 22
                picAtt22.Visible = True
                sclAtt22(0) = CurAttData(0)
                sclAtt22(1) = CurAttData(1)
                sclAtt22(2) = CurAttData(2)
                sclAtt22Light = CurAttData(3)
                lblAtt = "22 - Light Source"
            Case 23
                picAtt23.Visible = True
                sclAtt23(0) = CurAttData(0)
                sclAtt23(1) = CurAttData(1)
                For A = 0 To 4
                    If ExamineBit(CurAttData(2), A) Then chkAtt23(A) = 1
                Next A
                If ExamineBit(CurAttData(3), 0) Then sclAtt23(0) = -sclAtt23(0)
                If ExamineBit(CurAttData(3), 1) Then sclAtt23(1) = -sclAtt23(1)
                lblAtt = "23 - Shift Tile"
            Case 24 'Mining
                picAtt24.Visible = True
                lblAtt = "24 - Mining"
            Case 25 'NPC
                picAtt25.Visible = True
                sclAtt25NPC = CurAttData(0)
                lblAtt = "25 - NPC"
            Case 26
                picAtt26.Visible = True
                lblAtt = "25 - Movement"
        End Select
    End If
End Sub

Private Sub lblAtt2Map_DblClick()
txtAtt2Map.Text = Val(lblAtt2Map)
txtAtt2Map.Visible = True
End Sub

Private Sub sclAtt19_Change()
    lblAtt19 = sclAtt19
    Dim R1 As RECT, r2 As RECT
    Dim A As Long, b As Long
    Select Case sclAtt19Source
            Case 0
                R1.Top = 0: R1.Bottom = 32: R1.Left = 0: R1.Right = 32
                A = sclAtt19 - 1
                b = A Mod 64
                r2.Top = (b \ 8) * 32: r2.Left = (b Mod 8) * 32: r2.Right = r2.Left + 32: r2.Bottom = r2.Top + 32
                A = (A \ 64) + 1
                If A > 0 And A <= 8 Then
                    sfcObjects(A).Surface.BltToDC picAtt19Object.hdc, r2, R1
                End If
            Case 1
                R1.Top = 0: R1.Bottom = 32: R1.Left = 0: R1.Right = 32
                A = sclAtt19 - 1
                r2.Top = (A \ 16) * 32: r2.Left = (A Mod 16) * 32: r2.Right = r2.Left + 32: r2.Bottom = r2.Top + 32
                sfcSprites.Surface.BltToDC picAtt19Object.hdc, r2, R1
    End Select
        picAtt19Object.Refresh
End Sub

Private Sub sclAtt19_Scroll()
    sclAtt19_Change
End Sub

Private Sub sclAtt19Frame_Change()
    sclAtt19_Change
End Sub

Private Sub sclAtt19Source_Change()
    sclAtt19.max = 255
    Select Case sclAtt19Source.Value
        Case 0
            lblAtt19Source = "Objects"
            sclAtt19.max = 512
        Case 1
            lblAtt19Source = "Sprites"
        Case 2
            lblAtt19Source = "Effects"
    End Select
    sclAtt19_Change
End Sub

Private Sub sclAtt20Height_Change()
    lblAtt20Height = sclAtt20Height
End Sub

Private Sub sclAtt20Height_Scroll()
    sclAtt20Height_Change
End Sub

Private Sub sclAtt22_Change(Index As Integer)
    lblAtt22(Index).Caption = sclAtt22(Index)
End Sub

Private Sub sclAtt22_Scroll(Index As Integer)
    sclAtt22_Change Index
End Sub

Private Sub sclAtt22Light_Change()
    lblAtt22Light = Lights(sclAtt22Light).Name
End Sub

Private Sub sclAtt22Light_Scroll()
    sclAtt22Light_Change
End Sub

Private Sub sclAtt23_Change(Index As Integer)
    lblAtt23(Index) = sclAtt23(Index)
End Sub

Private Sub sclAtt23_Scroll(Index As Integer)
    sclAtt23_Change Index
End Sub


Private Sub sclAtt24_Change(Index As Integer)
    
    lblAtt24(Index) = sclAtt24(Index)

End Sub
Private Sub sclAtt24_SCroll(Index As Integer)
    
    lblAtt24(Index) = sclAtt24(Index)

End Sub

Private Sub sclAtt25NPC_Change()
    lblAtt25NPC = sclAtt25NPC
    LblAtt25Name = NPC(sclAtt25NPC).Name
End Sub

Private Sub sclAtt2Map_Change()
    lblAtt2Map = sclAtt2Map
End Sub


Private Sub sclAtt2Map_Scroll()
    sclAtt2Map_Change
End Sub


Private Sub sclAtt2X_Change()
    lblAtt2X = sclAtt2X
End Sub


Private Sub sclAtt2X_Scroll()
    sclAtt2X_Change
End Sub


Private Sub sclAtt2Y_Change()
    lblAtt2Y = sclAtt2Y
End Sub


Private Sub sclAtt2Y_Scroll()
    sclAtt2Y_Change
End Sub


Private Sub sclAtt3Key_Change()
    lblAtt3Key = sclAtt3Key
End Sub

Private Sub sclAtt3Key_Scroll()
    sclAtt3Key_Change
End Sub


Private Sub sclAtt3Str_Change()
lblAtt3Str = sclAtt3Str
End Sub

Private Sub sclAtt3Str_Scroll()
sclAtt3Str_Change
End Sub

Private Sub sclAtt6Data_Change(Index As Integer)
    lblAtt6Data(Index) = sclAtt6Data(Index)
End Sub

Private Sub sclAtt6Data_Scroll(Index As Integer)
    sclAtt6Data_Change (Index)
End Sub

Private Sub sclAtt6Num_Change()
    lblAtt6Num = sclAtt6Num
End Sub

Private Sub sclAtt6Num_Scroll()
    sclAtt6Num_Change
End Sub


Private Sub sclAtt7Mag_Change()
    lblAtt7Mag = Int(sclAtt7Mag * 4)
End Sub

Private Sub sclAtt7Mag_Scroll()
    sclAtt7Mag_Change
End Sub

Private Sub sclAtt7Obj_Change()
    lblAtt7Obj = sclAtt7Obj
    lblAtt7Name = Object(CInt(lblAtt7Obj.Caption)).Name
End Sub


Private Sub sclAtt7Obj_Scroll()
    sclAtt7Obj_Change
End Sub


Private Sub sclAtt7Val_Change()
    lblAtt7Val = sclAtt7Val
End Sub


Private Sub sclAtt7Val_Scroll()
    sclAtt7Val_Change
End Sub


Private Sub sclAtt8Hall_Change()
    lblAtt8Hall = sclAtt8Hall
End Sub

Private Sub sclAtt8Hall_Scroll()
    sclAtt8Hall_Change
End Sub


Private Sub sclAtt8X_Change()
    lblAtt8X = sclAtt8X
End Sub


Private Sub sclAtt8X_Scroll()
    sclAtt8X_Change
End Sub


Private Sub sclAtt8Y_Change()
    lblAtt8Y = sclAtt8Y
End Sub


Private Sub sclAtt8Y_Scroll()
    sclAtt8Y_Change
End Sub


Private Sub sclAtt9Damage_Change()
    lblAtt9Damage = sclAtt9Damage
End Sub


Private Sub sclAtt9Damage_Scroll()
    sclAtt9Damage_Change
End Sub


Private Sub sclHalfPlayer_Change()
    lblAtt24(4) = sclHalfPlayer
End Sub

Private Sub sclRunSpeed_Change()
lblAtt24(2) = sclRunSpeed
End Sub

Private Sub sclWalkSpeed_Change()
lblAtt24(3) = sclWalkSpeed
End Sub

Private Sub txtAtt2Map_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 10 Or KeyCode = 13 Then
    txtAtt2Map_LostFocus
End If
End Sub

Private Sub txtAtt2Map_LostFocus()
If Val(txtAtt2Map) >= 1 And Val(txtAtt2Map) <= 5000 Then
    sclAtt2Map.Value = Val(txtAtt2Map)
End If
txtAtt2Map.Visible = False
End Sub
