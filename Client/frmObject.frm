VERSION 5.00
Begin VB.Form frmObject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seyerdin Online [Item Editor]"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7620
   ControlBox      =   0   'False
   Icon            =   "frmObject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   554
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   508
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar sclMagicAmp 
      Height          =   255
      Left            =   960
      Max             =   255
      TabIndex        =   123
      Top             =   4920
      Value           =   100
      Width           =   3495
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "Unpickupable"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   122
      Top             =   6240
      Width           =   1335
   End
   Begin VB.PictureBox picFrame 
      AutoRedraw      =   -1  'True
      Height          =   540
      Left            =   6120
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   121
      Top             =   4560
      Width           =   540
   End
   Begin VB.HScrollBar sclEquipmentPicture 
      Height          =   255
      Left            =   2160
      Max             =   255
      TabIndex        =   119
      Top             =   4560
      Width           =   3495
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "Sprite > 255"
      Enabled         =   0   'False
      Height          =   255
      Index           =   6
      Left            =   6120
      TabIndex        =   117
      Top             =   600
      Width           =   1335
   End
   Begin VB.HScrollBar sclLevel 
      Height          =   255
      Left            =   1320
      Max             =   10
      Min             =   1
      TabIndex        =   107
      Top             =   1920
      Value           =   1
      Width           =   3495
   End
   Begin VB.ListBox lstClass 
      Appearance      =   0  'Flat
      Height          =   2280
      ItemData        =   "frmObject.frx":000C
      Left            =   6000
      List            =   "frmObject.frx":002E
      Style           =   1  'Checkbox
      TabIndex        =   106
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox txtDescription 
      Height          =   1575
      Left            =   120
      MaxLength       =   512
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   105
      Top             =   6600
      Width           =   7335
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "Dual Wield [w]"
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   72
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "Undepositable"
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   71
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "Two Handed [w]"
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   70
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "Unrepairable"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   69
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "Unbreakable"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   68
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CheckBox chkFlags 
      Caption         =   "Undroppable"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   67
      Top             =   5760
      Width           =   1215
   End
   Begin VB.HScrollBar sclMinLevel 
      Height          =   255
      Left            =   1320
      Max             =   255
      Min             =   1
      TabIndex        =   12
      Top             =   1560
      Value           =   1
      Width           =   3495
   End
   Begin VB.PictureBox picPicture 
      AutoRedraw      =   -1  'True
      Height          =   540
      Left            =   5280
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   8
      Top             =   600
      Width           =   540
   End
   Begin VB.HScrollBar sclPicture 
      Height          =   255
      Left            =   1320
      Max             =   512
      Min             =   1
      TabIndex        =   1
      Top             =   1200
      Value           =   1
      Width           =   3495
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
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   0
      Top             =   600
      Width           =   3855
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      Top             =   6000
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   6000
      Width           =   1455
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2280
      Width           =   3495
   End
   Begin VB.PictureBox picRing 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   5595
      TabIndex        =   21
      Top             =   2640
      Visible         =   0   'False
      Width           =   5655
      Begin VB.HScrollBar sclRingAmount 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   46
         Top             =   720
         Width           =   3495
      End
      Begin VB.HScrollBar sclRingHP 
         Height          =   255
         Left            =   1200
         Max             =   255
         Min             =   1
         TabIndex        =   43
         Top             =   360
         Value           =   1
         Width           =   3495
      End
      Begin VB.ComboBox cmbRingType 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   0
         Width           =   3495
      End
      Begin VB.Label lblCaption 
         Caption         =   "Amount:"
         Height          =   255
         Index           =   12
         Left            =   0
         TabIndex        =   48
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblRingAmount 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   47
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "HP / 10:"
         Height          =   255
         Index           =   11
         Left            =   0
         TabIndex        =   45
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblRingHP 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   44
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "Type:"
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picWeapon 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   5595
      TabIndex        =   15
      Top             =   2640
      Visible         =   0   'False
      Width           =   5655
      Begin VB.HScrollBar sclWeaponMin 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   32
         Top             =   360
         Width           =   3495
      End
      Begin VB.HScrollBar sclWeaponhp 
         Height          =   255
         Left            =   1200
         Max             =   255
         Min             =   1
         TabIndex        =   26
         Top             =   0
         Value           =   1
         Width           =   3495
      End
      Begin VB.HScrollBar sclWeaponMax 
         Height          =   255
         Left            =   1200
         Max             =   10000
         TabIndex        =   25
         Top             =   720
         Width           =   3495
      End
      Begin VB.HScrollBar sclASpeed 
         Height          =   255
         Left            =   960
         Max             =   10
         TabIndex        =   22
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblWeaponMax 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   33
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "HP / 10"
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblWeaponHP 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   30
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "Min:"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   29
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblWeaponMin 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   28
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "Max"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblSpeed 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label cptSpeed 
         Caption         =   "Speed:"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   735
      End
   End
   Begin VB.PictureBox picPotion 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   5595
      TabIndex        =   19
      Top             =   2640
      Visible         =   0   'False
      Width           =   5655
      Begin VB.HScrollBar sclPotionStrength 
         Height          =   255
         Left            =   1200
         Max             =   1000
         TabIndex        =   110
         Top             =   840
         Width           =   3495
      End
      Begin VB.ComboBox cmbPotionType 
         Height          =   315
         ItemData        =   "frmObject.frx":0091
         Left            =   1200
         List            =   "frmObject.frx":0093
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   0
         Width           =   3495
      End
      Begin VB.HScrollBar sclPotionValue 
         Height          =   255
         Left            =   1200
         Max             =   1000
         TabIndex        =   50
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label lblPotionStrength 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   112
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "Value:"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   111
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblCaption 
         Caption         =   "Type:"
         Height          =   255
         Index           =   14
         Left            =   0
         TabIndex        =   53
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblCaption 
         Caption         =   "Value:"
         Height          =   255
         Index           =   13
         Left            =   0
         TabIndex        =   52
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblPotionValue 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   51
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.PictureBox picMoney 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   5595
      TabIndex        =   113
      Top             =   2640
      Visible         =   0   'False
      Width           =   5655
      Begin VB.HScrollBar sclStackSize 
         Height          =   255
         LargeChange     =   10
         Left            =   1080
         Max             =   255
         TabIndex        =   114
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label lblStackSize 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   116
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "Stack Size:"
         Height          =   255
         Index           =   26
         Left            =   0
         TabIndex        =   115
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.PictureBox picProjectile 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   5595
      TabIndex        =   89
      Top             =   2640
      Visible         =   0   'False
      Width           =   5655
      Begin VB.HScrollBar sclProjectileAmmoType 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   102
         Top             =   1080
         Width           =   3495
      End
      Begin VB.HScrollBar sclProjectileRange 
         Height          =   255
         Left            =   1200
         Max             =   12
         TabIndex        =   93
         Top             =   360
         Width           =   3495
      End
      Begin VB.HScrollBar sclProjectileHP 
         Height          =   255
         Left            =   1200
         Max             =   255
         Min             =   1
         TabIndex        =   92
         Top             =   0
         Value           =   1
         Width           =   3495
      End
      Begin VB.HScrollBar sclProjectileSpeed 
         Height          =   255
         Left            =   720
         Max             =   40
         TabIndex        =   91
         Top             =   1440
         Value           =   10
         Width           =   1695
      End
      Begin VB.HScrollBar sclProjectilePlus 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   90
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label lblCaption 
         Caption         =   "Ammo Type:"
         Height          =   255
         Index           =   23
         Left            =   0
         TabIndex        =   104
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblProjectileAmmoType 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   103
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "HP / 10"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   101
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblProjectileHP 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   100
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "Range:"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   99
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblProjectileRange 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   98
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "Plus:"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   97
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Speed:"
         Height          =   255
         Left            =   0
         TabIndex        =   95
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblProjectilePlus 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   94
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblProjectileSpeed 
         Alignment       =   2  'Center
         Caption         =   "10"
         Height          =   255
         Left            =   2520
         TabIndex        =   96
         Top             =   1440
         Width           =   1695
      End
   End
   Begin VB.PictureBox picAmmo 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   5595
      TabIndex        =   73
      Top             =   2640
      Visible         =   0   'False
      Width           =   5655
      Begin VB.HScrollBar sclAmmoType 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   86
         Top             =   1440
         Width           =   3495
      End
      Begin VB.HScrollBar sclAmmoMax 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   83
         Top             =   1080
         Width           =   3495
      End
      Begin VB.HScrollBar sclAmmoAnimation 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   76
         Top             =   360
         Width           =   3495
      End
      Begin VB.HScrollBar sclAmmoLimit 
         Height          =   255
         Left            =   1200
         Max             =   255
         Min             =   1
         TabIndex        =   75
         Top             =   0
         Value           =   1
         Width           =   3495
      End
      Begin VB.HScrollBar sclAmmoMin 
         Height          =   255
         Left            =   1200
         Max             =   255
         TabIndex        =   74
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label lblCaption 
         Caption         =   "Ammo Type:"
         Height          =   255
         Index           =   22
         Left            =   0
         TabIndex        =   88
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblAmmoType 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   87
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "Max:"
         Height          =   255
         Index           =   21
         Left            =   0
         TabIndex        =   85
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblAmmoMax 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   84
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "Limit"
         Height          =   255
         Index           =   20
         Left            =   0
         TabIndex        =   82
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblAmmoLimit 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   81
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "Animation:"
         Height          =   255
         Index           =   19
         Left            =   0
         TabIndex        =   80
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblAmmoAnimation 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   79
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "Min:"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   78
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblAmmoMin 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   77
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.PictureBox picArmor 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   5595
      TabIndex        =   16
      Top             =   2640
      Visible         =   0   'False
      Width           =   5655
      Begin VB.HScrollBar sclArmorResist 
         Height          =   255
         Left            =   1500
         Max             =   255
         TabIndex        =   148
         Top             =   720
         Width           =   3200
      End
      Begin VB.HScrollBar sclArmorMDefense 
         Height          =   255
         Left            =   1500
         Max             =   255
         TabIndex        =   147
         Top             =   1080
         Width           =   3200
      End
      Begin VB.HScrollBar sclArmorMResist 
         Height          =   255
         Left            =   1500
         Max             =   255
         TabIndex        =   146
         Top             =   1440
         Width           =   3200
      End
      Begin VB.HScrollBar sclArmorHP 
         Height          =   255
         Left            =   1500
         Max             =   255
         Min             =   1
         TabIndex        =   40
         Top             =   0
         Value           =   1
         Width           =   3200
      End
      Begin VB.HScrollBar sclArmorDefense 
         Height          =   255
         Left            =   1500
         Max             =   1000
         TabIndex        =   35
         Top             =   360
         Width           =   3200
      End
      Begin VB.HScrollBar sclData 
         Height          =   255
         Index           =   5
         Left            =   1560
         Max             =   255
         TabIndex        =   34
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblArmorMResist 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   154
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblArmorMDefense 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   153
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblArmorResist 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   152
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "Physical Resistance: "
         Height          =   255
         Index           =   35
         Left            =   0
         TabIndex        =   151
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblCaption 
         Caption         =   "Magical Defense:"
         Height          =   255
         Index           =   34
         Left            =   0
         TabIndex        =   150
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblCaption 
         Caption         =   "Magical Resistance: "
         Height          =   255
         Index           =   33
         Left            =   0
         TabIndex        =   149
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblCaption 
         Caption         =   "HP / 10:"
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblArmorHP 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   38
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "Defense:"
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   37
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblArmorDefense 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   36
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.PictureBox picShield 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   5595
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   5655
      Begin VB.HScrollBar sclshieldmagicchance 
         Height          =   255
         Left            =   1550
         Max             =   255
         TabIndex        =   132
         Top             =   1080
         Width           =   3195
      End
      Begin VB.HScrollBar sclshieldmagicpercent 
         Height          =   255
         Left            =   1550
         Max             =   255
         TabIndex        =   131
         Top             =   1440
         Width           =   3195
      End
      Begin VB.HScrollBar sclShieldDamagePercent 
         Height          =   255
         Left            =   1550
         Max             =   255
         TabIndex        =   127
         Top             =   720
         Value           =   100
         Width           =   3195
      End
      Begin VB.HScrollBar sclShieldDefense 
         Height          =   255
         Left            =   1550
         Max             =   255
         TabIndex        =   62
         Top             =   360
         Width           =   3195
      End
      Begin VB.HScrollBar sclShieldHP 
         Height          =   255
         Left            =   1550
         Max             =   255
         Min             =   1
         TabIndex        =   61
         Top             =   0
         Value           =   1
         Width           =   3195
      End
      Begin VB.Label lblshieldmagicchance 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   136
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblshieldmagicpercent 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   135
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "Magic Block Chance:"
         Height          =   255
         Index           =   29
         Left            =   0
         TabIndex        =   134
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblCaption 
         Caption         =   "Magic Block %:"
         Height          =   255
         Index           =   28
         Left            =   0
         TabIndex        =   133
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblCaption 
         Caption         =   "Block %:"
         Height          =   255
         Index           =   27
         Left            =   0
         TabIndex        =   129
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblShieldDamagePercent 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   128
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblShieldDefense 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   66
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "Block Chance:"
         Height          =   255
         Index           =   18
         Left            =   0
         TabIndex        =   65
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblShieldHP 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   64
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "HP / 10:"
         Height          =   255
         Index           =   17
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picHelmet 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   5595
      TabIndex        =   18
      Top             =   2640
      Visible         =   0   'False
      Width           =   5655
      Begin VB.HScrollBar sclHelmMResist 
         Height          =   255
         Left            =   1500
         Max             =   255
         TabIndex        =   143
         Top             =   1440
         Width           =   3200
      End
      Begin VB.HScrollBar sclHelmMDefense 
         Height          =   255
         Left            =   1500
         Max             =   255
         TabIndex        =   140
         Top             =   1080
         Width           =   3200
      End
      Begin VB.HScrollBar sclHelmResist 
         Height          =   255
         Left            =   1500
         Max             =   255
         TabIndex        =   137
         Top             =   720
         Width           =   3200
      End
      Begin VB.HScrollBar sclHelmetDefense 
         Height          =   255
         Left            =   1500
         Max             =   1000
         TabIndex        =   56
         Top             =   360
         Width           =   3200
      End
      Begin VB.HScrollBar sclHelmetHP 
         Height          =   255
         Left            =   1500
         Max             =   255
         Min             =   1
         TabIndex        =   55
         Top             =   0
         Value           =   1
         Width           =   3200
      End
      Begin VB.Label lblCaption 
         Caption         =   "Magical Resistance: "
         Height          =   255
         Index           =   32
         Left            =   0
         TabIndex        =   145
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblHelmMResist 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   144
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "Magical Defense:"
         Height          =   255
         Index           =   31
         Left            =   0
         TabIndex        =   142
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblHelmMDefense 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   141
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "Physical Resistance: "
         Height          =   255
         Index           =   30
         Left            =   0
         TabIndex        =   139
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblHelmResist 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   138
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblHelmetDefense 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   60
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "Physical Defense:"
         Height          =   255
         Index           =   16
         Left            =   0
         TabIndex        =   59
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblHelmetHP 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   255
         Left            =   4800
         TabIndex        =   58
         Top             =   0
         Width           =   855
      End
      Begin VB.Label lblCaption 
         Caption         =   "HP / 10:"
         Height          =   255
         Index           =   15
         Left            =   0
         TabIndex        =   57
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox picKey 
      Height          =   1815
      Left            =   120
      ScaleHeight     =   1755
      ScaleWidth      =   5595
      TabIndex        =   20
      Top             =   2640
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CheckBox chkKeyUnlim 
         Caption         =   "Unlimited Use"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Label lblPicture 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Left            =   4920
      TabIndex        =   130
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblCaption 
      Caption         =   "Magic Amp:"
      Height          =   255
      Index           =   25
      Left            =   360
      TabIndex        =   126
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblMagicAmp 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
      Height          =   255
      Left            =   4560
      TabIndex        =   125
      Top             =   4920
      Width           =   855
   End
   Begin VB.Label lblCaption 
      Caption         =   "Magic Amp:"
      Height          =   255
      Index           =   24
      Left            =   0
      TabIndex        =   124
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label lblEquipmentPicture 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   5520
      TabIndex        =   120
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Equipped Object Picture:"
      Height          =   255
      Left            =   240
      TabIndex        =   118
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Item Level:"
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
      TabIndex        =   109
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Left            =   4920
      TabIndex        =   108
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Restricted:"
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
      Left            =   6000
      TabIndex        =   14
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblMinLevel 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Min Level:"
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
      TabIndex        =   11
      Top             =   1560
      Width           =   1095
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
      Left            =   1320
      TabIndex        =   10
      Top             =   120
      Width           =   4455
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
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Picture:"
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
      TabIndex        =   7
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Type:"
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
      Top             =   2280
      Width           =   1095
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
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmObject"
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
Dim ChangingTypes As Boolean
Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnOk_Click()
Dim bytFlags As Byte, intClass As Integer, A As Long, St As String
For A = 0 To 7
    If chkFlags(A).Value Then
        SetBit bytFlags, A
    End If
Next A
For A = 0 To MAX_CLASS - 1
    If lstClass.Selected(A) = True Then
        intClass = intClass Or (2 ^ A)
    End If
Next A

Select Case cmbType.ListIndex
    Case 0, 9 'None,Deed
        St = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
    Case 6
         St = Chr$(sclStackSize) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
         
    Case 1 'Weapon
        St = Chr$(sclWeaponhp) + Chr$(sclWeaponMin) + DoubleChar(sclWeaponMax) + Chr$(sclASpeed) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(sclMagicAmp)
    Case 2 'Shield
        St = Chr$(sclShieldHP) + Chr$(sclShieldDefense) + Chr$(sclShieldDamagePercent) + Chr$(sclshieldmagicchance) + Chr$(sclshieldmagicpercent) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(sclMagicAmp)
    Case 3 'Armor
        St = Chr$(sclArmorHP) + DoubleChar(sclArmorDefense) + Chr$(sclArmorResist) + Chr$(sclArmorMDefense) + Chr$(sclArmorMResist) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(sclMagicAmp)
    Case 4 'Helmet
        St = Chr$(sclHelmetHP) + DoubleChar(sclHelmetDefense) + Chr$(sclHelmResist) + Chr$(sclHelmMDefense) + Chr$(sclHelmMResist) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(sclMagicAmp)
    Case 8 'Ring
        St = Chr$(cmbRingType.ListIndex) + Chr$(sclRingHP) + Chr$(sclRingAmount) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(sclMagicAmp)
        
    Case 5 'Potion
        St = Chr$(cmbPotionType.ListIndex) + DoubleChar$(sclPotionValue) + DoubleChar$(sclPotionStrength) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
    Case 7 'Key
        St = Chr$(chkKeyUnlim) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
    Case 10 'Projectile
        St = Chr$(sclProjectileHP) + Chr$(sclProjectileRange) + Chr$(sclProjectilePlus) + Chr$(sclProjectileAmmoType) + Chr$(sclProjectileSpeed) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(sclMagicAmp)
    Case 11 'Ammo
        St = Chr$(sclAmmoLimit) + Chr$(sclAmmoAnimation) + Chr$(sclAmmoMin) + Chr$(sclAmmoMax) + Chr$(sclAmmoType) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
End Select

    If chkFlags(6) = 1 Then
        SendSocket Chr$(21) + DoubleChar(lblNumber) + Chr$(sclPicture - 255) + Chr$(cmbType.ListIndex) + Chr$(bytFlags) + St + Chr$(sclMinLevel) + DoubleChar$(intClass) + Chr$(sclLevel) + Chr$(sclEquipmentPicture) + txtName + Chr$(0) + txtDescription
    Else
        SendSocket Chr$(21) + DoubleChar(lblNumber) + Chr$(sclPicture) + Chr$(cmbType.ListIndex) + Chr$(bytFlags) + St + Chr$(sclMinLevel) + DoubleChar$(intClass) + Chr$(sclLevel) + Chr$(sclEquipmentPicture) + txtName + Chr$(0) + txtDescription
    End If
    Unload Me
End Sub


Private Sub cmbType_Click()
    ChangingTypes = True

    Select Case cmbType.ListIndex
        Case 0 'None
            picShield.Visible = False
            picArmor.Visible = False
            picWeapon.Visible = False
            picHelmet.Visible = False
            picPotion.Visible = False
            picKey.Visible = False
            picRing.Visible = False
            picMoney.Visible = False
            sclMagicAmp.Enabled = False
        Case 1 'Weapon
            picShield.Visible = False
            picArmor.Visible = False
            picWeapon.Visible = True
            picHelmet.Visible = False
            picPotion.Visible = False
            picKey.Visible = False
            picRing.Visible = False
            picMoney.Visible = False
            sclMagicAmp.Enabled = True
        Case 2 'Shield
            picShield.Visible = True
            picArmor.Visible = False
            picWeapon.Visible = False
            picHelmet.Visible = False
            picPotion.Visible = False
            picKey.Visible = False
            picRing.Visible = False
            picMoney.Visible = False
             sclMagicAmp.Enabled = True
        Case 3 'Armor
            picShield.Visible = False
            picArmor.Visible = True
            picWeapon.Visible = False
            picHelmet.Visible = False
            picPotion.Visible = False
            picKey.Visible = False
            picRing.Visible = False
            picProjectile.Visible = False
            picAmmo.Visible = False
            picMoney.Visible = False
             sclMagicAmp.Enabled = True
        Case 4 'Helmut
            picShield.Visible = False
            picArmor.Visible = False
            picWeapon.Visible = False
            picHelmet.Visible = True
            picPotion.Visible = False
            picKey.Visible = False
            picRing.Visible = False
            picProjectile.Visible = False
            picAmmo.Visible = False
            picMoney.Visible = False
             sclMagicAmp.Enabled = True
        Case 5 'Potion
            With cmbPotionType
                .Clear
                .AddItem "Gives HP"
                .AddItem "Takes HP"
                .AddItem "Give Mana"
                .AddItem "Takes Mana"
                .AddItem "Gives Energy"
                .AddItem "Takes Energy"
                .AddItem "Cures Poison"
                .AddItem "Causes Poison"
                .AddItem "Causes Regen"
                .AddItem "Cures Mute"
                .AddItem "Causes Mute"
            End With
            picShield.Visible = False
            picArmor.Visible = False
            picWeapon.Visible = False
            picHelmet.Visible = False
            picPotion.Visible = True
            picKey.Visible = False
            picRing.Visible = False
            picProjectile.Visible = False
            picAmmo.Visible = False
            picMoney.Visible = False
            sclMagicAmp.Enabled = False
        Case 6 'Money
            picShield.Visible = False
            picArmor.Visible = False
            picWeapon.Visible = False
            picHelmet.Visible = False
            picPotion.Visible = False
            picKey.Visible = False
            picRing.Visible = False
            picProjectile.Visible = False
            picAmmo.Visible = False
            picMoney.Visible = True
            sclMagicAmp.Enabled = False
        Case 7 'Key
            picShield.Visible = False
            picArmor.Visible = False
            picWeapon.Visible = False
            picHelmet.Visible = False
            picPotion.Visible = False
            picKey.Visible = True
            picRing.Visible = False
            picProjectile.Visible = False
            picAmmo.Visible = False
            picMoney.Visible = False
            sclMagicAmp.Enabled = False
        Case 8 'Ring
            With cmbRingType
                .Clear
                .AddItem "Modifies Attack"
                .AddItem "Modifies Defense"
                .AddItem "Nothing"
            End With
            picShield.Visible = False
            picArmor.Visible = False
            picWeapon.Visible = False
            picHelmet.Visible = False
            picPotion.Visible = False
            picKey.Visible = False
            picRing.Visible = True
            picProjectile.Visible = False
            picAmmo.Visible = False
            picMoney.Visible = False
            sclMagicAmp.Enabled = True
        Case 9 'Guild Deed
            picShield.Visible = False
            picArmor.Visible = False
            picWeapon.Visible = False
            picHelmet.Visible = False
            picPotion.Visible = False
            picKey.Visible = False
            picRing.Visible = False
            picProjectile.Visible = False
            picAmmo.Visible = False
            picMoney.Visible = False
            sclMagicAmp.Enabled = False
        Case 10 'Projectile
            picShield.Visible = False
            picArmor.Visible = False
            picWeapon.Visible = False
            picHelmet.Visible = False
            picPotion.Visible = False
            picKey.Visible = False
            picRing.Visible = False
            picProjectile.Visible = True
            picAmmo.Visible = False
            picMoney.Visible = False
            sclMagicAmp.Enabled = True
        Case 11 'Ammo
            picShield.Visible = False
            picArmor.Visible = False
            picWeapon.Visible = False
            picHelmet.Visible = False
            picPotion.Visible = False
            picKey.Visible = False
            picRing.Visible = False
            picProjectile.Visible = False
            picAmmo.Visible = True
            picMoney.Visible = False
            sclMagicAmp.Enabled = False
        End Select
    ChangingTypes = False
End Sub
Private Sub Form_Load()
Dim hdcObjects As Long, A As Long, b As Long
    A = (sclPicture - 1)
    b = A Mod 64
    A = (A \ 64) + 1
    If Int(A \ 64) + 1 <= 4 Then
        hdcObjects = sfcObjects(A).Surface.GetDC
            BitBlt picPicture.hdc, 0, 0, 32, 32, hdcObjects, b Mod 8, (b \ 8) * 32, SRCCOPY
        sfcObjects(A).Surface.ReleaseDC hdcObjects
        picPicture.Refresh
    End If
    cmbType.AddItem "<None>"
    cmbType.AddItem "Weapon"
    cmbType.AddItem "Shield"
    cmbType.AddItem "Armor"
    cmbType.AddItem "Helmet"
    cmbType.AddItem "Potion"
    cmbType.AddItem "Money"
    cmbType.AddItem "Key"
    cmbType.AddItem "Ring"
    cmbType.AddItem "Guild Deed"
    cmbType.AddItem "Projectile"
    cmbType.AddItem "Ammo"
    
    frmObject_Loaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmObject_Loaded = False
End Sub



Private Sub sclAmmoAnimation_Change()
lblAmmoAnimation = sclAmmoAnimation
End Sub

Private Sub sclAmmoAnimation_Scroll()
sclAmmoAnimation_Change
End Sub

Private Sub sclAmmoMin_Change()
lblAmmoMin = sclAmmoMin
End Sub

Private Sub sclAmmoMin_Scroll()
sclAmmoMin_Change
End Sub

Private Sub sclAmmoMax_Change()
lblAmmoMax = sclAmmoMax
End Sub

Private Sub sclAmmoMax_Scroll()
sclAmmoMax_Change
End Sub

Private Sub sclAmmoLimit_Change()
lblAmmoLimit = sclAmmoLimit
End Sub

Private Sub sclAmmoLimit_Scroll()
sclAmmoLimit_Change
End Sub

Private Sub sclAmmoType_Change()
lblAmmoType = sclAmmoType
End Sub

Private Sub sclAmmoType_Scroll()
sclAmmoType_Change
End Sub

Private Sub sclArmorDefense_Change()
lblArmorDefense = sclArmorDefense
End Sub

Private Sub sclArmorDefense_Scroll()
sclArmorDefense_Change
End Sub

Private Sub sclArmorHP_Change()
lblArmorHP = sclArmorHP
End Sub

Private Sub sclArmorHP_Scroll()
sclArmorHP_Change
End Sub

Private Sub sclArmorMDefense_Change()
lblArmorMDefense = sclArmorMDefense
End Sub

Private Sub sclArmorMResist_Change()
lblArmorMResist = sclArmorMResist
End Sub

Private Sub sclArmorResist_Change()
lblArmorResist = sclArmorResist
End Sub

Private Sub sclASpeed_Change()
lblSpeed = sclASpeed
End Sub

Private Sub sclASpeed_Scroll()
sclASpeed_Change
End Sub

Private Sub sclEquipmentPicture_Change()


Dim A As Long, R1 As RECT, r2 As RECT

    lblEquipmentPicture = sclEquipmentPicture.Value
    A = sclEquipmentPicture - 1
        If A >= 0 Then
    R1.Top = 0: R1.Bottom = 32: R1.Left = 0: R1.Right = 32
    r2.Top = (A \ 16) * 32: r2.Left = (A Mod 16) * 32: r2.Right = r2.Left + 32: r2.Bottom = r2.Top + 32
    On Error Resume Next
    sfcObjectFrames.Surface.BltToDC picFrame.hdc, r2, R1
        


    
    picFrame.Refresh
    End If

End Sub

Private Sub sclHelmetDefense_Change()
lblHelmetDefense = sclHelmetDefense
End Sub

Private Sub sclHelmetDefense_Scroll()
sclHelmetDefense_Change
End Sub

Private Sub sclHelmetHP_Change()
lblHelmetHP = sclHelmetHP
End Sub

Private Sub sclHelmetHP_Scroll()
sclHelmetHP_Change
End Sub

Private Sub sclHelmMDefense_Change()
lblHelmMDefense = sclHelmMDefense
End Sub

Private Sub sclHelmMResist_Change()
lblHelmMResist = sclHelmMResist
End Sub

Private Sub sclHelmResist_Change()
lblHelmResist = sclHelmResist
End Sub

Private Sub sclLevel_Change()
    lblLevel = sclLevel
End Sub

Private Sub sclLevel_Scroll()
    sclLevel_Change
End Sub

Private Sub sclMagicAmp_Change()
    lblMagicAmp.Caption = sclMagicAmp
End Sub

Private Sub sclMinLevel_Change()
lblMinLevel.Caption = sclMinLevel.Value
End Sub

Private Sub sclMinLevel_Scroll()
sclMinLevel_Change
End Sub

Private Sub sclPicture_Change()
Dim hdcObjects As Long, A As Long, b As Long
lblPicture.Caption = sclPicture

        If sclPicture > 255 Then
            chkFlags(6) = 1
        Else
            chkFlags(6) = 0
        End If

    A = sclPicture - 1
    b = A Mod 64
    A = (A \ 64) + 1

    hdcObjects = sfcObjects(A).Surface.GetDC
        BitBlt picPicture.hdc, 0, 0, 32, 32, hdcObjects, (b Mod 8) * 32, (b \ 8) * 32, SRCCOPY
    sfcObjects(A).Surface.ReleaseDC hdcObjects
    
    picPicture.Refresh
End Sub


Private Sub sclPicture_Scroll()
    sclPicture_Change
End Sub


Private Sub sclPotionStrength_Change()
    lblPotionStrength = sclPotionStrength
End Sub

Private Sub sclPotionStrength_Scroll()
    sclPotionStrength_Change
End Sub

Private Sub sclPotionValue_Change()
lblPotionValue = sclPotionValue
End Sub

Private Sub sclPotionValue_Scroll()
sclPotionValue_Change
End Sub

Private Sub sclProjectileAmmoType_Change()
lblProjectileAmmoType = sclProjectileAmmoType
End Sub

Private Sub sclProjectileAmmoType_Scroll()
sclProjectileAmmoType_Change
End Sub

Private Sub sclProjectileHP_Change()
lblProjectileHP = sclProjectileHP
End Sub

Private Sub sclProjectileHP_Scroll()
sclProjectileHP_Change
End Sub

Private Sub sclProjectilePlus_Change()
lblProjectilePlus = sclProjectilePlus
End Sub

Private Sub sclProjectilePlus_Scroll()
sclProjectilePlus_Change
End Sub

Private Sub sclProjectileRange_Change()
lblProjectileRange = sclProjectileRange
End Sub

Private Sub sclProjectileRange_Scroll()
sclProjectileRange_Change
End Sub

Private Sub sclProjectileSpeed_Change()
lblProjectileSpeed = sclProjectileSpeed
End Sub

Private Sub sclProjectileSpeed_Scroll()
sclProjectileSpeed_Change
End Sub

Private Sub sclRingAmount_Change()
lblRingAmount = sclRingAmount
End Sub

Private Sub sclRingAmount_Scroll()
sclRingAmount_Change
End Sub

Private Sub sclRingHP_Change()
    lblRingHP = sclRingHP
End Sub

Private Sub sclRingHP_Scroll()
sclRingHP_Change
End Sub

Private Sub sclShieldDamagePercent_Change()
    lblShieldDamagePercent = sclShieldDamagePercent
End Sub

Private Sub sclShieldDefense_Change()
lblShieldDefense = sclShieldDefense
End Sub

Private Sub sclShieldDefense_Scroll()
sclShieldDefense_Change
End Sub

Private Sub sclShieldHP_Change()
lblShieldHP = sclShieldHP
End Sub

Private Sub sclShieldHP_Scroll()
sclShieldHP_Change
End Sub

Private Sub sclshieldmagicchance_Change()
lblshieldmagicchance = sclshieldmagicchance
End Sub

Private Sub sclshieldmagicpercent_Change()
lblshieldmagicpercent = sclshieldmagicpercent
End Sub

Private Sub sclStackSize_Change()
lblStackSize = sclStackSize
End Sub

Private Sub sclWeaponMin_Change()
lblWeaponMin = sclWeaponMin
End Sub

Private Sub sclWeaponMin_Scroll()
sclWeaponMin_Change
End Sub

Private Sub sclWeaponhp_Change()
lblWeaponHP = sclWeaponhp
End Sub

Private Sub sclWeaponhp_Scroll()
sclWeaponhp_Change
End Sub

Private Sub sclWeaponMax_Change()
lblWeaponMax = sclWeaponMax
End Sub

Private Sub sclWeaponMax_Scroll()
sclWeaponMax_Change
End Sub
