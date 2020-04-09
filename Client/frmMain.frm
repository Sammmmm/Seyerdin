VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H006BA6CE&
   BorderStyle     =   0  'None
   Caption         =   "Seyerdin Online"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPlayerScan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2250
      Left            =   2880
      ScaleHeight     =   2220
      ScaleWidth      =   6045
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   6075
      Begin VB.CommandButton cmdUpdateScan 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   1920
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton btnHideScan 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   57
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton btnHideItem 
         Caption         =   "Hide"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   4080
         Width           =   495
      End
      Begin VB.CommandButton btnToggleItems 
         Caption         =   "Show Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   50
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblScannedItemData 
         Caption         =   "2000000000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   4920
         TabIndex        =   70
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label lblScanData 
         Caption         =   "255.255.255.255"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   4560
         TabIndex        =   69
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblScannedItemData 
         Caption         =   "255"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   68
         Top             =   4440
         Width           =   375
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Storage5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   4200
         TabIndex        =   67
         Top             =   3720
         Width           =   720
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Storage4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   3360
         TabIndex        =   66
         Top             =   3720
         Width           =   720
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Storage3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   2520
         TabIndex        =   65
         Top             =   3720
         Width           =   720
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Storage2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   1680
         TabIndex        =   64
         Top             =   3720
         Width           =   720
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Storage1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   855
         TabIndex        =   63
         Top             =   3720
         Width           =   720
      End
      Begin VB.Label lblScanData 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   5520
         TabIndex        =   62
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblScanDesc 
         Caption         =   "AttackSpeed:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   4440
         TabIndex        =   61
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblScanData 
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   5520
         TabIndex        =   60
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblScanDesc 
         Caption         =   "Strength:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   2760
         TabIndex        =   59
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblScannedItemData 
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   55
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label lblScannedItemData 
         Caption         =   "Increased Chance of Sudden Infant Death Syndrome"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   54
         Top             =   4800
         Width           =   3855
      End
      Begin VB.Label lblScannedItemData 
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   53
         Top             =   4440
         Width           =   735
      End
      Begin VB.Label lblScannedItemData 
         Caption         =   "Increased attack velocity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   52
         Top             =   4440
         Width           =   3855
      End
      Begin VB.Label lblScannedItemData 
         Caption         =   "Furry Winged Eyes of the Caterpillar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   51
         Top             =   4080
         Width           =   3975
      End
      Begin VB.Label lblScanData 
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   5520
         TabIndex        =   49
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblScanData 
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   3840
         TabIndex        =   48
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblScanData 
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   3840
         TabIndex        =   47
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblScanData 
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   3840
         TabIndex        =   46
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblScanData 
         Caption         =   "1000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   3840
         TabIndex        =   45
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblScanData 
         Caption         =   "9999/9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   6
         Left            =   1320
         TabIndex        =   44
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblScanData 
         Caption         =   "9999/9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   43
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblScanData 
         Caption         =   "9999/9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   4
         Left            =   1320
         TabIndex        =   42
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblScanData 
         Caption         =   "255"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   41
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblScanData 
         Caption         =   "CLASS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   40
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblScanData 
         Caption         =   "PLAYERNAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   39
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblScanData 
         Caption         =   "USERNAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   38
         Top             =   120
         Width           =   2895
      End
      Begin VB.Line lineScan 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Index           =   10
         X1              =   6000
         X2              =   120
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line lineScan 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Index           =   9
         X1              =   6000
         X2              =   120
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line lineScan 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Index           =   8
         X1              =   6000
         X2              =   6000
         Y1              =   2280
         Y2              =   3630
      End
      Begin VB.Line lineScan 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Index           =   7
         X1              =   100
         X2              =   100
         Y1              =   2280
         Y2              =   3630
      End
      Begin VB.Line lineScan 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Index           =   6
         X1              =   6000
         X2              =   120
         Y1              =   3345
         Y2              =   3345
      End
      Begin VB.Line lineScan 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Index           =   5
         X1              =   6000
         X2              =   120
         Y1              =   3080
         Y2              =   3080
      End
      Begin VB.Line lineScan 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Index           =   4
         X1              =   6000
         X2              =   120
         Y1              =   2805
         Y2              =   2805
      End
      Begin VB.Line lineScan 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Index           =   3
         X1              =   6000
         X2              =   120
         Y1              =   2530
         Y2              =   2530
      End
      Begin VB.Line lineScan 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Index           =   2
         X1              =   4525
         X2              =   4525
         Y1              =   2280
         Y2              =   3630
      End
      Begin VB.Line lineScan 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Index           =   1
         X1              =   3050
         X2              =   3050
         Y1              =   2280
         Y2              =   3630
      End
      Begin VB.Line lineScan 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Index           =   0
         X1              =   1560
         X2              =   1560
         Y1              =   2280
         Y2              =   3630
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item20"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   4545
         TabIndex        =   37
         Top             =   3365
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item19"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   3070
         TabIndex        =   36
         Top             =   3365
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item18"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   1595
         TabIndex        =   35
         Top             =   3365
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item17"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   34
         Top             =   3365
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   4545
         TabIndex        =   33
         Top             =   3090
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   3070
         TabIndex        =   32
         Top             =   3090
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   1595
         TabIndex        =   31
         Top             =   3090
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item13"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   30
         Top             =   3090
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   4545
         TabIndex        =   29
         Top             =   2825
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   3070
         TabIndex        =   28
         Top             =   2825
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   1595
         TabIndex        =   27
         Top             =   2825
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   26
         Top             =   2825
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   4545
         TabIndex        =   25
         Top             =   2550
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   3070
         TabIndex        =   24
         Top             =   2550
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1595
         TabIndex        =   23
         Top             =   2550
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   2550
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4545
         TabIndex        =   21
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3070
         TabIndex        =   20
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1595
         TabIndex        =   19
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblScannedItem 
         Caption         =   "Item1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lblScanDesc 
         Caption         =   "Intelligence:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   4440
         TabIndex        =   17
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblScanDesc 
         Caption         =   "Constitution:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   4440
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblScanDesc 
         Caption         =   "Wisdom:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   2760
         TabIndex        =   15
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblScanDesc 
         Caption         =   "Endurance:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   2760
         TabIndex        =   14
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblScanDesc 
         Caption         =   "Agility:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   2760
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblScanDesc 
         Caption         =   "M/MaxM:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   12
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lblScanDesc 
         Caption         =   "E/MaxE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   11
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblScanDesc 
         Caption         =   "Hp/MaxHp:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblScanDesc 
         Caption         =   "Level:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblScanDesc 
         Caption         =   "Class:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblScanDesc 
         Caption         =   "Playername:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblScanDesc 
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox picMiniMap 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1830
      Left            =   9000
      ScaleHeight     =   122
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   174
      TabIndex        =   74
      Top             =   4440
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.PictureBox picChat 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawMode        =   16  'Merge Pen
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2325
      Left            =   120
      ScaleHeight     =   155
      ScaleMode       =   0  'User
      ScaleWidth      =   785
      TabIndex        =   73
      Top             =   6360
      Width           =   11775
   End
   Begin VB.PictureBox lstSkills 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   120
      ScaleHeight     =   180
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   190
      TabIndex        =   71
      Top             =   1200
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.PictureBox picDrop 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   3720
      ScaleHeight     =   1650
      ScaleWidth      =   4065
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox txtDrop 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         ForeColor       =   &H0080C0FF&
         Height          =   315
         Left            =   1080
         TabIndex        =   72
         Text            =   "0"
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblDropTitle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Drop how much?"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3840
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   35
         Left            =   2520
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblMenu 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   34
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.PictureBox picViewport 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   3090
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   0
      Top             =   555
      Width           =   5760
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

Public ChatSizeGotFocus As Boolean
Public CurMouseX As Long
Public CurMouseY As Long
Private MainWndProc As Long
Private Fullscreen As Long

Private Sub btnHideItem_Click()
frmMain.picPlayerScan.Height = 270
frmMain.picPlayerScan.Width = 410
End Sub

Private Sub btnHideScan_Click()
frmMain.picPlayerScan.Visible = False
End Sub

Private Sub btnToggleItems_Click()
If btnToggleItems.Caption = "Hide Items" Then
    frmMain.picPlayerScan.Height = 150
    frmMain.picPlayerScan.Width = 405
    btnToggleItems.Caption = "Show Items"
Else
    frmMain.picPlayerScan.Height = 270
    frmMain.picPlayerScan.Width = 410
    btnToggleItems.Caption = "Hide Items"
End If
End Sub


Private Sub cmdUpdateScan_Click()
Dim A As Long
A = FindPlayer(lblScanData(1))
If A >= 1 Then
    SendSocket Chr$(18) + Chr$(22) + Chr$(A)
Else
    MsgBox "The player has left the game!"
End If
End Sub

Private Sub Form_DblClick()
    Dim A As Long, b As Long
  '  If CurMouseY < 20 * WindowScaleY Then
   '     If Fullscreen Then
   '         Fullscreen = False
   '         ResizeGameWindow
   '     Else
   '         Fullscreen = True
   '         FullscreenGameWindow
   '     End If
   ' End If
        
    If CurMouseX > INVDestX And CurMouseX < INVDestX + INVWIDTH Then
        If CurMouseY > INVDestY And CurMouseY < INVDestY + INVHEIGHT Then
            If StorageOpen Then
                If CurInvObj > 0 And CurInvObj <= 20 Then
                    A = Character.Inv(CurInvObj).Object
                    If A Then
                        If Object(A).Type = 6 Or Object(A).Type = 11 Then
                            CurStorageObj = 0
                            TempVar2 = Character.Inv(CurInvObj).Value
                            TempVar1 = TempVar2
                            TempVar3 = CurInvObj
                            picDrop.Visible = True
                            lblDropTitle = "Deposit how much?"
                            txtDrop = TempVar1
                            txtDrop.SetFocus
                            txtDrop.SelStart = 0
                            txtDrop.SelLength = Len(txtDrop)
                            CurrentWindowFlags = CurrentWindowFlags Or WINDOW_FLAG_INVISIBLE
                        Else
                            SendSocket Chr$(75) + Chr$(1) + Chr$(LastNPCX) + Chr$(LastNPCY) + Chr$(CurInvObj) + QuadChar(0)
                        End If
                    End If
                ElseIf CurInvObj > 20 And CurInvObj <= 25 Then
                    PrintChat "You must unequip this item first.", 7, Options.FontSize
                End If
            ElseIf Character.Trading = True Then
                If CurInvObj > 0 And CurInvObj <= 20 Then
                    If Character.Inv(CurInvObj).Object > 0 Then
                        If Object(Character.Inv(CurInvObj).Object).Type = 6 Or Object(Character.Inv(CurInvObj).Object).Type = 11 Then
                            CurStorageObj = 0
                            TempVar2 = Character.Inv(CurInvObj).Value
                            TempVar1 = TempVar2
                            TempVar3 = CurInvObj
                            picDrop.Visible = True
                            lblDropTitle = "Trade how much?"
                            txtDrop = TempVar1
                            txtDrop.SetFocus
                            txtDrop.SelStart = 0
                            txtDrop.SelLength = Len(txtDrop)
                            CurrentWindowFlags = CurrentWindowFlags Or WINDOW_FLAG_INVISIBLE
                        Else
                            If ExamineBit(Object(Character.Inv(CurInvObj).Object).Flags, 4) Then
                                PrintChat "You cannot trade this object.", 7, Options.FontSize
                                Exit Sub
                            End If
                            
                            
                            b = 0
                            For A = 1 To 10
                                If TradeData.Slot(A) = CurInvObj Then
                                    b = 1
                                End If
                            Next A
                            If b = 0 Then
                                For A = 1 To 10
                                    If TradeData.YourObjects(A).Object = 0 And TradeData.Slot(A) = 0 Then
                                        b = A
                                        Exit For
                                    End If
                                Next A
                                If b > 0 And b <= 10 Then
                                    SendSocket Chr$(73) + Chr$(4) + Chr$(b) + Chr$(CurInvObj) + String$(4, 0)
                                    With Character.Inv(CurInvObj)
                                        TradeData.Slot(b) = CurInvObj
                                        TradeData.YourObjects(b).Object = .Object
                                        TradeData.YourObjects(b).Value = .Value
                                        TradeData.YourObjects(b).Prefix = .Prefix
                                        TradeData.YourObjects(b).PrefixValue = .PrefixValue
                                        TradeData.YourObjects(b).Suffix = .Suffix
                                        TradeData.YourObjects(b).SuffixValue = .SuffixValue
                                    End With
                                Else
                                    PrintChat "You can not currently trade any more items.  Try dropping some.", 7, Options.FontSize
                                End If
                            Else
                                PrintChat "You are already trading this object!", 7, Options.FontSize
                            End If
                        End If
                    End If
                End If
            Else
                If CurInvObj > 0 And CurInvObj <= 20 Then
                    With Character.Inv(CurInvObj)
                        If .Object > 0 Then
                            SendSocket Chr$(10) + Chr$(CurInvObj) 'Use Obj
                        Else
                            PrintChat "No such object.", 7, Options.FontSize, 15
                        End If
                    End With
                ElseIf CurInvObj > 20 And CurInvObj <= 25 Then
                    With Character.Equipped(CurInvObj - 20)
                        If .Object > 0 Then
                            If .Object > 0 Then
                                SendSocket Chr$(11) + Chr$(CurInvObj - 20) 'Stop Using Obj
                            Else
                                SendSocket Chr$(10) + Chr$(CurInvObj) 'Use Obj
                            End If
                        Else
                            PrintChat "No such object.", 7, Options.FontSize
                        End If
                    End With
                End If
            End If
            SetTab tsInventory
        End If
    End If
End Sub

Public Sub form_keydown(KeyCode As Integer, Shift As Integer)
If txtDrop.Visible Then Exit Sub

    Select Case KeyCode
        Case 112, 113, 114, 115, 116, 117, 118, 119, 120, 121 'F1-F10
            With Macro(KeyCode - 112)
                If ((Shift And vbAltMask)) Then
                    If fTime < GetTickCount Then
                        fTime = GetTickCount + IIf(Character.Access > 0, 250, 2500)
                        If .Text <> "" Then
                            ChatString = .Text
                            Chat.Enabled = True
                            DrawChatString
                            If .LineFeed = True Then
                                Form_KeyPress 200
                            End If
                        End If
                    End If
                Else
                    'If SkillListBox.MouseState = False Then
                      ''  Character.MacroSkill = 0
                      ''  If .Skill > 0 Then
                      ''      If Skills(.Skill).TargetType And TT_TILE Then
                      ''          CastingSpell = True
                      ''          Character.MacroSkill = .Skill
                      ''      Else
                      ''          UseSkill .Skill
                      ''      End If
                      ''  End If
                    'Else
                    'If SkillListBox.MouseState = True Then
                    '    With SkillListBox
                    '        .MouseState = True
                    '        DrawLstBox
                    '        If .Data(.Selected) > 0 Then
                    '            'Set Macro
                    '            c = 0
                    '            For b = vbKeyF1 To vbKeyF10
                    '                If GetKeyState(b) < 0 Then
                    '                    If .Selected > 0 Then
                    '                        If .Data(.Selected) <= MAX_SKILLS Then
                    '                            For A = 0 To 9
                    '                                If Macro(A).Skill = .Data(.Selected) Then
                    '                                    Macro(A).Skill = 0
                    '                                    c = 1
                    '                                End If
                    '                            Next A
                    '                            Macro(b - vbKeyF1).Skill = .Data(.Selected)
                    '                            c = 1
                    '                            DrawLstBox
                    '                        End If
                    '                        Exit For
                    '                    End If
                    '                End If
                    '            Next b
                    '            If c Then
                    '                SaveSkillMacros
                    '            End If
                    '        End If
                    '    End With
                    'End If
                End If
            End With
        Case 33 'PgUp
            If ChatScrollBack < 999 Then
                If Chat.AllChat(ChatScrollBack + 1).used Then
                    ChatScrollBack = ChatScrollBack + 1
                End If
                DrawChat
            Else
                'Beep
            End If
        Case 34 'PgDn
            If ChatScrollBack > 0 Then
                ChatScrollBack = ChatScrollBack - 1
                DrawChat
            Else
                'Beep
            End If
        Case 18 'Alt
            keyAlt = True
    End Select
End Sub

Public Sub Form_KeyPress(KeyAscii As Integer)
If txtDrop.Visible Then Exit Sub
    Dim A As Long, b As Long, C As Long
    Dim St1 As String
If WidgetKeyDown(Widgets, KeyAscii) Then
    If KeyAscii >= 32 And KeyAscii <= 127 Then
        If ChatString = "" Then
            If KeyAscii = KeyCodeList(Options.BroadcastKey).KeyCode Or KeyAscii = KeyCodeList(Options.BroadcastKey).CapitalKeyCode Then ' ;
                Chat.Enabled = True
                ChatString = "/BROADCAST "
                DrawChatString
            ElseIf KeyAscii = KeyCodeList(Options.TellKey).KeyCode Or KeyAscii = KeyCodeList(Options.TellKey).CapitalKeyCode Then ' '
                Chat.Enabled = True
                ChatString = "/TELL "
                DrawChatString
            ElseIf KeyAscii = KeyCodeList(Options.GuildKey).KeyCode Or KeyAscii = KeyCodeList(Options.GuildKey).CapitalKeyCode Then '\'
                Chat.Enabled = True
                ChatString = "/GUILD CHAT "
                DrawChatString
            ElseIf KeyAscii = KeyCodeList(Options.PartyKey).KeyCode Or KeyAscii = KeyCodeList(Options.PartyKey).CapitalKeyCode Then '\'
                Chat.Enabled = True
                ChatString = "/PARTY CHAT "
                DrawChatString
            ElseIf KeyAscii = KeyCodeList(Options.SayKey).KeyCode Or KeyAscii = KeyCodeList(Options.SayKey).CapitalKeyCode Then  '\'
                Chat.Enabled = True
                ChatString = "/SAY "
                DrawChatString
            ElseIf KeyAscii = 47 Then '/
                Chat.Enabled = True
                ChatString = Chr$(KeyAscii)
                DrawChatString
            Else
                If Chat.Enabled Then
                    ChatString = Chr$(KeyAscii)
                    DrawChatString
                End If
            End If
        Else
            If Chat.Enabled Then
                If Len(ChatString) < 255 Then 'add
                    ChatString = ChatString + Chr$(KeyAscii)
                    DrawChatString
                Else
                    'Beep
                End If
            End If
        End If
    ElseIf KeyAscii = 8 Then 'delete
        If Chat.Enabled Then
            If Len(ChatString) > 0 Then
                ChatString = Left$(ChatString, Len(ChatString) - 1)
                DrawChatString
            Else
                'Beep
            End If
        End If
    ElseIf KeyAscii = 27 Then
        Chat.Enabled = False
        ChatString = ""
        DrawChatString
        CurrentTarget.TargetType = TT_NO_TARGET
        CurrentTarget.Target = 0
    ElseIf KeyAscii = KeyCodeList(Options.ChatKey).KeyCode Or KeyAscii = KeyCodeList(Options.ChatKey).CapitalKeyCode Or KeyAscii = 200 Then 'KeyAscii = 10 Or
            
        
             If (Options.PickupKey = Options.ChatKey) Then
                    If Character.Trading = False And KeyAscii <> 200 Then
                        'Pick up object
                        For A = 0 To 49
                            With map.Object(A)
                                If .Object > 0 And .x = cX And .y = cY Then
                                    
                                        SendSocket Chr$(8)
                                        'Chat.Enabled = False
                                        'DrawChatString
                                    If Not Chat.Enabled Then
                                        Exit Sub
                                    End If
                                End If
                            End With
                        Next A
                    End If
                End If

    
        If Not Chat.Enabled And KeyAscii <> 200 Then
            Chat.Enabled = True
            DrawChatString
        Else
            'If Options.AltKeysEnabled Then
                Chat.Enabled = False
                DrawChatString
            'End If
            If ChatString <> "" Then
                If CMap <= 5000 Then ChatString = Replace(ChatString, "%MAP+1", CMap + 1) Else ChatString = Replace(ChatString, "%MAP+1", 5000)
                ChatString = Replace(ChatString, "%MAP", CMap)
                If cX < 11 Then ChatString = Replace(ChatString, "%X+1", cX + 1) Else ChatString = Replace(ChatString, "%X+1", 11)
                ChatString = Replace(ChatString, "%X", cX)
                If cY < 11 Then ChatString = Replace(ChatString, "%Y+1", cY + 1) Else ChatString = Replace(ChatString, "%Y+1", 11)
                ChatString = Replace(ChatString, "%Y", cY)
                ChatString = Replace(ChatString, "%LOCATION", "[" + CStr(CMap) + "," + CStr(cX) + "," + CStr(cY) + "]")
                ChatString = Replace(ChatString, "%LASTKILLER", CLastKiller)
                ChatString = Replace(ChatString, "%LASTKILLED", CLastKilled)
                If Character.Guild > 0 Then ChatString = Replace(ChatString, "%GUILD", Guild(Character.Guild).Name) Else ChatString = Replace(ChatString, "%GUILD", "")
                ChatString = Replace(ChatString, "%NAME", Character.Name)
                
                If UCase$(Left$(ChatString, 7)) = "/RENOWN" Or UCase$(Left$(ChatString, 5)) = "/TIPS" Or UCase$(Left$(ChatString, 5)) = "/TASK" Then
                
                Else
                            If UCase$(Left$(ChatString, 3)) = "/P " Then ChatString = "/PARTY CHAT" + Mid$(ChatString, 3)
                            If UCase$(Left$(ChatString, 3)) = "/G " Then ChatString = "/GUILD CHAT" + Mid$(ChatString, 3)
                            If UCase$(Left$(ChatString, 1)) <> "/" Then
                                ChatString = "/s " + ChatString
                            End If
                End If
                
                If Left$(ChatString, 1) = "/" Then
                    If Len(ChatString) > 1 Then
                        GetSections Mid$(ChatString, 2), 1
                        Select Case UCase$(Section(1))
                            Case "R", "RE", "REP", "REPL", "REPLY"
                                    If LastPlayerTellNum > 0 Then
                                        If player(LastPlayerTellNum).Name = LastPlayerTellName Then
                                            If Character.Squelched = False Then
                                                A = LastPlayerTellNum
                                                If A > 0 Then
                                                    If Suffix <> "" Then
                                                        If Character.Mana >= 2 Then
                                                            SendSocket Chr$(14) + Chr$(A) + Suffix
                                                            PrintChat "You tell " + player(A).Name + ", " + Chr$(34) + SwearFilter(Suffix) + Chr$(34), 10, Options.FontSize, 15
                                                        Else
                                                            PrintChat "You do not have enough mana to tell!", 14, Options.FontSize, 15
                                                        End If
                                                    Else
                                                        PrintChat "What do you want to tell " + player(A).Name + "?", 14, Options.FontSize, 15
                                                    End If
                                                Else
                                                    PrintChat "That player isn't here any more.", 14, Options.FontSize, 15
                                                End If
                                            Else
                                              PrintChat "You are squelched!", 14, Options.FontSize, 15
                                            End If
                                        End If
                                    Else
                                        PrintChat "You have no one to reply to.", 14, Options.FontSize, 15
                                    End If
                            Case "LIGHT"
                                If Character.Access Then
                                GetSections Suffix, 3
                                    LightSource(0).Red = (Val(Section(1)) And 255)
                                    LightSource(0).Green = (Val(Section(2)) And 255)
                                    LightSource(0).Blue = (Val(Section(3)) And 255)
                                End If
                            Case "UNLOADGRAPHICS"
                                For A = 1 To NumTextures
                                    Set DynamicTextures(A).Texture = Nothing
                                    DynamicTextures(A).Loaded = False
                                    DynamicTextures(A).LastUsed = 0
                                Next A
                            Case "TIME"
                                If World.Hour > 12 Then
                                    PrintChat "The time in Charon is approximately " & CStr(World.Hour - 12) & ":00 PM", 15, Options.FontSize
                                Else
                                    PrintChat "The time in Charon is approximately " & CStr(World.Hour) & ":00 AM", 15, Options.FontSize
                                End If
                            Case "HELP"
                                GetSections Suffix, 1
                                Select Case UCase$(Section(1))
                                    Case "OPTIONS"
                                        PrintChat "Options command: Allows you to change several game options.", 14, Options.FontSize, 15
                                        PrintChat "SYNTAX: /Options", 14, Options.FontSize, 15
                                    Case "MACROS"
                                        PrintChat "Macros command: Allows you create macros that are executed with the function keys.", 14, Options.FontSize, 15
                                        PrintChat "SYNTAX: /Macros", 14, Options.FontSize, 15
                                    Case "BUY", "TRADE", "SELL"
                                        PrintChat "Trade command: Allows you to buy and sell goods from the nearest NPC.", 14, Options.FontSize, 15
                                        PrintChat "SYNTAX: /Buy", 14, Options.FontSize, 15
                                    Case "GUILD"
                                        PrintChat "Guild command: Offers a variety of guild related functions.  Type '/guild help' for more information.", 14, Options.FontSize, 15
                                        PrintChat "SYNTAX: /Guild <command>", 14, Options.FontSize, 15
                                    Case "BALANCE"
                                        PrintChat "Balance command: Tells you how much gold is in your bank account.  May only be used from inside a bank.", 14, Options.FontSize, 15
                                        PrintChat "SYNTAX: /Balance", 14, Options.FontSize, 15
                                    Case "WITHDRAW"
                                        PrintChat "Withdraw command: Allows you to withdraw gold from your bank account.  May only be used from inside a bank.", 14, Options.FontSize, 15
                                        PrintChat "SYNTAX: /Withdraw <amount>", 14, Options.FontSize, 15
                                    Case "DEPOSIT"
                                        PrintChat "Deposit command: Allows you to deposit gold into your bank account.  May only be used from inside a bank.", 14, Options.FontSize, 15
                                        PrintChat "SYNTAX: /Deposit <amount>", 14, Options.FontSize, 15
                                    Case "GUILDS"
                                        PrintChat "Guilds command: Lists all of the guilds in the game.", 14, Options.FontSize, 15
                                        PrintChat "SYNTAX: /Guilds", 14, Options.FontSize, 15
                                    'Case "DESCRIBE"
                                    '    PrintChat "Describe command: Allows you to write a description for your character.  Others may read this by left clicking on your character sprite.", 14, Options.FontSize, 15
                                    '    PrintChat "SYNTAX: /DESCRIBE <text>", 14, Options.FontSize, 15
                                    Case "BROADCAST"
                                        PrintChat "Broadcast command: Lets you talk to everyone in the game at once.  Uses energy.", 14, Options.FontSize, 15
                                        PrintChat "SYNTAX: /BROADCAST <message>", 14, Options.FontSize, 15
                                    Case "EMOTE"
                                        PrintChat "Emote command: Lets you describe what you are doing.", 14, Options.FontSize, 15
                                        PrintChat "SYNTAX: /EMOTE <actions>", 14, Options.FontSize, 15
                                    Case "SAY"
                                        PrintChat "Say command: Lets you talk to everyone on the map.", 14, Options.FontSize, 15
                                        PrintChat "SYNTAX: /SAY <message>", 14, Options.FontSize, 15
                                    Case "TELL"
                                        PrintChat "Tell command: Lets you describe what you are doing.", 14, Options.FontSize, 15
                                        PrintChat "SYNTAX: /TELL <player> <message>", 14, Options.FontSize, 15
                                    Case "WHERE"
                                         PrintChat "Where command: Gives you the coordinates of your current location.  Useful for reporting map bugs to gods.", 14, Options.FontSize, 15
                                         PrintChat "SYNTAX: /Where", 14, Options.FontSize, 15
                                    Case "WHO"
                                        PrintChat "Who command: Shows you who is online.", 14, Options.FontSize, 15
                                        PrintChat "SYNTAX: /WHO", 14, Options.FontSize, 15
                                    Case "YELL"
                                        PrintChat "Yell command: Lets you talk to everyone on the map and nearby maps.", 14, Options.FontSize, 15
                                        PrintChat "SYNTAX: /YELL <message>", 14, Options.FontSize, 15
                                    Case "MESSAGES"
                                        PrintChat "Messages command: Lets you view your logged messages while being 'away' (used in conjunction with '/AMSG')", 14, Options.FontSize, 15
                                        PrintChat "SYNTAX: /MESSAGES", 14, Options.FontSize, 15
                                    Case "AMSG"
                                        PrintChat "Away Message command: This sets your away message to be displayed while option 'Away' is checked and a user '/tells' you.", 14, Options.FontSize, 15
                                        PrintChat "SYNTAX: /AMSG <message>", 14, Options.FontSize, 15
                                    Case "PLAYERTRADE"
                                        PrintChat "Player Trade command:Trade items with other players.", 14, Options.FontSize
                                        PrintChat "SYNTAX: '/PLAYERTRADE <player>' to initiate a trade or '/PLAYERTRADE CANCEL' to cancel a request or '/PLAYERTRADE ACCEPT' to accept a request", 14, Options.FontSize, 15
                                    Case ""
                                        PrintChat "Available Commands: BALANCE BROADCAST DEPOSIT EMOTE GUILD GUILDS HELP MACROS OPTIONS SAY STATS TELL TRADE TRAIN WHERE WITHDRAW WHO YELL AMSG MESSAGES", 14, Options.FontSize, 15
                                    Case Else
                                        PrintChat "No such command!", 14, Options.FontSize, 15
                                End Select
                                
                            Case "OPTIONS"
                                frmOptions.Show
                            
                            Case "DEBUG"
                                'PrintChat D3DInitType, 14, Options.FontSize, 15
                            
                            Case "MAIN", "MAI", "MA", "M"
                                    If Suffix <> "" Then
                                        If Character.Squelched = False Then
                                            SendSocket Chr$(15) + Chr$(0) + Suffix
                                            PrintChat Character.Name + ": " + SwearFilter(Suffix), 13, Options.FontSize, 0
                                        Else
                                            PrintChat "You are squelched!", 14, Options.FontSize, 15
                                        End If
                                    Else
                                        PrintChat "What do you want to main-broadcast?", 14, Options.FontSize, 15
                                    End If
    
                            
                            
                            Case "MACROS"
                                frmMacros.Show
                                
                            Case "AWAYMSG", "AMSG", "AMESSAGE", "AWAY"
                                If Suffix <> "" Then
                                    Options.AwayMsg = SwearFilter(Suffix)
                                    WriteString "Options", "Amsg", Options.AwayMsg
                                    PrintChat "Your away message has been set, use type '/options' to turn away on!", 14, Options.FontSize, 15
                                Else
                                    PrintChat "Your away message must not be blank!", 14, Options.FontSize, 15
                                    Options.AwayMsg = "I am currently AFK. I will be with you soon :)"
                                    WriteString "Options", "Amsg", Options.AwayMsg
                                End If
                            Case "MSGS", "MSG", "MESSAGES", "LOG"
                                frmLog.Show

                            Case "BROADCAST", "BROADCAS", "BROADCA", "BROADC", "BROAD", "BROA", "BRO", "BR", "B"
                                If Suffix <> "" Then
                                  If Character.Squelched = False Then
                                    If LastBroadcast = vbNullString Then
                                        LastBroadcast = SwearFilter(Suffix)
                                        SendSocket Chr$(15) + Chr$(0) + LastBroadcast
                                        'PrintChat Character.Name + ": " + SwearFilter(Suffix), RGB(200, 0, 220), Options.FontSize, Character.CurChannel, True
                                        If Character.Access = 0 Then
                                        '    PrintChat Character.Name + ": " + SwearFilter(Suffix), RGB(230, 0, 230), Options.FontSize, 0, True
                                        Else
                                        '    PrintChat Character.Name + ": " + SwearFilter(Suffix), RGB(195, 0, 120), Options.FontSize, 0, True
                                        End If
                                    Else
                                        PrintChat "Please wait to send another message!", 14, Options.FontSize
                                    End If
                                  Else
                                    PrintChat "You are squelched!", 14, Options.FontSize
                                  End If
                                Else
                                    PrintChat "What do you want to broadcast?", 14, Options.FontSize, 15
                                End If
                                
                            'Case "DESCRIBE", "DESCRIB", "DESCRI", "DESCR", "DESC", "DES", "DE", "D"
                            '    If Len(Suffix) > 0 Then
                            '        SendSocket Chr$(28) + SwearFilter(Suffix)
                            '        PrintChat "Your description has been changed.", 14, Options.FontSize, 15
                            '    Else
                            '        PrintChat "You must enter a description.", 14, Options.FontSize, 15
                            '    End If
                                
                            Case "EMOTE", "EMOT", "EMO", "EM", "E"
                                If Suffix <> "" Then
                                  If Character.Squelched = False Then
                                    SendSocket Chr$(16) + Suffix
                                    PrintChat Character.Name + " " + SwearFilter(Suffix), 3, Options.FontSize
                                  Else
                                    PrintChat "You are squelched!", 14, Options.FontSize, 3
                                  End If
                                Else
                                    PrintChat "What do you want to do?", 14, Options.FontSize, 3
                                End If
                                
                            Case "SAY", "SA", "S"
                                If Character.Squelched = False Then
                                    If Suffix <> "" Then
                                        SendSocket Chr$(6) + Chr$(0) + Suffix
                                        PrintChat "You say, " + Chr$(34) + SwearFilter(Suffix) + Chr$(34), 7, Options.FontSize, 1
                                    Else
                                        PrintChat "What do you want to say?", 14, Options.FontSize, 1
                                    End If
                                Else
                                    PrintChat "You are squelched!", 14, Options.FontSize
                                End If
    
                            Case "IGNORE"
                                If Suffix <> "" Then
                                    A = FindPlayer(Suffix)
                                    If A > 0 Then
                                        With player(A)
                                            If .Ignore = True Then
                                                .Ignore = False
                                                PrintChat "You are no longer ignoring " + .Name + ".", 14, Options.FontSize, 15
                                            Else
                                                .Ignore = True
                                                PrintChat "You are now ignoring " + .Name + ".", 14, Options.FontSize, 15
                                            End If
                                        End With
                                    Else
                                        PrintChat "No such player!", 14, Options.FontSize
                                    End If
                                Else
                                    St1 = ""
                                    b = 0
                                    For A = 1 To MAXUSERS
                                        With player(A)
                                            If .Sprite > 0 And .Ignore = True Then
                                                b = b + 1
                                                St1 = St1 + ", " + .Name
                                            End If
                                        End With
                                    Next A
                                    If b > 0 Then
                                        St1 = Mid$(St1, 2)
                                        PrintChat "You are currently ignoring " + CStr(b) + " people:" + St1, 14, Options.FontSize, 15
                                    Else
                                        PrintChat "You are not ignoring anybody!", 14, Options.FontSize, 15
                                    End If
                                End If
                                
                            Case "TELL", "TEL", "TE", "T"
                                If Suffix <> "" Then
                                  If Character.Squelched = False Then
                                    GetSections Suffix, 1
                                    A = FindPlayer(Section(1))
                                       If A > 0 Then
                                            If Suffix <> "" Then
                                                If Character.Mana >= 2 Then
                                                    SendSocket Chr$(14) + Chr$(A) + Suffix
                                                    PrintChat "You tell " + player(A).Name + ", " + Chr$(34) + SwearFilter(Suffix) + Chr$(34), 10, Options.FontSize, 2
                                                Else
                                                    PrintChat "You do not have enough mana to tell!", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                PrintChat "What do you want to tell " + player(A).Name + "?", 14, Options.FontSize, 2
                                            End If
                                        Else
                                            PrintChat "No such player!", 14, Options.FontSize, 15
                                        End If
    
                                  Else
                                    PrintChat "You are squelched!", 14, Options.FontSize, 15
                                  End If
                                Else
                                    PrintChat "What, and to whom, do you want to tell?", 14, Options.FontSize, 15
                                End If
                            Case "WHERE", "WHER", "WHE"
                                If Character.Access > 0 Then
                                    PrintChat "You are at location [" + CStr(CMap) + ", " + CStr(cX) + ", " + CStr(cY) + "]", 14, Options.FontSize, 15
                                End If
                            Case "WHO", "WH", "W"
                                St1 = ""
                                b = 0
                                For A = 1 To MAXUSERS
                                    With player(A)
                                        If .Sprite > 0 And A <> Character.Index Then
                                            b = b + 1
                                            St1 = St1 + ", " + .Name
                                        End If
                                    End With
                                Next A
                                If b > 0 Then
                                    St1 = Mid$(St1, 2)
                                    PrintChat "There are " + CStr(b) + " other players online: " + St1, 14, Options.FontSize, 15
                                Else
                                    PrintChat "There are no other players online.", 14, Options.FontSize, 15
                                End If
                            Case "YELL", "YEL", "YE", "Y"
                                If Suffix <> "" Then
                                  If Character.Squelched = False Then
                                    SendSocket Chr$(17) + Chr$(0) + Suffix
                                    PrintChat "You yell, """ + SwearFilter(Suffix) + Chr$(34), 7, Options.FontSize, 4
                                  Else
                                    PrintChat "You are squelched!", 14, Options.FontSize
                                  End If
                                Else
                                    PrintChat "What do you want to yell?", 14, Options.FontSize, 4
                                End If
                                
                            Case "FRAMERATE", "FRAMERAT", "FRAMERA", "FRAMER", "FRAME", "FRAM", "FRA", "FR", "F"
                                PrintChat "Current Frame Rate: " + CStr(FrameRate), 14, Options.FontSize
                            Case "GUILDS"
                                frmGuilds.Show
                                With frmGuilds
                                .nextTab = 0
                                .GuildRank = Character.GuildRank
                                .playerGuild = Character.Guild
                                .changed = True
                                .ctab = 0
                                .setGuildTab
                                .lstEnabled = True
                                End With
                                
                            Case "GUILD"
                                GetSections Suffix, 1
                                Select Case UCase$(Section(1))
                                    Case "BUY"
                                        If Character.Guild > 0 And Character.GuildRank >= 2 Then
                                            SendSocket Chr$(43)
                                        Else
                                            PrintChat "You must be an Officer of a guild to use that command.", 14, Options.FontSize
                                        End If
                                        
                                    Case "WHO"
                                        If Character.Guild > 0 Then
                                            St1 = ""
                                            b = 0
                                            For A = 1 To MAXUSERS
                                                With player(A)
                                                    If .Sprite > 0 And A <> Character.Index And .Guild = Character.Guild Then
                                                        b = b + 1
                                                        St1 = St1 + ", " + .Name
                                                    End If
                                                End With
                                            Next A
                                            If b > 0 Then
                                                St1 = Mid$(St1, 2)
                                                PrintChat "There are " + CStr(b) + " other guild members online: " + St1, 14, Options.FontSize, 5
                                            Else
                                                PrintChat "There are no other guild members online.", 14, Options.FontSize, 5
                                            End If
                                        Else
                                            PrintChat "You are not in a guild!", 14, Options.FontSize
                                        End If
                                        
                                    Case "CHAT"
                                        If Character.Guild > 0 Then
                                            If Suffix <> "" Then
                                                SendSocket Chr$(41) + Suffix 'Guild Chat
                                                PrintChat Character.Name + " -> Guild: " + Suffix, CHannelColors(4), Options.FontSize, 4, True
                                            Else
                                                PrintChat "You must specify a message!", CHannelColors(4), Options.FontSize, 5, True
                                            End If
                                        Else
                                            PrintChat "You are not in a guild!", 14, Options.FontSize, 15
                                        End If
                                        
                                    Case "INVITE"
                                        If Character.Guild > 0 And Character.GuildRank >= 2 Then
                                            GetSections Suffix, 1
                                            If Section(1) <> "" Then
                                                A = FindPlayer(Section(1))
                                                If A > 0 Then
                                                    SendSocket Chr$(34) + Chr$(A)
                                                    PrintChat player(A).Name + " has been invited to join your guild.", 15, Options.FontSize, 5
                                                Else
                                                    PrintChat "No such player!", 14, Options.FontSize, 5
                                                End If
                                            Else
                                                PrintChat "Must specify a name.", 14, Options.FontSize, 5
                                            End If
                                        Else
                                            PrintChat "You must be an Officer of a guild to use that command.", 14, Options.FontSize, 15
                                        End If
                                        
                                    Case "JOIN"
                                        If Character.Trading = False Then
                                            If Character.Guild = 0 Then
                                                If Character.Level >= 10 Then
                                                    If Character.Energy <> Character.MaxEnergy Or Character.HP <> Character.MaxHP Then
                                                        PrintChat "You must have full HP and Energy to join a guild.", 14, Options.FontSize, 15
                                                    Else
                                                         SendSocket Chr$(31) 'Join Guild
                                                    End If
                                                Else
                                                    PrintChat "You must be level 10 to join a guild!", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                PrintChat "You are already in a guild.  If you would like to join a new guild, you must first leave this guild by typing '/guild leave'.", 14, Options.FontSize, 15
                                            End If
                                        Else
                                            PrintChat "You cannot do that while trading!", 14, Options.FontSize, 15
                                        End If
                                    Case "LEAVE"
                                        If Character.Guild > 0 Then
                                            If Character.HP = Character.MaxHP And Character.Energy = Character.MaxEnergy Then
                                                SendSocket Chr$(32) 'Leave Guild
                                            Else
                                                PrintChat "You must have full HP and Energy to leave a guild.", 14, Options.FontSize, 15
                                            End If
                                        Else
                                            PrintChat "You are not in a guild!", 14, Options.FontSize, 15
                                        End If
                                    
                                    Case "PAY"
                                        If Character.Trading = False Then
                                            If Character.Guild > 0 Then
                                                'If Map.NPC > 0 Then
                                                    GetSections Suffix, 1
                                                    If Section(1) <> "" Then
                                                        If CDbl(Int(Val(Section(1)))) <= 2147483647# Then
                                                            A = Int(Val(Section(1)))
                                                            If A >= 0 Then
                                                                SendSocket Chr$(40) + QuadChar(A) 'Pay balance
                                                            Else
                                                                PrintChat "Invalid value.", 14, Options.FontSize, 5
                                                            End If
                                                        Else
                                                            PrintChat "Invalid value.", 14, Options.FontSize, 5
                                                        End If
                                                    Else
                                                        PrintChat "You must specify an amount!", 14, Options.FontSize, 5
                                                    End If
                                                'Else
                                                '    PrintChat "You must be in a bank!", 14, Options.FontSize
                                                'End If
                                            Else
                                                PrintChat "You are not in a guild!", 14, Options.FontSize, 15
                                            End If
                                        Else
                                            PrintChat "You cannot do that while trading!", 14, Options.FontSize, 15
                                        End If
                                    Case "EDIT"
                                        If Character.Guild > 0 Then
                                            SendSocket Chr$(39) + Chr$(Character.Guild)
                                        Else
                                            PrintChat "You are not in a guild!", 14, Options.FontSize, 15
                                        End If
                                        
                                    Case "HALLINFO"
                                        SendSocket Chr$(47)
                                        
                                    Case "BALANCE"
                                        If Character.Guild > 0 Then
                                            SendSocket Chr$(46)
                                        Else
                                            PrintChat "You are not in a guild!", 14, Options.FontSize, 15
                                        End If
                                    
                                    Case "HELP"
                                        PrintChat "Guild commands: BALANCE BUY CHAT HALLINFO INVITE JOIN LEAVE NEW PAY EDIT", 14, Options.FontSize, 5
                                
                                    
                                    Case Else
                                        PrintChat "Invalid guild command.", 14, Options.FontSize, 5
                                End Select
                                
                            Case "GOD"
                                If Character.Access > 0 Then
                                    GetSections Suffix, 1
                                    Select Case UCase$(Section(1))
                                        Case "CHAT"
                                            If Character.Access > 0 Then
                                                If Suffix <> "" Then
                                                    SendSocket Chr$(18) + Chr$(14) + Suffix
                                                    PrintChat "<" + Character.Name + ">: " + Suffix, 11, Options.FontSize, 15
                                                End If
                                            End If
                                        Case "BOOT"
                                            If Character.Access >= 4 Then
                                                GetSections Suffix, 1
                                                If Section(1) <> "" Then
                                                    A = FindPlayer(Section(1))
                                                    If A >= 1 Then
                                                        SendSocket Chr$(18) + Chr$(9) + Chr$(A) + Suffix
                                                    Else
                                                        PrintChat "Boot: No such player", 14, Options.FontSize, 15
                                                    End If
                                                Else
                                                    PrintChat "Boot: Not enough parameters", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        Case "SQUELCH"
                                            If Character.Access >= 8 Then
                                                GetSections Suffix, 2
                                                If Section(1) <> "" Then
                                                    A = FindPlayer(Section(1))
                                                    If A >= 1 Then
                                                        If Val(Section(2)) >= 1 And Val(Section(2)) <= 255 Then
                                                            SendSocket Chr$(18) + Chr$(19) + Chr$(A) + Chr$(Val(Section(2)))
                                                        Else
                                                            PrintChat "Squelch: Invalid Time! (1-255 seconds)", 14, Options.FontSize, 15
                                                        End If
                                                    Else
                                                        PrintChat "Squelch: No such player", 14, Options.FontSize, 15
                                                    End If
                                                Else
                                                    SendSocket Chr$(18) + Chr$(19)
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        Case "UNSQUELCH"
                                            If Character.Access >= 8 Then
                                                GetSections Suffix, 1
                                                If Section(1) <> "" Then
                                                    A = FindPlayer(Section(1))
                                                    If A >= 1 Then
                                                        SendSocket Chr$(18) + Chr$(19) + Chr$(A) + Chr$(0)
                                                    Else
                                                        PrintChat "Unsquelch: No such player", 14, Options.FontSize, 15
                                                    End If
                                                Else
                                                    PrintChat "Unsquelch: Not enough parameters", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        Case "BAN"
                                            If Character.Access >= 4 Then
                                                GetSections Suffix, 2
                                                If Section(2) <> "" Then
                                                    A = FindPlayer(Section(1))
                                                    If A >= 1 Then
                                                        b = Int(Val(Section(2)))
                                                        If b >= 1 And b <= 255 Then
                                                            SendSocket Chr$(18) + Chr$(10) + Chr$(A) + Chr$(b) + Suffix
                                                        Else
                                                            PrintChat "Ban: Unnacceptable number of days", 14, Options.FontSize, 15
                                                        End If
                                                    Else
                                                        PrintChat "Ban: No such player", 14, Options.FontSize, 15
                                                    End If
                                                Else
                                                    PrintChat "Ban: Nt enough parameters", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        Case "DISBAND", "DISBAN", "DELETE", "REMOVE"
                                            If Character.Access >= 10 Then
                                                GetSections Suffix, 1
                                                A = FindGuild(Section(1))
                                                If A >= 1 Then
                                                    SendSocket Chr$(18) + Chr$(5) + Chr$(A)
                                                    PrintChat "Guild " + Guild(A).Name + " disbanded!", 14, Options.FontSize, 15
                                                Else
                                                    PrintChat "No such guild!", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        Case "RESETMAP"
                                            If Character.Access >= 4 Then
                                                SendSocket Chr$(18) + Chr$(8)
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                            
                                        Case "TP"
                                            If Character.Access >= 4 Then
                                                SendSocket Chr$(18) + Chr$(18)
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        
                                        Case "PERMABAN"
                                            If Character.Access >= 10 Then
                                                GetSections Suffix, 1
                                                If Trim$(Section(1)) <> "" Then
                                                    A = FindPlayer(Section(1))
                                                    If A >= 1 Then
                                                        SendSocket Chr$(18) + Chr$(16) + Chr$(A)
                                                    End If
                                                End If
                                            End If
                                                
                                        Case "SHUTDOWN"
                                            If Character.Access = 10 Then
                                                SendSocket Chr$(18) + Chr$(13)
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                                                    
                                        Case "FORWARD", "FUSER", "FORWARDU"
                                            If Character.Access >= 6 Then
                                                If Options.Away = False Then
                                                    GetSections Suffix, 1
                                                    If Section(1) <> "" And FindPlayer(Section(1)) <> 0 Then
                                                        Options.ForwardUser = Section(1)
                                                        PrintChat "Forwarding all '/tells' to " + Section(1), 14, Options.FontSize, 15
                                                        Options.Away = False
                                                    Else
                                                        Options.ForwardUser = ""
                                                        PrintChat "User not found, forward messages is OFF!", 14, Options.FontSize, 15
                                                    End If
                                                Else
                                                    PrintChat "Message forwarding cannot be used while beiung 'Away'. Please use the options menu to turn off away.", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        Case "MOTD", "MOT", "MO", "M"
                                            If Character.Access >= 9 Then
                                                If Suffix <> "" Then
                                                    SendSocket Chr$(18) + Chr$(4) + Suffix
                                                    PrintChat "MOTD changed.", 14, Options.FontSize
                                                Else
                                                    PrintChat "You must specify a message!", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        Case "SETNAME", "SETNAM", "SETNA", "SETN"
                                            If Character.Access >= 8 Then
                                                GetSections Suffix, 1
                                                A = FindPlayer(Section(1))
                                                If A >= 1 Then
                                                    If Len(Suffix) >= 3 And Len(Suffix) <= 15 Then
                                                        SendSocket Chr$(18) + Chr$(7) + Chr$(A) + Suffix
                                                        PrintChat player(A).Name + "'s name changed.", 14, Options.FontSize, 15
                                                    Else
                                                        PrintChat "Name may be no longer than 15 characters!", 14, Options.FontSize, 15
                                                    End If
                                                Else
                                                    PrintChat "No such player!", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        Case "SETMYNAME", "SETMYNAM", "SETMYNA", "SETMYN"
                                            If Character.Access >= 8 Then
                                                If Len(Suffix) > 0 Then
                                                    If Len(Suffix) <= 15 Then
                                                        SendSocket Chr$(18) + Chr$(7) + Chr$(Character.Index) + Suffix
                                                        PrintChat "Name changed.", 14, Options.FontSize, 15
                                                    Else
                                                        PrintChat "Name may be no longer than 15 characters!", 14, Options.FontSize, 15
                                                    End If
                                                Else
                                                    PrintChat "You must specify a new name!", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        Case "SETSPRITE"
                                            If Character.Access >= 8 Then
                                                GetSections Suffix, 2
                                                A = FindPlayer(Section(1))
                                                If A >= 1 Then
                                                    b = Int(Val(Section(2)))
                                                    If b >= 0 And b <= 255 Then
                                                        SendSocket Chr$(18) + Chr$(6) + Chr$(A) + Chr$(b)
                                                        PrintChat player(A).Name + "'s sprite has been changed.", 14, Options.FontSize, 15
                                                    Else
                                                        PrintChat "Invalid sprite number!", 14, Options.FontSize, 15
                                                    End If
                                                Else
                                                    PrintChat "No such player!", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        Case "SETCLASS"
                                            If Character.Access >= 10 Then
                                                GetSections Suffix, 2
                                                A = FindPlayer(Section(1))
                                                If UCase(Section(1)) = UCase(Character.Name) Then A = Character.Index
                                                If A >= 1 Then
                                                    b = Int(Val(Section(2)))
                                                    If b >= 1 And b <= 10 Then
                                                        SendSocket Chr$(18) + Chr$(25) + Chr$(A) + Chr$(b)
                                                        PrintChat player(A).Name + " class has been changed.", 14, Options.FontSize, 15
                                                    Else
                                                        PrintChat "Invalid class number! (1-10)", 14, Options.FontSize, 15
                                                    End If
                                                Else
                                                    PrintChat "No such player!", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                            
                                         Case "SETSTATUS"
                                            If Character.Access >= 8 Then
                                                GetSections Suffix, 2
                                                A = FindPlayer(Section(1))
                                                If A >= 1 Then
                                                    b = Int(Val(Section(2)))
                                                    If b >= 0 And b <= 100 Then
                                                        SendSocket Chr$(18) + Chr$(17) + Chr$(A) + Chr$(b)
                                                        PrintChat player(A).Name + "'s status has been changed.", 14, Options.FontSize, 15
                                                    Else
                                                        PrintChat "Invalid status number!", 14, Options.FontSize, 15
                                                    End If
                                                Else
                                                    PrintChat "No such player!", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                            
                                        Case "SETGUILDSPRITE"
                                            If Character.Access >= 10 Then
                                                GetSections Suffix, 2
                                                A = FindGuild(Section(1))
                                                If A >= 1 Then
                                                    b = Int(Val(Section(2)))
                                                    If b <= 255 Then
                                                        SendSocket Chr$(18) + Chr$(15) + Chr$(A) + Chr$(b)
                                                        PrintChat Guild(A).Name + "'s sprite has been changed.", 14, Options.FontSize, 15
                                                    Else
                                                        PrintChat "Invalid sprite number!", 14, Options.FontSize, 15
                                                    End If
                                                Else
                                                    PrintChat "No such guild!", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        
                                        Case "SETMYSPRITE"
                                            If Character.Access >= 8 Then
                                                If Val(Suffix) > 0 Then
                                                    A = Int(Val(Suffix))
                                                    If A > 255 Then A = 255
                                                    SendSocket Chr$(18) + Chr$(6) + Chr$(Character.Index) + Chr$(A)
                                                    PrintChat "Sprite changed.", 14, Options.FontSize
                                                Else
                                                    PrintChat "You must specify a sprite number!", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                            
                                        Case "SETMYSTATUS"
                                            If Character.Access >= 8 Then
                                                If Val(Suffix) > 0 Then
                                                    A = Int(Val(Suffix))
                                                    If A > 255 Then A = 255
                                                    SendSocket Chr$(18) + Chr$(17) + Chr$(Character.Index) + Chr$(A)
                                                    PrintChat "Status changed.", 14, Options.FontSize, 15
                                                Else
                                                    PrintChat "You must specify a status number!", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        Case "GLOBAL", "GLOBA", "GLOB", "GLO", "GL", "G"
                                            If Suffix <> "" Then
                                                SendSocket Chr$(18) + Chr$(0) + Suffix
                                            Else
                                                PrintChat "You must specify a message!", 14, Options.FontSize, 15
                                            End If
                                        Case "SAVEWARP"
                                            GetSections Suffix, 1
                                            If Section(1) <> "" Then
                                                WriteString "Warps", LCase$(Section(1)) + "-map", CStr(CMap)
                                                WriteString "Warps", LCase$(Section(1)) + "-x", CStr(cX)
                                                WriteString "Warps", LCase$(Section(1)) + "-y", CStr(cY)
                                                PrintChat "Warp location saved.", 14, Options.FontSize
                                            Else
                                                PrintChat "You must specify a name for this warp location.", 14, Options.FontSize, 15
                                            End If
                                        Case "SAVEDWARP", "SAVEDWAR", "SAVEDWA", "SAVEDW"
                                            GetSections Suffix, 1
                                            If Section(1) <> "" Then
                                                A = ReadInt("Warps", Section(1) + "-map")
                                                If A > 0 Then
                                                    b = ReadInt("Warps", Section(1) + "-x")
                                                    C = ReadInt("Warps", Section(1) + "-y")
                                                    SendSocket Chr$(18) + Chr$(1) + DoubleChar(CInt(A)) + Chr$(b) + Chr$(C)
                                                    PrintChat "You have been warped.", 14, Options.FontSize, 15
                                                Else
                                                    PrintChat "No such warp location.", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                PrintChat "You must specify a warp location.", 14, Options.FontSize, 15
                                            End If
                                        
                                        Case "WARP", "WAR", "WA", "W"
                                            If Character.Access >= 4 Then
                                                GetSections Suffix, 3
                                                If Section(1) <> "" Then
                                                    A = Int(Val(Section(1)))
                                                    If A >= 1 And A <= 5000 Then
                                                        If Section(3) <> "" Then
                                                            b = Int(Val(Section(2)))
                                                            C = Int(Val(Section(3)))
                                                            If b >= 0 And b <= 11 And C >= 0 And C <= 11 Then
                                                                SendSocket Chr$(18) + Chr$(1) + DoubleChar(CInt(A)) + Chr$(b) + Chr$(C)
                                                            Else
                                                                PrintChat "Warp: Invalid parameters", 14, Options.FontSize, 15
                                                            End If
                                                        Else
                                                            SendSocket Chr$(18) + Chr$(1) + DoubleChar(CInt(A)) + Chr$(5) + Chr$(5)
                                                        End If
                                                    Else
                                                        PrintChat "Warp:  No such map", 14, Options.FontSize, 15
                                                    End If
                                                Else
                                                    PrintChat "Warp: Not enough parameters", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                            
                                        Case "WARPME", "WARPM"
                                            If Character.Access >= 2 Then
                                                GetSections Suffix, 1
                                                If Section(1) <> "" Then
                                                    A = FindPlayer(Section(1))
                                                    If A > 0 Then
                                                        SendSocket Chr$(18) + Chr$(2) + Chr$(A)
                                                        PrintChat "You have been warped to " + player(A).Name + ".", 14, Options.FontSize
                                                    Else
                                                        PrintChat "No such player", 14, Options.FontSize, 15
                                                    End If
                                                Else
                                                    PrintChat "You must specify a player to warp to.", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                            
                                        Case "WARPTOME", "WARPTOM", "WARPTO", "WARPT"
                                            If Character.Access >= 3 Then
                                                GetSections Suffix, 1
                                                If Section(1) <> "" Then
                                                    A = FindPlayer(Section(1))
                                                    If A > 0 Then
                                                        SendSocket Chr$(18) + Chr$(3) + Chr$(A) + DoubleChar(CInt(CMap)) + Chr$(cX) + Chr$(cY)
                                                        PrintChat player(A).Name + " has been warped to you.", 14, Options.FontSize, 15
                                                    Else
                                                        PrintChat "No such player", 14, Options.FontSize, 15
                                                    End If
                                                Else
                                                    PrintChat "You must specify a player to warp to.", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                            
                                        Case "GETI", "GETIP"
                                            If Character.Access >= 3 Then
                                                GetSections Suffix, 1
                                                If Section(1) <> "" Then
                                                    A = FindPlayer(Section(1))
                                                    If A > 0 Then
                                                        SendSocket Chr$(18) + Chr$(21) + Chr$(A)
                                                    Else
                                                        PrintChat "No such player", 14, Options.FontSize, 15
                                                    End If
                                                Else
                                                    PrintChat "You must specify a player to get the IP from.", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        Case "EDITMAP"
                                            If Character.Access >= 5 Then
                                                If MapEdit = False Then
                                                    SendSocket Chr$(80) + Chr$(player(Character.Index).map)
                                                    OpenMapEdit
    
                                                    
                                                Else
                                                    PrintChat "The map editor is already open!", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        Case "BANS"
                                            If Character.Access >= 4 Then
                                                SendSocket Chr$(80) + Chr$(player(Character.Index).map)
                                                If frmList_Loaded = False Then Load frmList
                                                With frmList
                                                    .lstBans.Clear
                                                    .lstBans.Visible = True
                                                    .lstMonsters.Visible = False
                                                    .lstObjects.Visible = False
                                                    .lstNPCs.Visible = False
                                                    .lstHalls.Visible = False
                                                    .lstPrefix.Visible = False
                                                    .lstLights.Visible = False
                                                    .lstScripts.Visible = False
                                                End With
                                                SendSocket Chr$(18) + Chr$(12)
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                            
                                        Case "EDITHALL"
                                            If Character.Access >= 10 Then
                                                SendSocket Chr$(80) + Chr$(player(Character.Index).map)
                                                GetSections Suffix, 1
                                                If Section(1) <> "" And Val(Section(1)) >= 1 And Val(Section(1)) <= 255 Then
                                                    SendSocket Chr$(48) + Chr$(Val(Section(1)))
                                                Else
                                                    If frmList_Loaded = False Then Load frmList
                                                    With frmList
                                                        .lstHalls.Visible = True
                                                        .lstObjects.Visible = False
                                                        .lstMonsters.Visible = False
                                                        .lstNPCs.Visible = False
                                                        .lstBans.Visible = False
                                                        .lstPrefix.Visible = False
                                                        .lstLights.Visible = False
                                                        .lstScripts.Visible = False
                                                        .btnOk.Caption = "Edit"
                                                        .Show
                                                    End With
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        
                                        Case "EDITPREFIX"
                                            If Character.Access >= 5 Then
                                                SendSocket Chr$(80) + Chr$(player(Character.Index).map)
                                                GetSections Suffix, 1
                                                If Section(1) <> "" And Val(Section(1)) >= 1 And Val(Section(1)) <= 255 Then
                                                    
                                                Else
                                                    If frmList_Loaded = False Then Load frmList
                                                    With frmList
                                                        .lstHalls.Visible = False
                                                        .lstObjects.Visible = False
                                                        .lstMonsters.Visible = False
                                                        .lstNPCs.Visible = False
                                                        .lstBans.Visible = False
                                                        .lstPrefix.Visible = True
                                                        .lstLights.Visible = False
                                                        .lstScripts.Visible = False
                                                        .Show
                                                    End With
                                                End If
                                            End If
                                            
                                        Case "EDITSCRIPT"
                                            If Character.Access >= 10 Then
                                                GetSections Suffix, 1
                                                If Section(1) <> "" Then
                                                    If Len(Section(1)) <= 15 Then
                                                        SendSocket Chr$(59) + Section(1)
                                                    Else
                                                        PrintChat "Error: Script name too long!", 14, Options.FontSize, 15
                                                    End If
                                                Else
                                                    PrintChat "Must specify a script name!", 14, Options.FontSize, 15
                                                End If
                                            End If
                                        Case "EDITSCRIPTS"
                                            If Character.Access >= 10 Then
                                                SendSocket Chr$(74)
                                            End If
                                        Case "EDITOBJECT"
                                            If Character.Access >= 6 Then
                                                GetSections Suffix, 1
                                                If Section(1) <> "" And Val(Section(1)) >= 1 And Val(Section(1)) <= MAXITEMS Then
                                                    SendSocket Chr$(19) + DoubleChar(Val(Section(1)))
                                                Else
                                                    If frmList_Loaded = False Then Load frmList
                                                    With frmList
                                                        .lstObjects.Visible = True
                                                        .lstMonsters.Visible = False
                                                        .lstNPCs.Visible = False
                                                        .lstBans.Visible = False
                                                        .lstHalls.Visible = False
                                                        .lstPrefix.Visible = False
                                                        .lstLights.Visible = False
                                                        .lstScripts.Visible = False
                                                        .btnOk.Caption = "Edit"
                                                        .Show
                                                    End With
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        
                                        Case "EDITMONSTER"
                                            If Character.Access >= 6 Then
                                                SendSocket Chr$(80) + Chr$(player(Character.Index).map)
                                                GetSections Suffix, 1
                                                If Section(1) <> "" And Val(Section(1)) >= 1 And Val(Section(1)) <= 255 Then
                                                    SendSocket Chr$(20) + Chr$(Val(Section(1)))
                                                Else
                                                    If frmList_Loaded = False Then Load frmList
                                                    With frmList
                                                        .lstMonsters.Visible = True
                                                        .lstObjects.Visible = False
                                                        .lstNPCs.Visible = False
                                                        .lstBans.Visible = False
                                                        .lstHalls.Visible = False
                                                        .lstPrefix.Visible = False
                                                        .lstLights.Visible = False
                                                        .lstScripts.Visible = False
                                                        .btnOk.Caption = "Edit"
                                                        .Show
                                                    End With
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        Case "EDITNPC"
                                            If Character.Access >= 6 Then
                                            SendSocket Chr$(80) + Chr$(player(Character.Index).map)
                                                GetSections Suffix, 1
                                                If Section(1) <> "" And Val(Section(1)) >= 1 And Val(Section(1)) <= 255 Then
                                                    SendSocket Chr$(50) + Chr$(Val(Section(1)))
                                                Else
                                                    If frmList_Loaded = False Then Load frmList
                                                    With frmList
                                                        .lstNPCs.Visible = True
                                                        .lstObjects.Visible = False
                                                        .lstMonsters.Visible = False
                                                        .lstBans.Visible = False
                                                        .lstHalls.Visible = False
                                                        .lstPrefix.Visible = False
                                                        .lstLights.Visible = False
                                                        .lstScripts.Visible = False
                                                        .btnOk.Caption = "Edit"
                                                        .Show
                                                    End With
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        Case "EDITLIGHT"
                                            If Character.Access >= 5 Then
                                            SendSocket Chr$(80) + Chr$(player(Character.Index).map)
                                                'GetSections Suffix, 1
                                                'If Section(1) <> "" And Val(Section(1)) >= 1 And Val(Section(1)) <= 50 Then
                                                    If frmList_Loaded = False Then Load frmList
                                                    With frmList
                                                        .lstMonsters.Visible = False
                                                        .lstObjects.Visible = False
                                                        .lstNPCs.Visible = False
                                                        .lstBans.Visible = False
                                                        .lstHalls.Visible = False
                                                        .lstPrefix.Visible = False
                                                        .lstLights.Visible = True
                                                        .lstScripts.Visible = False
                                                        .btnOk.Caption = "Edit"
                                                        .Show
                                                    End With
                                                'End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        Case "SCAN"
                                            If Character.Access >= 8 Then
                                                GetSections Suffix, 1
                                                If Section(1) <> "" Then
                                                    A = FindPlayer(Section(1))
                                                    If A >= 1 Then
                                                        SendSocket Chr$(18) + Chr$(22) + Chr$(A)
                                                    Else
                                                        PrintChat "Scan: No such player", 14, Options.FontSize, 15
                                                    End If
                                                Else
                                                    PrintChat "Scan: Not enough parameters", 14, Options.FontSize, 15
                                                End If
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        
                                        Case "FREEMAP"
                                            If Character.Access > 0 Then
                                                SendSocket Chr$(18) + Chr$(23)
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        Case "SPAWN"
                                            If Character.Access > 9 Then
                                                'frmMain.lblSpawnObj(2) = Object(sclSpawnObj).Name
                                                'frmMain.picSpawnObj.Visible = True
                                            Else
                                                'printchat "You do not have access to that command!", 14, Options.FontSize, 15
                                            End If
                                        Case Else
                                            'PrintChat "No such god command!", 14, Options.FontSize, 15
                                    End Select
                                Else
                                    'printchat "You do not have god access!", 14, Options.FontSize, 15
                                End If
                            
                            Case "ADMIN"
                                GetSections Suffix, 1
                                Select Case UCase$(Section(1))
                                    Case "ACCESS"
                                        'If Character.Access >= 10 Then
                                            GetSections Suffix, 3
                                            If Section(2) <> "" Then
                                                If UCase$(Section(1)) = UCase$(Character.Name) Then
                                                    A = Character.Index
                                                Else
                                                    A = FindPlayer(Section(1))
                                                End If
                                                    If A > 0 Then
                                                        b = Val(Section(2))
                                                        If b >= 0 And b <= 11 Then
                                                            SendSocket Chr$(64) + Chr$(1) + Chr$(A) + Chr$(b) + Trim$(Section(3))
                                                        Else
                                                            PrintChat "Invalid level.", 14, Options.FontSize, 15
                                                        End If
                                                    Else
                                                        PrintChat "No such player.", 14, Options.FontSize, 15
                                                    End If
                                            Else
                                                PrintChat "You must specify a player and an access level.", 14, Options.FontSize, 15
                                            End If
                                        'Else
                                        '    PrintChat "Invalid Admin Access", 14, Options.FontSize
                                        'End If
                                End Select
                            Case "PING"
                                If GetTickCount > PingTime + 2000 Then
                                    PingTime = GetTickCount
                                    SendSocket (Chr$(81))
                                Else
                                    PrintChat "You must wait to ping again.", 7, Options.FontSize, 15
                                End If
                            Case "PARTY", "PART", "PAR", "PA", "P" 'Party RA RA
                                GetSections Suffix, 1
                                Select Case UCase$(Section(1))
                                    Case "HELP"
                                        PrintChat "Party Commands: CREATE JOIN LEAVE INVITE WHO", 14, Options.FontSize, 6
                                    Case "CREATE"
                                        If Character.Party = 0 Then
                                            SendSocket Chr$(69) + Chr$(1) + Chr$(1)
                                        Else
                                            PrintChat "You are already in a party!", 14, Options.FontSize, 6
                                        End If
                                    Case "LEAVE"
                                        If Character.Party <> 0 Then
                                            SendSocket Chr$(69) + Chr$(2)
                                        Else
                                            PrintChat "You are not in a party!", 14, Options.FontSize, 6
                                        End If
                                    Case "JOIN"
                                        If Character.Party = 0 Then
                                            SendSocket Chr$(69) + Chr$(4)
                                        Else
                                            PrintChat "You are already in a party!", 14, Options.FontSize, 6
                                        End If
                                    Case "INVITE"
                                        GetSections Suffix, 1
                                        If Section(1) <> "" Then
                                            A = FindPlayer(Section(1))
                                            If A >= 1 Then
                                                If player(A).Party = 0 Then
                                                    If Character.Party = 0 Then
                                                        SendSocket Chr$(69) + Chr$(1) + Chr$(0)
                                                    End If
                                                    SendSocket Chr$(69) + Chr$(3) + Chr$(A)
                                                    PrintChat "You have invited " & player(A).Name & " to join your party!", 14, Options.FontSize, 6
                                                Else
                                                    If player(A).Party <> Character.Party Then
                                                        PrintChat player(A).Name + " is already in a party!", 14, Options.FontSize, 6
                                                    Else
                                                        PrintChat player(A).Name + " is already a member of your party!", 14, Options.FontSize, 6
                                                    End If
                                                End If
                                            Else
                                                PrintChat "No such player!", 14, Options.FontSize, 6
                                            End If
                                        Else
                                            PrintChat "You must enter a players name!", 14, Options.FontSize, 6
                                        End If
                                    Case "CHAT"
                                        If Suffix <> "" Then
                                            If Character.Party > 0 Then
                                                SendSocket Chr$(69) + Chr$(6) + Suffix
                                                PrintChat "[Party Chat] -> " + Suffix, CHannelColors(3), Options.FontSize, 6, True
                                            Else
                                                PrintChat "You are not in a party!", 14, Options.FontSize, 6
                                            End If
                                        End If
                                    Case "WHO"
                                        If Character.Party > 0 Then
                                            St1 = ""
                                            b = 0
                                            For A = 1 To MAXUSERS
                                                With player(A)
                                                    If .Sprite > 0 And A <> Character.Index And .Party = Character.Party Then
                                                        b = b + 1
                                                        St1 = St1 + ", " + .Name
                                                    End If
                                                End With
                                            Next A
                                            If b > 0 Then
                                                St1 = Mid$(St1, 2)
                                                PrintChat "There are " + CStr(b) + " other party members online: " + St1, 14, Options.FontSize, 6
                                            Else
                                                PrintChat "There are no other party members online.", 14, Options.FontSize, 6
                                            End If
                                        End If
                                End Select
                            Case "PLAYERTRADE"
                                GetSections Suffix, 1
                                Select Case UCase$(Section(1))
                                    Case "CANCEL"
                                        If TradeData.Tradestate(0) = 1 Then
                                            SendSocket Chr$(73) + Chr$(1)
                                        Else
                                            PrintChat "You have not requested a trade!", 14, Options.FontSize
                                        End If
                                    Case "ACCEPT"
                                        If TradeData.Tradestate(0) = TRADE_STATE_INVITED Then
                                            SendSocket Chr$(73) + Chr$(2)
                                        Else
                                            PrintChat "You have not been invited to trade!", 14, Options.FontSize
                                        End If
                                    Case Else
                                        If Section(1) <> "" Then
                                            A = FindPlayer(UCase(Section(1)))
                                            If A >= 1 Then
                                                SendSocket Chr$(73) + Chr$(1) + Chr$(A)
                                            Else
                                                PrintChat "No such player!", 14, Options.FontSize
                                            End If
                                        Else
                                            PrintChat "You must enter a players name!", 14, Options.FontSize
                                        End If
                                End Select
                            Case Else
                                St1 = UCase$(Section(1))
                                GetSections Suffix, 3
                                SendSocket Chr$(62) + St1 + Chr$(0) + Section(1) + Chr$(0) + Section(2) + Chr$(0) + Section(3)
                        End Select
                    End If
                Else
                    If Character.Squelched = False Then
                    SendSocket Chr$(6) + ChatString
                    PrintChat "You say, " + Chr$(34) + ChatString + Chr$(34), 7, Options.FontSize
                    Else
                    PrintChat "You are squelched!", 14, Options.FontSize
                    End If
                End If
                ChatString = ""
                DrawChatString
            'Else

            End If
        End If

    End If
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If txtDrop.Visible Then Exit Sub
    Select Case KeyCode
        Case 18 'Alt
            keyAlt = False
        Case KeyCodeList(Options.CycleKey).KeyCode
            tabDown = False
    End Select
End Sub

Private Sub Form_Load()
    frmMain_Loaded = True
    Set Me.Icon = frmMenu.Icon
    Set Me.Picture = LoadPicture("Data\Graphics\Interface\Interface.rsc")
    
    Me.Top = Screen.Height / 2 - 300
    Me.Left = Screen.Width / 2 - 400
    
    On Error Resume Next
    MainWndProc = SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim StillMove As Boolean
    Dim A As Long, b As Long
    StillMove = True
    
    If x > INVDestX And x < INVDestX + INVWIDTH Then
        If y > INVDestY And y < INVDestY + INVHEIGHT Then
            StartInvX = x
            StartInvY = y
            x = x - INVDestX
            y = y - INVDestY
            CurInvObj = Int(x / (34 * WindowScaleX)) + (Int(y / (35 * WindowScaleY)) * 5) + 1
            If CurInvObj < 1 Then CurInvObj = 1
            If CurInvObj > 25 Then CurInvObj = 25
            
            If Button = 2 And StorageOpen = False Then
                If CurInvObj <= 20 Then
                    If Character.Inv(CurInvObj).Object > 0 Then
                        If Character.Trading = False Then
                            If Object(Character.Inv(CurInvObj).Object).Type = 6 Or Object(Character.Inv(CurInvObj).Object).Type = 11 Then
                                'Money
                                TempVar2 = Character.Inv(CurInvObj).Value
                                TempVar1 = TempVar2
                                TempVar3 = CurInvObj
                                txtDrop = TempVar1
                                picDrop.Visible = True
                                lblDropTitle = "Drop How Much?"
                                txtDrop.SetFocus
                                txtDrop.SelStart = 0
                                txtDrop.SelLength = Len(txtDrop)
                            Else
                                SendSocket Chr$(9) + Chr$(CurInvObj) + QuadChar(0)
                            End If
                        Else
                            PrintChat "You cannot do that while trading!", 14, Options.FontSize
                        End If
                    End If
                ElseIf CurInvObj > 20 And CurInvObj <= 25 Then
                    If Character.Trading = False Then
                        If Character.Equipped(CurInvObj - 20).Object > 0 Then
                            If Object(Character.Equipped(CurInvObj - 20).Object).Type = 6 Or Object(Character.Equipped(CurInvObj - 20).Object).Type = 11 Then
                                'Money
                                TempVar1 = 0
                                TempVar2 = Character.Equipped(CurInvObj - 20).Value
                                TempVar1 = TempVar2
                                TempVar3 = CurInvObj
                                picDrop.Visible = True
                                lblDropTitle = "Drop How Much?"
                                txtDrop = TempVar1
                                txtDrop.SetFocus
                                txtDrop.SelStart = 0
                                txtDrop.SelLength = Len(txtDrop)
                            Else
                                SendSocket Chr$(9) + Chr$(CurInvObj) + QuadChar(0)
                            End If
                        End If
                    Else
                        PrintChat "You cannot do that while trading!", 14, Options.FontSize
                    End If
                End If
            ElseIf Button = 1 Then
                DraggingObj = True
            End If

            SetTab tsInventory
            
            StillMove = False
            Exit Sub
        End If
    End If
    
    If x > 8 * WindowScaleX And x < 198 * WindowScaleX Then
        If y > 45 * WindowScaleY And y < 412 * WindowScaleY Then
            Select Case GetTab
                Case tsStats
                    If x > 149 * WindowScaleX And x < 160 * WindowScaleX Then
                        If y > 168 * WindowScaleY And y < 179 * WindowScaleY Then
                            If Character.strength + Tstr < 50 And TempVar5 > 0 Then
                                Tstr = Tstr + 1
                                TempVar5 = TempVar5 - 1
                                DrawTrainBars
                            End If
                        ElseIf y > 200 * WindowScaleY And y < 211 * WindowScaleY Then
                            If Character.Agility + TAgi < 50 And TempVar5 > 0 Then
                                TAgi = TAgi + 1
                                TempVar5 = TempVar5 - 1
                                DrawTrainBars
                            End If
                        ElseIf y > 231 * WindowScaleY And y < 242 * WindowScaleY Then
                            If Character.Endurance + TEnd < 50 And TempVar5 > 0 Then
                                TEnd = TEnd + 1
                                TempVar5 = TempVar5 - 1
                                DrawTrainBars
                            End If
                        ElseIf y > 263 * WindowScaleY And y < 274 * WindowScaleY Then
                            If Character.Wisdom + TWis < 50 And TempVar5 > 0 Then
                                TWis = TWis + 1
                                TempVar5 = TempVar5 - 1
                                DrawTrainBars
                            End If
                        ElseIf y > 296 * WindowScaleY And y < 307 * WindowScaleY Then
                            If Character.Constitution + TCon < 50 And TempVar5 > 0 Then
                                TCon = TCon + 1
                                TempVar5 = TempVar5 - 1
                                DrawTrainBars
                            End If
                        ElseIf y > 328 * WindowScaleY And y < 339 * WindowScaleY Then
                            If Character.Intelligence + TInt < 50 And TempVar5 > 0 Then
                                TInt = TInt + 1
                                TempVar5 = TempVar5 - 1
                                DrawTrainBars
                            End If
                        End If
                    End If
                    If x > 164 * WindowScaleX And x < 175 * WindowScaleX Then
                        If y > 168 * WindowScaleY And y < 179 * WindowScaleY Then
                            If Tstr > 0 Then
                                Tstr = Tstr - 1
                                TempVar5 = TempVar5 + 1
                                DrawTrainBars
                            End If
                        ElseIf y > 200 * WindowScaleY And y < 211 * WindowScaleY Then
                            If TAgi > 0 Then
                                TAgi = TAgi - 1
                                TempVar5 = TempVar5 + 1
                                DrawTrainBars
                            End If
                        ElseIf y > 231 * WindowScaleY And y < 242 * WindowScaleY Then
                            If TEnd > 0 Then
                                TEnd = TEnd - 1
                                TempVar5 = TempVar5 + 1
                                DrawTrainBars
                            End If
                        ElseIf y > 263 * WindowScaleY And y < 274 * WindowScaleY Then
                            If TWis > 0 Then
                                TWis = TWis - 1
                                TempVar5 = TempVar5 + 1
                                DrawTrainBars
                            End If
                        ElseIf y > 296 * WindowScaleY And y < 307 * WindowScaleY Then
                            If TCon > 0 Then
                                TCon = TCon - 1
                                TempVar5 = TempVar5 + 1
                                DrawTrainBars
                            End If
                        ElseIf y > 328 * WindowScaleY And y < 339 * WindowScaleY Then
                            If TInt > 0 Then
                                TInt = TInt - 1
                                TempVar5 = TempVar5 + 1
                                DrawTrainBars
                            End If
                        End If
                    End If
                    If x > 128 * WindowScaleX And x < 180 * WindowScaleX Then 'Save Button when Updating Stats
                        If y > 135 * WindowScaleY And y < 156 * WindowScaleY Then
                            If Tstr + TAgi + TEnd + TWis + TCon + TInt > 0 Then
                                With Character
                                    .strength = .strength + Tstr
                                    .Agility = .Agility + TAgi
                                    .Endurance = .Endurance + TEnd
                                    .Wisdom = .Wisdom + TWis
                                    .Constitution = .Constitution + TCon
                                    .Intelligence = .Intelligence + TInt
                                    .StatPoints = TempVar5
                                End With
                                DrawTrainBars
                                SendSocket Chr$(30) + Chr$(Tstr) + Chr$(TAgi) + Chr$(TEnd) + Chr$(TInt) + Chr$(TWis) + Chr$(TCon)
                                Tstr = 0: TAgi = 0: TEnd = 0: TWis = 0: TCon = 0: TInt = 0
                            End If
                        End If
                    End If
                    
                    If x > 28 * WindowScaleX And x < 144 * WindowScaleX Then          'Stat Descriptions
                        If y >= 152 * WindowScaleY And y <= 181 * WindowScaleY Then     'Strength
                            PrintChat "Strength increases your melee damage.", 14, Options.FontSize, 15
                        ElseIf y >= 184 * WindowScaleY And y <= 213 * WindowScaleY Then 'Agility
                            PrintChat "Agility increases your chance to dodge and your chance to score a critical hit.", 14, Options.FontSize, 15
                        ElseIf y >= 216 * WindowScaleY And y <= 244 * WindowScaleY Then 'Endurance
                            PrintChat "Endurance increases your Energy, Energy regeneration, and your chance to block.", 14, Options.FontSize, 15
                        ElseIf y >= 248 * WindowScaleY And y <= 276 * WindowScaleY Then 'Wisdom
                            PrintChat "Piety increases all stats and regeneration, as well as magic resistance.", 14, Options.FontSize, 15
                        ElseIf y >= 280 * WindowScaleY And y <= 309 * WindowScaleY Then 'Constitution
                            PrintChat "Constitution increases your Health and Health regeneration.", 14, Options.FontSize, 15
                        ElseIf y >= 312 * WindowScaleY And y <= 341 * WindowScaleY Then 'Intelligence
                            PrintChat "Intelligence increases your Mana and Mana regeneration.", 14, Options.FontSize, 15
                        End If
                    End If
                Case tsParty
                    If Character.Party > 0 Then
                        A = ((y - 55 * WindowScaleY) \ (43 * WindowScaleY)) + 1
                        If A > 0 And A <= 10 Then
                            If PartyIndex(A) > 0 Then
                                If player(PartyIndex(A)).Party = Character.Party Then
                                    CurrentTarget.TargetType = TT_PLAYER
                                    CurrentTarget.Target = PartyIndex(A)
                                End If
                            End If
                        End If
                    End If
                Case tsSkills
                    If y >= 66 * WindowScaleY And y < 82 * WindowScaleY Then
                        If SkillListBox.YOffset > 0 Then SkillListBox.YOffset = SkillListBox.YOffset - 1 'TODO RESIZE?
                    End If
                    If y > 258 * WindowScaleY And y < 274 * WindowScaleY Then
                        If SkillListBox.YOffset < 255 - 14 Then SkillListBox.YOffset = SkillListBox.YOffset + 1 'TODO RESIZE?
                    End If
                    DrawLstBox
            End Select
            StillMove = False
        End If
    End If
    
    If y > 23 * WindowScaleY And y < 44 * WindowScaleY Then 'Tabs
        If x > 8 * WindowScaleX And x < 68 * WindowScaleX Then
            SetTab tsStats
        ElseIf x > 73 * WindowScaleX And x < 132 * WindowScaleX Then
            SetTab tsSkills
        ElseIf x > 138 * WindowScaleX And x < 196 * WindowScaleX Then
            SetTab tsParty
        End If
    Else
        If y >= 385 * WindowScaleY And y <= 407 * WindowScaleY Then
            If x >= 138 * WindowScaleX And x <= 191 * WindowScaleX Then
                If CurrentTab = tsStats Then

                    calculateEvasion
                    calculateAttackDamage
                    calculateBlock
                    calculatePoisonResist
                    calculateHpRegen
                    calculateManaRegen
                    calculateCriticalChance
                    calculateMagicResist
                    calculateAttackSpeed
                    calculateLeachHp
                    calculateMagicFind
                    calculateEnergyRegen
                                        
                    SetTab tsStats2
                ElseIf CurrentTab = tsStats2 Then
                    SetTab tsStats
                ElseIf CurrentTab = tsSkills Then
                    If frmSkillTree.Visible = False Then
                        Load frmSkillTree
                        frmSkillTree.Show vbModeless, frmMain
                    End If
                End If
            End If
        End If
    End If
    
    If x > 774 * WindowScaleX And x < 794 * WindowScaleX Then 'Mini Map Tabs
        If y > 297 * WindowScaleY And y < 356 * WindowScaleY Then 'Map
            SetMiniMapTab tsMap
        ElseIf y > 360 * WindowScaleY And y < 419 * WindowScaleY Then
            SetMiniMapTab tsButtons
        End If
    End If
    
    If x >= 780 * WindowScaleX And y >= 581 * WindowScaleY Then
        NewChatWindow
    End If
    
    If MiniMapTab = tsButtons Then
        If x > 603 * WindowScaleX And x < 685 * WindowScaleX Then
            If y > 328 * WindowScaleY And y < 355 * WindowScaleY Then   'Helpme
                ChatString = "/helpme"
                Form_KeyPress 200
                StillMove = False
            End If
            If y > 358 * WindowScaleY And y < 385 * WindowScaleY Then 'Renown
                ChatString = "/Renown"
                Form_KeyPress 200
                StillMove = False
            End If
            If y > 390 * WindowScaleY And y < 417 * WindowScaleY Then 'Options
                frmOptions.Show
                frmOptions.ZOrder vbBringToFront
                StillMove = False
            End If
        End If
        If x > 688 * WindowScaleX And x < 770 * WindowScaleX Then
            If y > 328 * WindowScaleY And y < 355 * WindowScaleY Then 'Guilds
                frmGuilds.Show
                With frmGuilds
                .GuildRank = Character.GuildRank
                .playerGuild = Character.Guild
                .nextTab = 0
                .changed = True
                .ctab = 0
                .setGuildTab
                .lstEnabled = True
                End With
                StillMove = False
            End If
            If y > 358 * WindowScaleY And y < 385 * WindowScaleY Then
                Dim St1 As String
                St1 = ""
                b = 0
                For A = 1 To MAXUSERS
                    With player(A)
                        If .Sprite > 0 And A <> Character.Index Then
                            b = b + 1
                            St1 = St1 + ", " + .Name
                        End If
                    End With
                Next A
                If b > 0 Then
                    St1 = Mid$(St1, 2)
                    PrintChat "There are " + CStr(b) + " other players online: " + St1, 14, Options.FontSize
                Else
                    PrintChat "There are no other players online.", 14, Options.FontSize
                End If
                StillMove = False
            End If
            If y > 388 * WindowScaleY And y < 415 * WindowScaleY Then
                frmMacros.Show
                frmMacros.ZOrder vbBringToFront
                StillMove = False
            End If
        End If
    End If
    If x > 720 * WindowScaleX And x < 736 * WindowScaleX And y > 3 * WindowScaleY And y < 20 * WindowScaleY Then
        Me.WindowState = 1
    End If
    If x > 739 * WindowScaleX And x < 756 * WindowScaleX And y > 3 * WindowScaleY And y < 20 * WindowScaleY Then
        If Fullscreen Then
            Fullscreen = False
            ResizeGameWindow
        Else
            Fullscreen = True
            FullscreenGameWindow
        End If
    End If
    If x > 759 * WindowScaleX And x < 777 * WindowScaleX And y > 3 * WindowScaleY And y < 20 * WindowScaleY Then
        If Character.Access = 0 Then
            If Character.HP < Character.MaxHP Then
                PrintChat "You must have full HP to log out.", 15, Options.FontSize
                Exit Sub
            End If
        End If
        If frmSkillTree.Visible Then Unload frmSkillTree
        A = GetTickCount + 300
        Transition 0, 0, 0, 0, 10
        While GetTickCount < A
        Wend
        CloseClientSocket 0, True
        DisableScreen = False
    End If

    
    
    If StillMove Then
       Dim ReturnVal As Long
       ReleaseCapture
       ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim A As Long, b As Long, C As Long, D As Long, tDC As Long, R1 As RECT, r2 As RECT
    'Keep track of mouse x and y's because well .. there's just no
    'way to know what they are in the Double Click Procedure
    CurMouseX = x
    CurMouseY = y
    DoEvents
    A = x 'X / WindowScaleX
    b = y
    If x >= 524 * WindowScaleX And x <= 522 * WindowScaleX + 28 * WindowScaleX And y > 4 * WindowScaleY And y < 30 * WindowScaleY Then
        If Not hoverCombat Then
            hoverCombat = True
            drawTopBar
        End If
    Else
        If hoverCombat Then
            hoverCombat = False
            drawTopBar
        End If
    End If
    If Button = 1 Then
        If CurInvObj > 0 And CurInvObj <= 20 And DraggingObj Then
            If Character.Inv(CurInvObj).Object > 0 Then
                If A < INVDestX + 16 Then A = INVDestX + 16
                If b < INVDestY + 16 Then b = INVDestY + 16
                If A > INVDestX + INVWIDTH - 18 Then A = INVDestX + INVWIDTH - 18
                If b > INVDestY + INVHEIGHT - 50 Then b = INVDestY + INVWIDTH - 50
                If Sqr((A - StartInvX) ^ 2 + (b - StartInvY) ^ 2) >= 16 Or ReplaceCursor Then
                    If ReplaceCursor Then
                        R1.Top = 0: R1.Left = 32: R1.Right = 64: R1.Bottom = 32
                        r2.Top = OldInvY - 16: r2.Left = OldInvX - 16: r2.Right = r2.Left + 32: r2.Bottom = r2.Top + 32
                        sfcCursor.Surface.BltToDC frmMain.hdc, R1, r2
                    End If
                    tDC = sfcCursor.Surface.GetDC
                    BitBlt tDC, 0, 0, 32, 32, frmMain.hdc, A - 16, b - 16, SRCCOPY
                    BitBlt tDC, 32, 0, 32, 32, frmMain.hdc, A - 16, b - 16, SRCCOPY
                    sfcCursor.Surface.ReleaseDC tDC
                    If ExamineBit(Object(Character.Inv(CurInvObj).Object).Flags, 6) Then
                        C = (Object(Character.Inv(CurInvObj).Object).Picture + 255 - 1)
                    Else
                        C = (Object(Character.Inv(CurInvObj).Object).Picture - 1)
                    End If
                    D = C Mod 64
                    C = (C \ 64) + 1
                    If C > 0 And C <= NumObjects Then
                        Draw sfcCursor, 0, 0, 32, 32, sfcObjects(C), (D Mod 8) * 32, (D \ 8) * 32, True
                    End If
                    R1.Top = 0: R1.Left = 0: R1.Right = 32: R1.Bottom = 32
                    r2.Top = b - 16: r2.Left = A - 16: r2.Right = r2.Left + 32: r2.Bottom = r2.Top + 32
                    sfcCursor.Surface.BltToDC frmMain.hdc, R1, r2
                    ReplaceCursor = True
                    OldInvX = A
                    OldInvY = b
                    frmMain.Refresh
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim DropObj As Long, R1 As RECT, r2 As RECT
    If Button = 1 Then
        If ReplaceCursor = True And CurInvObj <> DropObj And DraggingObj Then
            DraggingObj = False
            x = x - INVDestX
            y = y - INVDestY
            x = x / WindowScaleX
            y = y / WindowScaleY
            If x < 0 Then x = 0
            If y < 0 Then y = 0
            If x > INVWIDTH - 6 Then x = INVWIDTH - 6
            If y > INVHEIGHT - 38 Then y = INVHEIGHT - 38
            DropObj = Int(x / 34) + (Int(y / 35) * 5) + 1
            If DropObj < 1 Then DropObj = 1
            If DropObj > 20 Then DropObj = 20
            R1.Top = 0: R1.Left = 32: R1.Right = 64: R1.Bottom = 32
            r2.Top = OldInvY - 16: r2.Left = OldInvX - 16: r2.Right = r2.Left + 32: r2.Bottom = r2.Top + 32
            sfcCursor.Surface.BltToDC frmMain.hdc, R1, r2
            frmMain.Refresh
            
            'Now figure out what to do
            If CurInvObj > 0 And CurInvObj <= 20 Then
                If DropObj > 0 And DropObj <= 20 Then
                    If Character.Inv(CurInvObj).Object > 0 Then
                        SendSocket Chr$(66) + Chr$(CurInvObj) + Chr$(DropObj)
                    End If
                ElseIf DropObj > 20 And DropObj <= 25 Then
'                    B = Character.Inv(CurInvObj).Object
'                    If B > 0 Then
'                        A = 0
'                        Select Case DropObj - 20
'                            Case 1 'Weapon
'                                If Object(B).Type = 1 Or Object(B).Type = 10 Then A = 1
'                            Case 2 'Shield
'                                If Object(B).Type = 2 Or Object(B).Type = 11 Then A = 1
'                            Case 3 'Armor
'                                If Object(B).Type = 3 Then A = 1
'                            Case 4  'Helmet
'                                If Object(B).Type = 4 Then A = 1
'                            Case 5 'Ring
'                                If Object(B).Type = 8 Then A = 1
'                        End Select
'                        If A Then SendSocket Chr$(10) + Chr$(CurInvObj) 'Use Obj
'                    End If
                End If
            ElseIf CurInvObj > 20 And CurInvObj <= 25 Then
                If Character.Equipped(CurInvObj - 20).Object > 0 Then
                    If DropObj > 0 And DropObj <= 20 Then
                        If Object(Character.Inv(DropObj).Object).Type = Object(Character.Equipped(CurInvObj - 20).Object).Type Then
                            SendSocket Chr$(10) + Chr$(DropObj)
                        Else
                            SendSocket Chr$(66) + Chr$(CurInvObj) + Chr$(DropObj)
                        End If
                    End If
                End If
            End If
            
            CurInvObj = DropObj
        Else
            DraggingObj = False
        End If
    End If
    ReplaceCursor = False
    'ReleaseCapture
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain_Loaded = False
    Set Me.Picture = Nothing
    On Error Resume Next
    Call SetWindowLong(Me.hwnd, GWL_WNDPROC, MainWndProc)
End Sub

Private Sub lblMenu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
  lblMenu(Index).BackColor = QBColor(15)
End Sub


Private Sub lblMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim A As Long, b As Long
    lblMenu(Index).BackColor = QBColor(8)
    If x >= 0 And x <= lblMenu(Index).Width * Screen.twipsPerpixelX And y >= 0 And y <= lblMenu(Index).Height * Screen.twipsPerpixelY Then
        Select Case Index
            Case 34 'Drop/Cancel
                picDrop.Visible = False
                If CurrentWindow = WINDOW_TRADE Or CurrentWindow = WINDOW_STORAGE Then
                    CurrentWindowFlags = (CurrentWindowFlags And (Not WINDOW_FLAG_INVISIBLE))
                End If
            Case 35 'Drop/Ok
                If TempVar1 > 0 And TempVar1 <= TempVar2 Then
                    If CurrentWindow = WINDOW_TRADE Then
                        CurrentWindowFlags = (CurrentWindowFlags And (Not WINDOW_FLAG_INVISIBLE))
                        picDrop.Visible = False
                        If Character.Trading Then
                            If TempVar3 > 0 And TempVar3 <= 20 Then
                                If TempVar1 > 0 And TempVar1 <= Character.Inv(TempVar3).Value Then
                                    b = 0
                                    For A = 1 To 10
                                        If TradeData.Slot(A) = TempVar3 Then
                                            b = A
                                            Exit For
                                        End If
                                    Next A
                                    If b = 0 Then
                                        For A = 1 To 10
                                            If TradeData.YourObjects(A).Object = 0 And TradeData.Slot(A) = 0 Then
                                                b = A
                                                Exit For
                                            End If
                                        Next A
                                    End If
                                    If b > 0 And b <= 10 Then
                                        SendSocket Chr$(73) + Chr$(4) + Chr$(b) + Chr$(TempVar3) + QuadChar$(TempVar1)
                                        With Character.Inv(TempVar3)
                                            TradeData.Slot(b) = TempVar3
                                            TradeData.YourObjects(b).Object = .Object
                                            TradeData.YourObjects(b).Value = TempVar1
                                            TradeData.YourObjects(b).Prefix = .Prefix
                                            TradeData.YourObjects(b).PrefixValue = .PrefixValue
                                            TradeData.YourObjects(b).Suffix = .Suffix
                                            TradeData.YourObjects(b).SuffixValue = .SuffixValue
                                        End With
                                    Else
                                        PrintChat "You can not currently trade any more items.  Try dropping some.", 7, Options.FontSize
                                    End If
                                End If
                            End If
                        End If
                    ElseIf StorageOpen = False Then
                        SendSocket Chr$(9) + Chr$(TempVar3) + QuadChar(TempVar1)
                    Else
                        If CurStorageObj > 0 Then
                            SendSocket Chr$(75) + Chr$(2) + Chr$(LastNPCX) + Chr$(LastNPCY) + Chr$(TempVar3) + QuadChar(TempVar1)
                        ElseIf CurInvObj > 0 Then
                            SendSocket Chr$(75) + Chr$(1) + Chr$(LastNPCX) + Chr$(LastNPCY) + Chr$(TempVar3) + QuadChar(TempVar1)
                        End If
                        CurrentWindowFlags = CurrentWindowFlags And (Not WINDOW_FLAG_INVISIBLE)
                    End If
                    picDrop.Visible = False
                End If
        End Select
    End If
End Sub

Private Sub lblScannedItem_Click(Index As Integer)
Dim A As Long
A = Index + 1
If PlayerScanItem(Index + 1).objNum > 0 Then
    lblScannedItemData(0).Caption = ""
    If PlayerScanItem(Index + 1).Prefix > 0 Then
        lblScannedItemData(0).Caption = Prefix(PlayerScanItem(Index + 1).Prefix).Name + " "
    End If
    lblScannedItemData(0).Caption = lblScannedItemData(0).Caption + Object(PlayerScanItem(Index + 1).objNum).Name
    If PlayerScanItem(Index + 1).Suffix > 0 Then
        lblScannedItemData(0).Caption = lblScannedItemData(0).Caption + " " + Prefix(PlayerScanItem(Index + 1).Suffix).Name
    End If
    lblScannedItemData(1).Caption = ModString(PlayerScanItem(Index + 1).PrefixType)
    lblScannedItemData(2).Caption = PlayerScanItem(Index + 1).PrefixVal
    lblScannedItemData(3).Caption = ModString(PlayerScanItem(Index + 1).SuffixType)
    lblScannedItemData(4).Caption = PlayerScanItem(Index + 1).SuffixVal
    lblScannedItemData(5).Caption = PlayerScanItem(Index + 1).objNum
    lblScannedItemData(6).Caption = PlayerScanItem(Index + 1).ObjValue
Else
    lblScannedItemData(0) = "-Nothing-"
    lblScannedItemData(1) = "-Nothing-"
    lblScannedItemData(2) = "0"
    lblScannedItemData(3) = "-Nothing-"
    lblScannedItemData(4) = "0"
    lblScannedItemData(5) = "0"
    lblScannedItemData(6) = "0"
End If
frmMain.picPlayerScan.Height = 345
frmMain.picPlayerScan.Width = 410
End Sub

Private Sub lstSkills_Click()
    CastingSpell = False
End Sub

Private Sub lstSkills_DblClick()
    If lstSkills.Tag < lstSkills.Width - 15 Then
        With SkillListBox
            If .Data(.Selected) > 0 Then
                Character.MacroSkill = 0
                If Skills(.Data(.Selected)).TargetType = TT_NO_TARGET Then
                    CurrentTarget.TargetType = TT_NO_TARGET
                    UseSkill .Data(.Selected)
                ElseIf Skills(.Data(.Selected)).TargetType = TT_TILE Then
                    If GetTickCount >= Character.LocalSpellTick(.Data(.Selected)) And GetTickCount >= Character.GlobalSpellTick Then
                        CastingSpell = True
                        Character.MacroSkill = .Data(.Selected)
                    End If
                Else
                    If CurrentTarget.TargetType = TT_PLAYER Or CurrentTarget.TargetType = TT_CHARACTER Or CurrentTarget.TargetType = TT_MONSTER Then
                        UseSkill .Data(.Selected)
                    Else
                        If CurrentTarget.TargetType = TT_NO_TARGET Or CurrentTarget.TargetType = 0 Then
                            If Skills(.Data(.Selected)).TargetType And TT_CHARACTER Then
                                CurrentTarget.TargetType = TT_CHARACTER
                                UseSkill .Data(.Selected)
                            End If
                        End If
                    End If
                End If
            End If
        End With
    End If
End Sub

Private Sub picMiniMap_Paint()
    doMiniMapDraw = True
End Sub

Private Sub picPlayerScan_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
       ReleaseCapture
       Call SendMessage(picPlayerScan.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub

Private Sub picViewport_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim A As Long, b As Long, MapX As Long, MapY As Long, SpriteX As Long, SpriteY As Long, St As String
    MapX = x \ TileSizeX
    MapY = y \ TileSizeY
    
    
    
    'ParticleEngineF.Add cX * 32 + 16, cY * 32, 12, 25, 25, 200, 30, 0.1, 0, 200, 2, Int(10 * Rnd) + 1, 8, TT_NO_TARGET, 0

    SpriteX = x \ TileSizeX
    If y > TileSizeY \ 2 Then
        SpriteY = (y + TileSizeY \ 2) \ TileSizeY
    End If
    If GUIWindow_MouseDown(x, y, Button) = False Then
        If MapEdit = True Then
            If MapX >= 0 And MapX <= 11 And MapY >= 0 And MapY <= 11 Then
                If Button = 1 Then
                    If frmMapEdit.lblMapOptions(1).BackColor = QBColor(15) Then
                        Select Case EditMode
                            Case 0 'Ground
                                CurTile = EditMap.Tile(MapX, MapY).Ground
                            Case 1 'BGTile1
                                CurTile = EditMap.Tile(MapX, MapY).BGTile1
                            Case 2 'Anim
                                CurAnim(1) = EditMap.Tile(MapX, MapY).Anim(1)
                                CurAnim(2) = EditMap.Tile(MapX, MapY).Anim(2)
                            Case 3 'FGTile
                                CurTile = EditMap.Tile(MapX, MapY).FGTile
                            Case 4 'Attribute
                                CurAtt = EditMap.Tile(MapX, MapY).Att
                                CurAttData(0) = EditMap.Tile(MapX, MapY).AttData(0)
                                CurAttData(1) = EditMap.Tile(MapX, MapY).AttData(1)
                                CurAttData(2) = EditMap.Tile(MapX, MapY).AttData(2)
                                CurAttData(3) = EditMap.Tile(MapX, MapY).AttData(3)
                                NewAtt = CurAtt + 50
                                frmMapAtt.Show 0 ' 1 'vbModal
                            Case 5 'Ground2
                                CurTile = EditMap.Tile(MapX, MapY).Ground2
                            Case 7 'Wall
                                CurWall = EditMap.Tile(MapX, MapY).WallTile
                        End Select
                        frmMapEdit.lblMapOptions(1).BackColor = QBColor(8)
                        frmMapEdit.RedrawTile
                    Else
                        Select Case EditMode
                            Case 0 'Ground
                                If EditMap.Tile(MapX, MapY).Ground <> CurTile Then
                                    SetMapUndo MapX, MapY, 1, CLng(EditMap.Tile(MapX, MapY).Ground)
                                    EditMap.Tile(MapX, MapY).Ground = CurTile
                                End If
                            Case 1 'BGTile1
                                If EditMap.Tile(MapX, MapY).BGTile1 <> CurTile Then
                                    SetMapUndo MapX, MapY, 3, CLng(EditMap.Tile(MapX, MapY).BGTile1)
                                    EditMap.Tile(MapX, MapY).BGTile1 = CurTile
                                End If
                            Case 2 'Anim
                                If keyAlt = True Then
                                    If CurSubX > 0 And CurSubX < TileSizeX \ 2 Then
                                        If CurSubY > TileSizeY \ 2 And CurSubY < TileSizeY Then 'Add to frame delay
                                            If ((EditMap.Tile(MapX, MapY).Anim(2) \ 4) And 15) < 15 Then
                                                EditMap.Tile(MapX, MapY).Anim(2) = ((((EditMap.Tile(MapX, MapY).Anim(2) \ 4) And 15) + 1) * 4) Or (EditMap.Tile(MapX, MapY).Anim(2) And 3)
                                            End If
                                        End If
                                    End If
                                    If CurSubX > TileSizeX \ 2 And CurSubX < TileSizeX Then
                                        If CurSubY > TileSizeY \ 2 And CurSubY < TileSizeY Then 'subtract frame delay
                                            If ((EditMap.Tile(MapX, MapY).Anim(2) \ 4) And 15) > 0 Then
                                                EditMap.Tile(MapX, MapY).Anim(2) = ((((EditMap.Tile(MapX, MapY).Anim(2) \ 4) And 15) - 1) * 4) Or (EditMap.Tile(MapX, MapY).Anim(2) And 3)
                                            End If
                                        End If
                                    End If
                                    If CurSubX > TileSizeX \ 2 And CurSubX < TileSizeX * 2 / 3 Then
                                        If CurSubY > 0 And CurSubY < TileSizeY \ 2 Then 'Toggle Flag
                                            ToggleBit EditMap.Tile(MapX, MapY).Anim(2), 7
                                        End If
                                    End If
                                    If CurSubX >= TileSizeX * 2 / 3 And CurSubX < TileSizeX Then
                                        If CurSubY > 0 And CurSubY < TileSizeY \ 2 Then 'Toggle Flag
                                            ToggleBit EditMap.Tile(MapX, MapY).Anim(2), 6
                                        End If
                                    End If
                                Else
                                    EditMap.Tile(MapX, MapY).Anim(1) = CurAnim(1)
                                    EditMap.Tile(MapX, MapY).Anim(2) = CurAnim(2)
                                End If
                            Case 3 'FGTile
                                If EditMap.Tile(MapX, MapY).FGTile <> CurTile Then
                                    SetMapUndo MapX, MapY, 5, CLng(EditMap.Tile(MapX, MapY).FGTile)
                                    EditMap.Tile(MapX, MapY).FGTile = CurTile
                                End If
                            Case 4 'Attribute
                                If EditMap.Tile(MapX, MapY).Att <> CurAtt Or EditMap.Tile(MapX, MapY).AttData(0) <> CurAttData(0) Or EditMap.Tile(MapX, MapY).AttData(1) <> CurAttData(1) Or EditMap.Tile(MapX, MapY).AttData(2) = CurAttData(2) Or EditMap.Tile(MapX, MapY).AttData(3) <> CurAttData(3) Then
                                    SetMapUndo MapX, MapY, 6, CLng(EditMap.Tile(MapX, MapY).Att), EditMap.Tile(MapX, MapY).AttData(0), EditMap.Tile(MapX, MapY).AttData(1), EditMap.Tile(MapX, MapY).AttData(2), EditMap.Tile(MapX, MapY).AttData(3)
                                    EditMap.Tile(MapX, MapY).Att = CurAtt
                                    EditMap.Tile(MapX, MapY).AttData(0) = CurAttData(0)
                                    EditMap.Tile(MapX, MapY).AttData(1) = CurAttData(1)
                                    EditMap.Tile(MapX, MapY).AttData(2) = CurAttData(2)
                                    EditMap.Tile(MapX, MapY).AttData(3) = CurAttData(3)
                                End If
                            Case 5 'Ground2
                                If EditMap.Tile(MapX, MapY).Ground2 <> CurTile Then
                                    SetMapUndo MapX, MapY, 2, CLng(EditMap.Tile(MapX, MapY).Ground2)
                                    EditMap.Tile(MapX, MapY).Ground2 = CurTile
                                End If
                            Case 7 'Wall
                                If EditMap.Tile(MapX, MapY).WallTile <> CurWall Then
                                    SetMapUndo MapX, MapY, 7, CLng(EditMap.Tile(MapX, MapY).WallTile)
                                    EditMap.Tile(MapX, MapY).WallTile = CurWall
                                End If
                        End Select
                    End If
                ElseIf Button = 2 Then
                    Select Case EditMode
                        Case 0 'Ground
                            EditMap.Tile(MapX, MapY).Ground = 0
                        Case 1 'BGTile1
                            EditMap.Tile(MapX, MapY).BGTile1 = 0
                        Case 2 'Anim
                            If keyAlt = False Then
                                EditMap.Tile(MapX, MapY).Anim(1) = 0
                                EditMap.Tile(MapX, MapY).Anim(2) = 0
                            End If
                        Case 3 'FGTile
                            EditMap.Tile(MapX, MapY).FGTile = 0
                        Case 4 'Attribute
                            EditMap.Tile(MapX, MapY).Att = 0
                            EditMap.Tile(MapX, MapY).AttData(0) = 0
                            EditMap.Tile(MapX, MapY).AttData(1) = 0
                            EditMap.Tile(MapX, MapY).AttData(2) = 0
                            EditMap.Tile(MapX, MapY).AttData(3) = 0
                        Case 5 'Ground2
                            EditMap.Tile(MapX, MapY).Ground2 = 0
                        Case 7 'Wall
                            EditMap.Tile(MapX, MapY).WallTile = 0
                    End Select
                End If
            End If
        ElseIf CastingSpell = True Then
            If Button = 1 Then
                'For A = 1 To MAXUSERS
                '    With player(A)
                '        If .map = CMap And .X = SpriteX And .Y = SpriteY And .Status <> 9 Then
                            'CurrentTarget.Target = A
                            'CurrentTarget.TargetType = TT_PLAYER
                            'CurrentTarget.X = .X
                            'CurrentTarget.Y = .Y
                            'Call UseSkill(SkillListBox.Data(SkillListBox.Selected))
                            'Exit Sub
                '        End If
                '    End With
                'Next A
                'If A = 256 Then
                '    For A = 0 To 9
                '        With map.Monster(A)
                '            If .Monster > 0 And .X = SpriteX And .Y = SpriteY Then
                                'TargetMonster = A
                                'CurrentTarget.Target = A
                                'CurrentTarget.TargetType = TT_MONSTER
                                'CurrentTarget.X = .X
                                'CurrentTarget.Y = .Y
                                'Call UseSkill(SkillListBox.Data(SkillListBox.Selected))
                                'Exit Sub
                '            End If
                '        End With
                '    Next A
                'End If
                If A = 10 Then
                    If cX = SpriteX And cY = SpriteY Then
                        'CurrentTarget.Target = Character.Index
                        'CurrentTarget.TargetType = TT_CHARACTER
                        'Call UseSkill(SkillListBox.Data(SkillListBox.Selected))
                        Exit Sub
                    End If
                End If
                If map.Tile(MapX, MapY).Att <> 11 And map.Tile(MapX, MapY).Att <> 14 And map.Tile(MapX, MapY).Att <> 19 Then
                    CurrentTarget.Target = 0
                    CurrentTarget.TargetType = TT_TILE
                    CurrentTarget.x = MapX
                    CurrentTarget.y = MapY
                    If Character.MacroSkill > 0 Then
                        Call UseSkill(Character.MacroSkill)
                    Else
                        If SkillListBox.Selected > 0 Then UseSkill (SkillListBox.Data(SkillListBox.Selected))
                    End If
                    Exit Sub
                End If
            End If
        Else
       
            If RemoteWidgetParent <> "" Then
                With Widgets.item(RemoteWidgetParent)
                    If .x * WindowScaleX <= Int(x) And (.x + .Width) * WindowScaleX >= Int(x) And .y * WindowScaleY <= Int(y) And (.y + .Height) * WindowScaleY >= Int(y) Then
                        If Widgets.item(RemoteWidgetParent).x <= Int(x) \ WindowScaleX And Widgets.item(RemoteWidgetParent).y <= Int(y) \ WindowScaleY And Widgets.item(RemoteWidgetParent).x + Widgets.item(RemoteWidgetParent).Width >= Int(x) \ WindowScaleX And Widgets.item(RemoteWidgetParent).y + Widgets.item(RemoteWidgetParent).Height >= Int(y) \ WindowScaleY Then
                            SendSocket Chr$(91) + DoubleChar(Int(x / WindowScaleX)) + DoubleChar(Int(y / WindowScaleY))
                        End If
                        'Widgets.item(RemoteWidgetParent).Key
                        WidgetMouseDown Widgets.item(RemoteWidgetParent).Children, CLng(Button), Int(x), Int(y)
                        Exit Sub
                    End If '
                End With
            End If 'Else
                If MapX >= 0 And MapX <= 11 And MapY >= 0 And MapY <= 11 Then
                    For A = 1 To 10
                        If map.Fish(A).x = MapX And map.Fish(A).y = MapY And map.Fish(A).TimeStamp > GetTickCount Then
                            SendSocket Chr$(90) + Chr$(MapX) + Chr$(MapY)
                            map.Fish(A).TimeStamp = 0 'Fish Click
                            map.Fish(A).x = 13
                            map.Fish(A).y = 13
                            Exit Sub
                        End If
                    Next A
                    If map.Tile(MapX, MapY).Att = 14 Then
                        If ExamineBit(map.Tile(MapX, MapY).AttData(0), 1) Then
                            SendSocket Chr$(76) + Chr$(0) + Chr(MapX) + Chr(MapY)
                            Exit Sub
                        End If
                    ElseIf (map.Tile(MapX, MapY).Att = 6 And ExamineBit(map.Tile(MapX, MapY).AttData(3), 0)) Then
                        SendSocket Chr$(76) + Chr$(0) + Chr(MapX) + Chr(MapY)
                        Exit Sub
                    End If
                    
                    If map.Tile(MapX, MapY).Att = 19 Then
                        If map.Tile(MapX, MapY).AttData(3) Then
                            SendSocket Chr$(76) + Chr$(0) + Chr(MapX) + Chr(MapY)
                            Exit Sub
                        End If
                    End If
        
                    For A = 1 To MAXUSERS
                        With player(A)
                            If .map = CMap And .x = SpriteX And .y = SpriteY And .Status <> 9 Then
                                    If Button = 1 Then
                                        If .Guild > 0 Then
                                            'St = St + .Name " a member of " + Guild(.Guild).Name & "."
                                        End If
                                        PrintChat St, 8, Options.FontSize
                                        CurrentTarget.TargetType = TT_PLAYER
                                        CurrentTarget.Target = A
                                        SendSocket Chr$(76) + Chr$(2) + Chr$(A)
                                    ElseIf Button = 2 Then
                                        CreatePlayerClickWindow A
                                    End If
                                Exit Sub
                            End If
                        End With
                    Next A
                    If A = MAXUSERS + 1 Then
                        For A = 0 To 9
                            With map.Monster(A)
                                If .Monster > 0 Then
                                    If .x = SpriteX And .y = SpriteY Then
                                        CurrentTarget.TargetType = TT_MONSTER
                                        CurrentTarget.Target = A
                                        TargetMonster = A
                                        SendSocket Chr$(76) + Chr$(3) + Chr$(A)
                                    End If
                                    If Monster(map.Monster(A).Monster).Flags2 And MONSTER_LARGE Then
                                        If .x + 1 = SpriteX And .y = SpriteY Then
                                            CurrentTarget.TargetType = TT_MONSTER
                                            CurrentTarget.Target = A
                                            TargetMonster = A
                                            SendSocket Chr$(76) + Chr$(3) + Chr$(A)
                                        End If
                                        If .x + 1 = SpriteX And .y + 1 = SpriteY Then
                                            CurrentTarget.TargetType = TT_MONSTER
                                            CurrentTarget.Target = A
                                            TargetMonster = A
                                            SendSocket Chr$(76) + Chr$(3) + Chr$(A)
                                        End If
                                        If .x = SpriteX And .y + 1 = SpriteY Then
                                            CurrentTarget.TargetType = TT_MONSTER
                                            CurrentTarget.Target = A
                                            TargetMonster = A
                                            SendSocket Chr$(76) + Chr$(3) + Chr$(A)
                                        End If
                                    
                                    End If
                                End If
                            End With
                            With map.DeadBody(A)
                                If .Sprite > 0 And .x = MapX And .y = MapY Then
                                    If .BodyType = TT_MONSTER Then
                                        PrintChat "You see the remains of a " & .Name & "!", 8, Options.FontSize
                                    ElseIf .BodyType = TT_PLAYER Then
                                        PrintChat "You see the remains of " & .Name & "!", 8, Options.FontSize
                                    End If
                                End If
                            End With
                        Next A
                    End If
                    If A = 10 Then
                        If cX = SpriteX And cY = SpriteY Then
                            If Button = 1 Then
                                CurrentTarget.TargetType = TT_CHARACTER
                                CurrentTarget.Target = 0
                                A = 1
                            Else
                                CreateCharacterClickWindow
                            End If
                        Else
                            If map.Tile(MapX, MapY).Att = 25 Then 'NPC Attribute
                                If Button = 2 Or Button = 1 Then
                                    CreateNPCClickWindow map.Tile(MapX, MapY).AttData(0), MapX, MapY
                                    LastNPCX = MapX
                                    LastNPCY = MapY
                                End If
                            End If
                            If map.Tile(MapX, MapY).Att = 3 Then 'Key
                                If Button = 2 Or Button = 1 Then
                                With map.Tile(MapX, MapY)
                                    b = .AttData(0) + .AttData(3) * 256&
                                    If b > 0 Then PrintChat "This door can be opened with a " & Object(b).Name, 8, Options.FontSize, 6
                                End With
                                    
                                
                                    
                                End If
                            End If
                        End If
                    End If
                    'If A = 10 Then
                        b = 0
                        For A = 0 To 49
                            With map.Object(A)
                                If .x = MapX And .y = MapY Then
                                    If .Object > 0 Then
                                        If b = 0 Then b = A + 56
                                        If 56 + A > CurInvObj Then
                                            b = 56 + A
                                            CurInvObj = 56 + A
                                            SetTab tsInventory
                                            DrawCurInvObj
                                            Exit For
                                        End If
                                    End If
                                End If
                            End With
                        Next A
                        If b <> CurInvObj And b > 0 Then
                            CurInvObj = b
                            SetTab tsInventory
                            DrawCurInvObj
                        End If
                    'End If
                
                'If Character.GuildRank = 3 Then
                If Character.Guild > 0 Then
                    For x = 0 To 11
                    For y = 0 To 11
                        If map.Tile(x, y).Att = 8 Then
                            If map.Tile(x, y).AttData(2) > 0 Then
                                If map.Tile(x, y).AttData(2) = Guild(Character.Guild).hallNum And map.Tile(x, y).AttData(0) = MapX And map.Tile(x, y).AttData(1) = MapY Then
                                     If Not frmHallAccess_Loaded Then Load frmHallAccess
                                        With frmHallAccess
                                            If Character.GuildRank = 3 Then
                                                .x = MapX
                                                .y = MapY
                                                .map = CMap
                                                .rdRank(0).Value = True
                                                .default = 0
                                                If ExamineBit(map.Tile(x, y).AttData(3), 3) Then
                                                    .rdRank(1).Value = True
                                                    .default = 1
                                                End If
                                                If ExamineBit(map.Tile(x, y).AttData(3), 4) Then
                                                    .rdRank(2).Value = True
                                                    .default = 2
                                                End If
                                                If ExamineBit(map.Tile(x, y).AttData(3), 5) Then
                                                    .rdRank(3).Value = True
                                                    .default = 3
                                                End If
                                                
                    
                                                .Show
                                                x = 12
                                                y = 12
                                            Else
                                                If ExamineBit(map.Tile(x, y).AttData(3), 2) Then PrintChat "This door is only accessible by everyone in the guild.", 8, Options.FontSize, 6
                                                If ExamineBit(map.Tile(x, y).AttData(3), 3) Then PrintChat "This door is accessible by members.", 8, Options.FontSize, 6
                                                If ExamineBit(map.Tile(x, y).AttData(3), 4) Then PrintChat "This door is accessible by the guild lords.", 8, Options.FontSize, 6
                                                If ExamineBit(map.Tile(x, y).AttData(3), 5) Then PrintChat "This door is only accessible by the guildmaster.", 8, Options.FontSize, 6
                                            End If
                                            Exit For
                                            
                                        End With
                                        
                                End If
                            End If
                        End If
                    Next y
                    Next x
                End If
                'end if
            End If
            'End If
        End If
    End If
End Sub

Private Sub picViewport_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    CurX = Int(x / (32)) '* WindowScaleX))
    CurSubX = x Mod (32) ' * WindowScaleX)
    CurY = Int(y / (32)) '* WindowScaleY))
    CurSubY = y Mod (32) '* WindowScaleY)
    hoverCombat = False
    If MapEdit = True Then
        picViewport_MouseDown Button, Shift, x, y
    Else
        If RemoteWidgetParent <> "" Then WidgetMouseMove Widgets.item(RemoteWidgetParent).Children, CLng(x), CLng(y)
    End If
End Sub


Public Sub LstAddItem(Caption As String, Data As Byte, Optional Index As Byte)
If Index < 1 Or Index > 255 Then
    For Index = 1 To 255
        If SkillListBox.Caption(Index) = "" Then Exit For
        If Index = 255 Then Exit Sub
    Next Index
End If

    With SkillListBox
        .Caption(Index) = Caption
        .Data(Index) = Data
        '.YOffset = 0
    End With
    DrawLstBox
End Sub

Private Sub lstSkills_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim A As Long
x = x / WindowScaleX
y = y / WindowScaleY
lstSkills.Tag = x
If Button = 1 Then
    SetCapture lstSkills.hwnd
    With SkillListBox
        .Selected = Int(y / 15) + 1 + .YOffset
        .MouseState = True
        If x > 180 Then 'Train it
            A = .Data(.Selected)
            If A <= MAX_SKILLS And A <> SKILL_INVALID Then
                If Character.SkillLevels(A) < Skills(A).MaxLevel Then
                    If Character.skillPoints > 0 Then
                        Character.skillPoints = Character.skillPoints - 1
                        Character.SkillLevels(A) = Character.SkillLevels(A) + 1
                        SendSocket Chr$(30) + Chr$(1) + Chr$(A)
                        UpdateSkills
                        If frmSkillTree.Visible Then Unload frmSkillTree
                    End If
                End If
            End If
        End If
        DrawLstBox
        drawSkills
    End With
ElseIf Button = 2 Then
    With SkillListBox
        .Selected = Int(y / 15) + 1 + .YOffset
        .MouseState = True
        DrawLstBox
'        If .Data(.Selected) > 0 Then
'            'Set Macro
'            For B = vbKeyF1 To vbKeyF10
'                If GetKeyState(B) < 0 Then
'                    If .Selected > 0 Then
'                        If .Data(.Selected) <= MAX_SKILLS Then
'                            Macro(B - vbKeyF1).Skill = .Data(.Selected)
'                            WriteString "Macros", "Skill" + CStr(B - vbKeyF1), .Data(.Selected)
'                            DrawLstBox
'                            Exit For
'                        End If
'                    End If
'                End If
'            Next B
'        End If
    End With
End If
End Sub

Private Sub lstSkills_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim A As Byte, b As Byte
x = x / WindowScaleX
y = y / WindowScaleY
With SkillListBox
b = .Selected
    If .MouseState = True Then
        If y < 0 Then
            If .YOffset > 0 Then .YOffset = .YOffset - 1
            .Selected = .YOffset + 1
            A = 1
        End If
        If y > (165) Then
            If .YOffset < 255 - 14 And .Selected < 255 Then
                .YOffset = .YOffset + 1
                .Selected = .YOffset + 13
                A = 1
            End If
        End If
        If A = 0 Then
            If Int(y / 15) + 1 + .YOffset <= 254 Then
                .Selected = Int(y / 15) + 1 + .YOffset
            End If
        End If
        If b <> .Selected Then
            DrawLstBox
        End If
    End If
End With
End Sub

Private Sub lstSkills_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
SkillListBox.MouseState = False
DrawLstBox
ReleaseCapture
End Sub

Public Sub DrawLstBox()
Dim A As Long, C As Long, St As String
lstSkills.Cls
lstSkills.FontSize = 8 * WindowScaleY
    With SkillListBox
        For A = .YOffset To .YOffset + 17
            If .Data(A + 1) <= MAX_SKILLS And .Data(A + 1) <> SKILL_INVALID Then
                If GetTickCount < Character.LocalSpellTick(.Data(A + 1)) Then
                    C = lstSkills.Width * ((Character.LocalSpellTick(.Data(A + 1)) - GetTickCount) / Skills(.Data(A + 1)).LocalTick)
                    lstSkills.ForeColor = &H363636
                    lstSkills.FillColor = &H363636
                    lstSkills.Line (lstSkills.Width - C, (A - .YOffset) * WindowScaleY * 15)-(lstSkills.Width, (A - .YOffset) * WindowScaleY * 15 + 15 * WindowScaleY), &H464646, BF
                    RedrawSkills = True
                End If
                If (A + 1 = .Selected) And .MouseState Then
                    lstSkills.ForeColor = vbBlue
                    lstSkills.FillColor = vbBlack
                    Rectangle lstSkills.hdc, 0, (A - .YOffset) * 15 * WindowScaleY, lstSkills.ScaleWidth * WindowScaleX, (A - .YOffset) * 15 * WindowScaleY + 15 * WindowScaleY
                ElseIf A + 1 = .Selected Then
                    lstSkills.ForeColor = vbGreen
                    lstSkills.FillColor = vbBlack
                    Rectangle lstSkills.hdc, 0, (A - .YOffset) * 15 * WindowScaleY, lstSkills.ScaleWidth * WindowScaleX, (A - .YOffset) * 15 * WindowScaleY + 15 * WindowScaleY
                End If
                lstSkills.ForeColor = vbWhite
                St = .Caption(A + 1)
                For C = 0 To 9
                    If Macro(C).Skill = .Data(A + 1) Then
                        St = St & " [" & KeyCodeList(Options.SpellKey(C)).Text & "]"
                    End If
                Next C
                lstSkills.ForeColor = Skills(.Data(A + 1)).Color
                TextOut lstSkills.hdc, 0, (A - .YOffset) * 15 * WindowScaleY, St, Len(St)
                St = Character.SkillLevels(.Data(A + 1)) & "/" & Skills(.Data(A + 1)).MaxLevel
                TextOut lstSkills.hdc, 145 * WindowScaleX, (A - .YOffset) * 15 * WindowScaleY, St, Len(St)
                If Character.SkillLevels(.Data(A + 1)) < Skills(.Data(A + 1)).MaxLevel Then
                    If Character.skillPoints > 0 Then
                        TextOut lstSkills.hdc, 182 * WindowScaleX, (A - .YOffset) * 15 * WindowScaleY, "+", 1
                    End If
                End If
            End If
        Next A
    End With
    lstSkills.Refresh
End Sub

Private Sub picViewport_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not MapEdit Then
        If RemoteWidgetParent <> "" Then WidgetMouseUp Widgets.item(RemoteWidgetParent).Children, CLng(Button), CLng(x), CLng(y)
    End If
End Sub

Private Sub txtDrop_Change()
If Val(txtDrop) > 2147483647 Then
    txtDrop = 2147483647
End If
txtDrop = Val(txtDrop)
TempVar1 = Val(txtDrop)
End Sub

Private Sub txtDrop_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Or KeyAscii = 10 Then
    lblMenu_MouseDown 35, 1, 0, 0, 0
    lblMenu_MouseUp 35, 1, 0, 0, 0
End If
End Sub
