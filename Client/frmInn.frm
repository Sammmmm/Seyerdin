VERSION 5.00
Begin VB.Form frmInn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Odyssey Online Classic [Editing Inn]"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar sclCost 
      Height          =   255
      Left            =   1080
      Max             =   1000
      Min             =   1
      TabIndex        =   19
      Top             =   3240
      Value           =   1
      Width           =   4695
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3360
      TabIndex        =   18
      Top             =   3600
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   4920
      TabIndex        =   17
      Top             =   3600
      Width           =   1455
   End
   Begin VB.HScrollBar sclStartY 
      Height          =   255
      Left            =   1080
      Max             =   11
      TabIndex        =   13
      Top             =   2880
      Width           =   4695
   End
   Begin VB.HScrollBar sclStartX 
      Height          =   255
      Left            =   1080
      Max             =   11
      TabIndex        =   12
      Top             =   2520
      Width           =   4695
   End
   Begin VB.HScrollBar sclStartMap 
      Height          =   255
      Left            =   1080
      Max             =   5000
      Min             =   1
      TabIndex        =   11
      Top             =   2160
      Value           =   1
      Width           =   4695
   End
   Begin VB.TextBox txtMessage2 
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
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   9
      Top             =   1560
      Width           =   5175
   End
   Begin VB.TextBox txtMessage1 
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
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   8
      Top             =   1080
      Width           =   5175
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
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   7
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label lblCost 
      Alignment       =   2  'Center
      Caption         =   "1"
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
      Left            =   5880
      TabIndex        =   20
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label lblStartY 
      Alignment       =   2  'Center
      Caption         =   "0"
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
      Left            =   5880
      TabIndex        =   16
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label lblStartX 
      Alignment       =   2  'Center
      Caption         =   "0"
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
      Left            =   5880
      TabIndex        =   15
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lblStartMap 
      Alignment       =   2  'Center
      Caption         =   "1"
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
      Left            =   5880
      TabIndex        =   14
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label7 
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
      TabIndex        =   10
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label6 
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
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label5 
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
      TabIndex        =   5
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Message2:"
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
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Message1:"
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
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
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
      TabIndex        =   2
      Top             =   600
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
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
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
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmInn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnOk_Click()
Dim St As String
Dim Name As String * 15
Name = txtName.Text
St = Chr$(lblNumber) + DoubleChar(sclStartMap) + Chr$(sclStartX) + Chr$(sclStartY) + DoubleChar(sclCost) + Name + txtMessage1.Text + Chr$(0) + txtMessage2.Text
SendSocket Chr$(67) + St
Unload Me
End Sub

Private Sub Form_Load()
frmInn_Loaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmInn_Loaded = False
End Sub

Private Sub sclCost_Change()
    lblCost = sclCost
End Sub

Private Sub sclCost_Scroll()
    lblCost = sclCost
End Sub

Private Sub sclStartMap_Change()
    lblStartMap = sclStartMap
End Sub

Private Sub sclStartMap_Scroll()
    lblStartMap = sclStartMap
End Sub

Private Sub sclStartX_Change()
    lblStartX = sclStartX
End Sub

Private Sub sclStartX_Scroll()
    lblStartX = sclStartX
End Sub

Private Sub sclStartY_Change()
    lblStartY = sclStartY
End Sub

Private Sub sclStartY_Scroll()
    lblStartY = sclStartY
End Sub
