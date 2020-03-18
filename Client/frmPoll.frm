VERSION 5.00
Begin VB.Form frmPoll 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seyerdin Online [Poll]"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3075
   ControlBox      =   0   'False
   Icon            =   "frmPoll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   3075
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMax 
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Text            =   "1"
      Top             =   600
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   480
      TabIndex        =   13
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   1800
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtChoice 
      Height          =   285
      Index           =   4
      Left            =   1080
      TabIndex        =   11
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtChoice 
      Height          =   285
      Index           =   3
      Left            =   1080
      TabIndex        =   10
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtChoice 
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   9
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtChoice 
      Height          =   285
      Index           =   1
      Left            =   1080
      TabIndex        =   8
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtChoice 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1080
      MaxLength       =   75
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
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
      TabIndex        =   14
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Choice5:"
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
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Choice4:"
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
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Choice3:"
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
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Choice2:"
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
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Choice1:"
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
      TabIndex        =   1
      Top             =   1200
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
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmPoll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
If Not IsNumeric(txtMax) And Val(txtMax) > 0 And Val(txtMax) <= 5 Then
    MsgBox "You must make enter a maximum between 1 and 5!", vbOKOnly
    Exit Sub
End If

If Trim$(txtName) = "" Then
    MsgBox "You must enter a topic!", vbOKOnly
    Exit Sub
End If

If Trim$(txtChoice(0).Text) = "" Then
    MsgBox "The first choice must not be blank!", vbOKOnly
    Exit Sub
End If

SendSocket Chr$(18) + Chr$(20) + Chr$(2) + Chr$(txtMax) + txtName + Chr$(0) + txtChoice(0) + Chr$(0) + txtChoice(1) + Chr$(0) + txtChoice(2) + Chr$(0) + txtChoice(3) + Chr$(0) + txtChoice(4)
Unload Me
End Sub
