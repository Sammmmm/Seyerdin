VERSION 5.00
Begin VB.Form frmName 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server Description"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOk 
      Caption         =   "Update"
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.TextBox txtDescription 
      Height          =   1455
      Left            =   1680
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Description:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frmName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    If SRConnected Then
        SRName = txtName
        SRDescription = txtDescription
        SendSR Chr$(5) + SRName + Chr$(0) + SRDescription
    End If
    Unload Me
End Sub


Private Sub Form_Load()
    txtName = SRName
    txtDescription = SRDescription
    frmName_Open = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmName_Open = False
End Sub


