VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Odyssey Classic Server Login"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Login"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CheckBox chkNewAccount 
      Caption         =   "Create a new account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1080
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   1080
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Pass:"
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
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "User:"
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
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnCancel_Click()
    Unload Me
End Sub


Private Sub btnOk_Click()
    If ValidName(txtUser) And ValidName(txtPassword) And Len(txtUser) >= 3 And Len(txtPassword) >= 3 And InStr(txtUser, " ") = 0 And InStr(txtPassword, " ") = 0 Then
        User = txtUser
        Password = txtPassword
        If chkNewAccount = 0 Then
            NewAccount = False
        Else
            NewAccount = True
        End If
        UseRegistry = True
        SRConnect
        Unload Me
    Else
        MsgBox "User name and password must be between 3 and 15 characters long, must start with a letter, and may only contain letters, numbers, and the underscore character.", vbOKOnly, TitleString
    End If
End Sub


Private Sub Form_Load()
    txtUser = User
    txtPassword = Password
    If NewAccount = True Then
        chkNewAccount = 1
    Else
        chkNewAccount = 0
    End If
    frmLogin_Open = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmLogin_Open = False
End Sub


