VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOk 
      Caption         =   "Change"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1800
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "New Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
End Sub


Private Sub btnOk_Click()
    If Len(txtPassword) >= 3 And ValidName(txtPassword) And InStr(txtPassword, " ") = 0 Then
        Password = txtPassword
        DataRS.Edit
        DataRS!Password = Password
        DataRS.Update
        SendSR Chr$(6) + Password
        Unload Me
    Else
        MsgBox "Password must be between 3 and 15 characters long, must start with a letter, and may only contain letters, numbers, and the underscore character.", vbOKOnly, TitleString
    End If
End Sub


Private Sub Form_Load()
    frmPassword_Open = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmPassword_Open = False
End Sub


