VERSION 5.00
Begin VB.Form frmAccount 
   BorderStyle     =   0  'None
   Caption         =   "Seyerdin Online [New Account]"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5460
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   2400
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtPass1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label btnOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label btnCancel 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   2640
      Width           =   1455
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnCancel_Click()
    Unload Me
    frmMenu.Show
End Sub


Private Sub btnOk_Click()
    Dim St As String, A As Long
    txtUser = Trim$(txtUser)
    txtPass1 = Trim$(txtPass1)
    txtPass2 = Trim$(txtPass2)
    
    If Len(txtUser) >= 3 Then
        If Len(txtPass1) >= 3 Then
            A = Asc(Left$(txtUser, 1))
            If (A >= 65 And A <= 90) Or (A >= 97 And A <= 122) Then
                If UCase$(txtPass1) = UCase$(txtPass2) Then
                    User = txtUser
                    Pass = txtPass1
                    NewAccount = True
                    
                    WriteString "Login", "User", User
                    
                    frmWait.Show
                    frmWait.lblStatus = "Connecting ..."
                    frmWait.Refresh
                    
                    ClientSocket = ConnectSock(ServerIP, ServerPort, St, gHW, True)
                    
                    Me.Hide
                Else
                    MsgBox "Your two passwords do not match, please re-enter!", vbOKOnly, TitleString
                End If
            Else
                MsgBox "User name must start with a letter!", vbOKOnly, TitleString
            End If
        Else
            MsgBox "Password must be atleast 3 characters long!", vbOKOnly + vbExclamation, TitleString
        End If
    Else
        MsgBox "User name must be atleast 3 characters long!", vbOKOnly + vbExclamation, TitleString
    End If
End Sub
Private Sub Form_Load()
    frmAccount_Loaded = True
    Set Me.Picture = LoadPicture("Data/Graphics/Interface/Account.rsc")
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Me.WindowState = 0 Then
       Dim ReturnVal As Long
       ReleaseCapture
       ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmAccount_Loaded = False
    Set Me.Picture = Nothing
End Sub


Private Sub txtPass1_Change()
    If txtUser <> "" And txtPass1 <> "" And txtPass2 <> "" Then
        If btnOk.Enabled = False Then
            btnOk.Enabled = True
        End If
    Else
        If btnOk.Enabled = True Then
            btnOk.Enabled = False
        End If
    End If
End Sub

Private Sub txtPass1_GotFocus()
    txtPass1.SelStart = 0
    txtPass1.SelLength = Len(txtPass1)
End Sub


Private Sub txtPass2_Change()
    If txtUser <> "" And txtPass1 <> "" And txtPass2 <> "" Then
        If btnOk.Enabled = False Then
            btnOk.Enabled = True
        End If
    Else
        If btnOk.Enabled = True Then
            btnOk.Enabled = False
        End If
    End If
End Sub

Private Sub txtPass2_GotFocus()
    txtPass2.SelStart = 0
    txtPass2.SelLength = Len(txtPass2)
End Sub


Private Sub txtUser_Change()
    If txtUser <> "" And txtPass1 <> "" And txtPass2 <> "" Then
        If btnOk.Enabled = False Then
            btnOk.Enabled = True
        End If
    Else
        If btnOk.Enabled = True Then
            btnOk.Enabled = False
        End If
    End If
End Sub
Private Sub txtUser_GotFocus()
    txtUser.SelStart = 0
    txtUser.SelLength = Len(txtUser)
End Sub


Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 95 Then
        'Valid Key
    Else
        KeyAscii = 0
        Beep
    End If
End Sub


