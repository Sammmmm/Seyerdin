VERSION 5.00
Begin VB.Form frmNewCharacter 
   BorderStyle     =   0  'None
   Caption         =   "Seyerdin Online [New Character]"
   ClientHeight    =   5235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4380
   ControlBox      =   0   'False
   Icon            =   "frmNewCharacter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   349
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   292
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesc 
      Height          =   855
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3720
      Width           =   4215
   End
   Begin VB.Timer SpriteTimer 
      Interval        =   250
      Left            =   0
      Top             =   0
   End
   Begin VB.OptionButton optGender 
      BackColor       =   &H00404040&
      Caption         =   "Female"
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Top             =   1575
      Width           =   195
   End
   Begin VB.OptionButton optGender 
      BackColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   1575
      Value           =   -1  'True
      Width           =   195
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      Height          =   540
      Left            =   3795
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   2400
      Width           =   540
   End
   Begin VB.ComboBox cmbClass 
      Height          =   315
      ItemData        =   "frmNewCharacter.frx":000C
      Left            =   1560
      List            =   "frmNewCharacter.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      MaxLength       =   15
      TabIndex        =   0
      ToolTipText     =   $"frmNewCharacter.frx":0010
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label btnCancel 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   1320
      TabIndex        =   16
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label btnOK 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   2880
      TabIndex        =   15
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label lblConstitution 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label lblWisdom 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label lblIntelligence 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblEndurance 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   11
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblAgility 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   10
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label lblStrength 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblMana 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label lblEnergy 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   1080
      TabIndex        =   7
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label lblHP 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   2520
      Width           =   735
   End
End
Attribute VB_Name = "frmNewCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim D As Byte, W As Byte, A As Byte
Private Sub btnOk_Click()
    Dim Gender As Byte, A As Long
    
    If Len(txtName) >= 3 Then
        A = Asc(Mid$(txtName, 1, 1))
        If (A >= 65 And A <= 90) Or (A >= 97 And A <= 122) Then
            If optGender(0) = True Then Gender = 0 Else Gender = 1
                
            frmWait.Show
            frmWait.Caption = "Creating new character ..."
            
            SendSocket Chr$(2) + Chr$(cmbClass.ListIndex + 1) + Chr$(1) + Chr$(Gender) + txtName + Chr$(0) + txtDesc
            
            Me.Hide
        Else
            MsgBox "Name must start with a letter!", vbOKOnly + vbExclamation, TitleString
        End If
    Else
        MsgBox "Name must be atleast 3 characters long!", vbOKOnly + vbExclamation, TitleString
    End If
End Sub

Private Sub cmbClass_Click()
    Select Case cmbClass.ListIndex
        Case 0
            lblHP = 30 + 10
            lblEnergy = 30
            lblMana = 30
        Case 1
            lblHP = 30
            lblEnergy = 30
            lblMana = 30 + 10
        Case 2
            lblHP = 30 + 5
            lblEnergy = 30 + 5
            lblMana = 30
        Case 3
            lblHP = 30
            lblEnergy = 30
            lblMana = 30 + 10
    End Select
    lblStrength = 10
    lblAgility = 10
    lblEndurance = 10
    lblWisdom = 10
    lblConstitution = 10
    lblIntelligence = 10
End Sub

Private Sub Form_Load()
    Dim A As Long
    For A = 1 To MAX_CLASS
        cmbClass.AddItem Class(A).Name
    Next A
    cmbClass.ListIndex = 0
    Set Me.Picture = LoadPicture("Data/Graphics/Interface/NewCharacter.rsc")
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Me.WindowState = 0 Then
       Dim ReturnVal As Long
       ReleaseCapture
       ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Me.Picture = Nothing
End Sub


Private Sub Label7_Click()

End Sub

Private Sub SpriteTimer_Timer()
    Dim Frame As Byte, tSurf As Direct3DSurface8, tDC As Long
    If Int(Rnd * 10) = 0 Then
        A = 2
    End If
    
    If A > 0 Then
        A = A - 1
        Frame = D * 3 + 2
    Else
        Frame = D * 3 + W
        W = 1 - W
        If Int(Rnd * 10) = 0 Then
            D = (D + 1) Mod 4
        End If
    End If
    
    Dim R1 As RECT, R2 As RECT
    A = (cmbClass.ListIndex * 2)
    R1.Top = (A \ 16) * 32: R1.Bottom = R1.Top + 32: R1.Left = (A Mod 16) * 32: R1.Right = R1.Left + 32
    R2.Top = 0: R2.Left = 0: R2.Bottom = 32: R2.Right = 32
    If optGender(0) = True Then
        sfcSprites.Surface.BltToDC picSprite.hdc, R1, R2
        picSprite.Refresh
    Else
        R1.Top = R1.Top + 32
        R1.Bottom = R1.Top + 32
        sfcSprites.Surface.BltToDC picSprite.hdc, R1, R2
        picSprite.Refresh
    End If
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or (KeyAscii >= 32 And KeyAscii <= 127) Then
        'Valid Key
    Else
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub txtName_Change()
    If txtName <> "" Then
        If btnOk.Enabled = False Then
            btnOk.Enabled = True
        End If
    Else
        If btnOk.Enabled = True Then
            btnOk.Enabled = False
        End If
    End If
End Sub
Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Then
        'Valid Key
    Else
        KeyAscii = 0
        Beep
    End If
End Sub
