VERSION 5.00
Begin VB.Form frmCard 
   Caption         =   "Edit Card"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar sclFront 
      Height          =   255
      Left            =   1200
      Max             =   255
      Min             =   1
      TabIndex        =   12
      Top             =   1800
      Value           =   1
      Width           =   3495
   End
   Begin VB.HScrollBar SclBack 
      Height          =   255
      Left            =   1200
      Max             =   255
      Min             =   1
      TabIndex        =   11
      Top             =   2520
      Value           =   1
      Width           =   3495
   End
   Begin VB.HScrollBar sclSides 
      Height          =   255
      Left            =   1200
      Max             =   255
      Min             =   1
      TabIndex        =   10
      Top             =   2160
      Value           =   1
      Width           =   3495
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   4320
      TabIndex        =   8
      Top             =   4560
      Width           =   1455
   End
   Begin VB.ListBox lstType 
      Appearance      =   0  'Flat
      Height          =   810
      ItemData        =   "frmCard.frx":0000
      Left            =   4920
      List            =   "frmCard.frx":0010
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
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
      TabIndex        =   2
      Top             =   480
      Width           =   3855
   End
   Begin VB.HScrollBar sclPicture 
      Height          =   255
      Left            =   1200
      Max             =   255
      Min             =   1
      TabIndex        =   1
      Top             =   1080
      Value           =   1
      Width           =   3495
   End
   Begin VB.PictureBox picSprite 
      AutoRedraw      =   -1  'True
      Height          =   540
      Left            =   5160
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   480
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Back:"
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
      TabIndex        =   15
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Sides:"
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
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Front:"
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
      TabIndex        =   13
      Top             =   1800
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
      Left            =   0
      TabIndex        =   6
      Top             =   480
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
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   1095
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
      Left            =   0
      TabIndex        =   4
      Top             =   0
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
      Left            =   1200
      TabIndex        =   3
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnOk_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
lstType.Selected(0) = True
End Sub

Private Sub sclPicture_Change()
             Dim hdcObjects As Long, A As Long, B As Long
    
    If lstType.Selected(0) Or lstType.Selected(1) Then
        lblSprite.Caption = sclSprite
        picSprite.Height = 36
        picSprite.Width = 36
        Dim R1 As RECT, R2 As RECT
        A = sclSprite - 1
        R1.Top = 0: R1.Bottom = 32: R1.Left = 0: R1.Right = 32
        R2.Top = (A \ 16) * 32: R2.Left = (A Mod 16) * 32: R2.Right = R2.Left + 32: R2.Bottom = R2.Top + 32
        On Error Resume Next
            sfcSprites.Surface.BltToDC picSprite.hdc, R2, R1
        On Error GoTo 0
        picSprite.Refresh
    End If
    If lstType.Selected(2) Then

    
        A = sclPicture - 1
        B = A Mod 64
        A = (A \ 64) + 1
    
        hdcObjects = sfcObjects(A).Surface.GetDC
            BitBlt picSprite.hdc, 0, 0, 32, 32, hdcObjects, (B Mod 8) * 32, (B \ 8) * 32, SRCCOPY
        sfcObjects(A).Surface.ReleaseDC hdcObjects
        
        picSprite.Refresh
    End If
    If lstType.Selected(3) Then
        
    End If
    
End Sub
