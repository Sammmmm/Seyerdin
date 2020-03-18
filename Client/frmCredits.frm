VERSION 5.00
Begin VB.Form frmCredits 
   BorderStyle     =   0  'None
   Caption         =   "Seyerdin Online [Credits]"
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   292
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label btnOk 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   3840
      Width           =   1455
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnOk_Click()
    Unload Me
    frmMenu.Show
End Sub

Private Sub Form_Load()
    Set Me.Picture = LoadPicture("Data/Graphics/Interface/Credits.rsc")
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
