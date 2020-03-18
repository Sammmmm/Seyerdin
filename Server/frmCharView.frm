VERSION 5.00
Begin VB.Form frmCharView 
   Caption         =   " "
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   6075
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstPlayers 
      Height          =   5130
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblY 
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblX 
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblMap 
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblDirections 
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblName 
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Y:"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "X:"
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Map:"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Direction:"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmCharView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim A As Long
lstPlayers.Clear
For A = 1 To GetMaxUsers
    lstPlayers.AddItem A & ". " & Player(A).Name
Next A
End Sub

Private Sub lblDirection_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub lstPlayers_Click()
With lstPlayers
    If .ListIndex >= 0 Then
        With Player(.ListIndex + 1)
            lblDirection = .D
            lblName = .Name
            lblX = .X
            lblY = .Y
            lblMap = .Map
        End With
    End If
End With
End Sub
