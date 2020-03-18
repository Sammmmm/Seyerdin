VERSION 5.00
Begin VB.Form frmNPC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seyerdin Online [Editing NPC]"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   ControlBox      =   0   'False
   Icon            =   "frmNPC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   516
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkFlag 
      Caption         =   "Shop"
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   45
      Top             =   840
      Width           =   1215
   End
   Begin VB.HScrollBar sclDirection 
      Height          =   255
      Left            =   6120
      Max             =   3
      TabIndex        =   44
      Top             =   1680
      Width           =   1815
   End
   Begin VB.PictureBox picSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   8520
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   39
      Top             =   60
      Width           =   480
   End
   Begin VB.HScrollBar sclSprite 
      Height          =   255
      Left            =   6120
      Max             =   255
      TabIndex        =   37
      Top             =   600
      Value           =   1
      Width           =   2535
   End
   Begin VB.PictureBox picPortrait 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   4560
      ScaleHeight     =   960
      ScaleWidth      =   960
      TabIndex        =   36
      Top             =   480
      Width           =   960
   End
   Begin VB.HScrollBar sclPortrait 
      Height          =   255
      Left            =   1320
      Max             =   100
      TabIndex        =   35
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Can Repair"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   33
      Top             =   480
      Width           =   1215
   End
   Begin VB.CheckBox chkFlag 
      Caption         =   "Banker"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   31
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6120
      TabIndex        =   14
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   7680
      TabIndex        =   15
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton btnUpdate 
      Caption         =   "<-- Update"
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   6480
      Width           =   3855
   End
   Begin VB.TextBox txtTakeValue 
      Height          =   315
      Left            =   8160
      MaxLength       =   9
      TabIndex        =   12
      Top             =   6120
      Width           =   975
   End
   Begin VB.ComboBox cmbTakeObject 
      Height          =   315
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox txtGiveValue 
      Height          =   315
      Left            =   8160
      MaxLength       =   9
      TabIndex        =   10
      Top             =   5760
      Width           =   975
   End
   Begin VB.ComboBox cmbGiveObject 
      Height          =   315
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   5760
      Width           =   1695
   End
   Begin VB.ListBox lstSaleItems 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   1320
      TabIndex        =   8
      Top             =   5400
      Width           =   3855
   End
   Begin VB.TextBox txtSayText5 
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
      Left            =   1320
      MaxLength       =   255
      TabIndex        =   7
      Top             =   4920
      Width           =   7815
   End
   Begin VB.TextBox txtSayText4 
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
      Left            =   1320
      MaxLength       =   255
      TabIndex        =   6
      Top             =   4440
      Width           =   7815
   End
   Begin VB.TextBox txtSayText3 
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
      Left            =   1320
      MaxLength       =   255
      TabIndex        =   5
      Top             =   3960
      Width           =   7815
   End
   Begin VB.TextBox txtSayText2 
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
      Left            =   1320
      MaxLength       =   255
      TabIndex        =   4
      Top             =   3480
      Width           =   7815
   End
   Begin VB.TextBox txtSayText1 
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
      Left            =   1320
      MaxLength       =   255
      TabIndex        =   3
      Top             =   3000
      Width           =   7815
   End
   Begin VB.TextBox txtLeaveText 
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
      Left            =   1320
      MaxLength       =   255
      TabIndex        =   2
      Top             =   2520
      Width           =   7815
   End
   Begin VB.TextBox txtJoinText 
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
      Left            =   1320
      MaxLength       =   255
      TabIndex        =   1
      Top             =   2040
      Width           =   7815
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
      Left            =   1320
      MaxLength       =   15
      TabIndex        =   0
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label17 
      Caption         =   "Direction:"
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
      Left            =   6120
      TabIndex        =   43
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblDirection 
      Alignment       =   2  'Center
      Caption         =   "Up"
      Height          =   255
      Left            =   8040
      TabIndex        =   42
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblPortrait 
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
      Left            =   3960
      TabIndex        =   41
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblSprite 
      Alignment       =   2  'Center
      Caption         =   "1"
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
      Left            =   8640
      TabIndex        =   40
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label16 
      Caption         =   "Sprite:"
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
      Left            =   6120
      TabIndex        =   38
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label15 
      Caption         =   "Portrait"
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
      TabIndex        =   34
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label14 
      Caption         =   "Flags:"
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
      TabIndex        =   32
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblItemNumber 
      Height          =   375
      Left            =   6360
      TabIndex        =   30
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label13 
      Caption         =   "Item Number:"
      Height          =   375
      Left            =   5280
      TabIndex        =   29
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "TakeObject:"
      Height          =   375
      Left            =   5280
      TabIndex        =   28
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "GiveObject:"
      Height          =   375
      Left            =   5280
      TabIndex        =   27
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Sale Items:"
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
      TabIndex        =   26
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "SayText5:"
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
      TabIndex        =   25
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "SayText4:"
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
      TabIndex        =   24
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "SayText3:"
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
      TabIndex        =   23
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "SayText2:"
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
      TabIndex        =   22
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "SayText1:"
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
      TabIndex        =   21
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "LeaveText:"
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
      TabIndex        =   20
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "JoinText:"
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
      TabIndex        =   19
      Top             =   2040
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
      Left            =   120
      TabIndex        =   18
      Top             =   1560
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
      Left            =   120
      TabIndex        =   17
      Top             =   120
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
      Left            =   1320
      TabIndex        =   16
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmNPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    Seyerdin Online - A MMO RPG based on Odyssey Online Classic - In memory of Clay Rance
'    Copyright (C) 2020  Samuel Cook and Eric Robinson
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.

Option Explicit
Private Sub btnCancel_Click()
    Me.Hide
End Sub


Private Sub btnOk_Click()
    Dim A As Long, St As String
    Dim Flags As Byte
    
    For A = 0 To 2
        If chkFlag(A) = 1 Then
            SetBit Flags, CByte(A)
        Else
            ClearBit Flags, CByte(A)
        End If
    Next A
    
    St = Chr$(51) + Chr$(lblNumber) + Chr$(Flags) + Chr$(sclPortrait) + Chr$(sclSprite) + Chr$(sclDirection)
    For A = 0 To 9
        With SaleItem(A)
            St = St + DoubleChar(.GiveObject) + QuadChar(.GiveValue) + DoubleChar(.TakeObject) + QuadChar(.TakeValue)
        End With
    Next A
    
    St = St + txtName + Chr$(0) + txtJoinText + Chr$(0) + txtLeaveText + Chr$(0) + txtSayText1 + Chr$(0) + txtSayText2 + Chr$(0) + txtSayText3 + Chr$(0) + txtSayText4 + Chr$(0) + txtSayText5
    SendSocket St
    Me.Hide
End Sub
Private Sub btnUpdate_Click()
    Dim A As Long, b As Long, C As Long
    b = Int(Val(txtGiveValue))
    C = Int(Val(txtTakeValue))
    If b >= 0 And C >= 0 Then
        A = lstSaleItems.ListIndex
        With SaleItem(A)
            .GiveObject = cmbGiveObject.ListIndex
            .GiveValue = b
            .TakeObject = cmbTakeObject.ListIndex
            .TakeValue = C
        End With
        UpdateSaleItem A
    Else
        MsgBox "Invalid Give or Take Values!", vbOKOnly + vbExclamation, TitleString
    End If
End Sub

Private Sub Form_Load()
    Dim A As Long
    For A = 0 To 9
        lstSaleItems.AddItem CStr(A) + ":"
    Next A
    cmbGiveObject.AddItem "<nothing>"
    cmbTakeObject.AddItem "<nothing>"
    For A = 1 To MAXITEMS
        cmbGiveObject.AddItem CStr(A) + ": " + Object(A).Name
        cmbTakeObject.AddItem CStr(A) + ": " + Object(A).Name
    Next A
    lstSaleItems.ListIndex = 0
    frmNPC_Loaded = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmNPC_Loaded = False
End Sub


Private Sub lstSaleItems_Click()
    With SaleItem(lstSaleItems.ListIndex)
        cmbGiveObject.ListIndex = .GiveObject
        txtGiveValue = .GiveValue
        cmbTakeObject.ListIndex = .TakeObject
        txtTakeValue = .TakeValue
    End With
End Sub


Private Sub sclDirection_Change()
    Select Case sclDirection
        Case 0: lblDirection = "Up"
        Case 1: lblDirection = "Down"
        Case 2: lblDirection = "Left"
        Case 3: lblDirection = "Right"
    End Select
End Sub

Private Sub sclDirection_Scroll()
    sclDirection_Change
End Sub

Private Sub sclPortrait_Change()
    If sclPortrait > 0 And Exists(PORTRAIT_DIR & sclPortrait & ".rsc") Then
        On Error Resume Next
        Set picPortrait.Picture = LoadPicture(PORTRAIT_DIR & sclPortrait & ".rsc")
    Else
        Set picPortrait.Picture = Nothing
        picPortrait.Cls
    End If
    picPortrait.Refresh
    lblPortrait = sclPortrait
End Sub

Private Sub sclSprite_Change()
    lblSprite = sclSprite
    Dim R1 As RECT, r2 As RECT
    Dim A As Long
    R1.Top = 0: R1.Bottom = 32: R1.Left = 0: R1.Right = 32
    A = sclSprite - 1
    If A >= 0 And A < 255 Then
        r2.Top = (A \ 16) * 32: r2.Left = (A Mod 16) * 32: r2.Right = r2.Left + 32: r2.Bottom = r2.Top + 32
        sfcSprites.Surface.BltToDC picSprite.hdc, r2, R1
    Else
        picSprite.Cls
    End If
    picSprite.Refresh
End Sub

Private Sub sclSprite_Scroll()
    sclSprite_Change
End Sub
