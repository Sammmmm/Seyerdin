VERSION 5.00
Begin VB.Form frmLight 
   Caption         =   "Seyerdin Online [Light Editor]"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   ScaleHeight     =   4320
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   495
      Left            =   3120
      TabIndex        =   27
      Top             =   3720
      Width           =   1335
   End
   Begin VB.HScrollBar sclRate 
      Height          =   255
      Left            =   960
      Max             =   255
      TabIndex        =   24
      Top             =   2040
      Width           =   2895
   End
   Begin VB.HScrollBar sclFlicker 
      Height          =   255
      Left            =   960
      Max             =   255
      TabIndex        =   17
      Top             =   1680
      Width           =   2895
   End
   Begin VB.HScrollBar sclRadius 
      Height          =   255
      Left            =   960
      Max             =   255
      TabIndex        =   16
      Top             =   1320
      Width           =   2895
   End
   Begin VB.HScrollBar sclIntensity 
      Height          =   255
      Left            =   960
      Max             =   255
      TabIndex        =   15
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1680
      TabIndex        =   13
      Top             =   3720
      Width           =   1335
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
      Left            =   1080
      MaxLength       =   15
      TabIndex        =   3
      Top             =   480
      Width           =   3375
   End
   Begin VB.HScrollBar sclRed 
      Height          =   255
      Left            =   960
      Max             =   255
      TabIndex        =   2
      Top             =   2640
      Width           =   2895
   End
   Begin VB.HScrollBar sclGreen 
      Height          =   255
      Left            =   960
      Max             =   255
      TabIndex        =   1
      Top             =   3000
      Width           =   2895
   End
   Begin VB.HScrollBar sclBlue 
      Height          =   255
      Left            =   960
      Max             =   255
      TabIndex        =   0
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label lblRate 
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
      Height          =   255
      Left            =   3960
      TabIndex        =   26
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "Rate:"
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
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lblIntensity 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   3960
      TabIndex        =   23
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblFlicker 
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
      Height          =   255
      Left            =   3960
      TabIndex        =   22
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblRadius 
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
      TabIndex        =   21
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "Flicker:"
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
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Radius:"
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
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Intensity:"
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
      Top             =   960
      Width           =   855
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
      TabIndex        =   12
      Top             =   120
      Width           =   855
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
      Left            =   1080
      TabIndex        =   11
      Top             =   120
      Width           =   2655
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
      TabIndex        =   10
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Red:"
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
      TabIndex        =   9
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Green:"
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
      TabIndex        =   8
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Blue:"
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
      TabIndex        =   7
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblGreen 
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
      TabIndex        =   6
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label lblBlue 
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
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lblRed 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   2640
      Width           =   495
   End
End
Attribute VB_Name = "frmLight"
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

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    SendSocket Chr$(79) + Chr$(lblNumber) + Chr$(sclRed) + Chr$(sclGreen) + Chr$(sclBlue) + Chr$(sclIntensity) + Chr$(sclRadius) + Chr$(sclFlicker) + Chr$(sclRate) + txtName
    Unload Me
End Sub

Private Sub cmdApply_Click()
    SendSocket Chr$(79) + Chr$(lblNumber) + Chr$(sclRed) + Chr$(sclGreen) + Chr$(sclBlue) + Chr$(sclIntensity) + Chr$(sclRadius) + Chr$(sclFlicker) + Chr$(sclRate) + txtName
End Sub

Private Sub Form_Load()
    frmLight_Loaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmLight_Loaded = False
End Sub

Private Sub sclBlue_Change()
    lblBlue = sclBlue
End Sub

Private Sub sclBlue_Scroll()
    sclBlue_Change
End Sub

Private Sub sclFlicker_Change()
    lblFlicker = sclFlicker
End Sub

Private Sub sclFlicker_Scroll()
    sclFlicker_Change
End Sub

Private Sub sclGreen_Change()
    lblGreen = sclGreen
End Sub

Private Sub sclGreen_Scroll()
    sclGreen_Change
End Sub

Private Sub sclIntensity_Change()
    lblIntensity = sclIntensity
End Sub

Private Sub sclIntensity_Scroll()
    sclIntensity_Change
End Sub

Private Sub sclRadius_Change()
    lblRadius = sclRadius
End Sub

Private Sub sclRadius_Scroll()
    sclRadius_Change
End Sub

Private Sub sclRate_Change()
    lblRate = sclRate
End Sub

Private Sub sclRate_Scroll()
    sclRate_Change
End Sub

Private Sub sclRed_Change()
    lblRed = sclRed
End Sub

Private Sub sclRed_Scroll()
    sclRed_Change
End Sub
