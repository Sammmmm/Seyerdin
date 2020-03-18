VERSION 5.00
Begin VB.Form frmPrefix 
   Caption         =   "Seyerdin Online [Prefix / Suffix Editor]"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4335
   Icon            =   "frmPrefix.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar sclRarity 
      Height          =   255
      Left            =   1200
      Max             =   10
      Min             =   1
      TabIndex        =   30
      Top             =   2160
      Value           =   1
      Width           =   2295
   End
   Begin VB.Frame frmLight 
      Caption         =   "Light Info"
      Height          =   1095
      Left            =   240
      TabIndex        =   23
      Top             =   4320
      Width           =   3855
      Begin VB.HScrollBar sclRadius 
         Height          =   255
         Left            =   840
         Max             =   255
         Min             =   1
         TabIndex        =   27
         Top             =   600
         Value           =   1
         Width           =   2055
      End
      Begin VB.HScrollBar sclIntensity 
         Height          =   255
         Left            =   840
         Max             =   127
         Min             =   -127
         TabIndex        =   25
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblRadius 
         Caption         =   "1"
         Height          =   255
         Left            =   3000
         TabIndex        =   29
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblIntensity 
         Caption         =   "0"
         Height          =   255
         Left            =   3000
         TabIndex        =   28
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "Radius"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Intensity:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.CheckBox chkType 
      Caption         =   "Suffix"
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cmbWeakness2 
      Height          =   315
      Left            =   1200
      TabIndex        =   21
      Top             =   3840
      Width           =   3015
   End
   Begin VB.ComboBox cmbWeakness1 
      Height          =   315
      Left            =   1200
      TabIndex        =   20
      Top             =   3480
      Width           =   3015
   End
   Begin VB.ComboBox cmbStrength2 
      Height          =   315
      Left            =   1200
      TabIndex        =   19
      Top             =   3120
      Width           =   3015
   End
   Begin VB.ComboBox cmbStrength1 
      Height          =   315
      Left            =   1200
      TabIndex        =   18
      Top             =   2760
      Width           =   3015
   End
   Begin VB.HScrollBar sclMax 
      Height          =   255
      Left            =   1200
      Max             =   255
      TabIndex        =   16
      Top             =   1800
      Value           =   1
      Width           =   2295
   End
   Begin VB.ComboBox cmbModify 
      Height          =   315
      ItemData        =   "frmPrefix.frx":000C
      Left            =   1200
      List            =   "frmPrefix.frx":000E
      TabIndex        =   10
      Top             =   1080
      Width           =   2895
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
      TabIndex        =   3
      Top             =   720
      Width           =   2895
   End
   Begin VB.HScrollBar sclMin 
      Height          =   255
      Left            =   1200
      Max             =   255
      TabIndex        =   2
      Top             =   1440
      Value           =   1
      Width           =   2295
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Rarity:"
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
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblRarity 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   31
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblMax 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   17
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Max:"
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
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Strength:"
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
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Strength:"
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
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Weakness:"
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
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Weakness:"
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
      TabIndex        =   11
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Modify:"
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
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
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
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
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
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Minimum:"
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
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblMin 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   1440
      Width           =   495
   End
End
Attribute VB_Name = "frmPrefix"
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
Unload Me
End Sub

Private Sub btnOk_Click()
    Dim Flags As Byte, St As String, A As Byte
    If chkType.Value = 1 Then
        SetBit Flags, 1
    Else
        ClearBit Flags, 1
    End If
    St = Chr$(lblNumber) + Chr$(cmbModify.ListIndex) + Chr$(sclMin) + Chr$(sclMax) + Chr$(Flags) + Chr$(cmbStrength1.ListIndex) + Chr$(cmbStrength2.ListIndex) + Chr$(cmbWeakness1.ListIndex) + Chr$(cmbWeakness2.ListIndex)
    A = Abs(sclIntensity)
    If sclIntensity < 0 Then
        SetBit A, 7
        St = St + Chr$(A) + Chr$(sclRadius)
    ElseIf sclIntensity > 0 Then
        ClearBit A, 7
        St = St + Chr$(A) + Chr$(sclRadius)
    Else
        St = St + Chr$(0) + Chr$(0)
    End If
    St = St + Chr$(sclRarity)
    St = St + Trim$(txtName.Text)
    SendSocket Chr$(71) & St
    Unload Me
End Sub

Private Sub Form_Load()
Dim A As Long

cmbStrength1.AddItem "<None>"
cmbStrength2.AddItem "<None>"
cmbWeakness1.AddItem "<None>"
cmbWeakness2.AddItem "<None>"

With cmbModify
    .AddItem "<None>"
    For A = 1 To maxmod
        .AddItem ModString(A)
    Next A
End With

    For A = 1 To 255
        cmbStrength1.AddItem CStr(A) + ": " + Prefix(A).Name
        cmbStrength2.AddItem CStr(A) + ": " + Prefix(A).Name
        cmbWeakness1.AddItem CStr(A) + ": " + Prefix(A).Name
        cmbWeakness2.AddItem CStr(A) + ": " + Prefix(A).Name
    Next A
End Sub

Private Sub sclIntensity_Change()
    lblIntensity = sclIntensity * 2
End Sub

Private Sub sclMax_Change()
If sclMax < sclMin Then sclMax = sclMin
lblMax = sclMax
End Sub

Private Sub sclMax_Scroll()
lblMax = sclMax
End Sub

Private Sub sclMin_Change()
If sclMax < sclMin Then sclMax = sclMin
lblMin = sclMin
lblMax = sclMax
End Sub

Private Sub sclMin_Scroll()
lblMax = sclMax
End Sub

Private Sub sclRadius_Change()
    lblRadius = sclRadius
End Sub

Private Sub sclRarity_Change()
    lblRarity = sclRarity
End Sub

Private Sub sclRarity_Scroll()
    sclRarity_Change
End Sub
