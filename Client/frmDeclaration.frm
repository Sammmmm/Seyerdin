VERSION 5.00
Begin VB.Form frmDeclaration 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seyerdin Online [Delcaration]"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   ControlBox      =   0   'False
   Icon            =   "frmDeclaration.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox cmbGuild 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Declaration Guild:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmDeclaration"
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
    TempVar2 = 0
    TempVar3 = 0
    Unload Me
End Sub

Private Sub btnOk_Click()
    TempVar3 = cmbGuild.ItemData(cmbGuild.ListIndex)
    Unload Me
End Sub

Private Sub Form_Load()
    Dim A As Long
    
    For A = 1 To 255
        If Guild(A).Name <> "" And A <> Character.Guild Then
            cmbGuild.AddItem Guild(A).Name
            cmbGuild.ItemData(cmbGuild.ListCount - 1) = A
        End If
    Next A
    
    If cmbGuild.ListCount > 0 Then
        cmbGuild.ListIndex = 0
    Else
        btnOk.Enabled = False
    End If
    Set Me.Icon = frmMenu.Icon
End Sub
