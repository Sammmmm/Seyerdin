VERSION 5.00
Begin VB.Form frmHallAccess 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Choose Door Access"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1665
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   1665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton rdRank 
      Caption         =   "Guildmaster"
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.OptionButton rdRank 
      Caption         =   "Officer"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.OptionButton rdRank 
      Caption         =   "Member"
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.OptionButton rdRank 
      Caption         =   "Initiate"
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "frmHallAccess"
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
Public x As Long
Public y As Long
Public map As Long
Public default
Private Sub btnOk_Click()
    If rdRank(0).Value = True And default <> 0 Then SendSocket Chr$(85) + DoubleChar(CInt(map)) + Chr$(x) + Chr$(y) + Chr$(0)
    If rdRank(1).Value = True And default <> 1 Then SendSocket Chr$(85) + DoubleChar(CInt(map)) + Chr$(x) + Chr$(y) + Chr$(1)
    If rdRank(2).Value = True And default <> 2 Then SendSocket Chr$(85) + DoubleChar(CInt(map)) + Chr$(x) + Chr$(y) + Chr$(2)
    If rdRank(3).Value = True And default <> 3 Then SendSocket Chr$(85) + DoubleChar(CInt(map)) + Chr$(x) + Chr$(y) + Chr$(3)
    Unload Me
End Sub
