VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seyerdin Online Mapper Utility [Login]"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLowQuality 
      Caption         =   "Low Quality"
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
      Left            =   1560
      TabIndex        =   13
      Top             =   2400
      Width           =   1575
   End
   Begin VB.ComboBox cmbSize 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1920
      Width           =   2535
   End
   Begin VB.HScrollBar sclStartMap 
      Height          =   255
      LargeChange     =   25
      Left            =   1560
      Max             =   5000
      Min             =   1
      TabIndex        =   8
      Top             =   1440
      Value           =   1
      Width           =   2535
   End
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Login"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CheckBox chkSavePassword 
      Caption         =   "Remember Password"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Quality:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Map Size:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblStartMap 
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
      Left            =   4200
      TabIndex        =   9
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "Start Map:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "User:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
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
    blnEnd = True
End Sub

Private Sub btnOk_Click()
    Dim St As String
    
    If chkLowQuality = 1 Then
        blnLowQuality = True
    Else
        blnLowQuality = False
    End If
    
    Select Case cmbSize.ListIndex
        Case 0 '1/64
            MapWidth = 4
            MapHeight = 4
        Case 1 '1/128
            MapWidth = 2
            MapHeight = 2
        Case 2 '1/256
            MapWidth = 1
            MapHeight = 1
        Case 3 '1/48
            MapWidth = 8
            MapHeight = 8
        Case 4 '1/24
            MapWidth = 16
            MapHeight = 16
        Case 5 '1/16
            MapWidth = 24
            MapHeight = 24
        Case 6 '1/12
            MapWidth = 32
            MapHeight = 32
        Case 7 '1/8
            MapWidth = 48
            MapHeight = 48
        Case 8 '1/6
            MapWidth = 64
            MapHeight = 64
        Case 9 '1/4
            MapWidth = 96
            MapHeight = 96
        Case 10 '1/3
            MapWidth = 128
            MapHeight = 128
        Case 11 '1/2
            MapWidth = 192
            MapHeight = 192
        Case 12 '1
            MapWidth = 384
            MapHeight = 384
    End Select
    
    User = txtUser
    Pass = txtPass
    MapNum = sclStartMap
    blnDone = False
    
    WriteString "Login", "User", User
    If chkSavePassword = 1 Then
        WriteString "Login", "Password", Pass
    Else
        DelItem "Login", "Password"
    End If
    WriteString "Login", "SavePassword", chkSavePassword
    
    frmWait.Show
    frmWait.lblStatus = "Connecting ..."
    frmWait.Refresh
    
    ClientSocket = ConnectSock(ServerIP, 3018, St, gHW, True)
    Me.Hide
End Sub
Sub DelItem(lpAppName As String, lpKeyName As String)
    WritePrivateProfileString lpAppName, lpKeyName, 0&, App.Path + "\odyssey.ini"
End Sub

Private Sub Form_Load()
    frmLogin_Loaded = True
    
    cmbSize.AddItem "1/256 actual size (1 pixil wide)"
    cmbSize.AddItem "1/128 actual size (2)"
    cmbSize.AddItem "1/64 actual size (4)"
    cmbSize.AddItem "1/48 actual size (8)"
    cmbSize.AddItem "1/24 actual size (16)"
    cmbSize.AddItem "1/16 actual size (24)"
    cmbSize.AddItem "1/12 actual size (32)"
    cmbSize.AddItem "1/8 actual size (48)"
    cmbSize.AddItem "1/6 actual size (64)"
    cmbSize.AddItem "1/4 actual size (96)"
    cmbSize.AddItem "1/3 actual size (128)"
    cmbSize.AddItem "1/2 actual size (192)"
    cmbSize.AddItem "actual size (384)"
    
    
    cmbSize.ListIndex = 7
    
    If ReadInt("Login", "SavePassword") > 0 Then
        chkSavePassword = 1
        txtPass = ReadString("Login", "Password")
    Else
        chkSavePassword = 0
    End If
    
    txtUser = ReadString("Login", "User")
    
    Me.Show
    
    If txtUser = "" Then
        txtUser.SetFocus
    Else
        txtPass.SetFocus
    End If
    
    gHW = Me.hWnd
End Sub
Private Sub Form_Unload(Cancel As Integer)
    frmLogin_Loaded = False
End Sub


Private Sub sclStartMap_Change()
    lblStartMap = sclStartMap
End Sub

Private Sub sclStartMap_Scroll()
    sclStartMap_Change
End Sub


Private Sub txtPass_Change()
    If txtUser <> "" And txtPass <> "" Then
        If btnOk.Enabled = False Then
            btnOk.Enabled = True
        End If
    Else
        If btnOk.Enabled = True Then
            btnOk.Enabled = False
        End If
    End If
End Sub

Private Sub txtUser_Change()
    If txtUser <> "" And txtPass <> "" Then
        If btnOk.Enabled = False Then
            btnOk.Enabled = True
        End If
    Else
        If btnOk.Enabled = True Then
            btnOk.Enabled = False
        End If
    End If
End Sub


