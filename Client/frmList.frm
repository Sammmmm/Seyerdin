VERSION 5.00
Begin VB.Form frmList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seyerdin Online"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3225
   ControlBox      =   0   'False
   Icon            =   "frmList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   215
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstScripts 
      Height          =   4545
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtContaining 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   1455
   End
   Begin VB.ListBox lstLights 
      Height          =   4545
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox lstPrefix 
      Height          =   4545
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox lstObjects 
      Height          =   4545
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox lstMonsters 
      Height          =   4545
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox lstHalls 
      Height          =   4545
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox lstBans 
      Height          =   4545
      ItemData        =   "frmList.frx":000C
      Left            =   120
      List            =   "frmList.frx":000E
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox lstNPCs 
      Height          =   4545
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Show All Containing:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmList"
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
    'Me.Hide
    Unload Me
End Sub

Private Sub btnOk_Click()
    Dim A As Long
    If lstObjects.Visible Then
        SendSocket Chr$(19) + DoubleChar(Mid(lstObjects.List(lstObjects.ListIndex), 1, InStr(1, lstObjects.List(lstObjects.ListIndex), ":") - 1))
    ElseIf lstMonsters.Visible Then
        SendSocket Chr$(20) + DoubleChar(Mid(lstMonsters.List(lstMonsters.ListIndex), 1, InStr(1, lstMonsters.List(lstMonsters.ListIndex), ":") - 1))
    ElseIf lstNPCs.Visible Then
        SendSocket Chr$(50) + Chr$(Mid(lstNPCs.List(lstNPCs.ListIndex), 1, InStr(1, lstNPCs.List(lstNPCs.ListIndex), ":") - 1))
    ElseIf lstHalls.Visible = True Then
        SendSocket Chr$(48) + Chr$(Mid(lstHalls.List(lstHalls.ListIndex), 1, InStr(1, lstHalls.List(lstHalls.ListIndex), ":") - 1))
    ElseIf lstBans.Visible Then
        SendSocket Chr$(57) + Chr$(lstBans.ItemData(lstBans.ListIndex))
    ElseIf lstPrefix.Visible Then
        SendSocket Chr$(70) + Chr$(Mid(lstPrefix.List(lstPrefix.ListIndex), 1, InStr(1, lstPrefix.List(lstPrefix.ListIndex), ":") - 1))
    ElseIf lstLights.Visible Then
        Load frmLight
        A = lstLights.ListIndex + 1
        frmLight.lblNumber = A
        With Lights(A)
            frmLight.txtName = .Name
            frmLight.sclIntensity = .Intensity
            frmLight.sclRadius = .Radius
            frmLight.sclFlicker = .MaxFlicker
            frmLight.sclRate = .FlickerRate
            frmLight.sclRed = .Red
            frmLight.sclGreen = .Green
            frmLight.sclBlue = .Blue
        End With
        frmLight.Show
    ElseIf lstScripts.Visible Then
        SendSocket Chr$(59) + lstScripts.List(lstScripts.ListIndex)
    End If
    'txtContaining.Text = ""
    'Unload Me
End Sub

Private Sub Form_Load()
    Dim A As Long
    For A = 1 To MAXITEMS
        lstMonsters.AddItem CStr(A) + ": " + Monster(A).Name
    Next A
    For A = 1 To 255
        lstHalls.AddItem CStr(A) + ": " + Hall(A).Name
        lstNPCs.AddItem CStr(A) + ": " + NPC(A).Name
        lstPrefix.AddItem CStr(A) + ": " + Prefix(A).Name & " (" & ModString(Prefix(A).ModType) & ")"
    Next A
    For A = 1 To 255
        lstLights.AddItem CStr(A) + ": " + Lights(A).Name
    Next A
    For A = 1 To MAXITEMS
        lstObjects.AddItem CStr(A) + ": " + Object(A).Name
    Next A
    frmList_Loaded = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmList_Loaded = False
End Sub

Private Sub lstBans_Click()
    btnOk.Enabled = True
End Sub

Private Sub lstBans_DblClick()
    btnOk_Click
End Sub

Private Sub lstHalls_Click()
    btnOk.Enabled = True
End Sub

Private Sub lstHalls_DblClick()
    btnOk_Click
End Sub

Private Sub lstLights_Click()
    btnOk.Enabled = True
End Sub

Private Sub lstLights_DblClick()
    btnOk_Click
End Sub
Private Sub lstMonsters_Click()
    btnOk.Enabled = True
End Sub

Private Sub lstMonsters_DblClick()
    btnOk_Click
End Sub

Private Sub lstNPCs_Click()
    btnOk.Enabled = True
End Sub

Private Sub lstNPCs_DblClick()
    btnOk_Click
End Sub

Private Sub lstObjects_Click()
    btnOk.Enabled = True
End Sub

Private Sub lstObjects_DblClick()
    btnOk_Click
End Sub

Private Sub lstPrefix_Click()
    btnOk.Enabled = True
End Sub

Private Sub lstPrefix_DblClick()
    btnOk_Click
End Sub

Private Sub lstScripts_Click()
    btnOk.Enabled = True
End Sub

Private Sub lstScripts_DblClick()
    btnOk_Click
End Sub

Private Sub txtContaining_Change()
    Dim A As Long
    lstMonsters.Clear
    lstObjects.Clear
    lstHalls.Clear
    lstNPCs.Clear
    lstPrefix.Clear
    lstScripts.Clear
    For A = 1 To MAXITEMS
        If InStr(1, Str(A), txtContaining.Text, vbTextCompare) Or InStr(1, Monster(A).Name, txtContaining.Text, vbTextCompare) Or txtContaining.Text = "" Then
            lstMonsters.AddItem CStr(A) + ": " + Monster(A).Name
        End If
    Next A
    For A = 1 To 255

        If InStr(1, Str(A), txtContaining.Text, vbTextCompare) Or InStr(1, Hall(A).Name, txtContaining.Text, vbTextCompare) Or txtContaining.Text = "" Then
            lstHalls.AddItem CStr(A) + ": " + Hall(A).Name
        End If
        If InStr(1, Str(A), txtContaining.Text, vbTextCompare) Or InStr(1, NPC(A).Name, txtContaining.Text, vbTextCompare) Or txtContaining.Text = "" Then
            lstNPCs.AddItem CStr(A) + ": " + NPC(A).Name
        End If
        If InStr(1, Str(A), txtContaining.Text, vbTextCompare) Or InStr(1, Prefix(A).Name, txtContaining.Text, vbTextCompare) Or txtContaining.Text = "" Then
            lstPrefix.AddItem CStr(A) + ": " + Prefix(A).Name
        End If
    Next A
    For A = 1 To MAXITEMS
        If InStr(1, Str(A), txtContaining.Text, vbTextCompare) Or InStr(1, Object(A).Name, txtContaining.Text, vbTextCompare) Or txtContaining.Text = "" Then
            lstObjects.AddItem CStr(A) + ": " + Object(A).Name
        End If
    Next A
    For A = 0 To UBound(scripts)
        If InStr(1, scripts(A), txtContaining.Text, vbTextCompare) Then
            lstScripts.AddItem (scripts(A))
        End If
        If scripts(A) = "" Then Exit For
    Next A
End Sub
