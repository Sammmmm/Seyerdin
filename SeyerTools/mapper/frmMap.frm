VERSION 5.00
Begin VB.Form frmMap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5895
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   424
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOk 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   5925
      Width           =   1215
   End
   Begin VB.PictureBox picMap 
      BackColor       =   &H00000000&
      Height          =   5820
      Left            =   30
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   384
      TabIndex        =   0
      Top             =   30
      Width           =   5820
   End
   Begin VB.Label Label6 
      Caption         =   "=Bad exit"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "=No exit"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      Height          =   135
      Left            =   1200
      TabIndex        =   5
      Top             =   5880
      Width           =   135
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   120
      TabIndex        =   4
      Top             =   6120
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "=Valid exit"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   120
      TabIndex        =   2
      Top             =   5880
      Width           =   135
   End
   Begin VB.Line lnExit 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   3
      X1              =   390
      X2              =   390
      Y1              =   0
      Y2              =   386
   End
   Begin VB.Line lnExit 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   386
   End
   Begin VB.Line lnExit 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   1
      X1              =   0
      X2              =   388
      Y1              =   390
      Y2              =   390
   End
   Begin VB.Line lnExit 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   2
      X2              =   390
      Y1              =   1
      Y2              =   1
   End
End
Attribute VB_Name = "frmMap"
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

Private Sub btnOk_Click()
    Unload Me
End Sub

Private Sub Form_Resize()
    btnOk.Top = Me.ScaleHeight - 29
    btnOk.Left = (Me.ScaleWidth - 81) / 2
    picMap.Width = Me.ScaleWidth - 7
    picMap.Height = Me.ScaleHeight - 40
End Sub

Private Sub picMap_KeyUp(KeyCode As Integer, Shift As Integer)
Dim A As Boolean
    A = False
    Select Case KeyCode
        Case 37 'left
            If Map(MapNum).ExitLeft > 0 Then
                If Map(Map(MapNum).ExitLeft).ExitRight > 0 Then
                    MapNum = Map(MapNum).ExitLeft
                    A = True
                End If
            End If
        Case 38 'up
            If Map(MapNum).ExitUp > 0 Then
                If Map(Map(MapNum).ExitUp).ExitDown > 0 Then
                    MapNum = Map(MapNum).ExitUp
                    A = True
                End If
            End If
        Case 39 'right
            If Map(MapNum).ExitRight > 0 Then
                If Map(Map(MapNum).ExitRight).ExitLeft > 0 Then
                    MapNum = Map(MapNum).ExitRight
                    A = True
                End If
            End If
        Case 40 'down
            If Map(MapNum).ExitDown > 0 Then
                If Map(Map(MapNum).ExitDown).ExitUp > 0 Then
                    MapNum = Map(MapNum).ExitDown
                    A = True
                End If
            End If
    End Select
    If blnDone = True And A = True Then
        Load frmWait
        With frmWait
            .lblStatus = "Drawing Map ..."
            .Show
            .Refresh
        End With
        DrawMap
        Load frmMap
        With frmMap
            BitBlt .picMap.hdc, 0, 0, 384, 384, hdcBuffer, 0, 0, SRCCOPY
            If Map(MapNum).ExitUp = 0 Then
                lnExit(0).BorderColor = &HFF&
            Else
                If Map(Map(MapNum).ExitUp).ExitDown > 0 Then
                    lnExit(0).BorderColor = &HFFFFFF
                Else
                    lnExit(0).BorderColor = &HFFFF&
                End If
            End If
            If Map(MapNum).ExitDown = 0 Then
                lnExit(1).BorderColor = &HFF&
            Else
                If Map(Map(MapNum).ExitDown).ExitUp > 0 Then
                    lnExit(1).BorderColor = &HFFFFFF
                Else
                    lnExit(1).BorderColor = &HFFFF&
                End If
            End If
            If Map(MapNum).ExitLeft = 0 Then
                lnExit(2).BorderColor = &HFF&
            Else
                If Map(Map(MapNum).ExitLeft).ExitRight > 0 Then
                    lnExit(2).BorderColor = &HFFFFFF
                Else
                    lnExit(2).BorderColor = &HFFFF&
                End If
            End If
            If Map(MapNum).ExitRight = 0 Then
                lnExit(3).BorderColor = &HFF&
            Else
                If Map(Map(MapNum).ExitRight).ExitLeft > 0 Then
                    lnExit(3).BorderColor = &HFFFFFF
                Else
                    lnExit(3).BorderColor = &HFFFF&
                End If
            End If
            .Caption = "The Odyssey Online Classic Mapper Utility [Map #" + CStr(MapNum) + ": " + Map(MapNum).Name + "]"
            .Show
        End With
        Unload frmWait
    End If
End Sub

Private Sub picMap_Paint()
    BitBlt picMap.hdc, 0, 0, 384, 384, hdcBuffer, 0, 0, SRCCOPY
End Sub


