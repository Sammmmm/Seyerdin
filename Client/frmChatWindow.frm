VERSION 5.00
Begin VB.Form frmChatWindow 
   BackColor       =   &H00000000&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   4395
   ClientLeft      =   120
   ClientTop       =   690
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   566
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picChat2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawMode        =   16  'Merge Pen
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2325
      Left            =   0
      ScaleHeight     =   155
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   785
      TabIndex        =   0
      Top             =   120
      Width           =   11775
   End
   Begin VB.Menu cOptions 
      Caption         =   "Chat Options"
      Begin VB.Menu Broadcasts 
         Caption         =   "Broadcasts"
         Checked         =   -1  'True
      End
      Begin VB.Menu Says 
         Caption         =   "Says"
         Checked         =   -1  'True
      End
      Begin VB.Menu Tells 
         Caption         =   "Tells"
         Checked         =   -1  'True
      End
      Begin VB.Menu Yells 
         Caption         =   "Yells"
         Checked         =   -1  'True
      End
      Begin VB.Menu Emotes 
         Caption         =   "Emotes"
         Checked         =   -1  'True
      End
      Begin VB.Menu GuildChats 
         Caption         =   "Guild Chats"
         Checked         =   -1  'True
      End
      Begin VB.Menu partychats 
         Caption         =   "Party Chats"
         Checked         =   -1  'True
      End
      Begin VB.Menu SystemMessages 
         Caption         =   "System Messages"
      End
   End
End
Attribute VB_Name = "frmChatWindow"
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

Private Sub Broadcasts_Click()
    Broadcasts.Checked = Not Broadcasts.Checked
    DrawChat2
End Sub

Private Sub Emotes_Click()
    Emotes.Checked = Not Emotes.Checked
    DrawChat2
End Sub

Private Sub Form_Resize()
    picChat2.Width = Me.ScaleWidth
    picChat2.Height = Me.ScaleHeight
    picChat2.Left = 0
    picChat2.Top = 0
    DrawChat2
End Sub

Private Sub picchat2_keydown(KeyCode As Integer, Shift As Integer)
frmMain.form_keydown KeyCode, Shift
End Sub
Private Sub picchat2_KeyPress(KeyAscii As Integer)
frmMain.Form_KeyPress KeyAscii
End Sub


Sub DrawChatLine2(ByVal St As String, ByRef A As Long, ByRef b As Long, ByRef textHeight As Byte)
    Dim C As Long, D As Long
                For C = 1 To Len(St)
                    If picChat2.TextWidth(Mid$(St, 1, C)) >= (picChat2.ScaleWidth - picChat2.CurrentX) Then
                      D = C
                      While (D > 1)
                        If Mid$(St, D, 1) = " " Then
                            C = D + 1
                            D = 1
                        End If
                        D = D - 1
                      Wend
                      DrawChatLine2 Mid$(St, C), A, b, textHeight
                      St = Mid$(St, 1, C - 1)
                      Exit For
                    End If
                Next C
                
                C = picChat2.ScaleHeight / textHeight
                C = C * textHeight
                C = picChat2.ScaleHeight - C
                
                TextOut picChat2.hdc, 0, picChat2.Top + ((b - 1) * textHeight) + C, St, Len(St)
                b = b - 1
                A = A - 1
                    

End Sub

Public Sub DrawChat2()
    Dim A As Long, b As Long, textHeight As Byte, curLine As Long, C As Long, D As Long
    textHeight = picChat2.textHeight("A")
    picChat2.FontSize = FontSize

    picChat2.Cls
    

    
    
    With picChat2
        b = picChat2.ScaleHeight / textHeight
        D = b
        A = Chat.ChatIndex - ChatScrollBack
        While (A > Chat.ChatIndex - ChatScrollBack - D + C And b <> 0)
            If A < 0 Then
                curLine = A + 1000
            Else
                curLine = IIf(A > 1000, A - 1000, A)
            End If
            If Not Chat.AllChat(curLine).used Then
                A = Chat.ChatIndex - ChatScrollBack - D
                b = 0
            Else
                If Chat.AllChat(curLine).Channel = 0 And Not Broadcasts.Checked Then
                    C = C - 1
                    A = A - 1
                ElseIf Chat.AllChat(curLine).Channel = 1 And Not Says.Checked Then
                    C = C - 1
                    A = A - 1
                ElseIf Chat.AllChat(curLine).Channel = 2 And Not Tells.Checked Then
                    C = C - 1
                    A = A - 1
                ElseIf Chat.AllChat(curLine).Channel = 3 And Not Emotes.Checked Then
                    C = C - 1
                    A = A - 1
                ElseIf Chat.AllChat(curLine).Channel = 4 And Not Yells.Checked Then
                    C = C - 1
                    A = A - 1
                ElseIf Chat.AllChat(curLine).Channel = 5 And Not GuildChats.Checked Then
                    C = C - 1
                    A = A - 1
                ElseIf Chat.AllChat(curLine).Channel = 6 And Not partychats.Checked Then
                    C = C - 1
                    A = A - 1
                ElseIf Chat.AllChat(curLine).Channel = 15 And Not SystemMessages.Checked Then
                    C = C - 1
                    A = A - 1
                Else
                    .ForeColor = Chat.AllChat(curLine).Color
                    DrawChatLine2 Chat.AllChat(curLine).Text, A, b, textHeight
                End If
            End If
        Wend

    End With
End Sub

Private Sub GuildChats_Click()
    GuildChats.Checked = Not GuildChats.Checked
    DrawChat2
End Sub

Private Sub partychats_Click()
    partychats.Checked = Not partychats.Checked
    DrawChat2
End Sub

Private Sub Says_Click()
    Says.Checked = Not Says.Checked
    DrawChat2
End Sub

Private Sub SystemMessages_Click()
    SystemMessages.Checked = Not SystemMessages.Checked
    DrawChat2
End Sub

Private Sub Tells_Click()
    Tells.Checked = Not Tells.Checked
    DrawChat2
End Sub
