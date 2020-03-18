VERSION 5.00
Object = "{E910F8E1-8996-4EE9-90F1-3E7C64FA9829}#1.1#0"; "vbaListView6.ocx"
Begin VB.Form frmGuilds 
   BorderStyle     =   0  'None
   Caption         =   "Seyerdin Online [Guilds]"
   ClientHeight    =   7020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5100
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   Icon            =   "frmGuilds.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   462.076
   ScaleMode       =   0  'User
   ScaleWidth      =   334.104
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pbManage 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7020
      Left            =   0
      ScaleHeight     =   468
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   340
      TabIndex        =   3
      Top             =   0
      Width           =   5100
      Begin VB.CommandButton Command1 
         Caption         =   "buy"
         Height          =   375
         Left            =   4080
         TabIndex        =   31
         Top             =   5400
         Width           =   375
      End
      Begin VB.HScrollBar sclSymbol1 
         Height          =   135
         Left            =   1080
         Max             =   255
         TabIndex        =   30
         Top             =   5400
         Width           =   2895
      End
      Begin VB.PictureBox pbSymbol 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         DrawMode        =   9  'Not Mask Pen
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   720
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   29
         Top             =   5400
         Width           =   300
      End
      Begin VB.HScrollBar sclSymbol2 
         Height          =   135
         Left            =   1080
         Max             =   255
         TabIndex        =   27
         Top             =   5520
         Width           =   2895
      End
      Begin VB.HScrollBar sclSymbol3 
         Height          =   135
         Left            =   1080
         Max             =   255
         TabIndex        =   28
         Top             =   5640
         Width           =   2895
      End
      Begin VB.TextBox txtMOTD 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   1680
         Left            =   345
         MaxLength       =   400
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1110
         Width           =   4395
      End
      Begin VB.ListBox lstMembers2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   1830
         IntegralHeight  =   0   'False
         Left            =   360
         TabIndex        =   4
         Top             =   3540
         Width           =   4395
      End
      Begin VB.ListBox lstDeclarations2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2235
         IntegralHeight  =   0   'False
         Left            =   360
         TabIndex        =   11
         Top             =   3540
         Width           =   4395
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   15
         Top             =   2955
         Width           =   915
      End
   End
   Begin VB.PictureBox pbMyGuild 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7020
      Left            =   0
      ScaleHeight     =   468
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   340
      TabIndex        =   2
      Top             =   0
      Width           =   5100
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   3270
         HideSelection   =   0   'False
         Left            =   360
         MaxLength       =   2048
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2895
         Width           =   4395
      End
      Begin VB.Label lblSaveInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   6240
         Width           =   915
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Message of the Day:"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1005
         Width           =   2055
      End
      Begin VB.Label lblmotd 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000B&
         Height          =   1575
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   4620
      End
   End
   Begin VB.PictureBox pbInfo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7020
      Left            =   0
      ScaleHeight     =   468
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   340
      TabIndex        =   1
      Top             =   0
      Width           =   5100
      Begin VB.PictureBox pbSymbol2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   360
         ScaleHeight     =   300
         ScaleWidth      =   300
         TabIndex        =   32
         Top             =   1680
         Width           =   300
      End
      Begin vbalListViewLib6.vbalListViewCtl lstMembers 
         Height          =   2445
         Left            =   360
         TabIndex        =   18
         Top             =   3720
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   4313
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         BackColor       =   0
         View            =   1
         LabelEdit       =   0   'False
         FullRowSelect   =   -1  'True
         AutoArrange     =   0   'False
         Appearance      =   0
         BorderStyle     =   0
         HeaderButtons   =   0   'False
         HeaderTrackSelect=   0   'False
         HideSelection   =   0   'False
         InfoTips        =   0   'False
         ScaleMode       =   3
         TileBackgroundPicture=   0   'False
      End
      Begin VB.ListBox lstDeclarations 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2445
         IntegralHeight  =   0   'False
         Left            =   360
         TabIndex        =   12
         Top             =   3720
         Width           =   4395
      End
      Begin VB.Label lblFounded 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   26
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Founded:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   25
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblMembersTotal 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2100
         TabIndex        =   24
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Online/Total Members:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label lblAverageRenown 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   22
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Average Renown:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label lblRenown 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1155
         TabIndex        =   20
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Renown:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   19
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label lblHall 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lbl1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Hall:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   615
      End
   End
   Begin VB.PictureBox pbGuilds 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7020
      Left            =   0
      ScaleHeight     =   468
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   340
      TabIndex        =   0
      Top             =   0
      Width           =   5100
      Begin vbalListViewLib6.vbalListViewCtl lstGuilds 
         Height          =   5295
         Left            =   210
         TabIndex        =   17
         Top             =   990
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   9340
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   16777215
         BackColor       =   0
         View            =   1
         LabelEdit       =   0   'False
         FullRowSelect   =   -1  'True
         AutoArrange     =   0   'False
         Appearance      =   0
         BorderStyle     =   0
         HeaderButtons   =   0   'False
         HeaderTrackSelect=   0   'False
         HideSelection   =   0   'False
         InfoTips        =   0   'False
         ScaleMode       =   3
         TileBackgroundPicture=   0   'False
      End
   End
End
Attribute VB_Name = "frmGuilds"
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
Public ctab As Byte
Dim ctab2 As Byte
Public nextTab As Byte
Public promoteidx As Byte
Public meIndex As Byte
Dim nextGuild As Byte
Dim prevGuild As Byte
Dim Selected As Integer
Public lstEnabled As Boolean
Public playerGuild As Byte
Public GuildRank As Byte
Public changed As Boolean

Private Sub Command1_Click()
    With Guild(Character.Guild)
        If sclSymbol1.Value = 0 And sclSymbol2.Value = 0 And sclSymbol3.Value = 0 Then
            PrintChat "You cannot buy no guild symbol!", 7, Options.FontSize, 15
        Else
            If .Symbol1 <> sclSymbol1.Value Or .Symbol2 <> sclSymbol2.Value Or .Symbol3 <> sclSymbol3.Value Then
                SendSocket Chr$(88) + Chr$(sclSymbol1.Value) + Chr$(sclSymbol2.Value) + Chr$(sclSymbol3.Value)
            Else
                PrintChat "You already own that symbol!.", 7, Options.FontSize
            End If
        End If
    End With
End Sub
Private Sub Form_Load()

    lstEnabled = True
    Dim A As Long
    Dim item As cListItem, header As cColumn
    
    Set header = lstGuilds.Columns.Add(, , "Name", , 106)
    Set header = lstGuilds.Columns.Add(, , "Members", , TextWidth("Members") + 20)
    Set header = lstGuilds.Columns.Add(, , "Avg. Renown", , TextWidth("Avg. Renown") + 19)
    Set header = lstGuilds.Columns.Add(, , "", , 0)


    Set header = lstMembers.Columns.Add(, , "Name", , 77)
    Set header = lstMembers.Columns.Add(, , "Rank", , 77)
    Set header = lstMembers.Columns.Add(, , "", , 0)
    Set header = lstMembers.Columns.Add(, , "Renown", , 55)
    Set header = lstMembers.Columns.Add(, , "Joined", , 77)

    With lstGuilds.ListItems
        For A = 1 To 255
            If Guild(A).Name <> "" Then
                Set item = .Add(A, , Guild(A).Name)
                item.SubItems(1).Caption = Str(Guild(A).MembersOnline) + "/" + Str(Guild(A).members)
                item.SubItems(2).Caption = Str(Guild(A).AverageRenown)
                item.SubItems(3).Caption = A
                

                
            End If
        Next A
    End With
    
    Selected = -1
    frmGuilds_Loaded = True
    Set Me.Icon = frmMenu.Icon
    
    
    
    changed = True
        pbInfo.Visible = False
        pbGuilds.Visible = True
        pbManage.Visible = False
        pbMyGuild.Visible = False


        lstGuilds.Columns(1).SortOrder = eSortOrderDescending
        lstGuilds.Columns(1).SortType = eLVSortString
        lstGuilds.ListItems.SortItems
    
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'If Button = 1 And Me.WindowState = 0 Then
    '   Dim ReturnVal As Long
    '   ReleaseCapture
    '   ReturnVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    'End If

        If x >= 235 And y > 434 And x <= 327 And y <= 461 Then
            Unload Me
        End If
    If lstEnabled Then
        If Selected >= 0 Then

            If x >= 4 And y >= 36 And x <= 79 And y <= 58 Then
                If ctab = 3 Or ctab = 2 Then
                    nextGuild = prevGuild
                    Selected = prevGuild
                    changed = True
                Else
                    nextGuild = Selected
                End If
                nextTab = 0
                setGuildTab
            End If
            If x >= 89 And y >= 36 And x <= 89 + 75 And y <= 58 Then
                nextTab = 1
                If ctab = 3 Or ctab = 2 Then
                    nextGuild = prevGuild
                    Selected = prevGuild
                    changed = True
                Else
                    nextGuild = Selected
                End If
                If changed = True Then
                    SendSocket Chr$(39) + Chr$(Selected)
                    lstEnabled = False
                    changed = False
                Else
                    setGuildTab
                End If
            End If
        End If
        If Selected >= 0 Or playerGuild > 0 Then
            If x >= 172 And y >= 36 And x <= 172 + 75 And y <= 58 And playerGuild > 0 Then
                nextTab = 2
                If ctab <> 3 And ctab <> 2 Then prevGuild = nextGuild
                Selected = playerGuild
                nextGuild = playerGuild
                If changed = True Then
                    SendSocket Chr$(39) + Chr$(Selected)
                    lstEnabled = False
                    changed = False
                Else
                    setGuildTab
                End If
            End If

            If x >= 256 And y >= 36 And x <= 256 + 75 And y <= 58 And playerGuild > 0 And GuildRank >= 2 Then
                nextTab = 3
                If ctab <> 3 And ctab <> 2 Then prevGuild = nextGuild
                Selected = playerGuild
                nextGuild = playerGuild
                If changed = True Then
                    ctab2 = 1
                    SendSocket Chr$(39) + Chr$(Selected)
                    lstEnabled = False
                    changed = False
                Else
                    setGuildTab
                End If
            End If
        End If
        If ctab = 1 Then
            If x >= 13 And y >= 218 And x <= 13 + 77 And y <= 242 Then
                ctab2 = 1
                Set pbInfo.Picture = LoadPicture("Data\Graphics\Interface\NSGuildGUI2a.rsc")
                lstDeclarations.Visible = False
                lstMembers.Visible = True
            End If
            If x >= 98 And y >= 218 And x <= 98 + 77 And y <= 242 Then
                ctab2 = 2
                Set pbInfo.Picture = LoadPicture("Data\Graphics\Interface\NSGuildGUI2b.rsc")
                lstDeclarations.Visible = True
                lstMembers.Visible = False
            End If
        ElseIf ctab = 2 Then
            If x >= 160 And y >= 435 And x <= 236 And y <= 461 Then
                nextTab = 1
                Selected = playerGuild
                prevGuild = playerGuild
                nextGuild = Selected
                setGuildTab

            End If
        ElseIf ctab = 3 Then
            If x >= 255 And y >= 200 And x <= 307 And y <= 222 Then
                SendSocket Chr$(82) + txtMOTD
                SendSocket Chr$(39) + Chr$(Selected)
            End If
            If x >= 13 And y >= 205 And x <= 13 + 77 And y <= 230 Then
                ctab2 = 1
                Set pbManage.Picture = LoadPicture("Data\Graphics\Interface\NSGuildGUI4a.rsc")
                lstDeclarations2.Visible = False
                lstMembers2.Visible = True
                sclSymbol1.Visible = True
                sclSymbol2.Visible = True
                sclSymbol3.Visible = True
                pbSymbol.Visible = True
            End If
            If x >= 98 And y >= 205 And x <= 98 + 77 And y <= 230 Then
                ctab2 = 2
                Set pbManage.Picture = LoadPicture("Data\Graphics\Interface\NSGuildGUI4b.rsc")
                lstDeclarations2.Visible = True
                lstMembers2.Visible = False
                sclSymbol1.Visible = False
                sclSymbol2.Visible = False
                sclSymbol3.Visible = False
                pbSymbol.Visible = False
            End If
            
            'If Selected >= 0 Then
                If x >= 255 And y >= 389 And x <= 255 + 62 And y <= 389 + 28 And ctab2 = 1 Then
                    Dim St As String
                    If lstMembers2.ListIndex >= 0 Then
                        If lstMembers2.ListCount <= 3 Then
                            St = " Because there are only " + CStr(lstMembers2.ListCount) + " members in your guild, removing this member will cause your guild to be deleted.  Are you SURE you wish to continue?"
                        End If
                        If MsgBox("Are you sure you wish to kick " + Chr$(34) + lstMembers2.List(lstMembers2.ListIndex) + Chr$(34) + " out of the guild?" + St, vbYesNo + vbQuestion, TitleString) = vbYes Then
                            SendSocket Chr$(35) + Chr$(lstMembers2.ItemData(lstMembers2.ListIndex))
                            SendSocket Chr$(39) + Chr$(Character.Guild)
                        End If
                    End If
                End If
            
            
                If x >= 23 And y >= 389 And x <= 23 + 62 And y <= 389 + 28 And ctab2 = 1 Then
                    If lstMembers2.ListIndex >= 0 Then
                        promoteidx = lstMembers2.ListIndex
                        
                        If (Character.GuildRank = 3) Then
                            If InStr(lstMembers2.Text, "Officer") Then
                                If MsgBox("WARNING: Guilds can only have one guildmaster, promoting this officer will demote you to officer.", vbOKCancel, "Promote New Guildmaster?") = vbOK Then
                                    SendSocket Chr$(36) + Chr$(lstMembers2.ItemData(lstMembers2.ListIndex)) + Chr$(1)
                                    SendSocket Chr$(36) + Chr$(lstMembers2.ItemData(meIndex)) + Chr$(0)
                                    SendSocket Chr$(39) + Chr$(Character.Guild)
                                End If
                                Exit Sub
                            End If
                        End If
                        
                        
                        SendSocket Chr$(36) + Chr$(lstMembers2.ItemData(lstMembers2.ListIndex)) + Chr$(1)
                        SendSocket Chr$(39) + Chr$(Character.Guild)
                    End If
                End If
                If x >= 89 And y >= 389 And x <= 89 + 62 And y <= 389 + 28 And ctab2 = 1 Then
                    If lstMembers2.ListIndex >= 0 Then
                        promoteidx = lstMembers2.ListIndex
                        
                        If (Character.GuildRank = 3 And promoteidx = meIndex) Then
                            MsgBox "You cannot demote yourself as guildmaster, you must promote a replacement.", vbOKOnly, "Cannot demote guildmaster"
                            Exit Sub
                        End If
                        
                        SendSocket Chr$(36) + Chr$(lstMembers2.ItemData(lstMembers2.ListIndex)) + Chr$(0)
                        SendSocket Chr$(39) + Chr$(Character.Guild)
                    End If
                End If
            'End If
            If x >= 255 And y >= 389 And x <= 255 + 62 And y <= 389 + 28 And ctab2 = 2 Then
                If lstDeclarations2.ListIndex >= 0 Then
                    SendSocket Chr$(38) + Chr$(lstDeclarations2.ItemData(lstDeclarations2.ListIndex))
                    SendSocket Chr$(39) + Chr$(Character.Guild)
                End If
            End If
            If x >= 23 And y >= 389 And x <= 23 + 62 And y <= 389 + 28 And ctab2 = 2 Then
                frmDeclaration.Show 1
                If TempVar3 > 0 Then
                    SendSocket Chr$(37) + Chr$(TempVar3) + Chr$(1)
                    SendSocket Chr$(39) + Chr$(Character.Guild)
                End If
            End If
            If x >= 89 And y >= 389 And x <= 89 + 62 And y <= 389 + 28 And ctab2 = 2 Then
                frmDeclaration.Show 1
                If TempVar3 > 0 Then
                    SendSocket Chr$(37) + Chr$(TempVar3) + Chr$(0)
                    SendSocket Chr$(39) + Chr$(Character.Guild)
                End If
            End If
            If x >= 15 And x <= 92 And y >= 435 And y <= 462 Then
                If MsgBox("Are you sure you wish to disband your guild?  Your guild will be delete, and you will not get refunded for any of the guild fees nor will you get the money in your guild's bank account.  Continue?", vbYesNo + vbQuestion, TitleString) = vbYes Then
                    SendSocket Chr$(42)
                    Unload Me
                End If
            End If
            If x >= 98 And x <= 175 And y >= 435 And y <= 462 Then
                If Guild(Character.Guild).hallNum <> 0 Then
                If MsgBox("Are you sure you wish to move out of your guild hall?", vbYesNo + vbQuestion, TitleString) = vbYes Then
                    SendSocket Chr$(44)
                    Unload Me
                End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set pbMyGuild.Picture = Nothing
    Set pbGuilds.Picture = Nothing
    Set pbInfo.Picture = Nothing
    Set pbManage.Picture = Nothing
        
    frmGuilds_Loaded = False
End Sub


Private Sub Label2_Click()
    SendSocket Chr$(82) + txtMOTD
    SendSocket Chr$(39) + Chr$(Selected)
End Sub



Private Sub lblSaveInfo_Click()
    SendSocket Chr$(83) + txtInfo
    SendSocket Chr$(39) + Chr$(Selected)
End Sub

Private Sub lstGuilds_ItemClick(item As cListItem)
    If nextGuild <> item.SubItems(3).Caption Then
        changed = True
        Selected = item.SubItems(3).Caption
        nextGuild = item.SubItems(3).Caption
    End If
End Sub
Private Sub lstGuilds_itemdblClick(item As cListItem)
    If lstEnabled Then
    If item.SubItems(3).Caption >= 0 Then
        If nextGuild <> item.SubItems(3).Caption Or changed = True Then
            Selected = item.SubItems(3).Caption
            SendSocket Chr$(39) + Chr$(Selected)
            lstEnabled = False
            nextGuild = Selected
        Else
            nextTab = 1
            nextGuild = Selected
            setGuildTab
        End If
    End If
    End If
End Sub
Private Sub lstGuilds_ColumnClick(column As cColumn)


    If column <> lstGuilds.Columns(1) Then lstGuilds.Columns(1).SortOrder = eSortOrderNone
    If column <> lstGuilds.Columns(2) Then lstGuilds.Columns(2).SortOrder = eSortOrderNone
    If column <> lstGuilds.Columns(3) Then lstGuilds.Columns(3).SortOrder = eSortOrderNone

    If column = lstGuilds.Columns(3) Then
        column.SortOrder = IIf(column.SortOrder = eSortOrderAscending, eSortOrderDescending, eSortOrderAscending)
        column.SortType = eLVSortNumeric
    Else
        column.SortOrder = IIf(column.SortOrder = eSortOrderAscending, eSortOrderDescending, eSortOrderAscending)
        column.SortType = eLVSortString
    End If
            lstGuilds.ListItems.SortItems
End Sub
Private Sub lstMembers_ColumnClick(column As cColumn)
    If column <> lstMembers.Columns(1) Then lstMembers.Columns(1).SortOrder = eSortOrderNone
    If column <> lstMembers.Columns(2) Then lstMembers.Columns(2).SortOrder = eSortOrderNone
    If column <> lstMembers.Columns(2) Then lstMembers.Columns(3).SortOrder = eSortOrderNone
    If column <> lstMembers.Columns(4) Then lstMembers.Columns(4).SortOrder = eSortOrderNone
    If column <> lstMembers.Columns(5) Then lstMembers.Columns(5).SortOrder = eSortOrderNone



     If column = lstMembers.Columns(2) Then
        lstMembers.Columns(3).SortOrder = IIf(lstMembers.Columns(3).SortOrder = eSortOrderAscending, eSortOrderDescending, eSortOrderAscending)
        lstMembers.Columns(3).SortType = eLVSortString
     ElseIf column = lstMembers.Columns(4) Then
        lstMembers.Columns(4).SortOrder = IIf(column.SortOrder = eSortOrderAscending, eSortOrderDescending, eSortOrderAscending)
        lstMembers.Columns(4).SortType = eLVSortNumeric
     Else
     
        column.SortOrder = IIf(column.SortOrder = eSortOrderAscending, eSortOrderDescending, eSortOrderAscending)
        column.SortType = eLVSortString
     End If
    
    If column = lstMembers.Columns(5) Then column.SortType = eLVSortDate

          lstMembers.ListItems.SortItems
End Sub

Public Sub setGuildTab()
        If nextTab = 0 Then
            ctab = 0
            pbGuilds.Visible = True
            pbInfo.Visible = False
            pbManage.Visible = False
            pbMyGuild.Visible = False
            Set pbGuilds.Picture = LoadPicture("Data\Graphics\Interface\NSGuildGUI.rsc")
        End If
        If nextTab = 1 Then
            ctab = 1
            ctab2 = 1
            lstDeclarations.Visible = False
            lstMembers.Visible = True
            Set pbInfo.Picture = LoadPicture("Data\Graphics\Interface\NSGuildGUI2a.rsc")
            pbInfo.Visible = True
            pbGuilds.Visible = False
            pbManage.Visible = False
            pbMyGuild.Visible = False
        End If
        If nextTab = 2 Then
            ctab = 2
            Set pbMyGuild.Picture = LoadPicture("Data\Graphics\Interface\NSGuildGUI3.rsc")
            pbManage.Visible = False
            pbInfo.Visible = False
            pbGuilds.Visible = False
            pbMyGuild.Visible = True
        End If
        If nextTab = 3 Then
            ctab = 3
            'ctab2 = 1
            
            If ctab2 <= 1 Then
                Set pbManage.Picture = LoadPicture("Data\Graphics\Interface\NSGuildGUI4a.rsc")
                lstDeclarations2.Visible = False
                lstMembers2.Visible = True
                sclSymbol1.Visible = True
                sclSymbol2.Visible = True
                sclSymbol3.Visible = True
                pbSymbol.Visible = True
            ElseIf ctab2 = 2 Then
                lstDeclarations2.Visible = True
                lstMembers2.Visible = False
                Set pbManage.Picture = LoadPicture("Data\Graphics\Interface\NSGuildGUI4b.rsc")
                sclSymbol1.Visible = False
                sclSymbol2.Visible = False
                sclSymbol3.Visible = False
                pbSymbol.Visible = False
            End If
                pbMyGuild.Visible = False
                pbInfo.Visible = False
                pbManage.Visible = True
                pbGuilds.Visible = False
                
                
        End If
End Sub




Private Sub pbGuilds_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseDown Button, Shift, x, y
End Sub

Private Sub pbInfo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseDown Button, Shift, x, y
End Sub

Private Sub pbManage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseDown Button, Shift, x, y
End Sub

Private Sub pbMyGuild_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_MouseDown Button, Shift, x, y
End Sub


Private Sub sclSymbol1_Change()
    Dim hdcObjects As Long
    
    BitBlt pbSymbol.hdc, 0, 0, 20, 20, 0, 0, 0, BLACKNESS
    If sclSymbol3.Value > 0 Then
        hdcObjects = sfcSymbols(3).Surface.GetDC
            BitBlt pbSymbol.hdc, 0, 0, 20, 20, hdcObjects, (sclSymbol3.Value Mod 16) * 20 - 20, (sclSymbol3.Value \ 16) * 20, SRCCOPY
            BitBlt pbSymbol2.hdc, 0, 0, 20, 20, hdcObjects, (sclSymbol3.Value Mod 16) * 20 - 20, (sclSymbol3.Value \ 16) * 20, SRCCOPY
        sfcSymbols(3).Surface.ReleaseDC hdcObjects
    End If
    If sclSymbol2.Value > 0 Then
        hdcObjects = sfcSymbols(2).Surface.GetDC
            TransparentBlt pbSymbol.hdc, 0, 0, 20, 20, hdcObjects, (sclSymbol2.Value Mod 16) * 20 - 20, (sclSymbol2.Value \ 16) * 20, SRCCOPY
            TransparentBlt pbSymbol2.hdc, 0, 0, 20, 20, hdcObjects, (sclSymbol2.Value Mod 16) * 20 - 20, (sclSymbol2.Value \ 16) * 20, SRCCOPY
        sfcSymbols(2).Surface.ReleaseDC hdcObjects
    End If
    If sclSymbol1.Value > 0 Then
        hdcObjects = sfcSymbols(1).Surface.GetDC
            TransparentBlt pbSymbol.hdc, 0, 0, 20, 20, hdcObjects, (sclSymbol1.Value Mod 16) * 20 - 20, (sclSymbol1.Value \ 16) * 20, SRCCOPY
            TransparentBlt pbSymbol2.hdc, 0, 0, 20, 20, hdcObjects, (sclSymbol1.Value Mod 16) * 20 - 20, (sclSymbol1.Value \ 16) * 20, SRCCOPY
        sfcSymbols(1).Surface.ReleaseDC hdcObjects
    End If


    pbSymbol.Refresh
    pbSymbol2.Refresh
End Sub
Private Sub sclSymbol2_Change()
sclSymbol1_Change
End Sub
Private Sub sclSymbol3_Change()
sclSymbol1_Change
End Sub
