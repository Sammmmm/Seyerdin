VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "The Odyssey Classic Odyssey Server"
   ClientHeight    =   1650
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7470
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   7470
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstLog 
      Height          =   1230
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7455
   End
   Begin VB.TextBox txtMessage 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   7455
   End
   Begin VB.Menu mnuServer 
      Caption         =   "&Server"
      Begin VB.Menu mnuServerOptions 
         Caption         =   "&Options"
      End
      Begin VB.Menu mnuBlockAccess 
         Caption         =   "B&lock God Access"
      End
      Begin VB.Menu mnuCompact 
         Caption         =   "&Compact Database"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "&Reports"
      Begin VB.Menu mnuReportsMaps 
         Caption         =   "&Free Maps"
      End
      Begin VB.Menu mnuReportsGods 
         Caption         =   "&Gods"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Admin"
      Begin VB.Menu mnuBan 
         Caption         =   "Ban"
      End
      Begin VB.Menu mnuSquelch 
         Caption         =   "Squelch"
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu mnuReset 
         Caption         =   "&Reset"
         Begin VB.Menu mnuDatabaseResetAccounts 
            Caption         =   "Accounts"
         End
         Begin VB.Menu mnuDatabaseResetObjects 
            Caption         =   "Objects"
         End
         Begin VB.Menu mnuDatabaseResetMonsters 
            Caption         =   "Monsters"
         End
         Begin VB.Menu mnuDatabaseResetScripts 
            Caption         =   "Scripts"
         End
         Begin VB.Menu mnuDatabaseResetNPCs 
            Caption         =   "NPCs"
         End
         Begin VB.Menu mnuDatabaseResetGuilds 
            Caption         =   "Guilds"
         End
         Begin VB.Menu mnuDatabaseResetGods 
            Caption         =   "Gods"
         End
         Begin VB.Menu mnuDatabaseResetBans 
            Caption         =   "Bans"
         End
         Begin VB.Menu mnuDatabaseResetMaps 
            Caption         =   "Maps"
         End
         Begin VB.Menu mnuDatabaseResetPrefixs 
            Caption         =   "Prefixs"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
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

Private Sub Form_Load()
    gHW = Me.hwnd
 '   With nid
 '       .hwnd = frmMain.hwndx
 '       .uId = vbNull
 '       .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
 '       .ucallbackMessage = WM_MOUSEMOVE
 '       .hIcon = frmLoading.Icon
 '       .szTip = TitleString & vbNullChar
 '       .cbSize = Len(nid)
 '       Shell_NotifyIcon NIM_ADD, nid
 '   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
'With nid
'    .hwnd = frmMain.hwnd
'    .uId = vbNull
'    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
'    .ucallbackMessage = WM_MOUSEMOVE
'    .hIcon = frmLoading.Icon
'    .szTip = "Seyerdin Online" & vbNullChar
'    .cbSize = Len(nid)
'End With
'Shell_NotifyIcon NIM_DELETE, nid
ShutdownServer
End Sub
Private Sub Form_Resize()
    If Not Me.WindowState = 1 Then
        lstLog.Width = Me.ScaleWidth
        lstLog.Height = Me.ScaleHeight - txtMessage.Height
        txtMessage.Top = lstLog.Height
        txtMessage.Width = Me.ScaleWidth
    End If
End Sub

Private Sub mnuBan_Click()
    Dim Name As String
       Dim Length As String
       Dim reason As String
       
       Name = InputBox("Enter name of player:", "Ban")
       Length = InputBox("Enter length of ban:", "Ban")
       reason = InputBox("Enter reason for ban:", "Ban")
       Dim A As Long
       
       A = FindPlayer(Name)
       If A > 0 Then
            BanPlayer A, 0, CByte(Length), reason, "Server Ban"
       End If
End Sub

Private Sub mnuBlockAccess_Click()
    mnuBlockAccess.Checked = Not mnuBlockAccess.Checked
End Sub

Private Sub mnuCompact_Click()
    CompactDb
End Sub

Private Sub mnuDatabaseResetAccounts_Click()
    If MsgBox("Are you *sure* you wish to delete every account?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every account -- continue?", vbYesNo) = vbYes Then
            If Not UserRS.BOF Then
                UserRS.MoveFirst
                While Not UserRS.EOF
                    DeleteAccount
                    UserRS.MoveNext
                Wend
            End If
            UserRS.Close
            Set UserRS = Nothing
            db.TableDefs.Delete "Accounts"
            CreateAccountsTable db
            Set UserRS = db.TableDefs("Accounts").OpenRecordset(dbOpenTable)
            UserRS.Index = "User"
        End If
    End If
End Sub

Private Sub mnuDatabaseResetBans_Click()
    Dim A As Long
    
    If MsgBox("Are you *sure* you wish to delete every Ban?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every Ban -- continue?", vbYesNo) = vbYes Then
            BanRS.Close
            Set BanRS = Nothing
            db.TableDefs.Delete "Bans"
            CreateBansTable
            Set BanRS = db.TableDefs("Bans").OpenRecordset(dbOpenTable)
            BanRS.Index = "Number"
            For A = 1 To 50
                With Ban(A)
                    .Banner = ""
                    .user = ""
                    .reason = ""
                    .Name = ""
                    .UnbanDate = 0
                End With
            Next A
        End If
    End If
End Sub

Private Sub mnuDatabaseResetGods_Click()

    If MsgBox("Are you *sure* you wish to delete every God?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every God -- continue?", vbYesNo) = vbYes Then
            If UserRS.BOF = False Then
                UserRS.MoveFirst
                While UserRS.EOF = False
                    If UserRS!Access > 0 Then
                        UserRS.Edit
                        UserRS!Access = 0
                        UserRS.Update
                    End If
                    UserRS.MoveNext
                Wend
            End If
        End If
    End If
End Sub

Private Sub mnuDatabaseResetGuilds_Click()
    Dim A As Long, B As Long
    
    If MsgBox("Are you *sure* you wish to delete every guild?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every guild -- continue?", vbYesNo) = vbYes Then
            GuildRS.Close
            Set GuildRS = Nothing
            db.TableDefs.Delete "Guilds"
            CreateGuildsTable db
            Set GuildRS = db.TableDefs("Guilds").OpenRecordset(dbOpenTable)
            GuildRS.Index = "Number"
            For A = 1 To 255
                With Guild(A)
                    .Bank = 0
                    .Bookmark = 0
                    .Name = ""
                    .Hall = 0
                    .DueDate = 0
                    .sprite = 0
                    .MOTD = 0
                    .Info = 0
                    For B = 0 To 9
                        With .Declaration(B)
                            .Guild = 0
                            .Type = 0
                        End With
                    Next B
                    For B = 0 To 19
                        With .Member(B)
                            .Name = ""
                            .Rank = 0
                            .Deaths = 0
                            .Kills = 0
                            .Renown = 0
                        End With
                    Next B
                End With
            Next A
        End If
    End If
End Sub

Private Sub mnuDatabaseResetMaps_Click()
    Dim A As Long
    If MsgBox("Are you *sure* you wish to delete every map?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every map -- continue?", vbYesNo) = vbYes Then
            
            MapRS.Close
            
            db.TableDefs.Delete "Maps"
            CreateMapsTable
        
            Set MapRS = db.TableDefs("Maps").OpenRecordset(dbOpenTable)
        
            If MapRS.BOF = False Then
                MapRS.MoveFirst
                While MapRS.EOF = False
                    A = MapRS!Number
                    If A > 0 Then
                        LoadMap A, MapRS!Data
                        ResetMap A
                    End If
                    MapRS.MoveNext
                Wend
            End If
            
            End
        End If
    End If
End Sub

Private Sub mnuDatabaseResetMonsters_Click()
    Dim A As Long
    
    If MsgBox("Are you *sure* you wish to delete every monster?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every monster -- continue?", vbYesNo) = vbYes Then
            MonsterRS.Close
            Set MonsterRS = Nothing
            db.TableDefs.Delete "Monsters"
            CreateMonstersTable
            Set MonsterRS = db.TableDefs("Monsters").OpenRecordset(dbOpenTable)
            MonsterRS.Index = "Number"
            For A = 1 To MAXITEMS
                With monster(A)
                    .Armor = 0
                    .Agility = 0
                    .Description = ""
                    .Flags = 0
                    .HP = 0
                    .Name = ""
                    .Sight = 0
                    .sprite = 0
                    .Min = 0
                    .Max = 0
                    .Object(0) = 0
                    .Object(1) = 0
                    .Object(2) = 0
                    .Value(0) = 0
                    .Value(1) = 0
                    .Value(2) = 0
                    .Experience = 0
                End With
            Next A
        End If
    End If
End Sub

Private Sub mnuDatabaseResetNPCs_Click()
    Dim A As Long, B As Long
    
    If MsgBox("Are you *sure* you wish to delete every NPC?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every NPC -- continue?", vbYesNo) = vbYes Then
            NPCRS.Close
            Set NPCRS = Nothing
            db.TableDefs.Delete "NPCS"
            CreateNPCsTable
            Set NPCRS = db.TableDefs("NPCS").OpenRecordset(dbOpenTable)
            NPCRS.Index = "Number"
            For A = 1 To 255
                With NPC(A)
                    .Name = ""
                    .JoinText = ""
                    .LeaveText = ""
                    .Flags = 0
                    For B = 0 To 9
                        With .SaleItem(B)
                            .GiveObject = 0
                            .GiveValue = 0
                            .TakeObject = 0
                            .TakeValue = 0
                        End With
                    Next B
                End With
            Next A
        End If
    End If
End Sub
Private Sub mnuDatabaseResetObjects_Click()
    Dim A As Long
    
    If MsgBox("Are you *sure* you wish to delete every object?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every object -- continue?", vbYesNo) = vbYes Then
            ObjectRS.Close
            Set ObjectRS = Nothing
            db.TableDefs.Delete "Objects"
            CreateObjectsTable db
            Set ObjectRS = db.TableDefs("Objects").OpenRecordset(dbOpenTable)
            ObjectRS.Index = "Number"
            
            For A = 1 To 255
                With Object(A)
                    .Data(0) = 0
                    .Data(1) = 0
                    .Data(2) = 0
                    .Data(3) = 0
                    .Flags = 0
                    .Name = ""
                    .Picture = 0
                    .Type = 0
                End With
            Next A
        End If
    End If
End Sub

Private Sub mnuDatabaseResetPrefixs_Click()
    Dim A As Long
    
    If MsgBox("Are you *sure* you wish to delete every prefix?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every prefix -- continue?", vbYesNo) = vbYes Then
            PrefixRS.Close
            Set PrefixRS = Nothing
            db.TableDefs.Delete "Prefix"
            CreatePreFixTable
            Set PrefixRS = db.TableDefs("Prefix").OpenRecordset(dbOpenTable)
            PrefixRS.Index = "Number"
            
            For A = 1 To 255
                With prefix(A)
                    .Name = ""
                    .Flags = 0
                    .Light.Radius = 0
                    .Light.Intensity = 0
                    .Max = 0
                    .Min = 0
                    .ModType = 0
                    .Strength1 = 0
                    .Strength2 = 0
                    .Weakness1 = 0
                    .Weakness2 = 0
                End With
            Next A
        End If
    End If
End Sub

Private Sub mnuDatabaseResetScripts_Click()
    If MsgBox("Are you *sure* you wish to delete every script?", vbYesNo) = vbYes Then
        If MsgBox("About to reset every script -- continue?", vbYesNo) = vbYes Then
            
            ScriptRS.Close
            
            db.TableDefs.Delete "Scripts"
            CreateScriptsTable
        
            Set ScriptRS = db.TableDefs("Scripts").OpenRecordset(dbOpenTable)
        
            End
        End If
    End If
End Sub

Private Sub mnuReportsGods_Click()
    Open "report.txt" For Output As #1
    
    Print #1, "***Seyerdin God Report ***"
    Print #1, ""
    
    If UserRS.BOF = False Then
        UserRS.MoveFirst
        While UserRS.EOF = False
            If UserRS!Access > 0 Then
                Print #1, "User=" + UserRS!user + " Name=" + UserRS!Name + " Access=" + CStr(UserRS!Access)
            End If
            UserRS.MoveNext
        Wend
    End If
    
    Close #1
    
    Shell "notepad.exe report.txt", vbNormalFocus
End Sub
Private Sub mnuReportsMaps_Click()
    Dim A As Long, StartFree As Long, IsFree As Boolean
    Open "report.txt" For Output As #1
    
    Print #1, "***Seyerdin Free Map Report ***"
    Print #1, ""
    
    IsFree = False
    
    For A = 1 To 5000
        MapRS.Seek "=", A
        If MapRS.NoMatch = False Then
            If IsFree = True Then
                If StartFree < A - 1 Then
                    Print #1, CStr(StartFree) + "-" + CStr(A - 1)
                Else
                    Print #1, CStr(A - 1)
                End If
                IsFree = False
            End If
        Else
            If IsFree = False Then
                StartFree = A
                IsFree = True
            End If
        End If
    Next A
    
    If IsFree = True Then
        If StartFree < 5000 Then
            Print #1, CStr(StartFree) + "-5000"
        Else
            Print #1, "5000"
        End If
    End If
    Close #1
    
    Shell "notepad.exe report.txt", vbNormalFocus
End Sub

Private Sub mnuServerOptions_Click()
    frmOptions.Show 1
End Sub

Private Sub mnuSquelch_Click()
    Dim Name As String
    Dim Length As String
    
    Name = InputBox("Enter name of player:", "Squelch")
    Length = InputBox("Enter length of squelch:", "Squelch")
    
    Dim A As Long
    
    A = FindPlayer(Name)
    If A > 0 Then
        With player(A)
            If .Mode = modePlaying Then
                .Squelched = CByte(Length)
                SendSocket A, Chr2(23) + Chr(.Squelched)
                If .Squelched > 0 Then
                    SendAll Chr2(56) + Chr2(15) + player(A).Name & " has been squelched by the server for " + CStr(.Squelched) + " minutes!"
                Else
                    SendToGods Chr2(56) + Chr2(15) + player(A).Name & " has been unsquelched by the server!"
                End If
            End If
        End With
    End If
End Sub

Private Sub txtMessage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 10 Then
        If Len(txtMessage) > 0 Then
            SendAll Chr2(30) + txtMessage
            PrintLog "Server Message: " + txtMessage
            txtMessage = ""
        End If
    End If
End Sub
