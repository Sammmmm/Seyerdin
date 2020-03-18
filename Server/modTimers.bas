Attribute VB_Name = "modTimers"
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

Public Sub ObjectTimer(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
Dim Mapiterator As Byte, mapNum As Integer, A As Byte
Dim ST1 As String

    For Mapiterator = 1 To MaxUsers
        If CurrentMaps(Mapiterator) > 0 Then
            mapNum = CurrentMaps(Mapiterator)
            With map(mapNum)
                ST1 = ""
                For A = 0 To 49
                    With .Object(A)
                        If .Object > 0 Then
                            If map(mapNum).Tile(.x, .y).Att <> 5 And .TimeStamp > 0 Then
                                If .TimeStamp < GetTickCount Then
                                    .Object = 0
                                    .TimeStamp = 0
                                    ST1 = ST1 + DoubleChar(2) + Chr2(15) + Chr2(A)
                                End If
                            End If
                        End If
                    End With
                Next A
                If ST1 <> "" Then
                    SendToMapRaw mapNum, ST1
                End If
            End With
        End If
    Next Mapiterator
End Sub

Public Sub PlayerTimer(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
Dim playerNum As Long, LastHP As Integer, LastEnergy As Integer, LastMana As Integer, D As Integer, A As Long, B As Long, C As Long
Dim ST1 As String


On Error GoTo Error_Handler

For playerNum = 1 To currentMaxUser
    With player(playerNum)
        If .InUse = True Then
            LastHP = .HP
            LastEnergy = .Energy
            LastMana = .Mana
            
            ST1 = ""
            .FloodTimer = 0
            .SquelchTimer = 0

            If GetTickCount - .LastMsg >= 30000 Then
                'If Not .Mode = modePlaying Then
                If (.Mode = modeNotConnected) Then
                    CloseClientSocket playerNum, True
                Else
                    A = SendData(.Socket, DoubleChar(1) & Chr2(6))
                    If A = SOCKET_ERROR Then
                        CloseClientSocket playerNum, True
                    End If
                End If
                'End If
                
                .LastMsg = GetTickCount
            End If
            
            If .Mode = modePlaying Then
                .DeferSends = True
                D = .HP

                For A = 1 To MaxPlayerTimers
                    If .ScriptTimer(A) > 0 And GetTickCount() >= .ScriptTimer(A) Then
                        Parameter(0) = playerNum
                        .ScriptTimer(A) = 0
                        RunScript .Script(A)
                    End If
                Next A
                
                For A = 1 To MAXSTATUS
                    If A > 30 Then
                        If .StatusData(A).timer > 0 Then
                            If .StatusData(A).timer = 1 Then
                                Select Case A
                                    Case SE_HPMOD To SE_CRITICALCHANCEMOD
                                        .CalculateStats = True
                                End Select
                            End If
                            .StatusData(A).timer = .StatusData(A).timer - 1
                        End If
                    Else
                        If .StatusEffect And (2 ^ A) Then
                            If .StatusData(A).timer > 0 Then
                                .StatusData(A).timer = .StatusData(A).timer - 1
                                Select Case A
                                    Case SE_EXHAUST
                                        .Energy = .Energy - .StatusData(SE_EXHAUST).Data(0)
                                    Case SE_POISON
                                        If .StatusData(SE_POISON).Data(0) > 1 Then
                                            If .HP = 1 Then
                                                RemoveStatusEffect playerNum, SE_POISON
                                            Else
                                                .HP = .HP - .StatusData(SE_POISON).Data(0) ' * 256& + .StatusData(SE_POISON).Data(1))
                                            End If
                                            If GetTickCount + 5000 > .combatTimer Then .combatTimer = GetTickCount + 5000
                                        Else
                                            .HP = .HP - 1
                                        End If
                                    Case SE_MUTE
                                    
                                    Case SE_CRIPPLE
                                        .HP = .HP - .StatusData(SE_CRIPPLE).Data(0) ' * 256& + .StatusData(SE_CRIPPLE).Data(1))
                                        .Energy = .Energy - IIf(.StatusData(SE_CRIPPLE).Data(0) / 8 > 0, .StatusData(SE_CRIPPLE).Data(0) / 8, 1) '* 256& + .StatusData(SE_CRIPPLE).Data(1))
                                        .Mana = .Mana - .StatusData(SE_CRIPPLE).Data(0) '* 256& + .StatusData(SE_CRIPPLE).Data(1))
                                        If GetTickCount + 5000 > .combatTimer Then .combatTimer = GetTickCount + 5000
                                    Case SE_BERSERK
                                    Case SE_REGENERATION
                                        If .HP + .StatusData(SE_REGENERATION).Data(1) > .MaxHP Then
                                            If .HP < .MaxHP Then
                                                CreateFloatingText .map, .x, .y, 10, CStr(Int(.MaxHP - .HP))
                                            Else
                                                CreateFloatingText .map, .x, .y, 10, "0"
                                            End If
                                            .HP = .MaxHP
                                        Else
                                            .HP = .HP + .StatusData(SE_REGENERATION).Data(1)
                                            CreateFloatingText .map, .x, .y, 10, CStr(.StatusData(SE_REGENERATION).Data(1))
                                        End If
                                End Select
                            Else
                                RemoveStatusEffect CByte(playerNum), A
                                Select Case A
                                    Case SE_HPMOD To SE_CRITICALCHANCEMOD
                                        .CalculateStats = True
                                End Select
                            End If
                        End If
                    End If
                Next A
                
                
                If .Buff.timer > 0 Then
                    .Buff.timer = .Buff.timer - 1
                    If .Buff.timer = 0 Then
                        .Buff.Type = 0
                        SendToMap2 .map, Chr2(115) + Chr2(playerNum) + Chr2(0)
                        .CalculateStats = True
                    End If
                End If
                

                A = 2 + .HPRegen + (.Level \ LevelsPerHpRegen) + GetStatPerBonus(.Constitution, ConstitutionPerHPRegen) + GetStatPerBonus(.Wisdom, PietyPerHPRegen)
                B = 2 + GetStatPerBonus(.Endurance, EndurancePerEnergyRegen)
                C = 1 + (.Level \ LevelsPerManaRegen) + GetStatPerBonus(.Intelligence, IntelligencePerManaRegen) + GetStatPerBonus(.Wisdom, PietyPerManaRegen)
                
                If .SkillLevel(SKILL_CONVALESCENCE) Then
                    A = A + 1 + .SkillLevel(SKILL_CONVALESCENCE) \ 2
                End If
                
                If GetTickCount >= .combatTimer Then
                    .HP = .HP + A * 3
                Else
                    .HP = .HP + A
                End If
                
                If .HP > .MaxHP Then .HP = .MaxHP
                If .HP < 1 Then .HP = 1
                
                .Energy = .Energy + B
                If .Energy > .MaxEnergy Then .Energy = .MaxEnergy
                If .Energy < 0 Then .Energy = 0
                
                
                    If .StatusData(SE_MANAMULT).Data(0) > 0 And .StatusData(SE_MANAMULT).timer > 0 Then
                            .Mana = .Mana + ((C * .StatusData(SE_MANAMULT).Data(0)) / 10)
                    Else
                        If GetTickCount >= .combatTimer Then
                            .Mana = .Mana + C * 3
                        Else
                            .Mana = .Mana + C
                        End If
                    End If
                If .Mana > .MaxMana Then .Mana = .MaxMana
                If .Mana < 1 Then .Mana = 1
                A = 0
                If LastHP <> .HP Then
                    ST1 = ST1 + DoubleChar(3) + Chr2(46) + DoubleChar(CInt(.HP))
                    A = 1
                End If
                If LastEnergy <> .Energy Then
                    ST1 = ST1 + DoubleChar(3) + Chr2(47) + DoubleChar(CInt(.Energy))
                End If
                If LastMana <> .Mana Then
                    ST1 = ST1 + DoubleChar(3) + Chr2(48) + DoubleChar(CInt(.Mana))
                    A = 1
                End If
                If ST1 <> "" Then
                    SendRaw2 playerNum, ST1
                End If
                If A = 1 Then
                        SendToPartyAllBut2 .Party, playerNum, Chr2(104) + Chr2(0) + Chr2(playerNum) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                        SendToGods2 Chr2(104) + Chr2(0) + Chr2(playerNum) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                        SendToGuildAllBut2 playerNum, CLng(.Guild), Chr2(104) + Chr2(0) + Chr2(A) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                End If
            End If
        End If
        If .CalculateStats Then CalculateStats playerNum
        .DeferSends = False
    End With
Next playerNum
    
    For A = 1 To currentMaxUser
         If player(A).InUse = True Then
            If player(A).St <> "" Then
                FlushSocket A
            End If
        End If
    Next A
            
   
            
Exit Sub
Error_Handler:
    Open App.Path + "/LOG.TXT" For Append As #1
        ST1 = ""
        Print #1, Err.Number & "/" & Err.Description & "  playertimer " & "/" & (idEvent)
    Close #1
    Unhook
    End
End Sub

Public Sub MinuteTimer(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
Dim A As Long, B As Long, C As Long
On Error GoTo Error_Handler

    World.HourCounter = World.HourCounter + 1
    If World.HourCounter = 4 Then
        World.HourCounter = 0
        World.Hour = World.Hour + 1
        If World.Hour > 24 Then World.Hour = 1
        If World.Hour = 7 Then
            blnNight = False
            'SendAll chr2(54)
        End If
        If World.Hour = 22 Then
            blnNight = True
            'SendAll chr2(55)
        End If
        SendAll2 Chr2(110) + Chr2(World.Hour)
    End If
    
    If (World.Hour Mod 2) = 0 And World.HourCounter = 1 Then
        For A = 1 To currentMaxUser
            player(A).speedhack = 0
        Next A
    End If
        
    
    If World.BackupInterval > 0 Then
        If World.HourCounter = 1 And World.Hour = 1 Then
            BackupCounter = 0
            'Backup Server Data
            For A = 1 To currentMaxUser
                If player(A).Mode = modePlaying Then
                    SavePlayerData A
                End If
            Next A
            SaveFlags
            SaveObjects
    
    
                    
        End If
    End If
    
    For A = 1 To currentMaxUser
        With player(A)
            If .InUse Then
                If .Mode = modePlaying Then
                    If .Squelched > 0 Then
                        .Squelched = .Squelched - 1
  '                      If .Squelched = 0 Then
  '                          If GetIpSquelchTime(.ip) >= 1 Then
  '                              SendSocket2 A, Chr2(23) + Chr2(.Squelched)
  '                              SendToGods Chr2(56) + Chr2(15) + .Name + " has been unsquelched!"
  '                          End If
  '                      End If
                    End If
                End If
            End If
        End With
    Next A
    
    For A = 0 To 255
        If ipSquelches(A).Time > 0 Then
            ipSquelches(A).Time = ipSquelches(A).Time - 1
            If ipSquelches(A).Time = 0 Then
                For B = 1 To currentMaxUser
                    If player(B).ip = ipSquelches(A).ip Then
                        SendSocket2 B, Chr2(23) + Chr2(0)
                        player(B).Squelched = 0
                    End If
                Next B
                'SendIp ipSquelches(A).ip, Chr2(23) + Chr(0)
                ipSquelches(A).ip = vbNullString
                
                SendToGods Chr2(56) + Chr2(15) + ipSquelches(A).ip + " has been unsquelched!"
            End If
        End If
    Next A
    
    If World.LastUpdate <> CLng(Date) Then
        World.LastUpdate = CLng(Date)
        DataRS.Edit
        DataRS!LastUpdate = World.LastUpdate
        DataRS.Update
        
        'Update Guilds
        For A = 1 To 255
            With Guild(A)
                If .Name <> "" Then
                    If .Bank < 0 And World.LastUpdate >= .DueDate Then
                        'Debt not payed, delete guild
                        DeleteGuild A, 0
                    ElseIf CountGuildMembers(A) < 3 Then
                        'Not enough members, guild deleted
                        DeleteGuild A, 1
                    Else
                        If .Bank >= 0 Then
                            C = 0
                        Else
                            C = 1
                        End If
                        
                        'Pay bill
                        .Bank = .Bank - 200
                        For B = 0 To 19
                            If .Member(B).Name <> "" Then
                                .Bank = .Bank - 50
                            End If
                        Next B
                        If .Hall > 0 Then
                            .Bank = .Bank - Hall(.Hall).Upkeep
                        End If
                        If C = 0 And .Bank < 0 Then
                            .DueDate = CLng(Date) + 2
                        End If
                        If .Bank >= 0 Then
                            SendToGuild A, Chr2(74) + Chr2(74) + QuadChar(.Bank)
                        Else
                            SendToGuild A, Chr2(74) + QuadChar(Abs(.Bank)) + QuadChar(.DueDate)
                        End If
                        GuildRS.Seek "=", A
                        If GuildRS.NoMatch = False Then
                            GuildRS.Edit
                            GuildRS!Bank = .Bank
                            GuildRS!DueDate = .DueDate
                            GuildRS.Update
                        End If
                    End If
                End If
            End With
        Next A
    End If
    
    'If Hour(Time) = 3 And Minute(Time) = 1 Then
    '    SendAll2 Chr2(56) + Chr2(9) + StrConv("Warning: The server will restart in one hour to do a nightly backup.", vbUnicode)
    'End If
    'If Hour(Time) = 3 And Minute(Time) = 31 Then
    '    SendAll2 Chr2(56) + Chr2(9) + StrConv("Warning: The server will restart in 30 minutes to do a nightly backup.", vbUnicode)
    'End If
   ' If Hour(Time) = 3 And Minute(Time) = 51 Then
   '     SendAll2 Chr2(56) + Chr2(9) + StrConv("Warning: The server will restart in 10 minutes to do a nightly backup.", vbUnicode)
   ' End If
   ' If Hour(Time) = 4 And Minute(Time) = 0 Then
   '     SendAll2 Chr2(56) + Chr2(9) + StrConv("Warning: The server will restart in 1 minute to do a nightly backup.", vbUnicode)
   ' End If
    If Hour(Time) = 4 And Minute(Time) = 1 Then
        CompactDb
    End If
    
    
    RunScript ("MINUTETIMER")
    DoEvents
    If Minute(Time) Mod 10 = 0 Then
        UpdateLeaderboards
    End If
    DoEvents
        
            
Exit Sub
Error_Handler:
    Open App.Path + "/LOG.TXT" For Append As #1
        Print #1, Err.Number & "/" & Err.Description & "  minutetimer " & "/" & (idEvent)
    Close #1
    Unhook
    End
End Sub

Public Sub SocketQueueTimer(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
Dim A As Long, B As Long
Dim ST1 As String

On Error GoTo Error_Handler

For A = 1 To MaxUsers
    If CloseSocketQueue(A).user > 0 Then
        If CloseSocketQueue(A).TimeStamp < GetTickCount Then
            B = CloseSocketQueue(A).user
            CloseSocketQueue(A).user = 0
            CloseSocketQueue(A).TimeStamp = 0
            CloseClientSocket B
        End If
    End If
Next A

For A = 1 To 100
    If uTimers(A).InUse Then
        With uTimers(A)
            If .timer = 1 Then
                Parameter(0) = .Parm(0)
                Parameter(1) = .Parm(1)
                Parameter(2) = .Parm(2)
                Parameter(3) = .Parm(3)
                ST1 = .Script
                .Parm(0) = 0
                .Parm(1) = 0
                .Parm(2) = 0
                .Parm(3) = 0
                .Script = ""
                .InUse = False
                .timer = 0
                RunScript ST1
            ElseIf .timer > 1 Then
                .timer = .timer - 1
            End If
        End With
    End If
Next A

            
Exit Sub
Error_Handler:
    Open App.Path + "/LOG.TXT" For Append As #1
        ST1 = ""
        Print #1, Err.Number & "/" & Err.Description & "  modTimers" & "/" & (idEvent)
    Close #1
    Unhook
    End
End Sub

Public Sub MapTimer(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
Dim A As Long, B As Long, C As Long, D As Long, Mapiterator As Long, mapNum As Long, J As Long, E As Long, i As Long, F As Long
Dim Cont As Boolean
Dim ST1 As String

On Error GoTo Error_Handler
For A = 1 To currentMaxUser
                player(A).DeferSends = True
            Next A
            For Mapiterator = 1 To currentMaxUser
                If CurrentMaps(Mapiterator) > 0 Then
                    mapNum = CurrentMaps(Mapiterator)
                    If mapNum > 0 Then
                        With map(mapNum)
                            If ExamineBit(.Flags(1), 4) Then 'FISH
                                If .FishCounter = 0 Then
                                    .FishCounter = 5 + CInt(Rnd * 15)
                                    If CInt(Rnd * 3) = 1 Then .FishCounter = CInt(Rnd * 3) + 1
                                    A = CInt(Rnd * 11)
                                    B = CInt(Rnd * 11)
                                    With .Tile(A, B)
                                        C = 0
                                        If .Ground = 1472 Or .Ground = 1496 Or .Ground = 1504 Or .Ground = 1603 Or .Ground = 863 Or .Ground = 864 Or .Ground = 857 Or .Ground = 849 Then
                                            C = 1
                                        ElseIf .Ground = 10 Or .Ground = 11 Or .Ground = 9 Or .Ground = 2 Or .Ground = 3 Or .Ground = 4 Or .Ground = 16 Or .Ground = 17 Then
                                            C = 1
                                        ElseIf .Ground = 18 Or .Ground = 26 Or .Ground = 27 Or .Ground = 879 Or .Ground = 880 Or .Ground = 881 Or .Ground = 886 Or .Ground = 887 Then
                                            C = 1
                                        ElseIf .Ground = 888 Or .Ground = 893 Or .Ground = 894 Or .Ground = 895 Or .Ground = 910 Or .Ground = 917 Or .Ground = 912 Or .Ground = 923 Then
                                            C = 1
                                        ElseIf .Ground = 924 Or .Ground = 930 Or .Ground = 931 Or .Ground = 925 Or .Ground = 936 Or .Ground = 943 Or .Ground = 1373 Or .Ground = 1374 Then
                                            C = 1
                                        ElseIf .Ground = 1394 Or .Ground = 1395 Or .Ground = 1396 Or .Ground = 1401 Or .Ground = 1402 Then
                                            C = 1
                                        End If
                                        If C = 1 Then
                                            SendToMap mapNum, Chr2(139) + Chr2(A) + Chr2(B)
                                            With map(mapNum)
                                            For C = 1 To 10
                                                If .Fish(C).TimeStamp < GetTickCount Then
                                                    .Fish(C).TimeStamp = GetTickCount + 2200
                                                    .Fish(C).x = A
                                                    .Fish(C).y = B
                                                    Exit For
                                                End If
                                            Next C
                                            End With
                                        End If
                                    End With
                                End If
                                .FishCounter = .FishCounter - 1
                                
                            End If
                            
                            ST1 = ""
                            For A = 0 To 9
                                If .Door(A).Att > 0 Or .Door(A).Wall > 0 Then
                                    If GetTickCount - .Door(A).t > 10000 Then
                                        .Tile(.Door(A).x, .Door(A).y).Att = .Door(A).Att
                                        .Tile(.Door(A).x, .Door(A).y).WallTile = .Door(A).Wall
                                        .Door(A).Att = 0
                                        .Door(A).Wall = 0
                                        ST1 = ST1 + DoubleChar(2) + Chr2(37) + Chr2(A)
                                    End If
                                End If
                            Next A
                            For A = 0 To 9
                                With .monster(A)
                                    If .monster > 0 Then

                                        If monster(.monster).sprite > 0 Then
                                            If (ExamineBit(monster(.monster).Flags, 1) = False Or blnNight = False) And (ExamineBit(monster(.monster).Flags, 2) = False Or blnNight = True) Then
                                                B = 0
                                                If monster(.monster).Flags And MONSTER_TICK Then
                                                    Parameter(0) = mapNum
                                                    Parameter(1) = A
                                                    Parameter(2) = .x
                                                    Parameter(3) = .y
                                                    Parameter(4) = .TargType
                                                    Parameter(5) = .Target
                                                    B = RunScript("MONSTERTICK" & .monster)
                                                End If
                                                If .PoisonLength > 0 Then 'Poison
                                                    .PoisonLength = .PoisonLength - 1
                                                    If (.PoisonLength Mod 4) = 0 Then
                                                        Parameter(0) = mapNum
                                                        Parameter(1) = A
                                                        Parameter(2) = .Poison - 100
                                                        Parameter(3) = .PoisonLength / 4
                                                        RunScript ("POISONTICK")
                                                        If .HP - (.Poison - 100) > monster(.monster).HP Then
                                                            .HP = monster(.monster).HP
                                                        Else
                                                            If .HP >= .Poison - 100 Then
                                                                .HP = .HP - (.Poison - 100)
                                                            Else
                                                                .HP = 1
                                                            End If
                                                        End If
                                                        SendToMap mapNum, Chr2(132) + Chr2(A) + DoubleChar(.HP)
                                                        If .PoisonLength = 0 Then .Poison = 0
                                                    End If
                                                End If
                                                PrintCrashDebug 5, 7
                                                If B = 0 Then
                                                    If .MonsterQueue(0).Action = QUEUE_EMPTY Then
                                                        If .AttackCounter = 0 Then
                                                            If .TargType = TargTypePlayer And .Target > 0 Then
                                                                .MoveSpeed = .MoveSpeed
                                                                B = .Target
                                                                C = .x
                                                                D = .y
                                                                Cont = Sqr(CSng(CLng(player(B).x) - C) ^ 2 + CSng(CLng(player(B).y) - D) ^ 2) <= 1
                                                                If Cont = False And monster(.monster).Flags2 And MONSTER_LARGE Then
                                                                    Cont = Sqr(CSng(CLng(player(B).x) - (C + 1)) ^ 2 + CSng(CLng(player(B).y) - (D)) ^ 2) <= 1
                                                                    If Cont = False Then Cont = Sqr(CSng(CLng(player(B).x) - (C + 1)) ^ 2 + CSng(CLng(player(B).y) - (D + 1)) ^ 2) <= 1
                                                                    If Cont = False Then Cont = Sqr(CSng(CLng(player(B).x) - (C)) ^ 2 + CSng(CLng(player(B).y) - (D + 1)) ^ 2) <= 1
                                                                End If
                                                                If Cont Then
                                                                    'Attack Player
                                                                    If .AttackCounter = 0 Then
                                                                        .AttackCounter = .AttackSpeed
                                                                            If player(B).x > C + (monster(.monster).Flags2 And MONSTER_LARGE) And .D <> 3 Then
                                                                                .D = 3
                                                                                ST1 = ST1 + DoubleChar(2) + Chr2(40) + Chr2(A * 16 + .D)
                                                                                Call MonsterMove(mapNum, A)
                                                                            ElseIf player(B).x < C And .D <> 2 Then
                                                                                .D = 2
                                                                                ST1 = ST1 + DoubleChar(2) + Chr2(40) + Chr2(A * 16 + .D)
                                                                                Call MonsterMove(mapNum, A)
                                                                            ElseIf player(B).y > D + (monster(.monster).Flags2 And MONSTER_LARGE) And .D <> 1 Then
                                                                                .D = 1
                                                                                ST1 = ST1 + DoubleChar(2) + Chr2(40) + Chr2(A * 16 + .D)
                                                                                Call MonsterMove(mapNum, A)
                                                                            ElseIf player(B).y < D And .D <> 0 Then
                                                                                .D = 0
                                                                                ST1 = ST1 + DoubleChar(2) + Chr2(40) + Chr2(A * 16 + .D)
                                                                                Call MonsterMove(mapNum, A)
                                                                            End If
                                                                        J = 0
                                                                        
                                                                        
nextAttack:
PrintCrashDebug 5, 10
                                                                        If .monster = 0 Then GoTo nexta
                                                                        If J <> 0 Then
                                                                            B = J
                                                                            C = .x
                                                                            D = .y
                                                                        End If
    PrintCrashDebug 10, 1
    PrintCrashDebug CLng(player(B).map), mapNum
    PrintCrashDebug CLng(player(B).x), CLng(player(B).y)
    PrintCrashDebug CLng(.x), CLng(.y)
                                                                        Cont = CanAttack(CLng(.x), CLng(.y), CLng(player(B).x), CLng(player(B).y), player(B).map)

                                                                        If Cont = False And monster(.monster).Flags2 And MONSTER_LARGE Then
                                                                            Cont = CanAttack(CLng(.x + 1), CLng(.y), CLng(player(B).x), CLng(player(B).y), player(B).map)
                                                                            If Cont = False Then Cont = CanAttack(CLng(.x + 1), CLng(.y + 1), CLng(player(B).x), CLng(player(B).y), player(B).map)
                                                                            If Cont = False Then Cont = CanAttack(CLng(.x), CLng(.y + 1), CLng(player(B).x), CLng(player(B).y), player(B).map)
                                                                        End If
                                                                        PrintCrashDebug 10, 2
                                                                        If Cont Then
                                                                            If (monster(.monster).Flags2 And MONSTER_LARGE) Then
                                                                                Cont = True
                                                                            Else
                                                                                Cont = False
                                                                            End If
                                                                            'STATUS: ABSCOND
                                                                            PrintCrashDebug 10, 3
                                                                            If GetStatusEffect(B, SE_ABSCOND) Then
                                                                                If Int(Rnd * 100) < player(B).StatusData(SE_ABSCOND).Data(0) Then
                                                                                    SetStatusEffect B, SE_INVISIBLE
                                                                                    player(B).StatusData(SE_INVISIBLE).Data(1) = CAN_MOVE
                                                                                    player(B).StatusData(SE_INVISIBLE).timer = 5
                                                                                End If
                                                                            End If
                                                                            PrintCrashDebug 5, 12

                                                                            C = Int(Rnd * 255)
                                                                            If C > PlayerEvasion(B) Then
                                                                                PrintCrashDebug 10, 6
                                                                                C = Int(Rnd * (monster(.monster).Max - monster(.monster).Min)) + monster(.monster).Min
                                                                                C = PlayerArmor(B, C)
                                                                                

                                                                                E = IIf(Int(Rnd * 100) < 2, 1, 0)
                                                                                Parameter(0) = A
                                                                                Parameter(1) = B
                                                                                Parameter(2) = C
                                                                                Parameter(3) = E
                                                                                J = C
                                                                                C = RunScript("MONSTERATTACK")
                                                                                If E = 1 Then C = C * 1.8
                                                                                If C < 0 Then C = 0
                                                                                If C > 9999 Then C = 9999
                                                                                PrintCrashDebug 5, 13
                                                                                '''''''''''''''''''''''mAttack area
                                                                                SendSocket B, Chr2(50) + Chr2(0) + Chr2(A) + DoubleChar(CInt(C))
                                                                                SendToMapAllBut mapNum, B, Chr2(41) + Chr2(A)
                                                                                With player(B)
                                                                                    If GetTickCount + 10000 > .combatTimer Then .combatTimer = GetTickCount + 10000
                                                                                PrintCrashDebug 10, 7
                                                                                    If C >= .HP Then
                                                                                        If .SkillLevel(SKILL_OPPORTUNIST) > 0 And .deathStamp + 60000 < GetTickCount Then
                                                                                            .deathStamp = GetTickCount
                                                                                            CreateFloatingText mapNum, .x, .y, 12, "Opportunist!"
                                                                                        Else
                                                                                            'Player Died
                                                                                            SendSocket B, Chr2(53) + DoubleChar(map(mapNum).monster(A).monster) 'Monster Killed You
                                                                                            SendToMapAllBut .map, B, Chr2(62) + Chr2(B) + DoubleChar(map(mapNum).monster(A).monster) 'Player was killed by monster
                                                                                            CreateFloatingEvent mapNum, .x, .y, FT_ENDED
                                                                                            PlayerDied B, False, A
                                                                                        End If
                                                                                    Else
                                                                                    PrintCrashDebug 10, 8
                                                                                        .HP = .HP - C
                                                                                        
                                                                                        SendToPartyAllBut .Party, B, Chr2(104) + Chr2(0) + Chr2(B) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                                                                                        SendToGods Chr2(104) + Chr2(0) + Chr2(B) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                                                                                        SendToGuildAllBut B, CLng(.Guild), Chr2(104) + Chr2(0) + Chr2(B) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                                                                                        
                                                                                        
                                                                                        
                                                                                        'Do status effects
                                                                                        If monster(map(mapNum).monster(A).monster).cStatusEffect Then 'If any status effects are set, should be a non zero value
                                                                                            For i = 0 To 4
                                                                                                If ExamineBit(monster(map(mapNum).monster(A).monster).cStatusEffect, i) Then
                                                                                                    Select Case i
                                                                                                        Case 1 'Poison
                                                                                                            If Int(Rnd * (100 + .ResistPoison)) < 10 Then
                                                                                                                .StatusData(SE_POISON).Data(0) = monster(map(mapNum).monster(A).monster).Level + 1
                                                                                                                .StatusData(SE_POISON).timer = 12
                                                                                                                SetStatusEffect CByte(B), SE_POISON
                                                                                                            End If
                                                                                                        Case 2 'Mute
                                                                                                            If Int(Rnd * (100 + .ResistPoison)) < 10 Then
                                                                                                                .StatusData(SE_MUTE).timer = monster(map(mapNum).monster(A).monster).Level / 10 + 3
                                                                                                                SetStatusEffect CByte(B), SE_MUTE
                                                                                                            End If
                                                                                                        Case 3 'Exhaust
                                                                                                            If Int(Rnd * (100 + .ResistPoison)) < 10 Then
                                                                                                                .StatusData(SE_EXHAUST).Data(0) = monster(map(mapNum).monster(A).monster).Level / 13 + 2
                                                                                                                .StatusData(SE_EXHAUST).timer = 10
                                                                                                                SetStatusEffect CByte(B), SE_EXHAUST
                                                                                                            End If
                                                                                                    End Select
                                                                                                End If
                                                                                            Next i
                                                                                        End If
                                                                                        PrintCrashDebug 10, 9
                                                                                        If GetStatusEffect(B, SE_INVISIBLE) Then
                                                                                            RemoveStatusEffect B, SE_INVISIBLE
                                                                                        End If
                                                                                        
                                                                                        If C > 0 Or .ShieldBlock Then
                                                                                            If .ShieldBlock And Object(.Equipped(2).Object).Data(2) <= 100 Then
                                                                                                If E = 1 Then
                                                                                                    CreateSizedFloatingText mapNum, .x, .y, 12, "Block - " & CStr(C), critsize, 0
                                                                                                    'CreateFloatingText mapNum, .x, .y, 12, "Block, Critical Hit! " & CStr(C)
                                                                                                Else
                                                                                                    CreateFloatingText mapNum, .x, .y, 12, "Block - " & CStr(C)
                                                                                                End If
                                                                                            Else
                                                                                                If E = 1 Then
                                                                                                    CreateSizedFloatingText mapNum, .x, .y, 12, CStr(C), critsize, 0
                                                                                                    'CreateFloatingText mapNum, .x, .y, 12, "Critical Hit! " & CStr(C)
                                                                                                Else
                                                                                                    CreateFloatingText mapNum, .x, .y, 12, CStr(C)
                                                                                                End If

                                                                                            End If
                                                                                        Else
                                                                                            CreateFloatingEvent mapNum, .x, .y, FT_INEFFECTIVE
                                                                                        End If
                                                                                        .ShieldBlock = False
                                                                                    End If
                                                                                End With
                                                                                PrintCrashDebug 5, 15
                                                                                If (C = 0 And J = C) Or C <> 0 Then
                                                                                    If GetStatusEffect(B, SE_RETRIBUTION) Then
                                                                                        AttackMonster B, A, player(B).StatusData(SE_RETRIBUTION).Data(0), True, False, True, False
                                                                                        If .monster <= 0 Or .HP = 0 Then GoTo nexta
                                                                                    End If
                                                                                End If
                                                                                PrintCrashDebug 10, 11
                                                                                PrintCrashDebug 10, mapNum
                                                                                If .monster > 0 Then
                                                                                If monster(.monster).Flags2 And MONSTER_LARGE Then
                                                                                    If J = 0 Then
                                                                                        For i = 1 To currentMaxUser
                                                                                            If player(i).InUse Then
                                                                                                If i <> .Target Then
                                                                                                    If (player(i).map = mapNum) Then
                                                                                                        If .D = 0 Then
                                                                                                            If .x = player(i).x Or .x + 1 = player(i).x Then
                                                                                                                If .y - 1 = (player(i).y) Then
                                                                                                                    C = -1
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                        If .D = 1 Then
                                                                                                            If .x = player(i).x Or .x + 1 = player(i).x Then
                                                                                                                If .y + 2 = (player(i).y) Then
                                                                                                                    C = -1
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                        If .D = 2 Then
                                                                                                            If .y = player(i).y Or .y + 1 = player(i).y Then
                                                                                                                If .x - 1 = (player(i).x) Then
                                                                                                                    C = -1
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                        If .D = 0 Then
                                                                                                            If .y = player(i).y Or .y + 1 = player(i).y Then
                                                                                                                If .x + 2 = (player(i).x) Then
                                                                                                                    C = -1
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                        
                                                                                                        If C = -1 Then
                                                                                                            J = i
                                                                                                            GoTo nextAttack
                                                                                                        End If
                                                                                                        
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        Next i
                                                                                    End If
                                                                                End If
                                                                                End If
                                                                      PrintCrashDebug 5, 16
                                                                            Else
                                                                                SendToMap mapNum, Chr2(41) + Chr2(A)
                                                                                CreateFloatingEvent mapNum, player(B).x, player(B).y, FT_MISS
                                                                                
                                                                                If monster(.monster).Flags2 And MONSTER_LARGE Then
                                                                                    If J = 0 Then
                                                                                        For i = 1 To currentMaxUser
                                                                                            If player(i).InUse Then
                                                                                                If i <> .Target Then
                                                                                                    If (player(i).map = mapNum) Then
                                                                                                        If .D = 0 Then
                                                                                                            If .x = player(i).x Or .x + 1 = player(i).x Then
                                                                                                                If .y - 1 = (player(i).y) Then
                                                                                                                    C = -1
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                        If .D = 1 Then
                                                                                                            If .x = player(i).x Or .x + 1 = player(i).x Then
                                                                                                                If .y + 2 = (player(i).y) Then
                                                                                                                    C = -1
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                        If .D = 2 Then
                                                                                                            If .y = player(i).y Or .y + 1 = player(i).y Then
                                                                                                                If .x - 1 = (player(i).x) Then
                                                                                                                    C = -1
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                        If .D = 0 Then
                                                                                                            If .y = player(i).y Or .y + 1 = player(i).y Then
                                                                                                                If .x + 2 = (player(i).x) Then
                                                                                                                    C = -1
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                        
                                                                                                        If C = -1 Then
                                                                                                            J = i
                                                                                                            GoTo nextAttack
                                                                                                        End If
                                                                                                        
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        Next i
                                                                                    End If
                                                                                End If
                                                                                           PrintCrashDebug 5, 17
                                                                            End If
                                                                        End If
                                                                    Else
                                                                        '.AttackCounter = .AttackCounter - 1
                                                                        PrintCrashDebug 5, 18
                                                                    End If
                                                                End If
                                                            ElseIf .TargType = TargTypeMonster And .Target > 0 Then
                                                                B = .Target
                                                                If B = A Then
                                                                    .Target = 0
                                                                    .TargType = 0
                                                                    .MoveSpeed = monster(.monster).MoveSpeed
                                                                    .AttackSpeed = monster(.monster).AttackSpeed
                                                                Else
                                                                    If map(mapNum).monster(B).monster > 0 Then
                                                                        .MoveSpeed = .MoveSpeed
                                                                        C = .x
                                                                        D = .y
                                                                        E = .D
                                                                        
                                                                        Cont = Sqr(CSng(CLng(map(mapNum).monster(B).x) - C) ^ 2 + CSng(CLng(map(mapNum).monster(B).y) - D) ^ 2) <= 1
                                                                        If monster(.monster).Flags2 And MONSTER_LARGE And Cont = False Then
                                                                            Cont = Sqr(CSng(CLng(map(mapNum).monster(B).x) - (C + 1)) ^ 2 + CSng(CLng(map(mapNum).monster(B).y) - D) ^ 2) <= 1
                                                                            If Cont = False Then Cont = Sqr(CSng(CLng(map(mapNum).monster(B).x) - (C + 1)) ^ 2 + CSng(CLng(map(mapNum).monster(B).y) - (D + 1)) ^ 2) <= 1
                                                                            If Cont = False Then Cont = Sqr(CSng(CLng(map(mapNum).monster(B).x) - C) ^ 2 + CSng(CLng(map(mapNum).monster(B).y) - (D + 1)) ^ 2) <= 1
                                                                        End If
                                                                        PrintCrashDebug 5, 19
                                                                        If Cont Then
                                                                            'Attack Monster
                                                                            If .AttackCounter = 0 Then
                                                                                .AttackCounter = .AttackSpeed
                                                                                B = .Target
                                                                                C = map(mapNum).monster(B).x
                                                                                D = map(mapNum).monster(B).y
                                                                                
                                                                                If monster(.monster).Flags2 And MONSTER_LARGE Then
                                                                                    If .x < C And (.y = D Or .y = D + 1) Then
                                                                                        If .D <> 3 Then
                                                                                            .D = 3
                                                                                            ST1 = ST1 + DoubleChar(2) + Chr2(40) + Chr2(A * 16 + .D)
                                                                                            Call MonsterMove(mapNum, A)
                                                                                        End If
                                                                                    ElseIf .x > C And (.y = D Or .y = D + 1) Then
                                                                                        If .D <> 2 Then
                                                                                            .D = 2
                                                                                            ST1 = ST1 + DoubleChar(2) + Chr2(40) + Chr2(A * 16 + .D)
                                                                                            Call MonsterMove(mapNum, A)
                                                                                        End If
                                                                                    ElseIf .y < D And (.x = C Or .x + 1 = C) Then
                                                                                        If .D <> 1 Then
                                                                                            .D = 1
                                                                                            ST1 = ST1 + DoubleChar(2) + Chr2(40) + Chr2(A * 16 + .D)
                                                                                            Call MonsterMove(mapNum, A)
                                                                                        End If
                                                                                    ElseIf .y > D And (.x = C Or .x + 1 = C) Then
                                                                                        If .D <> 0 Then
                                                                                            .D = 0
                                                                                            ST1 = ST1 + DoubleChar(2) + Chr2(40) + Chr2(A * 16 + .D)
                                                                                            Call MonsterMove(mapNum, A)
                                                                                        End If
                                                                                    End If
                                                                                Else
                                                                                    If .x < C And .D <> 3 Then
                                                                                        .D = 3
                                                                                        ST1 = ST1 + DoubleChar(2) + Chr2(40) + Chr2(A * 16 + .D)
                                                                                        Call MonsterMove(mapNum, A)
                                                                                    ElseIf .x > C And .D <> 2 Then
                                                                                        .D = 2
                                                                                        ST1 = ST1 + DoubleChar(2) + Chr2(40) + Chr2(A * 16 + .D)
                                                                                        Call MonsterMove(mapNum, A)
                                                                                    ElseIf .y < D And .D <> 1 Then
                                                                                        .D = 1
                                                                                        ST1 = ST1 + DoubleChar(2) + Chr2(40) + Chr2(A * 16 + .D)
                                                                                        Call MonsterMove(mapNum, A)
                                                                                    ElseIf .y > D And .D <> 0 Then
                                                                                        .D = 0
                                                                                        ST1 = ST1 + DoubleChar(2) + Chr2(40) + Chr2(A * 16 + .D)
                                                                                        Call MonsterMove(mapNum, A)
                                                                                    End If
                                                                                End If
                                                                                If .monster = 0 Then GoTo nexta
                                                                                If Int(Rnd * 255) > (monster(map(mapNum).monster(B).monster).Agility / 2) Then
                                                                                    C = Int(Rnd * (monster(.monster).Max - monster(.monster).Min)) + monster(.monster).Min
                                                                                    C = C - monster(map(mapNum).monster(B).monster).Armor
                                                                                    If C < 0 Then C = 0
                                                                                    If C > 9999 Then C = 9999
                                                                                    ST1 = ST1 + DoubleChar(2) + Chr2(41) + Chr2(A)
                                                                                    With map(mapNum).monster(B)
                                                                                        If C > 0 Then
                                                                                            CreateFloatingText mapNum, .x, .y, 12, CStr(C)
                                                                                        Else
                                                                                            CreateFloatingEvent mapNum, .x, .y, FT_INEFFECTIVE
                                                                                        End If
                                                                                        If C >= .HP Then
                                                                                            'monster Died
                                                                                            Parameter(0) = A
                                                                                            Parameter(1) = 1
                                                                                            Parameter(2) = B
                                                                                            Parameter(3) = mapNum
                                                                                            RunScript "MONSTERDIE" + CStr(.monster)
                                                                                            
                                                                                            RunScript "MONSTERDIE"
                                                                                            
                                                                                            .monster = 0
                                                                                            .Target = 0
                                                                                            .TargType = 0
                                                                                            
                                                                                            ST1 = ST1 + DoubleChar(2) + Chr2(39) + Chr2(B)
                                                                                        Else
                                                                                            .HP = .HP - C
                                                                                            .Target = A
                                                                                            .TargType = TargTypeMonster
                                                                                            SendToMap mapNum, Chr2(132) + Chr2(B) + DoubleChar(.HP)
                                                                                        End If
                                                                                    End With
                                                                                Else
                                                                                    CreateFloatingEvent mapNum, map(mapNum).monster(B).x, map(mapNum).monster(B).y, FT_MISS
                                                                                End If
                                                                            Else
                                                                                '.AttackCounter = 0
                                                                            End If
                                                                        End If
                                                                        PrintCrashDebug 5, 20
                                                                    Else
                                                                        .Target = 0
                                                                        .TargType = 0
                                                                        .MoveSpeed = monster(.monster).MoveSpeed
                                                                        .AttackSpeed = monster(.monster).AttackSpeed
                                                                    End If
                                                                End If
                                                            End If
                                                        Else
                                                        PrintCrashDebug 5, 21
                                                            'Fix Large Monsters
                                                            If monster(.monster).Flags2 And MONSTER_LARGE Then
                                                            
                                                                B = .Target
                                                                C = .x
                                                                D = .y
                                                                If B > 0 Then
                                                                    'close enough to attack
                                                                    Cont = Sqr(CSng(CLng(player(B).x) - C) ^ 2 + CSng(CLng(player(B).y) - D) ^ 2) <= 1
                                                                    If Cont = False And monster(.monster).Flags2 And MONSTER_LARGE Then
                                                                        Cont = Sqr(CSng(CLng(player(B).x) - (C + 1)) ^ 2 + CSng(CLng(player(B).y) - (D)) ^ 2) <= 1
                                                                        If Cont = False Then Cont = Sqr(CSng(CLng(player(B).x) - (C + 1)) ^ 2 + CSng(CLng(player(B).y) - (D + 1)) ^ 2) <= 1
                                                                        If Cont = False Then Cont = Sqr(CSng(CLng(player(B).x) - (C)) ^ 2 + CSng(CLng(player(B).y) - (D + 1)) ^ 2) <= 1
                                                                    End If
                                                                    If Cont Then
                                                                        Cont = CanAttack(CLng(.x), CLng(.y), CLng(player(B).x), CLng(player(B).y), player(B).map)
                                                                        If Cont = False And monster(.monster).Flags2 And MONSTER_LARGE Then
                                                                            Cont = CanAttack(CLng(.x + 1), CLng(.y), CLng(player(B).x), CLng(player(B).y), player(B).map)
                                                                            If Cont = False Then Cont = CanAttack(CLng(.x + 1), CLng(.y + 1), CLng(player(B).x), CLng(player(B).y), player(B).map)
                                                                            If Cont = False Then Cont = CanAttack(CLng(.x), CLng(.y + 1), CLng(player(B).x), CLng(player(B).y), player(B).map)
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                            If .AttackCounter > 0 Then .AttackCounter = .AttackCounter - 1
                                                        End If
                                                        PrintCrashDebug 5, 22
                                                        If .MoveCounter = 0 And Cont = False Then
                                                            .MoveCounter = .MoveSpeed
    
                                                            If .Target = 0 And .TargType = 0 Then
                                                                'If .MoveSpeed < 2 Then .MoveSpeed = 2
                                                                'Random Movement
                                                                PrintCrashDebug 5, 23
                                                                If Rnd < (monster(.monster).Wander / 100) Then
                                                                    If monster(.monster).Flags2 And MONSTER_LARGE Then
                                                                        .D = Int(Rnd * 4)
                                                                        PrintCrashDebug 5, 24
                                                                        Select Case .D
                                                                            Case 0 'Up
                                                                                If .y > 0 Then
                                                                                    If IsVacant(mapNum, .x, .y - 1, .D) = 1 And IsVacant(mapNum, .x + 1, .y - 1, .D) = 1 Then
                                                                                        .y = .y - 1
                                                                                        ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                                        Call MonsterMove(mapNum, A)
                                                                                    End If
                                                                                End If
                                                                            Case 1 'Down
                                                                                If .y < 10 Then
                                                                                    If IsVacant(mapNum, .x, .y + 2, .D) = 1 And IsVacant(mapNum, .x + 1, .y + 2, .D) = 1 Then
                                                                                        .y = .y + 1
                                                                                        ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                                        Call MonsterMove(mapNum, A)
                                                                                    End If
                                                                                End If
                                                                            Case 2 'Left
                                                                                If .x > 0 Then
                                                                                    If IsVacant(mapNum, .x - 1, .y, .D) = 1 And IsVacant(mapNum, .x - 1, .y + 1, .D) = 1 Then
                                                                                        .x = .x - 1
                                                                                        ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                                        Call MonsterMove(mapNum, A)
                                                                                    End If
                                                                                End If
                                                                            Case 3 'Right
                                                                                If .x < 10 Then
                                                                                    If IsVacant(mapNum, .x + 2, .y, .D) = 1 And IsVacant(mapNum, .x + 2, .y + 1, .D) = 1 Then
                                                                                        .x = .x + 1
                                                                                        ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                                        Call MonsterMove(mapNum, A)
                                                                                    End If
                                                                                End If
                                                                        End Select
                                                                        PrintCrashDebug 5, 25
                                                                    Else
                                                                        .D = Int(Rnd * 4)
                                                                        Select Case .D
                                                                            Case 0 'Up
                                                                                If .y > 0 Then
                                                                                    If IsVacant(mapNum, .x, .y - 1, .D) = 1 Then
                                                                                        .y = .y - 1
                                                                                        ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                                        Call MonsterMove(mapNum, A)
                                                                                    End If
                                                                                End If
                                                                            Case 1 'Down
                                                                                If .y < 11 Then
                                                                                    If IsVacant(mapNum, .x, .y + 1, .D) = 1 Then
                                                                                        .y = .y + 1
                                                                                        ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                                        Call MonsterMove(mapNum, A)
                                                                                    End If
                                                                                End If
                                                                            Case 2 'Left
                                                                                If .x > 0 Then
                                                                                    If IsVacant(mapNum, .x - 1, .y, .D) = 1 Then
                                                                                        .x = .x - 1
                                                                                        ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                                        Call MonsterMove(mapNum, A)
                                                                                    End If
                                                                                End If
                                                                            Case 3 'Right
                                                                                If .x < 11 Then
                                                                                    If IsVacant(mapNum, .x + 1, .y, .D) = 1 Then
                                                                                        .x = .x + 1
                                                                                        ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                                        Call MonsterMove(mapNum, A)
                                                                                    End If
                                                                                End If
                                                                        End Select
                                                                        PrintCrashDebug 5, 26
                                                                    End If
                                                                    If .monster = 0 Then GoTo nexta
                                                                End If
                                                            ElseIf .TargType = TargTypePlayer And .Target > 0 Then
                                                                'Move Toward Target
                                                                '.AttackCounter = 0
                                                                B = .Target
                                                                If player(B).Mode = modePlaying And player(B).map = mapNum Then
                                                                    C = .x
                                                                    D = .y
                                                                    E = .D
                                                                    PrintCrashDebug 5, 27
                                                                    Cont = Sqr(CSng(CLng(player(B).x) - C) ^ 2 + CSng(CLng(player(B).y) - D) ^ 2) > 1
                                                                    If Cont = False And monster(.monster).Flags2 And MONSTER_LARGE Then
                                                                        Cont = Sqr(CSng(CLng(player(B).x) - (C + 1)) ^ 2 + CSng(CLng(player(B).y) - (D)) ^ 2) > 1
                                                                        If Cont = False Then Cont = Sqr(CSng(CLng(player(B).x) - (C + 1)) ^ 2 + CSng(CLng(player(B).y) - (D + 1)) ^ 2) > 1
                                                                        If Cont = False Then Cont = Sqr(CSng(CLng(player(B).x) - (C)) ^ 2 + CSng(CLng(player(B).y) - (D + 1)) ^ 2) > 1
                                                                    End If
                                                                    
                                                                    If Cont Then
                                                                    PrintCrashDebug 5, 28
                                                                        If Rnd < 0.5 Then
                                                                            If C < player(B).x Then
                                                                                F = IsVacant(mapNum, C + 1, CByte(D), 3, A)
                                                                                'If C < 10 And Monster(.Monster).Flags2 And MONSTER_LARGE Then F = IsVacant(MapNum, C + 2, CByte(D), 3) And IsVacant(MapNum, C + 2, CByte(D) + 1, 3, .Monster)
                                                                                If C >= 10 And monster(.monster).Flags2 And MONSTER_LARGE Then F = 0
                                                                                If F = 1 Then
                                                                                    C = C + 1
                                                                                    E = 3
                                                                                ElseIf F > 1 Then
                                                                                    If ((monster(.monster).Flags And MONSTER_FRIENDLY) = 0 And (monster(.monster).Flags And MONSTER_GUARD) = 0) Then
                                                                                        F = F - 1
                                                                                        .Target = F
                                                                                        .TargType = TargTypePlayer
                                                                                    End If
                                                                                End If
                                                                            ElseIf C > player(B).x Then
                                                                                F = IsVacant(mapNum, C - 1, CByte(D), 2, A)
                                                                                'If Monster(.Monster).Flags2 And MONSTER_LARGE Then F = IsVacant(MapNum, C - 1, CByte(D), 2) And IsVacant(MapNum, C - 1, CByte(D) + 1, 2, .Monster)
                                                                                If F = 1 Then
                                                                                    C = C - 1
                                                                                    E = 2
                                                                                ElseIf F > 1 Then
                                                                                    If ((monster(.monster).Flags And MONSTER_FRIENDLY) = 0 And (monster(.monster).Flags And MONSTER_GUARD) = 0) Then
                                                                                        F = F - 1
                                                                                        .Target = F
                                                                                        .TargType = TargTypePlayer
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                            If C = .x And D = .y Then
                                                                                If D < player(B).y Then
                                                                                    F = IsVacant(mapNum, CByte(C), D + 1, 1, A)
                                                                                    'If D < 10 And Monster(.Monster).Flags2 And MONSTER_LARGE Then F = IsVacant(MapNum, CByte(C), D + 2, 1) And IsVacant(MapNum, CByte(C) + 1, D + 2, 1, .Monster)
                                                                                    If D >= 10 And monster(.monster).Flags2 And MONSTER_LARGE Then F = 0
                                                                                    If F = 1 Then
                                                                                        D = D + 1
                                                                                        E = 1
                                                                                    ElseIf F > 1 Then
                                                                                        If ((monster(.monster).Flags And MONSTER_FRIENDLY) = 0 And (monster(.monster).Flags And MONSTER_GUARD) = 0) Then
                                                                                            F = F - 1
                                                                                            .Target = F
                                                                                            .TargType = TargTypePlayer
                                                                                        End If
                                                                                    ElseIf Rnd < 0.2 Then
                                                                                        If Rnd < 0.5 Then
                                                                                            If C > 0 Then
                                                                                                F = IsVacant(mapNum, C - 1, CByte(D), 2, A)
                                                                                                'If Monster(.Monster).Flags2 And MONSTER_LARGE Then F = IsVacant(MapNum, C - 1, CByte(D), 2) And IsVacant(MapNum, C - 1, CByte(D) + 1, 2, .Monster)
                                                                                                If F = 1 Then
                                                                                                    C = C - 1
                                                                                                    E = 2
                                                                                                ElseIf F > 1 Then
                                                                                                    If ((monster(.monster).Flags And MONSTER_FRIENDLY) = 0 And (monster(.monster).Flags And MONSTER_GUARD) = 0) Then
                                                                                                        F = F - 1
                                                                                                        .Target = F
                                                                                                        .TargType = TargTypePlayer
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        Else
                                                                                            If C < 11 Then
                                                                                                F = IsVacant(mapNum, C + 1, CByte(D), 3, A)
                                                                                                'If C < 10 And Monster(.Monster).Flags2 And MONSTER_LARGE Then F = IsVacant(MapNum, C + 2, CByte(D), 3) And IsVacant(MapNum, C + 2, CByte(D) + 1, 3, .Monster)
                                                                                                If C >= 10 And monster(.monster).Flags2 And MONSTER_LARGE Then F = 0
                                                                                                If F = 1 Then
                                                                                                    C = C + 1
                                                                                                    E = 3
                                                                                                ElseIf F > 1 Then
                                                                                                    If ((monster(.monster).Flags And MONSTER_FRIENDLY) = 0 And (monster(.monster).Flags And MONSTER_GUARD) = 0) Then
                                                                                                        F = F - 1
                                                                                                        .Target = F
                                                                                                        .TargType = TargTypePlayer
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                ElseIf D > player(B).y Then
                                                                                    F = IsVacant(mapNum, CByte(C), D - 1, 0, A)
                                                                                    'If Monster(.Monster).Flags2 And MONSTER_LARGE Then F = IsVacant(MapNum, CByte(C), D - 1, 0) And IsVacant(MapNum, CByte(C) + 1, D - 1, 0, .Monster)
                                                                                    If F = 1 Then
                                                                                        D = D - 1
                                                                                        E = 0
                                                                                    ElseIf F > 1 Then
                                                                                        If ((monster(.monster).Flags And MONSTER_FRIENDLY) = 0 And (monster(.monster).Flags And MONSTER_GUARD) = 0) Then
                                                                                            F = F - 1
                                                                                            .Target = F
                                                                                            .TargType = TargTypePlayer
                                                                                        End If
                                                                                    ElseIf Rnd < 0.2 Then
                                                                                        If Rnd < 0.5 Then
                                                                                            If C > 0 Then
                                                                                                F = IsVacant(mapNum, C - 1, CByte(D), 2, A)
                                                                                                'If Monster(.Monster).Flags2 And MONSTER_LARGE Then F = IsVacant(MapNum, C - 1, CByte(D) + 1, 2) And IsVacant(MapNum, C - 1, CByte(D), 2, .Monster)
                                                                                                If F = 1 Then
                                                                                                    C = C - 1
                                                                                                    E = 2
                                                                                                ElseIf F > 1 Then
                                                                                                    If ((monster(.monster).Flags And MONSTER_FRIENDLY) = 0 And (monster(.monster).Flags And MONSTER_GUARD) = 0) Then
                                                                                                        F = F - 1
                                                                                                        .Target = F
                                                                                                        .TargType = TargTypePlayer
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        Else
                                                                                            If C < 11 Then
                                                                                                F = IsVacant(mapNum, C + 1, CByte(D), 3, A)
                                                                                                'If C < 10 And Monster(.Monster).Flags2 And MONSTER_LARGE Then F = IsVacant(MapNum, C + 2, CByte(D), 3) And IsVacant(MapNum, C + 2, CByte(D) + 1, 3, .Monster)
                                                                                                If C >= 10 And monster(.monster).Flags2 And MONSTER_LARGE Then F = 0
                                                                                                If F = 1 Then
                                                                                                    C = C + 1
                                                                                                    E = 3
                                                                                                ElseIf F > 1 Then
                                                                                                    If ((monster(.monster).Flags And MONSTER_FRIENDLY) = 0 And (monster(.monster).Flags And MONSTER_GUARD) = 0) Then
                                                                                                        F = F - 1
                                                                                                        .Target = F
                                                                                                        .TargType = TargTypePlayer
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                            PrintCrashDebug 5, 29
                                                                        Else
                                                                        PrintCrashDebug 5, 30
                                                                            If D < player(B).y Then
                                                                                F = IsVacant(mapNum, CByte(C), D + 1, 1, A)
                                                                                'If D < 10 And Monster(.Monster).Flags2 And MONSTER_LARGE Then F = IsVacant(MapNum, CByte(C), D + 2, 1) And IsVacant(MapNum, CByte(C) + 1, D + 2, 1, .Monster)
                                                                                If D >= 10 And monster(.monster).Flags2 And MONSTER_LARGE Then F = 0
                                                                                If F = 1 Then
                                                                                    D = D + 1
                                                                                    E = 1
                                                                                ElseIf F > 1 Then
                                                                                    If ((monster(.monster).Flags And MONSTER_FRIENDLY) = 0 And (monster(.monster).Flags And MONSTER_GUARD) = 0) Then
                                                                                        F = F - 1
                                                                                        .Target = F
                                                                                        .TargType = TargTypePlayer
                                                                                    End If
                                                                                End If
                                                                            ElseIf D > player(B).y Then
                                                                                F = IsVacant(mapNum, CByte(C), D - 1, 0, A)
                                                                                'If Monster(.Monster).Flags2 And MONSTER_LARGE Then F = IsVacant(MapNum, CByte(C), D - 1, 0) And IsVacant(MapNum, CByte(C) + 1, D - 1, 0, .Monster)
                                                                                If F = 1 Then
                                                                                    D = D - 1
                                                                                    E = 0
                                                                                ElseIf F > 1 Then
                                                                                    If ((monster(.monster).Flags And MONSTER_FRIENDLY) = 0 And (monster(.monster).Flags And MONSTER_GUARD) = 0) Then
                                                                                        F = F - 1
                                                                                        .Target = F
                                                                                        .TargType = TargTypePlayer
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                            If C = .x And D = .y Then
                                                                                If C < player(B).x Then
                                                                                    F = IsVacant(mapNum, C + 1, CByte(D), 3, A)
                                                                                    'If C < 10 And Monster(.Monster).Flags2 And MONSTER_LARGE Then F = IsVacant(MapNum, C + 2, CByte(D) + 1, 3) And IsVacant(MapNum, C + 2, CByte(D), 3, .Monster)
                                                                                    If C >= 10 And monster(.monster).Flags2 And MONSTER_LARGE Then F = 0
                                                                                    If F = 1 Then
                                                                                        C = C + 1
                                                                                        E = 3
                                                                                    ElseIf F > 1 Then
                                                                                        If ((monster(.monster).Flags And MONSTER_FRIENDLY) = 0 And (monster(.monster).Flags And MONSTER_GUARD) = 0) Then
                                                                                            F = F - 1
                                                                                            .Target = F
                                                                                            .TargType = TargTypePlayer
                                                                                        End If
                                                                                    ElseIf Rnd < 0.2 Then
                                                                                        If Rnd < 0.5 Then
                                                                                            If D > 0 Then
                                                                                                F = IsVacant(mapNum, CByte(C), D - 1, 0, A)
                                                                                                'If Monster(.Monster).Flags2 And MONSTER_LARGE Then F = IsVacant(MapNum, CByte(C) + 1, D - 1, 0) And IsVacant(MapNum, CByte(C), D - 1, 0, .Monster)
                                                                                                If F = 1 Then
                                                                                                    D = D - 1
                                                                                                    E = 0
                                                                                                ElseIf F > 1 Then
                                                                                                    If ((monster(.monster).Flags And MONSTER_FRIENDLY) = 0 And (monster(.monster).Flags And MONSTER_GUARD) = 0) Then
                                                                                                        F = F - 1
                                                                                                        .Target = F
                                                                                                        .TargType = TargTypePlayer
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        Else
                                                                                            If D < 11 Then
                                                                                                F = IsVacant(mapNum, CByte(C), D + 1, 1, A)
                                                                                                'If D < 10 And Monster(.Monster).Flags2 And MONSTER_LARGE Then F = IsVacant(MapNum, CByte(C), D + 2, 1) And IsVacant(MapNum, CByte(C) + 1, D + 2, 1, .Monster)
                                                                                                If D >= 10 And monster(.monster).Flags2 And MONSTER_LARGE Then F = 0
                                                                                                If F = 1 Then
                                                                                                    D = D + 1
                                                                                                    E = 1
                                                                                                ElseIf F > 1 Then
                                                                                                    If ((monster(.monster).Flags And MONSTER_FRIENDLY) = 0 And (monster(.monster).Flags And MONSTER_GUARD) = 0) Then
                                                                                                        F = F - 1
                                                                                                        .Target = F
                                                                                                        .TargType = TargTypePlayer
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                ElseIf C > player(B).x Then
                                                                                    F = IsVacant(mapNum, C - 1, CByte(D), 2, A)
                                                                                    'If Monster(.Monster).Flags2 And MONSTER_LARGE Then F = IsVacant(MapNum, C - 1, CByte(D) + 1, 2) And IsVacant(MapNum, C - 1, CByte(D), 2, .Monster)
                                                                                    If F = 1 Then
                                                                                        C = C - 1
                                                                                        E = 2
                                                                                    ElseIf F > 1 Then
                                                                                        If ((monster(.monster).Flags And MONSTER_FRIENDLY) = 0 And (monster(.monster).Flags And MONSTER_GUARD) = 0) Then
                                                                                            F = F - 1
                                                                                            .Target = F
                                                                                            .TargType = TargTypePlayer
                                                                                        End If
                                                                                    ElseIf Rnd < 0.2 Then
                                                                                        If Rnd < 0.5 Then
                                                                                            If D > 0 Then
                                                                                                F = IsVacant(mapNum, CByte(C), D - 1, 0, A)
                                                                                                'If Monster(.Monster).Flags2 And MONSTER_LARGE Then F = IsVacant(MapNum, CByte(C), D - 1, 0) And IsVacant(MapNum, CByte(C) + 1, D - 1, 0, .Monster)
                                                                                                If F = 1 Then
                                                                                                    D = D - 1
                                                                                                    E = 0
                                                                                                ElseIf F > 1 Then
                                                                                                    If ((monster(.monster).Flags And MONSTER_FRIENDLY) = 0 And (monster(.monster).Flags And MONSTER_GUARD) = 0) Then
                                                                                                        F = F - 1
                                                                                                        .Target = F
                                                                                                        .TargType = TargTypePlayer
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        Else
                                                                                            If D < 11 Then
                                                                                                F = IsVacant(mapNum, CByte(C), D + 1, 1, A)
                                                                                                'If D < 10 And Monster(.Monster).Flags2 And MONSTER_LARGE Then F = IsVacant(MapNum, CByte(C), D + 2, 1) And IsVacant(MapNum, CByte(C) + 1, D + 2, 1, .Monster)
                                                                                                If D >= 10 And monster(.monster).Flags2 And MONSTER_LARGE Then F = 0
                                                                                                If F = 1 Then
                                                                                                    D = D + 1
                                                                                                    E = 1
                                                                                                ElseIf F > 1 Then
                                                                                                    If ((monster(.monster).Flags And MONSTER_FRIENDLY) = 0 And (monster(.monster).Flags And MONSTER_GUARD) = 0) Then
                                                                                                        F = F - 1
                                                                                                        .Target = F
                                                                                                        .TargType = TargTypePlayer
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                        PrintCrashDebug 5, 34
                                                                        If C <> .x Or D <> .y Or E <> .D Then
                                                                            .x = C
                                                                            .y = D
                                                                            .D = E
                                                                            ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(C * 16 + D) + Chr2(.MoveSpeed)
                                                                            Call MonsterMove(mapNum, A)
                                                                            If .monster = 0 Then GoTo nexta
                                                                        End If
                                                                    'Attack Code Went Here
                                                                    End If
                                                                Else
                                                                    .Target = 0
                                                                    .TargType = 0
                                                                    .Distance = monster(.monster).Sight
                                                                    .MoveSpeed = monster(.monster).MoveSpeed
                                                                    .AttackSpeed = monster(.monster).AttackSpeed
                                                                End If
                                                                PrintCrashDebug 5, 35
                                                            ElseIf .TargType = TargTypeMonster Then
                                                                'Move Toward Target
                                                                B = .Target
                                                                If B = A Then
                                                                    .Target = 0
                                                                    .TargType = 0
                                                                    .MoveSpeed = monster(.monster).MoveSpeed
                                                                    .AttackSpeed = monster(.monster).AttackSpeed
                                                                Else
                                                                PrintCrashDebug 5, 36
                                                                    If map(mapNum).monster(B).monster > 0 Then
                                                                        C = .x
                                                                        D = .y
                                                                        E = .D
                                                                        If Sqr(CSng(CLng(map(mapNum).monster(B).x) - C) ^ 2 + CSng(CLng(map(mapNum).monster(B).y) - D) ^ 2) > 1 Then
                                                                             If (monster(.monster).Flags2 And MONSTER_LARGE) Then 'i'm sorry this is so horrible, blame vb6 because doing this like normal program flow gave "expression too complex" crash error
                                                                                If Sqr(CSng(CLng(map(mapNum).monster(B).x) - (C + 1)) ^ 2 + CSng(CLng(map(mapNum).monster(B).y) - D) ^ 2) > 1 Then
                                                                                    If Sqr(CSng(CLng(map(mapNum).monster(B).x) - (C + 1)) ^ 2 + CSng(CLng(map(mapNum).monster(B).y) - (D + 1)) ^ 2) > 1 Then
                                                                                        If Sqr(CSng(CLng(map(mapNum).monster(B).x) - C) ^ 2 + CSng(CLng(map(mapNum).monster(B).y) - (D + 1)) ^ 2) > 1 Then
                                                                                        Else
                                                                                            GoTo skipmovetotarg
                                                                                        End If
                                                                                    Else
                                                                                        GoTo skipmovetotarg
                                                                                    End If
                                                                                Else
                                                                                    GoTo skipmovetotarg
                                                                                End If
                                                                            End If
                                                                            PrintCrashDebug 5, 37
                                                                            '.AttackCounter = 0
                                                                            If Rnd < 0.5 Then
                                                                                If C < map(mapNum).monster(B).x Then
                                                                                    If IsVacant(mapNum, C + 1, CByte(D), 3, A) = 1 Then
                                                                                        C = C + 1
                                                                                        E = 3
                                                                                        GoTo skipmovetotarg
                                                                                    End If
                                                                                ElseIf C > map(mapNum).monster(B).x Then
                                                                                    If IsVacant(mapNum, C - 1, CByte(D), 2, A) = 1 Then
                                                                                        C = C - 1
                                                                                        E = 2
                                                                                        GoTo skipmovetotarg
                                                                                    End If
                                                                                End If
                                                                                If C = .x And D = .y Then
                                                                                    If D < map(mapNum).monster(B).y Then
                                                                                        If IsVacant(mapNum, CByte(C), D + 1, 1, A) = 1 Then
                                                                                            D = D + 1
                                                                                            E = 1
                                                                                        ElseIf Rnd < 0.2 Then
                                                                                            If Rnd < 0.5 Then
                                                                                                If C > 0 Then
                                                                                                    If IsVacant(mapNum, C - 1, CByte(D), 2, A) = 1 Then
                                                                                                        C = C - 1
                                                                                                        E = 2
                                                                                                    End If
                                                                                                End If
                                                                                            Else
                                                                                                If C < 11 Then
                                                                                                    If IsVacant(mapNum, C + 1, CByte(D), 3, A) = 1 Then
                                                                                                        C = C + 1
                                                                                                        E = 3
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    ElseIf D > map(mapNum).monster(B).y Then
                                                                                        If IsVacant(mapNum, CByte(C), D - 1, 0, A) = 1 Then
                                                                                            D = D - 1
                                                                                            E = 0
                                                                                        ElseIf Rnd < 0.2 Then
                                                                                            If Rnd < 0.5 Then
                                                                                                If C > 0 Then
                                                                                                    If IsVacant(mapNum, C - 1, CByte(D), 2, A) = 1 Then
                                                                                                        C = C - 1
                                                                                                        E = 2
                                                                                                    End If
                                                                                                End If
                                                                                            Else
                                                                                                If C < 11 Then
                                                                                                    If IsVacant(mapNum, C + 1, CByte(D), 3, A) = 1 Then
                                                                                                        C = C + 1
                                                                                                        E = 3
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                Else
                                                                                PrintCrashDebug 5, 38
                                                                                    If D < map(mapNum).monster(B).y Then
                                                                                        If IsVacant(mapNum, CByte(C), D + 1, 1, A) = 1 Then
                                                                                            D = D + 1
                                                                                            E = 1
                                                                                            GoTo skipmovetotarg
                                                                                        End If
                                                                                    ElseIf D > map(mapNum).monster(B).y Then
                                                                                        If IsVacant(mapNum, CByte(C), D - 1, 0, A) = 1 Then
                                                                                            D = D - 1
                                                                                            E = 0
                                                                                            GoTo skipmovetotarg
                                                                                        End If
                                                                                    End If
                                                                                    PrintCrashDebug 5, 39
                                                                                    If C = .x And D = .y Then
                                                                                        If C < map(mapNum).monster(B).x Then
                                                                                            If IsVacant(mapNum, C + 1, CByte(D), 3, A) = 1 Then
                                                                                                C = C + 1
                                                                                                E = 3
                                                                                            ElseIf Rnd < 0.2 Then
                                                                                                If Rnd < 0.5 Then
                                                                                                    If D > 0 Then
                                                                                                        If IsVacant(mapNum, CByte(C), D - 1, 0, A) = 1 Then
                                                                                                            D = D - 1
                                                                                                            E = 0
                                                                                                        End If
                                                                                                    End If
                                                                                                Else
                                                                                                    If D < 11 Then
                                                                                                        If IsVacant(mapNum, CByte(C), D + 1, 1, A) = 1 Then
                                                                                                            D = D + 1
                                                                                                            E = 1
                                                                                                        End If
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                            PrintCrashDebug 5, 40
                                                                                        ElseIf C > map(mapNum).monster(B).x Then
                                                                                            If IsVacant(mapNum, C - 1, CByte(D), 2) = 1 Then
                                                                                                C = C - 1
                                                                                                E = 2
                                                                                            ElseIf Rnd < 0.2 Then
                                                                                                If Rnd < 0.5 Then
                                                                                                    If D > 0 Then
                                                                                                        If IsVacant(mapNum, CByte(C), D - 1, 0, A) = 1 Then
                                                                                                            D = D - 1
                                                                                                            E = 0
                                                                                                        End If
                                                                                                    End If
                                                                                                Else
                                                                                                    If D < 11 Then
                                                                                                        If IsVacant(mapNum, CByte(C), D + 1, 1, A) = 1 Then
                                                                                                            D = D + 1
                                                                                                            E = 1
                                                                                                        End If
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
skipmovetotarg:
PrintCrashDebug 5, 41
                                                                                If C <> .x Or D <> .y Or E <> .D Then
                                                                                    .x = C
                                                                                    .y = D
                                                                                    .D = E
                                                                                    ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(C * 16 + D) + Chr2(.MoveSpeed)
                                                                                    Call MonsterMove(mapNum, A)
                                                                                    If .monster = 0 Then GoTo nexta
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    Else
                                                                        .Target = 0
                                                                        .TargType = 0
                                                                        .Distance = monster(.monster).Sight
                                                                        .MoveSpeed = monster(.monster).MoveSpeed
                                                                        .AttackSpeed = monster(.monster).AttackSpeed
                                                                    End If
                                                                End If
                                                            Else
                                                                .Target = 0
                                                                .TargType = 0
                                                                .MoveSpeed = monster(.monster).MoveSpeed
                                                                .AttackSpeed = monster(.monster).AttackSpeed
                                                            End If
                                                        Else
                                                            If .MoveCounter > 0 Then .MoveCounter = .MoveCounter - 1
                                                        End If
                                                    Else
                                                    PrintCrashDebug 5, 42
                                                        Dim ShiftQueue As Boolean
                                                        Select Case .MonsterQueue(0).Action
                                                            Case QUEUE_TURN
                                                                If .MoveCounter = 0 Then
                                                                    .MoveCounter = .MoveSpeed
                                                                    Select Case .MonsterQueue(0).lngData
                                                                        Case 0
                                                                            .D = 0
                                                                            ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                        Case 1
                                                                            .D = 1
                                                                            ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                        Case 2
                                                                            .D = 2
                                                                            ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                        Case 3
                                                                            .D = 3
                                                                            ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                    End Select
                                                                    If .monster = 0 Then GoTo nexta
                                                                    PrintCrashDebug 5, 43
                                                                    ShiftQueue = True
                                                                Else
                                                                    ShiftQueue = False
                                                                    If .MoveCounter > 0 Then .MoveCounter = .MoveCounter - 1
                                                                End If
                                                            Case QUEUE_SHIFT
                                                                If .MoveCounter = 0 Then
                                                                    .MoveCounter = .MoveSpeed
                                                                    Select Case .MonsterQueue(0).lngData
                                                                        Case 0
                                                                            '.D = 0
                                                                            If .y > 0 Then
                                                                                If IsVacant(mapNum, .x, .y - 1, 0) Then
                                                                                    If .y > 0 Then
                                                                                        .y = .y - 1
                                                                                    End If
                                                                                    MonsterMove mapNum, A
                                                                                End If
                                                                            End If
                                                                            ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                        Case 1
                                                                            '.D = 1
                                                                            If .y < 11 Then
                                                                                If IsVacant(mapNum, .x, .y + 1, 1) Then
                                                                                    .y = .y + 1
                                                                                    MonsterMove mapNum, A
                                                                                End If
                                                                            End If
                                                                            ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                        Case 2
                                                                            '.D = 2
                                                                            If .x > 0 Then
                                                                                If IsVacant(mapNum, .x - 1, .y, 2) Then
                                                                                    .x = .x - 1
                                                                                    MonsterMove mapNum, A
                                                                                End If
                                                                            End If
                                                                            ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                        Case 3
                                                                            '.D = 3
                                                                            If .x < 11 Then
                                                                                If IsVacant(mapNum, .x + 1, .y, 3) Then
                                                                                    .x = .x + 1
                                                                                    MonsterMove mapNum, A
                                                                                End If
                                                                            End If
                                                                            ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                    End Select
                                                                    If .monster = 0 Then GoTo nexta
                                                                PrintCrashDebug 5, 43
                                                                    ShiftQueue = True
                                                                Else
                                                                    ShiftQueue = False
                                                                    If .MoveCounter > 0 Then .MoveCounter = .MoveCounter - 1
                                                                End If
                                                            Case QUEUE_MOVE
                                                                If .MoveCounter = 0 Then
                                                                    .MoveCounter = .MoveSpeed
                                                                    Select Case .MonsterQueue(0).lngData
                                                                        Case 0
                                                                            .D = 0
                                                                            If .y > 0 Then
                                                                                If IsVacant(mapNum, .x, .y - 1, 0) Then
                                                                                    If .y > 0 Then
                                                                                        .y = .y - 1
                                                                                    End If
                                                                                    MonsterMove mapNum, A
                                                                                End If
                                                                            End If
                                                                            ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                        Case 1
                                                                            .D = 1
                                                                            If .y < 11 Then
                                                                                If IsVacant(mapNum, .x, .y + 1, 1) Then
                                                                                    .y = .y + 1
                                                                                    MonsterMove mapNum, A
                                                                                End If
                                                                            End If
                                                                            ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                        Case 2
                                                                            .D = 2
                                                                            If .x > 0 Then
                                                                                If IsVacant(mapNum, .x - 1, .y, 2) Then
                                                                                    .x = .x - 1
                                                                                    MonsterMove mapNum, A
                                                                                End If
                                                                            End If
                                                                            ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                        Case 3
                                                                            .D = 3
                                                                            If .x < 11 Then
                                                                                If IsVacant(mapNum, .x + 1, .y, 3) Then
                                                                                    .x = .x + 1
                                                                                    MonsterMove mapNum, A
                                                                                End If
                                                                            End If
                                                                            ST1 = ST1 + DoubleChar(4) + Chr2(40) + Chr2(A * 16 + (.D * 2)) + Chr2(.x * 16 + .y) + Chr2(.MoveSpeed)
                                                                    End Select
                                                                    If .monster = 0 Then GoTo nexta
                                                                    PrintCrashDebug 5, 43
                                                                    ShiftQueue = True
                                                                Else
                                                                    ShiftQueue = False
                                                                    If .MoveCounter > 0 Then .MoveCounter = .MoveCounter - 1
                                                                End If
                                                            Case QUEUE_SCRIPT
                                                                If .MonsterQueue(0).strData <> "" Then
                                                                PrintCrashDebug 5, 44
                                                                    Parameter(0) = mapNum
                                                                    Parameter(1) = A
                                                                    Parameter(2) = .MonsterQueue(0).lngData
                                                                    Parameter(3) = .MonsterQueue(0).lngData1
                                                                    RunScript .MonsterQueue(0).strData
                                                                    PrintCrashDebug 5, 45
                                                                End If
                                                                ShiftQueue = True
                                                            Case QUEUE_PAUSE
                                                                ShiftQueue = True
                                                        End Select
                                                        If ShiftQueue Then
                                                        PrintCrashDebug 5, 46
                                                            If .CurrentQueue > 0 Then
                                                                For C = 0 To .CurrentQueue - 1
                                                                    .MonsterQueue(C).Action = .MonsterQueue(C + 1).Action
                                                                    .MonsterQueue(C).lngData = .MonsterQueue(C + 1).lngData
                                                                    .MonsterQueue(C).lngData1 = .MonsterQueue(C + 1).lngData1
                                                                    .MonsterQueue(C).strData = .MonsterQueue(C + 1).strData
                                                                Next C
                                                                .CurrentQueue = .CurrentQueue - 1
                                                            End If
                                                            If .CurrentQueue = 0 Then .MonsterQueue(.CurrentQueue).Action = QUEUE_EMPTY
                                    
                                                        End If
                                                    End If
                                                End If
                                            Else
                                                SendToMap mapNum, Chr2(39) + Chr2(A) 'Monster Died
                                                .monster = 0
                                            End If
                                        End If
                                    Else
                                    PrintCrashDebug 5, 47
                                        With map(mapNum).MonsterSpawn(Int(A / 2))
                                            If .monster > 0 Then
                                                If Int(Rnd * .Rate) = 0 Then
                                                    ST1 = ST1 + NewMapMonster(mapNum, A)
                                                End If
                                            End If
                                        End With
                                    End If
                                End With
nexta:
                            Next A
                            If ST1 <> "" Then
                                SendToMapRaw mapNum, ST1
                            End If
                        End With
                    End If
                End If
                DoEvents
            Next Mapiterator
            
            For A = 1 To currentMaxUser
                player(A).DeferSends = False
                If player(A).St <> "" Then
                    FlushSocket A
                End If
            Next A

Exit Sub
Error_Handler:
    Open App.Path + "/LOG.TXT" For Append As #1
        ST1 = ""
        Print #1, Err.Number & "/" & Err.Description & "  Maptimer " & "/" & (idEvent)
    Close #1
    Unhook
    End
End Sub


