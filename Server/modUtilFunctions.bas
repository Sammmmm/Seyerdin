Attribute VB_Name = "modUtilFunctions"
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

Dim oSHA As New clsSHAAlgorithm

Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public finishCurrentMoveAddr As Long
Public Chr2(0 To 255) As String

Public Function AttackPlayer(ByVal Index As Long, ByVal Target As Long, ByVal damage As Long, ByVal RunAttackScript As Boolean, Optional Magic As Boolean = False, Optional FloatText As Boolean = True, Optional DropAtFeet As Boolean = True, Optional Projectile As Boolean = False, Optional trap As Boolean = False, Optional showAttack As Boolean = True, Optional CantMiss As Boolean = False) As Long
Dim A As Long, B As Long, C As Long, D As Long, F As Long
If Target > 0 And Target < MaxUsers Then
    player(Index).combatTimer = GetTickCount + 30000 ' player(Index).combatTimer + 30000
    player(Target).combatTimer = GetTickCount + 30000 ' player(Target).combatTimer + 30000
    'STATUS: INVISIBLE
    If GetStatusEffect(Index, SE_INVISIBLE) Then
        RemoveStatusEffect Index, SE_INVISIBLE
        CantMiss = True
    End If
    If GetStatusEffect(Target, SE_INVISIBLE) Then
        RemoveStatusEffect Target, SE_INVISIBLE
    End If
    
    
    'STATUS: ABSCOND
    If GetStatusEffect(Target, SE_ABSCOND) Then
        If Int(Rnd * 100) < player(Target).StatusData(SE_ABSCOND).Data(0) Then
            SetStatusEffect Target, SE_INVISIBLE
            player(Target).StatusData(SE_INVISIBLE).Data(1) = CAN_MOVE
            player(Target).StatusData(SE_INVISIBLE).timer = 5
        End If
    End If
    
    With player(Target)
        If RunAttackScript Then
            Parameter(0) = Index
            Parameter(1) = Target
            Parameter(2) = damage
            Parameter(3) = 0
            B = damage
            If Magic Then
                If Projectile = False Then Parameter(3) = AT_MAGIC Else Parameter(3) = AT_PROJECTILE_MAGIC
            Else
                If Projectile = False Then Parameter(3) = AT_MELEE Else Parameter(3) = AT_PROJECTILE_PHYSICAL
            End If
            If trap Then Parameter(3) = AT_TRAP
            Parameter(4) = 0
            damage = RunScript("ATTACKPLAYER")
            If (damage = 0 And B = damage) Or damage > 0 Then
            
                If GetStatusEffect(Target, SE_RETRIBUTION) Then
                    AttackPlayer Target, Index, player(Target).StatusData(SE_RETRIBUTION).Data(0), False, True, True
                End If
            End If
        End If
        
        C = PlayerEvasion(Target)
        If Projectile Then C = C / 4 'harder to miss with magic
        
        If Magic Then
            C = (C * 2) / 3

            damage = PlayerMagicDamage(Index, damage)
            damage = PlayerMagicArmor(Target, damage)
            If damage < 0 Then damage = 0
        Else
            If trap = False Then
                D = PlayerCritDamage(Index, damage)
                If D <> damage Then
                    damage = D
                    D = 1
                Else
                    D = 0
                End If
            End If
            damage = PlayerArmor(Target, damage)
        End If
    
        If player(Index).Guild <> 0 Then
            If player(Index).Guild = player(Target).Guild Then
                damage = damage / 2
            End If
        End If
        
        B = Int(Rnd * 255)
        If B > C Or CantMiss Then 'c = miss chance
            F = 1
            If damage > 0 Or .ShieldBlock Then
                If .HP > damage Then
                    .HP = .HP - damage
                    
                    If Magic Then
                        
                    Else
                        .Energy = .Energy - 2
                    End If
                    If .Energy < 0 Then .Energy = 0
                                        
                    SendToPartyAllBut .Party, Target, Chr2(104) + Chr2(0) + Chr2(Target) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                    SendToGods Chr2(104) + Chr2(0) + Chr2(Target) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                    SendToGuildAllBut Target, CLng(.Guild), Chr2(104) + Chr2(0) + Chr2(Target) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                    
                    
                Else
                    If player(Target).SkillLevel(SKILL_OPPORTUNIST) > 0 And player(Target).deathStamp + 60000 < GetTickCount Then
                        .deathStamp = GetTickCount
                        If FloatText Then CreateFloatingText .map, .x, .y, 12, "Opportunist!"
                    Else
                        .HP = 0
                    End If
                End If
                
                If .ShieldBlock And Not Magic And Object(.Equipped(2).Object).Data(2) <= 100 Then
                    If D Then
                        If FloatText Then CreateSizedFloatingText .map, .x, .y, 12, "Block - " & CStr(damage), critsize, 0
                        'CreateFloatingText .map, .x, .y, 12, "Block, Critical Hit! - " & CStr(damage)
                    Else
                        If FloatText Then CreateFloatingText .map, .x, .y, 12, "Block - " & CStr(damage)
                    End If
                ElseIf .ShieldBlock And Magic And Object(.Equipped(2).Object).Data(4) <= 100 Then
                    If D Then
                        If FloatText Then CreateSizedFloatingText .map, .x, .y, 12, "Block - " & CStr(damage), critsize, 0
                        'CreateFloatingText .map, .x, .y, 12, "Block, Critical Hit! - " & CStr(damage)
                    Else
                        If FloatText Then CreateFloatingText .map, .x, .y, 12, "Block - " & CStr(damage)
                    End If
                Else
                    If D Then
                        If FloatText Then CreateSizedFloatingText .map, .x, .y, 12, CStr(damage), critsize, 0
                        'CreateFloatingText .map, .x, .y, 12, "Critical Hit! - " & CStr(damage)
                    Else
                        If FloatText Then CreateFloatingText .map, .x, .y, 12, CStr(damage)
                    End If

                End If
                SendSocket Target, Chr2(49) + Chr2(D) + Chr2(Index) + DoubleChar(CInt(damage)) + DoubleChar(player(Target).Energy)
            Else
                If FloatText Then CreateFloatingEvent .map, .x, .y, FT_INEFFECTIVE
            End If
            .ShieldBlock = False
        Else
            F = 0
            If FloatText Then CreateFloatingEvent .map, .x, .y, FT_MISS
        End If
    End With
    'If Player(Index).Energy > 0 Then Player(Index).Energy = Player(Index).Energy - 1
    'If AttackAnim Then B = 0
    
    If showAttack Then SendSocket Index, Chr2(43) + Chr2(D) + Chr2(Target) + DoubleChar(CInt(damage))
    If showAttack Then SendToMapAllBut player(Index).map, Index, Chr2(42) + Chr2(Index) + Chr2(0)
    If player(Target).HP = 0 And player(Target).Access = 0 Then
        'Player Died
        If player(Target).Status <> 1 Then
            If Not (player(Index).Guild > 0 And player(Target).Guild > 0) Then
                If ExamineBit(map(player(Index).map).Flags(0), 2) = False Then
                    player(Index).Status = 1
                End If
            End If
        End If
        SendSocket Target, Chr2(52) + Chr2(Index) 'Player Killed You
        F = player(Target).Experience
        PlayerDied CLng(Target), True, CLng(Index), False, True
        If player(Target).Level > 5 Then
            F = F - player(Target).Experience
            
            If player(Target).Level - 4 > player(Index).Level Then
                F = CLng(F * (1 - 0.08 * ((player(Target).Level - 4) - player(Index).Level)))
                If F < 0 Then F = 0
                GainExp CLng(Index), F, True, True
            Else
                GainExp CLng(Index), F, True, True
            End If
            SendSocket Index, Chr2(45) + Chr2(Target) + QuadChar(F) 'You Killed Player
            
        Else
            SendSocket Index, Chr2(45) + Chr2(Target) + QuadChar(0) 'You Killed Player
        End If
        SendAllButBut Index, Target, Chr2(61) + Chr2(Target) + Chr2(Index) '+ Map(Player(Index).Map).Name 'Player was killed by player
    Else
        RemoveStatusEffect Target, SE_INVISIBLE
    End If
End If
End Function

Sub PlayerTriggerTrap(ByVal Index As Byte, ByVal trapnum As Byte)
    Dim mapNum As Long
    mapNum = player(Index).map
    With map(mapNum)
    If trapnum <= MaxTraps Then
        If .trap(trapnum).player > 0 Then
            If IsPlaying(.trap(trapnum).player) Then
            If CanAttackPlayer(.trap(trapnum).player, Index, True) Then
                If .trap(trapnum).trapID = player(.trap(trapnum).player).trapID And IsPlaying(.trap(trapnum).player) And .trap(trapnum).strength > 0 Then
                    If .trap(trapnum).ActiveCounter <= GetTickCount Then
                        If .trap(trapnum).CreatedTime + 180000 >= GetTickCount Then
                            If .trap(trapnum).Type = 1 Then
                                AttackPlayer .trap(trapnum).player, Index, .trap(trapnum).strength, True, False, True, False, False, True, False
                                SendToMap player(Index).map, Chr2(111) + Chr2(Index) + Chr2(SKILL_EXPLOSIVETRAP) + Chr2(.trap(trapnum).x) + Chr2(.trap(trapnum).y)
                                .trap(trapnum).CreatedTime = 0
                                .trap(trapnum).strength = 0
                                .trap(trapnum).player = 0
                            End If
                            If .trap(trapnum).Type = 2 Then
                                'player(Index).Energy = player(Index).Energy - .trap(trapnum).strength
                                'set frozen
                                SendSocket Index, Chr2(101) + Chr2(1) + Chr2(player(Index).x * 16 + player(Index).y)
                                InternalScriptTimer Index, .trap(trapnum).strength / 10, "unfreeze"
                                
                                SendToMap player(Index).map, Chr2(111) + Chr2(Index) + Chr2(SKILL_BEARTRAP) + Chr2(.trap(trapnum).x) + Chr2(.trap(trapnum).y)
                                .trap(trapnum).CreatedTime = 0
                                .trap(trapnum).strength = 0
                                .trap(trapnum).player = 0
                            End If
                        End If
                    End If
                End If
            End If
        End If
        End If
        End If
    End With
End Sub
Sub MonsterTriggerTrap(ByVal mapNum As Long, ByVal monsterIndex As Byte, ByVal trapnum As Byte)
    With map(mapNum)
        If .trap(trapnum).player > 0 Then
            If IsPlaying(.trap(trapnum).player) Then
                If CanAttackMonster(.trap(trapnum).player, monsterIndex) Then
                    If .trap(trapnum).trapID = player(.trap(trapnum).player).trapID And IsPlaying(.trap(trapnum).player) And .trap(trapnum).strength > 0 Then
                        If .trap(trapnum).ActiveCounter <= GetTickCount Then
                            If .trap(trapnum).CreatedTime + 180000 >= GetTickCount Then
                                If .trap(trapnum).Type = 1 Then
                                    AttackMapMonster mapNum, .trap(trapnum).player, monsterIndex, .trap(trapnum).strength, False, True, True, False, True, False
                                    SendToMap mapNum, Chr2(111) + Chr2(monsterIndex) + Chr2(SKILL_EXPLOSIVETRAP) + Chr2(.trap(trapnum).x) + Chr2(.trap(trapnum).y)
                                    .trap(trapnum).CreatedTime = 0
                                    .trap(trapnum).strength = 0
                                    .trap(trapnum).player = 0
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub
Sub CreateTrap(ByVal mapNum As Long, ByVal x As Byte, ByVal y As Byte, ByVal trapType As Byte, ByVal trapStrength As Byte, trapID As Long, playerIndex As Long, Optional ByVal trapDelay As Long = 3100)
    Dim A As Long, B As Long
    With map(mapNum)
        B = 0
        For A = 0 To MaxTraps
            If .trap(A).x = x And .trap(A).y = y Then
                PrintDebug player(playerIndex).user & "(" & player(playerIndex).Name & ")" & " triggered his own trap."
                PlayerTriggerTrap playerIndex, A
            End If
            If .trap(A).CreatedTime = 0 Or .trap(A).CreatedTime + 180000 < GetTickCount Or Not IsPlaying(.trap(A).player) Then
                B = A
                Exit For
            ElseIf .trap(A).CreatedTime < .trap(B).CreatedTime Then
                B = A
            End If
        Next A
        With .trap(B)
            .CreatedTime = GetTickCount
            .ActiveCounter = .CreatedTime + trapDelay
            .strength = trapStrength
            .Type = trapType
            .x = x
            .y = y
            .trapID = trapID
            .player = playerIndex
            If trapDelay > 0 Then
                SendToMap mapNum, Chr2(137) + Chr2(x) + Chr2(y) + Chr2(B) + DoubleChar(trapDelay)
            Else
                SendSocket playerIndex, Chr2(100) + Chr2(x * 16 + y) + Chr2(0) + StrConv("Trap Set!", vbUnicode)
            End If
        End With
            
    End With
End Sub

Function AttackMonster(ByVal Index As Byte, ByVal Target As Byte, ByVal damage As Long, Magic As Boolean, RunAttackScript As Boolean, Optional FloatText As Boolean = True, Optional Projectile As Boolean = False) As Long
    Dim mapNum As Long, A As Long, D As Long, E As Long
    If Target >= 0 And Target <= 9 Then
    
        If GetStatusEffect(Index, SE_INVISIBLE) Then
            RemoveStatusEffect Index, SE_INVISIBLE
        End If
    
        mapNum = player(Index).map
        If map(mapNum).monster(Target).monster > 0 Then
            With map(mapNum).monster(Target)
                .Target = Index
                .TargType = TargTypePlayer
                If RunAttackScript Then
                    Parameter(0) = Index
                    Parameter(1) = Target
                    Parameter(2) = damage
                    Parameter(3) = 0
                    If Magic Then
                        If Projectile Then Parameter(3) = AT_PROJECTILE_MAGIC Else Parameter(3) = AT_MAGIC
                    Else
                        If Projectile Then Parameter(3) = AT_PROJECTILE_PHYSICAL Else Parameter(3) = AT_MELEE
                    End If
                    Parameter(4) = 0
                    Parameter(5) = mapNum
                    damage = RunScript("ATTACKMONSTER")
                    If damage = -1 Then Exit Function
                End If
                If Magic Then
                    damage = PlayerMagicDamage(Index, damage)
                End If
                If .HP > damage Then
                    .HP = .HP - damage
                    SendToMapAllBut mapNum, Index, Chr2(44) + Chr2(0) + Chr2(Target) + DoubleChar(CLng(damage))
                    SendSocket Index, Chr2(44) + Chr2(0) + Chr2(0) + Chr2(Target) + DoubleChar(CLng(damage))
                Else
                    GainExp CLng(Index), CLng(monster(.monster).Experience), True, False, monster(.monster).Level

                    For D = 0 To 2
                        If Int(Rnd * 100 + 1) <= monster(.monster).Chance(D) Then
                            E = monster(.monster).Object(D)
                            If E > 0 Then
                                NewMapObject mapNum, E, monster(.monster).Value(D), CLng(.x), CLng(.y), False, GLOBAL_MAGIC_DROP_RATE + player(Index).MagicBonus
                            End If
                        End If
                    Next D

                     If player(Index).Buff.Type = BUFF_NECROMANCY Then
                        A = player(Index).Buff.Data(0)
                        If monster(.monster).Level < player(Index).Level Then
                            A = 1 + (A * monster(.monster).Level) / player(Index).Level
                        End If
                        If A > 0 Then
                            If player(Index).HP < player(Index).MaxHP Then
                                SetPlayerHP CByte(Index), player(Index).HP + A
                                If player(Index).HP > player(Index).MaxHP Then SetPlayerHP CByte(Index), player(Index).MaxHP
                            End If
                            'CreateFloatingText MapNum, player(index).x, player(index).y, 10, CStr(a)
                            SendToMap mapNum, Chr2(111) + Chr2(Index) + Chr2(SKILL_NECROMANCY) + Chr2(A)
                        End If
                        
                        
                        'SKILL: EVOCATION
                        A = 1 + player(Index).SkillLevel(SKILL_EVOCATION) + ((player(Index).Wisdom - player(Index).WisMod(1)) \ 20)
                        If monster(.monster).Level < player(Index).Level Then
                            A = 1 + (A * monster(.monster).Level) / player(Index).Level
                        End If
                        If A > 0 Then
                            If player(Index).Mana < player(Index).MaxMana Then
                                SetPlayerMana CByte(Index), player(Index).Mana + A
                                If player(Index).Mana > player(Index).MaxMana Then SetPlayerMana CByte(Index), player(Index).MaxMana
                            End If
                            SendToMap mapNum, Chr2(111) + Chr2(Index) + Chr2(SKILL_EVOCATION) + Chr2(A)
                            'CreateFloatingText MapNum, player(index).x, player(index).y, 9, CStr(a)
                        End If

                        
                    End If
                    A = .monster
                    Parameter(0) = Index
                    Parameter(1) = 0
                    Parameter(2) = Target
                    Parameter(3) = player(Index).map
                    RunScript ("MONSTERDIE" + CStr(.monster))
                    Parameter(0) = Index
                    Parameter(1) = 0
                    Parameter(2) = Target
                    Parameter(3) = player(Index).map
                    RunScript ("MONSTERDIE")
                    'Monster Died
                    SendToMapAllBut mapNum, Index, Chr2(39) + Chr2(Target) + DoubleChar(CLng(damage)) 'Monster Died
                    SendSocket Index, Chr2(51) + Chr2(Target) + QuadChar(monster(A).Experience) + DoubleChar(CLng(damage)) 'You killed monster
                    .monster = 0
                End If
            End With
        End If
    End If
End Function
Function AttackMapMonster(ByVal mapNum As Long, ByVal Index As Byte, ByVal Target As Byte, ByVal damage As Long, Magic As Boolean, RunAttackScript As Boolean, Optional FloatText As Boolean = True, Optional Projectile As Boolean = False, Optional trap As Boolean = False, Optional showAttack As Boolean = True) As Long
    Dim A As Long, D As Long, E As Long
    
    PrintDebug player(Index).user & "(" & player(Index).Name & ")" & " monster triggered a trap on map " & mapNum
    
    If Target >= 0 And Target <= 9 Then
    
        If GetStatusEffect(Index, SE_INVISIBLE) And trap = False Then
            RemoveStatusEffect Index, SE_INVISIBLE
        End If
    
        If map(mapNum).monster(Target).monster > 0 Then
            With map(mapNum).monster(Target)
                .Target = Index
                .TargType = TargTypePlayer
                If RunAttackScript Then
                    Parameter(0) = Index
                    Parameter(1) = Target
                    Parameter(2) = damage
                    Parameter(3) = 0
                    If Magic Then
                        If Projectile Then Parameter(3) = AT_PROJECTILE_MAGIC Else Parameter(3) = AT_MAGIC
                    Else
                        If Projectile Then Parameter(3) = AT_PROJECTILE_PHYSICAL Else Parameter(3) = AT_MELEE
                    End If
                    If trap Then Parameter(3) = AT_TRAP
                    Parameter(4) = 0
                    Parameter(5) = mapNum
                    damage = RunScript("ATTACKMONSTER")
                    If damage = -1 Then Exit Function
                End If
                If Magic Then
                    damage = PlayerMagicDamage(Index, damage)
                End If
                If .HP > damage Then
                    .HP = .HP - damage
                    If showAttack Then
                        SendToMapAllBut mapNum, Index, Chr2(44) + Chr2(0) + Chr2(Target) + DoubleChar(CLng(damage))
                        SendSocket Index, Chr2(44) + Chr2(0) + Chr2(0) + Chr2(Target) + DoubleChar(CLng(damage))
                    Else
                        SendToMapAllBut mapNum, Index, Chr2(138) + Chr2(0) + Chr2(Target) + DoubleChar(CLng(damage))
                        SendSocket Index, Chr2(138) + Chr2(0) + Chr2(0) + Chr2(Target) + DoubleChar(CLng(damage))
                    End If
                Else
                    'Monster Died
                    GainExp CLng(Index), CLng(monster(.monster).Experience), True, False, monster(.monster).Level

                    For D = 0 To 2
                        If Int(Rnd * 100 + 1) <= monster(.monster).Chance(D) Then
                            E = monster(.monster).Object(D)
                            If E > 0 Then
                                NewMapObject mapNum, E, monster(.monster).Value(D), CLng(.x), CLng(.y), False, GLOBAL_MAGIC_DROP_RATE + player(Index).MagicBonus
                            End If
                        End If
                    Next D

                    If player(Index).Buff.Type = BUFF_NECROMANCY Then
                        A = player(Index).Buff.Data(0)
                        If monster(.monster).Level < player(Index).Level Then
                            A = A - (player(Index).Level - monster(.monster).Level)
                        End If
                        If A > 0 Then
                            If player(Index).HP < player(Index).MaxHP Then
                                player(Index).HP = player(Index).HP + A
                                If player(Index).HP > player(Index).MaxHP Then player(Index).HP = player(Index).MaxHP
                                SendSocket Index, Chr2(46) + DoubleChar$(player(Index).HP)
                            End If
                            CreateFloatingText mapNum, player(Index).x, player(Index).y, 10, CStr(A)
                        End If
                    End If
                    'SKILL: EVOCATION
                    If player(Index).Buff.Type = BUFF_EVOCATION Then
                        A = player(Index).Buff.Data(0)
                        If monster(.monster).Level < player(Index).Level Then
                            A = A - (player(Index).Level - monster(.monster).Level)
                        End If
                        If A > 0 Then
                            If player(Index).Mana < player(Index).MaxMana Then
                                player(Index).Mana = player(Index).Mana + A
                                If player(Index).Mana > player(Index).MaxMana Then player(Index).Mana = player(Index).MaxMana
                                SendSocket Index, Chr2(48) + DoubleChar$(player(Index).Mana)
                            End If
                            CreateFloatingText mapNum, player(Index).x, player(Index).y, 9, CStr(A)
                        End If
                    End If
                    A = .monster
                    Parameter(0) = Index
                    Parameter(1) = 0
                    Parameter(2) = Target
                    Parameter(3) = mapNum
                    RunScript ("MONSTERDIE" + CStr(.monster))
                    Parameter(0) = Index
                    Parameter(1) = 0
                    Parameter(2) = Target
                    Parameter(3) = mapNum
                    RunScript ("MONSTERDIE")
                    SendToMapAllBut mapNum, Index, Chr2(39) + Chr2(Target) + DoubleChar(CLng(damage)) 'Monster Died
                    SendSocket Index, Chr2(51) + Chr2(Target) + QuadChar(monster(A).Experience) + DoubleChar(CLng(damage)) 'You killed monster
                    .monster = 0
                End If
            End With
        End If
    End If
End Function



Public Sub SetPlayerHP(Index As Long, ByVal HP As Long)
    With player(Index)
        If .Mode = modePlaying Then
            If HP > 9999 Then HP = 9999
            If HP < 1 Then HP = 1
            .HP = HP
            SendSocket Index, Chr2(46) + DoubleChar(CInt(HP))
                SendToPartyAllBut .Party, Index, Chr2(104) + Chr2(0) + Chr2(Index) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                 SendToGods Chr2(104) + Chr2(0) + Chr2(Index) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                SendToGuildAllBut Index, CLng(.Guild), Chr2(104) + Chr2(0) + Chr2(Index) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
        End If
    End With
End Sub

Public Function GeneratePrefixList()
    Dim A As Long
    
    For A = 1 To 10
        NumPrefix(A) = 0
    Next A
    
    For A = 1 To 255
        With prefix(A)
            If Not ExamineBit(.Flags, 1) Then 'Suffix flag not checked
                If Len(.Name) > 0 Then 'Actually a Prefix
                    If .Rarity > 0 And .Rarity <= 10 Then
                        NumPrefix(.Rarity) = NumPrefix(.Rarity) + 1
                        SortedPrefixList(.Rarity, NumPrefix(.Rarity)) = A
                    End If
                End If
            End If
        End With
    Next A
End Function

Public Function GenerateSuffixList()
    Dim A As Long
    
    For A = 1 To 10
        NumSuffix(A) = 0
    Next A
    
    For A = 1 To 255
        With prefix(A)
            If ExamineBit(.Flags, 1) Then  'Suffix flag checked
                If Len(.Name) > 0 Then 'Actually a suffix
                    If .Rarity > 0 And .Rarity <= 10 Then
                        NumSuffix(.Rarity) = NumSuffix(.Rarity) + 1
                        SortedSuffixList(.Rarity, NumSuffix(.Rarity)) = A
                    End If
                End If
            End If
        End With
    Next A
End Function

Sub SetIpSquelchTime(ip As String, Time As Long)
 Dim A As Long
 Dim B As Long
 For A = 0 To 255
    If ipSquelches(A).ip = vbNullString And B = 0 Then B = A
    If ipSquelches(A).ip = ip Then
        ipSquelches(A).Time = Time
        If Time = 0 Then ipSquelches(A).ip = vbNullString
        Exit Sub
    End If
 Next A
 ipSquelches(B).ip = ip
 ipSquelches(B).Time = Time
End Sub

Function GetIpSquelchTime(ip As String) As Byte
 Dim A As Long
 For A = 0 To 255

    If ipSquelches(A).ip = ip Then
        GetIpSquelchTime = ipSquelches(A).Time
        Exit Function
    End If
 Next A
GetIpSquelchTime = 0
End Function



Public Function GetTickCount() As Double
Dim curFreq As Currency
Dim curTime As Currency
Dim tDouble As Double

QueryPerformanceFrequency curFreq
QueryPerformanceCounter curTime

tDouble = (curTime / (curFreq / 1000))
While tDouble > 2147483647
    tDouble = tDouble - 2147483647
Wend
 
GetTickCount = CLng(tDouble)

End Function

Sub CreateFloatingText(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, ByVal Color As Long, ByVal Text As String)
    If mapNum >= 1 And mapNum <= 5000 Then
        If Len(Text) >= 1 Then
            SendToMap mapNum, Chr2(100) + Chr2(x * 16 + y) + Chr2(Color) + Text
        End If
    End If
End Sub
Sub CreateSizedFloatingText(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, ByVal Color As Long, ByVal Text As String, ByVal Mult As Long, ByVal Life As Long)
    If mapNum >= 1 And mapNum <= 5000 Then
        If Len(Text) >= 1 Then
            SendToMap mapNum, Chr2(131) + Chr2(x * 16 + y) + Chr2(Color) + Chr2(Mult) + Chr2(Life) + Text
        End If
    End If
End Sub
Sub CreateFloatingTextAllBut(ByVal Index As Long, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, ByVal Color As Byte, ByVal Text As String)
    If mapNum >= 1 And mapNum <= 5000 Then
        If Len(Text) >= 1 Then
            SendToMapAllBut mapNum, Index, Chr2(100) + Chr2(x * 16 + y) + Chr2(Color) + Text
        End If
    End If
End Sub

Sub CreateIndividualFloatingText(ByVal Index As Long, ByVal Color As Byte, ByVal Text As String)
If (Index >= 0) Then
    With player(Index)
        If .map >= 1 And .map <= 5000 Then
            If Len(Text) >= 1 Then
                SendSocket Index, Chr2(100) + Chr2(.x * 16 + .y) + Chr2(Color) + Text
            End If
        End If
    End With
End If
End Sub
Sub GenerateEXPLevels()
EXPLevel(1) = 500
EXPLevel(2) = 900
EXPLevel(3) = 1400
EXPLevel(4) = 2000
EXPLevel(5) = 2700
EXPLevel(6) = 3500
EXPLevel(7) = 4400
EXPLevel(8) = 5400
EXPLevel(9) = 6500

Dim A As Double, EXP As Double
For A = 10 To 50
    EXP = EXPLevel(A - 1) + ((A - 5) / 1.2) * 1000
    EXPLevel(A) = EXP
Next A
End Sub

Sub InitScriptTable()

 scriptTable.SetSize 20000 'if scripts get lost make this number BIGGER
 scriptTable.RemoveAll
 
If ScriptRS.BOF = False Then
    ScriptRS.MoveFirst
End If
    While ScriptRS.EOF = False
        scriptTable.Add ScriptRS!Name, StrConv(ScriptRS!Data, vbFromUnicode)
        ScriptRS.MoveNext
    Wend

End Sub
Sub WriteString(fileName As String, lpAppName As String, lpKeyName As String, lpString As String)
    Dim Valid As Long
    Valid = WritePrivateProfileString&(lpAppName, lpKeyName, lpString, App.Path + "\" + fileName + ".ini")
End Sub
Function ReadString(fileName As String, lpAppName As String, lpKeyName As String) As String
    Dim lpReturnedString As String, Valid As Long
    lpReturnedString = Space$(256)
    Valid = GetPrivateProfileString&(lpAppName, lpKeyName, "", lpReturnedString, 256, App.Path + "\" + fileName + ".ini")
    ReadString = Left$(lpReturnedString, Valid)
End Function
Function ReadInt(fileName As String, lpAppName As String, lpKeyName As String) As Integer
    ReadInt = GetPrivateProfileInt&(lpAppName, lpKeyName, 0, App.Path + "\" + fileName + ".ini")
End Function


Public Sub InitConstants()
    Dim A As Long
    SkillsPerLevel = ReadInt("Server", "Settings", "SkillsPerLevel")
    StatsPerLevel = ReadInt("Server", "Settings", "StatsPerLevel")
    MaxUsers = ReadInt("Server", "Settings", "MaxUsers")
    MaxLevel = ReadInt("Server", "Settings", "MaxLevel")
    DeathDropItemsLevel = ReadInt("Server", "Settings", "DeathDropItemsLevel")
    GLOBAL_OBJECT_RESET_RATE = ReadInt("Server", "Settings", "global_object_reset_rate")
    GLOBAL_OBJECT_RESET_RATE = GLOBAL_OBJECT_RESET_RATE * 1000
    GLOBAL_DEATH_DROP_RESET_RATE = ReadInt("Server", "Settings", "global_death_reset_rate")
    GLOBAL_DEATH_DROP_RESET_RATE = GLOBAL_DEATH_DROP_RESET_RATE * 1000
    GLOBAL_MAGIC_DROP_RATE = ReadInt("Server", "Settings", "global_magic_drop_rate")
    finishCurrentMoveAddr = GetValue(AddressOf FinishCurrentMove)
    
    StatRate1 = ReadInt("Server", "Settings", "StatRate1")
    StatRate2 = ReadInt("Server", "Settings", "StatRate2")
    
    baseEnergyRegen = ReadInt("Server", "Settings", "BaseEnergyRegen")
    BaseHPRegen = ReadInt("Server", "Settings", "baseHpRegen")
    baseManaRegen = ReadInt("Server", "Settings", "baseManaRegen")
    
    LevelsPerHpRegen = ReadInt("Server", "Settings", "LevelsPerHpRegen")
    LevelsPerManaRegen = ReadInt("Server", "Settings", "LevelsPerManaRegen")
    
    
    For A = 1 To 3
        StrengthPerDamage(A) = ReadInt("Server", "Settings", "StrengthPerDamage" & A)
        
        AgilityPerCritChance(A) = ReadInt("Server", "Settings", "AgilityPerCritChance" & A)
        AgilityPerDodgeChance(A) = ReadInt("Server", "Settings", "AgilityPerDodgeChance" & A)
        
        EndurancePerBlockChance(A) = ReadInt("Server", "Settings", "EndurancePerBlockChance" & A)
        EndurancePerEnergy(A) = ReadInt("Server", "Settings", "EndurancePerEnergy" & A)
        EndurancePerEnergyRegen(A) = ReadInt("Server", "Settings", "EndurancePerEnergyRegen" & A)
        
        ManaPerIntelligence(A) = ReadInt("Server", "Settings", "ManaPerIntelligence" & A)
        IntelligencePerManaRegen(A) = ReadInt("Server", "Settings", "IntelligencePerManaRegen" & A)
        
        ConstitutionPerHPRegen(A) = ReadInt("Server", "Settings", "ConstitutionPerHPRegen" & A)
        HPPerConstitution(A) = ReadInt("Server", "Settings", "HPPerConstitution" & A)
        
        PietyPerMagicResist(A) = ReadInt("Server", "Settings", "PietyPerMagicResist" & A)
        PietyPerHPRegen(A) = ReadInt("Server", "Settings", "PietyPerHPRegen" & A)
        PietyPerManaRegen(A) = ReadInt("Server", "Settings", "PietyPerManaRegen" & A)
        PietyPerHP(A) = ReadInt("Server", "Settings", "PietyPerHP" & A)
        PietyPerMana(A) = ReadInt("Server", "Settings", "PietyPerMana" & A)
        PietyPerDodge(A) = ReadInt("Server", "Settings", "PietyPerDodge" & A)
        PietyPerCrit(A) = ReadInt("Server", "Settings", "PietyPerCrit" & A)
        PietyPerBlock(A) = ReadInt("Server", "Settings", "PietyPerBlock" & A)
        
        GenericStatPerBonus(A) = ReadInt("Server", "Settings", "GenericStatPerBonus" & A)
        GenericPietyPerBonus(A) = ReadInt("Server", "Settings", "GenericPietyPerBonus" & A)
    Next A

    ReDim CloseSocketQueue(1 To MaxUsers) As SocketQueueData
    ReDim player(1 To MaxUsers + 1) As PlayerData
    ReDim CurrentMaps(1 To MaxUsers) As Long
End Sub

Public Sub CompactDb()
Dim A As Long, LingerType As LingerType
    PrintLog "Backup started at " & Date & ", " & Time
    'If ListeningSocket <> INVALID_SOCKET Then
    '    closesocket ListeningSocket
    'End If
    'PrintLog "Closing open sockets"
    'For A = 1 To MaxUsers
    '    If player(A).InUse = True And player(A).Mode = modePlaying Then
    '    If IsPlaying(A) Then SavePlayerData A
    '        CloseClientSocket A
    '    End If
    'Next A
    PrintLog "Sockets closed"
    PrintLog "Saving Flags..."
    SaveFlags
    PrintLog "Flags saved."
    PrintLog "Saving Objects..."
    SaveObjects
    PrintLog "Objects Saved"
    
    PrintLog "Closing Recordsets"
    UserRS.Close
    GuildRS.Close
    NPCRS.Close
    MonsterRS.Close
    ObjectRS.Close
    PrefixRS.Close
    DataRS.Close
    MapRS.Close
    BanRS.Close
    PrintLog "RecordSets Closed"
    db.Close
    WS.Close
    
    

    PrintLog "Compacting Database"
    Set WS = DBEngine.Workspaces(0)
    If Exists("server.tmp") Then Kill "server.tmp"
    Name "server.dat" As "server.tmp"
    CompactDatabase "server.tmp", "server.dat", , 0, ";pwd=" + Chr2(100) + Chr2(114) + Chr2(97) + Chr2(99) + Chr2(111)
    On Error Resume Next
        'MkDir ("Backups")
        FileCopy "Server.dat", "Backups\Server" & DateTime.Month(Date) & "-" & DateTime.Day(Date) & "-" & DateTime.Year(Date) & "  " & DateTime.Hour(Time) & "." & DateTime.Minute(Time) & "." & DateTime.Second(Time) & ".dat"
        'Kill "Backups\NightlyBackup.dat"
        'FileCopy "Server.dat", "Backups\NightlyBackup1" & Date & "." & Time & ".dat"
    On Error GoTo 0
    PrintLog "Database Compacted.  Opening Database"
        Set db = WS.OpenDatabase("server.dat", 12, False, ";pwd=" + Chr2(100) + Chr2(114) + Chr2(97) + Chr2(99) + Chr2(111))
        
    PrintLog "Database Opened"
    Kill "server.tmp"
    
    PrintLog "Initializing DB"
    InitDB
    PrintLog "DB Intialized"
    'Listen for connections
    With LingerType
        .l_onoff = 1
        .l_linger = 0
    End With
    PrintLog "Opening Listening Socket"
   ' ListeningSocket = ListenForConnect(ReadInt("Server", "Settings", "Port"), gHW, 1025)
   ' If ListeningSocket = INVALID_SOCKET Then
   '     MsgBox "Unable to create listening socket!", vbOKOnly + vbExclamation, TitleString
   '     EndWinsock
   '     Unhook
   '     End
   ' End If
   ' If setsockopt(ListeningSocket, SOL_SOCKET, SO_LINGER, LingerType, LenB(LingerType)) Then
   '     MsgBox "Unable to create listening socket!", vbOKOnly + vbExclamation, TitleString
   '     EndWinsock
   '     Unhook
   '     End
   ' End If
   ' If setsockopt(ListeningSocket, IPPROTO_TCP, TCP_NODELAY, 1&, 1) <> 0 Then
   '     MsgBox "Unable to create listening socket!", vbOKOnly + vbExclamation, TitleString
   '     EndWinsock
   '     Unhook
   '     End
   ' End If
    PrintLog "Socket Initialized, All systems go. - " & Date & ", " & Time
End Sub

Public Sub GlobalReduceRenown(AmountInThousanths As Long)
    Dim amount As Single
    amount = 1 - (0.001 * AmountInThousanths)
        db.Execute ("UPDATE ACCOUNTS SET RENOWN = RENOWN * " & amount)
End Sub

Public Sub UpdateLeaderboards()
    Dim rsL As Recordset
    Dim Count As Byte, A As Byte, B As Long, C As Long, D As Byte, Guilds As Byte
    Dim boldStart As String, boldEnd As String
    Count = 1
    
    On Error GoTo EndUpdateLB
    
    If (Exists("./Leaderboards/Renown.html")) Then Kill "./Leaderboards/Renown.html"
    Open App.Path + "/Leaderboards/Renown.html" For Append As #1
        Print #1, "<html><body><ul>"
        Print #1, "<style type='text/css'>tr.d0 td {background-color: #2B1B17; color: #FFFFCC;}tr.d1 td {background-color: #342826; color: #FFFFCC;} tr.d2 td {background-color: Black; color: #FFFFCC;} tr.d3 td {background-color: #453532; color: #FFFFCC;}</style>"
        Print #1, "<table>"
        Print #1, "<tr class='d0'>"
        
        Print #1, "<td width='50px'><b>Rank</b></td>"
        Print #1, "<td width = '200px'><b>Player</b></td>"
        Print #1, "<td width = '80px'><b>Renown</b></td>"
        Print #1, "<td width = '150px'><b>Class</b></td>"
        Print #1, "</tr>"
        
        Set rsL = db.OpenRecordset("SELECT TOP 25 NAME, RENOWN, CLASS FROM ACCOUNTS WHERE ACCESS = 0 AND CLASS <> 0 ORDER BY RENOWN DESC")
        rsL.MoveFirst
        While rsL.EOF = False
            If Count <= 10 Then
                If (Count Mod 2 = 0) Then
                    Print #1, "<tr class='d0'>"
                Else
                    Print #1, "<tr class='d1'>"
                End If
            Else
                Print #1, "<tr class='d3'>"
            End If
            If Count = 1 Then
                boldStart = "<b><font size='5' > "
                boldEnd = "</font></b>"
            Else
                If Count = 2 Then
                    boldStart = "<b><font size='4' > "
                    boldEnd = "</font></b>"
                Else
                    If Count = 3 Then
                        boldStart = "<b><font size='3' > "
                        boldEnd = "</font></b>"
                    Else
                        boldStart = ""
                        boldEnd = ""
                    End If
                End If
            End If
            Print #1, "<td>" & boldStart & "#" & Count & boldEnd & "</td>"
            Print #1, "<td>" & boldStart & rsL!Name & boldEnd & "</td>"
            Print #1, "<td>" & boldStart & rsL!Renown & boldEnd & "</td>"
            Print #1, "<td>" & boldStart & Class(rsL!Class).Name & boldEnd & "</td>"
            Print #1, "</tr>"
            rsL.MoveNext
            Count = Count + 1
            If Count = 11 Then
                Print #1, "<tr class='d2' height=5><td colspan='4'/></tr>"
            End If
        Wend
        
        Print #1, "</table></body></html>"
    Close #1
    DoEvents
    Count = 1
    If (Exists("./Leaderboards/GuildBank.html")) Then Kill "./Leaderboards/GuildBank.html"
    Open App.Path + "/Leaderboards/GuildBank.html" For Append As #1
        Print #1, "<html><body><ul>"
        Print #1, "<style type='text/css'>tr.d0 td {background-color: #2B1B17; color: #FFFFCC;}tr.d1 td {background-color: #342826; color: #FFFFCC;}</style>"
        Print #1, "<table>"
        Print #1, "<tr class='d0'>"
        
        Print #1, "<td width='50px'><b>Rank</b></td>"
        Print #1, "<td width = '200px'><b>Guild</b></td>"
        Print #1, "<td width = '80px'><b>Bank</b></td>"
        Print #1, "</tr>"
        
        Set rsL = db.OpenRecordset("SELECT TOP 100 NAME, BANK FROM GUILDS ORDER BY BANK DESC")
        If (rsL.RecordCount > 0) Then
            rsL.MoveFirst
            While rsL.EOF = False
                If Count <= 10 Then
                    If (Count Mod 2 = 0) Then
                        Print #1, "<tr class='d0'>"
                    Else
                        Print #1, "<tr class='d1'>"
                    End If
                Else
                    Print #1, "<tr class='d3'>"
                End If
                If Count = 1 Then
                    boldStart = "<b><font size='5' > "
                    boldEnd = "</font></b>"
                Else
                    If Count = 2 Then
                        boldStart = "<b><font size='4' > "
                        boldEnd = "</font></b>"
                    Else
                        If Count = 3 Then
                            boldStart = "<b><font size='3' > "
                            boldEnd = "</font></b>"
                        Else
                            boldStart = ""
                            boldEnd = ""
                        End If
                    End If
                End If
                
                Print #1, "<td>" & boldStart & "#" & Count & boldEnd & "</td>"
                Print #1, "<td>" & boldStart & rsL!Name & boldEnd & "</td>"
                Print #1, "<td>" & boldStart & rsL!Bank & boldEnd & "</td>"
                Print #1, "</tr>"
                rsL.MoveNext
                Count = Count + 1
                
                If Count = 11 Then
                    Print #1, "<Tr><td colspan='3' height='2px' style=""{background-color: Black;}""/></tr>"
                End If
            Wend
        End If
        Print #1, "</table></body></html>"
    Close #1
        
    Guilds = 0
    For A = 1 To 254
        If (Guild(A).Name <> "") Then
            Guilds = Guilds + 1
        
            Count = 0
            C = 0
            For B = 0 To 19
                If Guild(A).Member(B).Name <> "" Then
                    Count = Count + 1
                    C = C + Guild(A).Member(B).Renown
                End If
            Next B
            Guild(A).MemberCount = Count
            Guild(A).AverageRenown = C / Count
        End If
    Next A
        
    DoEvents
    Count = 1
    If (Exists("./Leaderboards/GuildRenown.html")) Then Kill "./Leaderboards/GuildRenown.html"
    Open App.Path + "/Leaderboards/GuildRenown.html" For Append As #1
        Print #1, "<html><body><ul>"
        Print #1, "<style type='text/css'>tr.d0 td {background-color: #2B1B17; color: #FFFFCC;}tr.d1 td {background-color: #342826; color: #FFFFCC;}</style>"
        Print #1, "<table>"
        Print #1, "<tr class='d0'>"
        
        Print #1, "<td width='50px'><b>Rank</b></td>"
        Print #1, "<td width = '200px'><b>Guild</b></td>"
        Print #1, "<td width = '80px'><b>Average Renown</b></td>"
        Print #1, "</tr>"
      
            C = 2147000000
            For Count = 1 To Guilds
                B = 0
                For A = 1 To 254
                    If Guild(A).AverageRenown > B And Guild(A).AverageRenown < C Then
                        B = Guild(A).AverageRenown
                        D = A
                    End If
                Next A
                C = B
                If Count <= 10 Then
                    If (Count Mod 2 = 0) Then
                        Print #1, "<tr class='d0'>"
                    Else
                        Print #1, "<tr class='d1'>"
                    End If
                Else
                    Print #1, "<tr class='d3'>"
                End If
                If Count = 1 Then
                    boldStart = "<b><font size='5' > "
                    boldEnd = "</font></b>"
                Else
                    If Count = 2 Then
                        boldStart = "<b><font size='4' > "
                        boldEnd = "</font></b>"
                    Else
                        If Count = 3 Then
                            boldStart = "<b><font size='3' > "
                            boldEnd = "</font></b>"
                        Else
                            boldStart = ""
                            boldEnd = ""
                        End If
                    End If
                End If
                
                Print #1, "<td>" & boldStart & "#" & Count & boldEnd & "</td>"
                Print #1, "<td>" & boldStart & Guild(D).Name & boldEnd & "</td>"
                Print #1, "<td>" & boldStart & Guild(D).AverageRenown & boldEnd & "</td>"
                Print #1, "</tr>"
                
                If Count = 11 Then
                    Print #1, "<Tr><td colspan='3' height='2px' style=""{background-color: Black;}""/></tr>"
                End If
            Next Count
        Print #1, "</table></body></html>"
    Close #1
        
         
        
        
    DoEvents
    Count = 1
    If (Exists("./Leaderboards/Info.html")) Then Kill "./Leaderboards/Info.html"
    Open App.Path + "/Leaderboards/Info.html" For Append As #1
        Print #1, "<html><body><ul>"
        Print #1, "<style type='text/css'>tr.d0 td {background-color: #2B1B17; color: #FFFFCC;}tr.d1 td {background-color: #342826; color: #FFFFCC;}</style>"
        Print #1, "<table>"
        Print #1, "<tr class='d0'>"
        
        Print #1, "<td width = '500px'><b>Information</b></td>"
        Print #1, "</tr>"
        
        Print #1, "<tr class='d1'><td>Last Updated: " & Hour(Time) & ":" & Minute(Now) & " " & Month(Now) & "/" & Day(Now) & "/" & Year(Now) & " (PST)</td></tr>"
        Print #1, "<tr class='d0'><td>Players Online: " & NumUsers & "</td></tr>"
        If (World.Flag(90) > 0) Then Print #1, "<tr class='d1'><td>Fort Owner: " & Guild(World.Flag(90)).Name & "</td></tr>"
        Print #1, "<tr class='d0'><td>Fort Opening Time: " & World.Flag(92) & ":00 (PST)</td></tr>"
        
        Print #1, "</table></body></html>"
    Close #1
    
EndUpdateLB:
End Sub

Sub AddToMovePlayerMoveQueue(Index As Long, ByVal packet As String)
    Dim A As Byte, x As Long, y As Long, St As String
    With player(Index)
        If .moveQueue(4) <> "" Then
            PrintLog "player " & player(Index).Name & " warped back!"
            ClearMoveQueue (Index)
            SendSocket Index, Chr2(10) + Chr2((.x * 16) + .y)
            Exit Sub
        End If
        
        For A = 0 To 4
            If .moveQueue(A) = "" Then
                .moveQueue(A) = packet
                Exit For
            End If
        Next A
        If .moveQueue(1) = "" Then ExecuteNextPlayerMove (Index)
    End With
End Sub

Sub FinishCurrentMove(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
KillTimer gHW, idEvent
With player(idEvent - 1000)
    .moveQueue(0) = .moveQueue(1)
    .moveQueue(1) = .moveQueue(2)
    .moveQueue(2) = .moveQueue(3)
    .moveQueue(3) = .moveQueue(4)
    .moveQueue(4) = ""
    If .moveQueue(0) <> "" Then ExecuteNextPlayerMove (idEvent - 1000)
End With
End Sub

Sub ClearMoveQueue(Index As Long)
        Dim A As Long
        With player(Index)
        For A = 0 To 4
            .moveQueue(A) = ""
        Next A
        End With
        KillTimer gHW, Index + 1000
End Sub

Sub ExecuteNextPlayerMove(Index As Long)
Dim St As String
Dim i As Long, J As Long, A As Long, B As Long, C As Long, mapNum As Long, D As Long, k As Long, L As Byte, E As Long, F As Long, G As Long, H As Long
With player(Index)

    St = .moveQueue(0)
    mapNum = .map
    If Len(St) = 3 Then
        If Asc(Mid$(St, 1, 1)) = .WalkCode Then
            If .Frozen = 0 Then
                'SKILL: HIDDEN
                If GetStatusEffect(Index, SE_INVISIBLE) Then
                    
                    If .StatusData(SE_INVISIBLE).Data(1) = NO_MOVE Then
                        RemoveStatusEffect Index, SE_INVISIBLE
                    End If
                End If
            
                i = .x
                J = .y
                A = Asc(Mid$(St, 2, 1))
                B = A And 15
                A = (A And 240) / 16
                C = Asc(Mid$(St, 3, 1)) \ 8
                If A < 0 Or A > 11 Or B < 0 Or B > 11 Then
                    BootPlayer Index, 0, "Walk_Around"
                    Exit Sub
                End If
                If Abs(A - CLng(.x)) + Abs(B - CLng(.y)) <= 1 Then
                    If .x <> A Or .y <> B Then
                        If (C) <> 7 And (C) <> 8 And (C) <> 16 And .Access = 0 Then
                            If map(mapNum).Tile(A, B).Att <> 26 Then
                              Hacker Index, "D.985"
                            ElseIf map(mapNum).Tile(A, B).AttData(0) - 1 <> C And map(mapNum).Tile(A, B).AttData(0) <> C And map(mapNum).Tile(A, B).AttData(1) <> C Then
                              Hacker Index, "D.986"
                            End If
                        End If
                        If map(.map).Tile(A, B).Att = 24 Then
                          If map(.map).Tile(.x, .y).Att <> 24 Then
                              SetPlayerFlag Index, 240, .map
                              SetPlayerFlag Index, 241, .x
                              SetPlayerFlag Index, 242, .y
                          End If
                        End If
                        .x = A
                        .y = B
                        .walkStamp = GetTickCount
                        D = 1
    
                            If map(mapNum).Tile(A, B).Att <> 26 Then
                              If (C) >= 16 Then
                                  'Use Energy
                                    If .Energy > 0 Then .Energy = .Energy - 1
                                    SendSocket Index, Chr2(47) + DoubleChar(CInt(.Energy))
                              End If
                            Else
                             If Not ExamineBit(map(mapNum).Tile(A, B).AttData(2), 0) Then
                               If (C) = map(mapNum).Tile(A, B).AttData(1) Then
                                  'Use Energy
                                    If .Energy > 0 Then .Energy = .Energy - 1
                               End If
                                    If ExamineBit(map(mapNum).Tile(A, B).AttData(2), 1) Then If .Energy > 0 Then .Energy = .Energy - 1
                                    If ExamineBit(map(mapNum).Tile(A, B).AttData(2), 2) Then If .Energy > 1 Then .Energy = .Energy - 2
                                    
                                    SendSocket Index, Chr2(47) + DoubleChar(CInt(.Energy))
                              End If
                            End If
    
                    Else
                        D = 0
                    End If
                    k = .D
                    .D = (Asc(Mid$(St, 3, 1)) And 7)
                    
                    If .D > 3 And .Access < 5 Then
                        Hacker Index, "D.99"
                    End If
                    
                    'check traps
                    For L = 0 To MaxTraps
                        With map(.map).trap(L)
                            If A = .x And B = .y Then
                                PrintDebug player(Index).user & "(" & player(Index).Name & ")" & " triggered a trap."
                                PlayerTriggerTrap Index, L
                            End If
                        End With
                    Next L
                    
                    If GetSkillLevel(Index, SKILL_BLOODTHIRSTY) > 0 And GetStatusEffect(Index, SE_BERSERK) Then
                       C = C + 2
                    End If
                    If .Access > 5 Then
                        SetTimer gHW, Index + 1000, 40, finishCurrentMoveAddr
                    Else
                        SetTimer gHW, Index + 1000, (3200 \ C) - 10, finishCurrentMoveAddr
                    End If
                    C = C * 8 + .D
                    If .D = 5 Then C = ((C And 248) Or k)

                    SendToMapAllBut mapNum, Index, Chr2(10) + Chr2(Index) + Chr2(.x * 16 + .y) + Chr2(C)
                    SendToPartyAllBut .Party, Index, Chr2(104) + Chr2(3) + Chr2(Index) + Chr2(.x * 16 + .y)
                    SendToGuildAllBut Index, CLng(.Guild), Chr2(104) + Chr2(3) + Chr2(Index) + Chr2(.x * 16 + .y)
                    
                    'Check if monsters notice
                    If .Access = 0 Then
                        For C = 0 To 9
                            If map(mapNum).monster(C).monster > 0 Then
                                With map(mapNum).monster(C)
                                    E = .x
                                    F = .y
                                    G = .Distance
                                End With
                                H = Sqr((CLng(.x) - E) * (CLng(.x) - E) + (CLng(.y) - F) * (CLng(.y) - F))
                                If H <= G Then
                                    If GetStatusEffect(Index, SE_INVISIBLE) = 0 Or (monster(map(mapNum).monster(C).monster).Flags And MONSTER_SEE_INVISIBLE) Then
                                        Parameter(0) = Index
                                        Parameter(1) = C
                                        Parameter(2) = H
                                        Parameter(3) = mapNum
                                        If RunScript("MONSTERSEE" + CStr(map(mapNum).monster(C).monster)) = 0 Then
                                            If ExamineBit(monster(map(mapNum).monster(C).monster).Flags, 3) = False Then
                                                'Isn't Friendly
                                                If ExamineBit(monster(map(mapNum).monster(C).monster).Flags, 0) = False Or .Status = 1 Then
                                                    With map(mapNum).monster(C)
                                                        If Index <> .Target Then
                                                            .Target = Index
                                                            .TargType = TargTypePlayer
                                                        End If
                                                        .Distance = H
                                                    End With
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Next C
                    End If
    
    
                    If (.x <> i Or .y <> J) And .D < 5 Then
                        Select Case .D
                            Case 1
                                If J = .y - 1 Then '''''''''''
                                  If ExamineBit(map(mapNum).Tile(.x, .y).WallTile, 0) Then WarpPlayer Index, .map, i, J
                                  If ExamineBit(map(mapNum).Tile(i, J).WallTile, 5) Then WarpPlayer Index, .map, i, J
                                  If map(mapNum).Tile(.x, .y).Att = 17 Then If ExamineBit(map(mapNum).Tile(.x, .y).AttData(0), 3) Then Hacker Index, "D.6.3"
                                End If
                            Case 0
                                If J = .y + 1 Then
                                  If ExamineBit(map(mapNum).Tile(.x, .y).WallTile, 1) Then WarpPlayer Index, .map, i, J
                                  If ExamineBit(map(mapNum).Tile(i, J).WallTile, 4) Then WarpPlayer Index, .map, i, J
                                  If map(mapNum).Tile(.x, .y).Att = 17 Then If ExamineBit(map(mapNum).Tile(.x, .y).AttData(0), 1) Then Hacker Index, "D.7.3"
                                End If
                            Case 3
                                If i = .x - 1 Then
                                  If ExamineBit(map(mapNum).Tile(.x, .y).WallTile, 2) Then WarpPlayer Index, .map, i, J
                                  If ExamineBit(map(mapNum).Tile(i, J).WallTile, 7) Then WarpPlayer Index, .map, i, J
                                  If map(mapNum).Tile(.x, .y).Att = 17 Then If ExamineBit(map(mapNum).Tile(.x, .y).AttData(0), 7) Then Hacker Index, "D.8.3"
                                End If
                            Case 2
                                If i = .x + 1 Then
                                  If ExamineBit(map(mapNum).Tile(.x, .y).WallTile, 3) Then WarpPlayer Index, .map, i, J
                                  If ExamineBit(map(mapNum).Tile(i, J).WallTile, 6) Then WarpPlayer Index, .map, i, J
                                  If map(mapNum).Tile(.x, .y).Att = 17 Then If ExamineBit(map(mapNum).Tile(.x, .y).AttData(0), 5) Then Hacker Index, "D.9.3"
                                End If
                        End Select
                    End If
                    If .D < 5 Then
                        Select Case map(mapNum).Tile(.x, .y).Att
                            Case 1, 13, 15 'Wall
                                If D = 1 Then
                                    .x = i
                                    .y = J
                                    WarpPlayer Index, .map, .x, .y
                                End If
                            Case 24 'Mine
                              If map(mapNum).Tile(.x, .y).AttData(3) > 0 Then
                                If D = 1 Then
                                    .x = i
                                    .y = J
                                    WarpPlayer Index, .map, .x, .y
                                End If
                              End If
                            Case 14
                                If ExamineBit(map(mapNum).Tile(.x, .y).AttData(0), 0) Then
                                    If D = 1 Then
                                        .x = i
                                        .y = J
                                        WarpPlayer Index, .map, .x, .y
                                    End If
                                End If
                            Case 17
                                If .x <> i Or .y <> J Then
                                    Select Case .D
                                        Case 1
                                            If ExamineBit(map(mapNum).Tile(.x, .y).AttData(0), 0) Then WarpPlayer Index, .map, i, J
                                        Case 0
                                            If ExamineBit(map(mapNum).Tile(.x, .y).AttData(0), 2) Then WarpPlayer Index, .map, i, J
                                        Case 3
                                            If ExamineBit(map(mapNum).Tile(.x, .y).AttData(0), 4) Then WarpPlayer Index, .map, i, J
                                        Case 2
                                            If ExamineBit(map(mapNum).Tile(.x, .y).AttData(0), 6) Then WarpPlayer Index, .map, i, J
                                    End Select
                                End If
                            Case 2 'Warp
                                A = map(mapNum).Tile(.x, .y).AttData(2)
                                B = map(mapNum).Tile(.x, .y).AttData(3)
                                C = CLng(map(mapNum).Tile(.x, .y).AttData(0)) * 256 + CLng(map(mapNum).Tile(.x, .y).AttData(1))
                                If A <= 11 And B <= 11 And C >= 1 And C <= 5000 Then
                                    WarpPlayer Index, C, A, B
                                Else
                                    AddSocketQue Index, 0
                                End If
                            Case 3 'Key Door
                                Partmap Index
                                .map = mapNum
                                .x = i
                                .y = J
                                JoinMap Index
                            Case 4 'Door
                                C = FreeMapDoorNum(mapNum)
                                If C >= 0 Then
                                    With map(mapNum).Door(C)
                                        .x = A
                                        .y = B
                                        .t = GetTickCount
                                        If ExamineBit(map(mapNum).Tile(A, B).AttData(3), 1) Then
                                            .Att = 4
                                            map(mapNum).Tile(A, B).Att = 0
                                        End If
                                            .Wall = map(mapNum).Tile(A, B).WallTile
                                            map(mapNum).Tile(A, B).WallTile = 0
                                    End With
                                    SendToMap mapNum, Chr2(36) + Chr2(C) + Chr2(A) + Chr2(B) + Chr2(3)
                                End If
                            Case 6
                                Parameter(0) = Index
                                Parameter(1) = map(mapNum).Tile(A, B).AttData(1)
                                Parameter(2) = map(mapNum).Tile(A, B).AttData(2)
                                Parameter(3) = map(mapNum).Tile(A, B).AttData(3)
                                Parameter(4) = A
                                Parameter(5) = B
                                RunScript ("NEWS" & map(mapNum).Tile(A, B).AttData(0))
                            Case 8 'Touch Plate
                                F = map(mapNum).Tile(A, B).AttData(2) ' is a guild tp
                                G = 0
                                If ExamineBit(map(mapNum).Tile(A, B).AttData(3), 3) Then G = 1
                                If ExamineBit(map(mapNum).Tile(A, B).AttData(3), 4) Then G = 2
                                If ExamineBit(map(mapNum).Tile(A, B).AttData(3), 5) Then G = 3
                                                                                              
                                If F > 0 Then
                                    If .Guild > 0 Then
                                        If .GuildRank >= G And Guild(.Guild).Hall = F Then
                                            G = 1
                                        Else
                                            G = 0
                                        End If
                                    Else
                                        G = 0
                                    End If
                                Else
                                    G = 1
                                End If
                                If G = 1 Then
                                    D = map(mapNum).Tile(A, B).AttData(0) 'x
                                    E = map(mapNum).Tile(A, B).AttData(1) 'y
                                    L = map(mapNum).Tile(A, B).AttData(3)
                                    If D <= 11 And E <= 11 Then
                                        If map(mapNum).Tile(D, E).Att > 0 Or map(mapNum).Tile(D, E).WallTile > 0 Then
                                            C = FreeMapDoorNum(mapNum)
                                            If C >= 0 Then
                                                With map(mapNum).Door(C)
                                                    .x = D
                                                    .y = E
                                                    .t = GetTickCount
                                                    .Att = map(mapNum).Tile(D, E).Att
                                                    .Wall = map(mapNum).Tile(D, E).WallTile
                                                    If ExamineBit(L, 1) Then map(mapNum).Tile(D, E).Att = 0
                                                    If ExamineBit(L, 0) Then map(mapNum).Tile(D, E).WallTile = 0
                                                End With
                                                SendToMap mapNum, Chr2(36) + Chr2(C) + Chr2(D) + Chr2(E) + Chr2(L)
                                            End If
                                        End If
                                    End If
                                End If
                            Case 9 'Damage Tile
                                Call SetPlayerHP(Index, player(Index).HP - map(mapNum).Tile(.x, .y).AttData(0))
                            Case 11 'Script
                                If D = 1 Then
                                    Parameter(0) = Index
                                    Parameter(1) = MC_WALK
                                    Parameter(2) = 0
                                    Parameter(3) = 0
                                    Parameter(4) = .x
                                    Parameter(5) = .y
                                    RunScript "MAP" + CStr(mapNum) + "_" + CStr(A) + "_" + CStr(B)
                                End If
                        End Select
                    End If
                    If .D = 5 Then .D = k
                Else
                    PrintLog "player " & player(Index).Name & " warped back."
                    ClearMoveQueue (Index)
                    SendSocket Index, Chr2(10) + Chr2((.x * 16) + .y)
                End If
            End If
        End If
    ElseIf Len(St) = 1 Then
        A = Asc(Mid$(St, 1, 1))
        If A < 4 Then
            .D = A
            SendToMapAllBut mapNum, Index, Chr2(10) & Chr2(Index) & Chr2(A)
            'Check if monsters notice
            If .Access = 0 Then
                For C = 0 To 9
                    If map(mapNum).monster(C).monster > 0 Then
                        With map(mapNum).monster(C)
                            E = .x
                            F = .y
                            G = .Distance
                        End With
                        H = Sqr((CLng(.x) - E) ^ 2 + (CLng(.y) - F) ^ 2)
                        If H <= G Then
                            If GetStatusEffect(Index, SE_INVISIBLE) = 0 Or (monster(map(mapNum).monster(C).monster).Flags And MONSTER_SEE_INVISIBLE) Then
                                Parameter(0) = Index
                                Parameter(1) = C
                                Parameter(2) = H
                                Parameter(3) = mapNum
                                If RunScript("MONSTERSEE" + CStr(map(mapNum).monster(C).monster)) = 0 Then
                                    If ExamineBit(monster(map(mapNum).monster(C).monster).Flags, 3) = False Then
                                        'Isn't Friendly
                                        If ExamineBit(monster(map(mapNum).monster(C).monster).Flags, 0) = False Or .Status = 1 Then
                                            With map(mapNum).monster(C)
                                                If Index <> .Target Then
                                                    .Target = Index
                                                    .TargType = TargTypePlayer
                                                    .Distance = H
                                                Else
                                                    .Distance = H
                                                End If
                                            End With
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next C
            End If
        Else
            Hacker Index, "A.16.1"
        End If
        FinishCurrentMove 0, 0, Index + 1000, 0
    Else
        Hacker Index, "A.16"
    End If
    'EndTime = GetTickCount
    'PrintLog (EndTime - StartTime)
End With
End Sub

Function GetStatPerBonus(ByVal statValue As Long, ByRef statLimits() As Byte) As Long
    Dim bonus As Single, currentStatValue As Long
    
    If statLimits(1) > 0 Then
        currentStatValue = statValue
        If statValue > StatRate2 - StatRate1 Then
            currentStatValue = StatRate2 - StatRate1
        Else
            currentStatValue = statValue
        End If
        bonus = currentStatValue / statLimits(1)
        statValue = statValue - currentStatValue
        
        If statValue > 0 And statLimits(2) > 0 Then
            If statValue > StatRate2 - StatRate1 Then currentStatValue = StatRate2 - StatRate1
            bonus = bonus + (currentStatValue / statLimits(2))
            statValue = statValue - currentStatValue
        
            If statValue > 0 And statLimits(3) > 0 Then
                bonus = bonus + (statValue / statLimits(3))
            End If
        End If
    End If
    
    GetStatPerBonus = Fix(bonus)
    bonus = bonus - Fix(bonus)
    bonus = bonus * 100
    If Int(Rnd * 100) < bonus Then GetStatPerBonus = GetStatPerBonus + 1
    
End Function

Function GetStatPerBonusHigh(ByVal statValue As Long, ByRef statLimits() As Byte) As Long
    Dim bonus As Single, currentStatValue As Long
    
    If statLimits(1) > 0 Then
        currentStatValue = statValue
        If statValue > StatRate2 - StatRate1 Then
            currentStatValue = StatRate2 - StatRate1
        Else
            currentStatValue = statValue
        End If
        bonus = currentStatValue / statLimits(1)
        statValue = statValue - currentStatValue
        
        If statValue > 0 And statLimits(2) > 0 Then
            If statValue > StatRate2 - StatRate1 Then currentStatValue = StatRate2 - StatRate1
            bonus = bonus + (currentStatValue / statLimits(2))
            statValue = statValue - currentStatValue
        
            If statValue > 0 And statLimits(3) > 0 Then
                bonus = bonus + (currentStatValue / statLimits(3))
            End If
        End If
    End If
    
    GetStatPerBonusHigh = Fix(bonus)
    bonus = bonus - Fix(bonus)
    bonus = bonus * 100
    If bonus > 3 Then GetStatPerBonusHigh = GetStatPerBonusHigh + 1
    
End Function


Function GetBonusPerStat(ByVal statValue As Long, ByRef statLimits() As Byte) As Long
    Dim bonus As Long, currentStatValue As Long
    
    If statLimits(1) > 0 Then
        currentStatValue = statValue
        If statValue > StatRate1 Then currentStatValue = StatRate1
        bonus = currentStatValue * statLimits(1)
        statValue = statValue - currentStatValue
        
        If statValue > 0 And statLimits(2) > 0 Then
            If statValue > StatRate2 - StatRate1 Then
                currentStatValue = StatRate2 - StatRate1
            Else
                currentStatValue = statValue
            End If
            bonus = bonus + (currentStatValue * statLimits(2))
            statValue = statValue - currentStatValue
        
            If statValue > 0 And statLimits(3) > 0 Then
                bonus = bonus + (statValue * statLimits(3))
            End If
        End If
    End If
    
    GetBonusPerStat = bonus
End Function

Function GetGenericStatBonus(ByVal statValue As Long, Optional ByVal pietyValue As Long = 0) As Long
    GetGenericStatBonus = GetStatPerBonus(statValue, GenericStatPerBonus) + GetStatPerBonus(statValue, GenericPietyPerBonus)
End Function

Sub InternalScriptTimer(ByVal Index As Long, ByVal Seconds As Long, ByVal Script As String)
Dim A As Long
    If Index >= 1 And Index <= MaxUsers Then
        If Seconds > 86400 Then Seconds = 86400
        If Seconds < 0 Then Seconds = 0
        With player(Index)
            If .Mode = modePlaying Then
                For A = 1 To MaxPlayerTimers
                    If .ScriptTimer(A) = 0 Then
                        .Script(A) = Script
                        .ScriptTimer(A) = GetTickCount + Seconds * 1000
                        Exit For
                        'Parameter(0) = Index
                        '.ScriptTimer = 0
                        'ScriptRunning = False
                        'RunScript .Script
                        'ScriptRunning = True
                    End If
                Next A
            End If
        End With
    End If
End Sub

Function SHA256(Value As String)
    SHA256 = oSHA.SHA256FromString(Value)
End Function
