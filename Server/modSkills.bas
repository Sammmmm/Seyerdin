Attribute VB_Name = "modSkills"
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

'Class Constants
Public Const C_CLERIC = 1
Public Const C_MAGE = 2
Public Const C_THIEF = 3
Public Const C_WARRIOR = 4
Public Const C_PALADIN = 5
Public Const C_CRUSADER = 6
Public Const C_NECROMANCER = 7
Public Const C_ENCHANTER = 8
Public Const C_BARBARIAN = 9
Public Const C_CHARLATAN = 10


'Target Type Constants
Public Const TT_CHARACTER = 1
Public Const TT_PLAYER = 2
Public Const TT_MONSTER = 4
Public Const TT_TILE = 8
Public Const TT_NO_TARGET = 16

'Skill Flag Constants
Public Const SF_USELOS = 1      'Use the Line Of Sight to determine if a target can be hit
Public Const SF_HOSTILE = 2     'Hostile can only be used in hostile enviroments
Public Const SF_FRIENDLY = 4    'Friendly skills can be used anywhere
Public Const SF_USERANGE = 8    'Use range to determine if target is attackable
Public Const SF_UNKNOWN = 16    'This determines whether or not this skill is unknown until learned
Public Const SF_AUTO = 32
Public Const SF_RUNSCRIPT = 64  'Runs SPELL# script when used
Public Const SF_DYNRANGE = 128   'Determines whether the range increases with skill increase
Public Const SF_DYNMANA = 256
Public Const SF_DYNENERGY = 512

'EXP Table Constants
Public Const ET_STANDARD = 0

'SkillType Constants
Public Const _
            ST_ELECTRICITY = 1, _
            ST_FIRE = 2, _
            ST_SLASHING = 4, _
            ST_ICE = 8, _
            ST_HOLY = 16
            
'Status Effect Constants
Public Const _
            SE_EXHAUST = 1, _
            SE_POISON = 2, _
            SE_MUTE = 3, _
            SE_CRIPPLE = 4, _
            SE_BLIND = 5, _
            SE_BERSERK = 7, _
            SE_RETRIBUTION = 8, _
            SE_REGENERATION = 9, _
            SE_INVULNERABILITY = 10, _
            SE_HASTE = 11, _
            SE_SHATTERSHIELD = 12, _
            SE_DEADLYCLARITY = 13, _
            SE_REUSEME = 14, _
            SE_INVISIBLE = 15, _
            SE_MASTERDEFENSE = 16, _
            SE_EVANESCENCE = 17, _
            SE_ETHEREALITY = 18, _
            SE_ABSCOND = 19, _
            SE_LOSTARCANA = 20
Public Const _
            SE_FIERYESSENCE = 21, _
            SE_ENERGYMOD = 22, _
            SE_MANAMOD = 23, _
            SE_STRENGTHMOD = 24, _
            SE_AGILITYMOD = 25, _
            SE_ENDURANCEMOD = 26, _
            SE_WISDOMMOD = 27, _
            SE_CONSTITUTIONMOD = 28, _
            SE_INTELLIGENCEMOD = 29, _
            SE_HPREGENMOD = 30, _
            SE_MPREGENMOD = 31, _
            SE_POISONRESISTMOD = 32, _
            SE_MAGICRESISTMOD = 33, _
            SE_ATTACKSPEEDMOD = 34, _
            SE_CRITICALCHANCEMOD = 35, _
            SE_HPMOD = 36, _
            SE_MANAMULT = 37, _
            SE_ALL = 2147483647

Public Const NO_MOVE = 0
Public Const CAN_MOVE = 1

'Buff Constants
Public Const _
            BUFF_EMPOWER = 1, _
            BUFF_WILLPOWER = 2, _
            BUFF_ZEAL = 3, _
            BUFF_VITALITY = 4, _
            BUFF_HOLYARMOR = 5, _
            BUFF_NECROMANCY = 6, _
            BUFF_EVOCATION = 7, _
            BUFF_MANASHIELD = 8

'Skill Requirement
Public Type SkillRequirementType
    SKILL As Byte
    Level As Byte
End Type

'Skill information
Type SkillType
    Class As Integer
    Level(1 To 10) As Byte
    CostPerLevel As Single
    CostPerSLevel As Integer
    CostConstant As Integer
    TargetType As Byte
    Type As Long
    DataLength As Byte
    Flags As Long
    Range As Byte
    MaxLevel As Byte
    EXPTable As Byte
    damage As Integer
    Requirements(1) As SkillRequirementType
    LocalTick As Long
    GlobalTick As Long
End Type

'Target Information
Type TargetData
    TargetType As Byte
    Target As Byte
    x As Byte
    y As Byte
End Type

Public Const MAX_SKILLS = 255

Public Const _
    SKILL_INVALID = 0, _
    SKILL_HEAL = 1, _
    SKILL_CUREPOISON = 2, _
    SKILL_PARTYHEAL = 3, _
    SKILL_GREATERHEAL = 4, _
    SKILL_EMPOWER = 5, _
    SKILL_WILLPOWER = 6, _
    SKILL_ZEAL = 7, _
    SKILL_VITALITY = 8, _
    SKILL_REGENERATE = 9, _
    SKILL_HOLYARMOR = 10, _
    SKILL_INVULNERABILITY = 11, _
    SKILL_HEAVENLYSTRIKE = 12, _
    SKILL_LIGHTNING = 13, _
    SKILL_HEAVENSTORM = 14, _
    SKILL_RETRIBUTION = 15, _
    SKILL_NECROMANCY = 16, _
    SKILL_EVOCATION = 17, _
    SKILL_DARKNESS = 18, _
    SKILL_REINCARNATION = 19, _
    SKILL_FIREBALL = 20, _
    SKILL_FLAMEWAVE = 21, _
    SKILL_ASTRALGLARE = 22
Public Const _
    SKILL_SNOWSTORM = 23, _
    SKILL_TEMPEST = 24, _
    SKILL_BLIZZARD = 25, _
    SKILL_SHATTERSHIELD = 26, _
    SKILL_MANASHIELD = 27, _
    SKILL_STEALESSENCE = 28, _
    SKILL_TELEPORT = 29, _
    SKILL_PORTAL = 30, _
    SKILL_VOIDBOLT = 31, _
    SKILL_EVANESCENCE = 32, _
    SKILL_INVISIBILITY = 33, _
    SKILL_MANARESERVES = 34, _
    SKILL_SPELLPOWER = 35, _
    SKILL_EXPLOSIVETRAP = 36, _
    SKILL_BEARTRAP = 37, _
    SKILL_PARALYZE = 38, _
    SKILL_EVASION = 39, _
    SKILL_ETHEREALITY = 40, _
    SKILL_SECONDWIND = 41, _
    SKILL_STEALTH = 42, _
    SKILL_BACKSTAB = 43, _
    SKILL_PIERCE = 44, _
    SKILL_POIGNANCY = 45, _
    SKILL_DEADLYCLARITY = 46
Public Const _
    SKILL_DESOLATION = 47, _
    SKILL_ENVENOM = 48, _
    SKILL_BLIGHT = 49, _
    SKILL_CRIPPLE = 50, _
    SKILL_BASH = 51, _
    SKILL_FURY = 52, _
    SKILL_MAIM = 53, _
    SKILL_GREATSTRENGTH = 54, _
    SKILL_ARMORMASTERY = 55, _
    SKILL_SHIELDMASTERY = 56, _
    SKILL_MASTERDEFENSE = 57, _
    SKILL_GREATFORTITUDE = 58, _
    SKILL_TAUNT = 59, _
    SKILL_FRIGHTEN = 60, _
    SKILL_SILENCE = 61, _
    SKILL_TWOHAND = 62, _
    SKILL_DAMAGE = 63, _
    SKILL_DEFENSE = 64, _
    SKILL_SPEED = 65, _
    SKILL_IMMUNITY = 66, _
    SKILL_BERSERK = 67, _
    SKILL_BLOODLETTER = 68, _
    SKILL_VENGEANCE = 69, _
    SKILL_CONVALESCENCE = 70
Public Const _
    SKILL_PRIESTHOOD = 71, _
    SKILL_DEVOTION = 72, _
    SKILL_SPIRITSTOUCH = 73, _
    SKILL_SORCERY = 74, _
    SKILL_DRAWESSENCE = 75, _
    SKILL_MYSTICISM = 76, _
    SKILL_AGILITY = 77, _
    SKILL_OPPORTUNIST = 78, _
    SKILL_PERCEPTION = 79, _
    SKILL_VIGOR = 80, _
    SKILL_GREATFORCE = 81, _
    SKILL_COLOSSUS = 82, _
    SKILL_GUARDIAN = 83, _
    SKILL_VENGEFUL = 84, _
    SKILL_BURNINGSOUL = 85, _
    SKILL_FANATACISM = 86, _
    SKILL_FIERYESSENCE = 87, _
    SKILL_BLACKMAGIC = 88, _
    SKILL_LOSTARCANA = 89, _
    SKILL_SIPHONLIFE = 90, _
    SKILL_SUMMONBLADE = 91, _
    SKILL_SUMMONSTAFF = 92, _
    SKILL_SUMMONJEWEL = 93, _
    SKILL_WIZARDRY = 94
Public Const _
    SKILL_MAGICALCONDUIT = 95, _
    SKILL_RESTLESS = 96, _
    SKILL_FEROCITY = 97, _
    SKILL_BLOODTHIRSTY = 98, _
    SKILL_JUDGMENT = 99


    
    

Public Skills(1 To MAX_SKILLS) As SkillType
Public Sub InitSkills()

Dim St As String * 331, A As Long, B As Long
Dim CurrentSkill As SkillType
Dim tByte(1 To 4) As Byte
'ReDim Skills(0)
Open App.Path & "/skilldata.dat" For Binary As #1
    If LOF(1) Mod 331 = 0 And LOF(1) > 0 Then
        A = LOF(1) / Len(CurrentSkill)
        For B = 1 To MAX_SKILLS
            Get #1, , St
            
            With Skills(B)
                .Class = GetInt(Mid$(St, 289, 2))
                .TargetType = Asc(Mid$(St, 295, 1))
                .Type = Asc(Mid$(St, 296, 1)) * 16777216 + Asc(Mid$(St, 297, 1)) * 65536 + Asc(Mid$(St, 298, 1)) * 256& + Asc(Mid$(St, 299, 1))
                .Flags = Asc(Mid$(St, 300, 1)) * 16777216 + Asc(Mid$(St, 301, 1)) * 65536 + Asc(Mid$(St, 302, 1)) * 256& + Asc(Mid$(St, 303, 1))
                .Range = Asc(Mid$(St, 304, 1))
                .MaxLevel = Asc(Mid$(St, 305, 1))
                .EXPTable = Asc(Mid$(St, 306, 1))
                .Requirements(0).SKILL = Asc(Mid$(St, 307, 1))
                .Requirements(0).Level = Asc(Mid$(St, 308, 1))
                .Requirements(1).SKILL = Asc(Mid$(St, 309, 1))
                .Requirements(1).Level = Asc(Mid$(St, 310, 1))
                tByte(1) = Asc(Mid$(St, 311, 1))
                tByte(2) = Asc(Mid$(St, 312, 1))
                tByte(3) = Asc(Mid$(St, 313, 1))
                tByte(4) = Asc(Mid$(St, 314, 1))
                MemCopy .CostPerLevel, tByte(1), 4
                .CostPerSLevel = Asc(Mid$(St, 315, 1))
                .CostConstant = Asc(Mid$(St, 316, 1))
                .Level(1) = Asc(Mid$(St, 317, 1))
                .Level(2) = Asc(Mid$(St, 318, 1))
                .Level(3) = Asc(Mid$(St, 319, 1))
                .Level(4) = Asc(Mid$(St, 320, 1))
                .Level(5) = Asc(Mid$(St, 321, 1))
                .Level(6) = Asc(Mid$(St, 322, 1))
                .Level(7) = Asc(Mid$(St, 323, 1))
                .Level(8) = Asc(Mid$(St, 324, 1))
                .Level(9) = Asc(Mid$(St, 325, 1))
                .Level(10) = Asc(Mid$(St, 326, 1))
                .GlobalTick = GetInt(Mid$(St, 327, 2)) * 1000
                .LocalTick = GetInt(Mid$(St, 329, 2)) * 1000
            End With
        Next B
    End If
Close #1
    InitSkillEXPTable
End Sub

Public Function CanUseSkill(Index As Long, SKILL As Byte) As Byte
Dim A As Byte

CanUseSkill = 1

With player(Index)
    If GetStatusEffect(Index, SE_MUTE) = 0 Then
        If ((2 ^ (.Class - 1)) And Skills(SKILL).Class) Then
            If .Level >= Skills(SKILL).Level(.Class) Then
                If .Mana >= GetSkillManaCost(Index, SKILL) Then
                    If GetTickCount > (.GlobalSpellTick - 150) Then
                        If GetTickCount > (.LocalSpellTick(SKILL) - 150) Then
                            If (Skills(SKILL).Flags And SF_UNKNOWN) Then
                                If .SkillLevel(SKILL) > 0 Then
                                    CanUseSkill = 1
                                Else
                                    CanUseSkill = 0
                                End If
                            Else
                                CanUseSkill = 1
                            End If
                            For A = 0 To 1
                                With Skills(SKILL).Requirements(A)
                                    If .SKILL > SKILL_INVALID And .SKILL < MAX_SKILLS Then
                                        If player(Index).SkillLevel(.SKILL) < .Level Then
                                            CanUseSkill = 0
                                        End If
                                    End If
                                End With
                            Next A
                        Else
                            CanUseSkill = 1
                        End If
                    Else
                        CanUseSkill = 1
                    End If
                Else
                    CanUseSkill = 0
                End If
            Else
                CanUseSkill = 0
            End If
        Else
            CanUseSkill = 0
        End If
    Else
        CanUseSkill = 0
    End If
End With
End Function

Public Function UseSkill(Index As Long, SKILL As Byte, ByRef St As String) As Byte
Dim CurrentTarget As TargetData
Dim A As Long, B As Long, x As Long, y As Long
Dim ST1 As String

If SKILL > MAX_SKILLS Then Exit Function
If CanUseSkill(Index, SKILL) Then
    With player(Index)
        ST1 = ""
        If GetSkillManaCost(Index, SKILL) > 0 Then
            If .SkillLevel(SKILL_WIZARDRY) And ExamineBit(Skills(SKILL).Flags, SF_HOSTILE) Then
                .Mana = .Mana - CInt(GetSkillManaCost(Index, SKILL) * 0.66)
            Else
                If .SkillLevel(SKILL_MYSTICISM) And Int(Rnd * 10) < 1 Then
                    SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(SKILL_MYSTICISM)
                Else
                    .Mana = .Mana - GetSkillManaCost(Index, SKILL)
                End If
            End If
            
            SendToPartyAllBut .Party, Index, Chr2(104) + Chr2(0) + Chr2(Index) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
            SendToGods Chr2(104) + Chr2(0) + Chr2(Index) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
            SendToGuildAllBut Index, CLng(.Guild), Chr2(104) + Chr2(0) + Chr2(Index) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
            
            ST1 = ST1 + DoubleChar(3) + Chr2(48) + DoubleChar(CInt(.Mana))
        End If
        .LocalSpellTick(SKILL) = GetTickCount + Skills(SKILL).LocalTick
        .GlobalSpellTick = GetTickCount + Skills(SKILL).GlobalTick
        If ST1 <> "" Then
            SendRaw Index, ST1
        End If
    End With

    With player(Index)
        If St = "" Then
            If Skills(SKILL).TargetType And TT_NO_TARGET Then
                CurrentTarget.TargetType = TT_NO_TARGET
            End If
        Else
            'CurrentTarget.TargetType = Asc(Mid$(St, 1, 1))
        End If
        
        If ((Skills(SKILL).TargetType = TT_NO_TARGET) And Len(St) = 0) Or (Len(St) = 4) Then
            LoadTargetData St, CurrentTarget
            Parameter(0) = Index
            Parameter(1) = CurrentTarget.TargetType
            Parameter(2) = CurrentTarget.Target
            Parameter(3) = CurrentTarget.x
            Parameter(4) = CurrentTarget.y
            Parameter(5) = SKILL
            
            A = 0
            
            Select Case SKILL
                Case SKILL_HEAL
                    If CurrentTarget.TargetType = TT_CHARACTER Then
                        A = Index
                    ElseIf CurrentTarget.TargetType = TT_PLAYER Then
                        A = CurrentTarget.Target
                    End If
                    B = RunScript("SPELL" & SKILL_HEAL)
                    'If CanTargetPlayer(Index, SKILL_HEAL, A) Then
                    If A > 0 And A <= MaxUsers Then
                        With player(A)
                            If .HP < .MaxHP Then
                                .HP = .HP + B
                                If .HP > .MaxHP Then .HP = .MaxHP
                                SendSocket A, Chr2(46) + DoubleChar(.HP)
                                SendToPartyAllBut .Party, A, Chr2(104) + Chr2(0) + Chr2(A) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                                SendToGods Chr2(104) + Chr2(0) + Chr2(A) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                                SendToGuildAllBut A, CLng(.Guild), Chr2(104) + Chr2(0) + Chr2(A) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                            End If
                            SendToMap .map, Chr2(111) + Chr2(A) + Chr2(SKILL_HEAL) + Chr2(.x) + Chr2(.y) + DoubleChar(B)
                        End With
                    End If
                    'end if
                Case SKILL_SPIRITSTOUCH
                    B = RunScript("SPELL" & SKILL_SPIRITSTOUCH)
                    If CurrentTarget.TargetType = TT_CHARACTER Then
                        A = Index
                    ElseIf CurrentTarget.TargetType = TT_PLAYER Then
                        A = CurrentTarget.Target
                    End If
     
                    If A > 0 And A <= MaxUsers Then
                        With player(A)
                            If .InUse And .Mode = modePlaying Then
                                RemoveStatusEffect CByte(A), SE_POISON
                                RemoveStatusEffect CByte(A), SE_BLIND
                                RemoveStatusEffect CByte(A), SE_MUTE
                                RemoveStatusEffect CByte(A), SE_CRIPPLE
                                RemoveStatusEffect CByte(A), SE_EXHAUST
                            End If
                            If .HP < .MaxHP Then
                                .HP = .HP + B
                                If .HP > .MaxHP Then .HP = .MaxHP
                                SendSocket A, Chr2(46) + DoubleChar(.HP)
                                SendToPartyAllBut .Party, A, Chr2(104) + Chr2(0) + Chr2(A) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                                SendToGods Chr2(104) + Chr2(0) + Chr2(A) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                                SendToGuildAllBut A, CLng(.Guild), Chr2(104) + Chr2(0) + Chr2(A) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                            End If
                            SendToMap .map, Chr2(111) + Chr2(A) + Chr2(SKILL_SPIRITSTOUCH) + Chr2(.x) + Chr2(.y) + DoubleChar(B)
                        End With
                    End If
                    
              
                Case SKILL_SUMMONJEWEL, SKILL_SUMMONSTAFF, SKILL_SUMMONBLADE
                    RunScript ("SPELL" & SKILL)
                Case SKILL_CUREPOISON
                    If CurrentTarget.TargetType = TT_CHARACTER Then
                        A = Index
                    ElseIf CurrentTarget.TargetType = TT_PLAYER Then
                        A = CurrentTarget.Target
                    End If
                    If A > 0 And A <= MaxUsers Then
                        With player(A)
                            If .InUse And .Mode = modePlaying Then
                                RemoveStatusEffect CByte(A), SE_POISON
                                
                                If .SkillLevel(SKILL_CUREPOISON) > 9 Then RemoveStatusEffect CByte(A), SE_BLIND
                                If .SkillLevel(SKILL_CUREPOISON) > 7 Then RemoveStatusEffect CByte(A), SE_MUTE
                                If .SkillLevel(SKILL_CUREPOISON) > 5 Then RemoveStatusEffect CByte(A), SE_CRIPPLE
                                If .SkillLevel(SKILL_CUREPOISON) > 3 Then RemoveStatusEffect CByte(A), SE_EXHAUST
                                
                                SendToMap .map, Chr2(111) + Chr2(A) + Chr2(SKILL_CUREPOISON) + Chr2(player(A).x) + Chr2(player(A).y)
                            End If
                        End With
                    End If
                Case SKILL_PARTYHEAL
                    B = RunScript("SPELL" & SKILL_PARTYHEAL) '10 + (.SkillLevel(SKILL_PARTYHEAL) * 3) + (.Wisdom \ 10)
                    ST1 = DoubleChar(B)
                    For A = 1 To currentMaxUser
                        If .Party > 0 Or .Guild > 0 Then
                            If ((player(A).Party = .Party And .Party > 0) Or (player(A).Guild = .Guild And .Guild > 0)) And player(A).map = .map And A <> Index Then
                                If (CLng(player(A).x) - CLng(player(Index).x)) ^ 2 + (CLng(player(A).y) - CLng(player(Index).y)) ^ 2 <= Skills(SKILL_PARTYHEAL).Range ^ 2 Then
                                    'If player(a).HP < player(a).MaxHP Then
                                        player(A).HP = player(A).HP + B
                                        If player(A).HP > player(A).MaxHP Then player(A).HP = player(A).MaxHP
                                        If A = Index Then SendSocket A, Chr2(46) + DoubleChar(player(A).HP)
                                        SendToPartyAllBut .Party, A, Chr2(104) + Chr2(0) + Chr2(A) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                                        SendToGods Chr2(104) + Chr2(0) + Chr2(A) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                                        SendToGuildAllBut A, CLng(.Guild), Chr2(104) + Chr2(0) + Chr2(A) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                                        ST1 = ST1 + Chr2(A)
                                    'End If
                                End If
                            End If
                        End If
                    Next A
                    If ST1 <> "" Then SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(SKILL_PARTYHEAL) + ST1
                Case SKILL_GREATERHEAL
                    B = RunScript("SPELL" & SKILL_GREATERHEAL)
                    ST1 = DoubleChar$(B)
                    For A = 1 To currentMaxUser
                        With player(A)
                            If .map = player(Index).map Then
                                If (CLng(.x) - CLng(player(Index).x)) ^ 2 + (CLng(.y) - CLng(player(Index).y)) ^ 2 <= Skills(SKILL_GREATERHEAL).Range ^ 2 Then
                                    'If .HP < .MaxHP Then
                                        .HP = .HP + B
                                        If .HP > .MaxHP Then .HP = .MaxHP
                                        SendSocket A, Chr2(46) + DoubleChar(.HP)
                                        SendToPartyAllBut .Party, A, Chr2(104) + Chr2(0) + Chr2(A) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                                        SendToGods Chr2(104) + Chr2(0) + Chr2(A) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                                        SendToGuildAllBut A, CLng(.Guild), Chr2(104) + Chr2(0) + Chr2(A) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                                        ST1 = ST1 + Chr2(A)
                                    'End If
                                End If
                            End If
                        End With
                    Next A
                    If Len(ST1) > 0 Then
                        SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(SKILL_GREATERHEAL) + ST1
                    End If
                Case SKILL_EMPOWER
                    If CurrentTarget.TargetType = TT_CHARACTER Then
                        A = Index
                        ClearBuff (Index)
                        CalculateStats Index, False
                    ElseIf CurrentTarget.TargetType = TT_PLAYER Then
                        A = CurrentTarget.Target
                    End If
                    If A > 0 And A <= MaxUsers Then
                        B = RunScript("SPELL" & SKILL_EMPOWER)
                        'If CurrentTarget.TargetType = TT_PLAYER Then b = CInt(b / 1.5)
                        player(A).Buff.Type = BUFF_EMPOWER
                        player(A).Buff.timer = 300
                        player(A).Buff.Data(0) = B
                        SetBuff A
                        CalculateStats A
                    End If
                Case SKILL_WILLPOWER
                    If CurrentTarget.TargetType = TT_CHARACTER Then
                        A = Index
                        ClearBuff (Index)
                        CalculateStats Index, False
                    ElseIf CurrentTarget.TargetType = TT_PLAYER Then
                        A = CurrentTarget.Target
                    End If
                    If A > 0 And A <= MaxUsers Then
                        B = RunScript("SPELL" & SKILL_WILLPOWER)
                        'If CurrentTarget.TargetType = TT_PLAYER Then b = CInt(b / 1.5)
                        player(A).Buff.Type = BUFF_WILLPOWER
                        player(A).Buff.timer = 300
                        player(A).Buff.Data(0) = B
                        SetBuff A
                        CalculateStats A
                    End If
                Case SKILL_ZEAL
                    If CurrentTarget.TargetType = TT_CHARACTER Then
                        A = Index
                        ClearBuff (Index)
                        CalculateStats Index, False
                    ElseIf CurrentTarget.TargetType = TT_PLAYER Then
                        A = CurrentTarget.Target
                    End If
                    If A > 0 And A <= MaxUsers Then
                        B = RunScript("SPELL" & SKILL_ZEAL)
                        'If CurrentTarget.TargetType = TT_PLAYER Then b = CInt(b / 2)
                        player(A).Buff.Type = BUFF_ZEAL
                        player(A).Buff.timer = 300
                        player(A).Buff.Data(0) = B
                        SetBuff A
                        CalculateStats A
                    End If
                Case SKILL_VITALITY
                    'If CurrentTarget.TargetType = TT_CHARACTER Then
                    '    a = Index
                    'ElseIf CurrentTarget.TargetType = TT_PLAYER Then
                    '    a = CurrentTarget.Target
                    'End If
                    If CurrentTarget.TargetType = TT_CHARACTER Then
                        A = Index
                        ClearBuff (Index)
                        CalculateStats Index, False
                    ElseIf CurrentTarget.TargetType = TT_PLAYER Then
                        A = CurrentTarget.Target
                    End If
                    If A > 0 And A <= MaxUsers Then
                        B = RunScript("SPELL" & SKILL_VITALITY)
                        player(A).Buff.Type = BUFF_VITALITY
                        player(A).Buff.timer = 300
                        player(A).Buff.Data(0) = B
                        SetBuff A
                        CalculateStats A
                    End If
                Case SKILL_REGENERATE
                    If CurrentTarget.TargetType = TT_CHARACTER Then
                        A = Index
                    ElseIf CurrentTarget.TargetType = TT_PLAYER Then
                        A = CurrentTarget.Target
                    End If
                    If A > 0 And A <= MaxUsers Then
                        B = RunScript("SPELL" & SKILL_REGENERATE)
                        SetStatusEffect A, SE_REGENERATION
                        player(A).StatusData(SE_REGENERATION).Data(0) = B \ 256
                        player(A).StatusData(SE_REGENERATION).Data(1) = B And 255
                        player(A).StatusData(SE_REGENERATION).timer = 15
                    End If
                Case SKILL_RETRIBUTION
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("Spell" & SKILL_RETRIBUTION)
                        SetStatusEffect Index, SE_RETRIBUTION
                        player(Index).StatusData(SE_RETRIBUTION).Data(0) = B
                        player(Index).StatusData(SE_RETRIBUTION).timer = 30
                    End If
                Case SKILL_HOLYARMOR
                    If CurrentTarget.TargetType = TT_CHARACTER Then
                        A = Index
                        ClearBuff (Index)
                        CalculateStats Index
                    ElseIf CurrentTarget.TargetType = TT_PLAYER Then
                        A = CurrentTarget.Target
                    End If
                    If A > 0 And A <= MaxUsers Then
                        B = RunScript("Spell" & SKILL_HOLYARMOR)
                        
                        player(A).Buff.Type = BUFF_HOLYARMOR
                        player(A).Buff.timer = 300
                        player(A).Buff.Data(0) = B
                        SetBuff A
                        CalculateStats A
                    End If
                Case SKILL_INVULNERABILITY
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        SetStatusEffect Index, SE_INVULNERABILITY
                        player(Index).StatusData(SE_INVULNERABILITY).timer = 5
                    End If
                Case SKILL_BERSERK
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("Spell" & SKILL_BERSERK)
                        SetStatusEffect Index, SE_BERSERK
                        player(Index).StatusData(SE_BERSERK).timer = B
                        CalculateStats Index
                    End If
                
                Case SKILL_NECROMANCY
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = 1 + .SkillLevel(SKILL_NECROMANCY) + ((.Wisdom - .WisMod(1)) \ 20)
                        player(Index).Buff.Type = BUFF_NECROMANCY
                        player(Index).Buff.timer = 300
                        player(Index).Buff.Data(0) = B
                        SetBuff Index
                    End If
                Case SKILL_EVOCATION
                    'If CurrentTarget.TargetType = TT_NO_TARGET Then
                    '    B = 3 + .SkillLevel(SKILL_EVOCATION)
                    '    player(index).Buff.Type = BUFF_EVOCATION
                    '    player(index).Buff.timer = 150
                    '    player(index).Buff.Data(0) = B
                    '    SetBuff index
                    'End If
                Case SKILL_HEAVENSTORM
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("SPELL" & SKILL_HEAVENSTORM)
                        ST1 = AttackAreaEffect(Index, SKILL_HEAVENSTORM, .x, .y, 2, B, True)
                        SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(SKILL_HEAVENSTORM) + Chr2(.x) + Chr2(.y) + ST1
                    End If
                Case SKILL_SIPHONLIFE
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("SPELL" & SKILL_SIPHONLIFE)
                        ST1 = AttackAreaEffect(Index, SKILL_SIPHONLIFE, .x, .y, 2, B, True)
                        SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(SKILL_SIPHONLIFE) + Chr2(.x) + Chr2(.y) + ST1
                    End If
                Case SKILL_SNOWSTORM
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("SPELL" & SKILL_SNOWSTORM)
                        ST1 = AttackAreaEffect(Index, SKILL_SNOWSTORM, .x, .y, 1, B, True)
                        SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(SKILL_SNOWSTORM) + Chr2(.x) + Chr2(.y) + ST1
                    End If
                Case SKILL_TEMPEST
                    B = RunScript("SPELL" & SKILL_TEMPEST)
                    Select Case CurrentTarget.TargetType
                        Case TT_TILE
                            If LOS(.map, .x, .y, CurrentTarget.x, CurrentTarget.y, 0) Then
                                ST1 = AttackAreaEffect(Index, SKILL_TEMPEST, CurrentTarget.x, CurrentTarget.y, 1, B, True)
                                SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(SKILL_SNOWSTORM) + Chr2(CurrentTarget.x) + Chr2(CurrentTarget.y) + ST1
                            End If
'                        Case TT_MONSTER
'                            If CanTargetMonster(Index, SKILL_TEMPEST, CurrentTarget.Target) Then
'                                St1 = AttackAreaEffect(Index, SKILL_TEMPEST, Map(.Map).Monster(CurrentTarget.Target).X, Map(.Map).Monster(CurrentTarget.Target).Y, 1, B)
'                                SendToMap .Map, chr2(111) + chr2(Index) + chr2(SKILL_SNOWSTORM) + chr2(Map(.Map).Monster(CurrentTarget.Target).X) + chr2(Map(.Map).Monster(CurrentTarget.Target).Y) + St1
'                            End If
'                        Case TT_PLAYER
'                            If CanTargetPlayer(Index, SKILL_TEMPEST, CurrentTarget.Target, 0) Then
'                                St1 = AttackAreaEffect(Index, SKILL_TEMPEST, Player(CurrentTarget.Target).X, Player(CurrentTarget.Target).Y, 1, B)
'                                SendToMap .Map, chr2(111) + chr2(Index) + chr2(SKILL_SNOWSTORM) + chr2(Player(CurrentTarget.Target).X) + chr2(Player(CurrentTarget.Target).Y) + St1
'                            End If
                    End Select
                Case SKILL_BLIZZARD
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("SPELL" & SKILL_BLIZZARD)
                        ST1 = AttackAreaEffect(Index, SKILL_BLIZZARD, .x, .y, 2, B, True)
                        SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(SKILL_BLIZZARD) + Chr2(.x) + Chr2(.y) + ST1
                    End If
                Case SKILL_SHATTERSHIELD
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("SPELL" & SKILL_SHATTERSHIELD)
                        SetStatusEffect Index, SE_SHATTERSHIELD
                        .StatusData(SE_SHATTERSHIELD).Data(0) = B \ 256
                        .StatusData(SE_SHATTERSHIELD).Data(1) = B And 255
                        .StatusData(SE_SHATTERSHIELD).timer = 60
                    End If
                Case SKILL_MANARESERVES
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        CalculateStats Index
                        'SetStatusEffect Index, SE_MANARESERVES
                        '.StatusData(SE_MANARESERVES).Data(0) = b \ 256
                        '.StatusData(SE_MANARESERVES).Data(1) = b And 255
                        '.StatusData(SE_MANARESERVES).timer = 60
                    End If
                    
                Case SKILL_MANASHIELD
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("SPELL" & SKILL_MANASHIELD)
                        .Buff.Type = BUFF_MANASHIELD
                        .Buff.Data(0) = B
                        .Buff.timer = 60
                    End If
                        SetBuff Index
                Case SKILL_EVANESCENCE
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        SetStatusEffect Index, SE_EVANESCENCE
                        .StatusData(SE_EVANESCENCE).timer = 8
                    End If
                Case SKILL_FIERYESSENCE
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        SetStatusEffect Index, SE_FIERYESSENCE
                        .StatusData(SE_FIERYESSENCE).timer = 10
                    End If
                Case SKILL_TELEPORT
                    If CurrentTarget.TargetType = TT_TILE Then
                        If CurrentTarget.x >= 0 And CurrentTarget.x <= 11 And CurrentTarget.y >= 0 And CurrentTarget.y <= 11 Then
                            If LOS(.map, .x, .y, CurrentTarget.x, CurrentTarget.y, 1) Then
                                A = .x
                                B = .y
                                WarpPlayer Index, .map, CurrentTarget.x, CurrentTarget.y
                                SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(SKILL) + Chr2(A * 16 + B) + Chr2(CurrentTarget.x * 16 + CurrentTarget.y)
                            End If
                        End If
                    End If
                Case SKILL_PORTAL
                    B = RunScript("SPELL" & SKILL_PORTAL)
                'Case SKILL_WIZARDSBALANCE
                '    If CurrentTarget.TargetType = TT_NO_TARGET Then
                '        A = (.MaxHP \ 10)
                '        B = (.MaxEnergy \ 10)
                '        If .HP > A And .Energy > B Then
                '            .HP = .HP - A
                '            .Energy = .Energy - B
                '            .Mana = .Mana + (.MaxMana * 0.15)
                '            If .Mana > .MaxMana Then .Mana = .MaxMana
                '            SendToMap .Map, chr2(111) + chr2(Index) + chr2(Skill)
                '            St1 = DoubleChar(3) + chr2(46) + DoubleChar(.HP) + DoubleChar(3) + chr2(47) + DoubleChar(.Energy) + DoubleChar(3) + chr2(48) + DoubleChar(.Mana)
                '            SendRaw Index, St1
                '        End If
                '    End If
       
                
                Case SKILL_FLAMEWAVE
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("SPELL" & SKILL_FLAMEWAVE)
                        ST1 = AttackAreaEffect(Index, SKILL_FLAMEWAVE, .x, .y, 3, B, True)
                        SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(SKILL_FLAMEWAVE) + Chr2(.x) + Chr2(.y) + Chr(.D) + ST1
                    End If
                Case SKILL_DRAWESSENCE
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("SPELL" & SKILL_DRAWESSENCE)
                    End If
                Case SKILL_INVISIBILITY
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("SPELL" & SKILL_INVISIBILITY)
                        SetStatusEffect Index, SE_INVISIBLE
                        .StatusData(SE_INVISIBLE).Data(1) = NO_MOVE
                        .StatusData(SE_INVISIBLE).timer = B
                    End If
                Case SKILL_EXPLOSIVETRAP
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("SPELL" & SKILL_EXPLOSIVETRAP)
                        With player(Index)
                            If GetStatusEffect(Index, SE_INVISIBLE) Then
                                CreateTrap .map, .x, .y, 1, B, .trapID, Index, 0
                            Else
                                CreateTrap .map, .x, .y, 1, B, .trapID, Index
                            End If
                        End With
                    End If
                Case SKILL_BEARTRAP
                     If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("SPELL" & SKILL_BEARTRAP)
                        With player(Index)
                            If GetStatusEffect(Index, SE_INVISIBLE) Then
                                CreateTrap .map, .x, .y, 2, B, .trapID, Index, 0
                            Else
                                CreateTrap .map, .x, .y, 2, B, .trapID, Index
                            End If
                        End With
                        'SetStatusEffect index, SE_INVISIBLE
                        '.StatusData(SE_INVISIBLE).Data(1) = NO_MOVE
                        '.StatusData(SE_INVISIBLE).timer = B
                    End If
                'Case SKILL_HIDE
                '    If CurrentTarget.TargetType = TT_NO_TARGET Then
                '        If .HP = .MaxHP Then
                '            Parameter(0) = index
                '            b = RunScript("SPELL" & SKILL_HIDE)
                '            SetStatusEffect index, SE_INVISIBLE
                '            .StatusData(SE_INVISIBLE).Data(1) = NO_MOVE
                '            .StatusData(SE_INVISIBLE).timer = b
                '        End If
                '    End If
                'Case SKILL_HAUNT
                '    LoadTargetData St, CurrentTarget
                '    If CurrentTarget.TargetType = TT_TILE Then
                '        If GetStatusEffect(index, SE_INVISIBLE) Then
                '            If LOS(.map, .x, .y, CurrentTarget.x, CurrentTarget.y, 1) Then
                '                a = .x
                '                b = .y
                '                WarpPlayer index, .map, CurrentTarget.x, CurrentTarget.y
                '                SendSocket index, chr2(111) + chr2(index) + chr2(Skill) + chr2(a * 16 + b) + chr2(CurrentTarget.x * 16 + CurrentTarget.y)
                '            End If
                '        End If
                '    End If
                Case SKILL_STEALTH
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        If .HP = .MaxHP Then
                            B = RunScript("SPELL" & SKILL_STEALTH)
                            SetStatusEffect Index, SE_INVISIBLE
                            .StatusData(SE_INVISIBLE).Data(1) = CAN_MOVE
                            .StatusData(SE_INVISIBLE).timer = B
                        End If
                    End If
                Case SKILL_ETHEREALITY
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("SPELL" & SKILL_ETHEREALITY)
                        SetStatusEffect Index, SE_ETHEREALITY
                        .StatusData(SE_ETHEREALITY).Data(0) = B
                        .StatusData(SE_ETHEREALITY).timer = 5
                    End If
                Case SKILL_SECONDWIND
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("SPELL" & SKILL_SECONDWIND)
                        SetPlayerEnergy Index, .Energy + B
                        SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(SKILL_SECONDWIND)
                    End If
                Case SKILL_DEADLYCLARITY
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("SPELL" & SKILL_DEADLYCLARITY)
                        SetStatusEffect Index, SE_DEADLYCLARITY
                        .StatusData(SE_DEADLYCLARITY).Data(0) = B
                        .StatusData(SE_DEADLYCLARITY).timer = 10
                    End If
                Case SKILL_BLIGHT
                    'If CurrentTarget.TargetType = TT_NO_TARGET Then
                    x = Asc(Mid$(St, 3, 1))
                    y = Asc(Mid$(St, 4, 1))
                    With player(Index)
                    If LOS(.map, .x, .y, x, y, 0) Then
                        Parameter(1) = x
                        Parameter(2) = y
                        B = RunScript("SPELL" & SKILL_BLIGHT)
                        
                        For A = 1 To currentMaxUser
                            If A <> Index Then
                            If IsPlaying(A) Then
                                If .map = player(A).map Then
                                    If Abs(x - player(A).x) <= 1 And Abs(y - player(A).y) <= 1 And player(A).Access = 0 Then
                                        If CanAttackPlayer(Index, A) Then
                                        If LOS(.map, .x, .y, player(A).x, player(A).y, 1) Then
                                            SetStatusEffect A, SE_BLIND
                                            player(A).StatusData(SE_BLIND).timer = 1
                                            player(A).combatTimer = GetTickCount + 30000
                                            
                                            
                                            SetStatusEffect A, SE_POISON
                                            player(A).StatusData(SE_POISON).Data(0) = B
                                            player(A).StatusData(SE_POISON).timer = 10
                                                                            
                                        End If
                                        End If
                                    End If
                                End If
                            End If
                            End If
                        Next A
                        SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(SKILL_BLIGHT) + Chr2(.x) + Chr2(.y) + Chr2(x) + Chr2(y) + ST1
                    End If
                    End With
                        
                    'End If
                'General AttackSkills
                Case SKILL_BACKSTAB, SKILL_HEAVENLYSTRIKE, SKILL_BASH, SKILL_MAIM, SKILL_BLOODLETTER, SKILL_ENVENOM, SKILL_CRIPPLE, SKILL_STEALESSENCE, SKILL_JUDGMENT
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        .AttackSkill = SKILL
                    End If
                Case SKILL_FURY
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("SPELL" & SKILL_FURY)
                        ST1 = AttackAreaEffect(Index, SKILL_FURY, .x, .y, 1, B, False)
                        SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(SKILL_FURY) + Chr2(.x) + Chr2(.y) + ST1
                    End If
                Case SKILL_MASTERDEFENSE
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("SPELL" & SKILL_MASTERDEFENSE)
                        SetStatusEffect Index, SE_MASTERDEFENSE
                        .StatusData(SE_MASTERDEFENSE).timer = B
                    End If
                Case Else
                    If CurrentTarget.TargetType = TT_NO_TARGET Then
                        B = RunScript("SPELL" & SKILL)
                    End If
            End Select
        End If
    End With
End If
End Function

Public Function CanTargetMonster(Index As Long, SKILL As Byte, Target As Long) As Byte
Dim mapNum As Long
mapNum = player(Index).map
CanTargetMonster = 1
If Target < 10 Then
    If map(mapNum).monster(Target).monster > 0 Then
        If ExamineBit(map(mapNum).Flags(0), 5) = False Then
            If (Not SKILL = SKILL_INVALID) And (Skills(SKILL).Flags And SF_USELOS) Then
                If LOS(CLng(player(Index).map), CLng(player(Index).x), CLng(player(Index).y), CLng(map(mapNum).monster(Target).x), CLng(map(mapNum).monster(Target).y), 0) Then
                    CanTargetMonster = 1
                Else
                    CanTargetMonster = 0
                    Exit Function
                End If
            Else
                CanTargetMonster = 1
            End If
            If (Not SKILL = SKILL_INVALID) And (Skills(SKILL).Flags And SF_USERANGE) Then
                If Sqr((CLng(player(Index).x) - CLng(map(mapNum).monster(Target).x)) ^ 2 + (CLng(player(Index).y) - CLng(map(mapNum).monster(Target).y)) ^ 2) <= Skills(SKILL).Range Then
                    CanTargetMonster = 1
                Else
                    CanTargetMonster = 0
                    Exit Function
                End If
            End If
        Else
            CanTargetMonster = 0
        End If
    Else
        CanTargetMonster = 0
    End If
Else
    CanTargetMonster = 0
End If
End Function

Public Function CanTargetPlayer(Index As Long, SKILL As Byte, Target As Long, Friendly As Byte) As Byte
If Friendly Then
    If Target > 0 And Target <= MaxUsers Then
        If Target = Index Then
            CanTargetPlayer = 1
        Else
            With player(Index)
                If .map = player(Target).map Then
                    If .Mode = modePlaying And player(Target).Mode = modePlaying Then
                        If LOS(CLng(.map), CLng(.x), CLng(.y), CLng(player(Index).x), CLng(player(Index).y), 0) Then
                            CanTargetPlayer = 1
                        End If
                    End If
                End If
            End With
        End If
    End If
Else
    If Target > 0 And Target <= MaxUsers Then
        If Target <> Index Then
            With player(Index)
                If ExamineBit(map(.map).Flags(0), 0) = False Then
                    If .Mode = modePlaying And player(Target).Mode = modePlaying Then
                        If .map = player(Target).map Then
                            'If .Access = 0 And Player(Target).Access = 0 Then
                                If .Guild > 0 Or CBool(ExamineBit(map(.map).Flags(0), 6)) Then
                                    If player(Target).Guild > 0 Or CBool(ExamineBit(map(.map).Flags(0), 6)) Then
                                        'If .Guild = 0 Or .Guild <> player(Target).Guild Then
                                            If (Not SKILL = SKILL_INVALID) And (Skills(SKILL).Flags And SF_USELOS) Then
                                                If LOS(CLng(.map), CLng(.x), CLng(.y), CLng(player(Index).x), CLng(player(Index).y), 0) Then
                                                    CanTargetPlayer = 1
                                                Else
                                                    CanTargetPlayer = 0
                                                    Exit Function
                                                End If
                                            Else
                                                CanTargetPlayer = 1
                                            End If
                                            If (Not SKILL = SKILL_INVALID) And (Skills(SKILL).Flags And SF_USERANGE) Then
                                                If Sqr((CLng(.x) - CLng(player(Target).x)) ^ 2 + (CLng(.y) - CLng(player(Target).y)) ^ 2) <= Skills(SKILL).Range Then
                                                    CanTargetPlayer = 1
                                                Else
                                                    CanTargetPlayer = 0
                                                    Exit Function
                                                End If
                                            End If
                                        'End If
                                    End If
                                End If
                            'End If
                        End If
                    End If
                End If
            End With
        End If
    End If
End If
End Function

Function LOS(ByVal mapNum As Integer, ByVal Sourcex As Long, ByVal Sourcey As Long, ByVal Targetx As Long, ByVal Targety As Long, ExtendedLOS As Byte) As Boolean
    Dim A As Long, B As Long
    Dim dX As Long
    Dim dY As Long
    Dim AbsDx As Long
    Dim AbsDy As Long
    
    If (map(mapNum).Tile(Targetx, Targety).Att = 27) Then
        LOS = False
        Exit Function
    End If
    If Sourcex < 0 Or Sourcex > 11 Or Targetx < 0 Or Targetx > 11 Or Sourcey < 0 Or Sourcey > 11 Or Targety < 0 Or Targety > 11 Then
        LOS = False
        Exit Function
    End If
    
    dX = Targetx - Sourcex
    dY = Targety - Sourcey
    AbsDx = Abs(dX)
    AbsDy = Abs(dY)
    
    Dim NumSteps As Long
    If (AbsDx > AbsDy) Then
        NumSteps = AbsDx
    Else
        NumSteps = AbsDy
    End If
    
    If (NumSteps = 0) Then
        LOS = True
        Exit Function
    Else
        Dim CurrentX As Long, LastX As Long
        CurrentX = Sourcex * 64
        LastX = CurrentX
        Dim CurrentY As Long, LastY As Long
        CurrentY = Sourcey * 64
        LastY = CurrentY
        Dim Xincr As Long
        Xincr = (dX * 64) \ NumSteps
        If AbsDx = 3 Then Xincr = Xincr + 1
        Dim Yincr As Long
        Yincr = (dY * 64) \ NumSteps
        If AbsDy = 3 Then Yincr = Yincr + 1
        For B = NumSteps To 0 Step -1
            If CurrentX < 0 Then CurrentX = 0
            If CurrentY < 0 Then CurrentY = 0
            A = map(mapNum).Tile(Int(CurrentX \ 64), Int(CurrentY \ 64)).Att
            If ((ExtendedLOS = 1) And (A = 15 Or A = 14 Or A = 11 Or A = 13)) Or A = 1 Or A = 2 Or A = 3 Or A = 21 Then
                LOS = False
                Exit Function
            End If
            If NotLegalPath(LastX \ 64, LastY \ 64, CurrentX \ 64, CurrentY \ 64, mapNum) Then
                LOS = False
                Exit Function
            End If
            CurrentX = CurrentX + Xincr
            CurrentY = CurrentY + Yincr
        Next
    End If
LOS = True
End Function
Function NotLegalPath(ByVal fX As Long, ByVal fY As Long, ByVal tX As Long, ByVal tY As Long, mapNum As Integer)
Dim FromDir As Long

'0up,1down,2left,3right,4upleft,5upright,6downleft,7downright
If Abs(fX - tX) + Abs(fY - tY) > 0 Then
    If fX > tX Then 'Moving left
        If fY > tY Then 'Moving up
            FromDir = 4
        ElseIf fY < tY Then 'Moving down
            FromDir = 6
        Else
            FromDir = 2
        End If
    ElseIf fX < tX Then 'Moving right
        If fY > tY Then 'Moving up
            FromDir = 5
        ElseIf fY < tY Then 'Moving down
            FromDir = 7
        Else
            FromDir = 3
        End If
    Else
        If fY > tY Then 'moving up
            FromDir = 0
        ElseIf fY < tY Then 'moving down
            FromDir = 1
        End If
    End If
    Select Case FromDir
        Case 1 'Down
            If ExamineBit(map(mapNum).Flags(1), 5) = False Then
                If ExamineBit(map(mapNum).Tile(tX, tY).WallTile, 0) Then GoTo notlegal
            End If
            If ExamineBit(map(mapNum).Tile(fX, fY).WallTile, 5) Then GoTo notlegal
        Case 0 'Up
            If ExamineBit(map(mapNum).Flags(1), 5) = False Then
                If ExamineBit(map(mapNum).Tile(tX, tY).WallTile, 1) Then GoTo notlegal
            End If
            If ExamineBit(map(mapNum).Tile(fX, fY).WallTile, 4) Then GoTo notlegal
        Case 3 'Right
            If ExamineBit(map(mapNum).Flags(1), 5) = False Then
                If ExamineBit(map(mapNum).Tile(tX, tY).WallTile, 2) Then GoTo notlegal
            End If
            If ExamineBit(map(mapNum).Tile(fX, fY).WallTile, 7) Then GoTo notlegal
        Case 2 'Left
            If ExamineBit(map(mapNum).Flags(1), 5) = False Then
                If ExamineBit(map(mapNum).Tile(tX, tY).WallTile, 3) Then GoTo notlegal
            End If
            If ExamineBit(map(mapNum).Tile(fX, fY).WallTile, 6) Then GoTo notlegal
        Case 4 'Up Left
            'try going up and then left
            If fY > 0 Then
                If ExamineBit(map(mapNum).Tile(fX, fY - 1).WallTile, 1) = 0 Or ExamineBit(map(mapNum).Flags(1), 5) Then
                    If ExamineBit(map(mapNum).Tile(fX, fY).WallTile, 4) = 0 Then
                        'Up is good, try left
                        If ExamineBit(map(mapNum).Tile(tX, tY).WallTile, 3) = 0 Or ExamineBit(map(mapNum).Flags(1), 5) Then
                            If ExamineBit(map(mapNum).Tile(fX, fY - 1).WallTile, 6) = 0 Then 'Up/Left is good
                                NotLegalPath = False
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
            'try going left and then up
            If fX > 0 Then
                If ExamineBit(map(mapNum).Tile(fX - 1, fY).WallTile, 3) = 0 Or ExamineBit(map(mapNum).Flags(1), 5) Then
                    If ExamineBit(map(mapNum).Tile(fX, fY).WallTile, 6) = 0 Then
                        'left is good, try up
                        If ExamineBit(map(mapNum).Tile(tX, tY).WallTile, 1) = 0 Or ExamineBit(map(mapNum).Flags(1), 5) Then
                            If ExamineBit(map(mapNum).Tile(fX - 1, fY).WallTile, 4) = 0 Then
                                NotLegalPath = False
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
            NotLegalPath = True
            Exit Function
        Case 5 'Up Right
            'try going up and then right
            If fY > 0 Then
                If ExamineBit(map(mapNum).Tile(fX, fY - 1).WallTile, 1) = 0 Or ExamineBit(map(mapNum).Flags(1), 5) Then
                    If ExamineBit(map(mapNum).Tile(fX, fY).WallTile, 4) = 0 Then
                        'Up is good, try right
                        If ExamineBit(map(mapNum).Tile(tX, tY).WallTile, 2) = 0 Or ExamineBit(map(mapNum).Flags(1), 5) Then
                            If ExamineBit(map(mapNum).Tile(fX, fY - 1).WallTile, 7) = 0 Then 'Up/Left is good
                                NotLegalPath = False
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
            'try going right and then up
            If fX < 11 Then
                If ExamineBit(map(mapNum).Tile(fX + 1, fY).WallTile, 2) = 0 Or ExamineBit(map(mapNum).Flags(1), 5) Then
                    If ExamineBit(map(mapNum).Tile(fX, fY).WallTile, 7) = 0 Then
                        'right is good, try up
                        If ExamineBit(map(mapNum).Tile(tX, tY).WallTile, 1) = 0 Or ExamineBit(map(mapNum).Flags(1), 5) Then
                            If ExamineBit(map(mapNum).Tile(fX + 1, fY).WallTile, 4) = 0 Then
                                NotLegalPath = False
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
            NotLegalPath = True
            Exit Function
        Case 6 'Down Left
            'try going down and then left
            If fY < 11 Then
                If ExamineBit(map(mapNum).Tile(fX, fY + 1).WallTile, 0) = 0 Or ExamineBit(map(mapNum).Flags(1), 5) Then
                    If ExamineBit(map(mapNum).Tile(fX, fY).WallTile, 5) = 0 Then
                        'down is good, try left
                        If ExamineBit(map(mapNum).Tile(tX, tY).WallTile, 3) = 0 Or ExamineBit(map(mapNum).Flags(1), 5) Then
                            If ExamineBit(map(mapNum).Tile(fX, fY + 1).WallTile, 6) = 0 Then 'down/Left is good
                                NotLegalPath = False
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
            'try going left and then down
            If fX > 0 Then
                If ExamineBit(map(mapNum).Tile(fX - 1, fY).WallTile, 3) = 0 Or ExamineBit(map(mapNum).Flags(1), 5) Then
                    If ExamineBit(map(mapNum).Tile(fX, fY).WallTile, 6) = 0 Then
                        'left is good, try down
                        If ExamineBit(map(mapNum).Tile(tX, tY).WallTile, 0) = 0 Or ExamineBit(map(mapNum).Flags(1), 5) Then
                            If ExamineBit(map(mapNum).Tile(fX - 1, fY).WallTile, 5) = 0 Then
                                NotLegalPath = False
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
            NotLegalPath = True
            Exit Function
        Case 7 'Down/Right
            'try going down and then right
            If fY > 0 Then
                If ExamineBit(map(mapNum).Tile(fX, fY - 1).WallTile, 0) = 0 Or ExamineBit(map(mapNum).Flags(1), 5) Then
                    If ExamineBit(map(mapNum).Tile(fX, fY).WallTile, 5) = 0 Then
                        'down is good, try right
                        If ExamineBit(map(mapNum).Tile(tX, tY).WallTile, 2) = 0 Or ExamineBit(map(mapNum).Flags(1), 5) Then
                            If ExamineBit(map(mapNum).Tile(fX, fY - 1).WallTile, 7) = 0 Then 'Up/Left is good
                                NotLegalPath = False
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
            'try going right and then down
            If fX < 11 Then
                If ExamineBit(map(mapNum).Tile(fX + 1, fY).WallTile, 2) = 0 Or ExamineBit(map(mapNum).Flags(1), 5) Then
                    If ExamineBit(map(mapNum).Tile(fX, fY).WallTile, 7) = 0 Then
                        'right is good, try down
                        If ExamineBit(map(mapNum).Tile(tX, tY).WallTile, 0) = 0 Or ExamineBit(map(mapNum).Flags(1), 5) Then
                            If ExamineBit(map(mapNum).Tile(fX + 1, fY).WallTile, 5) = 0 Then
                                NotLegalPath = False
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
            NotLegalPath = True
            Exit Function
    End Select
    
    NotLegalPath = False
    Exit Function
End If

NotLegalPath = False
Exit Function
notlegal:
NotLegalPath = True
End Function

Public Sub LoadTargetData(St As String, ByRef Target As TargetData)
If Len(St) = 4 Then
    Target.TargetType = Asc(Mid$(St, 1, 1))
    Target.Target = Asc(Mid$(St, 2, 1))
    Target.x = Asc(Mid$(St, 3, 1))
    Target.y = Asc(Mid$(St, 4, 1))
End If
End Sub

Public Sub GainSkillEXP(Index As Long, SKILL As Byte) 'EXP as long (damage/2)
With player(Index)
    If .SkillLevel(SKILL) < Skills(SKILL).MaxLevel Then
        '.SkillEXP(Skill) = .SkillEXP(Skill) + EXP
        If .SkillEXP(SKILL) > SkillEXPTable(Skills(SKILL).EXPTable, .SkillLevel(SKILL) + 1) Then
            .SkillEXP(SKILL) = 0
            .SkillLevel(SKILL) = .SkillLevel(SKILL) + 1
            SendSocket Index, Chr2(123) + Chr2(1) + Chr2(SKILL) + Chr2(.SkillLevel(SKILL))
            SendSocket Index, Chr2(123) + Chr(SKILL) + QuadChar(0)
        Else
            SendSocket Index, Chr2(123) + Chr(SKILL) + QuadChar(.SkillEXP(SKILL))
        End If
    End If
End With
End Sub

Public Sub InitSkillEXPTable()
Dim A As Long
For A = 1 To 50
    SkillEXPTable(ET_STANDARD, A) = (Int(A ^ 1.6) * 100)
Next A
End Sub

Public Function GetSkillManaCost(Index As Long, SKILL As Byte) As Long

    If SKILL > SKILL_INVALID And SKILL <= MAX_SKILLS Then
        With Skills(SKILL)
            GetSkillManaCost = (.CostPerLevel * player(Index).Level) + (.CostPerSLevel * player(Index).SkillLevel(SKILL)) + .CostConstant
        End With
    End If
    
End Function

Sub SetStatusEffect(ByVal Index As Long, StatusEffect As Long)
    If StatusEffect < 32 Then
    Dim A As Long
    player(Index).StatusEffect = player(Index).StatusEffect Or (2 ^ StatusEffect)
    SendToMap player(Index).map, Chr2(112) + Chr2(Index) + QuadChar(player(Index).StatusEffect)
    If (StatusEffect And SE_INVISIBLE) Then
        For A = 0 To 9
            If map(player(Index).map).monster(A).Target = Index Then
                If map(player(Index).map).monster(A).TargType = TT_PLAYER Then
                    map(player(Index).map).monster(A).Target = 0
                    map(player(Index).map).monster(A).TargType = TT_NO_TARGET
                End If
            End If
        Next A
    End If
    If StatusEffect = SE_POISON Then CreateFloatingEvent player(Index).map, player(Index).x, player(Index).y, FT_POISON
    If StatusEffect = SE_EXHAUST Then CreateFloatingEvent player(Index).map, player(Index).x, player(Index).y, FT_EXHAUST
    End If
End Sub
Function GetStatusEffect(ByVal Index As Long, StatusEffect As Long) As Byte
    If (player(Index).StatusEffect And (2 ^ StatusEffect)) Then
        GetStatusEffect = 1
        Exit Function
    End If
    GetStatusEffect = 0
End Function

Public Sub RemoveStatusEffect(ByVal Index As Long, StatusEffect As Long)
Dim A As Long
    With player(Index)
        If StatusEffect = SE_ALL Then
            .StatusEffect = 0
            For A = 1 To MAXSTATUS
                .StatusData(A).Data(0) = 0
                .StatusData(A).Data(1) = 0
                .StatusData(A).Data(2) = 0
                .StatusData(A).Data(3) = 0
                .StatusData(A).timer = 0
            Next A
        Else
            .StatusEffect = .StatusEffect And Not (2 ^ StatusEffect)
            .StatusData(StatusEffect).Data(0) = 0
            .StatusData(StatusEffect).Data(1) = 0
            .StatusData(StatusEffect).Data(2) = 0
            .StatusData(StatusEffect).Data(3) = 0
            .StatusData(StatusEffect).timer = 0
        End If
        
        If StatusEffect = SE_ALL Or StatusEffect = SE_BERSERK Or (StatusEffect >= SE_ENERGYMOD And StatusEffect <= SE_MPREGENMOD) Then
            CalculateStats Index
        End If
        
        SendToMap .map, Chr2(112) + Chr2(Index) + QuadChar(.StatusEffect)
    End With
End Sub

Sub SetBuff(ByVal Index As Long)
    SendToMap player(Index).map, Chr2(115) + Chr2(Index) + Chr2(player(Index).Buff.Type)
End Sub

Sub ClearBuff(ByVal Index As Long)
    player(Index).Buff.Type = 0
    CalculateStats Index
    SendToMap player(Index).map, Chr2(115) + Chr2(Index) + Chr2(0)
End Sub

Public Function AttackAreaEffect(Index As Long, SKILL As Byte, x As Byte, y As Byte, AreaType As Byte, damage As Long, Magic As Boolean) As String
    Dim A As Long, NumPoints As Long, B As Long, St As String
    Dim PointList() As Point
    Select Case AreaType
        Case 1
            NumPoints = 8
            ReDim PointList(0 To NumPoints)         '-0-
            PointList(0).x = x                      '123
            PointList(0).y = y - 1                  '-4-
            PointList(1).x = x - 1
            PointList(1).y = y
            PointList(2).x = x
            PointList(2).y = y
            PointList(3).x = x + 1
            PointList(3).y = y
            PointList(4).x = x
            PointList(4).y = y + 1
            
            PointList(5).x = x - 1
            PointList(5).y = y - 1
            PointList(6).x = x - 1
            PointList(6).y = y + 1
            PointList(7).x = x + 1
            PointList(7).y = y - 1
            PointList(8).x = x + 1
            PointList(8).y = y + 1
            
        Case 2
            NumPoints = 12
            ReDim PointList(0 To NumPoints)
            PointList(0).x = x              'OOXOO
            PointList(0).y = y - 2          'OXXXO
            PointList(1).x = x - 1          'XXXXX
            PointList(1).y = y - 1          'OXXXO
            PointList(2).x = x              'OOXOO
            PointList(2).y = y - 1
            PointList(3).x = x + 1
            PointList(3).y = y - 1
            PointList(4).x = x - 2
            PointList(4).y = y
            PointList(5).x = x - 1
            PointList(5).y = y
            PointList(6).x = x
            PointList(6).y = y
            PointList(7).x = x + 1
            PointList(7).y = y
            PointList(8).x = x + 2
            PointList(8).y = y
            PointList(9).x = x - 1
            PointList(9).y = y + 1
            PointList(10).x = x
            PointList(10).y = y + 1
            PointList(11).x = x + 1
            PointList(11).y = y + 1
            PointList(12).x = x
            PointList(12).y = y + 2
        Case 3 'Directional
            NumPoints = 4
            ReDim PointList(0 To NumPoints)
            Select Case player(Index).D
                Case 0 'Up
                    PointList(0).x = x
                    PointList(0).y = y - 1
                    PointList(1).x = x
                    PointList(1).y = y - 2
                    PointList(2).x = x
                    PointList(2).y = y - 3
                    PointList(3).x = x
                    PointList(3).y = y - 4
                Case 1 'Down
                    PointList(0).x = x
                    PointList(0).y = y + 1
                    PointList(1).x = x
                    PointList(1).y = y + 2
                    PointList(2).x = x
                    PointList(2).y = y + 3
                    PointList(3).x = x
                    PointList(3).y = y + 4
                Case 2 'Left
                    PointList(0).x = x - 1
                    PointList(0).y = y
                    PointList(1).x = x - 2
                    PointList(1).y = y
                    PointList(2).x = x - 3
                    PointList(2).y = y
                    PointList(3).x = x - 4
                    PointList(3).y = y
                Case 3 'Right
                    PointList(0).x = x + 1
                    PointList(0).y = y
                    PointList(1).x = x + 2
                    PointList(1).y = y
                    PointList(2).x = x + 3
                    PointList(2).y = y
                    PointList(3).x = x + 4
                    PointList(3).y = y
            End Select
    End Select
    Dim damagecounter As Long
    damagecounter = 0
    For A = 0 To NumPoints
        With PointList(A)
            If .x >= 0 And .x <= 11 And .y >= 0 And .y <= 11 Then
                For B = 0 To 9
                    If map(player(Index).map).monster(B).x = .x And map(player(Index).map).monster(B).y = .y Then
                        If CanTargetMonster(Index, SKILL, B) Then
                            If LOS(CLng(player(Index).map), CLng(x), CLng(y), .x, .y, 0) Then
                                AttackMonster Index, B, damage, Magic, True
                                damagecounter = damagecounter + 1
                                GainSkillEXP Index, SKILL
                                St = St & Chr2(.x) + Chr2(.y) + DoubleChar(CInt(damage))
                            End If
                        End If
                    End If
                    If map(player(Index).map).monster(B).monster > 0 Then
                        If monster(map(player(Index).map).monster(B).monster).Flags2 And MONSTER_LARGE Then
                            If map(player(Index).map).monster(B).x + 1 = .x And map(player(Index).map).monster(B).y = .y Then
                                If CanTargetMonster(Index, SKILL, B) Then
                                    If LOS(CLng(player(Index).map), CLng(x), CLng(y), .x + 1, .y, 0) Then
                                        AttackMonster Index, B, damage, Magic, True
                                        GainSkillEXP Index, SKILL
                                        damagecounter = damagecounter + 1
                                        St = St & Chr2(.x) + Chr2(.y) + DoubleChar(CInt(damage))
                                    End If
                                End If
                            End If
                            If map(player(Index).map).monster(B).x + 1 = .x And map(player(Index).map).monster(B).y + 1 = .y Then
                                If CanTargetMonster(Index, SKILL, B) Then
                                    If LOS(CLng(player(Index).map), CLng(x), CLng(y), .x + 1, .y + 1, 0) Then
                                        AttackMonster Index, B, damage, Magic, True
                                        damagecounter = damagecounter + 1
                                        GainSkillEXP Index, SKILL
                                        St = St & Chr2(.x) + Chr2(.y) + DoubleChar(CInt(damage))
                                    End If
                                End If
                            End If
                            If map(player(Index).map).monster(B).x = .x And map(player(Index).map).monster(B).y + 1 = .y Then
                                If CanTargetMonster(Index, SKILL, B) Then
                                    If LOS(CLng(player(Index).map), CLng(x), CLng(y), .x, .y + 1, 0) Then
                                        AttackMonster Index, B, damage, Magic, True
                                        damagecounter = damagecounter + 1
                                        GainSkillEXP Index, SKILL
                                        St = St & Chr2(.x) + Chr2(.y) + DoubleChar(CInt(damage))
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next B
                For B = 1 To currentMaxUser
                    If player(B).x = .x And player(B).y = .y And player(B).map = player(Index).map Then
                        If CanTargetPlayer(Index, SKILL, B, 0) Then
                            If LOS(CLng(player(Index).map), CLng(x), CLng(y), .x, .y, 0) Then
                                AttackPlayer Index, B, damage, True, Magic, True, True, False, False, False
                                damagecounter = damagecounter + 1
                                GainSkillEXP Index, SKILL
                                St = St & Chr2(.x) + Chr2(.y) + DoubleChar(CInt(damage))
                            End If
                        End If
                    End If
                Next B
            End If
        End With
    Next A
    If SKILL = SKILL_SIPHONLIFE Then
            SetStatusEffect Index, SE_REGENERATION
            player(Index).StatusData(SE_REGENERATION).Data(0) = ((damagecounter * damage / 2) / 15) \ 256
            player(Index).StatusData(SE_REGENERATION).Data(1) = ((damagecounter * damage / 2) / 15) And 255
            player(Index).StatusData(SE_REGENERATION).timer = 15
    End If
    AttackAreaEffect = St
End Function

Public Sub AttackSkillPlayer(Index As Long, Target As Long)
Dim A As Long, B As Long
    Parameter(0) = Index
    Parameter(1) = TT_PLAYER
    Parameter(2) = Target
    Parameter(3) = player(Target).x
    Parameter(4) = player(Target).y
    Parameter(5) = player(Index).AttackSkill
    A = RunScript("SPELL" & player(Index).AttackSkill)
    Select Case player(Index).AttackSkill
        Case SKILL_JUDGMENT
            A = A + PlayerDamage(Index)
            If A < 0 Then A = 0
            If A > 9999 Then A = 9999
            SendToMap player(Index).map, Chr2(111) + Chr2(Index) + Chr2(player(Index).AttackSkill) + Chr2(Target) + Chr2(TT_PLAYER)
            AttackPlayer Index, Target, A, True, True
            
            SetStatusEffect Index, SE_REGENERATION
            player(Index).StatusData(SE_REGENERATION).Data(0) = CInt(A / 15 * 0.8) \ 256
            player(Index).StatusData(SE_REGENERATION).Data(1) = CInt(A / 15 * 0.8) And 255
            player(Index).StatusData(SE_REGENERATION).timer = 15

        Case SKILL_HEAVENLYSTRIKE
            A = A + PlayerDamage(Index)
            If A < 0 Then A = 0
            If A > 9999 Then A = 9999
            SendToMap player(Index).map, Chr2(111) + Chr2(Index) + Chr2(player(Index).AttackSkill) + Chr2(Target) + Chr2(TT_PLAYER)
            AttackPlayer Index, Target, A, True, True
        Case SKILL_BASH, SKILL_MAIM
            A = A + PlayerDamage(Index)
            If A < 0 Then A = 0
            If A > 9999 Then A = 9999
            SendToMap player(Index).map, Chr2(111) + Chr2(Index) + Chr2(player(Index).AttackSkill) + Chr2(Target) + Chr2(TT_PLAYER)
            AttackPlayer Index, Target, A, True, False
        Case SKILL_BLOODLETTER
            B = A
            A = A + PlayerDamage(CLng(Index))
            AttackPlayer Index, Target, A, True
            With player(Index)
                If B > player(Index).HP Then
                    SetPlayerHP Index, 1
                Else
                    SetPlayerHP Index, .HP - B
                End If
                SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(SKILL_BLOODLETTER) + Chr2(Target) + DoubleChar$(B) + Chr2(TT_PLAYER)
            End With
        Case SKILL_BACKSTAB
            'If player(index).D = player(Target).D Then
                If GetStatusEffect(Index, SE_INVISIBLE) Then
                    If GetStatusEffect(Index, SE_INVISIBLE) Then RemoveStatusEffect Index, SE_INVISIBLE
                    Parameter(0) = Index
                    Parameter(1) = TT_PLAYER
                    Parameter(2) = Target
                    Parameter(3) = player(Target).x
                    Parameter(4) = player(Target).y
                    Parameter(5) = player(Index).AttackSkill
                    A = RunScript("SPELL" & SKILL_BACKSTAB)
                    If player(Index).D <> player(Target).D Then A = (A * 3) / 4
                    If (player(Index).D = 0 And player(Target).D = 1) Or (player(Index).D = 1 And player(Target).D = 0) Then A = (A * 2) / 5
                    If (player(Index).D = 2 And player(Target).D = 3) Or (player(Index).D = 3 And player(Target).D = 2) Then A = (A * 2) / 5
                    
                    A = A + PlayerDamage(CLng(Index))
                    If A < 0 Then A = 0
                    If A > 9999 Then A = 9999
                    SendToMap player(Index).map, Chr2(111) + Chr2(Index) + Chr2(SKILL_BACKSTAB) + Chr2(Target) + Chr2(TT_PLAYER)
                    AttackPlayer Index, Target, A, True
                End If
            'End If
        Case SKILL_ENVENOM
            SetStatusEffect Target, SE_POISON
            player(Target).StatusData(SE_POISON).Data(0) = A ' \ 256
         '   player(Target).StatusData(SE_POISON).Data(1) = a And 255
            player(Target).StatusData(SE_POISON).timer = 10
            A = PlayerDamage(Index)
            AttackPlayer Index, Target, A, True
        Case SKILL_CRIPPLE
            SetStatusEffect Target, SE_CRIPPLE
            player(Target).StatusData(SE_CRIPPLE).Data(0) = A '\ 256
         '   player(Target).StatusData(SE_CRIPPLE).Data(1) = a ' And 255
            player(Target).StatusData(SE_CRIPPLE).timer = 8
            A = PlayerDamage(Index)
            SendToMap player(Index).map, Chr2(111) + Chr2(Index) + Chr2(player(Index).AttackSkill) + Chr2(Target) + Chr2(TT_PLAYER)
            AttackPlayer Index, Target, A, True
        Case SKILL_STEALESSENCE
            
            SetStatusEffect Index, SE_MANAMULT
            player(Index).StatusData(SE_MANAMULT).Data(0) = 30
            player(Index).StatusData(SE_MANAMULT).timer = A
            A = PlayerDamage(Index)
            'SetStatusEffect Target, SE_MANAMULT
            'Player(Target).StatusData(SE_MANAMULT).Data(0) = 5
            'Player(Target).StatusData(SE_MANAMULT).Timer = A / 2
            AttackPlayer Index, Target, A, True
        
    End Select
    player(Index).AttackSkill = SKILL_INVALID
End Sub
 
Public Sub AttackSkillMonster(Index As Long, Target As Byte)
Dim A As Long, B As Long
With player(Index)
    Parameter(0) = Index
    Parameter(1) = TT_MONSTER
    Parameter(2) = Target
    Parameter(3) = map(.map).monster(Target).x
    Parameter(4) = map(.map).monster(Target).y
    Parameter(5) = .AttackSkill
    
    A = RunScript("SPELL" & player(Index).AttackSkill)
    Select Case .AttackSkill
        Case SKILL_HEAVENLYSTRIKE, SKILL_BASH, SKILL_MAIM
            A = A + PlayerDamage(CLng(Index))
            A = A - monster(map(.map).monster(Target).monster).Armor
            If A < 0 Then A = 0
            If A > 9999 Then A = 9999
            SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(.AttackSkill) + Chr2(Target) + Chr2(TT_MONSTER)
            AttackMonster Index, Target, A, IIf(.AttackSkill = SKILL_HEAVENLYSTRIKE, True, False), True
            
        Case SKILL_JUDGMENT
            A = A + PlayerDamage(Index)
            A = A - monster(map(.map).monster(Target).monster).Armor
            If A < 0 Then A = 0
            If A > 9999 Then A = 9999
            SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(.AttackSkill) + Chr2(Target) + Chr2(TT_MONSTER)
            AttackMonster Index, Target, A, True, True

            SetStatusEffect Index, SE_REGENERATION
            player(Index).StatusData(SE_REGENERATION).Data(0) = CInt(A / 15 * 0.8) \ 256
            player(Index).StatusData(SE_REGENERATION).Data(1) = CInt(A / 15 * 0.8) And 255
            player(Index).StatusData(SE_REGENERATION).timer = 15
            
        Case SKILL_BLOODLETTER
            B = A
            If .HP - B < 1 Then
                SetPlayerHP Index, 1
            Else
                SetPlayerHP Index, .HP - B
            End If
            A = A + PlayerDamage(CLng(Index))
            A = A - monster(map(.map).monster(Target).monster).Armor
            If A < 0 Then A = 0
            If A > 9999 Then A = 9999
            SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(SKILL_BLOODLETTER) + Chr2(Target) + DoubleChar$(B) + Chr2(TT_MONSTER)
            AttackMonster Index, Target, A, False, True
        Case SKILL_BACKSTAB
            'If .D = map(.map).monster(Target).D Then
             
                If GetStatusEffect(Index, SE_INVISIBLE) Then
                    If GetStatusEffect(Index, SE_INVISIBLE) Then RemoveStatusEffect Index, SE_INVISIBLE
                    Parameter(0) = Index
                    Parameter(1) = TT_MONSTER
                    Parameter(2) = Target
                    Parameter(3) = map(.map).monster(Target).x
                    Parameter(4) = map(.map).monster(Target).y
                    Parameter(5) = .AttackSkill
                    A = RunScript("SPELL" & SKILL_BACKSTAB)
                    If .D <> map(.map).monster(Target).D Then A = (A * 3) / 4
                    A = A + PlayerDamage(CLng(Index))
                    A = A - monster(map(.map).monster(Target).monster).Armor
                    If A < 0 Then A = 0
                    If A > 9999 Then A = 9999
                    SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(SKILL_BACKSTAB) + Chr2(Target) + Chr2(TT_MONSTER)
                    AttackMonster Index, Target, A, False, True
                End If
            'End If
        Case SKILL_ENVENOM, SKILL_CRIPPLE
            'Can't Poison Monsters (Yet)
            A = PlayerDamage(Index)
            A = A - monster(map(.map).monster(Target).monster).Armor
            If A < 0 Then A = 0
            If A > 9999 Then A = 9999
            AttackMonster Index, Target, A, False, True
            SendToMap .map, Chr2(111) + Chr2(Index) + Chr2(.AttackSkill) + Chr2(Target) + Chr2(TT_MONSTER)
        Case SKILL_STEALESSENCE
            
            SetStatusEffect Index, SE_MANAMULT
            player(Index).StatusData(SE_MANAMULT).Data(0) = 30
            player(Index).StatusData(SE_MANAMULT).timer = A
            A = PlayerDamage(Index)
            A = A - monster(map(.map).monster(Target).monster).Armor
            'SetStatusEffect Target, SE_MANAMULT
            'Player(Target).StatusData(SE_MANAMULT).Data(0) = 5
            'Player(Target).StatusData(SE_MANAMULT).Timer = A / 2
            AttackMonster Index, Target, A, False, True
    End Select
    .AttackSkill = SKILL_INVALID
End With
End Sub
