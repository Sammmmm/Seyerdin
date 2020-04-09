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
Public Const TT_NPC = 32

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

'StautsEffect Constants
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
            SE_MANARESERVES = 14, _
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
    Skill As Byte
    Level As Byte
End Type

'Skill information
Type SkillType
    Name As String
    Description As String
    Class As Integer
    Level(1 To 10) As Byte
    CostPerLevel As Single
    CostPerSLevel As Integer
    CostConstant As Integer
    TargetType As Byte
    Type As Long
    Flags As Long
    Range As Byte
    MaxLevel As Byte
    EXPTable As Byte
    Requirements(1) As SkillRequirementType
    LocalTick As Long
    GlobalTick As Long
    Icon As Byte
    Color As Long
End Type

'Target Information
Type TargetData
    TargetType As Byte
    Target As Byte
    x As Byte
    y As Byte
End Type

Public Const MAX_SKILLS = 120

Public Skills(MAX_SKILLS) As SkillType
Public CurrentTarget As TargetData
Public TargetPulse As Long
Public LastTab As Long

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
    SKILL_WIZARDSBALANCE = 28, SKILL_STEALESSENCE = 28, _
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

Global CastingSpell As Boolean

Public Sub InitSkills()

Dim St As String * 331, A As Long, b As Long
Dim tByte(1 To 4) As Byte
Dim CurrentSkill As SkillType

Open App.Path & "/Data/Cache/skilldata" + IIf(ServerHasCustomSkilldata, ServerId, "") + ".dat" For Binary As #1
    If LOF(1) Mod 331 = 0 And LOF(1) > 0 Then
        A = LOF(1) / Len(CurrentSkill)
        For b = 1 To MAX_SKILLS
            Get #1, , St
            
            With Skills(b)
                .Name = ClipString$(Mid$(St, 1, 32))
                .Description = ClipString$(Mid$(St, 33, 256))
                
                .Class = GetInt(Mid$(St, 289, 2))
                .Icon = Asc(Mid$(St, 291, 1))
                .Color = RGB(Asc(Mid$(St, 292, 1)), Asc(Mid$(St, 293, 1)), Asc(Mid$(St, 294, 1)))
                .TargetType = Asc(Mid$(St, 295, 1))
                .Type = Asc(Mid$(St, 296, 1)) * 16777216 + Asc(Mid$(St, 297, 1)) * 65536 + Asc(Mid$(St, 298, 1)) * 256& + Asc(Mid$(St, 299, 1))
                .Flags = Asc(Mid$(St, 300, 1)) * 16777216 + Asc(Mid$(St, 301, 1)) * 65536 + Asc(Mid$(St, 302, 1)) * 256& + Asc(Mid$(St, 303, 1))
                .Range = Asc(Mid$(St, 304, 1))
                .MaxLevel = Asc(Mid$(St, 305, 1))
                .EXPTable = Asc(Mid$(St, 306, 1))
                .Requirements(0).Skill = Asc(Mid$(St, 307, 1))
                .Requirements(0).Level = Asc(Mid$(St, 308, 1))
                .Requirements(1).Skill = Asc(Mid$(St, 309, 1))
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
        Next b
    End If
Close #1

    InitSkillEXPTable
End Sub

Public Function CanUseSkill(ByVal Skill As Byte, DisplayError As Byte, NoStatCheck As Byte, Optional TargetCheck As Byte = 1) As Byte
Dim SkillError As Byte, A As Byte

If TargetCheck = 1 Then
    With CurrentTarget
        Select Case .TargetType
            Case TT_PLAYER
                If .Target > 0 Then
                    .x = player(.Target).x
                    .y = player(.Target).y
                    If Skills(Skill).Flags And SF_HOSTILE Then
                        If map.Tile(.x, .y).Att = 21 Or map.Tile(cX, cY).Att = 21 Then
                            SkillError = 12
                            GoTo RaiseSkillError
                        End If
                    End If
                End If
            Case TT_MONSTER
                .x = map.Monster(.Target).x
                .y = map.Monster(.Target).y
        End Select
    End With
End If

With Character
    If GetStatusEffect(Character.Index, SE_MUTE) = 0 Then
        If (2 ^ (.Class - 1) And Skills(Skill).Class) Then
            If .Level >= Skills(Skill).Level(.Class) Then
                If NoStatCheck Or (.Mana >= ((Skills(Skill).CostPerLevel * .Level) + (Skills(Skill).CostPerSLevel * .SkillLevels(Skill)) + Skills(Skill).CostConstant)) Then
                    If NoStatCheck Or GetTickCount > .GlobalSpellTick Then
                        If NoStatCheck Or GetTickCount >= .LocalSpellTick(Skill) Then
                            If (Skills(Skill).Flags And SF_UNKNOWN) Then
                                If .SkillLevels(Skill) > 0 Then
                                    CanUseSkill = 0
                                Else
                                    SkillError = 10
                                    GoTo RaiseSkillError
                                End If
                            Else
                                CanUseSkill = 0
                            End If
                            For A = 0 To 1
                                With Skills(Skill).Requirements(A)
                                    If .Skill > SKILL_INVALID And .Skill < MAX_SKILLS Then
                                        If Character.SkillLevels(.Skill) < .Level Then
                                            SkillError = 11
                                            GoTo RaiseSkillError
                                        End If
                                    End If
                                End With
                            Next A
                        Else
                            SkillError = 7
                            GoTo RaiseSkillError
                        End If
                    Else
                        SkillError = 7
                        GoTo RaiseSkillError
                    End If
                Else
                    SkillError = 8
                    GoTo RaiseSkillError
                End If
            Else
                SkillError = 1
                GoTo RaiseSkillError
            End If
        Else
            SkillError = 2
            GoTo RaiseSkillError
        End If
    Else
        SkillError = 9
        GoTo RaiseSkillError
    End If
    
    If Skill = SKILL_STEALTH And .HP < .MaxHP Then
        SkillError = 13
        GoTo RaiseSkillError
    End If
    
    
End With

CanUseSkill = 0
Exit Function

RaiseSkillError:
If DisplayError Then
    Select Case SkillError
        Case 1
            If spamdelay = 0 Then
                spamdelay = 2
                PrintChat "You are not high enough level to use this.", 8, Options.FontSize
            End If
        Case 2 To 5
            'PrintChat "Your class cannot use this.", 8, Options.FontSize
        Case 6
            PrintChat "Insufficient HP.", 8, Options.FontSize
        Case 7
            'PrintChat "You must wait to cast this again.", 8, Options.FontSize
        Case 8
            If spamdelay = 0 Then
                spamdelay = 1
                PrintChat "Insufficient Mana.", 8, Options.FontSize
            End If
        Case 9
            If spamdelay = 0 Then
                spamdelay = 1
                PrintChat "You can not cast spells at this time. [Mute]", 8, Options.FontSize
            End If
        Case 10
            If spamdelay = 0 Then
                spamdelay = 2
                PrintChat "You have not learned this skill!", 8, Options.FontSize
            End If
        Case 11
            If spamdelay = 0 Then
                spamdelay = 2
                PrintChat "You do not meet the requirements to use this skill!", 8, Options.FontSize
            End If
        Case 12
            If spamdelay = 0 Then
                spamdelay = 1
                PrintChat "You or your target are on a No Attack Tile", 8, Options.FontSize
            End If
        Case 13
            If spamdelay = 0 Then
                spamdelay = 1
                PrintChat "You cannot cast stealth without full hp!", 8, Options.FontSize
            End If
    End Select
End If
CanUseSkill = SkillError
End Function

Public Function UseSkill(ByVal Skill As Byte) As Byte
CastingSpell = False

If combatCounter < 10 Then combatCounter = 10
If CurrentTab = tsStats2 And combatCounter = 10 Then
    calculateHpRegen
    calculateManaRegen
    DrawMoreStats
End If
drawTopBar
    If Character.SkillLevels(Skill) > 0 Then
        If CanUseSkill(Skill, 1, 0) = 0 Then
            resetMagic
            With CurrentTarget
                If Skills(Skill).TargetType = TT_NO_TARGET Then
                    SendSocket Chr$(72) + Chr$(Skill)
                    Character.GlobalSpellTick = GetTickCount + Skills(Skill).GlobalTick
                    Character.LocalSpellTick(Skill) = GetTickCount + Skills(Skill).LocalTick
                    RedrawSkills = True
                ElseIf (.TargetType And Skills(Skill).TargetType) And .TargetType <> TT_NO_TARGET Then
                    If (.TargetType And (TT_MONSTER Or TT_PLAYER Or TT_TILE)) Then
                        If Sqr((cX - .x) ^ 2 + (cY - .y) ^ 2) > Skills(Skill).Range Then
                            If spamdelay = 0 Then
                                spamdelay = 1
                                PrintChat "Out of range.", 8, Options.FontSize
                            End If
                            Exit Function
                        End If
                    End If
                    If CBool(.TargetType And (TT_MONSTER Or TT_PLAYER Or TT_TILE)) And CBool(Skills(Skill).Flags And SF_USELOS) Then
                        Select Case .TargetType
                            Case TT_MONSTER
                                If .Target < 10 Then
                                    If Not LOS(cX, cY, CLng(.x), CLng(.y), 1) Then
                                        If spamdelay = 0 Then
                                            spamdelay = 1
                                            PrintChat "No line of sight.", 8, Options.FontSize
                                        End If
                                        Exit Function
                                    End If
                                End If
                            Case TT_PLAYER
                                If .Target > 0 Then
                                    If Not LOS(cX, cY, CLng(.x), CLng(.y), 1) Then
                                        If spamdelay = 0 Then
                                            spamdelay = 1
                                            PrintChat "No line of sight.", 8, Options.FontSize
                                        End If
                                        Exit Function
                                    End If
                                End If
                            Case TT_TILE
                                If .x < 12 And .y < 12 Then
                                    If Not LOS(cX, cY, CLng(.x), CLng(.y), 1) Then '(Skill = SKILL_TELEPORT)) Then
                                        If spamdelay = 0 Then
                                            spamdelay = 1
                                            PrintChat "No line of sight.", 8, Options.FontSize
                                        End If
                                        Exit Function
                                    End If
                                End If
                        End Select
                    End If
                    SendSocket Chr$(72) + Chr$(Skill) + Chr$(.TargetType) + Chr$(.Target) + Chr$(.x) + Chr$(.y)
                    Character.GlobalSpellTick = GetTickCount + Skills(Skill).GlobalTick
                    Character.LocalSpellTick(Skill) = GetTickCount + Skills(Skill).LocalTick
                    
         
                    RedrawSkills = True
                Else
                    If CBool(Skills(Skill).Flags And SF_FRIENDLY) Then
                        .TargetType = TT_CHARACTER
                        .Target = 0
                        UseSkill Skill
                    Else
                        If spamdelay = 0 Then
                            spamdelay = 1
                            PrintChat "Invalid target.", 8, Options.FontSize, 15
                        End If
                    End If
                End If
                
                If Skill = SKILL_TELEPORT Or Skill = SKILL_INVULNERABILITY Or Skill = SKILL_BEARTRAP Then
                    Character.LocalSpellTick(Skill) = GetTickCount + Skills(Skill).LocalTick - (Skills(Skill).LocalTick / ((Skills(Skill).MaxLevel + Skills(Skill).MaxLevel / 3)) * Character.SkillLevels(Skill))
                End If
                    
                
            End With
        End If
    Else
        If spamdelay = 0 Then
            spamdelay = 2
            PrintChat "You do not have " + Skills(Skill).Name + " trained!", 8, Options.FontSize, 15
        End If
    End If
End Function

Function LOS(SourceX As Long, SourceY As Long, TargetX As Long, TargetY As Long, ExtendedLOS As Byte) As Boolean
    Dim A As Long, b As Long
    Dim dX As Long
    dX = TargetX - SourceX
    Dim dY As Long
    dY = TargetY - SourceY
    Dim AbsDx As Long
    AbsDx = Abs(dX)
    Dim AbsDy As Long
    AbsDy = Abs(dY)
    
    Dim NumSteps As Long
    If (AbsDx > AbsDy) Then
        NumSteps = AbsDx
    Else
        NumSteps = AbsDy
    End If
    
    If (map.Tile(TargetX, TargetY).Att = 27) Or (map.Tile(TargetX, TargetY).Att = 22) Or (map.Tile(TargetX, TargetY).Att = 23) Or (map.Tile(TargetX, TargetY).Att = 18) Then
        LOS = False
        Exit Function
    End If
    If (NumSteps = 0) Then
        LOS = True
        Exit Function
    Else
        Dim CurrentX As Long
        Dim LastX As Long
        CurrentX = SourceX * 64
        LastX = CurrentX
        Dim CurrentY As Long
        Dim LastY As Long
        CurrentY = SourceY * 64
        LastY = CurrentY
        Dim Xincr As Long
        Dim Yincr As Long
        Xincr = (dX * 64) \ NumSteps
        Yincr = (dY * 64) \ NumSteps
        If AbsDx = 3 Or AbsDy = 3 Then
            Xincr = Xincr + 1
            Yincr = Yincr + 1
        End If
        
        
        For b = NumSteps To 0 Step -1
            If CurrentX < 0 Then CurrentX = 0
            If CurrentY < 0 Then CurrentY = 0
            A = map.Tile(Int(CurrentX \ 64), Int(CurrentY \ 64)).Att
            If (ExtendedLOS And (A = 15 Or A = 14 Or A = 11 Or A = 13)) Or A = 1 Or A = 2 Or A = 3 Or A = 21 Then
                LOS = False
                Exit Function
            End If
            If NotLegalPath(LastX \ 64, LastY \ 64, CurrentX \ 64, CurrentY \ 64) Then 'Map.Tile(Int(CurrentX / 64), Int(CurrentY / 64)).WallTile Then
                LOS = False
                Exit Function
            End If
            LastX = CurrentX
            LastY = CurrentY
            CurrentX = CurrentX + Xincr
            CurrentY = CurrentY + Yincr
        Next
    End If
LOS = True
End Function

Function GetStatusEffect(ByVal Index As Long, StatusEffect As Long) As Byte
    If Index = Character.Index Then
            If (Character.StatusEffect And (2 ^ StatusEffect)) Then
                GetStatusEffect = 1
            End If
    Else
            If (player(Index).StatusEffect And (2 ^ StatusEffect)) Then
                GetStatusEffect = 1
            End If
    End If
End Function

Function NotLegalPath(ByVal fX As Long, ByVal fY As Long, ByVal tx As Long, ByVal ty As Long) As Boolean
Dim FromDir As Long
If fX < 0 Or fX > 11 Or fY < 0 Or fY > 11 Or tx < 0 Or tx > 11 Or ty < 0 Or ty > 11 Then
    NotLegalPath = True
    Exit Function
End If

'0up,1down,2left,3right,4upleft,5upright,6downleft,7downright
If Abs(fX - tx) + Abs(fY - ty) > 0 Then
    If fX > tx Then 'Moving left
        If fY > ty Then 'Moving up
            FromDir = 4
        ElseIf fY < ty Then 'Moving down
            FromDir = 6
        Else
            FromDir = 2
        End If
    ElseIf fX < tx Then 'Moving right
        If fY > ty Then 'Moving up
            FromDir = 5
        ElseIf fY < ty Then 'Moving down
            FromDir = 7
        Else
            FromDir = 3
        End If
    Else
        If fY > ty Then 'moving up
            FromDir = 0
        ElseIf fY < ty Then 'moving down
            FromDir = 1
        End If
    End If
    Select Case FromDir
        Case 1 'Down
            If ExamineBit(map.Flags(1), 5) = False Then
                If ExamineBit(map.Tile(tx, ty).WallTile, 0) Then GoTo notlegal
            End If
            If ExamineBit(map.Tile(fX, fY).WallTile, 5) Then GoTo notlegal
        Case 0 'Up
            If ExamineBit(map.Flags(1), 5) = False Then
                If ExamineBit(map.Tile(tx, ty).WallTile, 1) Then GoTo notlegal
            End If
            If ExamineBit(map.Tile(fX, fY).WallTile, 4) Then GoTo notlegal
        Case 3 'Right
            If ExamineBit(map.Flags(1), 5) = False Then
                If ExamineBit(map.Tile(tx, ty).WallTile, 2) Then GoTo notlegal
            End If
            If ExamineBit(map.Tile(fX, fY).WallTile, 7) Then GoTo notlegal
        Case 2 'Left
            If ExamineBit(map.Flags(1), 5) = False Then
                If ExamineBit(map.Tile(tx, ty).WallTile, 3) Then GoTo notlegal
            End If
            If ExamineBit(map.Tile(fX, fY).WallTile, 6) Then GoTo notlegal
        Case 4 'Up Left
            'try going up and then left
            If ExamineBit(map.Tile(fX, fY - 1).WallTile, 1) = 0 Or ExamineBit(map.Flags(1), 5) Then
                If ExamineBit(map.Tile(fX, fY).WallTile, 4) = 0 Then
                    'Up is good, try left
                    If ExamineBit(map.Tile(tx, ty).WallTile, 3) = 0 Or ExamineBit(map.Flags(1), 5) Then
                        If ExamineBit(map.Tile(fX, fY - 1).WallTile, 6) = 0 Then 'Up/Left is good
                            NotLegalPath = False
                            Exit Function
                        End If
                    End If
                End If
            End If
            'try going left and then up
            If ExamineBit(map.Tile(fX - 1, fY).WallTile, 3) = 0 Or ExamineBit(map.Flags(1), 5) Then
                If ExamineBit(map.Tile(fX, fY).WallTile, 6) = 0 Then
                    'left is good, try up
                    If ExamineBit(map.Tile(tx, ty).WallTile, 1) = 0 Or ExamineBit(map.Flags(1), 5) Then
                        If ExamineBit(map.Tile(fX - 1, fY).WallTile, 4) = 0 Then
                            NotLegalPath = False
                            Exit Function
                        End If
                    End If
                End If
            End If
            NotLegalPath = True
            Exit Function
        Case 5 'Up Right
            'try going up and then right
            If ExamineBit(map.Tile(fX, fY - 1).WallTile, 1) = 0 Or ExamineBit(map.Flags(1), 5) Then
                If ExamineBit(map.Tile(fX, fY).WallTile, 4) = 0 Then
                    'Up is good, try right
                    If ExamineBit(map.Tile(tx, ty).WallTile, 2) = 0 Or ExamineBit(map.Flags(1), 5) Then
                        If ExamineBit(map.Tile(fX, fY - 1).WallTile, 7) = 0 Then 'Up/Left is good
                            NotLegalPath = False
                            Exit Function
                        End If
                    End If
                End If
            End If
            'try going right and then up
            If ExamineBit(map.Tile(fX + 1, fY).WallTile, 2) = 0 Or ExamineBit(map.Flags(1), 5) Then
                If ExamineBit(map.Tile(fX, fY).WallTile, 7) = 0 Then
                    'right is good, try up
                    If ExamineBit(map.Tile(tx, ty).WallTile, 1) = 0 Or ExamineBit(map.Flags(1), 5) Then
                        If ExamineBit(map.Tile(fX + 1, fY).WallTile, 4) = 0 Then
                            NotLegalPath = False
                            Exit Function
                        End If
                    End If
                End If
            End If
            NotLegalPath = True
            Exit Function
        Case 6 'Down Left
            'try going down and then left
            If ExamineBit(map.Tile(fX, fY + 1).WallTile, 0) = 0 Or ExamineBit(map.Flags(1), 5) Then
                If ExamineBit(map.Tile(fX, fY).WallTile, 5) = 0 Then
                    'down is good, try left
                    If ExamineBit(map.Tile(tx, ty).WallTile, 3) = 0 Or ExamineBit(map.Flags(1), 5) Then
                        If ExamineBit(map.Tile(fX, fY + 1).WallTile, 6) = 0 Then 'down/Left is good
                            NotLegalPath = False
                            Exit Function
                        End If
                    End If
                End If
            End If
            'try going left and then down
            If ExamineBit(map.Tile(fX - 1, fY).WallTile, 3) = 0 Or ExamineBit(map.Flags(1), 5) Then
                If ExamineBit(map.Tile(fX, fY).WallTile, 6) = 0 Then
                    'left is good, try down
                    If ExamineBit(map.Tile(tx, ty).WallTile, 0) = 0 Or ExamineBit(map.Flags(1), 5) Then
                        If ExamineBit(map.Tile(fX - 1, fY).WallTile, 5) = 0 Then
                            NotLegalPath = False
                            Exit Function
                        End If
                    End If
                End If
            End If
            NotLegalPath = True
            Exit Function
        Case 7 'Down/Right
            'try going down and then right
            If ExamineBit(map.Tile(fX, fY + 1).WallTile, 0) = 0 Or ExamineBit(map.Flags(1), 5) Then
                If ExamineBit(map.Tile(fX, fY).WallTile, 5) = 0 Then
                    'down is good, try right
                    If ExamineBit(map.Tile(tx, ty).WallTile, 2) = 0 Or ExamineBit(map.Flags(1), 5) Then
                        If ExamineBit(map.Tile(fX, fY + 1).WallTile, 7) = 0 Then 'Up/Left is good
                            NotLegalPath = False
                            Exit Function
                        End If
                    End If
                End If
            End If
            'try going right and then down
            If ExamineBit(map.Tile(fX + 1, fY).WallTile, 2) = 0 Or ExamineBit(map.Flags(1), 5) Then
                If ExamineBit(map.Tile(fX, fY).WallTile, 7) = 0 Then
                    'right is good, try down
                    If ExamineBit(map.Tile(tx, ty).WallTile, 0) = 0 Or ExamineBit(map.Flags(1), 5) Then
                        If ExamineBit(map.Tile(fX + 1, fY).WallTile, 5) = 0 Then
                            NotLegalPath = False
                            Exit Function
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

Public Function RunSkill(Index As Long, Skill As Byte, St As String)

On Error GoTo Error_Handler

Dim A As Long, b As Long, C As Long, D As Long
Select Case Skill
    Case SKILL_EXPLOSIVETRAP
        A = Asc(Mid$(St, 1, 1))
        b = Asc(Mid$(St, 2, 1))
        ParticleEngineF.Add A * 32 + 12, b * 32 + 12, 3, 0, 0, 0, 20, 7.5, 5, 200, 8, 6, 0, 0, 0
    Case SKILL_HEAL
        A = Asc(Mid$(St, 1, 1))
        b = Asc(Mid$(St, 2, 1))
        C = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
        FloatingText.Add A * 32, b * 32 - 16, CStr(C), &HFF00FF00
        If Index = Character.Index Then
            'CreateCharacterEffect 5, 50, 8, 0, 0
            ParticleEngineF.Add 0, 0, 4, 0, 255, 0, 32, 0, 1.2, 20, 8, 3, 0, TT_CHARACTER, 0
        Else
            'CreatePlayerEffect Index, 5, 50, 8, 0, 0
            ParticleEngineF.Add 0, 0, 4, 0, 255, 0, 32, 0, 1.2, 20, 8, 3, 0, TT_PLAYER, Index
        End If
    Case SKILL_CUREPOISON
        A = Asc(Mid$(St, 1, 1))
        b = Asc(Mid$(St, 2, 1))
        If Index = Character.Index Then
            CreateCharacterEffect 65, 90, 5, 0
        Else
            CreatePlayerEffect Index, 65, 90, 5, 0
        End If
    Case SKILL_SECONDWIND
        If Index = Character.Index Then
            CreateCharacterEffect 66, 90, 8, 1
        Else
            CreatePlayerEffect Index, 66, 90, 8, 1
        End If
    Case SKILL_MYSTICISM
        If Index = Character.Index Then
            ParticleEngineF.Add Cxo, CYO, 4, 0, 0, 255, 32, 0, 1.2, 20, 8, 5, 0, TT_CHARACTER, 0
        Else
            ParticleEngineF.Add player(Index).XO, player(Index).YO, 4, 0, 0, 255, 32, 0, 1.2, 20, 8, 5, 0, TT_PLAYER, Index
        End If
    Case SKILL_NECROMANCY
        If Index = Character.Index Then
            CreateCharacterEffect 53, 60, 8, 0
        Else
            CreatePlayerEffect Index, 53, 60, 8, 0
        End If
    Case SKILL_EVOCATION
        If Index = Character.Index Then
            CreateCharacterEffect 57, 80, 8, 1
        Else
            CreatePlayerEffect Index, 57, 80, 8, 1
        End If
    Case SKILL_PARTYHEAL
        C = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
        If (Index = Character.Index) Then
            ParticleEngineF.Add 0, 0, 4, 120, 255, 255, 40, 0, -1.2, 20, 8, 3, 0, TT_CHARACTER, 0
        Else
            ParticleEngineF.Add 0, 0, 4, 120, 255, 255, 40, 0, -1.2, 20, 8, 3, 0, TT_PLAYER, Index
        End If
        If Len(St) > 2 Then
            For D = 3 To Len(St)
                A = Asc(Mid$(St, D, 1))
            
                If A = Character.Index Then
                    FloatingText.Add Cxo, CYO - 16, CStr(C), &HFF00FF00
                    'CreateTileEffect A, B, 5, 50, 8, 0, 0
                    ParticleEngineF.Add Cxo * 32, -20, 4, 120, 255, 255, 32 + 16, 0, 1.2, 20, 8, 3, 0, TT_CHARACTER, 0
                Else
                    FloatingText.Add player(A).XO, player(A).YO - 16, CStr(C), &HFF00FF00
                    'CreateTileEffect A, B, 5, 50, 8, 0, 0
                    ParticleEngineF.Add player(A).XO * 32, -20, 4, 120, 255, 255, 32 + 16, 0, 1.2, 20, 8, 3, 0, TT_PLAYER, A
                End If
            Next D
        End If
    Case SKILL_GREATERHEAL
        A = GetInt(St)
        If Len(St) > 2 Then
            For b = 3 To Len(St)
                C = Asc(Mid$(St, b, 1))
                If C = Character.Index Then
                    FloatingText.Add Cxo, CYO - 16, CStr(A), &HFF00FF00
                    'CreateCharacterEffect 5, 50, 8, 0, 0
                Else
                    If player(C).map = CMap Then
                        FloatingText.Add player(C).x * 32, player(C).y * 32 - 16, CStr(A), &HFF00FF00
                        'CreatePlayerEffect C, 5, 50, 8, 0, 0
                    End If
                End If
            Next b
        End If
        If Index = Character.Index Then
            ParticleEngineF.Add Cxo + 16, CYO, 1, 0, &HFF, 0, 65, 0, 1#, 600, 8, 3, 100, TT_NO_TARGET, 0
        Else
            With player(Index)
                ParticleEngineF.Add .XO + 16, .YO, 1, 0, &HFF, 0, 65, 0, 1#, 600, 8, 3, 100, TT_NO_TARGET, 0
            End With
        End If
    Case SKILL_HEAVENSTORM
        A = Asc(Mid$(St, 1, 1))
        b = Asc(Mid$(St, 2, 1))
        CreateAreaEffect A, b, 2, 77, 50
        If Len(St) >= 6 Then
            For C = 0 To ((Len(St) - 2) \ 4) - 1
                FloatingText.Add Asc(Mid$(St, 3 + (C * 4), 1)), Asc(Mid$(St, 4 + (C * 4), 1)), Asc(Mid$(St, 5 + (C * 4), 1)) * 256 + Asc(Mid$(St, 6 + (C * 4), 1)), D3DColorARGB(255, 255, 0, 0)
            Next C
        End If
    Case SKILL_SIPHONLIFE
        A = Asc(Mid$(St, 1, 1))
        b = Asc(Mid$(St, 2, 1))
        CreateAreaEffect A, b, 2, 40, 50
        If Len(St) >= 6 Then
            For C = 0 To ((Len(St) - 2) \ 4) - 1
                FloatingText.Add Asc(Mid$(St, 3 + (C * 4), 1)), Asc(Mid$(St, 4 + (C * 4), 1)), Asc(Mid$(St, 5 + (C * 4), 1)) * 256 + Asc(Mid$(St, 6 + (C * 4), 1)), D3DColorARGB(255, 255, 0, 0)
            Next C
        End If
    
    Case SKILL_SNOWSTORM
        A = Asc(Mid$(St, 1, 1))
        b = Asc(Mid$(St, 2, 1))
        'CreateAreaEffect A, B, 1, 78, 25
        'ParticleEngineF.Add A * 32 + 16, B * 32, 1, &HFF, &HFF, &HFF, 40, 0, 1, 250, 1, 2, 48, TT_NO_TARGET, 0
        ParticleEngineF.Add A * 32 + 16, b * 32, 5, 255, 255, 255, 40, 0.7, 0, 200, 2, 2, 36, TT_NO_TARGET, 0
        If Len(St) >= 6 Then
            For C = 0 To ((Len(St) - 2) \ 4) - 1
                'FloatingText.Add Asc(Mid$(St, 3 + (C * 4), 1)), Asc(Mid$(St, 4 + (C * 4), 1)), Asc(Mid$(St, 5 + (C * 4), 1)) * 256 + Asc(Mid$(St, 6 + (C * 4), 1)), D3DColorARGB(255, 255, 0, 0)
            Next C
        End If
    Case SKILL_TEMPEST
        A = Asc(Mid$(St, 1, 1))
        b = Asc(Mid$(St, 2, 1))
        CreateAreaEffect A, b, 1, 50, 50
        If Len(St) >= 6 Then
            For C = 0 To ((Len(St) - 2) \ 4) - 1
                FloatingText.Add Asc(Mid$(St, 3 + (C * 4), 1)), Asc(Mid$(St, 4 + (C * 4), 1)), Asc(Mid$(St, 5 + (C * 4), 1)) * 256 + Asc(Mid$(St, 6 + (C * 4), 1)), D3DColorARGB(255, 255, 0, 0)
            Next C
        End If
    Case SKILL_FURY
        A = Asc(Mid$(St, 1, 1))
        b = Asc(Mid$(St, 2, 1))
        CreateAreaEffect A, b, 3, 63, 50
        If Len(St) >= 6 Then
            For C = 0 To ((Len(St) - 2) \ 4) - 1
                FloatingText.Add Asc(Mid$(St, 3 + (C * 4), 1)), Asc(Mid$(St, 4 + (C * 4), 1)), Asc(Mid$(St, 5 + (C * 4), 1)) * 256 + Asc(Mid$(St, 6 + (C * 4), 1)), D3DColorARGB(255, 255, 0, 0)
            Next C
        End If
    Case SKILL_BLIZZARD
        A = Asc(Mid$(St, 1, 1))
        b = Asc(Mid$(St, 2, 1))
        CreateAreaEffect A, b, 2, 49, 50
        ParticleEngineF.Add A * 32 + 16, b * 32, 5, 255, 255, 255, 40, 0.7, 0, 350, 4, 2, 90, TT_NO_TARGET, 0
        If Len(St) >= 6 Then
            For C = 0 To ((Len(St) - 2) \ 4) - 1
                FloatingText.Add Asc(Mid$(St, 3 + (C * 4), 1)), Asc(Mid$(St, 4 + (C * 4), 1)), Asc(Mid$(St, 5 + (C * 4), 1)) * 256 + Asc(Mid$(St, 6 + (C * 4), 1)), D3DColorARGB(255, 255, 0, 0)
            Next C
        End If
    Case SKILL_BLIGHT
        A = Asc(Mid$(St, 3, 1)) 'targetX
        b = Asc(Mid$(St, 4, 1)) 'targetY
        'C = Asc(Mid$(St, 1, 1)) 'pX
        'D = Asc(Mid$(St, 2, 1)) 'pY
        'ParticleEngineF.Add C * 32 + 16, D * 32, 9, 25, 25, 25, 10, CSng(A * 32 + 16), CSng(B * 32 + 16), 350, 4, 2, 90, TT_NO_TARGET, 0
        'CreateAreaEffect A, B, 2, 49, 50
        ParticleEngineF.Add A * 32 + 16, b * 32, 10, 15, 40, 15, 65, 45, 1#, 400, 8, 5, 55, TT_NO_TARGET, 0
    Case SKILL_FIREBALL
        A = Asc(Mid$(St, 1, 1))
        b = Asc(Mid$(St, 2, 1))
        CreateTileEffect A, b, 1, 50, 8, 0
    Case SKILL_FLAMEWAVE
        A = Asc(Mid$(St, 1, 1))
        b = Asc(Mid$(St, 2, 1))
        C = Asc(Mid$(St, 3, 1)) 'Direction
        For D = 1 To 4
            Select Case C
                Case 0
                    'ParticleEngineF.Add A * 32 + 16, B * 32 + 16, 3, 200, 0, 0, 150, 0.7, 25, 16, 5, 0, 0
                    If b - D >= 0 Then
                        CreateTileEffect A, b - D, 9, 50, 8, 0
                    End If
                Case 1
                    If b + D <= 11 Then
                        CreateTileEffect A, b + D, 9, 50, 8, 0
                    End If
                Case 2
                    If A - D >= 0 Then
                        CreateTileEffect A - D, b, 9, 50, 8, 0
                    End If
                Case 3
                    If A + D <= 11 Then
                        CreateTileEffect A + D, b, 9, 50, 8, 0
                    End If
            End Select
        Next D
        If Len(St) >= 7 Then
            For C = 0 To ((Len(St) - 3) \ 4) - 1
                FloatingText.Add Asc(Mid$(St, 4 + (C * 4), 1)), Asc(Mid$(St, 5 + (C * 4), 1)), Asc(Mid$(St, 6 + (C * 4), 1)) * 256 + Asc(Mid$(St, 7 + (C * 4), 1)), D3DColorARGB(255, 255, 0, 0)
            Next C
        End If
    Case SKILL_ASTRALGLARE
        A = Asc(Mid$(St, 1, 1))
        b = Asc(Mid$(St, 2, 1))
        'CreateTileEffect A, B, 6, 50, 8, 0, 0
        ParticleEngineF.Add A * 32 + 16, b * 32 + 16, 3, 255, 0, 0, 15, 7.5, 5, 250, 8, 5, 0, 0, 0
    Case SKILL_VOIDBOLT
        A = Asc(Mid$(St, 1, 1))
        b = Asc(Mid$(St, 2, 1))
        CreateTileEffect A, b, 19, 50, 8, 0
    Case SKILL_HEAVENLYSTRIKE
        A = Asc(Mid$(St, 1, 1))
        'B = Asc(Mid$(St, 2, 1))
        'CreateTileEffect A, B, 76, 50, 8, 0
        'A = Asc(Mid$(St, 1, 1))
        If Asc(Mid$(St, 2, 1)) = TT_PLAYER Then
            If A = Character.Index Then
                CreateCharacterEffect 65, 50, 8, 0
            Else
                CreatePlayerEffect A, 65, 50, 8, 0
            End If
        Else
            CreateMonsterEffect A, 65, 50, 8, 0, 0, 0, True
        End If
    Case SKILL_BACKSTAB
        A = Asc(Mid$(St, 1, 1))
        If Asc(Mid$(St, 2, 1)) = TT_PLAYER Then
            If A = Character.Index Then
                CreateCharacterEffect 52, 50, 8, 0
            Else
                CreatePlayerEffect A, 52, 50, 8, 0
            End If
        Else
            CreateMonsterEffect A, 52, 50, 8, 0, 0, 0, True
        End If
        
        'B = Asc(Mid$(St, 2, 1))
        'FloatingText.Add player(A).x * 32, player(A).y * 32 - 16, "Backstab!", &HFFFFFFFF
    Case SKILL_TELEPORT
        If Len(St) = 2 Then
            A = Asc(Mid$(St, 1, 1))
            b = Asc(Mid$(St, 2, 1))
            CreateTileEffect A \ 16, A And 15, 26, 90, 8, 0
            CreateTileEffect b \ 16, b And 15, 26, 90, 8, 0
        End If
    Case SKILL_BEARTRAP
        If Len(St) = 2 Then
            A = Asc(Mid$(St, 1, 1))
            b = Asc(Mid$(St, 2, 1))
            CreateTileEffect A, b, 53, 56, 8, 0
        End If
    Case SKILL_BASH
        A = Asc(Mid$(St, 1, 1))
        'B = Asc(Mid$(St, 2, 1))
        'CreateTileEffect A, B, 31, 50, 8, 0
        If Asc(Mid$(St, 2, 1)) = TT_PLAYER Then
            If A = Character.Index Then
                CreateCharacterEffect 61, 50, 8, 0
            Else
                CreatePlayerEffect A, 61, 50, 8, 0
            End If
        Else
            CreateMonsterEffect A, 61, 50, 8, 0, 0, 0, True
        End If
    Case SKILL_JUDGMENT
        A = Asc(Mid$(St, 1, 1))
        'B = Asc(Mid$(St, 2, 1))
        'CreateTileEffect A, B, 31, 50, 8, 0
        If Asc(Mid$(St, 2, 1)) = TT_PLAYER Then
            If A = Character.Index Then
                CreateCharacterEffect 24, 80, 4, 0
            Else
                CreatePlayerEffect A, 24, 80, 4, 0
            End If
        Else
            CreateMonsterEffect A, 24, 80, 4, 0, 0, 0, True
        End If
    Case SKILL_CRIPPLE
        A = Asc(Mid$(St, 1, 1))
        'B = Asc(Mid$(St, 2, 1))
        'CreateTileEffect A, B, 31, 50, 8, 0
        If Asc(Mid$(St, 2, 1)) = TT_PLAYER Then
            If A = Character.Index Then
                CreateCharacterEffect 63, 40, 8, 1
            Else
                CreatePlayerEffect A, 63, 40, 8, 1
            End If
        Else
            CreateMonsterEffect A, 63, 40, 8, 1, 0, 0, True
        End If
    Case SKILL_MAIM
        A = Asc(Mid$(St, 1, 1))
        'B = Asc(Mid$(St, 2, 1))
        'CreateTileEffect A, B, 31, 50, 8, 0
        If Asc(Mid$(St, 2, 1)) = TT_PLAYER Then
            CreatePlayerEffect A, 33, 50, 8, 0
        Else
            CreateMonsterEffect A, 33, 50, 8, 0, 0, 0, True
        End If
    Case SKILL_BLOODLETTER
        If Len(St) = 4 Then
            A = Asc(Mid$(St, 1, 1))
            b = GetInt(Mid$(St, 2, 2))
            If A > 0 And A <= MAXUSERS Then
                If Asc(Mid$(St, 4, 1)) = TT_PLAYER Then
                    If A = Character.Index Then
                        CreateCharacterEffect 32, 50, 8, 0
                    Else
                        CreatePlayerEffect A, 32, 50, 8, 0
                    End If
                Else
                    CreateMonsterEffect A, 32, 50, 8, 0, 0, 0, True
                End If
            End If
            If Index <> Character.Index Then
                FloatingText.Add player(Index).x * 32, player(Index).y * 32 - 16, CStr(b), &HFFFF0000
            Else
                FloatingText.Add cX * 32, cY * 32 - 16, CStr(b), &HFFFF0000
            End If
        End If
    Case SKILL_SPIRITSTOUCH
        A = Asc(Mid$(St, 1, 1))
        b = Asc(Mid$(St, 2, 1))
        C = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
        FloatingText.Add A * 32, b * 32 - 16, CStr(C), &HFF00FF00
        If Index = Character.Index Then
            'CreateCharacterEffect 5, 50, 8, 0, 0
            ParticleEngineF.Add Cxo, CYO, 4, 255, 255, 255, 32, 0, 1.2, 20, 8, 3, 0, TT_CHARACTER, 0
        Else
            'CreatePlayerEffect Index, 5, 50, 8, 0, 0
            ParticleEngineF.Add player(Index).XO, player(Index).YO, 4, 255, 255, 255, 32, 0, 1.2, 20, 8, 3, 0, TT_PLAYER, Index
        End If
End Select
 
 
Exit Function
Error_Handler:
PrintChat Skills(Skill).Name & " has caused an error.  Please notify an administrator.", 15, Options.FontSize
'Open App.Path + "/LOG.TXT" For Append As #1
'    St1 = ""
'    If Len(St) > 0 Then
'        B = Len(St)
'        For A = 1 To B
'            St1 = St1 & Asc(Mid$(St, A, 1)) & "-"
'        Next A
'    End If
'    Print #1, Err.Number & "/" & Err.Description & "/" & Skill & "/" & Len(St) & "/" & St1
'Close #1
'Unhook
'EndWinsock
'End
End Function

Sub CreateAreaEffect(x As Long, y As Long, AreaType As Long, EffectNum As Long, Speed As Long)
    Dim A As Long, NumPoints As Long
    Dim PointList() As Point
    Select Case AreaType
        Case 1
            NumPoints = 4
            ReDim PointList(0 To NumPoints)         'OXO
            PointList(0).x = x                      'XXX
            PointList(0).y = y - 1                  'OXO
            PointList(1).x = x - 1
            PointList(1).y = y
            PointList(2).x = x
            PointList(2).y = y
            PointList(3).x = x + 1
            PointList(3).y = y
            PointList(4).x = x
            PointList(4).y = y + 1
        Case 3
            NumPoints = 3
            ReDim PointList(0 To NumPoints)         'OXO
            PointList(0).x = x                      'XXX
            PointList(0).y = y - 1                  'OXO
            PointList(1).x = x - 1
            PointList(1).y = y
            PointList(3).x = x + 1
            PointList(3).y = y
            PointList(2).x = x
            PointList(2).y = y + 1
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
        Case 4
            NumPoints = 8
            ReDim PointList(0 To NumPoints)         'XXX
            PointList(0).x = x                      'XXX
            PointList(0).y = y - 1                  'XXX
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
            
    End Select
    
    For A = 0 To NumPoints
        With PointList(A)
            If .x >= 0 And .x <= 12 And .y >= 0 And .y <= 11 Then
                CreateTileEffect .x, .y, EffectNum, Speed, 8, 0
            End If
        End With
    Next A
End Sub

Public Sub InitSkillEXPTable()
Dim A As Long
For A = 1 To 50
    SkillEXPTable(ET_STANDARD, A) = Int(A ^ 1.6) * 100
Next A
End Sub
