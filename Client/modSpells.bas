Attribute VB_Name = "modSpells"
Option Explicit

Public Enum ClassEnum
    Mage = 1
    Knight = 2
    Paladin = 4
    Barbarian = 8
    Cleric = 16
    Necromancer = 32
    Thief = 64
End Enum

Public Enum StatusEffects 'Supports from 1 to 30
    Poisoned = 1
    Regen = 2
    Disease = 3
    Mute = 4
    Vex = 5
End Enum

Public Const _
    splCureLightWounds = 1, _
    splBolt = 2, _
    splPoison = 3, _
    splHolyFlame = 4, _
    splSlash = 5, _
    splPacify = 6, _
    splHoly = 7, _
    splBackStab = 8, _
    splTorch = 9, _
    splFreeze = 10, _
    splRepel = 11, _
    splKnockBack = 12, _
    splDoubleSlash = 13, _
    splZap = 14, _
    splRepair = 15, _
    splSealWounds = 16, _
    splGoldSpy = 17, _
    splStealGold = 18, _
 _
    splEnergyWeb = 21, _
    splIceBlast = 22, _
    splHellfire = 23
Public Const _
    splHeal = 24, _
    splRevitalize = 25, _
    splGreaterHeal = 26, _
    splWaterofLife = 27, _
    splDrain = 28, _
    splRefresh = 29, _
    splInvoke = 30, _
    splFirestorm = 31, _
    splVengeanceBlade = 32, _
    splStormBlade = 33, _
    splPickLock = 34, _
    SplBash = 35, _
    SplBoneBlast = 36, _
    SplDash = 37, _
    SplCrush = 38, _
    SplHolyNova = 39, _
    SplQuickHit = 40, _
    SplMeteor = 41, _
    SplFireCross = 42, _
    SplFireBomb = 43, _
    SplRage = 44, _
    SplDivineTouch = 45, _
    splTorturedExistance = 46, _
    SplLightningStab = 47
Public Const _
    splAssassinate = 48, _
    splShardofIce = 49, _
    splIceBlade = 50, _
    splPurity = 51, _
    splRegen = 52, _
    splCancel = 53, _
    splVex = 54, _
    splDisease = 55, _
    splMute = 56, _
    splGroupHeal = 57
    
Type Spells
    Name As String
    MinLevel As Byte
    Cost As Integer
    Class As ClassEnum
    Type As Byte  'Friendly = 1, hostile = 2
    Range As Byte
    Recovery As Integer
    Description As String
End Type

Public Const Friendly = 1
Public Const Hostile = 2

Public Const TargTypePlayer = 1
Public Const TargTypeMonster = 2
Public Const TargTypeCharacter = 3

Public Const MaxSpells As Byte = 100
Public Spells(1 To MaxSpells) As Spells

Public Sub CreateSpell(Spell As Byte, Name As String, MinLevel As Byte, Cost As Integer, Class As ClassEnum, tType As Byte, Range As Byte, Recovery As Integer, Description As String)
With Spells(Spell)
    .Name = Name
    .MinLevel = MinLevel
    .Cost = Cost
    .Class = Class
    .Type = tType
    .Range = Range
    .Recovery = Recovery
    .Description = Description
End With
End Sub

Public Sub InitSpells()
'----Init Cleric Spells-----'
CreateSpell splCureLightWounds, "Cure Light Wounds", 1, 3, Cleric Or Paladin, Friendly, 4, 25, "Cures target for small amount of HP - Range:4"
CreateSpell splHoly, "Holy", 45, 150, Cleric, Hostile, 3, 150, "A harmful light that burns target - Range:3"
CreateSpell splSealWounds, "Seal Wounds", 10, 30, Cleric Or Paladin, Friendly, 4, 35, "Cures target for medium amount of HP - Range:4"
CreateSpell splHeal, "Heal", 20, 50, Cleric Or Paladin, Friendly, 4, 40, "Cures target for medium amount of HP - Range:4"
CreateSpell splRevitalize, "Revitalize", 30, 80, Cleric, Friendly, 4, 60, "Heals target for large amount of HP - Range:4"
CreateSpell splGreaterHeal, "Greater Heal", 40, 120, Cleric, Friendly, 4, 80, "Cures target for large amount of HP - Range:4"
CreateSpell splWaterofLife, "Water of Life", 50, 150, Cleric, Friendly, 6, 100, "Cures target for very large amount of HP - Range:6"
CreateSpell splRefresh, "Refresh", 8, 25, Cleric, Friendly, 3, 50, "Restores targets energy - Range:3"
CreateSpell SplDivineTouch, "Divine Touch", 5, 8, Cleric, Friendly, 6, 35, "A DivineTouch that burns evil - Range:6"
CreateSpell splPurity, "Purity", 6, 8, Cleric, Friendly, 5, 20, "Cure target of poison - Range:5"
CreateSpell splRegen, "Regen", 25, 300, Cleric, Friendly, 99, 300, "Heals target over slow period of time - Range:99"
CreateSpell splGroupHeal, "Group Heal", 15, 100, Cleric, Friendly, 0, 100, "Cures members of guild or party"
'----End Cleric Spells------'
'----Init Knight Spells-----'
CreateSpell splSlash, "Slash", 15, 20, Knight, Hostile, 1, 200, "Visciously attacks target - Range:1"
CreateSpell splKnockBack, "Knock Back", 99, 40, Knight, Hostile, 1, 100, "Knocks target back"
CreateSpell splDoubleSlash, "Double Slash", 50, 30, Knight, Hostile, 1, 500, "A double reckless attack for random damage - Range:1"
CreateSpell splStormBlade, "Storm Blade", 35, 25, Knight, Hostile, 25, 250, "Shoot forth whirlwinds, decimating opponents - Range:1"
CreateSpell SplDash, "Dash", 1, 5, Knight, Hostile, 1, 35, "Dash Attacks Target for up to 15 Damage - Range:1"
CreateSpell splIceBlade, "Ice Blade", 40, 40, Knight, Hostile, 1, 500, "A blade so cold it burns target up to 200 dmg - Range:1"
'----End Knight Spells------'
'----Init Mage Spells-------'
CreateSpell splBolt, "Bolt", 1, 3, Mage, Hostile, 6, 35, "Small bolt of lightning strikes target - Range:6"
CreateSpell splTorch, "Torch", 5, 8, Mage, Hostile, 6, 35, "Sets target on fire - Range:6"
CreateSpell splZap, "Zap", 15, 50, Mage, Hostile, 5, 50, "Strong bolt for medium damage - Range:5"
CreateSpell splFreeze, "Freeze", 99, 3, Mage, Hostile, 10, 35, "Freezes Target - Range:10"
CreateSpell splEnergyWeb, "Energy Web", 25, 75, Mage, Hostile, 4, 50, "A burning web of energy that melts target - Range:4"
CreateSpell splIceBlast, "Ice Blast", 35, 100, Mage, Hostile, 4, 50, "Heavy blast of ice - Range:4"
CreateSpell splHellfire, "Hellfire", 40, 150, Mage, Hostile, 2, 175, "A massive fire blast - Range:3"
CreateSpell splInvoke, "Invoke", 25, 0, Mage, Friendly, 0, 50, "Restores MP at the cost of casters life"
CreateSpell splFirestorm, "Firestorm", 26, 50, Mage, Hostile, 4, 75, "Firestorm descends upon target - Range:5"
CreateSpell SplMeteor, "Meteor", 45, 200, Mage, Hostile, 2, 350, "A firey rock fired from the hands of God - Range:2"
CreateSpell splShardofIce, "Shard of Ice", 10, 30, Mage, Hostile, 5, 35, "A flying shard of ice stabs target - Range:5"
'----End Mage Spells -------'
'----Init Necro Spells-------'
CreateSpell splPacify, "Pacify", 5, 15, Necromancer, Hostile, 3, 100, "Calms monster to stop in peace - Range:3"
CreateSpell splRepel, "Repel", 99, 30, Necromancer, Friendly, 0, 150, ""
CreateSpell splDrain, "Drain", 10, 30, Necromancer, Hostile, 1, 100, "Damages foe healing the caster - Range:1"
CreateSpell SplBoneBlast, "Bone Blast", 10, 14, Necromancer, Hostile, 4, 50, "Fires sharp bones at target - Range:4"
CreateSpell splTorturedExistance, "Tortured Existance", 20, 75, Necromancer, Hostile, 4, 75, "Causes Pain upon target - Range:4"
CreateSpell splPoison, "Poison", 10, 46, Necromancer, Hostile, 1, 35, "Infect target with venom"
CreateSpell splVex, "Vex", 25, 100, Necromancer, Hostile, 1, 500, "Perplex target - Range:1"
CreateSpell splDisease, "Disease", 15, 110, Necromancer, Hostile, 1, 500, "Afflict target with Disease - Range:1"
CreateSpell splMute, "Mute", 23, 120, Necromancer, Hostile, 1, 500, "Cause a viel of silence to descend upon target - Range:1"
'----End Necro Spells------'
'----Init Paladin Spells----'
CreateSpell splHolyFlame, "Holy Flame", 1, 3, Paladin, Hostile, 4, 35, "A holy flame shoots forth - Range:4"
CreateSpell splVengeanceBlade, "Vengeance Blade", 15, 25, Paladin, Hostile, 1, 60, "Attack evil with holy magic - Range:1"
CreateSpell splVengeanceBlade, "Holy Nova", 25, 75, Paladin, Hostile, 4, 75, "Suffocates enemy while burning with holylight - Range:4"
CreateSpell SplFireCross, "FireCross", 10, 30, Paladin, Hostile, 1, 150, "Strikes fear into evil +50 Damage - Range:1"
CreateSpell SplLightningStab, "Lightning Stab", 55, 100, Paladin, Hostile, 1, 500, "A Devastating Shock emits from Weapon - Range:1"
'----End Paladin Spells-----'
'----Init Barbarian Spells--'
CreateSpell splRepair, "Repair", 20, 0, Barbarian, Friendly, 0, 500, "Repairs objects"
CreateSpell SplBash, "Bash", 15, 20, Barbarian, Hostile, 1, 350, "Deals +150 damage to target. - Range:1"
CreateSpell SplCrush, "Crush", 1, 5, Barbarian, Hostile, 1, 50, "Crushes Target for up to 25 Damage - Range:1"
CreateSpell SplRage, "Rage", 50, 40, Barbarian, Hostile, 1, 500, "A reckless attack for random damage - Range:1"
'----End Barbarian Spells--'
'----Init Theif Spells-----'
CreateSpell splGoldSpy, "Gold Spy", 1, 7, Thief, Friendly, 1, 80, "Senses gold in targets inventory"
CreateSpell splStealGold, "Steal Gold", 5, 12, Thief, Hostile, 1, 20, "Steals gold from target"
CreateSpell splPickLock, "Pick Lock", 10, 50, Thief, Friendly, 0, 1000, "Attempts to pick a lock"
CreateSpell SplQuickHit, "Quick Hit", 5, 20, Thief, Hostile, 1, 75, "Deals up to 50 damage- Range:1"
CreateSpell SplFireBomb, "FireBomb", 20, 30, Thief, Hostile, 1, 250, "A FireBomb that explodes on impact +150 Damage - Range:1"
CreateSpell splAssassinate, "Assassinate", 35, 40, Thief, Hostile, 1, 300, "A quick fatal stab that adds +200 Damage - Range:1"
'----End Theif Spells-----'
End Sub
'Display Spells
Public Sub DisplaySpells()
Dim A As Long, B As Long
For A = 1 To 255
    SpellListBox.Caption(A) = ""
    SpellListBox.Data(A) = 0
Next A
For B = 1 To 255
    For A = 1 To MaxSpells
        'If 2 ^ (Character.Class - 1) And Spells(A).Class Then
            If Character.level >= Spells(A).MinLevel And Spells(A).MinLevel = B Then
                frmMain.LstAddItem Spells(A).Name, CByte(A)
            End If
        'End If
    Next A
Next B
End Sub

Public Sub CastSpell(Spell As Long, Target As Long, TargType As Long)
Dim A As Long, B As Long
If Character.Recover < Character.Recovery Then Exit Sub
If Spell = 0 Then Exit Sub

If Spell = splRepair Then
    If Character.Mana <> Character.MaxMana Then
        PrintChat "Not enough mana to cast this spell!", 15
        Exit Sub
    End If
Else
    If Character.Mana - Spells(Spell).Cost < 0 Then
        PrintChat "Not enough mana to cast this spell!", 15
        Exit Sub
    End If
End If

'Check Range And LOS
Select Case TargType
    Case TargTypeMonster
        With Map.Monster(Target)
            A = Sqr((cx - .X) ^ 2 + (cy - .Y) ^ 2)
            If A > 1 Then If Not LOS(cx, cy, .X, .Y) Then B = 1
        End With
    Case TargTypePlayer
        With Player(Target)
            A = Sqr((cx - .X) ^ 2 + (cy - .Y) ^ 2)
            If A > 1 Then If Not LOS(cx, cy, .X, .Y) Then B = 1
        End With
End Select

If A > Spells(Spell).Range Or B = 1 Then
    PrintChat "You cannot cast that from here!", 7
    Exit Sub
End If

    If TargType = TargTypePlayer Then
        If Spells(Spell).Type = Friendly Then
            PrintChat "You cast " + Spells(Spell).Name + "!", 15
            Select Case Spell
                Case splRepair
                    PrintChat "You cannot cast this spell on other players!", 7
                    Exit Sub
                Case splGoldSpy
                    SendSocket Chr$(72) + Chr$(Spell) + Chr$(Target)
                Case Else
                    SendSocket Chr$(72) + Chr$(Spell) + Chr$(TargTypePlayer) + Chr$(Target)
            End Select
        ElseIf Spells(Spell).Type = Hostile Then
            With Player(Target)
                If .Map = CMap And .Sprite > 0 Then
                    If ExamineBit(Map.Flags, 0) = False Then
                        If .Guild > 0 Or ExamineBit(Map.Flags, 6) = True Then
                            If .Guild = 0 Or .Guild <> Character.Guild Then
                                'If Not .Party = Character.Party And .Party > 0 Then
                                    Select Case Spell
                                        Case splRepair, splPickLock
                                            PrintChat "You cannot cast this spell on other players!", 7
                                            Exit Sub
                                        Case splStealGold
                                            SendSocket Chr$(72) + Chr$(Spell) + Chr$(Target)
                                        Case Else
                                            SendSocket Chr$(72) + Chr$(Spell) + Chr$(TargTypePlayer) + Chr$(Target)
                                    End Select
                                    PrintChat "You cast " + Spells(Spell).Name + "!", 15
                                    Exit Sub
                                'Else
                                '    PrintChat "You cannot hit party members!", 7
                                'End If
                            Else
                                PrintChat "You cannot cast that on them!", 7
                                Exit Sub
                            End If
                        Else 'Not in guild
                            PrintChat "You cannot cast that here!", 7
                            Exit Sub
                        End If
                    Else 'Friendly
                        PrintChat "You cannot cast that in a friendly zone!", 7
                        Exit Sub
                    End If
                End If
            End With
        End If
    ElseIf TargType = TargTypeMonster Then
        If ExamineBit(Map.Flags, 5) = 0 Then
            Select Case Spell
                Case splStealGold, splGoldSpy, splRepair, splPickLock
                    PrintChat "You cannot cast this spell on monsters!", 7
                    Exit Sub
                Case splPacify
                    SendSocket Chr$(72) + Chr$(Spell) + Chr$(Target)
                Case Else
                    SendSocket Chr$(72) + Chr$(Spell) + Chr$(TargTypeMonster) + Chr$(Target)
            End Select
            PrintChat "You cast " + Spells(Spell).Name + "!", 15
        Else
            PrintChat "You cannot cast this on these monsters!", 7
            Exit Sub
        End If
    ElseIf TargType = TargTypeCharacter Then
        If Spells(Spell).Type = Friendly Then
            Select Case Spell
                Case splStealGold, splGoldSpy, splBolt, splTorch, splZap, splHolyFlame, splPoison, splSlash, splDoubleSlash, splFreeze, splPacify, SplBoneBlast, SplBash, SplHolyNova, SplDash, SplCrush, SplMeteor, SplQuickHit, SplFireCross, SplFireBomb, SplRage, SplLightningStab, splAssassinate, SplDivineTouch, splTorturedExistance, splShardofIce, splHoly
                    PrintChat "You cannot cast this spell on yourself!", 7
                    Exit Sub
                Case splRepair 'Repair (Barbarian)
                    SendSocket Chr$(72) + Chr$(splRepair) + Chr$(CurInvObj)
                Case splInvoke, splPickLock, splGroupHeal
                    SendSocket Chr$(72) + Chr$(Spell)
                Case Else 'I figure since this is the most common thing sent by a spell, we can save lines and just do this
                    SendSocket Chr$(72) + Chr$(Spell) + Chr$(TargTypeCharacter) + Chr$(Target)
            End Select
            PrintChat "You cast " + Spells(Spell).Name + "!", 15
        ElseIf Spells(Spell).Type = Hostile Then
            Select Case Spell
                Case splStormBlade
                    SendSocket Chr$(72) + Chr$(splStormBlade)
                Case Else
                    PrintChat "You cannot cast this spell on yourself!", 7
                    Exit Sub
            End Select
            PrintChat "You cast " + Spells(Spell).Name + "!", 15
        End If
    End If
Character.Recovery = Spells(Spell).Recovery
Character.Recover = 0
End Sub

Public Sub RunSpell(St As String)
Dim A As Long, B As Long, C As Long, D As Long, E As Long, X As Long, Y As Long
A = Asc(Mid$(St, 1, 1))
Select Case A
'---------------------------Healing Spells--------------------------
    Case splCureLightWounds, splSealWounds, splHeal, splRevitalize, splGreaterHeal 'Healing Spells (Cleric)
        B = Asc(Mid$(St, 2, 1)) 'Type
        C = Asc(Mid$(St, 3, 1)) 'Target
        D = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)) 'Heal Amount
            If B = TargTypeCharacter Then
                Call CreateTileEffect(cx, cy, 5, 75, 7, 0, 0)
                FloatingText.Add cx, cy, CStr(D), 10
            Else
                With Player(C)
                    Call CreateTileEffect(.X, .Y, 5, 75, 7, 0, 0)
                    FloatingText.Add .X, .Y, CStr(D), 10
                End With
            End If
    Case splWaterofLife
        B = Asc(Mid$(St, 2, 1)) 'Type
        C = Asc(Mid$(St, 3, 1)) 'Target
        D = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)) 'Heal Amount
            If B = TargTypeCharacter Then
                Call CreateTileEffect(cx, cy, 58, 75, 7, 0, 0)
                FloatingText.Add cx, cy, CStr(D), 10
            Else
                With Player(C)
                    Call CreateTileEffect(.X, .Y, 58, 125, 8, 0, 0)
                    FloatingText.Add .X, .Y, CStr(D), 10
                End With
            End If
    Case splRefresh 'Refresh (Cleric)
        B = Asc(Mid$(St, 2, 1))
        If B = Character.Index Then
            Call CreateTileEffect(cx, cy, 35, 75, 8, 1, 0)
        Else
            Call CreateTileEffect(Player(B).X, Player(B).Y, 35, 75, 8, 1, 0)
        End If
    Case splInvoke
        B = Asc(Mid$(St, 2, 1))
        If B = Character.Index Then
            Call CreateTileEffect(cx, cy, 25, 75, 8, 3, 0)
        Else
            Call CreateTileEffect(Player(B).X, Player(B).Y, 25, 75, 8, 3, 0)
        End If
    Case splGroupHeal
        B = Asc(Mid$(St, 2, 1))
        D = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
        CreateTileEffect Player(B).X, Player(B).Y, 5, 75, 8, 0, 0
        FloatingText.Add Player(B).X, Player(B).Y, CStr(D), 10
'------------------------Offensive Target Spells--------------------
    Case splBolt 'Bolt (Mage)
        C = Asc(Mid$(St, 2, 1)) 'Type
        D = Asc(Mid$(St, 3, 1)) 'Target
        E = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)) 'Attack Amount
            Select Case C
                Case TargTypeMonster
                    With Map.Monster(D)
                        Call CreateTileEffect(.X, .Y, 8, 75, 6, 0, 0)
                    End With
                Case TargTypePlayer
                    If D = Character.Index Then
                        With Character
                            Call CreateTileEffect(cx, cy, 8, 75, 6, 0, 0)
                        End With
                    Else
                        With Player(D)
                            Call CreateTileEffect(.X, .Y, 8, 75, 6, 0, 0)
                        End With
                    End If
            End Select
    Case splHolyFlame 'Holy Flame
        C = Asc(Mid$(St, 2, 1)) 'Type
        D = Asc(Mid$(St, 3, 1)) 'Target
        E = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)) 'Attack Amount
            Select Case C
                Case TargTypeMonster
                    With Map.Monster(D)
                        Call CreateTileEffect(.X, .Y, 1, 75, 7, 0, 0)
                    End With
                Case TargTypePlayer
                    If D = Character.Index Then
                        With Character
                            Call CreateTileEffect(cx, cy, 1, 75, 6, 0, 0)
                        End With
                    Else
                        With Player(D)
                            Call CreateTileEffect(.X, .Y, 1, 75, 6, 0, 0)
                        End With
                    End If
            End Select
    Case splVengeanceBlade
        C = Asc(Mid$(St, 2, 1)) 'Type
        D = Asc(Mid$(St, 3, 1)) 'Target
        E = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)) 'Amount
        Select Case C
            Case TargTypeMonster
                With Map.Monster(D)
                    Call CreateTileEffect(.X, .Y, 45, 75, 8, 0, 0)
                    Call CreateTileEffect(.X, .Y, 53, 75, 8, 0, 0)
                End With
            Case TargTypePlayer
                If D = Character.Index Then
                    Call CreateTileEffect(cx, cy, 45, 75, 8, 0, 0)
                    Call CreateTileEffect(cx, cy, 53, 75, 8, 0, 0)
                Else
                    With Player(D)
                        Call CreateTileEffect(.X, .Y, 45, 75, 8, 0, 0)
                        Call CreateTileEffect(.X, .Y, 53, 75, 8, 0, 0)
                    End With
                End If
        End Select
    Case splTorch 'Torch
        C = Asc(Mid$(St, 2, 1)) 'Type
        D = Asc(Mid$(St, 3, 1)) 'Target
        E = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)) 'Attack Amount
            Select Case C
                Case TargTypeMonster
                    With Map.Monster(D)
                        Call CreateTileEffect(.X, .Y, 18, 25, 7, 0, 0)
                    End With
                Case TargTypePlayer
                    With Player(D)
                        Call CreateTileEffect(.X, .Y, 18, 25, 7, 0, 0)
                    End With
            End Select
    Case splZap 'Zap (Mage)
        C = Asc(Mid$(St, 2, 1)) 'Type
        D = Asc(Mid$(St, 3, 1)) 'Target
        E = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)) 'Attack Amount
            Select Case C
                Case TargTypeMonster
                    With Map.Monster(D)
                        Call CreateTileEffect(.X, .Y, 7, 75, 7, 0, 0)
                    End With
                Case TargTypePlayer
                    With Player(D)
                        Call CreateTileEffect(.X, .Y, 7, 75, 7, 0, 0)
                    End With
            End Select
    Case splEnergyWeb 'Energy Web (Mage)
        C = Asc(Mid$(St, 2, 1)) 'Type
        D = Asc(Mid$(St, 3, 1)) 'Target
        E = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)) 'Attack Amount
            Select Case C
                Case TargTypeMonster
                    With Map.Monster(D)
                        Call CreateTileEffect(.X, .Y, 4, 75, 4, 0, 0)
                    End With
                Case TargTypePlayer
                    With Player(D)
                        Call CreateTileEffect(.X, .Y, 4, 75, 4, 0, 0)
                    End With
            End Select
    Case splIceBlast 'Ice Blast (Mage)
        C = Asc(Mid$(St, 2, 1)) 'Type
        D = Asc(Mid$(St, 3, 1)) 'Target
        E = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)) 'Attack Amount
            Select Case C
                Case TargTypeMonster
                    With Map.Monster(D)
                        Call CreateTileEffect(.X, .Y, 17, 75, 5, 0, 0)
                    End With
                Case TargTypePlayer
                    With Player(D)
                        Call CreateTileEffect(.X, .Y, 17, 75, 5, 0, 0)
                    End With
            End Select
    Case splHellfire 'Hellfire (Mage)
        C = Asc(Mid$(St, 2, 1)) 'Type
        D = Asc(Mid$(St, 3, 1)) 'Target
        E = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)) 'Attack Amount
            Select Case C
                Case TargTypeMonster
                    With Map.Monster(D)
                        Call CreateTileEffect(.X, .Y, 6, 75, 7, 0, 0)
                        Call CreateTileEffect(.X, .Y, 7, 75, 7, 0, 0)
                    End With
                Case TargTypePlayer
                    With Player(D)
                        Call CreateTileEffect(.X, .Y, 6, 75, 7, 0, 0)
                        Call CreateTileEffect(.X, .Y, 7, 75, 7, 0, 0)
                    End With
            End Select
    Case splFirestorm
        C = Asc(Mid$(St, 2, 1))
        D = Asc(Mid$(St, 3, 1))
        CreateTileEffect C, D, 9, 75, 8, 0, 0
        CreateTileEffect C - 1, D, 9, 75, 8, 0, 0
        CreateTileEffect C + 1, D, 9, 75, 8, 0, 0
        CreateTileEffect C, D - 1, 9, 75, 8, 0, 0
        CreateTileEffect C, D + 1, 9, 75, 8, 0, 0
    Case SplBoneBlast 'BoneBlast (Necromancer)
        C = Asc(Mid$(St, 2, 1)) 'Type
        D = Asc(Mid$(St, 3, 1)) 'Target
        E = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)) 'Attack Amount
            Select Case C
                Case TargTypeMonster
                    With Map.Monster(D)
                        Call CreateTileEffect(.X, .Y, 44, 75, 8, 0, 0)
                    End With
                Case TargTypePlayer
                    With Player(D)
                        Call CreateTileEffect(.X, .Y, 44, 75, 8, 0, 0)
                    End With
            End Select
    Case SplHolyNova 'Holy Nova
        C = Asc(Mid$(St, 2, 1)) 'Type
        D = Asc(Mid$(St, 3, 1)) 'Target
        E = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)) 'Attack Amount
            Select Case C
                Case TargTypeMonster
                    With Map.Monster(D)
                        Call CreateTileEffect(.X, .Y, 3, 75, 8, 0, 0)
                    End With
                Case TargTypePlayer
                    If D = Character.Index Then
                        With Character
                            Call CreateTileEffect(cx, cy, 3, 75, 8, 0, 0)
                        End With
                    Else
                        With Player(D)
                            Call CreateTileEffect(.X, .Y, 3, 75, 8, 0, 0)
                        End With
                    End If
            End Select
    Case SplMeteor 'Meteor
        C = Asc(Mid$(St, 2, 1)) 'Type
        D = Asc(Mid$(St, 3, 1)) 'Target
        E = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)) 'Attack Amount
            Select Case C
                Case TargTypeMonster
                    With Map.Monster(D)
                        Call CreateTileEffect(.X, .Y, 43, 75, 8, 0, 0)
                    End With
                Case TargTypePlayer
                    If D = Character.Index Then
                        With Character
                            Call CreateTileEffect(cx, cy, 43, 75, 8, 0, 0)
                        End With
                    Else
                        With Player(D)
                            Call CreateTileEffect(.X, .Y, 43, 75, 7, 0, 0)
                        End With
                    End If
            End Select
    Case SplDivineTouch 'DivineTouch
        C = Asc(Mid$(St, 2, 1)) 'Type
        D = Asc(Mid$(St, 3, 1)) 'Target
        E = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)) 'Attack Amount
            Select Case C
                Case TargTypeMonster
                    With Map.Monster(D)
                        Call CreateTileEffect(.X, .Y, 25, 75, 8, 0, 0)
                    End With
                Case TargTypePlayer
                    With Player(D)
                        Call CreateTileEffect(.X, .Y, 25, 75, 8, 0, 0)
                    End With
            End Select
    Case splTorturedExistance 'SplTortured Existence (Necromancer)
        C = Asc(Mid$(St, 2, 1)) 'Type
        D = Asc(Mid$(St, 3, 1)) 'Target
        E = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)) 'Attack Amount
            Select Case C
                Case TargTypeMonster
                    With Map.Monster(D)
                        Call CreateTileEffect(.X, .Y, 55, 75, 8, 0, 0)
                    End With
                Case TargTypePlayer
                    With Player(D)
                        Call CreateTileEffect(.X, .Y, 55, 75, 8, 0, 0)
                    End With
            End Select
    Case splShardofIce 'Shard of Ice (Mage)
        C = Asc(Mid$(St, 2, 1)) 'Type
        D = Asc(Mid$(St, 3, 1)) 'Target
        E = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)) 'Attack Amount
            Select Case C
                Case TargTypeMonster
                    With Map.Monster(D)
                        Call CreateTileEffect(.X, .Y, 49, 75, 8, 0, 0)
                    End With
                Case TargTypePlayer
                    With Player(D)
                        Call CreateTileEffect(.X, .Y, 49, 75, 8, 0, 0)
                    End With
            End Select
    Case splHoly 'Holy (Cleric)
        C = Asc(Mid$(St, 2, 1)) 'Type
        D = Asc(Mid$(St, 3, 1)) 'Target
        E = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)) 'Attack Amount
            Select Case C
                Case TargTypeMonster
                    With Map.Monster(D)
                        Call CreateTileEffect(.X, .Y, 26, 75, 8, 0, 0)
                    End With
                Case TargTypePlayer
                    With Player(D)
                        Call CreateTileEffect(.X, .Y, 26, 75, 8, 0, 0)
                    End With
            End Select
'---------------------------Area Spells----------------------------

'---------------------Close Range Target Spells--------------------
    Case splSlash
        B = Asc(Mid$(St, 2, 1)) 'Target Type
        C = Asc(Mid$(St, 3, 1)) 'Target
        Select Case B
            Case TargTypeMonster
                With Map.Monster(C)
                    Call CreateTileEffect(.X, .Y, 48, 75, 8, 0, 0)
                End With
            Case TargTypePlayer
                If B = Character.Index Then
                    Call CreateTileEffect(cx, cy, 48, 75, 8, 0, 0)
                Else
                    With Player(C)
                        Call CreateTileEffect(.X, .Y, 48, 75, 8, 0, 0)
                    End With
                End If
        End Select
    Case SplBash 'Bash
        B = Asc(Mid$(St, 2, 1)) 'Target Type
        C = Asc(Mid$(St, 3, 1)) 'Target
        Select Case B
            Case TargTypeMonster
                With Map.Monster(C)
                    Call CreateTileEffect(.X, .Y, 61, 75, 4, 0, 0)
                End With
            Case TargTypePlayer
                If B = Character.Index Then
                    Call CreateTileEffect(cx, cy, 61, 75, 4, 0, 0)
                Else
                    With Player(C)
                        Call CreateTileEffect(.X, .Y, 61, 75, 4, 0, 0)
                    End With
                End If
        End Select
    Case splDoubleSlash
        B = Asc(Mid$(St, 2, 1)) 'Target Type
        C = Asc(Mid$(St, 3, 1)) 'Target
        Select Case B
            Case TargTypeMonster
                With Map.Monster(C)
                    Call CreateTileEffect(.X, .Y, 33, 75, 8, 0, 0)
                End With
            Case TargTypePlayer
                If B = Character.Index Then
                    Call CreateTileEffect(cx, cy, 33, 75, 8, 0, 0)
                Else
                    With Player(C)
                        Call CreateTileEffect(.X, .Y, 33, 75, 8, 0, 0)
                    End With
                End If
        End Select
    Case splStormBlade
        A = Asc(Mid$(St, 2, 1)) 'Caster
        B = Asc(Mid$(St, 3, 1)) 'Direction
        If Character.Index = A Then
            C = cx
            D = cy
        Else
            C = Player(A).X
            D = Player(A).Y
        End If
        Select Case B
            Case 0
                Call CreateTileEffect(C, D - 1, 52, 50, 8, 0, 0)
                Call CreateTileEffect(C, D - 2, 52, 75, 8, 0, 0)
            Case 1
                Call CreateTileEffect(C, D + 1, 52, 50, 8, 0, 0)
                Call CreateTileEffect(C, D + 2, 52, 75, 8, 0, 0)
            Case 2
                Call CreateTileEffect(C - 1, D, 52, 50, 8, 0, 0)
                Call CreateTileEffect(C - 2, D, 52, 75, 8, 0, 0)
            Case 3
                Call CreateTileEffect(C + 1, D, 52, 50, 8, 0, 0)
                Call CreateTileEffect(C + 2, D, 52, 75, 8, 0, 0)
        End Select
    Case splGoldSpy 'Gold Spy
        B = Asc(Mid$(St, 2, 1)) 'Target
        C = Asc(Mid$(St, 3, 1)) * 16777216 + Asc(Mid$(St, 4, 1)) * 65536 + Asc(Mid$(St, 5, 1)) * 256& + Asc(Mid$(St, 6, 1)) 'Spy Amount
        PrintChat Player(B).Name & " has " & C & " gold.", 14
    Case splStealGold 'Steal Gold
        If Asc(Mid$(St, 2, 1)) = 255 Then
            PrintChat "Your purse feels slightly lighter", 14
            Exit Sub
        End If
        B = Asc(Mid$(St, 2, 1)) 'Target
        C = Asc(Mid$(St, 3, 1)) * 16777216 + Asc(Mid$(St, 4, 1)) * 65536 + Asc(Mid$(St, 5, 1)) * 256& + Asc(Mid$(St, 6, 1)) 'Steal Amount
        PrintChat "You stole " & C & " gold from " & Player(B).Name, 14
    Case SplQuickHit
        B = Asc(Mid$(St, 2, 1)) 'Target Type
        C = Asc(Mid$(St, 3, 1)) 'Target
        Select Case B
            Case TargTypeMonster
                With Map.Monster(C)
                    Call CreateTileEffect(.X, .Y, 12, 75, 6, 0, 0)
                End With
            Case TargTypePlayer
                If B = Character.Index Then
                    Call CreateTileEffect(cx, cy, 12, 75, 6, 0, 0)
                Else
                    With Player(C)
                        Call CreateTileEffect(.X, .Y, 12, 75, 6, 0, 0)
                    End With
                End If
        End Select
    Case SplFireCross
        B = Asc(Mid$(St, 2, 1)) 'Target Type
        C = Asc(Mid$(St, 3, 1)) 'Target
        Select Case B
            Case TargTypeMonster
                With Map.Monster(C)
                    Call CreateTileEffect(.X, .Y, 20, 75, 4, 0, 0)
                End With
            Case TargTypePlayer
                If B = Character.Index Then
                    Call CreateTileEffect(cx, cy, 20, 75, 4, 0, 0)
                Else
                    With Player(C)
                        Call CreateTileEffect(.X, .Y, 20, 75, 4, 0, 0)
                    End With
                End If
        End Select
    Case SplDash
        B = Asc(Mid$(St, 2, 1)) 'Target Type
        C = Asc(Mid$(St, 3, 1)) 'Target
        Select Case B
            Case TargTypeMonster
                With Map.Monster(C)
                    Call CreateTileEffect(.X, .Y, 59, 75, 4, 0, 0)
                End With
            Case TargTypePlayer
                If B = Character.Index Then
                    Call CreateTileEffect(cx, cy, 59, 75, 4, 0, 0)
                Else
                    With Player(C)
                        Call CreateTileEffect(.X, .Y, 59, 75, 4, 0, 0)
                    End With
                End If
        End Select
    Case SplCrush
        B = Asc(Mid$(St, 2, 1)) 'Target Type
        C = Asc(Mid$(St, 3, 1)) 'Target
        Select Case B
            Case TargTypeMonster
                With Map.Monster(C)
                    Call CreateTileEffect(.X, .Y, 40, 75, 8, 0, 0)
                End With
            Case TargTypePlayer
                If B = Character.Index Then
                    Call CreateTileEffect(cx, cy, 40, 75, 8, 0, 0)
                Else
                    With Player(C)
                        Call CreateTileEffect(.X, .Y, 40, 75, 8, 0, 0)
                    End With
                End If
        End Select
    Case SplFireBomb
        B = Asc(Mid$(St, 2, 1)) 'Target Type
        C = Asc(Mid$(St, 3, 1)) 'Target
        Select Case B
            Case TargTypeMonster
                With Map.Monster(C)
                    Call CreateTileEffect(.X, .Y, 19, 75, 5, 0, 0)
                End With
            Case TargTypePlayer
                If B = Character.Index Then
                    Call CreateTileEffect(cx, cy, 19, 75, 5, 0, 0)
                Else
                    With Player(C)
                        Call CreateTileEffect(.X, .Y, 19, 75, 5, 0, 0)
                    End With
                 End If
          End Select
    Case SplRage
        B = Asc(Mid$(St, 2, 1)) 'Target Type
        C = Asc(Mid$(St, 3, 1)) 'Target
        Select Case B
            Case TargTypeMonster
                With Map.Monster(C)
                    Call CreateTileEffect(.X, .Y, 33, 75, 8, 0, 0)
                End With
            Case TargTypePlayer
                If B = Character.Index Then
                    Call CreateTileEffect(cx, cy, 33, 75, 8, 0, 0)
                Else
                    With Player(C)
                        Call CreateTileEffect(.X, .Y, 33, 75, 8, 0, 0)
                    End With
                 End If
          End Select
    Case SplLightningStab 'Lightning Stab (Paladin)
        B = Asc(Mid$(St, 2, 1)) 'Target Type
        C = Asc(Mid$(St, 3, 1)) 'Target
        Select Case B
            Case TargTypeMonster
                With Map.Monster(C)
                    Call CreateTileEffect(.X, .Y, 54, 75, 8, 0, 0)
                End With
            Case TargTypePlayer
                If B = Character.Index Then
                    Call CreateTileEffect(cx, cy, 54, 75, 8, 0, 0)
                Else
                    With Player(C)
                        Call CreateTileEffect(.X, .Y, 54, 75, 8, 0, 0)
                    End With
                 End If
          End Select
    Case splAssassinate 'Assassinate (Thief)
        B = Asc(Mid$(St, 2, 1)) 'Target Type
        C = Asc(Mid$(St, 3, 1)) 'Target
        Select Case B
            Case TargTypeMonster
                With Map.Monster(C)
                    Call CreateTileEffect(.X, .Y, 15, 75, 2, 0, 0)
                End With
            Case TargTypePlayer
                If B = Character.Index Then
                    Call CreateTileEffect(cx, cy, 15, 75, 2, 0, 0)
                Else
                    With Player(C)
                        Call CreateTileEffect(.X, .Y, 15, 75, 2, 0, 0)
                    End With
                 End If
          End Select
    Case splIceBlade 'Ice Blade (Knight)
        B = Asc(Mid$(St, 2, 1)) 'Target Type
        C = Asc(Mid$(St, 3, 1)) 'Target
        Select Case B
            Case TargTypeMonster
                With Map.Monster(C)
                    Call CreateTileEffect(.X, .Y, 17, 75, 5, 0, 0)
                End With
            Case TargTypePlayer
                If B = Character.Index Then
                    Call CreateTileEffect(cx, cy, 17, 75, 5, 0, 0)
                Else
                    With Player(C)
                        Call CreateTileEffect(.X, .Y, 17, 75, 5, 0, 0)
                    End With
                 End If
          End Select
    Case splPurity
        B = Asc(Mid$(St, 2, 1)) 'Target Type
        C = Asc(Mid$(St, 3, 1)) 'Target
        Select Case B
            Case TargTypePlayer, TargTypeCharacter
                If C = Character.Index Then
                    Call CreateTileEffect(cx, cy, 65, 75, 8, 0, 0)
                Else
                    With Player(C)
                        Call CreateTileEffect(.X, .Y, 65, 75, 8, 0, 0)
                    End With
                End If
        End Select
End Select
End Sub

Public Function GetStatusEffect(Index As Long, Effect As StatusEffects) As Boolean
    If Index = Character.Index Then
        If ((2 ^ (Effect - 1)) And Character.StatusEffect) Then
            GetStatusEffect = True
        End If
    Else
        If ((2 ^ (Effect - 1)) And Player(Index).StatusEffect) Then
            GetStatusEffect = True
        End If
    End If
End Function

Function LOS(Ax As Long, Ay As Long, Bx As Long, By As Long) As Boolean
    Dim A As Long, B As Long
    Dim dX As Long
    dX = Bx - Ax
    Dim dY As Long
    dY = By - Ay
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
    
    If (NumSteps = 0) Then
        LOS = True
        Exit Function
    Else
        Dim CurrentX As Long
        CurrentX = Ax * 64
        Dim CurrentY As Long
        CurrentY = Ay * 64
        Dim Xincr As Long
        Xincr = (dX * 64) / NumSteps
        Dim Yincr As Long
        Yincr = (dY * 64) / NumSteps
        For B = NumSteps To 0 Step -1
            If CurrentX < 0 Then CurrentX = 0
            If CurrentY < 0 Then CurrentY = 0
            A = Map.Tile(Int(CurrentX / 64), Int(CurrentY / 64)).Att
            If A = 1 Or A = 2 Or A = 3 Or A = 10 Or Map.Tile(Int(CurrentX / 64), Int(CurrentY / 64)).WallTile Then
                LOS = False
                Exit Function
            End If
            CurrentX = CurrentX + Xincr
            CurrentY = CurrentY + Yincr
        Next
    End If
LOS = True
End Function





