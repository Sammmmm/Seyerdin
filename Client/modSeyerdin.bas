Attribute VB_Name = "modSeyerdin"
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

''''new doevents tests

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private CurrentProcessHandle As Long

Public Const TitleString = "Seyerdin Online"

Public Const ClientVer = 58 'version

Global FloatingText As New clsFloatTexts
Global Projectiles As New clsProjectiles
Global Effects As New clsEffects
Global ParticleEngineB As New clsParticleEngine
Global ParticleEngineF As New clsParticleEngine

#Const DEBUGWINDOWPROC = 0
#Const USEGETPROP = 0



'SendMessage Constants
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_LINEFROMCHAR = &HC9

'EM_SCROLL EM_LINELENGTH EM_SETRECTNP EM_GETPASSWORDCHAR EM_LINEFROMCHAR EM_SETHANDLE EM_UNDO EM_REPLACESEL EM_GETWORDBREAKPROC EM_GETTHUMB

'SetBkMode Constants
Public Const TRANSPARENT = 1
Public Const OPAQUE = 2

'WaitForTerm
Public Const SYNCHRONIZE = &H100000
Private Const INFINITE = &HFFFFFFFF

'BitBlt Constants
Public Const BLACKNESS = &H42
Public Const WHITENESS = &HFF0062
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const SRCAND = &H8800C6
Public Const NOTSRCCOPY = &H330008
Public Const SRCINVERT = &H660046
Public Const DSTINVERT = &H550009

'DrawText Constants
Public Const DT_TOP = &H0
Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_RIGHT = &H2
Public Const DT_VCENTER = &H4
Public Const DT_BOTTOM = &H8
Public Const DT_WORDBREAK = &H10
Public Const DT_SINGLELINE = &H20
Public Const DT_EXPANDTABS = &H40
Public Const DT_TABSTOP = &H80
Public Const DT_NOCLIP = &H100
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_CALCRECT = &H400
Public Const DT_NOPREFIX = &H800
Public Const DT_INTERNAL = &H1000

Public Const GWL_WNDPROC = -4
Public lpPrevWndProc As Long
Public gHW As Long

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public ClientSocket As Long, SocketData As String



Public Const pttCharacter = 0
Public Const pttPlayer = 1
Public Const pttMonster = 2
Public Const pttTile = 3
Public Const pttProject = 4

Public CurX As Long, CurY As Long, CurSubX As Long, CurSubY As Long

'Map Editing
Public MapEdit As Boolean, EditMode As Byte, drawCooldowns As Boolean
Public CurTile As Integer, TopY As Long
Public CurAnim(1 To 2) As Byte
Public NewAtt As Integer, CurAtt As Integer, CurAttData(0 To 3) As Byte, CurWall As Byte

Public map As MapData, EditMap As MapData, ClipboardMap As MapData
Public TileAnim(0 To 11, 0 To 11) As TileAnimData 'Temporary data for map animations
Public StatusColors(0 To 100) As Long
Public player(1 To MAXUSERS) As PlayerData
Public Monster(1 To MAXITEMS) As MonsterData
Public Object(1 To MAXITEMS) As ObjectData
Public Guild(1 To 255) As GuildData
Public Hall(1 To 255) As HallData
Public NPC(1 To 255) As NPCData
Public SaleItem(0 To 9) As NPCSaleItemData
Public Character As CharacterData
Public Options As OptionsData
Public Macro(0 To 9) As MacroData
Public Prefix(1 To 255) As PrefixType

Public cX As Long, cY As Long, CMap As Long, CMap2 As Long, CX2, CY2  'Your Location
Public CMapString As String
Public CWalkCode As Byte
Public Cxo As Long, CYO As Long, CDir As Long, CWalkStep As Long, cWalkStart As Long
Public CAttack As Long, CWalk As Long
Public cRadius As Long
Public CLastKiller As String, CLastKilled As String

Public NewAccount As Boolean
Public User As String, Pass As String, Email As String

Public frmMain_Loaded As Boolean
Public frmMapEdit_Loaded As Boolean
Public frmMonster_Loaded As Boolean
Public frmObject_Loaded As Boolean
Public frmList_Loaded As Boolean
Public frmLight_Loaded As Boolean
Public frmMapProperties_Loaded As Boolean
Public frmGuilds_Loaded As Boolean
Public frmHallAccess_Loaded As Boolean
Public frmGuild_Loaded As Boolean
Public frmNPC_Loaded As Boolean
Public frmMacros_Loaded As Boolean
Public frmOptions_Loaded As Boolean
Public frmNewGuild_Loaded As Boolean
Public frmBan_Loaded As Boolean
Public frmHall_Loaded As Boolean

Public blnEnd As Boolean, blnPlaying As Boolean

Public keyAlt As Boolean

'Misc Variables
'Public ChannelNames(0 To 14) As String
Public CHannelColors(0 To 14) As Long

Public MapCacheFile As String
Public AttackTimer As Long
Public ChatString As String
Public blnNight As Boolean
Public alternateFrame As Boolean
Public ServerIP As String
Public ServerPort As String
Public ServerId As String
Public ServerHasCustomSkilldata As Boolean
Public ServerHasCustomClasses As Boolean

Public TargetFps As Long
Public CurInvObj As Long
Public CurStorageObj As Long
Public Freeze As Boolean
Public FlickerCount As Byte
Public NextTransition As Long
Public CurrentMIDI As Long

Public SecondCounter As Long, FrameCounter As Long, FrameTimer As Long, FrameRate As Long, combatCounter As Byte, hoverCombat As Boolean
Public CurFrame As Long
Public TempVar1 As Long, TempVar2 As Long, TempVar3 As Long, TempVar4 As Long, TempVar5 As Long, TempVar6 As Long, TempVar7 As Long
Public Tstr As Long, TAgi As Long, TEnd As Long, TWis As Long, TCon As Long, TInt As Long
Public ChatScrollBack As Long
Public RequestedMap As Boolean
Public ComputerID As String
Public CPacketsSent As Byte


Public Section(1 To 30) As String, Suffix As String

Public DisableScreen As Boolean

Public tabDown As Boolean
Public altDown As Long

Public BOOLD3DERROR As Boolean
Public STOPMAPCHECK As Boolean

'Info Text
Public InfoText(0 To 3) As String
Public InfoTextTimer As Long

Public MapData As String * 2379

    Dim t As Long


Sub WaitForTerm(pid As Long)
    Dim phnd As Long
    phnd = OpenProcess(SYNCHRONIZE, 0, pid)
    If phnd <> 0 Then
        Call WaitForSingleObject(phnd, INFINITE)
        Call CloseHandle(phnd)
    End If
End Sub

Sub CreateMapCache()
    Dim St1 As String * 2379, A As Long
    St1 = String$(2379, 0)
    St1 = RC4_EncryptString(St1, MapKey)
    Open MapCacheFile For Random As #1 Len = 2379
    For A = 1 To 5000
        Put #1, , St1
    Next A
    Close #1
End Sub

Sub DelItem(lpAppName As String, lpKeyName As String)
    WritePrivateProfileString lpAppName, lpKeyName, 0&, App.Path + "\Data\Cache\Seyerdin.ini"
End Sub
Sub LoadMacros()
    Dim A As Long
    For A = 0 To 9
        With Macro(A)
            .Text = ReadStr("Macros", "Text" + CStr(A + 1))
            If Len(.Text) > 255 Then .Text = Left$(.Text, 255)
            If ReadInt("Macros", "LineFeed" + CStr(A + 1)) = 0 Then
                .LineFeed = False
            Else
                .LineFeed = True
            End If
        End With
    Next A
End Sub

Sub LoadSkillMacros()
    Dim A As Long, b As Byte
    Dim DatPath As String
    DatPath = "Data\Cache\" & User & ".dat" & ServerId
    If Exists(DatPath) Then
        Open DatPath For Random As #1 Len = 1
            For A = 0 To 9
                Get #1, 1025 + A, b
                Macro(A).Skill = b
                'Get #1, 1025, StSkills
            Next A
        Close #1
    Else
        For A = 0 To 9
            Macro(A).Skill = 0
        Next A
    End If
End Sub

Sub SaveSkillMacros()

    If User <> "" Then
        Dim A As Long
        Dim b As Byte
        Dim DatPath As String
        DatPath = "Data\Cache\" & User & ".dat" & ServerId
        If Exists(DatPath) Then
            Open DatPath For Random As #1 Len = 1
                For A = 0 To 9
                    b = Macro(A).Skill
                    Seek #1, 1025 + A
                    Put #1, , b
                Next A
            Close #1
        End If
    End If
End Sub

Sub MonsterDied(Index As Long)
    Dim tmpEffect As clsEffect
    For Each tmpEffect In Effects
        With tmpEffect
            If .Sprite > 0 Then
                If .TargetType = TT_MONSTER Then
                    If .Target = Index Then
                        If tmpEffect.deathCont Then
                            tmpEffect.TargetType = TT_TILE
                        Else
                            Effects.Remove tmpEffect.Key
                        End If
                    End If
                End If
            End If
        End With
    Next
    If CurrentTarget.TargetType = TT_MONSTER Then
        If CurrentTarget.Target = Index Then
            CurrentTarget.TargetType = 0
        End If
    End If
End Sub
Sub PlayerLeftMap(Index As Long)
    Dim tmpEffect As clsEffect
    
    For Each tmpEffect In Effects
        With tmpEffect
            If .Sprite > 0 Then
                If .TargetType = TT_PLAYER Then
                    If .Target = Index Then
                        Effects.Remove tmpEffect.Key
                    End If
                End If
            End If
        End With
    Next
    
    player(Index).map = 0
    If player(Index).LightSourceNumber > 0 Then
        LightSource(player(Index).LightSourceNumber).Intensity = 0
        LightSource(player(Index).LightSourceNumber).Radius = 0
        LightSource(player(Index).LightSourceNumber).Type = LT_NONE
        player(Index).LightSourceNumber = 0
    End If
End Sub
Function QuadChar(Num As Long) As String
    QuadChar = Chr$(Int(Num / 16777216) Mod 256) + Chr$(Int(Num / 65536) Mod 256) + Chr$(Int(Num / 256) Mod 256) + Chr$(Num Mod 256)
End Function

Sub CreateModString()
ModString(0) = "None"
ModString(1) = "Strength"                   'works
ModString(2) = "Agility"                    'works
ModString(3) = "Endurance"                  'works
ModString(4) = "Constitution"               'works
ModString(5) = "Piety"                      'works
ModString(6) = "Intelligence"               'works
ModString(7) = "HP Regen"                   'works
ModString(8) = "MP Regen"                   'works
ModString(9) = "Magic Resistance %"           'works
ModString(10) = "Indestructible"            'works
ModString(11) = "All Stats"                 'works
ModString(12) = "Increased Magic Item %"    'works
ModString(13) = "Leach HP %"                 'works
ModString(14) = "Increased Attack Speed"    'works (I think..)
ModString(15) = "Increased Critical Hit"   'works
ModString(16) = "Resist Poison %"            'works
ModString(17) = "Durability"                'works

ModString(18) = "Attack Damage"
ModString(19) = "Magic Damage"              'Needs Tested
ModString(20) = "Magic Defense"               'Needs Tested
ModString(21) = "Physical Defense"            'Needs Tested
ModString(22) = "Physical Resistance %"       'Needs Tested
ModString(23) = "Power"                     'Needs Tested
ModString(24) = "Shield Block %"
ModString(25) = "Dodge %"
ModString(26) = "NameOnly"
ModString(27) = "Hidden"
ModString(28) = "Hidden name = Name of Bonus"
ModString(29) = "Name = Name of Bonus"
End Sub

Sub CreateClassData()
    Dim A As Long
    
    For A = 1 To 10
        With Class(A)
            .Name = ReadStr("Class" + CStr(A), "Name", "Classes", ServerHasCustomClasses)
            .StartHP = ReadInt("Class" + CStr(A), "StartHP", "Classes", ServerHasCustomClasses)
            .StartEnergy = ReadInt("Class" + CStr(A), "StartEnergy", "Classes", ServerHasCustomClasses)
            .StartMana = ReadInt("Class" + CStr(A), "StartMana", "Classes", ServerHasCustomClasses)
            .StartStrength = ReadInt("Class" + CStr(A), "StartStrength", "Classes", ServerHasCustomClasses)
            .StartAgility = ReadInt("Class" + CStr(A), "StartAgility", "Classes", ServerHasCustomClasses)
            .StartEndurance = ReadInt("Class" + CStr(A), "StartEndurance", "Classes", ServerHasCustomClasses)
            .StartWisdom = ReadInt("Class" + CStr(A), "StartWisdom", "Classes", ServerHasCustomClasses)
            .StartConstitution = ReadInt("Class" + CStr(A), "StartConstitution", "Classes", ServerHasCustomClasses)
            .StartIntelligence = ReadInt("Class" + CStr(A), "StartIntelligence", "Classes", ServerHasCustomClasses)
            .Description = ReadStr("Class" + CStr(A), "Description", "Classes", ServerHasCustomClasses)
            .MaleSprite = ReadInt("Class" + CStr(A), "MaleSprite", "Classes", ServerHasCustomClasses)
            .FemaleSprite = ReadInt("Class" + CStr(A), "FemaleSprite", "Classes", ServerHasCustomClasses)
            .Enabled = ReadInt("Class" + CStr(A), "Enabled", "Classes", ServerHasCustomClasses)
        End With
    Next A
End Sub

Sub MoveToTile()
    SendSocket Chr$(7) + Chr$(CWalkCode) + Chr$(cX * 16 + cY) + Chr$((CWalkStep * 8) + CDir)
    If GetStatusEffect(Character.Index, SE_BERSERK) And Character.SkillLevels(SKILL_BLOODTHIRSTY) > 0 Then CWalkStep = CWalkStep + 2
    If GetStatusEffect(Character.Index, SE_INVISIBLE) And Character.SkillLevels(SKILL_OPPORTUNIST) > 0 Then CWalkStep = CWalkStep + 1
    If GetStatusEffect(Character.Index, SE_LOSTARCANA) And CWalkStep > 2 Then CWalkStep = CWalkStep - 2
    If Character.HP < Character.MaxHP / 4 And Character.SkillLevels(SKILL_OPPORTUNIST) > 0 Then CWalkStep = CWalkStep + 2
    
    If map.Tile(cX, cY).Att = 2 Then
        Freeze = True
        Transition 2, 0, 0, 0, 10
    ElseIf map.Tile(cX, cY).Att = 9 Then
        FloatingText.Add cX * 32, cY * 32, CStr(map.Tile(cX, cY).AttData(0)), &HFFFF0000
    End If
    
    If STOPMAPCHECK = True Then
        STOPMAPCHECK = False
        If MiniMapTab = tsMap Then
            CreateMiniMap
        End If
    End If
End Sub

Sub SaveOptions()
    Dim A As Long
    With Options
        WriteString "Options", "Saved", "1"
        If .MIDI = True Then
            WriteString "Options", "MIDI", "1"
        Else
            WriteString "Options", "MIDI", "0"
        End If
        If .Wav = True Then
            WriteString "Options", "Wav", "1"
        Else
            WriteString "Options", "Wav", "0"
        End If
        If .Broadcasts = True Then
            WriteString "Options", "Broadcasts", "1"
        Else
            WriteString "Options", "Broadcasts", "0"
        End If
        If .Says = True Then
            WriteString "Options", "Says", "1"
        Else
            WriteString "Options", "Says", "0"
        End If
        If .Says = True Then
            WriteString "Options", "Tells", "1"
        Else
            WriteString "Options", "Tells", "0"
        End If
        If .Emotes = True Then
            WriteString "Options", "Emotes", "1"
        Else
            WriteString "Options", "Emotes", "0"
        End If
        If .Yells = True Then
            WriteString "Options", "Yells", "1"
        Else
            WriteString "Options", "Yells", "0"
        End If
        If .fullredraws = True Then
            WriteString "Options", "FullRedraws", "1"
        Else
            WriteString "Options", "FullRedraws", "0"
        End If
        If .Away = True Then
            WriteString "Options", "Away", "1"
        Else
            WriteString "Options", "Away", "0"
        End If
        If .dontDisplayHelms = True Then
            WriteString "Options", "dontDisplayHelms", "1"
        Else
            WriteString "Options", "dontDisplayHelms", "0"
        End If
        If .ShowHP = True Then
            WriteString "Options", "ShowHP", "1"
        Else
            WriteString "Options", "ShowHP", "0"
        End If
        If .MName = True Then
            WriteString "Options", "MName", "1"
        Else
            WriteString "Options", "MName", "0"
        End If
        If .WalkSound = True Then
            WriteString "Options", "WalkSound", "1"
        Else
            WriteString "Options", "WalkSound", "0"
        End If
        If .AutoRun = True Then
            WriteString "Options", "Autorun", "1"
        Else
            WriteString "Options", "Autorun", "0"
        End If
        WriteString "Options", "LightQuality", .LightQuality
        WriteString "Options", "FontSize", .FontSize
        WriteString "Options", "MusicVolume", .MusicVolume
        WriteString "Options", "SoundVolume", .SoundVolume
        WriteString "Options", "PauseTime", .pausetime
        
        WriteString "Options", "ResolutionIndex", .ResolutionIndex
        WriteString "Options", "TargetFps", .TargetFps
        
        If .DisableMultiSampling = True Then
            WriteString "Options", "DisableMultiSampling", "1"
        Else
            WriteString "Options", "DisableMultiSampling", "0"
        End If
        
        If .VsyncEnabled = True Then
            WriteString "Options", "VsyncEnabled", "1"
        Else
            WriteString "Options", "VsyncEnabled", "0"
        End If
        
  
  
        If .highpriority = True Then
            WriteString "Options", "priority", "1"
        Else
            WriteString "Options", "priority", "0"
        End If
        If .hightask = True Then
            WriteString "Options", "hightask", "1"
            'SetPriorityClass CurrentProcessHandle, HIGH_PRIORITY_CLASS
        Else
            WriteString "Options", "hightask", "0"
            'SetPriorityClass CurrentProcessHandle, NORMAL_PRIORITY_CLASS
        End If
        If .ShowFog = True Then
            WriteString "Options", "ShowFog", "1"
        Else
            WriteString "Options", "ShowFog", "0"
        End If
        

        'If .AltKeysEnabled = True Then
        '    WriteString "Options", "AlternateBindings", "1"
        'Else
        '    WriteString "Options", "AlternateBindings", "0"
        'End If

        WriteString "Options", "UpKey", .UpKey
        WriteString "Options", "DownKey", .DownKey
        WriteString "Options", "LeftKey", .LeftKey
        WriteString "Options", "RightKey", .RightKey
        WriteString "Options", "AttackKey", .AttackKey
        WriteString "Options", "RunKey", .RunKey
        WriteString "Options", "StrafeKey", .StrafeKey
        WriteString "Options", "PickUpKey", .PickupKey
        WriteString "Options", "CycleKey", .CycleKey
        
        WriteString "Options", "ChatKey", .ChatKey
        WriteString "Options", "BroadcastKey", .BroadcastKey
        WriteString "Options", "SayKey", .SayKey
        WriteString "Options", "TellKey", .TellKey
        WriteString "Options", "GuildKey", .GuildKey
        WriteString "Options", "PartyKey", .PartyKey
        
        
        For A = 0 To 9
            WriteString "Options", "SpellKey" + Str(A + 1), .SpellKey(A)
        Next A

        If Options.TargetFps > 0 Then
            TargetFps = 1000 / Options.TargetFps
        Else
            TargetFps = 0
        End If
        
    End With
End Sub
Sub LoadOptions(Optional settingDefaults As Boolean = False)
    Dim A As Long
    With Options
        If ReadInt("Options", "Saved") = 1 And settingDefaults = False Then
            If ReadInt("Options", "MIDI") = 1 Then
                .MIDI = True
            Else
                .MIDI = False
            End If
            If ReadInt("Options", "Wav") = 1 Then
                .Wav = True
            Else
                .Wav = False
            End If
            If ReadInt("Options", "Broadcasts") = 1 Then
                .Broadcasts = True
            Else
                .Broadcasts = False
            End If
            If ReadInt("Options", "Tells") = 0 Then
                .Tells = False
            Else
                .Tells = True
            End If
            If ReadInt("Options", "Yells") = 0 Then
                .Yells = False
            Else
                .Yells = True
            End If
            If ReadInt("Options", "Emotes") = 0 Then
                .Emotes = False
            Else
                .Emotes = True
            End If
            If ReadInt("Options", "Says") = 0 Then
                .Says = False
            Else
                .Says = True
            End If
            If ReadInt("Options", "FullRedraws") = 0 Then
                .fullredraws = False
            Else
                .fullredraws = True
            End If
            If ReadInt("Options", "dontDisplayHelms") = 0 Then
                .dontDisplayHelms = False
            Else
                .dontDisplayHelms = True
            End If
            If ReadInt("Options", "Away") = 1 Then
                .Away = True
                .AwayMsg = "Sorry, I'm not here to answer you. Will be back soon. :)"
            Else
                .Away = False
            End If
            .AwayMsg = ReadStr("Options", "Amsg")
            'If ReadInt("Options", "MName") = 1 Then
            '    .MName = True
            'Else
                .MName = False
            'End If
            If ReadInt("Options", "WalkSound") = 1 Then
                .WalkSound = True
            Else
                .WalkSound = False
            End If
            If ReadInt("Options", "Autorun") = 1 Then
                .AutoRun = True
            Else
                .AutoRun = False
            End If
            .LightQuality = ReadInt("Options", "LightQuality")
            .FontSize = ReadInt("Options", "FontSize")
            If .FontSize < 8 Or .FontSize > 12 Then .FontSize = 8
            .MusicVolume = ReadInt("Options", "MusicVolume")
            If .MusicVolume = 0 Then .MusicVolume = 64
            .SoundVolume = ReadInt("Options", "SoundVolume")
            If .SoundVolume = 0 Then .SoundVolume = 64
            .pausetime = ReadInt("Options", "PauseTime")
            
            
            If ReadInt("Options", "DisableMultiSampling") = 1 Then
                .DisableMultiSampling = True
            Else
                .DisableMultiSampling = False
            End If
            
            If ReadInt("Options", "VsyncEnabled") = 1 Then
                .VsyncEnabled = True
            Else
                .VsyncEnabled = False
            End If
        
            
            .ResolutionIndex = ReadInt("Options", "ResolutionIndex")
            .TargetFps = ReadInt("Options", "TargetFps")
            
            If .ResolutionIndex = 0 Then .ResolutionIndex = 1
            If .ResolutionIndex > 100 Then .ResolutionIndex = 1
            
            If ReadInt("Options", "priority") = 1 Then
                .highpriority = True
            Else
                .highpriority = False
            End If
            If ReadInt("Options", "hightask") = 1 Then
                .hightask = True
            Else
                .hightask = False
            End If
            If ReadInt("Options", "ShowHP") = 1 Then
                .ShowHP = True
            Else
                .ShowHP = False
            End If
            If ReadInt("Options", "ShowFog") = 1 Then
                .ShowFog = True
            Else
                .ShowFog = False
            End If
            
             'If ReadInt("Options", "AlternateBindings") = 1 Then
             '      .AltKeysEnabled = True
             '  Else
             '      .AltKeysEnabled = False
             '  End If
            .UpKey = ReadInt("Options", "UpKey")
            .DownKey = ReadInt("Options", "DownKey")
            .LeftKey = ReadInt("Options", "LeftKey")
            .RightKey = ReadInt("Options", "Rightkey")
            .PickupKey = ReadInt("Options", "PickUpKey")
            .AttackKey = ReadInt("Options", "AttackKey")
            .RunKey = ReadInt("Options", "RunKey")
            .StrafeKey = ReadInt("Options", "StrafeKey")
            .CycleKey = ReadInt("Options", "CycleKey")
            
            .ChatKey = ReadInt("Options", "ChatKey")
            .BroadcastKey = ReadInt("Options", "BroadcastKey")
            .TellKey = ReadInt("Options", "TellKey")
            .SayKey = ReadInt("Options", "SayKey")
            .GuildKey = ReadInt("Options", "GuildKey")
            .PartyKey = ReadInt("Options", "PartyKey")
            
             
             For A = 0 To 9
                .SpellKey(A) = ReadInt("Options", "SpellKey" + Str(A + 1))
             Next A
        Else
            .UpKey = FindKeyCode(vbKeyUp)
            .DownKey = FindKeyCode(vbKeyDown)
            .LeftKey = FindKeyCode(vbKeyLeft)
            .RightKey = FindKeyCode(vbKeyRight)
            .PickupKey = FindKeyCode(vbKeyReturn)
            .AttackKey = FindKeyCode(vbKeyControl)
            .StrafeKey = FindKeyCode(vbKeyMenu)
            .CycleKey = FindKeyCode(vbKeyTab)
            .RunKey = FindKeyCode(vbKeyShift)
            
            
            
            .ChatKey = FindKeyCode(vbKeyReturn)
            .BroadcastKey = 63
            .TellKey = 64
            .GuildKey = 62
            .SayKey = 66
            .PartyKey = 61
            
            For A = 0 To 9
                .SpellKey(A) = FindKeyCode(112 + A)
            Next A
            
            .MIDI = True
            .Wav = True
            .Broadcasts = True
            .Says = True
            .Yells = True
            .Tells = True
            .Emotes = True
            .Away = False
            .WalkSound = True
            .AutoRun = False
            .AwayMsg = ""
            .FontSize = 10
            .MusicVolume = 64
            .SoundVolume = 64
            .pausetime = 100
            .fullredraws = False
            .ShowHP = True
            .highpriority = True
            .hightask = False
            .ShowFog = True
            
            .ResolutionIndex = 4
            .TargetFps = 40
            .DisableMultiSampling = 0
            .VsyncEnabled = 0
            
            
            If settingDefaults = False Then SaveOptions
        End If
    End With
    
    
    If Options.TargetFps > 0 Then
        TargetFps = 1000 / Options.TargetFps
    Else
        TargetFps = 0
    End If

End Sub
Sub ShowMap()
    DrawMapTitle (map.Name)
    
    If map.MIDI > 0 Then
        If map.MIDI <> CurrentSongNum Then
            If Int(Rnd * 100) < 10 Then
                Sound_PlayStream CLng(map.MIDI)
            End If
        End If
    End If
    
    If frmMain.Visible = False Then
        frmMenu.Hide
        frmMain.Show
        SetTab tsStats
        'frmMain.ZOrder vbBringToFront
    Else
        If NextTransition = 6 Then
            NextTransition = 0
            Transition 1, 255, 0, 0, 10
        Else
            'Transition 1, 0, 0, 0, 1
            DisableScreen = False
        End If
    End If
    
    Dim x As Long
    Dim y As Long
    For x = 0 To 11
    For y = 0 To 11
        mapChangedBg(x, y) = True
    Next y
    Next x
    mapChanged = True
    
        'D3DDevice.SetRenderTarget mapTexture(2).GetSurfaceLevel(0), Nothing, 0
        'D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, ByVal &H0&, 1#, 0
        'D3DDevice.SetRenderTarget mapTexture(1).GetSurfaceLevel(0), Nothing, 0
        'D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, ByVal &H0&, 1#, 0
    
    Freeze = False
End Sub

Sub ClearBit(bytByte, Bit)
      bytByte = bytByte And Not (2 ^ Bit)
End Sub

Sub PlayWav(ByVal Number As Long, Optional distance As Long = 0)
    If Options.Wav = True Then
        'sndPlaySound GetFilename("sound" + CStr(Number) + ".wav"), 1
        
        Dim volumepercent As Single
        
        If distance = 0 Then
            volumepercent = 1
        Else
            volumepercent = 0.75
            volumepercent = 0.75 - 0.05 * distance
            If volumepercent < 0 Then volumepercent = 0
        End If
        
        If Number > 0 And Number <= 255 Then
            Sound_PlaySound Number, volumepercent
        End If
    End If
End Sub

Sub SetBit(bytByte, Bit)
    bytByte = bytByte Or (2 ^ Bit)
End Sub

Function ExamineBit(bytByte, Bit) As Byte
    ExamineBit = ((bytByte And (2 ^ Bit)) > 0)
End Function

Sub ToggleBit(bytByte, Bit)
    If ExamineBit(bytByte, Bit) Then
        ClearBit bytByte, Bit
    Else
        SetBit bytByte, Bit
    End If
End Sub

Sub CloseMapEdit()
    MapEdit = False
    frmMapEdit.Visible = False
    Unload frmMapEdit
    ReDoLightSources
End Sub

Sub DrawChatString()
    Dim A As Long, b As Long, yOff As Long
    Dim Text As String

    frmMain.FillStyle = vbSolid
    If Chat.Enabled Then
        frmMain.ForeColor = vbBlack
        frmMain.FillColor = vbBlack
    Else
        frmMain.ForeColor = RGB(44, 36, 37)
        frmMain.FillColor = RGB(44, 36, 37)
    End If
    Rectangle frmMain.hdc, 6 * WindowScaleX, 582 * WindowScaleY, 781 * WindowScaleX, 597 * WindowScaleY
    
    If WindowScaleY >= 1.4 Then yOff = 4
    
    If ChatString <> "" Then
        frmMain.Font = "Arial"
        frmMain.FontSize = 8 * WindowScaleY
        frmMain.ForeColor = vbWhite
        If frmMain.TextWidth(ChatString) < 770 * WindowScaleX Then
            TextOut frmMain.hdc, 10 * WindowScaleX, 582 * WindowScaleY + yOff, ChatString, Len(ChatString)
        Else
            For A = 1 To Len(ChatString)
                Text = Right(ChatString, A)
                If frmMain.TextWidth(Text) > 767 * WindowScaleX Then
                    Text = Right(Text, Len(Text) - 1)
                    b = frmMain.TextWidth(Text)
                    TextOut frmMain.hdc, (777 * WindowScaleX - b), 582 * WindowScaleY + yOff, Text, Len(Text)
                    Exit For
                End If
            Next A
        End If
    End If
    frmMain.Refresh
End Sub

Sub GetSections(ByVal St As String, NumSections)
    Dim A As Integer, W As Integer, Q As Boolean
    Dim CurChar As String * 1, LastChar As String * 1
    Erase Section
    Suffix = ""
    If Len(St) = 0 Then Exit Sub
    
    W = 1
    For A = 1 To Len(St)
        CurChar = Mid$(St, A, 1)
        Select Case Asc(CurChar)
            Case 32
                If Q = False Then
                    If Not LastChar = Chr$(32) Then W = W + 1
                    If W > NumSections Then Exit For
                Else
                    Section(W) = Section(W) + CurChar
                End If
            Case 34
                If Q = False Then Q = True Else Q = False
            Case Else
                Section(W) = Section(W) + CurChar
        End Select
        LastChar = CurChar
    Next A
    If A < Len(St) Then
        Suffix = Mid$(St, A + 1)
    Else
        Suffix = ""
    End If
End Sub
Sub GetSections2(St)
    Dim A As Long, b As Long, C As Long
    b = 1
    Erase Section
    For A = 1 To 30
        C = InStr(b, St, Chr$(0))
        If C - b = 0 Then
            Section(A) = ""
        ElseIf C <> 0 Then
            Section(A) = Mid$(St, b, C - b)
        Else
            Section(A) = Mid$(St, b, Len(St) - b + 1)
            Exit For
        End If
        b = C + 1
    Next A
End Sub

Sub GetSections3(St)
    Dim A As Long, b As Long, C As Long
    b = 1
    Erase Section
    For A = 1 To 30
        C = InStr(b + 13, St, Chr$(0))
        If C - b = 0 Then
            Section(A) = ""
        ElseIf C <> 0 Then
            Section(A) = Mid$(St, b, C - b)
        Else
            Section(A) = Mid$(St, b, Len(St) - b + 1)
            Exit For
        End If
        b = C + 1
    Next A
End Sub

Function ClipString(St As String) As String
    Dim A As Long
    For A = Len(St) To 1 Step -1
        If Mid$(St, A, 1) <> Chr$(32) And Mid$(St, A, 1) <> Chr$(0) Then
            ClipString = Mid$(St, 1, A)
            Exit Function
        End If
    Next A
End Function

Sub CopyMap(DestMap As MapData, SourceMap As MapData)
    Dim A As Long, x As Long, y As Long
    
    With DestMap
        .Name = SourceMap.Name
        .MIDI = SourceMap.MIDI
        '.NPC = SourceMap.NPC
        .ExitUp = SourceMap.ExitUp
        .ExitDown = SourceMap.ExitDown
        .ExitLeft = SourceMap.ExitLeft
        .ExitRight = SourceMap.ExitRight
        .BootLocation.map = SourceMap.BootLocation.map
        .BootLocation.x = SourceMap.BootLocation.x
        .BootLocation.y = SourceMap.BootLocation.y
        .Flags(0) = SourceMap.Flags(0)
        .Flags(1) = SourceMap.Flags(1)
        .Intensity = SourceMap.Intensity
        .Raining = SourceMap.Raining
        .Snowing = SourceMap.Snowing
        .Zone = SourceMap.Zone
        .Fog = SourceMap.Fog
        .SnowColor = SourceMap.SnowColor
        .RainCOlor = SourceMap.RainCOlor
        For A = 0 To 4
            .MonsterSpawn(A).Monster = SourceMap.MonsterSpawn(A).Monster
            .MonsterSpawn(A).Rate = SourceMap.MonsterSpawn(A).Rate
        Next A
        For y = 0 To 11
            For x = 0 To 11
                With .Tile(x, y)
                    .Ground = SourceMap.Tile(x, y).Ground
                    .Ground2 = SourceMap.Tile(x, y).Ground2
                    .BGTile1 = SourceMap.Tile(x, y).BGTile1
                    .Anim(1) = SourceMap.Tile(x, y).Anim(1)
                    .Anim(2) = SourceMap.Tile(x, y).Anim(2)
                    .FGTile = SourceMap.Tile(x, y).FGTile
                    .Att = SourceMap.Tile(x, y).Att
                    .AttData(0) = SourceMap.Tile(x, y).AttData(0)
                    .AttData(1) = SourceMap.Tile(x, y).AttData(1)
                    .AttData(2) = SourceMap.Tile(x, y).AttData(2)
                    .AttData(3) = SourceMap.Tile(x, y).AttData(3)
                    .WallTile = SourceMap.Tile(x, y).WallTile
                End With
            Next x
        Next y
        For A = 0 To 9
            With SourceMap.Door(A)
                If .Att > 0 Or .WallTile > 0 Then
                    DestMap.Tile(.x, .y).BGTile1 = .BGTile1
                    DestMap.Tile(.x, .y).Att = .Att
                    DestMap.Tile(.x, .y).WallTile = .WallTile
                End If
            End With
            .Door(A).Att = 0
            .Door(A).BGTile1 = 0
            .Door(A).WallTile = 0
        Next A
    End With
End Sub
Sub LoadMap(MapData As String)
    Dim A As Long, x As Long, y As Long
    If Len(MapData) = 2374 Or Len(MapData) = 2379 Then
        With map
            .Name = ClipString$(Mid$(MapData, 1, 30))
            .Version = Asc(Mid$(MapData, 31, 1)) * 16777216 + Asc(Mid$(MapData, 32, 1)) * 65536 + Asc(Mid$(MapData, 33, 1)) * 256& + Asc(Mid$(MapData, 34, 1))
            '.NPC = Asc(Mid$(MapData, 35, 1))
            .MIDI = Asc(Mid$(MapData, 36, 1))
            .ExitUp = Asc(Mid$(MapData, 37, 1)) * 256& + Asc(Mid$(MapData, 38, 1))
            .ExitDown = Asc(Mid$(MapData, 39, 1)) * 256& + Asc(Mid$(MapData, 40, 1))
            .ExitLeft = Asc(Mid$(MapData, 41, 1)) * 256& + Asc(Mid$(MapData, 42, 1))
            .ExitRight = Asc(Mid$(MapData, 43, 1)) * 256& + Asc(Mid$(MapData, 44, 1))
            .BootLocation.map = Asc(Mid$(MapData, 45, 1)) * 256& + Asc(Mid$(MapData, 46, 1))
            .BootLocation.x = Asc(Mid$(MapData, 47, 1))
            .BootLocation.y = Asc(Mid$(MapData, 48, 1))
            .Flags(0) = Asc(Mid$(MapData, 49, 1))
            .Intensity = Asc(Mid$(MapData, 50, 1))
            .Flags(1) = Asc(Mid$(MapData, 51, 1))
            .Raining = GetInt(Mid$(MapData, 52, 2))
            .Snowing = GetInt(Mid$(MapData, 54, 2))
            .Zone = Asc(Mid$(MapData, 56, 1))
            .Fog = Asc(Mid$(MapData, 57, 1))
            .SnowColor = Asc(Mid$(MapData, 58, 1))
            .RainCOlor = Asc(Mid$(MapData, 59, 1))
            If .Fog > 31 Then .Fog = 0
            For A = 0 To 4
                .MonsterSpawn(A).Monster = Asc(Mid$(MapData, 61 + A * 2))
                .MonsterSpawn(A).Rate = Asc(Mid$(MapData, 62 + A * 2))
            Next A

            For y = 0 To 11
                For x = 0 To 11

                    With .Tile(x, y)
                        A = 71 + y * 192 + x * 16
                        .Ground = Asc(Mid$(MapData, A, 1)) * 256& + Asc(Mid$(MapData, A + 1, 1))
                        .Ground2 = Asc(Mid$(MapData, A + 2, 1)) * 256& + Asc(Mid$(MapData, A + 3, 1))
                        .BGTile1 = Asc(Mid$(MapData, A + 4, 1)) * 256& + Asc(Mid$(MapData, A + 5, 1))
                        .Anim(1) = Asc(Mid$(MapData, A + 6, 1))
                        .Anim(2) = Asc(Mid$(MapData, A + 7, 1))
                        TileAnim(x, y).Frame = (.Anim(2) \ 4) And 15
                        TileAnim(x, y).Frame2 = 0
                        TileAnim(x, y).AnimDelay = 0
                        .FGTile = Asc(Mid$(MapData, A + 8, 1)) * 256& + Asc(Mid$(MapData, A + 9, 1))
                        .Att = Asc(Mid$(MapData, A + 10, 1))
                        .AttData(0) = Asc(Mid$(MapData, A + 11, 1))
                        .AttData(1) = Asc(Mid$(MapData, A + 12, 1))
                        .AttData(2) = Asc(Mid$(MapData, A + 13, 1))
                        .AttData(3) = Asc(Mid$(MapData, A + 14, 1))
                        If .Att = 25 Then
                            If .AttData(0) > 0 Then
                                FloatingText.Add x * 32 + 16, y * 32 - 13, NPC(.AttData(0)).Name, NPCStatusColors(1), False, 255, NPC(.AttData(0)).Name
                            End If
                        End If
                        .WallTile = Asc(Mid$(MapData, A + 15, 1))
                    End With
                Next x
            Next y
            If Len(MapData) = 2379 Then
                .MonsterSpawn(0).Monster = .MonsterSpawn(0).Monster + Asc(Mid$(MapData, 2375)) * 256
                .MonsterSpawn(1).Monster = .MonsterSpawn(1).Monster + Asc(Mid$(MapData, 2376)) * 256
                .MonsterSpawn(2).Monster = .MonsterSpawn(2).Monster + Asc(Mid$(MapData, 2377)) * 256
                .MonsterSpawn(3).Monster = .MonsterSpawn(3).Monster + Asc(Mid$(MapData, 2378)) * 256
                .MonsterSpawn(4).Monster = .MonsterSpawn(4).Monster + Asc(Mid$(MapData, 2379)) * 256
            End If
        End With
        
        ReDoLightSources
        If map.Raining > 0 And (map.Flags(1) And MAP_RAINING) Then
            InitRain map.Raining
        End If
        If map.Snowing > 0 And (map.Flags(1) And MAP_SNOWING) Then
            InitSnow map.Snowing
        End If
        If map.Fog > 0 And map.Fog <= 31 Then
            InitFog map.Fog, False
        ElseIf World.Fog > 0 Then
            'InitFog World.Fog
        Else
            'ClearFog True
        End If
        
    End If
End Sub
Sub Draw3dText(DC As Long, TargetRect As RECT, St As String, lngColor As Long, Height As Integer)
    Dim ShadowRect As RECT
    With ShadowRect
        .Top = TargetRect.Top + Height
        .Left = TargetRect.Left + Height
        .Bottom = TargetRect.Bottom + Height
        .Right = TargetRect.Right + Height
    End With
    SetBkMode DC, vbTransparent
    SetTextColor DC, RGB(10, 10, 10)
    DrawText DC, St, Len(St), ShadowRect, DT_CENTER Or DT_NOCLIP Or DT_WORDBREAK
    SetTextColor DC, lngColor
    DrawText DC, St, Len(St), TargetRect, DT_CENTER Or DT_NOCLIP Or DT_WORDBREAK
End Sub

Function IsVacant(x As Long, y As Long, FromDir As Byte) As Boolean
    Dim A As Long
    Select Case map.Tile(x, y).Att
        Case 1, 3, 13, 15 'Wall / Key Door / NP Tile / Chest / mine
            Exit Function
        Case 24
            If map.Tile(x, y).AttData(3) > 0 Then Exit Function
        Case 2 'Warp
        
        Case 14 'ClickTile
            If ExamineBit(map.Tile(x, y).AttData(0), 0) Then Exit Function
        Case 17
            Select Case FromDir
                Case 1
                    If ExamineBit(map.Tile(x, y).AttData(0), 0) Then Exit Function
                Case 0
                    If ExamineBit(map.Tile(x, y).AttData(0), 2) Then Exit Function
                Case 3
                    If ExamineBit(map.Tile(x, y).AttData(0), 4) Then Exit Function
                Case 2
                    If ExamineBit(map.Tile(x, y).AttData(0), 6) Then Exit Function
            End Select
    End Select
    Select Case FromDir
        Case 1
            If ExamineBit(map.Tile(x, y).WallTile, 0) Then Exit Function
            If ExamineBit(map.Tile(x, y - 1).WallTile, 5) Then Exit Function
            If map.Tile(x, y - 1).Att = 17 Then If ExamineBit(map.Tile(x, y - 1).AttData(0), 3) Then Exit Function
        Case 0
            If ExamineBit(map.Tile(x, y).WallTile, 1) Then Exit Function
            If ExamineBit(map.Tile(x, y + 1).WallTile, 4) Then Exit Function
            If map.Tile(x, y + 1).Att = 17 Then If ExamineBit(map.Tile(x, y + 1).AttData(0), 1) Then Exit Function
        Case 3
            If ExamineBit(map.Tile(x, y).WallTile, 2) Then Exit Function
            If ExamineBit(map.Tile(x - 1, y).WallTile, 7) Then Exit Function
            If map.Tile(x - 1, y).Att = 17 Then If ExamineBit(map.Tile(x - 1, y).AttData(0), 7) Then Exit Function
        Case 2
            If ExamineBit(map.Tile(x, y).WallTile, 3) Then Exit Function
            If ExamineBit(map.Tile(x + 1, y).WallTile, 6) Then Exit Function
            If map.Tile(x + 1, y).Att = 17 Then If ExamineBit(map.Tile(x + 1, y).AttData(0), 5) Then Exit Function
    End Select
    For A = 0 To 9
        With map.Monster(A)
            If .Monster > 0 Then
                If .x = x And .y = y Then
                    Exit Function
                End If
                If Monster(.Monster).Flags2 And MONSTER_LARGE Then
                    If .x + 1 = x And .y = y Then Exit Function
                    If .x + 1 = x And .y + 1 = y Then Exit Function
                    If .x = x And .y + 1 = y Then Exit Function
                End If
            End If
        End With
    Next A
    For A = 1 To MAXUSERS
        With player(A)
            If .map = CMap And .x = x And .y = y Then
                Exit Function
            End If
        End With
    Next A
    IsVacant = True
End Function
Sub OpenMapEdit()

    MapEdit = True
    CopyMap EditMap, map
    frmMapEdit.Visible = True
    frmMapEdit.ZOrder vbBringToFront
    frmMapEdit.RedrawTiles
    frmMapEdit.RedrawTile
    ReDoLightSources
End Sub

Sub PrepTargetDC(DC As Long)
    SetBkColor DC, RGB(255, 255, 255)
    SetTextColor DC, 0
End Sub

Function ReadInt(lpAppName, lpKeyName$, Optional Filename As String = "Seyerdin", Optional UseServerId As Boolean = True) As Integer
    ReadInt = GetPrivateProfileInt&(lpAppName, lpKeyName$, 0, App.Path + "\Data\Cache\" + Filename + IIf(UseServerId, ServerId, "") + ".ini")
End Function

Function ReadStr(lpAppName, lpKeyName$, Optional Filename As String = "Seyerdin", Optional UseServerId As Boolean = True) As String
    Dim lpReturnedString As String, Valid As Long
    lpReturnedString = Space$(256)
    Valid = GetPrivateProfileString&(lpAppName, lpKeyName, "", lpReturnedString, 256, App.Path + "\Data\Cache\" + Filename + IIf(UseServerId, ServerId, "") + ".ini")
    ReadStr = Left$(lpReturnedString, Valid)
End Function


Function SwearFilter(ByVal St As String) As String
If Options.UseFilter Then
    Dim A As Long
    
    
 '   A = InStr(UCase$(St), "FUCKER")
 '   While A > 0
 '       If Mid$(St, A, 6) = "FUCKER" Then
 '           St = Mid$(St, 1, A - 1) + "LOVER" + Mid$(St, A + 6)
 '       ElseIf Mid$(St, A, 6) = "FUCKER" Then
 '           St = Mid$(St, 1, A - 1) + "Lover" + Mid$(St, A + 6)
 '       Else
 '           St = Mid$(St, 1, A - 1) + "lover" + Mid$(St, A + 6)
 '       End If
 '       A = InStr(UCase$(St), "FUCKER")
 '   Wend
 '
 '   A = InStr(UCase$(St), "FUCK")
 '   While A > 0
 '       If Mid$(St, A, 4) = "FUCK" Then
 '           St = Mid$(St, 1, A - 1) + "LOVE" + Mid$(St, A + 4)
 '       ElseIf Mid$(St, A, 4) = "Fuck" Then
 '           St = Mid$(St, 1, A - 1) + "Love" + Mid$(St, A + 4)
 '       Else
 '           St = Mid$(St, 1, A - 1) + "love" + Mid$(St, A + 4)
 '       End If
 '       A = InStr(UCase$(St), "FUCK")
 '   Wend
 '
 '   A = InStr(UCase$(St), "FACK")
 '   While A > 0
 '       If Mid$(St, A, 4) = "FACK" Then
 '           St = Mid$(St, 1, A - 1) + "LOVE" + Mid$(St, A + 4)
 '       ElseIf Mid$(St, A, 4) = "Fack" Then
 '           St = Mid$(St, 1, A - 1) + "Love" + Mid$(St, A + 4)
 '       Else
 '           St = Mid$(St, 1, A - 1) + "love" + Mid$(St, A + 4)
 '       End If
 '       A = InStr(UCase$(St), "FUCK")
 '   Wend
 '
 '   A = InStr(UCase$(St), "FUK")
 '   While A > 0
 '       If Mid$(St, A, 3) = "FUK" Then
 '           St = Mid$(St, 1, A - 1) + "LOVE" + Mid$(St, A + 3)
 '       ElseIf Mid$(St, A, 3) = "Fuk" Then
 '           St = Mid$(St, 1, A - 1) + "Love" + Mid$(St, A + 3)
 '       Else
 '           St = Mid$(St, 1, A - 1) + "love" + Mid$(St, A + 3)
 '       End If
 '       A = InStr(UCase$(St), "FUK")
 '   Wend
 '
 '   A = InStr(UCase$(St), " FK")
 '   While A > 0
 '       If Mid$(St, A, 3) = " FK" Then
 '           St = Mid$(St, 1, A - 1) + " LOVE" + Mid$(St, A + 3)
 '       ElseIf Mid$(St, A, 3) = "Fk" Then
 '           St = Mid$(St, 1, A - 1) + " Love" + Mid$(St, A + 3)
 '       Else
 '           St = Mid$(St, 1, A - 1) + " love" + Mid$(St, A + 3)
 '       End If
 '       A = InStr(UCase$(St), " FK")
 '   Wend'

'
 '
 '   A = InStr(UCase$(St), "NIGGER")
 '   While A > 0
 '       If Mid$(St, A, 6) = "NIGGER" Then
 '           St = Mid$(St, 1, A - 1) + "SQUELCH ME PLEASE! " + Mid$(St, A + 6)
 '       ElseIf Mid$(St, A, 6) = "Nigger" Then
 '           St = Mid$(St, 1, A - 1) + "Squelch me please! " + Mid$(St, A + 6)
 '       Else
 '           St = Mid$(St, 1, A - 1) + "Squelch me please! " + Mid$(St, A + 6)
 '       End If
 '       A = InStr(UCase$(St), "NIGGER")
 '   Wend
 '
 '   A = InStr(UCase$(St), "NIG ")
 '   While A > 0
 '       If Mid$(St, A, 4) = "NIG " Then
 '           St = Mid$(St, 1, A - 1) + "SQUELCH ME PLEASE! " + Mid$(St, A + 4)
 '       ElseIf Mid$(St, A, 4) = "Nig " Then
 '           St = Mid$(St, 1, A - 1) + "Squelch me please! " + Mid$(St, A + 4)
 '       Else
 '           St = Mid$(St, 1, A - 1) + "Squelch me please! " + Mid$(St, A + 4)
 '       End If
 '       A = InStr(UCase$(St), "NIG ")
 '   Wend
 '
 '   A = InStr(UCase$(St), "NIKKA")
 '   While A > 0
 '       If Mid$(St, A, 5) = "NIKKA" Then
 '           St = Mid$(St, 1, A - 1) + "SQUELCH ME PLEASE! " + Mid$(St, A + 5)
 '       ElseIf Mid$(St, A, 5) = "Nikka" Then
 '           St = Mid$(St, 1, A - 1) + "Squelch me please! " + Mid$(St, A + 5)
 '       Else
 '           St = Mid$(St, 1, A - 1) + "Squelch me please! " + Mid$(St, A + 5)
 '       End If
 '       A = InStr(UCase$(St), "NIKKA")
 '   Wend
 '
    
 '   A = InStr(UCase$(St), "SHIT")
 '   While A > 0
 '       If Mid$(St, A, 4) = "shit" Then
 '           St = Mid$(St, 1, A - 1) + "poop" + Mid$(St, A + 4)
 '       ElseIf Mid$(St, A, 4) = "Shit" Then
 '           St = Mid$(St, 1, A - 1) + "Poop" + Mid$(St, A + 4)
 '       Else
 '           St = Mid$(St, 1, A - 1) + "POOP" + Mid$(St, A + 4)
 '       End If
 '       A = InStr(UCase$(St), "SHIT")
 '   Wend
 '
 '   A = InStr(UCase$(St), "ASSHOLE")
 '   While A > 0
 '       If Mid$(St, A, 7) = "ASSHOLE" Then
 '           St = Mid$(St, 1, A - 1) + "JERK" + Mid$(St, A + 7)
 '       ElseIf Mid$(St, A, 7) = "Asshole" Then
 '           St = Mid$(St, 1, A - 1) + "Jerk" + Mid$(St, A + 7)
 '       Else
 '           St = Mid$(St, 1, A - 1) + "jerk" + Mid$(St, A + 7)
 '       End If
 '       A = InStr(UCase$(St), "ASSHOLE")
 '   Wend
 '
 '   A = InStr(UCase$(St), " ASS ")
 '   While A > 0
 '       If Mid$(St, A, 5) = " ASS " Then
 '           St = Mid$(St, 1, A - 1) + " BUTT " + Mid$(St, A + 5)
 '       ElseIf Mid$(St, A, 5) = " Ass " Then
 '           St = Mid$(St, 1, A - 1) + " Butt " + Mid$(St, A + 5)
 '       Else
 '           St = Mid$(St, 1, A - 1) + " butt " + Mid$(St, A + 5)
 '       End If
 '       A = InStr(UCase$(St), " ASS ")
 '   Wend
    
 '   If Mid$(St, 1, 4) = "ASS " Then St = "BUTT " + Mid$(St, 5)
 '   If Mid$(St, 1, 4) = "Ass " Then St = "Butt " + Mid$(St, 5)
 '   If Mid$(St, 1, 4) = "ass " Then St = "butt " + Mid$(St, 5)
 '
 '   If St = "ASS" Then St = "BUTT"
 '   If St = "Ass" Then St = "Butt"
 '   If St = "ass" Then St = "butt"
 '
 '   'A = InStr(UCase$(St), "DAMNIT")
 '   'While A > 0
 '   '    If Mid$(St, A, 6) = "damnit" Then
 '   '        St = Mid$(St, 1, A - 1) + "darnit" + Mid$(St, A + 6)
 '   '    ElseIf Mid$(St, A, 6) = "Damnit" Then
    '        St = Mid$(St, 1, A - 1) + "Darnit" + Mid$(St, A + 6)
    '    Else
    '        St = Mid$(St, 1, A - 1) + "DARNIT" + Mid$(St, A + 6)
    '    End If
    '    A = InStr(UCase$(St), "DAMNIT")
    'Wend
    
 '   A = InStr(UCase$(St), "BITCH")
 '   While A > 0
 '       If Mid$(St, A, 5) = "BITCH" Then
 '           St = Mid$(St, 1, A - 1) + "BUNNY" + Mid$(St, A + 5)
 '       ElseIf Mid$(St, A, 5) = "Bitch" Then
 '           St = Mid$(St, 1, A - 1) + "Bunny" + Mid$(St, A + 5)
 '       Else
 '           St = Mid$(St, 1, A - 1) + "bunny" + Mid$(St, A + 5)
 '       End If
 '       A = InStr(UCase$(St), "BITCH")
 '   Wend
 '
 '   A = InStr(UCase$(St), "BICH")
 '   While A > 0
 '       If Mid$(St, A, 4) = "BICH" Then
 '           St = Mid$(St, 1, A - 1) + "BUNNY" + Mid$(St, A + 4)
 '       ElseIf Mid$(St, A, 4) = "Bich" Then
 '           St = Mid$(St, 1, A - 1) + "Bunny" + Mid$(St, A + 4)
 '       Else
 '           St = Mid$(St, 1, A - 1) + "bunny" + Mid$(St, A + 4)
 '       End If
 '       A = InStr(UCase$(St), "BICH")
 '   Wend
 '
'End If
    
    A = InStr(UCase$(St), "FUCKER")
    While A > 0
            St = Mid$(St, 1, A - 1) + "*****" + Mid$(St, A + 6)
            A = InStr(UCase$(St), "FUCKER")
    Wend
    
    A = InStr(UCase$(St), "FUCK")
    While A > 0
        If Mid$(St, A, 4) = "FUCK" Then
            St = Mid$(St, 1, A - 1) + "LOVE" + Mid$(St, A + 4)
        ElseIf Mid$(St, A, 4) = "Fuck" Then
            St = Mid$(St, 1, A - 1) + "Love" + Mid$(St, A + 4)
        Else
            St = Mid$(St, 1, A - 1) + "love" + Mid$(St, A + 4)
        End If
        A = InStr(UCase$(St), "FUCK")
    Wend
    
    A = InStr(UCase$(St), "FACK")
    While A > 0
        If Mid$(St, A, 4) = "FACK" Then
            St = Mid$(St, 1, A - 1) + "LOVE" + Mid$(St, A + 4)
        ElseIf Mid$(St, A, 4) = "Fack" Then
            St = Mid$(St, 1, A - 1) + "Love" + Mid$(St, A + 4)
        Else
            St = Mid$(St, 1, A - 1) + "love" + Mid$(St, A + 4)
        End If
        A = InStr(UCase$(St), "FUCK")
    Wend
    
    A = InStr(UCase$(St), "FUK")
    While A > 0
        If Mid$(St, A, 3) = "FUK" Then
            St = Mid$(St, 1, A - 1) + "LOVE" + Mid$(St, A + 3)
        ElseIf Mid$(St, A, 3) = "Fuk" Then
            St = Mid$(St, 1, A - 1) + "Love" + Mid$(St, A + 3)
        Else
            St = Mid$(St, 1, A - 1) + "love" + Mid$(St, A + 3)
        End If
        A = InStr(UCase$(St), "FUK")
    Wend
    
    A = InStr(UCase$(St), " FK")
    While A > 0
        If Mid$(St, A, 3) = " FK" Then
            St = Mid$(St, 1, A - 1) + " LOVE" + Mid$(St, A + 3)
        ElseIf Mid$(St, A, 3) = "Fk" Then
            St = Mid$(St, 1, A - 1) + " Love" + Mid$(St, A + 3)
        Else
            St = Mid$(St, 1, A - 1) + " love" + Mid$(St, A + 3)
        End If
        A = InStr(UCase$(St), " FK")
    Wend

    
    
    A = InStr(UCase$(St), "NIGGER")
    While A > 0
        If Mid$(St, A, 6) = "NIGGER" Then
            St = Mid$(St, 1, A - 1) + "SQUELCH ME PLEASE! " + Mid$(St, A + 6)
        ElseIf Mid$(St, A, 6) = "Nigger" Then
            St = Mid$(St, 1, A - 1) + "Squelch me please! " + Mid$(St, A + 6)
        Else
            St = Mid$(St, 1, A - 1) + "Squelch me please! " + Mid$(St, A + 6)
        End If
        A = InStr(UCase$(St), "NIGGER")
    Wend
    
    A = InStr(UCase$(St), "NIG ")
    While A > 0
        If Mid$(St, A, 4) = "NIG " Then
            St = Mid$(St, 1, A - 1) + "SQUELCH ME PLEASE! " + Mid$(St, A + 4)
        ElseIf Mid$(St, A, 4) = "Nig " Then
            St = Mid$(St, 1, A - 1) + "Squelch me please! " + Mid$(St, A + 4)
        Else
            St = Mid$(St, 1, A - 1) + "Squelch me please! " + Mid$(St, A + 4)
        End If
        A = InStr(UCase$(St), "NIG ")
    Wend
    
    A = InStr(UCase$(St), "NIKKA")
    While A > 0
        If Mid$(St, A, 5) = "NIKKA" Then
            St = Mid$(St, 1, A - 1) + "SQUELCH ME PLEASE! " + Mid$(St, A + 5)
        ElseIf Mid$(St, A, 5) = "Nikka" Then
            St = Mid$(St, 1, A - 1) + "Squelch me please! " + Mid$(St, A + 5)
        Else
            St = Mid$(St, 1, A - 1) + "Squelch me please! " + Mid$(St, A + 5)
        End If
        A = InStr(UCase$(St), "NIKKA")
    Wend
    
    
    A = InStr(UCase$(St), "SHIT")
    While A > 0
        If Mid$(St, A, 4) = "shit" Then
            St = Mid$(St, 1, A - 1) + "poop" + Mid$(St, A + 4)
        ElseIf Mid$(St, A, 4) = "Shit" Then
            St = Mid$(St, 1, A - 1) + "Poop" + Mid$(St, A + 4)
        Else
            St = Mid$(St, 1, A - 1) + "POOP" + Mid$(St, A + 4)
        End If
        A = InStr(UCase$(St), "SHIT")
    Wend
    
    A = InStr(UCase$(St), "ASSHOLE")
    While A > 0
        If Mid$(St, A, 7) = "ASSHOLE" Then
            St = Mid$(St, 1, A - 1) + "JERK" + Mid$(St, A + 7)
        ElseIf Mid$(St, A, 7) = "Asshole" Then
            St = Mid$(St, 1, A - 1) + "Jerk" + Mid$(St, A + 7)
        Else
            St = Mid$(St, 1, A - 1) + "jerk" + Mid$(St, A + 7)
        End If
        A = InStr(UCase$(St), "ASSHOLE")
    Wend
        
    A = InStr(UCase$(St), " ASS ")
    While A > 0
        If Mid$(St, A, 5) = " ASS " Then
            St = Mid$(St, 1, A - 1) + " BUTT " + Mid$(St, A + 5)
        ElseIf Mid$(St, A, 5) = " Ass " Then
            St = Mid$(St, 1, A - 1) + " Butt " + Mid$(St, A + 5)
        Else
            St = Mid$(St, 1, A - 1) + " butt " + Mid$(St, A + 5)
        End If
        A = InStr(UCase$(St), " ASS ")
    Wend
    
    If Mid$(St, 1, 4) = "ASS " Then St = "BUTT " + Mid$(St, 5)
    If Mid$(St, 1, 4) = "Ass " Then St = "Butt " + Mid$(St, 5)
    If Mid$(St, 1, 4) = "ass " Then St = "butt " + Mid$(St, 5)
    
    If St = "ASS" Then St = "BUTT"
    If St = "Ass" Then St = "Butt"
    If St = "ass" Then St = "butt"
    
    'A = InStr(UCase$(St), "DAMNIT")
    'While A > 0
    '    If Mid$(St, A, 6) = "damnit" Then
    '        St = Mid$(St, 1, A - 1) + "darnit" + Mid$(St, A + 6)
    '    ElseIf Mid$(St, A, 6) = "Damnit" Then
    '        St = Mid$(St, 1, A - 1) + "Darnit" + Mid$(St, A + 6)
    '    Else
    '        St = Mid$(St, 1, A - 1) + "DARNIT" + Mid$(St, A + 6)
    '    End If
    '    A = InStr(UCase$(St), "DAMNIT")
    'Wend
    
    A = InStr(UCase$(St), "BITCH")
    While A > 0
        If Mid$(St, A, 5) = "BITCH" Then
            St = Mid$(St, 1, A - 1) + "BUNNY" + Mid$(St, A + 5)
        ElseIf Mid$(St, A, 5) = "Bitch" Then
            St = Mid$(St, 1, A - 1) + "Bunny" + Mid$(St, A + 5)
        Else
            St = Mid$(St, 1, A - 1) + "bunny" + Mid$(St, A + 5)
        End If
        A = InStr(UCase$(St), "BITCH")
    Wend
    
    A = InStr(UCase$(St), "BICH")
    While A > 0
        If Mid$(St, A, 4) = "BICH" Then
            St = Mid$(St, 1, A - 1) + "BUNNY" + Mid$(St, A + 4)
        ElseIf Mid$(St, A, 4) = "Bich" Then
            St = Mid$(St, 1, A - 1) + "Bunny" + Mid$(St, A + 4)
        Else
            St = Mid$(St, 1, A - 1) + "bunny" + Mid$(St, A + 4)
        End If
        A = InStr(UCase$(St), "BICH")
    Wend


End If
SwearFilter = St
End Function

Sub UpdatePlayerColor(Index As Long)
    Dim A As Long
    With player(Index)
        If .Guild > 0 Then
            If .Guild = Character.Guild Then
                .Color = &HFF00FFFF
            Else
                .Color = &HFFFFFFFF
                If Character.Guild > 0 Then
                    For A = 0 To 9
                        If Character.GuildDeclaration(A).Guild = .Guild Then
                            If Character.GuildDeclaration(A).Type = 0 Then
                                .Color = &HFF00FF00
                            Else
                                .Color = &HFFFF0000
                            End If
                        End If
                    Next A
                End If
            End If
        Else
            .Color = &HFFC0C0C0
        End If
    End With
End Sub
Sub UpdatePlayersColors()
    Dim A As Long
    For A = 1 To MAXUSERS
        UpdatePlayerColor A
    Next A
End Sub
Sub UpdateSaleItem(A As Long)
    Dim St As String
    With SaleItem(A)
        If .GiveObject >= 1 And .TakeObject >= 1 Then
            St = CStr(A) + ": "
            If Object(.GiveObject).Type = 6 Then
                'Money
                St = St + CStr(.GiveValue) + " " + Object(.GiveObject).Name
            Else
                St = St + Object(.GiveObject).Name
            End If
            St = St + " in exchange for "
            If Object(.TakeObject).Type = 6 Then
                'Money
                St = St + CStr(.TakeValue) + " " + Object(.TakeObject).Name
            Else
                St = St + Object(.TakeObject).Name
            End If
            frmNPC.lstSaleItems.List(A) = St
        Else
            frmNPC.lstSaleItems.List(A) = CStr(A) + ":"
        End If
    End With
End Sub
Sub WriteString(lpAppName, lpKeyName As String, A)
    Dim lpString As String, Valid As Long
    lpString = A
    Valid = WritePrivateProfileString&(lpAppName, lpKeyName, lpString, App.Path + "\Data\Cache\Seyerdin" + ServerId + ".ini")
End Sub

Sub CheckKeys()
    Dim A As Long, b As Long, C As Long, D As Long
    A = 0
    If GetFocus <> frmMain.hwnd Then
        A = GetFocus
        Do Until A = 0 Or b = 1
            If GetParent(A) = frmMain.hwnd Then
                b = 1
            End If
            For C = 1 To 5
                If Not chatForms(C) Is Nothing Then
                    If chatForms(C).Visible Then
                        b = 1
                    End If
                End If
            Next C
            A = GetParent(A)
        Loop
        If b = 0 Then Exit Sub
    End If
    DoEvents
    If KeyDown(Options.CycleKey) Then tabDown = False
    
    'If GetKeyState(VK_ALT) < 0 And GetTickCount > altDown Then
    '    If GetKeyState(vbKeyB) < 0 Then
    '        altDown = GetTickCount + 100
    '        Options.Broadcasts = IIf(Options.Broadcasts, False, True)
    '        SaveOptions
    '        DrawChat
    '    End If
   ' End If

    If SkillListBox.MouseState = False Then
        If Not GetKeyState(vbKeyMenu) < 0 Then
            For A = 0 To 9
                If KeyDown(Options.SpellKey(A)) Then UseMacro (A)
            Next A
        End If
    Else
        With SkillListBox
            .MouseState = True
           ' frmMain.DrawLstBox
            If .Data(.Selected) > 0 Then
                'Set Macro
                C = 0
                For b = 0 To 9
                    If KeyDown(Options.SpellKey(b)) < 0 Then
                        If .Selected > 0 Then
                            If .Data(.Selected) <= MAX_SKILLS Then
                                For A = 0 To 9
                                    If Macro(A).Skill = .Data(.Selected) Then
                                        Macro(A).Skill = 0
                                        C = 1
                                    End If
                                Next A
                                Macro(b).Skill = .Data(.Selected)
                                C = 1
                                frmMain.DrawLstBox
                            End If
                            Exit For
                        End If
                    End If
                Next b
                If C Then
                    SaveSkillMacros
                End If
            End If
        End With
    End If
    
       ' If KeyDown(Options.PickupKey) Then
       '     If Character.Trading = False Then
       '         'Pick up object
       '         For A = 0 To 49
       '             With map.Object(A)
       '                 If .Object > 0 And .X = cX And .Y = cY Then
       '                     If Not ExamineBit(Object(.Object).Flags, 7) Then
       '                         SendSocket Chr$(8)
       '                         If Options.PickupKey = Options.ChatKey Then Chat.Enabled = False
       '                         DrawChatString
       '                         Exit For
       '                     End If
       '                 End If
       '             End With
       '         Next A
       '     End If
        'End If
    If (KeyDown(Options.AttackKey)) And Character.Frozen = False And Freeze = False Then
        If Not KeyDown(Options.CycleKey) < 0 Or tabDown Then
            If GetTickCount - AttackTimer >= Character.AttackSpeed Then
            Dim didAttack As Boolean
                resetAttack
                AttackTimer = GetTickCount
                Dim tx As Long, ty As Long
                Select Case CDir
                    Case 0
                        tx = cX
                        ty = cY - 1
                    Case 1
                        tx = cX
                        ty = cY + 1
                    Case 2
                        tx = cX - 1
                        ty = cY
                    Case 3
                        tx = cX + 1
                        ty = cY
                End Select
                If tx >= 0 And tx <= 11 And ty >= 0 And ty <= 11 Then
                    D = 0
                    If Character.Equipped(1).Object > 0 Then
                        If Object(Character.Equipped(1).Object).Type = 10 Then
                            If Character.Equipped(2).Object > 0 Then
                                If Object(Character.Equipped(2).Object).Type = 11 Then
                                    If Object(Character.Equipped(1).Object).ObjData(3) = Object(Character.Equipped(2).Object).ObjData(4) Then
                                        If Len(Character.CurProjectile.Key) = 0 Then 'Attack with a projectile
                                            Character.CurProjectile.Key = Projectiles.Add(256, CByte(CDir), 0.5, cX, cY, Character.Index, 0, 0, 0, 0, 0).Key
                                            SendSocket Chr$(24) + Chr$(0) + Chr(CDir)
                                            D = 1
                                        End If
                                    Else
                                        PrintChat "This ammo does not work with this ranged weapon!", 7, Options.FontSize
                                    End If
                                Else
                                    PrintChat "You do not have any ammo!", 7, Options.FontSize
                                End If
                            Else
                                PrintChat "You do not have any ammo!", 7, Options.FontSize
                            End If
                        End If
                    End If
                    If D = 0 Then
                        If CanAttack(tx, ty) Then
                            For A = 0 To 9
                                With map.Monster(A)
                                    If .Monster > 0 Then
                                        If .x = tx And .y = ty Then
                                            TargetMonster = A
                                            didAttack = True
                                            SendSocket Chr$(26) + Chr$(A)
                                            Exit For
                                        End If
                                        If Monster(.Monster).Flags2 And MONSTER_LARGE Then
                                            If .x + 1 = tx And .y = ty Then
                                                TargetMonster = A
                                                didAttack = True
                                                SendSocket Chr$(26) + Chr$(A)
                                                Exit For
                                            End If
                                            If .x + 1 = tx And .y + 1 = ty Then
                                                TargetMonster = A
                                                didAttack = True
                                                SendSocket Chr$(26) + Chr$(A)
                                                Exit For
                                            End If
                                            If .x = tx And .y + 1 = ty Then
                                                TargetMonster = A
                                                didAttack = True
                                                SendSocket Chr$(26) + Chr$(A)
                                                Exit For
                                            End If
                                        End If
                                    End If
                                End With
                            Next A
                            If A = 10 Then
                                For A = 1 To MAXUSERS
                                    With player(A)
                                        If .map = CMap And .Sprite > 0 Then
                                            b = 0
                                            If .x = tx And .y = ty Then b = 1
                                            C = .XO Mod 32
                                            D = .YO Mod 32
                                            If b = 0 Then
                                                If .D = CDir Or (.D = 1 And CDir = 0) Or (.D = 0 And CDir = 1) Or (.D = 2 And CDir = 3) Or (.D = 3 And CDir = 2) Then
                                                    b = 6
                                                Else
                                                    b = 10
                                                End If
                                                    If .x = tx + 1 And .y = ty And .D = 3 And Abs(C) <= b Then
                                                        b = 1 'hit while walking right
                                                    End If
                                                    If .x = tx - 1 And .y = ty And .D = 2 And Abs(C) >= 32 - b Then
                                                        b = 1  'hit while walking left
                                                    End If
                                                    If .y = ty + 1 And .x = tx And .D = 1 And Abs(D) <= b Then
                                                        b = 1  'hit while walking down
                                                    End If
                                                    If .y = ty - 1 And .x = tx And .D = 0 And Abs(D) >= 32 - b Then
                                                        b = 1 'hit while walking up
                                                    End If
                                            End If

                                                                               
                                                                               
                                            If b = 1 Then
                                                'If .Party = 0 Or Not .Party <> Character.Party Then
                                                    SendSocket Chr$(25) + Chr$(A)
                                                    combatCounter = 30
                                                    If CurrentTab = tsStats2 And combatCounter = 30 Then
                                                        calculateHpRegen
                                                        calculateManaRegen
                                                        DrawMoreStats
                                                    End If
                                                    didAttack = True
                                                    Exit For
                                                'Else
                                                '    PrintChat "You cannot attack party members!", 7
                                                'End If
                                            End If
                                        End If
                                    End With
                                Next A
                            End If
                            If Not didAttack Then
                                If map.Tile(tx, ty).Att = 24 Then
                                    SendSocket Chr$(89) + Chr$(tx) + Chr$(ty)
                                End If
                                    If combatCounter < 5 Then combatCounter = 5
                                    If CurrentTab = tsStats2 And combatCounter = 5 Then
                                        calculateHpRegen
                                        calculateManaRegen
                                        DrawMoreStats
                                    End If
                                    SendSocket Chr$(84)
                                    CAttack = 5
                                    PlayWav (14)
      
                            Else
                                If combatCounter < 10 Then combatCounter = 10
                                If CurrentTab = tsStats2 And combatCounter = 10 Then
                                    calculateHpRegen
                                    calculateManaRegen
                                    DrawMoreStats
                                End If
                            End If
                            drawTopBar
                        Else
                            If combatCounter < 5 Then combatCounter = 5
                            If CurrentTab = tsStats2 And combatCounter = 5 Then
                                calculateHpRegen
                                calculateManaRegen
                                DrawMoreStats
                            End If
                            If map.Tile(tx, ty).Att = 24 Then
                                SendSocket Chr$(89) + Chr$(tx) + Chr$(ty)
                            End If
                            drawTopBar
                            SendSocket Chr$(84)
                            CAttack = 5
                            PlayWav (14)
                        End If
                        If A = MAXUSERS + 1 Then
                            With map.Tile(tx, ty)
                                If .Att = 14 Then
                                    If ExamineBit(.AttData(0), 2) Then
                                        SendSocket Chr$(76) + Chr$(1) + Chr$(tx) + Chr$(ty)
                                        If combatCounter < 30 Then combatCounter = 30
                                        If CurrentTab = tsStats2 And combatCounter = 30 Then
                                            calculateHpRegen
                                            calculateManaRegen
                                            DrawMoreStats
                                        End If
                                        drawTopBar
                                    End If
                                End If
                            End With
                        End If
                    End If
                End If
            End If
        End If
    End If
    If cX * 32 = Cxo And cY * 32 = CYO Then
        If KeyDown(Options.StrafeKey) And Character.Access >= 5 Then
            If GetTickCount > altDown Then
            altDown = GetTickCount + 40
            CWalkStep = 8
            If (KeyDown(Options.UpKey)) And cY > 0 Then
                cY = cY - 1: CY2 = cY ^ 2 + 5
                SendSocket Chr$(7) + Chr$(CWalkCode) + Chr$(cX * 16 + cY) + Chr$((CWalkStep * 8) + 5)
            ElseIf (KeyDown(Options.DownKey)) And cY < 11 Then
                cY = cY + 1: CY2 = cY ^ 2 + 5
                SendSocket Chr$(7) + Chr$(CWalkCode) + Chr$(cX * 16 + cY) + Chr$((CWalkStep * 8) + 5)
            ElseIf (KeyDown(Options.LeftKey)) And cX > 0 Then
                cX = cX - 1: CX2 = cX ^ 2 + 5
                SendSocket Chr$(7) + Chr$(CWalkCode) + Chr$(cX * 16 + cY) + Chr$((CWalkStep * 8) + 5)
            ElseIf (KeyDown(Options.RightKey)) And cX < 11 Then
                cX = cX + 1: CX2 = cX ^ 2 + 5
                SendSocket Chr$(7) + Chr$(CWalkCode) + Chr$(cX * 16 + cY) + Chr$((CWalkStep * 8) + 5)
            End If
            End If
        Else
            If (KeyDown(Options.StrafeKey)) And (KeyDown(Options.RunKey)) Then
                If GetTickCount > altDown Then
                    If KeyDown(Options.UpKey) Then
                        altDown = GetTickCount + 50
                        CDir = 0
                        SendSocket Chr$(7) + Chr$(CDir)
                    End If
                    
                    If KeyDown(Options.DownKey) Then
                        altDown = GetTickCount + 50
                        CDir = 1
                        SendSocket Chr$(7) + Chr$(CDir)
                    End If
                    
                    If KeyDown(Options.LeftKey) Then
                        altDown = GetTickCount + 50
                        CDir = 2
                        SendSocket Chr$(7) + Chr$(CDir)
                    End If
                    
                    If KeyDown(Options.RightKey) Then
                        altDown = GetTickCount + 50
                        CDir = 3
                        SendSocket Chr$(7) + Chr$(CDir)
                    End If
                End If
            Else
                If GetTickCount - Character.LastMapSwitch > Options.pausetime Then
                    If (KeyDown(Options.UpKey)) And Character.Frozen = 0 Then
                        If CDir = 0 Or (KeyDown(Options.StrafeKey)) Then
                            If cY > 0 Then
                                If IsVacant(cX, cY - 1, CByte(0)) Then
                                    If map.Tile(cX, cY - 1).Att <> 2 Or (GetTickCount - Character.LastMapSwitch > 2000 Or Character.Access > 0 Or (map.Flags(0) And MAP_FRIENDLY)) Then
                                        If map.Tile(cX, cY - 1).Att = 2 Then Character.LastMapSwitch = GetTickCount
                                        cY = cY - 1: CY2 = cY ^ 2 + 5
                                        SetWalkSpeed
                                        cWalkStart = GetTickCount
                                        MoveToTile
                                    End If
                                End If
                            Else
                                If map.ExitUp > 0 And ExamineBit(map.Tile(cX, cY).WallTile, 4) = 0 Then
                                    If GetTickCount - Character.LastMapSwitch > 2000 Or Character.Access > 0 Or (map.Flags(0) And MAP_FRIENDLY) Then
                                        Character.LastMapSwitch = GetTickCount
                                        SendSocket Chr$(13) + Chr$(0)
                                        Freeze = True
                                        'Transition 0, 0, 0, 0, 1
                                    End If
                                End If
                            End If
                        Else
                            CDir = 0
                            SendSocket Chr$(7) + Chr$(CDir)
                        End If
                    ElseIf (KeyDown(Options.DownKey)) And Character.Frozen = 0 Then
                        If CDir = 1 Or (KeyDown(Options.StrafeKey)) Then
                            If cY < 11 Then
                                If IsVacant(cX, cY + 1, CByte(1)) Then
                                    If map.Tile(cX, cY + 1).Att <> 2 Or (GetTickCount - Character.LastMapSwitch > 2000 Or Character.Access > 0 Or (map.Flags(0) And MAP_FRIENDLY)) Then
                                        If map.Tile(cX, cY + 1).Att = 2 Then Character.LastMapSwitch = GetTickCount
                                        cY = cY + 1: CY2 = cY ^ 2 + 5
                                        SetWalkSpeed
                                        cWalkStart = GetTickCount
                                        MoveToTile
                                    End If
                                End If
                            Else
                                If map.ExitDown > 0 And ExamineBit(map.Tile(cX, cY).WallTile, 5) = 0 Then
                                    If GetTickCount - Character.LastMapSwitch > 2000 Or Character.Access > 0 Or (map.Flags(0) And MAP_FRIENDLY) Then
                                        Character.LastMapSwitch = GetTickCount
                                        SendSocket Chr$(13) + Chr$(1)
                                        Freeze = True
                                        'Transition 0, 0, 0, 0, 1
                                    End If
                                End If
                            End If
                        Else
                            CDir = 1
                            SendSocket Chr$(7) + Chr$(CDir)
                            'End If
                        End If
                    ElseIf (KeyDown(Options.LeftKey)) And Character.Frozen = False Then
                        If CDir = 2 Or (KeyDown(Options.StrafeKey)) Then
                            If cX > 0 Then
                                If IsVacant(cX - 1, cY, CByte(2)) Then
                                    If map.Tile(cX - 1, cY).Att <> 2 Or (GetTickCount - Character.LastMapSwitch > 2000 Or Character.Access > 0 Or (map.Flags(0) And MAP_FRIENDLY)) Then
                                        If map.Tile(cX - 1, cY).Att = 2 Then Character.LastMapSwitch = GetTickCount
                                        cX = cX - 1: CX2 = cX ^ 2 + 5
                                        SetWalkSpeed
                                        cWalkStart = GetTickCount
                                        MoveToTile
                                    End If
                                End If
                            Else
                                If map.ExitLeft > 0 And ExamineBit(map.Tile(cX, cY).WallTile, 6) = 0 Then
                                    If GetTickCount - Character.LastMapSwitch > 2000 Or Character.Access > 0 Or (map.Flags(0) And MAP_FRIENDLY) Then
                                        Character.LastMapSwitch = GetTickCount
                                        SendSocket Chr$(13) + Chr$(2)
                                        Freeze = True
                                        'Transition 0, 0, 0, 0, 1
                                    End If
                                End If
                            End If
                        Else
                            CDir = 2
                            SendSocket Chr$(7) + Chr$(CDir)
                        End If
                    ElseIf (KeyDown(Options.RightKey)) And Character.Frozen = False Then
                        If CDir = 3 Or (KeyDown(Options.StrafeKey)) Then
                            If cX < 11 Then
                                If IsVacant(cX + 1, cY, CByte(3)) Then
                                    If map.Tile(cX + 1, cY).Att <> 2 Or (GetTickCount - Character.LastMapSwitch > 2000 Or Character.Access > 0 Or (map.Flags(0) And MAP_FRIENDLY)) Then
                                        If map.Tile(cX + 1, cY).Att = 2 Then Character.LastMapSwitch = GetTickCount
                                        cX = cX + 1: CX2 = cX ^ 2 + 5
                                        SetWalkSpeed
                                        cWalkStart = GetTickCount
                                        MoveToTile
                                    End If
                                End If
                            Else
                                If map.ExitRight > 0 And ExamineBit(map.Tile(cX, cY).WallTile, 7) = 0 Then
                                    If GetTickCount - Character.LastMapSwitch > 2000 Or Character.Access > 0 Or (map.Flags(0) And MAP_FRIENDLY) Then
                                        Character.LastMapSwitch = GetTickCount
                                        SendSocket Chr$(13) + Chr$(3)
                                        Freeze = True
                                        'Transition 0, 0, 0, 0, 1
                                    End If
                                End If
                            End If
                        Else
                            CDir = 3
                            SendSocket Chr$(7) + Chr$(CDir)
                        End If
                    End If
                End If
            End If
        End If
    End If
    If KeyDown(Options.CycleKey) And frmScript.Visible = False And GetTickCount > LastTab Then
    
        LastTab = GetTickCount + 300
        TargetPulse = 150
        'Cycle through targets
        C = 0
        If CurrentTarget.TargetType = TT_CHARACTER Then
            CurrentTarget.TargetType = TT_MONSTER
            CurrentTarget.Target = 0
            If map.Monster(0).Monster > 0 Then
                C = 1
            End If
        End If
        If CurrentTarget.TargetType = TT_MONSTER And C = 0 Then
            A = CurrentTarget.Target
            If A < 9 Then
                For b = A + 1 To 9
                    If map.Monster(b).Monster > 0 Then
                        CurrentTarget.Target = b
                        C = 1
                        Exit For
                    End If
                Next b
                If (b = 9 And CurrentTarget.Target <> 9) Or C = 0 Then
                    CurrentTarget.TargetType = TT_PLAYER
                    CurrentTarget.Target = 1
                    If player(1).map = CMap And player(1).Status <> 9 Then
                        C = 1
                    End If
                End If
            Else
                CurrentTarget.TargetType = TT_PLAYER
                CurrentTarget.Target = 1
                If player(1).map = CMap And player(1).Status <> 9 Then
                    C = 1
                End If
            End If
        End If
        If CurrentTarget.TargetType = TT_PLAYER And C = 0 Then
            A = CurrentTarget.Target
            If A < MAXUSERS Then
                For b = A + 1 To MAXUSERS
                    If player(b).map = CMap And player(b).Status <> 9 Then
                        CurrentTarget.Target = b
                        C = 1
                        Exit For
                    End If
                Next b
                If (b = MAXUSERS And CurrentTarget.Target <> MAXUSERS) Or C = 0 Then
                    CurrentTarget.TargetType = TT_CHARACTER
                End If
            Else
                CurrentTarget.TargetType = TT_CHARACTER
            End If
        End If
        If C = 0 Then CurrentTarget.TargetType = TT_CHARACTER
    End If
End Sub

Sub MovePlayers()
Dim A As Long, b As Long, C As Long, D As Long, x As Long, y As Long, SX As Long, SY As Long, OX As Long, OY As Long

'Move You
If Cxo < cX * 32 Then
    D = Int(Cxo / 16)
    A = (cX * 32 - 32) + ((GetTickCount - cWalkStart) / (3200 / CWalkStep)) * 32
    If A >= cX * 32 Then
        Cxo = cX * 32
    Else
        Cxo = A
    End If
    
    If Int(Cxo / 16) <> D Then
        CWalk = 1 - CWalk
        If CWalk = 0 And Options.WalkSound = True Then PlayWav 4
    End If
    LightSource(0).x = Cxo + 16
    RedoOwnLight
ElseIf Cxo > cX * 32 Then
    D = Int(Cxo / 16)
    A = (cX * 32 + 32) - ((GetTickCount - cWalkStart) / (3200 / CWalkStep)) * 32
    If A <= cX * 32 Then
        Cxo = cX * 32
    Else
        Cxo = A
    End If
    If Int(Cxo / 16) <> D Then
        CWalk = 1 - CWalk
        If CWalk = 0 And Options.WalkSound = True Then PlayWav 4
    End If
    LightSource(0).x = Cxo + 16
    RedoOwnLight
End If
If CYO < cY * 32 Then
    D = Int(CYO / 16)
    A = (cY * 32 - 32) + ((GetTickCount - cWalkStart) / (3200 / CWalkStep)) * 32
    If A >= cY * 32 Then
        CYO = cY * 32
    Else
        CYO = A
    End If
    If Int(CYO / 16) <> D Then
        CWalk = 1 - CWalk
        If CWalk = 0 And Options.WalkSound = True Then PlayWav 4
    End If
    LightSource(0).y = CYO + 16
    RedoOwnLight
ElseIf CYO > cY * 32 Then
    D = Int(CYO / 16)
    A = (cY * 32 + 32) - ((GetTickCount - cWalkStart) / (3200 / CWalkStep)) * 32
    If A <= cY * 32 Then
        CYO = cY * 32
    Else
        CYO = A
    End If
    If Int(CYO / 16) <> D Then
        CWalk = 1 - CWalk
        If CWalk = 0 And Options.WalkSound = True Then PlayWav 4
    End If
    LightSource(0).y = CYO + 16
    RedoOwnLight
End If
If CAttack > 0 Then
    CAttack = CAttack - 1
End If
    
For A = 0 To 9
    With map.Monster(A)
        If .Monster > 0 Then
            C = Monster(.Monster).Sprite
            If C > 0 Then
                If .OX < .x Then 'Walk Right
                    '.XO = .XO + .WalkStep
                    'If Int(.XO / 16) * 16 = .XO Then .W = 1 - .W
                    If GetTickCount >= .EndTick Then
                        .OX = .x
                    Else
                        If .EndTick > .StartTick Then
                            b = ((GetTickCount - .StartTick) / (.EndTick - .StartTick) * 32)
                            If Abs(b) > 16 Then .W = 0
                            .XO = .OX * 32 + b
                        End If
                    End If
                    If .LightSourceNumber > 0 Then LightSource(.LightSourceNumber).x = .XO + 16
                ElseIf .OX > .x Then 'Left
                    '.XO = .XO - .WalkStep
                    'If Int(.XO / 16) * 16 = .XO Then .W = 1 - .W
                    If GetTickCount >= .EndTick Then
                        .OX = .x
                    Else
                        If .EndTick > .StartTick Then
                            b = ((GetTickCount - .StartTick) / (.EndTick - .StartTick) * 32)
                            If Abs(b) > 16 Then .W = 0
                            .XO = .OX * 32 - b
                        End If
                    End If
                    If .LightSourceNumber > 0 Then LightSource(.LightSourceNumber).x = .XO + 16
                End If
                If .OY < .y Then 'Down
                    '.YO = .YO + .WalkStep
                    'If Int(.YO / 16) * 16 = .YO Then .W = 1 - .W
                    If GetTickCount >= .EndTick Then
                        .OY = .y
                    Else
                        If .EndTick > .StartTick Then
                            b = ((GetTickCount - .StartTick) / (.EndTick - .StartTick) * 32)
                            If Abs(b) > 16 Then .W = 0
                            .YO = .OY * 32 + b
                        End If
                    End If
                    If .LightSourceNumber > 0 Then LightSource(.LightSourceNumber).y = .YO + 16
                ElseIf .OY > .y Then 'Up
                    '.YO = .YO - .WalkStep
                    'If Int(.YO / 16) * 16 = .YO Then .W = 1 - .W
                    If GetTickCount >= .EndTick Then
                        .OY = .y
                    Else
                        If .EndTick > .StartTick Then
                            b = ((GetTickCount - .StartTick) / (.EndTick - .StartTick) * 32)
                            If Abs(b) > 16 Then .W = 0
                            .YO = .OY * 32 - b
                        End If
                    End If
                    If .LightSourceNumber > 0 Then LightSource(.LightSourceNumber).y = .YO + 16
                End If
                If .A > 0 Then .A = .A - 1
            End If
        End If
    End With
    With map.DeadBody(A)
        If .Counter > 0 Then
            .Counter = .Counter - 1
            If .Counter = 0 Then
                .Sprite = 0
                .x = 0
                .y = 0
                .Counter = 0
            End If
        End If
    End With
Next A

For A = 1 To MAXUSERS
    With player(A)
        If .map = CMap Then
            'Move Player
            If .XO < .x * 32 Then
                D = Int(.XO / 16)
                b = (.x * 32 - 32) + ((GetTickCount - .WalkStart) / (3200 / .WalkStep)) * 32
                If b >= .x * 32 - 1 Then
                    .XO = .x * 32
                Else
                    .XO = b
                End If
                If player(A).LightSourceNumber > 0 Then LightSource(player(A).LightSourceNumber).x = .XO
                If Int(.XO / 16) <> D Then .W = 1 - .W
            ElseIf .XO > .x * 32 Then
                D = Int(.XO / 16)
                b = (.x * 32 + 32) - ((GetTickCount - .WalkStart) / (3200 / .WalkStep)) * 32
                If b <= .x * 32 Then
                    .XO = .x * 32
                Else
                    .XO = b
                End If
                If player(A).LightSourceNumber > 0 Then LightSource(player(A).LightSourceNumber).x = .XO
                If Int(.XO / 16) <> D Then .W = 1 - .W
            End If
            If .YO < .y * 32 Then
                D = Int(.YO / 16)
                b = (.y * 32 - 32) + ((GetTickCount - .WalkStart) / (3200 / .WalkStep)) * 32
                If b >= .y * 32 - 1 Then
                    .YO = .y * 32
                Else
                    .YO = b
                End If
                If player(A).LightSourceNumber > 0 Then LightSource(player(A).LightSourceNumber).y = .YO
                If Int(.YO / 16) <> D Then .W = 1 - .W
            ElseIf .YO > .y * 32 Then
                D = Int(.YO / 16)
                b = (.y * 32 + 32) - ((GetTickCount - .WalkStart) / (3200 / .WalkStep)) * 32
                If b <= .y * 32 Then
                    .YO = .y * 32
                Else
                    .YO = b
                End If
                If player(A).LightSourceNumber > 0 Then LightSource(player(A).LightSourceNumber).y = .YO
                If Int(.YO / 16) <> D Then .W = 1 - .W
            End If
            If .A > 0 Then .A = .A - 1
        End If
    End With
Next A



    TargetPulse = Sin(GetTickCount / 300) * 75 + 75

    Dim tmpEffect As clsEffect
    For Each tmpEffect In Effects
        With tmpEffect
            If .loopcount >= 0 Then
                If GetTickCount - .TimeStamp >= .Speed Then
                    .TimeStamp = GetTickCount
                    If .Frame < .TotalFrames Then
                        .Frame = .Frame + 1
                    Else
                        .loopcount = .loopcount - 1
                        .Frame = 0
                    End If
                End If
                Select Case .TargetType
                    Case TT_CHARACTER
                        .x = Cxo
                        .y = CYO
                    Case TT_PLAYER
                        If player(.Target).map = CMap Then
                            .x = player(.Target).XO
                            .y = player(.Target).YO
                        Else
                            Effects.Remove tmpEffect.Key
                        End If
                    Case TT_MONSTER
                        If map.Monster(.Target).Monster > 0 Then
                            .x = map.Monster(.Target).XO + .XO
                            .y = map.Monster(.Target).YO + .YO
                        Else
                            Effects.Remove tmpEffect.Key
                        End If
                End Select
            Else
                Effects.Remove tmpEffect.Key
            End If
        End With
    Next

    Dim ft As clsFloatText
    For Each ft In FloatingText
        With ft
            If GetTickCount >= .TimeStamp Then
                '.TimeStamp = GetTickCount + 30
                If .Life > 0 Then
                    If .Life < 255 Then
                        .Life = .Life - 1
                        If .Life = 0 Then FloatingText.Remove .Key
                    End If
                Else
                    .Step = .Step + 1
                    If .Step > 20 Then FloatingText.Remove .Key
                End If
            End If
        End With
    Next
     Dim PJ As clsProjectile
    For Each PJ In Projectiles
        D = 0
        With PJ
            .Frame = .Frame + 1
            If .Frame > 7 Then .Frame = 0
            A = ((GetTickCount - .StartTime) * .Speed)
            OX = .x
            OY = .y
            Select Case .D
                Case 0 'Up
                    .YO = .StartY * 32 - A
                    .y = Abs((.YO - 8) / 32)
                    If (.Reflect And .y <> .reflectY) Then
                        .y = .reflectY
                        If .YO < .y * 32 Then .YO = .y * 32
                    End If
                    If .y <= 11 Then
                        If .y > 0 Then
                            If ExamineBit(map.Flags(1), 5) = False Then
                                If .StartY <> .y Or .StartX <> .x Then If ExamineBit(map.Tile(.x, .y).WallTile, 1) Then D = 2
                            End If
                        End If
                        If .y >= -1 Then If .y + 1 <= .StartY Then If ExamineBit(map.Tile(.x, .y + 1).WallTile, 4) Then D = 2
                    End If
                Case 1 'Down
                    .YO = .StartY * 32 + A
                    .y = ((.YO + 8) / 32)
                    If (.Reflect And .y <> .reflectY) Then
                        .y = .reflectY
                        If .YO > .y * 32 Then .YO = .y * 32
                    End If
                    If .y >= 0 Then
                        If .y < 11 Then
                            If ExamineBit(map.Flags(1), 5) = False Then
                                If .StartY <> .y Or .StartX <> .x Then If ExamineBit(map.Tile(.x, .y).WallTile, 0) Then D = 2
                            End If
                        End If
                        If .y <= 12 Then If .y - 1 >= .StartY Then If ExamineBit(map.Tile(.x, .y - 1).WallTile, 5) Then D = 2
                    End If
                Case 2 'Left
                    .XO = .StartX * 32 - A '- 16
                    .x = Abs((.XO - 8) / 32)
                    If (.Reflect And .x <> .reflectX) Then
                        .x = .reflectX
                        If .XO < .x * 32 Then .XO = .x * 32
                    End If
                    If .x <= 11 Then
                        If .x > 0 Then
                            If ExamineBit(map.Flags(1), 5) = False Then
                                If .StartY <> .y Or .StartX <> .x Then If ExamineBit(map.Tile(.x, .y).WallTile, 3) Then D = 2
                            End If
                        End If
                        If .x >= -1 Then If .x + 1 <= .StartX Then If ExamineBit(map.Tile(.x + 1, .y).WallTile, 6) Then D = 2
                    End If
                Case 3 'Right
                    .XO = .StartX * 32 + A '- 16
                    .x = (.XO + 8) / 32
                    If (.Reflect And .x <> .reflectX) Then
                        .x = .reflectX
                        If .XO > .x * 32 Then .XO = .x * 32
                    End If
                    If .x >= 0 Then
                        If .x < 11 Then
                            If ExamineBit(map.Flags(1), 5) = False Then
                                If .StartY <> .y Or .StartX <> .x Then If ExamineBit(map.Tile(.x, .y).WallTile, 2) Then D = 2
                            End If
                        End If
                        If .x <= 12 Then If .x - 1 >= .StartX Then If ExamineBit(map.Tile(.x - 1, .y).WallTile, 7) Then D = 2
                    End If
            End Select
            

            If .LightSourceNumber > 0 Then
                LightSource(.LightSourceNumber).x = .XO + 16
                LightSource(.LightSourceNumber).y = .YO
            End If
            If Not (.XO > -32 And .XO < 384 And .YO > -32 And .YO < 384) Then
                If .D = 2 Then
                    If ExamineBit(map.Tile(0, .y).WallTile, 6) Then If .Sprite < 256 Then CreateTileEffectXYO -16, .YO, .Sprite, 75, 8, 0
                ElseIf .D = 0 Then
                    If ExamineBit(map.Tile(.x, 0).WallTile, 4) Then If .Sprite < 256 Then CreateTileEffectXYO .XO, -16, .Sprite, 75, 8, 0
                End If
                D = 1
            Else
                SX = .x - OX 'current - old X
                If SX = 0 Then SX = 1
                SX = SX / Abs(SX)
                For x = OX To .x Step SX
                    SY = .y - OY 'current - old X
                    If SY = 0 Then SY = 1
                    SY = SY / Abs(SY)
                    For y = OY To .y Step SY
                        
                            If x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
                                A = map.Tile(x, y).Att
                            Else
                                A = 0
                            End If
                            If .Reflect Or ((A = 28 And ((x <> .StartX And x <> .reflectX) Or (y <> .StartY And y <> .reflectY)))) Then 'mirror
                                b = .D
                                If Not .Reflect And ((x <> .StartX And x <> .reflectX) Or (y <> .StartY And y <> .reflectY)) Then
                                    .Reflect = True
                                    .reflectX = x
                                    .reflectY = y
                                End If
                                
                                If .D = 0 And (.YO - (.reflectY * 32)) <= 0 Then 'UP
                                    .D = 1
                                ElseIf .D = 1 And (.YO - (.reflectY * 32)) >= 0 Then 'DOWN
                                    .D = 0
                                ElseIf .D = 2 And (.XO - (.reflectX * 32)) <= 0 Then 'LEFT
                                    .D = 3
                                ElseIf .D = 3 And (.XO - (.reflectX * 32)) >= 0 Then 'RIGHT
                                    .D = 2
                                End If
                                
                                If b <> .D Then
                                    .Reflect = False
                                    .StartX = .reflectX
                                    .StartY = .reflectY
                                    .StartTime = GetTickCount
                                    .CanTargetSelf = True
                                End If
                            End If
                        'End If
                    Next y
                Next x
                    
                If .x >= 0 And .x <= 11 And .y >= 0 And .y <= 11 Then
                    If Not (A = 1 Or A = 2 Or A = 3 Or A = 21) Then
                        If .Key <> Character.CurProjectile.Key Then
                        
                        
                            If .Owner <> 255 Then
                                For A = 0 To 9
                                    If map.Monster(A).Monster > 0 Then
                                        If Round(map.Monster(A).XO / 32) = .x And Round(map.Monster(A).YO / 32) = .y Then
                                            D = 1
                                            If .Sprite < 256 Then CreateMonsterEffect A, .Sprite, 75, 8, 0
                                            Exit For
                                        End If
                                        If Monster(map.Monster(A).Monster).Flags2 And MONSTER_LARGE Then
                                            If Round(map.Monster(A).XO / 32) = .x - 1 And Round(map.Monster(A).YO / 32) = .y Then
                                                D = 1
                                                If .Sprite < 256 Then CreateMonsterEffect A, .Sprite, 75, 8, 0, 32, 0
                                                Exit For
                                            End If
                                            If Round(map.Monster(A).XO / 32) = .x - 1 And Round(map.Monster(A).YO / 32) = .y - 1 Then
                                                D = 1
                                                If .Sprite < 256 Then CreateMonsterEffect A, .Sprite, 75, 8, 0, 32, 32
                                                Exit For
                                            End If
                                            If Round(map.Monster(A).XO / 32) = .x And Round(map.Monster(A).YO / 32) = .y - 1 Then
                                                D = 1
                                                If .Sprite < 256 Then CreateMonsterEffect A, .Sprite, 75, 8, 0, 0, 32
                                                Exit For
                                            End If
                                        End If
                                    End If
                                Next A
                            Else
                                A = 10
                            End If
                            If A = 10 Then
                                For A = 1 To MAXUSERS
                                    If A <> Character.Index Then
                                        If player(A).map = CMap Then
                                            If A <> .Owner Or .CanTargetSelf Then
                                                If player(A).x = .x And player(A).y = .y Then
                                                    D = 1
                                                    If .Sprite < 256 Then CreatePlayerEffect A, .Sprite, 75, 8, 0
                          
                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    End If
                                Next A
                            End If
                            If A = MAXUSERS + 1 Then
                                If cX = .x And cY = .y Then
                                    If .Sprite < 256 Then CreateCharacterEffect .Sprite, 75, 8, 0
                                    If .Owner = 255 Then SendSocket Chr$(87) + Chr$(.Intensity)
                                    D = 1
                                End If
                            End If

                        End If
                    Else
                        D = 1
                        If .Sprite < 256 Then CreateTileEffectXYO .XO, .YO, .Sprite, 75, 8, 0
                    End If
                End If
            End If
            If D > 0 Then
                If .Key = Character.CurProjectile.Key Then
                    SendSocket Chr$(23) + Chr$(TT_NO_TARGET)
                    Character.CurProjectile.Key = vbNullString
                End If
                If D = 2 Then
                    If .Sprite < 256 Then
                        If .D = 3 Then
                             CreateTileEffectXYO .XO, .YO, .Sprite, 75, 8, 0
                        ElseIf .D = 2 Then
                            CreateTileEffectXYO .XO, .YO, .Sprite, 75, 8, 0
                        ElseIf .D = 1 Then
                            CreateTileEffectXYO .XO, .YO + 16, .Sprite, 75, 8, 0
                        ElseIf .D = 0 Then
                            CreateTileEffectXYO .XO, .YO + 16, .Sprite, 75, 8, 0
                        End If
                     
                     'CreateTileEffect .X, .Y, .Sprite, 75, 8, 0
                    End If
                End If
                If .Intensity Then
                    .Intensity = 0
                    ReDoLightSources
                End If
                Projectiles.Remove (.Key)
                D = 0
            End If
        End With
        
        'Now Check for collisions with your own projectile
        If Character.CurProjectile.Key <> "" Then
            With Projectiles(Character.CurProjectile.Key)
                If .x >= 0 And .x <= 11 And .y >= 0 And .y <= 11 Then
                    A = map.Tile(.x, .y).Att
                    
                    If Not (A = 1 Or A = 2 Or A = 3 Or A = 21 Or A = 4) Then
                        For A = 0 To 9
                            If map.Monster(A).Monster > 0 Then
                                
                                'how easy it is to hit a monster
                                'currently onidirectional checks
                                If Abs(map.Monster(A).XO - .XO) <= 18 And Abs(map.Monster(A).YO - .YO) <= 18 Then
                                    TargetMonster = A
                                    SendSocket Chr$(23) + Chr$(TT_MONSTER) + Chr$(A)
                                    D = 1
                                    If .Sprite < 256 Then CreateMonsterEffect A, .Sprite, 75, 8, 0
                                    Exit For
                                End If
                                
                                If Monster(map.Monster(A).Monster).Flags2 And MONSTER_LARGE Then
                                    If Abs(map.Monster(A).XO + 32 - .XO) <= 16 And Abs(map.Monster(A).YO - .YO) <= 16 Then
                                        TargetMonster = A
                                        SendSocket Chr$(23) + Chr$(TT_MONSTER) + Chr$(A)
                                        D = 1
                                        If .Sprite < 256 Then CreateMonsterEffect A, .Sprite, 75, 8, 0, 32, 0
                                        Exit For
                                    End If
                                    If Abs(map.Monster(A).XO + 32 - .XO) <= 16 And Abs(map.Monster(A).YO + 32 - .YO) <= 16 Then
                                        TargetMonster = A
                                        SendSocket Chr$(23) + Chr$(TT_MONSTER) + Chr$(A)
                                        D = 1
                                        If .Sprite < 256 Then CreateMonsterEffect A, .Sprite, 75, 8, 0, 32, 32
                                        Exit For
                                    End If
                                    If Abs(map.Monster(A).XO - .XO) <= 16 And Abs(map.Monster(A).YO + 32 - .YO) <= 16 Then
                                        TargetMonster = A
                                        SendSocket Chr$(23) + Chr$(TT_MONSTER) + Chr$(A)
                                        D = 1
                                        If .Sprite < 256 Then CreateMonsterEffect A, .Sprite, 75, 8, 0, 0, 32
                                        Exit For
                                    End If
                                End If
                                
                                
                            End If
                        Next A
                        If A = 10 Then
                            For A = 1 To MAXUSERS
                                If A <> Character.Index Then
                                    If player(A).map = CMap Then
                                        If player(A).x = .x And player(A).y = .y Then
                                            SendSocket Chr$(23) + Chr$(TT_PLAYER) + Chr$(A)
                                            D = 1
                                            If .Sprite < 256 Then CreatePlayerEffect A, .Sprite, 75, 8, 0
                                            Exit For
                                        End If
                                    End If
                                ElseIf .CanTargetSelf Then
                                    If cX = .x And cY = .y Then
                                        SendSocket Chr$(23) + Chr$(TT_PLAYER) + Chr$(Character.Index)
                                        D = 1
                                        If .Sprite < 256 Then CreateCharacterEffect .Sprite, 75, 8, 0
                                        Exit For
                                    End If
                                End If
                            Next A
                        End If
                    Else
                        SendSocket Chr$(23) + Chr$(TT_TILE) + Chr$(.x) + Chr$(.y)
                        D = 1
                        If .Sprite < 256 Then CreateTileEffect .x, .y, .Sprite, 75, 8, 0
                    End If
                Else
                    D = 1
                    SendSocket Chr$(23) + Chr$(TT_NO_TARGET)
                End If
                If D > 0 Then
                    Character.CurProjectile.Key = vbNullString
                    If D = 2 Then If .Sprite < 256 Then CreateTileEffect .x, .y, .Sprite, 75, 8, 0
                    If .Intensity Then
                        .Intensity = 0
                        ReDoLightSources
                    End If
                    Projectiles.Remove (.Key)
                End If
                
            End With
        End If
    Next
    Dim tmpPS As clsParticleSource
    For Each tmpPS In ParticleEngineB
        With tmpPS
            If tmpPS.Update = False Then
                ParticleEngineB.Remove tmpPS.Key
            End If
        End With
    Next
    For Each tmpPS In ParticleEngineF
        With tmpPS
            If tmpPS.Update = False Then
                ParticleEngineF.Remove tmpPS.Key
            End If
        End With
    Next
End Sub

Public Function Exists(Filename As String) As Boolean
On Error Resume Next
Open Filename For Input As #1
Close #1
If Err.Number <> 0 Then
   Exists = False
Else
   Exists = True
End If
End Function

Sub CheckFile(Filename As String)
    If Exists(Filename) = False Then
        MsgBox "Error: File " + Chr$(34) + Filename + Chr$(34) + " not found!", vbOKOnly + vbExclamation, TitleString
        End
    End If
End Sub

Sub CloseClientSocket(Action As Byte, NotifyServer As Boolean)
    Dim A As Long
    Character.Party = 0
    For A = 1 To MAXUSERS
        player(A).Party = 0
    Next A
    DrawPartyNames
    If NotifyServer Then
        SendSocket Chr$(0)
        A = GetTickCount
        While A + 100 > GetTickCount
        Wend
    End If
    closesocket ClientSocket
    ClientSocket = INVALID_SOCKET
    If frmMain_Loaded = True Then frmMain.Visible = False 'Unload frmMain
    If frmMapEdit_Loaded = True Then Unload frmMapEdit
    If frmMonster_Loaded = True Then Unload frmMonster
    If frmObject_Loaded = True Then Unload frmObject
    If frmList_Loaded = True Then Unload frmList
    If frmLight_Loaded = True Then Unload frmLight
    If frmGuilds_Loaded = True Then Unload frmGuilds
    If frmHallAccess_Loaded = True Then Unload frmHallAccess
    If frmBan_Loaded = True Then Unload frmBan
    If frmHall_Loaded = True Then Unload frmHall
    If frmOptions_Loaded = True Then Unload frmOptions
    If frmNewGuild_Loaded = True Then Unload frmNewGuild

    blnPlaying = False
    Select Case Action
        Case Else
            frmMenu.SetMenu 1 'MENU_MENU
            frmMenu.Show
    End Select
End Sub
Sub DeInitialize()

    Sound_Unload
    
    If ClientSocket <> INVALID_SOCKET Then
        closesocket ClientSocket
    End If
    
    'Unload Graphics
    Set sfcSymbols(1).Surface = Nothing
    Set sfcSymbols(2).Surface = Nothing
    Set sfcSymbols(3).Surface = Nothing
    Set sfcObjects(1).Surface = Nothing
    Set sfcObjects(2).Surface = Nothing
    Set sfcObjects(3).Surface = Nothing
    Set sfcObjects(4).Surface = Nothing
    Set sfcObjects(5).Surface = Nothing
    Set sfcObjects(6).Surface = Nothing
    Set sfcObjects(7).Surface = Nothing
    Set sfcObjects(8).Surface = Nothing
    Set sfcSprites.Surface = Nothing
    Set sfcInventory(0).Surface = Nothing
    Set sfcInventory(1).Surface = Nothing
    Set sfcInventory2.Surface = Nothing
    Set sfcChatTabs.Surface = Nothing
    
    Dim A As Long
    For A = 1 To NumTextures
        Set DynamicTextures(A).Texture = Nothing
    Next A
    For A = 1 To 10
        Set texParticles(A) = Nothing
    Next A
    Set TexFont.Texture = Nothing
    Set texAtts.Texture = Nothing
    Set texShade.Texture = Nothing
    Set texShadeEX.Texture = Nothing
    Set texLightsEX.Texture = Nothing
    
    Set D3D = Nothing
    Set DX8 = Nothing
    Set D3DX = Nothing
    Set D3DDevice = Nothing
    Set SwapChain(1) = Nothing
    Set mapTexture(1) = Nothing
    Set mapTexture(2) = Nothing
    Set RenderSurface(0) = Nothing
    Set RenderSurface(1) = Nothing
    
    UnloadMiniMap
    
    Set DD = Nothing
    Set Dx7 = Nothing
    

    
    'Unload Winsock
    EndWinsock
    
    'Unhook Form
    Unhook
    Unload frmMain
    
    Projectiles.Clear
    Effects.Clear
    Widgets.Clear
    FloatingText.Clear
    ParticleEngineB.Clear
    ParticleEngineF.Clear
    
    Set Projectiles = Nothing
    Set Effects = Nothing
    Set Widgets = Nothing
    Set FloatingText = Nothing
    Set ParticleEngineB = Nothing
    Set ParticleEngineF = Nothing
    End
End Sub
Public Sub Hook()
#If DEBUGWINDOWPROC Then
    On Error Resume Next
    Set m_SCHook = CreateWindowProcHook
    If Err Then
        MsgBox Err.Description
        Err.Clear
        Unhook
        Exit Sub
    End If
    On Error GoTo 0
    With m_SCHook
        .SetMainProc AddressOf WindowProc
        lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, .ProcAddress)
        .SetDebugProc lpPrevWndProc
    End With
#Else
    lpPrevWndProc = SetWindowLong(gHW, GWL_WNDPROC, AddressOf WindowProc)
#End If
End Sub

Public Sub Unhook()
    SetWindowLong gHW, GWL_WNDPROC, lpPrevWndProc
End Sub
Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = 1025 Then
        'Client Socket
        Select Case lParam And 255
            Case FD_CLOSE
                CloseClientSocket 1, False
            Case FD_CONNECT
                If lParam = FD_CONNECT Then
                    CPacketsSent = 0
                    SendSocket Chr$(61) + Chr$(ClientVer) + QuadChar(GetTickCount)
                    If NewAccount = True Then
                        frmMenu.SetStatusText "Sending Account Information . . .", 1
                        SendSocket Chr$(0) + User + Chr$(0) + SHA256(UCase(User) & UCase(Pass)) + Chr$(0) + Email
                    Else
                        frmMenu.SetStatusText "Sending Login Information . . .", 1
                        SendSocket Chr$(92) '+ ReadStr("Options", "UID")
                        SendSocket Chr$(93) + ReadStr("Options", "Keys")
                        SendSocket Chr$(1) + User + Chr$(0) + SHA256(UCase(User) & UCase(Pass))
                    End If
                Else
                    CloseClientSocket 4, False
                    frmMenu.SetStatusText "Error Connecting.", 2
                End If
            Case FD_READ
                If lParam = FD_READ Then ReceiveData
        End Select
    ElseIf uMsg = &H20A And hw = frmMain.hwnd Then 'Mouse Scroll
        If (wParam / 65536) > 0 Then
            frmMain.form_keydown 33, 0 'pgup
        Else
            frmMain.form_keydown 34, 0 'pgdn
        End If
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function

Function Crypt(St As String) As String
Dim A As Long, St1 As String
For A = 1 To Len(St)
    St1 = St1 & Chr$(219 Xor Asc(Mid$(St, A, 1)))
Next A
Crypt = St1
End Function

Function CheckSum(St As String) As Long
    Dim A As Long, b As Long
    For A = 1 To Len(St)
        b = b + Asc(Mid$(St, A, 1))
    Next A
    CheckSum = b
End Function

Function FindPlayer(ByVal St As String) As Long
    Dim A As Long, StLen As Long
    St = UCase$(St)
    
    'Search for exact match
    For A = 1 To MAXUSERS
        With player(A)
            If .Sprite > 0 And UCase$(.Name) = St Then
                FindPlayer = A
                Exit Function
            End If
        End With
    Next A
    
    'Search for partial match
    StLen = Len(St)
    For A = 1 To MAXUSERS
        With player(A)
            If .Sprite > 0 Then
                If Len(.Name) >= StLen Then
                    If UCase$(Left$(.Name, StLen)) = St Then
                        FindPlayer = A
                        Exit Function
                    End If
                End If
            End If
        End With
    Next A
End Function

Function FindGuild(ByVal St As String) As Long
    Dim A As Long, StLen As Long
    St = UCase$(St)
    
    'Search for exact match
    For A = 1 To 255
        With Guild(A)
            If UCase$(.Name) = St Then
                FindGuild = A
                Exit Function
            End If
        End With
    Next A
    
    'Search for partial match
    StLen = Len(St)
    For A = 1 To 255
        With Guild(A)
            If Len(.Name) >= StLen Then
                If UCase$(Left$(.Name, StLen)) = St Then
                    FindGuild = A
                    Exit Function
                End If
            End If
        End With
    Next
End Function
Function GetInt(Chars As String) As Long
    GetInt = Asc(Mid$(Chars, 1, 1)) * 256& + Asc(Mid$(Chars, 2, 1))
End Function

Function GetLong(Chars As String) As Long
    GetLong = Asc(Mid$(Chars, 1, 1)) * 16777216 + Asc(Mid$(Chars, 2, 1)) * 65536 + Asc(Mid$(Chars, 3, 1)) * 256& + Asc(Mid$(Chars, 4, 1))
End Function

Sub SendSocket(ByVal St As String)
Dim A As Long, b As Long, C As Byte
If CPacketsSent = 255 Then CPacketsSent = 0
CPacketsSent = (CPacketsSent + 1)
For A = 1 To Len(St)
    b = b + Asc(Mid$(St, A, 1)) + 7
Next A
    b = b Mod 256
    C = b Xor CPacketsSent
    C = Not C
    If SendData(ClientSocket, DoubleChar(Len(St) + 1) + St + Chr$(C)) = SOCKET_ERROR Then
        CloseClientSocket 0, False
    End If
End Sub
Function DoubleChar(Num As Integer) As String
    DoubleChar = Chr$(Int(Num / 256)) + Chr$(Num Mod 256)
End Function
Sub UploadMap()
    Dim MapData As String, St1 As String * 30
    Dim x As Long, y As Long
    With EditMap
        If .Version < 2147483647 Then
            .Version = .Version + 1
        Else
            .Version = 1
        End If
        St1 = .Name
        MapData = St1 + QuadChar(.Version) + Chr$(0) + Chr$(.MIDI) + _
        DoubleChar$(CLng(.ExitUp)) + DoubleChar$(CLng(.ExitDown)) + _
        DoubleChar$(CLng(.ExitLeft)) + DoubleChar$(CLng(.ExitRight)) + _
        DoubleChar(CLng(.BootLocation.map)) + Chr$(.BootLocation.x) + _
        Chr$(.BootLocation.y) + Chr$(.Flags(0)) + Chr$(.Intensity) + _
        Chr$(.Flags(1)) + DoubleChar$(.Raining) + DoubleChar$(.Snowing) + Chr$(.Zone) + _
        Chr$(.Fog) + Chr$(.SnowColor) + Chr$(.RainCOlor) + " " + _
        Chr$(.MonsterSpawn(0).Monster Mod 256) + Chr$(.MonsterSpawn(0).Rate) + _
        Chr$(.MonsterSpawn(1).Monster Mod 256) + Chr$(.MonsterSpawn(1).Rate) + _
        Chr$(.MonsterSpawn(2).Monster Mod 256) + Chr$(.MonsterSpawn(2).Rate) + _
        Chr$(.MonsterSpawn(3).Monster Mod 256) + Chr$(.MonsterSpawn(3).Rate) + _
        Chr$(.MonsterSpawn(4).Monster Mod 256) + Chr$(.MonsterSpawn(4).Rate)
        For y = 0 To 11
            For x = 0 To 11
                With .Tile(x, y)
                    If .Att = 24 Then
                        MapData = MapData + DoubleChar(CLng(.Ground)) + DoubleChar(CLng(.Ground2)) + DoubleChar(CLng(.BGTile1)) + Chr$(.Anim(1)) + Chr$(.Anim(2)) + DoubleChar(CLng(.FGTile)) + Chr$(.Att) + Chr$(.AttData(0)) + Chr$(.AttData(1)) + Chr$(.AttData(2)) + Chr$(.AttData(1)) + Chr$(.WallTile)
                    Else
                        MapData = MapData + DoubleChar(CLng(.Ground)) + DoubleChar(CLng(.Ground2)) + DoubleChar(CLng(.BGTile1)) + Chr$(.Anim(1)) + Chr$(.Anim(2)) + DoubleChar(CLng(.FGTile)) + Chr$(.Att) + Chr$(.AttData(0)) + Chr$(.AttData(1)) + Chr$(.AttData(2)) + Chr$(.AttData(3)) + Chr$(.WallTile)
                    End If
                    
                End With
            Next x
        Next y
        MapData = MapData + Chr$(Int(.MonsterSpawn(0).Monster / 256))
        MapData = MapData + Chr$(Int(.MonsterSpawn(1).Monster / 256))
        MapData = MapData + Chr$(Int(.MonsterSpawn(2).Monster / 256))
        MapData = MapData + Chr$(Int(.MonsterSpawn(3).Monster / 256))
        MapData = MapData + Chr$(Int(.MonsterSpawn(4).Monster / 256))
    End With
    SendSocket Chr$(12) + MapData
End Sub

Sub PrintChat(ByVal St As String, Color As Long, ByVal Size As Long, Optional ByVal Channel As Byte = 15, Optional ByVal getcolor As Boolean = False)
With Chat

        
If Not getcolor Then Color = QBColor(Color)
If Channel = 1 Then Channel = 0

        .ChatIndex = .ChatIndex + 1
        If .ChatIndex > 1000 Then .ChatIndex = 0
        .AllChat(.ChatIndex).Text = St
        .AllChat(.ChatIndex).Color = Color
        .AllChat(.ChatIndex).Size = Size
        .AllChat(.ChatIndex).used = True
        .AllChat(.ChatIndex).Channel = Channel

        DrawChat
 '   End If
'End If
End With

End Sub

Sub DrawChatLine(ByVal St As String, ByRef lineRow As Long, ByRef numRows As Long, ByRef textHeight As Byte)
    Dim C As Long, D As Long, yOff As Long, yOff2
                For C = 1 To Len(St)
                    If frmMain.picChat.TextWidth(Mid$(St, 1, C)) >= (frmMain.picChat.ScaleWidth - frmMain.picChat.CurrentX) Then
                      D = C
                      While (D > 1)
                        If Mid$(St, D, 1) = " " Then
                            C = D + 1
                            D = 1
                        End If
                        D = D - 1
                      Wend
                      DrawChatLine Mid$(St, C), lineRow, numRows, textHeight
                      St = Mid$(St, 1, C - 1)
                      Exit For
                    End If
                Next C

                TextOut frmMain.picChat.hdc, 0, frmMain.picChat.CurrentY + ((numRows - 1) * textHeight) - 1, St, Len(St)
                numRows = numRows - 1
End Sub

Sub DrawChat()
    Dim A As Long, b As Long, textHeight As Byte, curLine As Long, C As Long
    frmMain.picChat.FontSize = Options.FontSize
    textHeight = frmMain.picChat.textHeight("A")
    frmMain.picChat.CurrentY = (frmMain.picChat.ScaleHeight - textHeight) * WindowScaleY
    
    For A = 1 To 5
        If Not chatForms(A) Is Nothing Then
            If chatForms(A).Visible Then
                chatForms(A).DrawChat2
            End If
        End If
    Next A

    frmMain.picChat.Cls
    
    With frmMain.picChat
        b = frmMain.picChat.ScaleHeight / textHeight '12 * WindowScaleY
        frmMain.picChat.CurrentY = frmMain.picChat.ScaleHeight - textHeight * b
        A = Chat.ChatIndex - ChatScrollBack
        While (A > (Chat.ChatIndex - ChatScrollBack - 15 * WindowScaleY - C) And b <> 0)
            If A < 0 Then
                curLine = A + 1000
            Else
                curLine = IIf(A > 1000, A - 1000, A)
            End If
            If Not Chat.AllChat(curLine).used Then
                A = Chat.ChatIndex - ChatScrollBack - 15
                b = 0
            Else
                If Chat.AllChat(curLine).Channel = 0 And Not Options.Broadcasts Then
                    C = C + 1
                    A = A - 1
                ElseIf Chat.AllChat(curLine).Channel = 1 And Not Options.Says Then
                    C = C + 1
                    A = A - 1
                ElseIf Chat.AllChat(curLine).Channel = 2 And Not Options.Tells Then
                    C = C + 1
                    A = A - 1
                ElseIf Chat.AllChat(curLine).Channel = 3 And Not Options.Emotes Then
                    C = C + 1
                    A = A - 1
                ElseIf Chat.AllChat(curLine).Channel = 4 And Not Options.Yells Then
                    C = C + 1
                    A = A - 1
                Else
                    .ForeColor = Chat.AllChat(curLine).Color
                    DrawChatLine Chat.AllChat(curLine).Text, A, b, textHeight
                    A = A - 1
                End If
            End If
        Wend
    End With
End Sub

Sub AddLog(Message As String)
frmLog.txtLog.Text = frmLog.txtLog.Text + vbCrLf + Trim$(Message)
End Sub

Sub SendAwayMsg(player As Long, Message As String)
If Message <> "" Then
    If player > 0 Then
        SendSocket Chr$(14) + Chr$(player) + Message
    End If
End If
End Sub
Sub ConnectClient()
    Dim St As String
    Character.StatusEffect = 0
    Character.buff = 0
    ClientSocket = ConnectSock(ServerIP, ServerPort, St, gHW, True)
End Sub

Sub CreateStatusColors()
Dim A As Long

StatusColors(0) = &HFFC0C0C0
StatusColors(1) = &HFFC0C0C0
StatusColors(2) = &HFFFF0000
StatusColors(3) = &HFF0000FF
StatusColors(4) = &HFFC0C0FF
StatusColors(5) = &HFF004040
StatusColors(6) = &HFF000040
StatusColors(7) = &HFF8080FF
StatusColors(8) = &HFF8080FF
StatusColors(9) = &HFF808080
StatusColors(10) = &HFF000000
StatusColors(11) = &HFF004000
StatusColors(12) = &HFF008080
StatusColors(13) = &HFFC0C0FF
StatusColors(14) = &HFF804040
StatusColors(15) = &HFFFF8000
StatusColors(16) = &HFFC0FFC0
StatusColors(17) = &HFFC00000
StatusColors(18) = &HFFFF4040
StatusColors(19) = &HFF800080
StatusColors(20) = &HFFFFFFC0
'21 is blinky
StatusColors(22) = &HFFFF0000
StatusColors(23) = &HFF00FF00
StatusColors(24) = &HFFFF00FF
'25 is white/black alternating
StatusColors(26) = &HFFFFFFFF

For A = 27 To 100
    StatusColors(A) = QBColor(7)
    
Next A

    NPCStatusColors(1) = &HFFFFFFBB 'Normal NPC Color
    NPCStatusColors(2) = &HFFFFFF00 'Normal NPC Quest Color
    NPCStatusColors(3) = &HFF00FF00 'BRIGHT GREEN
    
    
End Sub
Sub CreateTileEffect(x As Long, y As Long, Sprite As Long, Speed As Long, TotalFrames As Long, loopcount As Integer)
If x < 0 Or x > 11 Or y < 0 Or y > 11 Then Exit Sub
Effects.Add x * 32, y * 32, Sprite, Speed, TotalFrames, loopcount, 0, TT_TILE
End Sub
Sub CreateTileEffectXYO(x As Long, y As Long, Sprite As Long, Speed As Long, TotalFrames As Long, loopcount As Integer)
If x < -16 Or x > 384 Or y < -16 Or y > 384 Then Exit Sub
Effects.Add x, y, Sprite, Speed, TotalFrames, loopcount, 0, TT_TILE
End Sub
Sub CreateMonsterEffect(MonsterNum As Long, Sprite As Long, Speed As Long, TotalFrames As Long, loopcount As Integer, Optional xOff As Byte = 0, Optional yOff As Byte = 0, Optional deathCont As Boolean = False)
    With map.Monster(MonsterNum)
        If .Monster > 0 Then
            Effects.Add .XO, .YO, Sprite, Speed, TotalFrames, loopcount, MonsterNum, TT_MONSTER, xOff, yOff, deathCont
        End If
    End With
End Sub
Sub CreatePlayerEffect(Index As Long, Sprite As Long, Speed As Long, TotalFrames As Long, loopcount As Integer)
    With player(Index)
        If .map = CMap Then
            Effects.Add .XO, .YO, Sprite, Speed, TotalFrames, loopcount, Index, TT_PLAYER
        End If
    End With
End Sub

Sub CreateCharacterEffect(Sprite As Long, Speed As Long, TotalFrames As Long, loopcount As Integer)
    Effects.Add Cxo, CYO, Sprite, Speed, TotalFrames, loopcount, 0, TT_CHARACTER
End Sub

Sub SetParty(Index As Integer, Party As Integer)
    If Index = Character.Index Then
        Character.Party = Party
    Else
        player(Index).Party = Party
    End If
    DrawPartyNames
End Sub


Sub TransparentBlt(hdc As Long, ByVal DestX As Long, ByVal DestY As Long, destWidth As Long, destHeight As Long, srcDC As Long, srcX As Long, srcY As Long, maskDC As Long)
    BitBlt hdc, DestX, DestY, destWidth, destHeight, maskDC, srcX, srcY, SRCAND
    BitBlt hdc, DestX, DestY, destWidth, destHeight, srcDC, srcX, srcY, SRCPAINT
    
End Sub
Function CanAttack(x As Long, y As Long) As Boolean
    If map.Tile(cX, cY).Att = 21 Then
        CanAttack = False
    Else
        If Abs(x - cX) > 1 Or Abs(y - cY) > 1 Then
            CanAttack = False
        Else
            If Abs(x - cX) = 1 And Abs(y - cY) = 1 Then
                CanAttack = False
            Else
                If y - cY = -1 Then 'up
                    If ExamineBit(map.Tile(x, y).WallTile, 1) Or ExamineBit(map.Tile(cX, cY).WallTile, 4) Then
                        CanAttack = False
                    Else
                        CanAttack = True
                    End If
                End If
                If y - cY = 1 Then 'down
                    If ExamineBit(map.Tile(x, y).WallTile, 0) Or ExamineBit(map.Tile(cX, cY).WallTile, 5) Then
                        CanAttack = False
                    Else
                        CanAttack = True
                    End If
                End If
                If x - cX = -1 Then 'left
                    If ExamineBit(map.Tile(x, y).WallTile, 3) Or ExamineBit(map.Tile(cX, cY).WallTile, 6) Then
                        CanAttack = False
                    Else
                        CanAttack = True
                    End If
                End If
                If x - cX = 1 Then 'right
                    If ExamineBit(map.Tile(x, y).WallTile, 2) Or ExamineBit(map.Tile(cX, cY).WallTile, 7) Then
                        CanAttack = False
                    Else
                        CanAttack = True
                    End If
                End If
            End If
        End If
    End If
End Function
Sub ReDoLightSources()
    Dim A As Long, b As Long
    Dim x As Byte, y As Byte
    b = 1 'own light source is always 0
    For A = 1 To MAXUSERS
        If player(A).Light.Intensity <> 0 And b <= 29 And A <> Character.Index And player(A).map = CMap Then
            LightSource(b).Intensity = player(A).Light.Intensity
            LightSource(b).Radius = player(A).Light.Radius
            LightSource(b).x = player(A).XO + 16
            LightSource(b).y = player(A).YO + 16
            LightSource(b).Type = LT_PLAYER
            LightSource(b).Flicker = 0
            player(A).LightSourceNumber = b
            b = b + 1
        Else
            player(A).LightSourceNumber = 0
        End If
    Next A
    For x = 0 To 11
        For y = 0 To 11
            If map.Tile(x, y).Att = 22 And b <= 29 Then
                LightSource(b).Intensity = map.Tile(x, y).AttData(0)
                LightSource(b).Radius = map.Tile(x, y).AttData(1)
                LightSource(b).MaxFlicker = (map.Tile(x, y).AttData(2) And 31)
                LightSource(b).FlickerRate = 4
                LightSource(b).x = x * 32 + 16
                LightSource(b).y = y * 32 + 16
                If map.Tile(x, y).AttData(3) > 0 And map.Tile(x, y).AttData(3) <= 50 Then
                    LightSource(b).Red = Lights(map.Tile(x, y).AttData(3)).Red
                    LightSource(b).Green = Lights(map.Tile(x, y).AttData(3)).Green
                    LightSource(b).Blue = Lights(map.Tile(x, y).AttData(3)).Blue
                Else
                    LightSource(b).Red = 0
                    LightSource(b).Green = 0
                    LightSource(b).Blue = 0
                End If
                LightSource(b).Type = LT_TILE
                b = b + 1
            End If
        Next y
    Next x
    For A = 0 To 49
        If map.Object(A).Object > 0 Then
            If map.Object(A).Light.Intensity <> 0 And b <= 29 Then
                LightSource(b).Intensity = map.Object(A).Light.Intensity
                LightSource(b).Radius = map.Object(A).Light.Radius
                LightSource(b).x = map.Object(A).x * 32 + 16 + map.Object(A).XOffset
                LightSource(b).y = map.Object(A).y * 32 + 16 + map.Object(A).YOffset
                LightSource(b).Type = LT_OBJECT
                LightSource(b).Flicker = 0
                b = b + 1
            End If
        End If
    Next A
    For A = 0 To 9
        If map.Monster(A).Monster > 0 Then
            If Monster(map.Monster(A).Monster).Light > 0 Then
                With Lights(Monster(map.Monster(A).Monster).Light)
                    LightSource(b).Intensity = .Intensity
                    LightSource(b).Radius = .Radius
                    LightSource(b).x = map.Monster(A).XO + 16
                    LightSource(b).y = map.Monster(A).YO + 16
                    LightSource(b).Red = .Red
                    LightSource(b).Green = .Green
                    LightSource(b).Blue = .Blue
                    LightSource(b).Flicker = 0
                    LightSource(b).Type = LT_MONSTER
                    map.Monster(A).LightSourceNumber = b
                    b = b + 1
                End With
            Else
                map.Monster(A).LightSourceNumber = 0
            End If
        Else
            map.Monster(A).LightSourceNumber = 0
        End If
    Next A
    Dim tmpProjectile As clsProjectile
    For Each tmpProjectile In Projectiles
        With tmpProjectile
            If .Radius > 0 Then
                LightSource(b).Intensity = .Intensity
                LightSource(b).Radius = .Radius
                LightSource(b).x = .XO + 16
                LightSource(b).y = .YO
                LightSource(b).Red = .Red
                LightSource(b).Green = .Green
                LightSource(b).Blue = .Blue
                LightSource(b).Type = LT_PROJECTILE
                LightSource(b).D = .D
                .LightSourceNumber = b
                b = b + 1
            End If
        End With
    Next
    
    For A = b To 29
        LightSource(A).Intensity = 0
        LightSource(A).Radius = 0
        LightSource(A).x = 0
        LightSource(A).y = 0
        LightSource(A).Type = LT_NONE
        LightSource(A).Flicker = 0
    Next A
    RedoOwnLight
End Sub
Sub RedoOwnLight()
    Character.Light.Intensity = 255
    Character.Light.Radius = 65
    LightSource(0).Intensity = Character.Light.Intensity
    LightSource(0).Radius = Character.Light.Radius
    LightSource(0).x = Cxo + 16
    LightSource(0).y = CYO
    LightSource(0).Type = 2
'    For A = 1 To 5
'        With Character.Equipped(A)
'            If .Object > 0 Then
'                If .Prefix > 0 Then
'                    If Prefix(.Prefix).Light.Intensity > 0 Then
'                        LightSource(0).Intensity = LightSource(0).Intensity + Prefix(.Prefix).Light.Intensity
'                        LightSource(0).Radius = LightSource(0).Radius + Prefix(.Prefix).Light.Radius
'                    End If
'                End If
'                If .Suffix > 0 Then
'                    If Prefix(.Suffix).Light.Intensity > 0 Then
'                        LightSource(0).Intensity = LightSource(0).Intensity + Prefix(.Suffix).Light.Intensity
'                        LightSource(0).Radius = LightSource(0).Radius + Prefix(.Suffix).Light.Radius
'                    End If
'                End If
'            End If
'        End With
'    Next A
End Sub

Sub CalculateAmbientAlpha()
Dim A As Integer
If (map.Flags(0) And MAP_INDOORS) Then
    A = 255 - map.Intensity
Else
    Select Case World.Hour
        Case 21
            A = 218 + (37.5 * World.Minute / 240)
        Case 22
            A = 180 + (37.5 * World.Minute / 240)
        Case 23
            A = 143 + (37.5 * World.Minute / 240)
        Case 24
            A = 105 + (37.5 * World.Minute / 240)
        Case 1, 2
            A = 105
        Case 3
            A = 105 + (37.5 - 37.5 * World.Minute / 240)
        Case 4
            A = 143 + (37.5 - 37.5 * World.Minute / 240)
        Case 5
            A = 180 + (37.5 - 37.5 * World.Minute / 240)
        Case 6
            A = 218 + (37.5 - 37.5 * World.Minute / 240)
        Case Else
            A = 255
    End Select
    A = A - map.Intensity
End If
If A < 0 Then A = 0
If A > 255 Then A = 255
AmbientAlpha = A
AmbientGreen = 0
AmbientBlue = 0
AmbientRed = 0
End Sub

Function RGB16(Red As Variant, Green As Variant, Blue As Variant) As Long
If Red > 31 Then Red = 31
If Green > 63 Then Green = 63
If Blue > 31 Then Blue = 31
RGB16 = Red * 2016 + Green * 32 + Blue
End Function

Public Sub WRITETOLOG(INFO As String)
Open AppPath & "packetlog.txt" For Append As #1
    Print #1, INFO
Close #1
End Sub

Function GetFreeDeadBody() As Byte
Dim A As Long, b As Long
b = 0
For A = 0 To 9
    If map.DeadBody(A).Sprite = 0 Then
        GetFreeDeadBody = A
        Exit Function
    Else
        If map.DeadBody(A).Counter < map.DeadBody(b).Counter Then
            b = A
        End If
    End If
Next A
GetFreeDeadBody = b
End Function

Public Sub SetMapUndo(x As Long, y As Long, Layer As Long, OldTile As Long, Optional AttData0 As Byte = 0, Optional AttData1 As Byte = 0, Optional AttData2 As Byte = 0, Optional AttData3 As Byte = 0)
Dim A As Long

With MapUndo(CurrentMapAction)
    If .x = x And .y = y And .Layer = Layer And .OldTile = OldTile And .AttData(0) = AttData0 And .AttData(1) = AttData1 And .AttData(2) = AttData2 And .AttData(3) = AttData3 Then Exit Sub
End With

If CurrentMapAction < 250 Then
    CurrentMapAction = CurrentMapAction + 1
Else
    For A = 0 To 249
        With MapUndo(A)
            .x = MapUndo(A + 1).x
            .y = MapUndo(A + 1).y
            .Layer = MapUndo(A + 1).Layer
            .OldTile = MapUndo(A + 1).OldTile
            .AttData(0) = MapUndo(A + 1).AttData(0)
            .AttData(1) = MapUndo(A + 1).AttData(1)
            .AttData(2) = MapUndo(A + 1).AttData(2)
            .AttData(3) = MapUndo(A + 1).AttData(3)
        End With
    Next A
End If

With MapUndo(CurrentMapAction)
    .x = x
    .y = y
    .Layer = Layer
    .OldTile = OldTile
    .AttData(0) = AttData0
    .AttData(1) = AttData1
    .AttData(2) = AttData2
    .AttData(3) = AttData3
End With
End Sub

Public Sub UndoMapEdit()
If MapEdit Then
    If CurrentMapAction > 0 Then
        
        With MapUndo(CurrentMapAction)
            Select Case .Layer
                Case 1  'Ground
                    EditMap.Tile(.x, .y).Ground = .OldTile
                Case 2  'Ground 2
                    EditMap.Tile(.x, .y).Ground2 = .OldTile
                Case 3  'BGTile1
                    EditMap.Tile(.x, .y).BGTile1 = .OldTile
                Case 4  'Anim
                    EditMap.Tile(.x, .y).Anim(1) = .OldTile
                Case 5  'FGTile
                    EditMap.Tile(.x, .y).FGTile = .OldTile
                Case 6  'Att
                    EditMap.Tile(.x, .y).Att = .OldTile
                    EditMap.Tile(.x, .y).AttData(0) = .AttData(0)
                    EditMap.Tile(.x, .y).AttData(1) = .AttData(1)
                    EditMap.Tile(.x, .y).AttData(2) = .AttData(2)
                    EditMap.Tile(.x, .y).AttData(3) = .AttData(3)
                Case 7
                    EditMap.Tile(.x, .y).WallTile = .OldTile
            End Select
        CurrentMapAction = CurrentMapAction - 1
        End With
    End If
End If
End Sub
Sub drawTopBar()
Dim R1 As RECT
Dim r2 As RECT
  r2.Top = 0
    r2.Left = 0
    r2.Bottom = 36
    r2.Right = 78
    
    R1.Top = 0
    R1.Left = 522
    R1.Right = R1.Left + 78
    R1.Bottom = 36
    
    R1.Top = R1.Top * WindowScaleY: R1.Bottom = R1.Bottom * WindowScaleY: R1.Left = R1.Left * WindowScaleX: R1.Right = R1.Right * WindowScaleX
    
    Draw sfcTimers, 0, 0, 78, 36, sfcInventory(0), 145, 393, True, 0
    If combatCounter > 0 Then
        
        Draw sfcTimers, 2, 2, 32, 32, sfcInventory(0), 223, 397, True, 0
        If hoverCombat Then Draw sfcTimers, 8, 11, 121, 17, sfcInventory(0), 243, 625, True, 0
        sfcTimers.Surface.BltToDC frmMain.hdc, r2, R1
        frmMain.FontSize = 9 * WindowScaleY
        If hoverCombat Then
            frmMain.ForeColor = vbWhite
            TextOut frmMain.hdc, 533 * WindowScaleX, 12 * WindowScaleY, "Combat: " & combatCounter, IIf(combatCounter > 9, 10, 9)
        End If
    End If
    
    frmMain.Refresh

End Sub

Sub resetAttack()
Dim R1 As RECT
Dim r2 As RECT
    drawCooldowns = True
    r2.Top = 0
    r2.Left = 0
    r2.Bottom = 32
    r2.Right = 64
    
    R1.Top = 2
    R1.Left = 212
    R1.Right = R1.Left + 64
    R1.Bottom = 34

    R1.Top = R1.Top * WindowScaleY: R1.Bottom = R1.Bottom * WindowScaleY: R1.Left = R1.Left * WindowScaleX: R1.Right = R1.Right * WindowScaleX

    Draw sfcTimers2, 0, 0, 32, 32, sfcInventory(0), 286 + 32, 292, True, 0
    sfcTimers2.Surface.BltToDC frmMain.hdc, r2, R1
    'frmMain.Refresh

End Sub
Sub resetMagic()
Dim R1 As RECT
Dim r2 As RECT

    r2.Top = 0
    r2.Left = 0
    r2.Bottom = 32
    r2.Right = 64
    
    R1.Top = 2
    R1.Left = 212
    R1.Right = R1.Left + 64
    R1.Bottom = 34
    R1.Top = R1.Top * WindowScaleY: R1.Bottom = R1.Bottom * WindowScaleY: R1.Left = R1.Left * WindowScaleX: R1.Right = R1.Right * WindowScaleX

    
    drawCooldowns = True
    Draw sfcTimers2, 32, 2, 28, 28, sfcInventory(0), 548, 315, True, 0
    sfcTimers2.Surface.BltToDC frmMain.hdc, r2, R1
    frmMain.Refresh

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

    
Sub YouDied()
    With Character
        'Reset Stat Bars
        .HP = .MaxHP
        .Energy = .MaxEnergy
        .Mana = .MaxMana
        DrawHP
        DrawEnergy
        DrawMana
    End With
    
    Freeze = True
    Transition 0, 255, 0, 0, 10
    NextTransition = 6
    PlayWav 8
End Sub


Sub Main()
'On Error GoTo ERRHANDLER
    Dim A As Long, AnimDelay As Long
    'Dim Rec_Pos As RECT
    'Speedhack protection Variables
    Dim LastTime As Long, NumTimeOff As Long, NumTimeOff2 As Long, SkipTimeOff As Long, TimeRatio As Long, speedhack As Byte
    SkipTimeOff = 10
    NumTimeOff2 = 7
    
    
   
    MapKey = Chr$(180) & Chr$(136) & Chr$(148) & Chr$(74) & Chr$(77) & Chr$(198) & Chr$(3) & Chr$(194) & Chr$(208) & Chr$(181) & Chr$(11) & Chr$(105) & Chr$(220) & Chr$(202) & Chr$(95) & Chr$(246) & Chr$(223) & Chr$(14) & Chr$(243) & Chr$(93) & Chr$(134) & Chr$(196) & Chr$(13) & Chr$(151) & Chr$(119) & Chr$(76) & Chr$(159) & Chr$(165) & Chr$(67) & Chr$(71) & Chr$(212) & Chr$(211) & Chr$(150) & Chr$(252) & Chr$(233) & Chr$(58) & Chr$(177) & Chr$(250) & Chr$(62) & Chr$(136) & Chr$(27) & Chr$(255) & Chr$(173) & Chr$(4) & Chr$(147) & Chr$(25) & Chr$(26) & Chr$(204) & Chr$(72) & Chr$(11) & Chr$(75) & Chr$(97) & Chr$(77) & Chr$(242) & Chr$(250) & Chr$(102) & Chr$(71) & Chr$(41) & Chr$(41) & Chr$(165) & Chr$(104) & Chr$(105) & Chr$(182) & Chr$(83) & Chr$(162) & Chr$(53) & Chr$(47) & Chr$(149) & Chr$(20) & Chr$(117) & Chr$(231) & Chr$(66) & Chr$(201) & Chr$(96) & Chr$(74) & Chr$(235) & Chr$(161) & Chr$(160) & Chr$(109) & Chr$(25) & Chr$(143) & Chr$(177) & Chr$(233) & Chr$(213) & Chr$(5) & _
             Chr$(139) & Chr$(234) & Chr$(110) & Chr(173) & Chr$(128) & Chr$(131) & Chr$(118) & Chr$(90) & Chr$(103) & Chr$(69) & Chr$(14) & Chr$(62) & Chr$(250) & Chr$(15) & Chr$(99) & Chr$(93) & Chr$(125) & Chr$(39) & Chr$(121) & Chr$(65) & Chr$(160) & Chr$(138) & Chr$(40) & Chr$(240) & Chr$(167) & Chr$(129) & Chr$(99) & Chr$(27) & Chr$(200) & Chr$(117) & Chr$(192) & Chr$(152) & Chr$(213) & Chr$(4) & Chr$(53) & Chr$(18) & Chr$(26) & Chr$(84) & Chr$(32) & Chr$(0) & Chr$(137) & Chr$(168) & Chr$(139) & Chr$(211) & Chr$(20) & Chr$(49) & Chr$(173) & Chr$(116) & Chr$(91) & Chr$(38) & Chr$(180) & Chr$(237) & Chr$(135) & Chr$(22) & Chr$(193) & Chr$(102) & Chr$(118) & Chr$(125) & Chr$(53) & Chr$(84) & Chr$(24) & Chr$(150) & Chr$(43) & Chr$(237) & Chr$(25) & Chr$(113) & Chr$(69) & Chr$(223) & Chr$(192) & Chr$(69) & Chr$(172) & Chr$(65) & Chr$(23) & Chr$(7) & Chr$(82) & Chr$(202) & Chr$(76) & Chr$(60) & Chr$(123) & Chr$(65) & Chr$(87) & Chr$(11) & Chr$(123) & Chr$(52) & Chr$(221) & Chr$(150) & _
             Chr$(193) & Chr$(237) & Chr$(84) & Chr$(138) & Chr$(20) & Chr$(162) & Chr$(245) & Chr$(29) & Chr$(236) & Chr$(158) & Chr$(89) & Chr$(38) & Chr$(122) & Chr$(56) & Chr$(254) & Chr$(33) & Chr$(7) & Chr$(88) & Chr$(140) & Chr$(236) & Chr$(137) & Chr$(104) & Chr$(216) & Chr$(211) & Chr$(172) & Chr$(184) & Chr$(255) & Chr$(86) & Chr$(126)
    
    Dim regstring As String
    regstring = ReadStr("Options", "Keys")
    If regstring = "" Then
        Randomize
        
        regstring = Chr$(2) + Chr$(120) + Chr$(45) + Chr$(96) + Chr$(4) + Chr$(2) + Chr$(120) + Chr$(3) + Chr$(17) + Chr$(245)
        For A = 1 To 20
            regstring = regstring + Chr$(CInt(Rnd * 254) + 1)
        Next A
        regstring = regstring + Chr$(17) + Chr$(245)
        For A = 1 To 18
            regstring = regstring + Chr$(CInt(Rnd * 254) + 1)
        Next A
        A = Len(regstring)
        'WriteString "Options" "Keys" regstring
    End If

    'CurrentProcessHandle = GetCurrentProcess

    ChDir (App.Path)

    If Not InStr(1, Command$, "-update") = 0 Then
        If Exists("Updater.exe") Then
            On Error GoTo ERRUPDATER
            Shell "Updater.exe", vbNormalFocus
            On Error Resume Next
            End
        End If
    End If

    
    If Exists("Updater.tmp") Then
        On Error Resume Next
        Do
            DoEvents
            Err.Clear
            FileCopy "Updater.tmp", "Updater.exe"
        Loop While Err.Number > 0
        Kill "Updater.tmp"
        On Error GoTo ERRHANDLER
    End If

    If GetTickCount < 0 Then
        MsgBox "There has been an error (gettickcount)", vbOKOnly, TitleString
        End
    End If
        
    SetServerData
    InitConstants
    CreateKeyCodeList
    LoadOptions

    CheckFiles
    Init_Directx
    If Not InitD3D(True) Then
       GoTo exitit
    End If

    CreateStatusColors

    CHannelColors(0) = QBColor(13)
    CHannelColors(1) = QBColor(13)
    CHannelColors(2) = RGB(122, 156, 211)
    CHannelColors(3) = RGB(155, 227, 199)
    CHannelColors(4) = RGB(255, 255, 255)
    CHannelColors(5) = RGB(255, 113, 115)
    CHannelColors(6) = RGB(71, 136, 68)

    MapCacheFile = "Data\Cache\cachem" + ServerId + ".dat"
    
    
    CreateModString
    CreateClassData
    

    
    LoadFontData
    LoadMacros
    
    
    'If Options.AltKeysEnabled Then
        Chat.Enabled = False
        DrawChatString
    'Else
    '    Chat.Enabled = True
    'End If
    
    If Options.hightask Then
        'SetPriorityClass CurrentProcessHandle, HIGH_PRIORITY_CLASS
    Else
        'SetPriorityClass CurrentProcessHandle, NORMAL_PRIORITY_CLASS
    End If
        
    ResizeGameWindow
    
    
    Sound_Init
    Sound_PlayStream 1
    
    Dim St As String


    
    If Exists(MapCacheFile) Then
        If FileLen(MapCacheFile) <> 11895000 Then
            CreateMapCache
        End If
    Else
        CreateMapCache
    End If

    If InStr(1, Command$, "-oneInstance") = 0 Then
        If App.PrevInstance = True Then
        '    MsgBox "You can only run one instance of Seyerdin at a time."
            'GoTo exitit
        End If
    End If
    

    Load frmMenu
    
    'Hook Form
    Hook
        
    'Load Winsock
    StartWinsock (St)
    frmMenu.Show False
    frmMenu.SetMenu 1

    InitSkills
    GenerateEXPLevels

    CurInvObj = 1
    'SetTab tsStats

    ReDoLightSources
    InitMiniMap
    SetMiniMapTab tsButtons
    InitRain 150
    InitSnow 150

    NumRainDrops = 0
    NumSnowFlakes = 0
    
    LastTime = Time
    AmbientRed = 255
    AmbientGreen = 255
50:     AmbientBlue = 255

           
    Draw sfcTimers2, 0, 0, 64, 32, sfcInventory(0), 148, 393, True, 0
    resetAttack
    resetMagic

    While blnEnd = False
        If BanString <> "" Then
            MsgBox BanString
            BanString = ""
        End If
        If blnPlaying = True Then
            If Freeze = False Then
                If AnimDelay = 15 Then
                    CurFrame = 1 - CurFrame
                    AnimDelay = 0
                Else
                    AnimDelay = AnimDelay + 1
                End If
                If AnimDelay = 7 Or AnimDelay = 15 Then FrameCounter = FrameCounter + 1
                If FrameCounter = 8 Then FrameCounter = 0
                
                
                If DD.TestCooperativeLevel <> DD_OK Then
                    A = DD.TestCooperativeLevel
                    'Do While DD.TestCooperativeLevel <> DD_OK
                        Select Case DD.TestCooperativeLevel
                            Case DDERR_WRONGMODE
                                Debug.Print "wrong mode "
                                
                            Case Else
                                Debug.Print "Other error: " & DD.TestCooperativeLevel
                        End Select
                    'Loop
                    Init_Directx
                    InitializeTextures
                    BOOLD3DERROR = True
                    STOPMAPCHECK = True
                    InitMiniMap
                    If CurrentTab = tsMap Then
                        SetMiniMapTab tsMap
                    End If
                    Draw sfcTimers2, 0, 0, 64, 32, sfcInventory(0), 148, 393, True, 0
                End If
                If D3DDevice.TestCooperativeLevel <> D3D_OK Then
                    Do Until D3DDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET
                        'Sleep 1
                        DoEvents
                    Loop
        
                    For A = 1 To NumTextures
                        Set DynamicTextures(A).Texture = Nothing
                        DynamicTextures(A).Loaded = False
                        DynamicTextures(A).LastUsed = 0
                    Next A
                    If Not InitD3D(False) Then
                       GoTo exitit
                    End If

                    InitMiniMap
                    If CurrentTab = tsMap Then
                        SetMiniMapTab tsMap
                    End If
                End If
                           
                            
                
                If MusicFading Then
                    If MusicFade > 0 Then
                        MusicFade = MusicFade - 1
                        Sound_SetStreamFadeVolume
                    Else
                        Sound_StopStream
                    End If
                End If


                MovePlayers
                SecondCounter = SecondCounter + 1
                CheckKeys
                
                If DisableScreen = False Then
                    DoEvents
                    If BOOLD3DERROR = True Then FixD3DError
                    checkDraw
                End If
            Else
                If D3DDevice.TestCooperativeLevel = D3D_OK Then
                    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
                End If
            End If

            
            If SecondCounter Mod 4 = 0 Then
              'Draw skills, using this as a delay just because its here
                If CurrentTab = tsSkills Then
                    If RedrawSkills Then
                        RedrawSkills = False
                        frmMain.DrawLstBox
                    End If
                End If
            End If
            

            If Second(Time) <> FrameTimer Then
                If spamdelay > 0 Then spamdelay = spamdelay - 1
                If alternateFrame Then
                    alternateFrame = False
                Else
                    alternateFrame = True
                End If
                            
                For A = 0 To 49
                    If map.Object(A).TimeStamp > 1000 Then
                        map.Object(A).TimeStamp = map.Object(A).TimeStamp - 1000
                    End If
                Next A
                
                If DisableScreen = False Then
                    Dim R1 As RECT, r2 As RECT

                    
                    

                    r2.Top = 0
                    r2.Left = 0
                    r2.Bottom = 36
                    r2.Right = 78
                    
                    R1.Top = 0
                    R1.Left = 522
                    R1.Right = R1.Left + 78
                    R1.Bottom = 36
                
        
                    Draw sfcTimers, 0, 0, 78, 36, sfcInventory(0), 145, 393, True, 0
                    If combatCounter > 0 Then
                        Draw sfcTimers, 2, 2, 32, 32, sfcInventory(0), 223, 397, True, 0
                        If hoverCombat Then Draw sfcTimers, 8, 11, 121, 17, sfcInventory(0), 243, 625, True, 0
                    End If
                    
                    R1.Top = R1.Top * WindowScaleY: R1.Bottom = R1.Bottom * WindowScaleY: R1.Left = R1.Left * WindowScaleX: R1.Right = R1.Right * WindowScaleX
                    sfcTimers.Surface.BltToDC frmMain.hdc, r2, R1
                    
                    If combatCounter > 0 Then
                        If hoverCombat Then
                            frmMain.FontSize = 9 * WindowScaleY
                            TextOut frmMain.hdc, 533 * WindowScaleX, 12 * WindowScaleY, "Combat: " & combatCounter, IIf(combatCounter > 9, 10, 9)
                            frmMain.FontSize = 9
                        End If
                        combatCounter = combatCounter - 1
                        If combatCounter = 0 And CurrentTab = tsStats2 Then
                            calculateHpRegen
                            calculateManaRegen
                            DrawMoreStats
                        End If
                    End If
                    If Not drawCooldowns Then frmMain.Refresh
                    

                End If

                If World.Minute > 0 Then World.Minute = World.Minute - 1
                CalculateAmbientAlpha
                If SkipTimeOff = 0 Then
                    If (Second(Time) - FrameTimer) > 1 Then
                        SkipTimeOff = SkipTimeOff + 1
                    Else
                        TimeRatio = (GetTickCount - LastTime)
                        If TimeRatio >= 2000 Then
                            'tickCountMod = TimeRatio
                            speedhack = speedhack + 1
                            If speedhack >= 3 Then SendSocket (Chr$(86) + Chr$(2))
                            'MsgBox "test"
                            'GoTo exitit
                        Else
                            If speedhack > 0 Then speedhack = speedhack - 1
                        End If
                        'PrintChat TimeRatio, 3, Options.FontSize, 0, False
                        If NumTimeOff ^ 4 + 7 <> NumTimeOff2 Then
                            GoTo exitit
                        End If
                        If TimeRatio >= (1099) Then 'TimeRatio <= (1000 - 500) Or
                            NumTimeOff = NumTimeOff + 5
                          '  PrintChat Str(TimeRatio), 4, 10
                        Else
                            If NumTimeOff > 0 Then NumTimeOff = NumTimeOff - 1
                        End If
                        If TimeRatio <= 900 Then
                           ' PrintChat Str(TimeRatio), 4, 10
                            If NumTimeOff >= 20 Then NumTimeOff = NumTimeOff - 20
                        End If
                        NumTimeOff2 = NumTimeOff ^ 4 + 7
                        If NumTimeOff > 100 Then
                            SendSocket (Chr$(86) + Chr$(1))
                            NumTimeOff = 10
                            NumTimeOff2 = NumTimeOff ^ 4 + 7
                            'GoTo exitit
                        End If
                    End If
                Else
                    SkipTimeOff = SkipTimeOff - 1
                End If
                FrameTimer = Second(Time)
                FrameRate = SecondCounter
                SecondCounter = 0
                LastTime = GetTickCount
                
                If Not STOPMAPCHECK Then
                    If (CMap ^ 2 + 5 <> CMap2) Or (cX ^ 2 + 5 <> CX2) Or (cY ^ 2 + 5 <> CY2) Then
                        End
                    End If
                End If
            
                For A = 0 To MaxTraps
                    With map.Trap(A)
                        If .Counter > 0 Then .Counter = .Counter - 1
                    End With
                Next A

                'Unload textures
                For A = 1 To NumTextures
                    If DynamicTextures(A).Loaded Then
                        If DynamicTextures(A).LastUsed < GetTickCount Then
                            DynamicTextures(A).Loaded = False
                            DynamicTextures(A).LastUsed = 0
                            Set DynamicTextures(A).Texture = Nothing
                        End If
                    End If
                Next A

                For A = 1 To 255
                    If Sounds(A).Loaded Then
                        If Sounds(A).LastUsed < GetTickCount Then
                            Sound_Free A
                        End If
                    End If
                Next A
            End If
        End If
        If Character.AttackSpeed > 0 And drawCooldowns Then
            A = ((GetTickCount - AttackTimer) / Character.AttackSpeed) * 32
            r2.Top = 0
            r2.Left = 0
            r2.Bottom = 32
            r2.Right = 64
          
            R1.Top = 2
            R1.Left = 212
            R1.Right = R1.Left + 64
            R1.Bottom = 34
            If A > 32 Then A = 32
            If A = 32 Then drawCooldowns = False
            Draw sfcTimers2, 0, 32 - A, 32, A, sfcInventory(0), 286, 292 + (32 - A), True, 0
            A = 28 - 28 * ((Character.GlobalSpellTick - GetTickCount) / 1000)
            If A > 28 Then A = 28
            If A < 28 And drawCooldowns = False Then drawCooldowns = True
            Draw sfcTimers2, 32, 30 - A, 28, A, sfcInventory(0), 548, 315 + 28 + (28 - A), True, 0
            
            R1.Top = R1.Top * WindowScaleY: R1.Bottom = R1.Bottom * WindowScaleY: R1.Left = R1.Left * WindowScaleX: R1.Right = R1.Right * WindowScaleX
            sfcTimers2.Surface.BltToDC frmMain.hdc, r2, R1
            frmMain.Refresh
        End If
        
        
        DoEvents
        If (Not Options.VsyncEnabled) Then
            A = TargetFps - 1 - (GetTickCount - t)
            If A > 0 Then
                If A < TargetFps And A > 1 Then
                    'B = GetTickCount
                    If Not Options.highpriority Then
                        Sleep A - 1
                    End If
                    'PrintChat GetTickCount - B, 5, 10, 15
                End If
            End If
            A = GetTickCount - t
            While GetTickCount - t < TargetFps
               DoEvents
            Wend
             '   tickCountMod = 0
        End If
        t = GetTickCount
        

    Wend
exitit:
    DeInitialize
Exit Sub
ERRUPDATER:
MsgBox "Failed to load Updater.exe.  Run this application with administrator permissions or run the updater directly."
Exit Sub
ERRHANDLER:
Open AppPath & "log.txt" For Append As #1
    Print #1, "THERE HAS BEEN AN ERROR - " & Err.Number & ": " & Err.Description & " (" & Erl & ")"
Close #1
End
End Sub

Sub checkDraw()
    D3DDevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, ByVal 0, 1, 0
    If (Character.StatusEffect And (2 ^ SE_BLIND)) = 0 Then
        CreateShadeMap
        CreateLightMap
        D3DDevice.BeginScene
            DrawNextFrame3D
            If MapEdit = False Then DrawShadeMap
            DrawWidgets Widgets
            DrawGUIWindow
        D3DDevice.EndScene
    End If
    If D3DDevice.TestCooperativeLevel = D3D_OK Then
        D3DDevice.Present ByVal 0&, ByVal 0&, 0, ByVal 0&
         
        If cX <> LastCX Or cY <> LastCY Or doMiniMapDraw Then
            DrawMiniMapSection
        End If
    End If


End Sub


