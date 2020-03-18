Attribute VB_Name = "modServer"
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
'Game Constants



Public TitleString As String
Public currentMaxUser As Byte
Public Const DownloadSite = "http://www.Seyerdin.com"
Public ServerAdminPass As String

Public Const CurrentClientVer = "56" 'client version
Public Word(1 To 50) As String

'Debugging Constants
#Const DEBUGWINDOWPROC = 0
#Const USEGETPROP = 0

Type SocketQueueData
    user As Long
    TimeStamp As Long
End Type

#If DEBUGWINDOWPROC Then
Private m_SCHook As WindowProcHook
#End If

Public Const modeNotConnected = 0
Public Const modeConnected = 1
Public Const modePlaying = 2

'Blacksmithy Consts
Public Const Cost_Per_Durability = 1
Public Const Cost_Per_Strength = 3
Public Const Cost_Per_Modifier = 25

'User Defined Types
Public Const MaxPlayerTimers = 4
Public Const pttCharacter = 0
Public Const pttPlayer = 1
Public Const pttMonster = 2
Public Const pttTile = 3
Public Const pttProject = 4

Public NumUsers As Long

Type InvObject
    Object As Integer
    Value As Long 'If money, holds value, otherwise holds current Dur
    prefix As Byte 'Holds Prefix
    prefixVal As Byte
    suffix As Byte 'Holds Suffix
    SuffixVal As Byte
    Affix As Byte 'Holds Affix
    AffixVal As Byte
    ObjectColor As Byte
    Flags(0 To 3) As Long
End Type

'Misc Variables
Public blnNight As Boolean
Public BackupCounter As Long
Public LastDate As Long

Public Declare Function QueryPerformanceCounter Lib "kernel32" ( _
    lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
    x As Currency) As Boolean
    
'Hook
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = -4
Public lpPrevWndProc As Long
Public gHW As Long

'Sockets
Public ListeningSocket As Long

#Const UseExperience = True
#Const CheckIPDupe = False
#Const AdminCheck = False
#Const GodChecking = False
#Const PublicServer = False

Function BanPlayer(A As Long, Index As Long, NumDays As Long, reason As String, Banner As String) As Boolean
    Dim C As Long, ST1 As String
    BanPlayer = False
    If A > 0 And A < 82 Then
        With player(A)
            If .Mode = modePlaying Then
                C = FreeBanNum
                If C >= 1 Then
                    ST1 = .user
                    If Len(ST1) > 0 Then
                        With Ban(C)
                            .user = player(A).user
                            .Name = player(A).Name
                            .ip = player(A).ip
                            .reason = reason
                            .Banner = Banner
                            .uId = player(A).uId
                            .iniUID = player(A).iniUID
                            .InUse = True
                            .UnbanDate = CLng(Date) + NumDays
                            BanRS.Seek "=", C
                            If BanRS.NoMatch = True Then
                                BanRS.AddNew
                                BanRS!Number = C
                            Else
                                BanRS.Edit
                            End If
                            BanRS!user = player(A).user
                            BanRS!Name = player(A).Name
                            BanRS!ip = player(A).ip
                            BanRS!reason = .reason
                            BanRS!UnbanDate = .UnbanDate
                            BanRS!uId = player(A).uId
                            BanRS!iniUID = player(A).iniUID
                            BanRS!Banner = .Banner
                            BanRS.Update
                            SendSocket A, Chr2(67) + Chr2(Index) + .reason
                            SendAllBut A, Chr2(66) + Chr2(A) + Chr2(Index) + .reason
                            AddSocketQue A, 0
                            BanPlayer = True
                        End With
                    End If
                'Else
                '    BanPlayer = False
                End If
            End If
        End With
    End If
End Function
Sub BootPlayer(A As Long, Index As Long, reason As String)
    If A > 0 And A < 82 Then
        With player(A)
            If .InUse = True Then
                If reason <> "" Then
                    If .Access = 0 Then SendSocket A, Chr2(67) + Chr2(Index) + reason
                    If .Mode = modePlaying Then
                        SendToGods Chr2(68) + Chr2(A) + Chr2(Index) + reason
                        PrintLog player(A).Name & " has been booted! " + reason
                    Else
                        'SendAllBut A, chr2(56) + chr2(15) + "User " + chr2(34) + .User + chr2(34) + " with name " + chr2(34) + .Name + chr2(34) + " has been booted: " + Reason
                    End If
                    If .Access = 0 Then AddSocketQue A, 3000
                Else
                    If .Access = 0 Then SendSocket A, Chr2(67) + Chr2(Index)
                    If .Mode = modePlaying Then
                        SendToGods Chr2(68) + Chr2(A) + Chr2(Index)
                        PrintLog player(A).Name + " has been booted!"
                    Else
                        'SendAllBut A, chr2(56) + chr2(15) + "User " + chr2(34) + .User + chr2(34) + " with name " + chr2(34) + .Name + chr2(34) + " has been booted!"
                    End If
                    If .Access = 0 Then AddSocketQue A, 3000
                End If
            End If
        End With
    End If
End Sub

Sub Hacker(Index As Long, Code As String)
    If Code <> "C.1" Then
        BootPlayer Index, 0, "Possible Hacking Attempt: Code '" + Code + "' from IP '" + player(Index).ip + "'"
    Else
        AddSocketQue Index, 3000
    End If
End Sub

Sub PrintDebug(St As String)
    Open "debug.log" For Append As #1
    Print #1, St
    Close #1
End Sub
Sub PrintCrashDebug(ByRef A As Long, ByRef B As Long)

    Dim stringcount As Long
    stringcount = StringPointer

    Open "debugCrash.log" For Append As #1
    Print #1, (A) & " - " & (B)
    Close #1
    
    For A = stringcount To StringPointer - 1
        SysFreeString StringStack(A)
    Next A
    StringPointer = stringcount
    
End Sub


Sub SaveFlags()
    Dim A As Long, St As String
    For A = 0 To 255
        St = St + QuadChar(World.Flag(A))
    Next A
    For A = 0 To 127
        DataRS.Edit
        DataRS("StringFlag" + CStr(A)) = World.StringFlag(A)
        DataRS.Update
    Next A
    DataRS.Edit
    DataRS!Flags = St
    DataRS.Update
End Sub
Sub SaveObjects()
    Dim A As Long, B As Long, St As String, C As Long
    For A = 1 To 5000
        With map(A)
            If .Keep = True Then
                For B = 0 To 49
                    With .Object(B)
                        If .Object > 0 Then
                            If map(A).Tile(.x, .y).Att = 5 Then
                                St = St + DoubleChar(CInt(A)) + Chr2(B) + Chr2(.x) + Chr2(.y) + DoubleChar(CLng(.Object)) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal)
                                For C = 0 To 3
                                    St = St + QuadChar(.Flags(C))
                                Next C
                                St = St + Chr2(.ObjectColor)
                            End If
                        End If
                    End With
                Next B
            End If
        End With
    Next A
    DataRS.Edit
    DataRS!ObjectData = St
    DataRS.Update
End Sub

Sub CheckGuild(Index As Long)
    If Guild(Index).Name <> "" Then
        If CountGuildMembers(Index) < 3 Then
            'Not enough players -- delete guild
            DeleteGuild Index, 1
        End If
    End If
End Sub

Function CheckSum(St As String) As Long
    Dim A As Long, B As Long
    For A = 1 To Len(St)
        B = B + Asc(Mid$(St, A, 1))
    Next A
    CheckSum = B
End Function
Function CountGuildMembers(Index As Long) As Long
    Dim A As Long, B As Long
    With Guild(Index)
        If .Name <> "" Then
            B = 0
            For A = 0 To 19
                If .Member(A).Name <> "" Then
                    B = B + 1
                End If
            Next A
            CountGuildMembers = B
        End If
    End With
End Function
Function CountGuildMembersOL(Index As Long) As Long
    Dim A As Long, B As Long
    If Index > 0 Then
    With Guild(Index)
        If .Name <> "" Then
            B = 0
            For A = 0 To 19
                If .Member(A).Name <> "" Then
                    If IsPlaying(FindPlayer(.Member(A).Name)) Then
                        B = B + 1
                    End If
                End If
            Next A
            CountGuildMembersOL = B
        End If
    End With
    Else
        CountGuildMembersOL = 0
    End If
End Function
Sub UpdateGuildInfo(A As Long)
    With Guild(A)
        If .Name <> "" Then
            SendAll Chr2(136) + Chr2(A) + Chr2(CountGuildMembers(A)) + Chr2(CountGuildMembersOL(A)) + QuadChar(GetGuildRenown(A) / CountGuildMembers(A)) + Chr2(.Symbol1) + Chr2(.Symbol2) + Chr2(.Symbol3) + Chr2(.Hall)
        End If
    End With
End Sub

Function GetGuildRenown(Index As Long) As Long
    Dim A As Long, B As Long
    With Guild(Index)
    If .Name <> "" Then
    B = 0
    For A = 0 To 19
        If .Member(A).Name <> "" Then
            B = B + .Member(A).Renown
        End If
    Next A
    GetGuildRenown = B
    End If
    End With
End Function



Sub CreateModString()
ModString(0) = "None"
ModString(1) = "Strength"
End Sub

Sub CreateClassData()
    Dim A As Long
    For A = 1 To 10
        With Class(A)
            .Name = ReadString("Classes", "Class" + CStr(A), "Name")
            .StartHP = ReadInt("Classes", "Class" + CStr(A), "StartHP")
            .StartEnergy = ReadInt("Classes", "Class" + CStr(A), "StartEnergy")
            .StartMana = ReadInt("Classes", "Class" + CStr(A), "StartMana")
            .StartStrength = ReadInt("Classes", "Class" + CStr(A), "StartStrength")
            .StartAgility = ReadInt("Classes", "Class" + CStr(A), "StartAgility")
            .StartEndurance = ReadInt("Classes", "Class" + CStr(A), "StartEndurance")
            .StartWisdom = ReadInt("Classes", "Class" + CStr(A), "StartWisdom")
            .StartConstitution = ReadInt("Classes", "Class" + CStr(A), "StartConstitution")
            .StartIntelligence = ReadInt("Classes", "Class" + CStr(A), "StartIntelligence")
            .HPIncrement = ReadInt("Classes", "Class" + CStr(A), "HPIncrement")
            .ManaIncrement = ReadInt("Classes", "Class" + CStr(A), "ManaIncrement")
            .Enabled = ReadInt("Classes", "Class" + CStr(A), "Enabled")
        End With
    Next A
End Sub

Sub DeleteCharacter()
    Dim A As Long, B As Long, St As String

    St = UserRS!Name
    For A = 1 To 255
        With Guild(A)
            If .Name <> "" Then
                For B = 0 To 19
                    With .Member(B)
                        If .Name = St Then
                            .Name = ""
                            CheckGuild A
                        End If
                    End With
                Next B
            End If
        End With
    Next A
End Sub
Sub DeleteGuild(Index As Long, reason As Byte)
    Dim A As Long, B As Long, C As Long
    
    With Guild(Index)
        If .Name <> "" Then
            .Name = ""
            GuildRS.Bookmark = .Bookmark
            GuildRS.Delete
        End If
        
        UserRS.Index = "Name"
        For A = 0 To 19
            With .Member(A)
                If .Name <> "" Then
                    B = FindPlayer(.Name)
                    If B > 0 Then
                        With player(B)
                            .Guild = 0
                            .GuildRank = 0
                            If Guild(Index).sprite > 0 Then
                                .sprite = .Class * 2 + .Gender - 1
                                SendAll Chr2(63) + Chr2(B) + Chr2(.sprite)
                            End If
                            SendSocket B, Chr2(75) + Chr2(reason)
                            SendAllBut B, Chr2(73) + Chr2(B) + Chr2(0)
                        End With
                    ElseIf Guild(Index).sprite > 0 Then
                        UserRS.Seek "=", .Name
                        If UserRS.NoMatch = False Then
                            B = UserRS!Class * 2 + UserRS!Gender - 1
                            If B >= 1 And B <= 255 Then
                                UserRS.Edit
                                UserRS!sprite = B
                                UserRS.Update
                            End If
                        End If
                    End If
                End If
            End With
        Next A
    End With
    
    'Check if other guilds have declarations
    For A = 1 To 255
        With Guild(A)
            If .Name <> "" Then
                C = 0
                For B = 0 To 4
                    With .Declaration(B)
                        If .Guild = Index Then
                            .Guild = 0
                            SendToGuild A, Chr2(71) + Chr2(B) + Chr2(0) + Chr2(0) 'Declaration Data
                            C = 1
                        End If
                    End With
                Next B
                If C = 1 Then
                    GuildRS.Bookmark = .Bookmark
                    GuildRS.Edit
                    For B = 0 To 4
                        With .Declaration(B)
                            GuildRS("DeclarationGuild" + CStr(B)) = .Guild
                            GuildRS("DeclarationType" + CStr(B)) = .Type
                        End With
                    Next B
                    GuildRS.Update
                End If
            End If
        End With
    Next A
    
    'Erase Join Requests
    For A = 1 To MaxUsers
        With player(A)
            If .JoinRequest = Index Then .JoinRequest = 0
        End With
    Next A
    
    SendAll Chr2(70) + Chr2(Index) 'Erase Guild
End Sub

Sub DeleteAccount()
    On Error Resume Next
    
    If UserRS!Class > 0 Then
        DeleteCharacter
    End If

    UserRS.Delete
    
    On Error GoTo 0
End Sub
Function FindBan(Index As Long) As Long
    Dim A As Long
    
    For A = 1 To 50
        If Ban(A).InUse = True Then
            If UCase$(Ban(A).user) = (player(Index).user) Or Ban(A).ip = player(Index).ip Or Ban(A).iniUID = player(Index).iniUID Or Ban(A).uId = player(Index).uId Then
                FindBan = A
                Exit Function
            End If
        End If
    Next A
End Function

Sub LoadObjectData(ObjectData As String)
    Dim A As Long, NumObjects As Long, B As Long
    NumObjects = Len(ObjectData) / 34 - 1
    For A = 0 To NumObjects
        With map(Asc(Mid$(ObjectData, A * 34 + 1, 1)) * 256 + Asc(Mid$(ObjectData, A * 34 + 2, 1))).Object(Asc(Mid$(ObjectData, A * 34 + 3, 1)))
            .x = Asc(Mid$(ObjectData, A * 34 + 4, 1))
            .y = Asc(Mid$(ObjectData, A * 34 + 5, 1))
            .Object = Asc(Mid$(ObjectData, A * 34 + 6, 1)) * 256 + Asc(Mid$(ObjectData, A * 34 + 7, 1))
            .Value = Asc(Mid$(ObjectData, A * 34 + 8, 1)) * 16777216 + Asc(Mid$(ObjectData, A * 34 + 9, 1)) * 65536 + Asc(Mid$(ObjectData, A * 34 + 10, 1)) * 256& + Asc(Mid$(ObjectData, A * 34 + 11, 1))
            .prefix = Asc(Mid$(ObjectData, A * 34 + 12, 1))
            .prefixVal = Asc(Mid$(ObjectData, A * 34 + 13, 1))
            .suffix = Asc(Mid$(ObjectData, A * 34 + 14, 1))
            .SuffixVal = Asc(Mid$(ObjectData, A * 34 + 15, 1))
            .Affix = Asc(Mid$(ObjectData, A * 34 + 16, 1))
            .AffixVal = Asc(Mid$(ObjectData, A * 34 + 17, 1))
            For B = 0 To 3
                .Flags(B) = Asc(Mid$(ObjectData, A * 34 + 18 + B * 4, 1)) * 16777216 + Asc(Mid$(ObjectData, A * 34 + 19 + B * 4, 1)) * 65536 + Asc(Mid$(ObjectData, A * 34 + 20 + B * 4, 1)) * 256& + Asc(Mid$(ObjectData, A * 34 + 21 + B * 4, 1))
            Next B
            .ObjectColor = Asc(Mid$(ObjectData, A * 34 + 34, 1))
        End With
    Next A
End Sub
Function NPCNum(ByVal Name As String) As Long
    Name = UCase$(Name)
    Dim A As Long
    For A = 1 To 255
        With NPC(A)
            If UCase$(.Name) = Name Then
                NPCNum = A
                Exit Function
            End If
        End With
    Next A
End Function
Function FindGuildMember(ByVal Name As String, GuildNum As Long) As Long
    Name = UCase$(Name)
    Dim A As Long
    With Guild(GuildNum)
        For A = 0 To 19
            If UCase$(.Member(A).Name) = Name Then
                FindGuildMember = A
                Exit Function
            End If
        Next A
    End With
    FindGuildMember = -1
End Function

Function FindInvObject(Index As Long, ByVal ObjectNum As Long, Equipped As Boolean, Optional Start As Long = 1) As Long
    FindInvObject = Start - 1

    Dim A As Long
    With player(Index)
        If Equipped Then
            For A = 1 To 5
                If .Equipped(A).Object = ObjectNum Then
                    FindInvObject = 20 + A
                    Exit Function
                End If
            Next A
        Else
            For A = Start To 20
                If .Inv(A).Object = ObjectNum Then
                    FindInvObject = A
                    Exit Function
                End If
            Next A
        End If
    End With
End Function

Function FindStackableStorageObject(Index As Long, ObjectNum As Long)
'PLEASE NOTE THAT THIS FUNCTION SETS THE CURSTORAGEPAGE
'TELLS THE CLIENT TO GO TO THAT PAGE
'AND THEN ADDS. THIS FUNCTION IS ONLY USED FOR STACKABLE OBJECTS!!!
    Dim A As Long, B As Long
    With player(Index)
        For A = 1 To STORAGEPAGES
            For B = 1 To 20
                If .Storage(A, B).Object = ObjectNum Then
                    If A <> .CurStoragePage Then
                        .CurStoragePage = A
                        SendSocket Index, Chr2(117) + Chr2(4) + Chr2(player(Index).CurStoragePage)
                    End If
                    FindStackableStorageObject = B
                    Exit Function
                End If
            Next B
        Next A
    End With
End Function


Function FindPlayer(ByVal Name As String) As Long
    Name = UCase$(Name)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .InUse = True And UCase$(.Name) = Name Then
                FindPlayer = A
                Exit Function
            End If
        End With
    Next A
End Function

Function FreeBanNum() As Long
    Dim A As Long
    For A = 1 To 50
        If Ban(A).InUse = False Then
            FreeBanNum = A
            Exit For
        End If
    Next A
End Function

Function FreeGuildDeclarationNum(GuildNum As Long) As Long
    Dim A As Long
    With Guild(GuildNum)
        For A = 0 To 4
            If .Declaration(A).Guild = 0 Then
                FreeGuildDeclarationNum = A
                Exit Function
            End If
        Next A
    End With
    FreeGuildDeclarationNum = -1
End Function
Function FreeGuildMemberNum(GuildNum As Long)
    Dim A As Long
    With Guild(GuildNum)
        For A = 0 To 19
            If .Member(A).Name = "" Then
                FreeGuildMemberNum = A
                Exit Function
            End If
        Next A
    End With
    FreeGuildMemberNum = -1
End Function

Function FreeGuildNum() As Long
    Dim A As Long
    For A = 1 To 255
        If Guild(A).Name = "" Then
            FreeGuildNum = A
            Exit Function
        End If
    Next A
End Function
Function FreeInvNum(Index As Long) As Long
    Dim A As Long
    With player(Index)
        For A = 1 To 20
            If .Inv(A).Object = 0 Then
                FreeInvNum = A
                Exit Function
            End If
        Next A
    End With
End Function

Function FreeStorageNum(Index As Long) As Long
    Dim A As Long
    With player(Index)
        For A = 1 To 20
            If .Storage(.CurStoragePage, A).Object = 0 Then
                FreeStorageNum = A
                Exit Function
            End If
        Next A
    End With
End Function

Function FreeMapDoorNum(mapNum As Long) As Long
    Dim A As Long
    With map(mapNum)
        For A = 0 To 9
            If .Door(A).Att = 0 And .Door(A).Wall = 0 Then
                FreeMapDoorNum = A
                Exit Function
            End If
        Next A
    End With
    FreeMapDoorNum = -1
End Function
Function FreeMapObj(mapNum As Long) As Long
    Dim A As Long
    If mapNum >= 1 Then
        With map(mapNum)
            For A = 0 To 49
                If .Object(A).Object = 0 Then
                    FreeMapObj = A
                    Exit Function
                End If
            Next A
        End With
    End If
    FreeMapObj = -1
End Function
Function FreePlayer() As Long
    Dim A As Long, B As Long, C As Long
    For A = 1 To currentMaxUser
        If player(A).Mode = modeNotConnected And player(A).LoginStamp + 20000 < GetTickCount Then
            CloseClientSocket (A)
        End If
    Next A
    For A = 1 To MaxUsers
        If player(A).InUse = False Then
            FreePlayer = A
            If A > currentMaxUser Then currentMaxUser = A
            Exit Function
        End If
    Next A
    B = GetTickCount
    C = 0
    For A = 1 To MaxUsers
        If player(A).Mode = modeNotConnected And player(A).LoginStamp < B Then
            B = player(A).LoginStamp
            C = A
        End If
    Next A
    If C > 0 Then
        CloseClientSocket C
        FreePlayer = C
    End If
End Function
Sub GainExp(Index As Long, ByVal EXP As Long, Split As Boolean, PK As Boolean, Optional IsMonster As Byte = 0, Optional overflow As Boolean = False)
Dim A As Long, PartyExp As Long, pMembers As Byte, F As Long, carryover As Long
Dim tDouble As Double
Dim St As String


    'Perform EXP Attenuation if monster
    If IsMonster Then
        If Abs(CLng(player(Index).Level) - CLng(IsMonster)) > 8 Then
            tDouble = CDbl(Abs(CLng(player(Index).Level) - CLng(IsMonster) - 8)) * 0.13
            tDouble = EXP * tDouble
            If tDouble > 0 Then EXP = EXP - tDouble Else EXP = 0
            If EXP < 1 Then EXP = 1
        End If
    End If

    If Split Then 'Split EXP if in party
        pMembers = 0
        If player(Index).Party > 0 Then
            For A = 1 To currentMaxUser
                If player(A).Party = player(Index).Party And player(A).InUse And player(A).map = player(Index).map And Not A = Index Then
                    pMembers = pMembers + 1
                End If
            Next A
            If pMembers > 0 And PK = False Then
                PartyExp = (EXP * (1.3 + ((pMembers - 1) * 0.05))) / (pMembers + 1)
                If PartyExp < 1 Then PartyExp = 1
                For A = 1 To currentMaxUser
                    If player(A).Party = player(Index).Party And player(A).InUse And player(A).map = player(Index).map And Not Index = A Then
                        If IsMonster And Abs(CLng(player(A).Level) - CLng(IsMonster)) > 5 Then
                            With player(A)
                                F = PartyExp
                                tDouble = CDbl(Abs(CLng(player(A).Level) - CLng(IsMonster) - 5)) * 0.13
                                tDouble = F * tDouble
                                If tDouble > 0 Then F = F - tDouble Else F = 0
                            End With
                            If F < 1 Then F = 1
                            'If F > 0 Then
                            GainExp A, F, False, False, IsMonster
                            'Else
                            '    CreateFloatingText Player(A).Map, Player(A).X, Player(A).Y, 14, "1"
                            'End If
                        Else
                            GainExp A, PartyExp, False, False
                        End If
                    End If
                Next A
                EXP = PartyExp
            End If
        End If
    End If

    With player(Index)
        If .Level <> MaxLevel Then
        
            If .Class = 10 Then EXP = EXP / 4
            CreateFloatingTextAllBut Index, .map, .x, .y, 14, CStr(EXP)
            If GetPlayerFlag(Index, 226) > 0 And overflow = False Then EXP = (EXP * GetPlayerFlag(Index, 226)) / 10
            
            CreateIndividualFloatingText Index, 14, CStr(EXP)
            
            'SKILL: AMBITION
            'If CanUseSkill(CByte(Index), SKILL_AMBITION) Then
            '    EXP = EXP + (.SkillLevel(SKILL_AMBITION) * EXP) / 100
            'End If
            
            If CDbl(.Experience) + CDbl(EXP) > 2147483647# Then
                .Experience = 2147483647
            Else
                .Experience = .Experience + EXP
            End If
            If .Experience >= EXPLevel(.Level) And .Level < MaxLevel Then
                PrintLog player(Index).user & "(" & player(Index).Name & ")" & " Leveled up from " & .Level & "   class: " & Class(.Class).Name
                carryover = .Experience - EXPLevel(.Level)
                .Level = .Level + 1
                Parameter(0) = Index
                RunScript ("LEVELUP")
                .Experience = 0
                .OldHP = .OldHP + Class(.Class).HPIncrement
                .OldMana = .OldMana + Class(.Class).ManaIncrement
                CalculateStats Index
                .HP = .MaxHP
                .Mana = .MaxMana
                .Energy = .MaxEnergy
                .StatPoints = .StatPoints + StatsPerLevel
                .SkillPoints = .SkillPoints + SkillsPerLevel
                
                St = DoubleChar(3) + Chr2(46) + DoubleChar(CInt(.HP))
                St = St + DoubleChar(3) + Chr2(47) + DoubleChar(CInt(.Energy))
                St = St + DoubleChar(3) + Chr2(48) + DoubleChar(CInt(.Mana))
                St = St + DoubleChar(11) + Chr2(59) + DoubleChar(CInt(.MaxHP)) + DoubleChar(CInt(.MaxEnergy)) + DoubleChar(CInt(.MaxMana)) + DoubleChar(CInt(.StatPoints)) + DoubleChar(CInt(.SkillPoints))
                
                SendRaw Index, St
    
    

                    SendToPartyAllBut .Party, Index, Chr2(104) + Chr2(0) + Chr2(Index) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                    SendToGods Chr2(104) + Chr2(0) + Chr2(Index) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                    SendToGuildAllBut Index, CLng(.Guild), Chr2(104) + Chr2(0) + Chr2(Index) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                    SendToPartyAllBut .Party, Index, Chr2(104) + Chr2(1) + Chr2(Index) + Chr2(.Level)
  
                
                If carryover > 0 Then GainExp Index, carryover, False, False, 0, True
                Exit Sub
            End If
            
            SendSocket Index, Chr2(60) + QuadChar(.Experience)
        End If
    End With
End Sub
Function GuildNum(ByVal Name As String) As Long
    Name = UCase$(Name)
    Dim A As Long
    For A = 1 To 255
        With Guild(A)
            If UCase$(.Name) = Name Then
                GuildNum = A
                Exit Function
            End If
        End With
    Next A
End Function

Function IsVacant(mapNum As Long, x As Byte, y As Byte, FromDir As Byte, Optional ByVal mNum As Long = -1) As Byte
    'printcrashdebug 6, 1
    Dim A As Long
        With map(mapNum)
         If mNum >= 0 Then
            If monster(map(mapNum).monster(mNum).monster).Flags2 And MONSTER_LARGE Then
    
                Select Case FromDir
                Case 0
                    Select Case .Tile(x, y + 1).Att
                        Case 1, 2, 3, 10, 15 'Wall / Warp / Door / No Monsters
                            Exit Function
                        Case 24
                            If .Tile(x, y + 1).AttData(3) > 0 Then Exit Function
                    End Select
                    Select Case .Tile(x + 1, y + 1).Att
                        Case 1, 2, 3, 10, 15 'Wall / Warp / Door / No Monsters
                            Exit Function
                        Case 24
                            If .Tile(x + 1, y + 1).AttData(3) > 0 Then Exit Function
                    End Select
                Case 1
                    Select Case .Tile(x, y).Att
                        Case 1, 2, 3, 10, 15 'Wall / Warp / Door / No Monsters
                            Exit Function
                        Case 24
                            If .Tile(x, y).AttData(3) > 0 Then Exit Function
                    End Select
                    Select Case .Tile(x + 1, y).Att
                        Case 1, 2, 3, 10, 15 'Wall / Warp / Door / No Monsters
                            Exit Function
                        Case 24
                            If .Tile(x + 1, y).AttData(3) > 0 Then Exit Function
                    End Select
                Case 2
                    Select Case .Tile(x + 1, y).Att
                        Case 1, 2, 3, 10, 15 'Wall / Warp / Door / No Monsters
                            Exit Function
                        Case 24
                            If .Tile(x + 1, y).AttData(3) > 0 Then Exit Function
                    End Select
                    Select Case .Tile(x + 1, y + 1).Att
                        Case 1, 2, 3, 10, 15 'Wall / Warp / Door / No Monsters
                            Exit Function
                        Case 24
                            If .Tile(x + 1, y + 1).AttData(3) > 0 Then Exit Function
                    End Select
                Case 3
                    Select Case .Tile(x, y).Att
                        Case 1, 2, 3, 10, 15 'Wall / Warp / Door / No Monsters
                            Exit Function
                        Case 24
                            If .Tile(x, y).AttData(3) > 0 Then Exit Function
                    End Select
                    Select Case .Tile(x, y + 1).Att
                        Case 1, 2, 3, 10, 15 'Wall / Warp / Door / No Monsters
                            Exit Function
                        Case 24
                            If .Tile(x, y + 1).AttData(3) > 0 Then Exit Function
                    End Select
                End Select
              
                
            Else
                Select Case .Tile(x, y).Att
                    Case 1, 2, 3, 10, 15 'Wall / Warp / Door / No Monsters
                        Exit Function
                    Case 24
                        If .Tile(x, y).AttData(3) > 0 Then Exit Function
                    Case 17
                        Select Case FromDir
                            Case 1
                                If ExamineBit(map(mapNum).Tile(x, y).AttData(0), 0) Then Exit Function
                            Case 0
                                If ExamineBit(map(mapNum).Tile(x, y).AttData(0), 2) Then Exit Function
                            Case 3
                                If ExamineBit(map(mapNum).Tile(x, y).AttData(0), 4) Then Exit Function
                            Case 2
                                If ExamineBit(map(mapNum).Tile(x, y).AttData(0), 6) Then Exit Function
                        End Select
                End Select
            End If
        Else
            Select Case .Tile(x, y).Att
             Case 1, 2, 3, 10, 15 'Wall / Warp / Door / No Monsters
                 Exit Function
             Case 24
                 If .Tile(x, y).AttData(3) > 0 Then Exit Function
             Case 17
                 Select Case FromDir
                     Case 1
                         If ExamineBit(map(mapNum).Tile(x, y).AttData(0), 0) Then Exit Function
                     Case 0
                         If ExamineBit(map(mapNum).Tile(x, y).AttData(0), 2) Then Exit Function
                     Case 3
                         If ExamineBit(map(mapNum).Tile(x, y).AttData(0), 4) Then Exit Function
                     Case 2
                         If ExamineBit(map(mapNum).Tile(x, y).AttData(0), 6) Then Exit Function
                 End Select
            End Select
        End If
'printcrashdebug 6, 3
        Select Case FromDir
            Case 1
                If ExamineBit(map(mapNum).Tile(x, y).WallTile, 0) Then Exit Function
                If ExamineBit(map(mapNum).Tile(x, y - 1).WallTile, 5) Then Exit Function
                If map(mapNum).Tile(x, y - 1).Att = 17 Then If ExamineBit(map(mapNum).Tile(x, y - 1).AttData(0), 3) Then Exit Function
                
                If mNum >= 0 And x <> 11 And y <> 11 Then
                    If monster(map(mapNum).monster(mNum).monster).Flags2 And MONSTER_LARGE Then
                        If ExamineBit(map(mapNum).Tile(x, y + 1).WallTile, 0) Then Exit Function
                        If ExamineBit(map(mapNum).Tile(x, y).WallTile, 5) Then Exit Function
                        If map(mapNum).Tile(x, y).Att = 17 Then If ExamineBit(map(mapNum).Tile(x, y).AttData(0), 3) Then Exit Function
                        If ExamineBit(map(mapNum).Tile(x + 1, y + 1).WallTile, 0) Then Exit Function
                        If ExamineBit(map(mapNum).Tile(x + 1, y).WallTile, 5) Then Exit Function
                        If map(mapNum).Tile(x + 1, y).Att = 17 Then If ExamineBit(map(mapNum).Tile(x + 1, y).AttData(0), 3) Then Exit Function
                    End If
                End If
            Case 0
                If ExamineBit(map(mapNum).Tile(x, y).WallTile, 1) Then Exit Function
                If ExamineBit(map(mapNum).Tile(x, y + 1).WallTile, 4) Then Exit Function
                If map(mapNum).Tile(x, y + 1).Att = 17 Then If ExamineBit(map(mapNum).Tile(x, y + 1).AttData(0), 1) Then Exit Function
                
                If mNum >= 0 And y <> 11 And x <> 11 Then
                    If monster(map(mapNum).monster(mNum).monster).Flags2 And MONSTER_LARGE Then
                        If ExamineBit(map(mapNum).Tile(x + 1, y).WallTile, 1) Then Exit Function
                        If ExamineBit(map(mapNum).Tile(x + 1, y + 1).WallTile, 4) Then Exit Function
                        If map(mapNum).Tile(x + 1, y + 1).Att = 17 Then If ExamineBit(map(mapNum).Tile(x + 1, y + 1).AttData(0), 1) Then Exit Function
                    End If
                End If
            Case 3
                If ExamineBit(map(mapNum).Tile(x, y).WallTile, 2) Then Exit Function
                If ExamineBit(map(mapNum).Tile(x - 1, y).WallTile, 7) Then Exit Function
                If map(mapNum).Tile(x - 1, y).Att = 17 Then If ExamineBit(map(mapNum).Tile(x - 1, y).AttData(0), 7) Then Exit Function
                                
                If mNum >= 0 And y <> 11 And x <> 11 Then
                    If monster(map(mapNum).monster(mNum).monster).Flags2 And MONSTER_LARGE Then
                        If ExamineBit(map(mapNum).Tile(x + 1, y).WallTile, 2) Then Exit Function
                        If ExamineBit(map(mapNum).Tile(x, y).WallTile, 7) Then Exit Function
                        If map(mapNum).Tile(x, y).Att = 17 Then If ExamineBit(map(mapNum).Tile(x, y).AttData(0), 7) Then Exit Function
                        If ExamineBit(map(mapNum).Tile(x + 1, y + 1).WallTile, 2) Then Exit Function
                        If ExamineBit(map(mapNum).Tile(x, y + 1).WallTile, 7) Then Exit Function
                        If map(mapNum).Tile(x, y + 1).Att = 17 Then If ExamineBit(map(mapNum).Tile(x, y + 1).AttData(0), 7) Then Exit Function
                    End If
                End If
            Case 2
                If ExamineBit(map(mapNum).Tile(x, y).WallTile, 3) Then Exit Function
                If ExamineBit(map(mapNum).Tile(x + 1, y).WallTile, 6) Then Exit Function
                If map(mapNum).Tile(x + 1, y).Att = 17 Then If ExamineBit(map(mapNum).Tile(x + 1, y).AttData(0), 5) Then Exit Function
                
                If mNum >= 0 And y <> 11 And x <> 11 Then
                    If monster(map(mapNum).monster(mNum).monster).Flags2 And MONSTER_LARGE Then
                        If ExamineBit(map(mapNum).Tile(x, y + 1).WallTile, 3) Then Exit Function
                        If ExamineBit(map(mapNum).Tile(x + 1, y + 1).WallTile, 6) Then Exit Function
                        If map(mapNum).Tile(x + 1, y + 1).Att = 17 Then If ExamineBit(map(mapNum).Tile(x + 1, y + 1).AttData(0), 5) Then Exit Function
                    End If
                End If
        End Select
        'printcrashdebug 6, 4
        For A = 0 To 9
            With .monster(A)
                If A <> mNum Then
                    If .monster > 0 Then
                        If .x = x And .y = y Then
                            Exit Function
                        End If
                        If monster(.monster).Flags2 And MONSTER_LARGE Then
                            If .x + 1 = x And .y + 1 = y Then
                                Exit Function
                            End If
                            If .x + 1 = x And .y = y Then
                                Exit Function
                            End If
                            If .x = x And .y + 1 = y Then
                                Exit Function
                            End If
                        End If
                        If mNum >= 0 Then
                            If monster(map(mapNum).monster(mNum).monster).Flags2 Then
                                If .x = x + 1 And .y = y Then
                                    Exit Function
                                End If
                                If .x = x + 1 And .y = y + 1 Then
                                    Exit Function
                                End If
                                If .x = x And .y = y + 1 Then
                                    Exit Function
                                End If
                                'If monster(map(MapNum).monster(a).monster).Flags2 And MONSTER_LARGE Then
                                
                                
                                'End If
                            End If
                        End If
                    End If
                End If
            End With
        Next A
        'printcrashdebug 6, 5
        For A = 1 To currentMaxUser
            With player(A)
                If .map = mapNum Then
                    If mNum >= 0 Then
                        If monster(map(mapNum).monster(mNum).monster).Flags2 And MONSTER_LARGE Then
                            If .x = x And .y = y Then
                                IsVacant = A + 1
                                Exit Function
                            End If
                            If .x = x + 1 And .y = y + 1 Then
                                IsVacant = A + 1
                                Exit Function
                            End If
                            If .x = x + 1 And .y = y Then
                                IsVacant = A + 1
                                Exit Function
                            End If
                            If .x = x And .y = y + 1 Then
                                IsVacant = A + 1
                                Exit Function
                            End If
                        Else
                            If .x = x And .y = y Then
                                IsVacant = A + 1
                                Exit Function
                            End If
                        End If
                    Else
                        If .x = x And .y = y Then
                            IsVacant = A + 1
                            Exit Function
                        End If
                    End If
                End If
            End With
        Next A
    End With
    IsVacant = 1
        'printcrashdebug 6, 20
End Function

'Function FindFreeTile(x As Long, y As Long, mapnum As Long, lastdir As Long)
'    FindFreeTile = 0
'    If x > 11 Or x < 0 Or y > 11 Or y < 0 Then Exit Function '

'    If map(mapnum).Tile(x, y).Att <> 1 And map(mapnum).Tile(x, y).Att <> 2 And map(mapnum).Tile(x, y).Att <> 3 And map(mapnum).Tile(x, y).Att <> 13 And map(mapnum).Tile(x, y).Att <> 24 Then
'        FindFreeTile = x * 100 + y
'        Exit Function
'    End If

'    If lastdir <> 1 Then
'        FindFreeTile = FindFreeTile(x, y - 1, map, 0)
'        If FindFreeTile > 0 Then Exit Function
'    End If
'
'
'End Function

Sub JoinGame(Index As Long)
    Dim A As Long, ST1 As String, B As Long
    
    With player(Index)

        .Mode = modePlaying
        If .Trading = True Then CloseTrade Index
        For A = 1 To 10
            .Trade.Slot(A) = 0
        Next A
               
        
        
        .Trade.State = 0
        .Trade.Trader = 0
        EquateLight Index
        SendAllBut Index, Chr2(6) + Chr2(Index) + Chr2(.sprite) + Chr2(.Status) + Chr2(.Guild) + Chr2(.Light.Intensity) + Chr2(.Light.Radius) + .Name
        If .Guild > 0 Then UpdateGuildInfo CByte(.Guild)
        ST1 = DoubleChar(1) + Chr2(24)
        
        .trapID = Int(10000 * Rnd)
        
        A = .map
        If map(A).BootLocation.map > 0 Then
            'Move player if not allowed to join on this map
            .map = map(A).BootLocation.map
            .x = map(A).BootLocation.x
            .y = map(A).BootLocation.y
        Else
            If map(.map).Tile(.x, .y).Att = 24 Then 'move off of trees
                .map = GetPlayerFlag(Index, 240)
                .x = GetPlayerFlag(Index, 241)
                .y = GetPlayerFlag(Index, 242)
            End If
        End If
        If .map < 1 Then .map = 1
        If .map > 5000 Then .map = 5000
        If .x > 11 Then .x = 11
        If .y > 11 Then .y = 11
        
        'Send Player Data
        For A = 1 To currentMaxUser
            If A <> Index Then
                With player(A)
                    If .Mode = modePlaying Then
                        ST1 = ST1 + DoubleChar(7 + Len(.Name)) + Chr2(6) + Chr2(A) + Chr2(.sprite) + Chr2(.Status) + Chr2(.Guild) + Chr2(.Light.Intensity) + Chr2(.Light.Radius) + .Name
                        If Len(ST1) > 1024 Then
                            SendRaw Index, ST1
                            ST1 = ""
                        End If
                    End If
                End With
            End If
        Next A
        
        
        
        'Send Inventory Data
        For A = 1 To 20
            If .Inv(A).Object > 0 Then
                ST1 = ST1 + DoubleChar$(15) + Chr2(17) + Chr2(A) + DoubleChar(.Inv(A).Object) + QuadChar(.Inv(A).Value) + Chr2(.Inv(A).prefix) + Chr2(.Inv(A).prefixVal) + Chr2(.Inv(A).suffix) + Chr2(.Inv(A).SuffixVal) + Chr2(.Inv(A).Affix) + Chr2(.Inv(A).AffixVal) + Chr2(.Inv(A).ObjectColor)
                
                If Len(ST1) > 1024 Then
                    SendRaw Index, ST1
                    ST1 = ""
                End If
            End If
        Next A
        
        For A = 1 To 5
            With .Equipped(A)
                If .Object > 0 Then
                    ST1 = ST1 + DoubleChar$(15) + Chr2(17) + Chr2(A + 20) + DoubleChar(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)
                    
                    If Len(ST1) > 1024 Then
                        SendRaw Index, ST1
                        ST1 = ""
                    End If
                End If
            End With
        Next A

        ST1 = ST1 + DoubleChar(2) + Chr2(110) + Chr2(World.Hour)
        If Len(ST1) > 0 Then
            SendRaw Index, ST1
        End If
        


        JoinMap Index

        Parameter(0) = Index
        RunScript ("JOINGAME")

  
  
        
        
        If .Guild > 0 Then SendSocket Index, Chr2(135) + Chr2(15) + Guild(.Guild).MOTD
        
        'Send Guild Data
        If .Guild > 0 Then
            ST1 = ""
            With Guild(.Guild)
                For A = 0 To 4
                    With .Declaration(A)
                        ST1 = ST1 + DoubleChar(4) + Chr2(71) + Chr2(A) + Chr2(.Guild) + Chr2(.Type)
                    End With
                Next A
                
                    If .Bank >= 0 Then
                        ST1 = ST1 + DoubleChar(5) + Chr2(74) + QuadChar(.Bank)
                    Else
                        ST1 = ST1 + DoubleChar(9) + Chr2(74) + QuadChar(Abs(.Bank)) + QuadChar(.DueDate)
                    End If
            End With
            If Len(ST1) > 0 Then
                SendRaw Index, ST1
            End If
        End If
            player(Index).MaxHP = player(Index).MaxHP
            
           CalculateStats Index
           player(Index).MaxHP = player(Index).MaxHP
            
    End With
    
Exit Sub
Error_Handler:
Open App.Path + "/LOG.TXT" For Append As #1
    ST1 = ""
    Print #1, "JOINGAME" & player(Index).Name & "/" & Err.Number & "/" & Err.Description & "/" & ST1 & "/" & "  modServer" & "  modServer"
Close #1
Unhook
End
End Sub
Sub SendToGuild(GuildNum As Long, St As String)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode = modePlaying And .Guild = GuildNum Then
                .St = .St + DoubleChar(Len(St)) + St
                If Not .DeferSends Then
                    If SendData(.Socket, .St) = SOCKET_ERROR Then
                        'CloseClientSocket Index
                    End If
                    .St = ""
                End If
            End If
        End With
    Next A
End Sub
Sub SendToGuildAllBut(Index As Long, GuildNum As Long, St As String)
    If player(Index).Guild > 0 Then
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode = modePlaying And .Guild = GuildNum And Index <> A Then
                .St = .St + DoubleChar(Len(St)) + St
                If Not .DeferSends Then
                    If SendData(.Socket, .St) = SOCKET_ERROR Then
                        'CloseClientSocket Index
                    End If
                    .St = ""
                End If
            End If
        End With
    Next A
    End If
End Sub

Sub SendToGuildAllBut2(Index As Long, GuildNum As Long, St As String)
    If player(Index).Guild > 0 Then
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode = modePlaying And .Guild = GuildNum And Index <> A Then
                .St = .St + DoubleChar(Len(St)) + St
            End If
        End With
    Next A
    End If
End Sub

Sub JoinMap(Index As Long, Optional runscripts As Boolean = True)
    Dim A As Long, B As Long, C As Integer, D As Integer, E As Integer, F As Integer, G As Integer, H As Integer, mapNum As Long, ST1 As String
    
    With player(Index)
        ClearMoveQueue (Index)
        .DeferSends = True
        mapNum = .map
        If .Trading = True Then
            CloseTrade Index
        End If
        If map(.map).NumPlayers = 0 Then
            B = 0
            For A = 1 To MaxUsers
                If CurrentMaps(A) = 0 Then
                    CurrentMaps(A) = mapNum
                    Exit For
                End If
            Next A
            For A = 0 To 49
                With map(mapNum).Object(A)
                    If .Object > 0 Then
                        If map(mapNum).Tile(.x, .y).Att <> 5 And .TimeStamp > 0 Then
                            B = B + 1
                            If .TimeStamp < GetTickCount Then
                                B = B - 1
                                .Object = 0
                                .TimeStamp = 0
                            End If
                        End If
                    End If
                End With
            Next A
            If B = 0 Then
                If map(mapNum).ResetTimer > 0 Then
                    
                    If GetTickCount > map(mapNum).ResetTimer + World.MapResetTime Then
                        Dim x As Long, y As Long
                        For y = 0 To 11
                            For x = 0 To 11
                                With map(mapNum).Tile(x, y)
                                    If .Att = 24 Then
                                        If GetTickCount - .TimeStamp > 60000 Then
                                            If .AttData(3) < .AttData(1) Then
                                                If .AttData(3) + (GetTickCount - .TimeStamp) / 59999 > .AttData(1) Then
                                                    .AttData(3) = .AttData(1)
                                                Else
                                                    .AttData(3) = .AttData(3) + (GetTickCount - .TimeStamp) / 59999
                                                End If
                                            End If
                                        End If
                                        .TimeStamp = GetTickCount
                                    End If
                                End With
                            Next x
                        Next y
                    End If
                    If GetTickCount > map(mapNum).ResetTimer + World.MapResetTime Then
                        SoftResetMap mapNum, False
                    End If
                End If
            End If
        End If
        
                                   
        
        If .WalkCode < 255 Then
            .WalkCode = .WalkCode + 1
        Else
            .WalkCode = 0
        End If
    
        With map(mapNum)
            .NumPlayers = .NumPlayers + 1
        End With
        ST1 = DoubleChar(15) + Chr2(12) + DoubleChar(CLng(mapNum)) + Chr2(.x) + Chr2(.y) + Chr2(.D) + Chr2(.WalkCode) + QuadChar(map(mapNum).Version) + QuadChar(map(mapNum).CheckSum)

        'Send Door Data
        For A = 0 To 9
            With map(mapNum).Door(A)
                If .Att > 0 Or .Wall > 0 Then
                    ST1 = ST1 + DoubleChar(4) + Chr2(36) + Chr2(A) + Chr2(.x) + Chr2(.y)
                End If
            End With
        Next A
        
        'Send Player Data
        For A = 1 To currentMaxUser
            If player(A).Mode = modePlaying And player(A).map = mapNum And A <> Index Then
                With player(A)
                    If (.alpha <> 0 Or .red <> 0 Or .blue <> 0 Or .green <> 0) Then
                        ST1 = ST1 + DoubleChar(17) + Chr2(8) + Chr2(A) + Chr2(.x) + Chr2(.y) + Chr2(.D)
                        ST1 = ST1 + Chr2(.red) + Chr2(.green) + Chr2(.blue) + Chr2(.alpha)
                    Else
                        ST1 = ST1 + DoubleChar(13) + Chr2(8) + Chr2(A) + Chr2(.x) + Chr2(.y) + Chr2(.D)
                    End If
                        If .Equipped(1).Object > 0 Then
                            ST1 = ST1 + DoubleChar(.Equipped(1).Object)
                        Else
                            ST1 = ST1 + DoubleChar(0)
                        End If
                        If .Equipped(2).Object > 0 Then
                            ST1 = ST1 + DoubleChar(.Equipped(2).Object)
                        Else
                            ST1 = ST1 + DoubleChar(0)
                        End If
                        If .Equipped(3).Object > 0 Then
                            ST1 = ST1 + DoubleChar(.Equipped(3).Object)
                        Else
                            ST1 = ST1 + DoubleChar(0)
                        End If
                        If .Equipped(4).Object > 0 Then
                            ST1 = ST1 + DoubleChar(.Equipped(4).Object)
                        Else
                            ST1 = ST1 + DoubleChar(0)
                        End If
                    
                    ST1 = ST1 + DoubleChar(6) + Chr2(112) + Chr2(A) + QuadChar(.StatusEffect)
                    ST1 = ST1 + DoubleChar(3) + Chr2(115) + Chr2(A) + Chr2(.Buff.Type)
                End With
                If Len(ST1) > 1024 Then
                    SendRaw Index, ST1
                    ST1 = ""
                End If
            End If
        Next A
        
        With map(mapNum)
            'Send Map Monster Data
            For A = 0 To 9
                With .monster(A)
                    If .monster > 0 Then
                        If .R <> 0 Or .G <> 0 Or .B <> 0 Or .A <> 0 Then
                            ST1 = ST1 + DoubleChar(13) + Chr2(38) + Chr2(A) + DoubleChar(.monster) + Chr2(.x) + Chr2(.y) + Chr2(.D) + DoubleChar(CLng(.HP)) + Chr2(.R) + Chr2(.G) + Chr2(.B) + Chr2(.A)
                        Else
                            ST1 = ST1 + DoubleChar(9) + Chr2(38) + Chr2(A) + DoubleChar(.monster) + Chr2(.x) + Chr2(.y) + Chr2(.D) + DoubleChar(CLng(.HP))
                        End If
                    End If
                End With
            Next A
            
            'Send Map Object Data
            For A = 0 To 49
                With .Object(A)
                    If .Object > 0 Then
                        If .prefix > 0 And .prefix < 256 Then
                            B = prefix(.prefix).Light.Radius
                            C = prefix(.prefix).Light.Intensity
                        End If
                        If .suffix > 0 And .suffix < 256 Then
                            B = B + prefix(.suffix).Light.Radius
                            C = C + prefix(.suffix).Light.Intensity
                        End If
                        If .Affix > 0 And .Affix < 256 Then
                            B = B + prefix(.Affix).Light.Radius
                            C = C + prefix(.Affix).Light.Intensity
                        End If
                        If B > 255 Then B = 255
                        If C > 255 Then C = 255
                        ST1 = ST1 + DoubleChar(24) + Chr2(14) + Chr2(A) + DoubleChar(.Object) + Chr2(.x) + Chr2(.y) + Chr2(B) + Chr(C) + QuadChar$(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + QuadChar(IIf(.TimeStamp - GetTickCount > 0, .TimeStamp - GetTickCount, 0)) + Chr2(Abs(.deathObj)) + Chr2(.ObjectColor)
                    End If
                    If Len(ST1) > 1024 Then
                        SendRaw Index, ST1
                        ST1 = ""
                    End If
                End With
            Next A
            ST1 = ST1 + DoubleChar(9) + Chr2(120) + Chr2(0) + DoubleChar$(.Raining.Raining) + DoubleChar$(.Snowing) + Chr2(.Fog) + Chr2(.FlickerDark) + Chr2(.FlickerLength)
        End With
        
        'St1 = St1 + DoubleChar(1) + chr2(22) 'End of map data
        
'        A = Map(MapNum).NPC
'        If A >= 1 Then
'            With NPC(A)
'                If .JoinText <> "" Then
'                    St1 = St1 + DoubleChar(2 + Len(.JoinText)) + chr2(88) + chr2(A) + .JoinText
'                End If
'            End With
'        End If
        
        If ST1 <> "" Then
            SendRaw Index, ST1
        End If
        SendToMapAllBut mapNum, Index, Chr2(112) + Chr2(Index) + QuadChar(CLng(.StatusEffect))
        
            If .Equipped(1).Object > 0 Then
                A = .Equipped(1).Object
            Else
                A = 0
            End If
            If .Equipped(2).Object > 0 Then
                B = .Equipped(2).Object
            Else
                B = 0
            End If
            If .Equipped(3).Object > 0 Then
                C = .Equipped(3).Object
            Else
                C = 0
            End If
            If .Equipped(4).Object > 0 Then
                D = .Equipped(4).Object
            Else
                D = 0
            End If
        
        If (.alpha <> 0 Or .red <> 0 Or .green <> 0 Or .blue <> 0) Then
            SendToMapAllBut mapNum, Index, Chr2(8) + Chr2(Index) + Chr2(.x) + Chr2(.y) + Chr2(.D) + Chr2(.red) + Chr2(.green) + Chr2(.blue) + Chr2(.alpha) + DoubleChar(A) + DoubleChar(B) + DoubleChar(C) + DoubleChar(D)
        Else
            SendToMapAllBut mapNum, Index, Chr2(8) + Chr2(Index) + Chr2(.x) + Chr2(.y) + Chr2(.D) + DoubleChar(A) + DoubleChar(B) + DoubleChar(C) + DoubleChar(D)
        End If
        If .Buff.Type > 0 Then SendToMapAllBut mapNum, Index, Chr2(115) + Chr2(Index) + Chr2(.Buff.Type)
        If .Party > 0 Then
            'Send Location Information
        End If
        
        If .Access = 0 Then
            For C = 0 To 9
                If map(mapNum).monster(C).monster > 0 Then
                  If monster(map(mapNum).monster(C).monster).Sight = 255 Then
                    If ExamineBit(monster(map(mapNum).monster(C).monster).Flags, 3) = False Then
                        'Isn't Friendly
                         If ExamineBit(monster(map(mapNum).monster(C).monster).Flags, 0) = False Or .Status = 1 Then
                            With map(mapNum).monster(C)
                                E = .x
                                F = .y
                                G = .Distance
                            End With
                            H = Sqr((CLng(.x) - E) ^ 2 + (CLng(.y) - F) ^ 2)
                            If H <= G Then
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
            Next C
        End If
        SendToPartyAllBut .Party, Index, Chr2(104) + Chr2(2) + Chr2(Index) + DoubleChar(.map) ' + map(.map).name
        SendToGuildAllBut Index, CLng(.Guild), Chr2(104) + Chr2(2) + Chr2(Index) + DoubleChar(.map) ' + map(.map).name
        Parameter(0) = Index
        Parameter(1) = player(Index).map
        If runscripts Then
            RunScript "JOINMAP" + CStr(mapNum)
            RunScript "JOINMAP"
        End If
        SendSocket Index, Chr2(22)
        
        For A = 0 To MaxTraps
            With map(mapNum).trap(A)
                If player(Index).x = .x And player(Index).y = .y Then
                    PrintDebug player(Index).user & "(" & player(Index).Name & ")" & " triggered a trap."
                    PlayerTriggerTrap Index, A
                End If
                If .ActiveCounter - 50 >= GetTickCount Then
                    SendSocket Index, Chr2(137) + Chr2(.x) + Chr2(.y) + Chr2(A) + DoubleChar((.ActiveCounter - 50) - GetTickCount)
                End If
            End With
        Next A

        ST1 = ""
        'If ExamineBit(.Flags(1), 5) Then 'check for exhausted trees
            With map(mapNum)
                For A = 0 To 11
                    For B = 0 To 11
                        With .Tile(A, B)
                            If .Att = 24 Then
                                If .AttData(1) > 0 Then
                                    If .AttData(3) = 0 Then
                                        ST1 = ST1 + DoubleChar(3) + Chr2(140) + Chr2(A) + Chr2(B)
                                    End If
                                End If
                            End If
                        End With
                    Next B
                Next A
                If ST1 <> "" Then SendRaw Index, ST1
            End With
        'End If
        .DeferSends = False
        FlushSocket (Index)
    End With
End Sub
Sub LoadMap(mapNum As Long, MapData As String)
    Dim A As Long, x As Long, y As Long
    If Len(MapData) >= 2374 Then
        'Characters 1-30 = Name
        '36 = Midi
        With map(mapNum)
            .Name = ClipString$(Mid$(MapData, 1, 30))
            .CheckSum = CheckSum(MapData)
            .Version = Asc(Mid$(MapData, 31, 1)) * 16777216 + Asc(Mid$(MapData, 32, 1)) * 65536 + Asc(Mid$(MapData, 33, 1)) * 256& + Asc(Mid$(MapData, 34, 1))
            .NPC = Asc(Mid$(MapData, 35, 1))
            .ExitUp = Asc(Mid$(MapData, 37, 1)) * 256 + Asc(Mid$(MapData, 38, 1))
            .ExitDown = Asc(Mid$(MapData, 39, 1)) * 256 + Asc(Mid$(MapData, 40, 1))
            .ExitLeft = Asc(Mid$(MapData, 41, 1)) * 256 + Asc(Mid$(MapData, 42, 1))
            .ExitRight = Asc(Mid$(MapData, 43, 1)) * 256 + Asc(Mid$(MapData, 44, 1))
            .BootLocation.map = Asc(Mid$(MapData, 45, 1)) * 256 + Asc(Mid$(MapData, 46, 1))
            .BootLocation.x = Asc(Mid$(MapData, 47, 1))
            .BootLocation.y = Asc(Mid$(MapData, 48, 1))
            .Flags(0) = Asc(Mid$(MapData, 49, 1))
            .Flags(1) = Asc(Mid$(MapData, 51, 1))
            .Zone = Asc(Mid$(MapData, 56, 1))
            For A = 0 To 4 '61 - 70
                .MonsterSpawn(A).monster = Asc(Mid$(MapData, 61 + A * 2))
                .MonsterSpawn(A).Rate = Asc(Mid$(MapData, 62 + A * 2))
            Next A
            '56
            .Keep = False
            For y = 0 To 11
                For x = 0 To 11
                    With .Tile(x, y)
                        A = 71 + y * 192 + x * 16
                        .Ground = Asc(Mid$(MapData, A, 1)) * 256 + Asc(Mid$(MapData, A + 1, 1))
                        .Att = Asc(Mid$(MapData, A + 10, 1))
                        .AttData(0) = Asc(Mid$(MapData, A + 11, 1))
                        .AttData(1) = Asc(Mid$(MapData, A + 12, 1))
                        .AttData(2) = Asc(Mid$(MapData, A + 13, 1))
                        .AttData(3) = Asc(Mid$(MapData, A + 14, 1))
                        .WallTile = Asc(Mid$(MapData, A + 15, 1))
                        Select Case .Att
                            Case 5
                                map(mapNum).Keep = True
                            Case 8
                                If .AttData(2) > 0 Then
                                    map(mapNum).Hall = .AttData(2)
                                End If
                            Case 17
                                .WallTile = Asc(Mid$(MapData, A + 11, 1))
                            Case 1
                                .WallTile = 15
                        End Select
                    End With
                Next x
            Next y
            If Len(MapData) = 2379 Then
                For A = 0 To 4 '61 - 70
                    .MonsterSpawn(A).monster = .MonsterSpawn(A).monster + Asc(Mid$(MapData, A + 2375)) * 256
                Next A
            End If
        End With
    End If
End Sub
Sub Main()
    Dim bDataRSError As Boolean
    Randomize timer
    Dim A As Long, CurDate As Long
    Dim St As String
    Dim LingerType As LingerType
    
    InitFunctionTable
    currentMaxUser = 0
    CurDate = CLng(Date)
    frmLoading.Show
    frmLoading.Refresh
    
    For A = 0 To 255
        Chr2(A) = Chr$(A)
    Next A
    
    If Exists("./debug2.log") Then Kill "./debug2.log"
    If Exists("./debug.log") Then Name "./debug.log" As "./debug2.log"
    
    If Exists("server.tmp") Then
        MsgBox "The server will not run when a server.tmp exists, for your own safety.", vbOKOnly
        End
    End If
    
    'Set System Password
    ServerAdminPass = ReadString("Server", "Settings", "ServerAdminPass")
    
    ChDir (App.Path)
    Set WS = DBEngine.Workspaces(0)
    If Exists("server.dat") Then
        frmLoading.lblStatus = "Opening Server Database.."
        frmLoading.lblStatus.Refresh

        Name "server.dat" As "server.tmp"
        CompactDatabase "server.tmp", "server.dat", , 0, ";pwd=" + Chr2(100) + Chr2(114) + Chr2(97) + Chr2(99) + Chr2(111)
        'mode=16
        'mode=adModeShareDenyNone
        'exclusive=no
        
        Kill "server.tmp"
        Set db = WS.OpenDatabase("server.dat", False, False, ";exclusive=Yes;pwd=" + Chr2(100) + Chr2(114) + Chr2(97) + Chr2(99) + Chr2(111))
    Else
        frmLoading.lblStatus = "Creating Server Database.."
        frmLoading.lblStatus.Refresh
        CreateDatabase
    End If
    'Leaving this here in case I need to update the DB again
    If InStr(Command$, "-UPDATE") Then
        UpdateDataBase
        MsgBox "Database Updated"
        End
    End If

    InitDB


        'a = GetTickCount
        ''CopyToWebDB
        'copyDb
        'PrintLog GetTickCount - a
        
        
    A = ReadInt("Server", "Settings", "Port")
    If A = 0 Then
        WriteString "Server", "Settings", "Port", "3017"
        WriteString "Server", "Settings", "Name", "Seyerdin Online Server"
    End If
    
    TitleString = ReadString("Server", "Settings", "Name")


    frmLoading.lblStatus = "Initializing Sockets.."
    frmLoading.lblStatus.Refresh
    
    frmMain.Show
    frmMain.Caption = TitleString + " [0]"
    Hook
    StartWinsock St
    
    'Listen for connections
    With LingerType
        .l_onoff = 1
        .l_linger = 0
    End With
            
    ListeningSocket = ListenForConnect(ReadInt("Server", "Settings", "Port"), gHW, 1025)
    If ListeningSocket = INVALID_SOCKET Then
        MsgBox "Unable to create listening socket!", vbOKOnly + vbExclamation, TitleString
        EndWinsock
        Unhook
        End
    End If
    If setsockopt(ListeningSocket, SOL_SOCKET, SO_LINGER, LingerType, LenB(LingerType)) Then
        MsgBox "Unable to create listening socket!", vbOKOnly + vbExclamation, TitleString
        EndWinsock
        Unhook
        End
    End If
    If setsockopt(ListeningSocket, IPPROTO_TCP, TCP_NODELAY, 1&, 1) <> 0 Then
        MsgBox "Unable to create listening socket!", vbOKOnly + vbExclamation, TitleString
        EndWinsock
        Unhook
        End
    End If


    Unload frmLoading
    
    InitSkills
    InitConstants
    
    GetBonusPerStat 38, HPPerConstitution
    
    GenerateEXPLevels
    frmMain.Show
    
    SetTimer gHW, 1, 10000, AddressOf ObjectTimer
    SetTimer gHW, 2, 2000, AddressOf PlayerTimer
    SetTimer gHW, 3, 60000, AddressOf MinuteTimer
    SetTimer gHW, 4, 1000, AddressOf SocketQueueTimer
    SetTimer gHW, 5, 190, AddressOf MapTimer



    PrintLog ("Seyerdin Online Server Version A" + CurrentClientVer + ".")
        
    'HashPasswords
        
    RunScript "FLAGS"
     
    If Command <> "-s" Then
        A = GetTickCount
        PrintLog "Running Startup Script..."
        RunScript "STARTUP"
        PrintLog "...Done (" & (GetTickCount - A) & " ms)"
    End If

Exit Sub
DataRSError:
    bDataRSError = True
    Resume Next
End Sub

Sub HashPasswords()
Dim user As String
Dim password As String

If Not UserRS.BOF Then
    UserRS.MoveFirst
    While Not UserRS.EOF
        If Not IsNull(UserRS!password) Then
            If Len(UserRS!password) <> 64 Then
                user = UserRS!user
                password = Cryp(UserRS!password)
                UserRS.Edit
                UserRS!password = Cryp(SHA256(UCase(user) & UCase(password)))
                UserRS.Update
            End If
        End If
        UserRS.MoveNext
    Wend
End If
End Sub

Function NewMapMonster(mapNum As Long, MonsterNum As Long) As String
    Dim tX As Long, tY As Long, TriesLeft As Long
    Dim MonsterType As Long, MonsterFlags As Byte
    Dim A As Long
    Randomize
    If Int(MonsterNum / 2) * 2 = MonsterNum Or ExamineBit(map(mapNum).Flags(0), 4) = True Then
        MonsterType = map(mapNum).MonsterSpawn(Int(MonsterNum / 2)).monster
        If MonsterType > 0 Then
            MonsterFlags = monster(MonsterType).Flags
            If (ExamineBit(MonsterFlags, 1) = False Or blnNight = False) And (ExamineBit(MonsterFlags, 2) = False Or blnNight = True) Then
                tX = Int(Rnd * 12)
                tY = Int(Rnd * 12)
                TriesLeft = 10
                If monster(MonsterType).Flags2 And MONSTER_LARGE Then
                    tX = Int(Rnd * 11)
                    tY = Int(Rnd * 11)
                End If

StartAgain:
                If monster(MonsterType).Flags2 And MONSTER_LARGE Then
                    While TriesLeft > 0 And ((map(mapNum).Tile(tX, tY).Att > 0 And Not map(mapNum).Tile(tX, tY).Att = 20) Or map(mapNum).Tile(tX, tY).WallTile > 0) Or ((map(mapNum).Tile(tX + 1, tY).Att > 0 And Not map(mapNum).Tile(tX + 1, tY).Att = 20) Or map(mapNum).Tile(tX + 1, tY).WallTile > 0) Or ((map(mapNum).Tile(tX + 1, tY + 1).Att > 0 And Not map(mapNum).Tile(tX + 1, tY + 1).Att = 20) Or map(mapNum).Tile(tX + 1, tY + 1).WallTile > 0) Or ((map(mapNum).Tile(tX, tY + 1).Att > 0 And Not map(mapNum).Tile(tX, tY + 1).Att = 20) Or map(mapNum).Tile(tX, tY + 1).WallTile > 0)
                        tX = Int(Rnd * 11)
                        tY = Int(Rnd * 11)
                        TriesLeft = TriesLeft - 1
                    Wend
                Else
                    While TriesLeft > 0 And ((map(mapNum).Tile(tX, tY).Att > 0 And Not map(mapNum).Tile(tX, tY).Att = 20) Or map(mapNum).Tile(tX, tY).WallTile > 0)
                        tX = Int(Rnd * 12)
                        tY = Int(Rnd * 12)
                        TriesLeft = TriesLeft - 1
                    Wend
                End If
                
                
                If TriesLeft > 0 Then
                    For A = 0 To 9
                        With map(mapNum).Door(A)
                            If ((.Att > 0) And (Not .Att = 20)) Or .Wall > 0 Then
                                If tX = .x And tY = .y Then
                                    TriesLeft = TriesLeft - 1
                                    GoTo StartAgain
                                End If
                                If monster(MonsterType).Flags2 And MONSTER_LARGE Then
                                    If tX + 1 = .x And tY = .y Then
                                        TriesLeft = TriesLeft - 1
                                        GoTo StartAgain
                                    End If
                                    If tX + 1 = .x And tY + 1 = .y Then
                                        TriesLeft = TriesLeft - 1
                                        GoTo StartAgain
                                    End If
                                    If tX = .x And tY + 1 = .y Then
                                        TriesLeft = TriesLeft - 1
                                        GoTo StartAgain
                                    End If
                                End If
                            End If
                        End With
                    Next A
                End If
                If TriesLeft > 0 Then
                    NewMapMonster = SpawnMapMonster(mapNum, MonsterNum, MonsterType, tX, tY)
                End If
            End If
        End If
    End If
End Function
Function NewMapMonster2(mapNum As Long, MonsterNum As Long, mIndex As Long) As String
    Dim tX As Long, tY As Long, TriesLeft As Long
    Dim MonsterType As Long, MonsterFlags As Byte
    Dim A As Long
    Randomize
    'If Int(MonsterNum / 2) * 2 = MonsterNum Or ExamineBit(Map(MapNum).Flags(0), 4) = True Then
        MonsterType = MonsterNum
        If MonsterType > 0 Then
            MonsterFlags = monster(MonsterType).Flags
            If (ExamineBit(MonsterFlags, 1) = False Or blnNight = False) And (ExamineBit(MonsterFlags, 2) = False Or blnNight = True) Then
                tX = Int(Rnd * 12)
                tY = Int(Rnd * 12)
                TriesLeft = 10
                If monster(MonsterType).Flags2 And MONSTER_LARGE Then
                    tX = Int(Rnd * 11)
                    tY = Int(Rnd * 11)
                End If
StartAgain:
                If monster(MonsterType).Flags2 And MONSTER_LARGE Then
                    While TriesLeft > 0 And ((map(mapNum).Tile(tX, tY).Att > 0 And Not map(mapNum).Tile(tX, tY).Att = 20) Or map(mapNum).Tile(tX, tY).WallTile > 0) Or ((map(mapNum).Tile(tX + 1, tY).Att > 0 And Not map(mapNum).Tile(tX + 1, tY).Att = 20) Or map(mapNum).Tile(tX + 1, tY).WallTile > 0) Or ((map(mapNum).Tile(tX + 1, tY + 1).Att > 0 And Not map(mapNum).Tile(tX + 1, tY + 1).Att = 20) Or map(mapNum).Tile(tX + 1, tY + 1).WallTile > 0) Or ((map(mapNum).Tile(tX, tY + 1).Att > 0 And Not map(mapNum).Tile(tX, tY + 1).Att = 20) Or map(mapNum).Tile(tX, tY + 1).WallTile > 0)
                        tX = Int(Rnd * 11)
                        tY = Int(Rnd * 11)
                        TriesLeft = TriesLeft - 1
                    Wend
                Else
                    While TriesLeft > 0 And ((map(mapNum).Tile(tX, tY).Att > 0 And Not map(mapNum).Tile(tX, tY).Att = 20) Or map(mapNum).Tile(tX, tY).WallTile > 0)
                        tX = Int(Rnd * 12)
                        tY = Int(Rnd * 12)
                        TriesLeft = TriesLeft - 1
                    Wend
                End If
                
                
                If TriesLeft > 0 Then
                    For A = 0 To 9
                        With map(mapNum).Door(A)
                            If ((.Att > 0) And (Not .Att = 20)) Or .Wall > 0 Then
                                If tX = .x And tY = .y Then
                                    TriesLeft = TriesLeft - 1
                                    GoTo StartAgain
                                End If
                                If monster(MonsterType).Flags2 And MONSTER_LARGE Then
                                    If tX + 1 = .x And tY = .y Then
                                        TriesLeft = TriesLeft - 1
                                        GoTo StartAgain
                                    End If
                                    If tX + 1 = .x And tY + 1 = .y Then
                                        TriesLeft = TriesLeft - 1
                                        GoTo StartAgain
                                    End If
                                    If tX = .x And tY + 1 = .y Then
                                        TriesLeft = TriesLeft - 1
                                        GoTo StartAgain
                                    End If
                                End If
                            End If
                        End With
                    Next A
                End If
                If TriesLeft > 0 Then
                    NewMapMonster2 = SpawnMapMonster(mapNum, mIndex, MonsterType, tX, tY)
                End If
            End If
        End If
    'End If
End Function

Function NewMapObject(mapNum As Long, ObjectNum As Long, Value As Long, x As Long, y As Long, Infinite As Boolean, MagicChance As Long, Optional durabilityPercent As Long = 100, Optional flag0 As Long = 0, Optional flag1 As Long = 0, Optional flag2 As Long = 0, Optional flag3 As Long = 0, Optional ObjectColor As Long = 0, Optional prefixNum As Long = 0, Optional prefixVal As Long = 0, Optional suffixNum As Long = 0, Optional SuffixVal As Long = 0, Optional AffixNum As Long = 0, Optional AffixVal As Long = 0) As Long
Randomize
    If ObjectNum = 0 Then Exit Function
    Dim A As Long, B As Long, C As Long, D As Long, Mul As Single
    If mapNum >= 1 Then
        If Object(ObjectNum).Type = 6 Or Object(ObjectNum).Type = 11 Then 'its money
            'search for other moneys of the same type
            For A = 0 To 49
                If map(mapNum).Object(A).Object = ObjectNum Then
                    If x = map(mapNum).Object(A).x And y = map(mapNum).Object(A).y Then
                        Exit For
                    End If
                End If
            Next A
            If A = 50 Then A = FreeMapObj(mapNum)
        Else
            A = FreeMapObj(mapNum)
        End If
        If A >= 0 Then
            With map(mapNum).Object(A)
                .x = x
                .y = y
                .prefix = prefixNum
                .suffix = suffixNum
                .Affix = AffixNum
                .prefixVal = prefixVal
                .SuffixVal = SuffixVal
                .AffixVal = AffixVal
                .ObjectColor = ObjectColor
                .Flags(0) = flag0
                .Flags(1) = flag1
                .Flags(2) = flag2
                .Flags(3) = flag3
                If Infinite = True Then
                    .TimeStamp = 0
                    .deathObj = False
                Else
                    .TimeStamp = GetTickCount + GLOBAL_OBJECT_RESET_RATE
                    .deathObj = False
                End If
                Select Case Object(ObjectNum).Type
                    Case 1, 2, 3, 4, 10 'Weapon, Shield, Armor, Helmut
                        .Value = (CLng(Object(ObjectNum).Data(0)) * 10 * durabilityPercent) / 100
                    Case 6, 11 'Money
                        If Value < 1 Then Value = 1
                        If map(mapNum).Object(A).Object <> 0 Then
                            .Value = .Value + Value
                        Else
                            .Value = Value
                        End If
                    Case 8 'Ring
                        .Value = (CLng(Object(ObjectNum).Data(1)) * 10 * durabilityPercent) / 100
                    Case Else
                        .Value = 0
                End Select
                .Object = ObjectNum
                If (Object(ObjectNum).Type >= 0 And Object(ObjectNum).Type <= 4) Or (Object(ObjectNum).Type = 8) Then
                    
                
                        If Int(Rnd * 100) < MagicChance Then
                            C = Int(Rnd * 20)
                            If C = 0 Then
                                Mul = 1.8
                                D = 2
                            ElseIf C > 0 And C < 4 Then
                                Mul = 1.4
                                D = 1
                            Else
                                Mul = 1
                                D = 0
                            End If
                            If Object(.Object).Level > 2 And D = 2 Then
                                B = Object(.Object).Level - Int(Rnd * 3)
                            Else
                                B = Int(Rnd * Object(.Object).Level) + 1
                            End If
                            If NumPrefix(B) > 0 Then
                                C = Int(Rnd * NumPrefix(B)) + 1
                                .prefix = SortedPrefixList(B, C)
                                If .prefix > 0 Then
                                    .prefixVal = Int(Rnd * (prefix(.prefix).Max - prefix(.prefix).Min + 1) + prefix(.prefix).Min) * Mul
                                End If
                            End If
                            .prefixVal = (.prefixVal Or (D * 64)) 'This is the "how good" flag
                        End If
                        If Int(Rnd * 100) < MagicChance Then
                            C = Int(Rnd * 20)
                            If C = 0 Then
                                Mul = 1.8
                                D = 2
                            ElseIf C > 0 And C < 4 Then
                                Mul = 1.4
                                D = 1
                            Else
                                Mul = 1
                                D = 0
                            End If
                            If Object(.Object).Level > 2 And D = 2 Then
                                B = Object(.Object).Level - Int(Rnd * 3)
                            Else
                                B = Int(Rnd * Object(.Object).Level) + 1
                            End If
                            If NumSuffix(B) > 0 Then
                                C = Int(Rnd * NumPrefix(B)) + 1
                                .suffix = SortedSuffixList(B, C)
                                If .suffix > 0 Then
                                    .SuffixVal = Int(Rnd * (prefix(.suffix).Max - prefix(.suffix).Min + 1) + prefix(.suffix).Min) * Mul
                                End If
                            End If
                            .SuffixVal = (.SuffixVal Or (D * 64)) 'This is the "how good" flag
                        End If
                        
       
                End If
                
                SendToMap mapNum, Chr2(14) + Chr2(A) + DoubleChar(CInt(ObjectNum)) + Chr2(x) + Chr2(y) + Chr2(0) + Chr2(0) + QuadChar$(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + QuadChar(IIf(.TimeStamp - GetTickCount > 0, .TimeStamp - GetTickCount, 0)) + Chr2(Abs(.deathObj)) + Chr2(.ObjectColor)
            End With
            NewMapObject = True
        End If
    End If
End Function

Sub Partmap(Index As Long, Optional LoseAggro As Boolean = True, Optional runscripts As Boolean = True)
    Dim A As Long, mapNum As Long

    With player(Index)
        If .Trading = True Then
            CloseTrade Index
        End If
        'Remove Widget Menu
        .Widgets.NumWidgets = 0
        .Widgets.MenuVisible = False
        'Destroy Projectiles
        .ProjectileType = 0
        .ProjectileDamage = 0
        .ProjectileX = 0
        .ProjectileY = 0
        mapNum = .map
        If mapNum > 0 Then
            Parameter(0) = Index
            If runscripts Then RunScript "PARTMAP" + CStr(mapNum)
            
            With map(mapNum)
                .NumPlayers = .NumPlayers - 1
                If .NumPlayers = 0 Then
                    For A = 1 To MaxUsers
                        If CurrentMaps(A) = mapNum Then
                            CurrentMaps(A) = 0
                            Exit For
                        End If
                    Next A
                    If ExamineBit(map(mapNum).Flags(0), 7) Then
                        ResetMap (mapNum)
                    End If
                End If
                If LoseAggro Then
                    For A = 0 To 9
                        With .monster(A)
                            If .Target = Index And .monster > 0 And .TargType = TargTypePlayer Then
                                .Target = 0
                                .TargType = 0
                                .Distance = monster(.monster).Sight
                            End If
                        End With
                    Next A
                End If
                If .NumPlayers = 0 Then
                    .ResetTimer = GetTickCount
                End If
            End With
            SendToMapAllBut mapNum, Index, Chr2(9) + Chr2(Index)
            '.map = 0
        End If
    End With
End Sub

Function PlayerCritDamage(Index As Long, ByVal damage As Long) As Long
    Dim A As Long
    A = GetStatPerBonus(player(Index).Agility, AgilityPerCritChance)
    'STATUS: DEADLY CLERITY
    If GetStatusEffect(Index, SE_DEADLYCLARITY) Then
        A = A + 10 + player(Index).StatusData(SE_DEADLYCLARITY).Data(0)
    End If
    'SKILL: POIGNANCY
    If Int(Rnd * 100) < (5 + player(Index).CriticalBonus + player(Index).SkillLevel(SKILL_PERCEPTION) * 10 + CInt(player(Index).SkillLevel(SKILL_POIGNANCY) / 2.5) + A) Then
        damage = damage * (1.5 + (player(Index).SkillLevel(SKILL_DESOLATION) * 0.02))
    End If
PlayerCritDamage = damage
End Function

Function PlayerMagicDamage(ByVal Index As Long, ByVal damage As Long) As Long
    Dim D As Long, A As Long
            D = 0
            Parameter(0) = Index
            If player(Index).SkillLevel(SKILL_SPELLPOWER) Then
                Parameter(0) = Index
                D = RunScript("SPELL" & SKILL_SPELLPOWER)
            End If
            
            Parameter(0) = Index
            If GetStatusEffect(Index, SE_EVANESCENCE) Then D = D + RunScript("SPELL" & SKILL_EVANESCENCE)
            
            If D Then
                damage = damage + ((damage * D) \ 100)
            End If
            

            damage = damage + player(Index).MagicDamageBonus

            
            'magic amp''''''''''''''''''''
            For A = 1 To 5
                If player(Index).Equipped(A).Object > 0 Then
                    D = Object(player(Index).Equipped(A).Object).Data(9)
                    damage = (damage * D) \ 100
                End If
            Next A

        PlayerMagicDamage = damage
End Function

Function PlayerEvasion(Target As Long) As Long
    Dim C As Long
    C = GetStatPerBonus(player(Target).Agility, AgilityPerDodgeChance) + GetStatPerBonus(player(Target).Wisdom, PietyPerDodge)
    'SKILL: EVASION
    If CanUseSkill(Target, SKILL_EVASION) Then
        C = C + player(Target).SkillLevel(SKILL_EVASION) * 1.5
    End If
    'STATUS: ETHEREALITY
    If GetStatusEffect(Target, SE_ETHEREALITY) Then
        C = C + player(Target).SkillLevel(SKILL_ETHEREALITY) * 2
    End If
    C = C + player(Target).SkillLevel(SKILL_AGILITY) * 17
    C = C + player(Target).DodgeBonus
    
    If C > 255 Then C = 255
    PlayerEvasion = C
End Function

Function PlayerMagicArmor(Index As Long, ByVal damage As Long) As Long
Dim resist As Long
Dim A As Long, Armor As Long, Modifier As Byte, B As Long

With player(Index)
    .ShieldBlock = False
    If .Equipped(EQ_SHIELD).Object > 0 Then
        With .Equipped(EQ_SHIELD)
            If Object(.Object).Type = 2 Then
                If Not ExamineBit(Object(.Object).Flags, 5) Then
                    'SKILL: SHIELD MASTERY
                    If Object(.Object).Data(3) > 0 And Int(Rnd * 255 + 1) < Object(.Object).Data(3) + player(Index).ShieldBlockBonus + player(Index).SkillLevel(SKILL_SHIELDMASTERY) + player(Index).SkillLevel(SKILL_GUARDIAN) * 5 + GetStatPerBonus(player(Index).Endurance, EndurancePerBlockChance) + GetStatPerBonus(player(Index).Wisdom, PietyPerBlock) Then
                        B = 0
                        damage = (damage * Object(.Object).Data(4)) \ 100
                        player(Index).ShieldBlock = True
                        If .prefix > 0 Then
                            If prefix(.prefix).ModType = 10 Then 'indest
                                B = 1
                            ElseIf prefix(.prefix).ModType = 17 Then 'durability
                                If Int(Rnd * (.prefixVal And 63)) > 0 Then
                                     B = 1
                                End If
                            End If
                        End If
                        If .suffix > 0 Then
                            If prefix(.suffix).ModType = 10 Then
                                B = 1
                            ElseIf prefix(.suffix).ModType = 17 Then
                                If Int(Rnd * (.SuffixVal And 63)) > 0 Then
                                    B = 1
                                End If
                            End If
                        End If
                        If ExamineBit(Object(.Object).Flags, 1) Then B = 1
                        If B = 0 Then
                            A = .Value - 1
                            If A <= 0 Then
                                SendSocket Index, Chr2(57) + Chr2(EQ_SHIELD)
                                .Object = 0
                                CalculateStats Index
                            Else
                                .Value = A
                                SendSocket Index, Chr2(119) + Chr2(20 + EQ_SHIELD) + QuadChar(.Value)
                            End If
                        End If
                        If damage = 0 Then
                            PlayerMagicArmor = 0
                            Exit Function
                        End If
                    End If
                End If
            End If
        End With
    End If


    If Rnd > 0.2 Then
        'Body Shot
        If .Equipped(EQ_ARMOR).Object > 0 Then
            'Uses Armor
            With .Equipped(EQ_ARMOR)
                Armor = Armor + Object(.Object).Data(4)
                resist = resist + Object(.Object).Data(5)
                B = 0
                If .prefix > 0 Then
                    If prefix(.prefix).ModType = 10 Then
                        B = 1
                    ElseIf prefix(.prefix).ModType = 17 Then
                        If Int(Rnd * (.prefixVal And 63)) > 0 Then
                             B = 1
                        End If
                    End If
                End If
                If .suffix > 0 Then
                    If prefix(.suffix).ModType = 10 Then
                        B = 1
                    ElseIf prefix(.suffix).ModType = 17 Then
                        If Int(Rnd * (.SuffixVal And 63)) > 0 Then
                            B = 1
                        End If
                    End If
                End If
                If ExamineBit(Object(.Object).Flags, 1) Then B = 1
                If B = 0 Then
                    A = .Value - 1
                    If A <= 0 Then
                        'Object Is Destroyed
                        SendSocket Index, Chr2(57) + Chr2(EQ_ARMOR)
                        .Object = 0
                        CalculateStats Index
                    Else
                        .Value = A
                        SendSocket Index, Chr2(119) + Chr2(20 + EQ_ARMOR) + QuadChar(.Value)
                    End If
                End If
            End With
        End If
    Else
        'Head Shot
        If .Equipped(EQ_HELMET).Object > 0 Then
            'Uses Helmet
            With .Equipped(EQ_HELMET)
                Armor = Armor + Object(.Object).Data(4)
                resist = resist + Object(.Object).Data(5)
                B = 0
                If .prefix > 0 Then
                    If prefix(.prefix).ModType = 10 Then
                        B = 1
                    ElseIf prefix(.prefix).ModType = 17 Then
                        If Int(Rnd * (.prefixVal And 63)) > 0 Then
                             B = 1
                        End If
                    End If
                End If
                If .suffix > 0 Then
                    If prefix(.suffix).ModType = 10 Then
                        B = 1
                    ElseIf prefix(.suffix).ModType = 17 Then
                        If Int(Rnd * (.SuffixVal And 63)) > 0 Then
                            B = 1
                        End If
                    End If
                End If
                If ExamineBit(Object(.Object).Flags, 1) Then B = 1
                'SKILL: CRAFTMANSHIP
                'If CanUseSkill(CByte(Index), SKILL_CRAFTSMANSHIP) Then
                '    If Int(Rnd * 20) < player(Index).SkillLevel(SKILL_CRAFTSMANSHIP) Then
                '        b = 1
                '    End If
                'End If
                If B = 0 Then
                    A = .Value - 1
                    If A <= 0 Then
                        'Object Is Destroyed
                        SendSocket Index, Chr2(57) + Chr2(EQ_HELMET)
                        .Object = 0
                        CalculateStats Index
                    Else
                        .Value = A
                        SendSocket Index, Chr2(119) + Chr2(20 + EQ_HELMET) + QuadChar(.Value)
                    End If
                End If
            End With
        End If
    End If
    
    If GetStatusEffect(Index, SE_INVULNERABILITY) Then
        damage = 0
    End If
    'SKILL: HOLY ARMOR
    Select Case .Buff.Type
        Case BUFF_HOLYARMOR
            Armor = Armor + (.Buff.Data(0))
    End Select
    

                                                                            
                                                                            
    'SKILL: Armor Mastery
    If .SkillLevel(SKILL_ARMORMASTERY) Then
        Armor = Armor + (Armor * (.SkillLevel(SKILL_ARMORMASTERY) + 5) \ 100)
    End If

    Armor = Armor + .MagicDefenseBonus


    'SKILL: BERSERK
    If GetStatusEffect(Index, SE_BERSERK) Then
        damage = damage * 1.23
    End If
    If Armor > damage Then
        damage = 0
    Else
        damage = damage - Armor
    End If

        'SKILL: SHATTERSHIELD
        If GetStatusEffect(Index, SE_SHATTERSHIELD) Then
            B = .StatusData(SE_SHATTERSHIELD).Data(0) * 256 + .StatusData(SE_SHATTERSHIELD).Data(1)
            SetPlayerMana Index, .Mana + (damage * B) / 100
            damage = damage - (damage * B) / 100
            SetPlayerMana Index, .Mana + B
            RemoveStatusEffect Index, SE_SHATTERSHIELD
        End If
        'SKILL: MANASHIELD
        If .Buff.Type = BUFF_MANASHIELD Then
            If .Mana >= (damage * .Buff.Data(0)) / 50 Then
                .Mana = .Mana - (damage * .Buff.Data(0)) / 50
                damage = (damage * (100 - .Buff.Data(0))) / 100
                If .Mana <= 0 Then
                    .Mana = 0
                    ClearBuff Index
                End If
                SendSocket Index, Chr2(48) + DoubleChar(.Mana)
            Else
                ClearBuff Index
            End If
        End If

    If GetStatusEffect(Index, SE_MASTERDEFENSE) Then
        resist = resist + 10
    End If
    
    damage = damage - ((damage * (.MagicResist + resist)) \ 100)
    If damage < 0 Then damage = 0
    PlayerMagicArmor = damage
End With
End Function

Function PlayerArmor(Index As Long, ByVal damage As Long) As Long
    Dim A As Long, Armor As Long, Modifier As Byte, B As Long, resist As Long
    With player(Index)
    
    .ShieldBlock = False
    If .Equipped(EQ_SHIELD).Object > 0 Then
        'Uses shield
        With .Equipped(EQ_SHIELD)
            If Object(.Object).Type = 2 Then
                If Not ExamineBit(Object(.Object).Flags, 5) Then
                    'SKILL: SHIELD MASTERY
                    If Object(.Object).Data(1) > 0 And Int(Rnd * 255 + 1) < Object(.Object).Data(1) + player(Index).ShieldBlockBonus + player(Index).SkillLevel(SKILL_SHIELDMASTERY) + player(Index).SkillLevel(SKILL_GUARDIAN) * 5 + GetStatPerBonus(player(Index).Endurance, EndurancePerBlockChance) + GetStatPerBonus(player(Index).Wisdom, PietyPerBlock) Then
                        B = 0
                        damage = (damage * Object(.Object).Data(2)) \ 100
                        player(Index).ShieldBlock = True
                        If .prefix > 0 Then
                            If prefix(.prefix).ModType = 10 Then 'indest
                                B = 1
                            ElseIf prefix(.prefix).ModType = 17 Then 'durability
                                If Int(Rnd * (.prefixVal And 63)) > 0 Then
                                     B = 1
                                End If
                            End If
                        End If
                        If .suffix > 0 Then
                            If prefix(.suffix).ModType = 10 Then
                                B = 1
                            ElseIf prefix(.suffix).ModType = 17 Then
                                If Int(Rnd * (.SuffixVal And 63)) > 0 Then
                                    B = 1
                                End If
                            End If
                        End If
                        If ExamineBit(Object(.Object).Flags, 1) Then B = 1
                        'SKILL: CRAFTMANSHIP
                        'If CanUseSkill(CByte(Index), SKILL_CRAFTSMANSHIP) Then
                        '    If Int(Rnd * 20) < player(Index).SkillLevel(SKILL_CRAFTSMANSHIP) Then
                        '        b = 1
                        '    End If
                        'End If
                        If B = 0 Then
                            A = .Value - 1
                            If A <= 0 Then
                                SendSocket Index, Chr2(57) + Chr2(EQ_SHIELD)
                                .Object = 0
                                CalculateStats Index
                            Else
                                .Value = A
                                SendSocket Index, Chr2(119) + Chr2(20 + EQ_SHIELD) + QuadChar(.Value)
                            End If
                        End If
                        If damage = 0 Then
                            PlayerArmor = 0
                            Exit Function
                        End If
                    End If
                End If
            End If
        End With
    End If
    
        If .Equipped(EQ_RING).Object > 0 Then
            'Has Ring
            With .Equipped(EQ_RING)
                If Object(.Object).Data(0) = 1 Then
                    Modifier = Object(.Object).Data(2)
                    B = 0
                    If .prefix > 0 Then
                        If prefix(.prefix).ModType = 10 Then
                            B = 1
                        ElseIf prefix(.prefix).ModType = 17 Then
                            If Int(Rnd * (.prefixVal And 63)) > 0 Then
                                 B = 1
                            End If
                        End If
                    End If
                    If .suffix > 0 Then
                        If prefix(.suffix).ModType = 10 Then
                            B = 1
                        ElseIf prefix(.suffix).ModType = 17 Then
                            If Int(Rnd * (.SuffixVal And 63)) > 0 Then
                                B = 1
                            End If
                        End If
                    End If
                    If ExamineBit(Object(.Object).Flags, 1) Then B = 1
                    'SKILL: CRAFTMANSHIP
                    'If CanUseSkill(CByte(Index), SKILL_CRAFTSMANSHIP) Then
                    '    If Int(Rnd * 20) < player(Index).SkillLevel(SKILL_CRAFTSMANSHIP) Then
                    '        b = 1
                    '    End If
                    'End If
                    
                    Armor = Armor + Modifier
                    If B = 0 Then
                        A = .Value - 1 'Durability
                        If A <= 0 Then
                            'Object Is Destroyed
                            SendSocket Index, Chr2(57) + Chr2(EQ_RING)
                            .Object = 0
                            CalculateStats Index
                        Else
                            .Value = A
                            SendSocket Index, Chr2(119) + Chr2(20 + EQ_RING) + QuadChar(.Value)
                        End If
                    End If
                End If
            End With
        End If

        If Rnd > 0.2 Then
            'Body Shot
            If .Equipped(EQ_ARMOR).Object > 0 Then
                'Uses Armor
                With .Equipped(EQ_ARMOR)
                    Armor = Armor + Object(.Object).Data(1) * 256 + Object(.Object).Data(2)
                    resist = resist + Object(.Object).Data(3)
                    B = 0
                    If .prefix > 0 Then
                        If prefix(.prefix).ModType = 10 Then
                            B = 1
                        ElseIf prefix(.prefix).ModType = 17 Then
                            If Int(Rnd * (.prefixVal And 63)) > 0 Then
                                 B = 1
                            End If
                        End If
                    End If
                    If .suffix > 0 Then
                        If prefix(.suffix).ModType = 10 Then
                            B = 1
                        ElseIf prefix(.suffix).ModType = 17 Then
                            If Int(Rnd * (.SuffixVal And 63)) > 0 Then
                                B = 1
                            End If
                        End If
                    End If
                    If ExamineBit(Object(.Object).Flags, 1) Then B = 1
                    'SKILL: CRAFTMANSHIP
                    'If CanUseSkill(CByte(Index), SKILL_CRAFTSMANSHIP) Then
                    '    If Int(Rnd * 20) < player(Index).SkillLevel(SKILL_CRAFTSMANSHIP) Then
                    '        b = 1
                    '    End If
                    'End If
                    If B = 0 Then
                        A = .Value - 1
                        If A <= 0 Then
                            'Object Is Destroyed
                            SendSocket Index, Chr2(57) + Chr2(EQ_ARMOR)
                            .Object = 0
                            CalculateStats Index
                        Else
                            .Value = A
                            SendSocket Index, Chr2(119) + Chr2(20 + EQ_ARMOR) + QuadChar(.Value)
                        End If
                    End If
                End With
            End If
        Else
            'Head Shot
            If .Equipped(EQ_HELMET).Object > 0 Then
                'Uses Helmet
                With .Equipped(EQ_HELMET)
                    Armor = Armor + Object(.Object).Data(1) * 256 + Object(.Object).Data(2)
                    resist = resist + Object(.Object).Data(3)
                    B = 0
                    If .prefix > 0 Then
                        If prefix(.prefix).ModType = 10 Then
                            B = 1
                        ElseIf prefix(.prefix).ModType = 17 Then
                            If Int(Rnd * (.prefixVal And 63)) > 0 Then
                                 B = 1
                            End If
                        End If
                    End If
                    If .suffix > 0 Then
                        If prefix(.suffix).ModType = 10 Then
                            B = 1
                        ElseIf prefix(.suffix).ModType = 17 Then
                            If Int(Rnd * (.SuffixVal And 63)) > 0 Then
                                B = 1
                            End If
                        End If
                    End If
                    If ExamineBit(Object(.Object).Flags, 1) Then B = 1
                    'SKILL: CRAFTMANSHIP
                    'If CanUseSkill(CByte(Index), SKILL_CRAFTSMANSHIP) Then
                    '    If Int(Rnd * 20) < player(Index).SkillLevel(SKILL_CRAFTSMANSHIP) Then
                    '        b = 1
                    '    End If
                    'End If
                    If B = 0 Then
                        A = .Value - 1
                        If A <= 0 Then
                            'Object Is Destroyed
                            SendSocket Index, Chr2(57) + Chr2(EQ_HELMET)
                            .Object = 0
                            CalculateStats Index
                        Else
                            .Value = A
                            SendSocket Index, Chr2(119) + Chr2(20 + EQ_HELMET) + QuadChar(.Value)
                        End If
                    End If
                End With
            End If
        End If
        
        'SKILL: TOUGHNESS
        'If CanUseSkill(CByte(Index), SKILL_TOUGHNESS) Then
        '    If .SkillLevel(SKILL_TOUGHNESS) > Damage Then
        '        Damage = 0
        '    Else
        '        Armor = Armor + .SkillLevel(SKILL_TOUGHNESS)
        '    End If
        'End If
        'SKILL: INVULERNABILITY
        If GetStatusEffect(Index, SE_INVULNERABILITY) Then
            damage = 0
        End If
        'SKILL: HOLY ARMOR
        Select Case .Buff.Type
            Case BUFF_HOLYARMOR
                Armor = Armor + (.Buff.Data(0))
        End Select
        

                                                                               
                                                                                
        'SKILL: Armor Mastery
        If .SkillLevel(SKILL_ARMORMASTERY) Then
            Armor = Armor + (Armor * (.SkillLevel(SKILL_ARMORMASTERY) + 5) \ 100)
        End If

        'SKILL: MASTER DEFENSE
        If GetStatusEffect(Index, SE_MASTERDEFENSE) Then
            Armor = Armor * 2
        End If
        
        Armor = Armor + .PhysicalDefenseBonus
        
        'SKILL: BERSERK
        If GetStatusEffect(Index, SE_BERSERK) Then
            damage = damage * 1.23
        End If
        If Armor > damage Then
            damage = 0
        Else
            damage = damage - Armor
        End If

            'SKILL: SHATTERSHIELD
            If GetStatusEffect(Index, SE_SHATTERSHIELD) Then
                B = .StatusData(SE_SHATTERSHIELD).Data(0) * 256 + .StatusData(SE_SHATTERSHIELD).Data(1)
                SetPlayerMana Index, .Mana + (damage * B) / 100
                damage = damage - (damage * B) / 100
                SetPlayerMana Index, .Mana + B
                RemoveStatusEffect Index, SE_SHATTERSHIELD
            End If
            'SKILL: MANASHIELD
            If .Buff.Type = BUFF_MANASHIELD Then
                If .Mana >= (damage * .Buff.Data(0)) / 50 Then
                    .Mana = .Mana - (damage * .Buff.Data(0)) / 50
                    damage = (damage * (100 - .Buff.Data(0))) / 100
                    If .Mana <= 0 Then
                        .Mana = 0
                        ClearBuff Index
                    End If
                    SendSocket Index, Chr2(48) + DoubleChar(.Mana)
                Else
                    ClearBuff Index
                End If
            End If


   
    
    damage = damage - ((damage * (resist + .PhysicalResistanceBonus)) \ 100)
    If damage < 0 Then damage = 0
    
    End With
    
    PlayerArmor = damage
End Function
Function PlayerDamage(Index As Long) As Long
Randomize
    Dim A As Long, B As Long, damage As Long, Modifier As Long
    With player(Index)
        If .Equipped(EQ_WEAPON).Object > 0 Then
            'Uses Weapon
            With .Equipped(EQ_WEAPON)
                If Object(.Object).Type = 1 Then
                    damage = Int(Rnd * ((Object(.Object).Data(2) * 256 + Object(.Object).Data(3)) - Object(.Object).Data(1))) + Object(.Object).Data(1)
                    B = 0
                ElseIf Object(.Object).Type = 10 Then
                    If player(Index).Equipped(EQ_AMMO).Object > 0 Then
                        If Object(player(Index).Equipped(EQ_AMMO).Object).Type = 11 Then 'Ammo
                            'Bow, etc +PLUS+
                            damage = Object(.Object).Data(2)
                            'Projectile Min-Max damage
                            damage = damage + Int(Rnd * (Object(player(Index).Equipped(EQ_AMMO).Object).Data(3) - Object(player(Index).Equipped(EQ_AMMO).Object).Data(2))) + Object(player(Index).Equipped(EQ_AMMO).Object).Data(2)
                        Else
                            damage = 0
                            Exit Function
                        End If
                    Else
                        damage = 0
                    End If
                End If
                If .prefix > 0 Then
                    If prefix(.prefix).ModType = 10 Then
                        B = 1
                    ElseIf prefix(.prefix).ModType = 17 Then
                        If Int(Rnd * (.prefixVal And 63)) > 0 Then
                             B = 1
                        End If
                    End If
                End If
                If .suffix > 0 Then
                    If prefix(.suffix).ModType = 10 Then
                        B = 1
                    ElseIf prefix(.suffix).ModType = 17 Then
                        If Int(Rnd * (.SuffixVal And 63)) > 0 Then
                            B = 1
                        End If
                    End If
                End If
                
                
                If ExamineBit(Object(.Object).Flags, 3) Then
                    damage = damage + (damage * GetSkillLevel(Index, SKILL_TWOHAND)) / 100
                End If
                
                
                
                If ExamineBit(Object(.Object).Flags, 1) Then B = 1
                'SKILL: CRAFTMANSHIP
                'If CanUseSkill(CByte(Index), SKILL_CRAFTSMANSHIP) Then
                '    If Int(Rnd * 20) < player(Index).SkillLevel(SKILL_CRAFTSMANSHIP) Then
                '        b = 1
                '    End If
                'End If
                If B = 0 Then
                    A = .Value - 1 '(Damage / 2)
                    If A <= 0 Then
                        'Object Is Destroyed
                        SendSocket Index, Chr2(57) + Chr2(EQ_WEAPON)
                        .Object = 0
                        CalculateStats Index
                    Else
                        .Value = A
                        SendSocket Index, Chr2(119) + Chr2(20 + EQ_WEAPON) + QuadChar(.Value)
                    End If
                End If
            End With
        End If
        
        If .Equipped(EQ_DUALWIELD).Object > 0 Then
            'May Have Dual Wield Weapon
            With .Equipped(EQ_DUALWIELD)
                If ExamineBit(Object(.Object).Flags, 5) Then
                    damage = damage + Int(Rnd * ((Object(.Object).Data(2) * 256 + Object(.Object).Data(3)) - Object(.Object).Data(1))) + Object(.Object).Data(1)
                    B = 0
                    If .prefix > 0 Then
                        If prefix(.prefix).ModType = 10 Then
                            B = 1
                        ElseIf prefix(.prefix).ModType = 17 Then
                            If Int(Rnd * (.prefixVal And 63)) > 0 Then
                                 B = 1
                            End If
                        End If
                    End If
                    If .suffix > 0 Then
                        If prefix(.suffix).ModType = 10 Then
                            B = 1
                        ElseIf prefix(.suffix).ModType = 17 Then
                            If Int(Rnd * (.SuffixVal And 63)) > 0 Then
                                B = 1
                            End If
                        End If
                    End If
                    If ExamineBit(Object(.Object).Flags, 1) Then B = 1
                    'SKILL: CRAFTMANSHIP
                    'If CanUseSkill(CByte(Index), SKILL_CRAFTSMANSHIP) Then
                    '    If Int(Rnd * 20) < player(Index).SkillLevel(SKILL_CRAFTSMANSHIP) Then
                    '        b = 1
                    '    End If
                    'End If
                    If B = 0 Then
                        A = .Value - 1 '(Damage / 2)
                        If A <= 0 Then
                            'Object Is Destroyed
                            SendSocket Index, Chr2(57) + Chr2(EQ_DUALWIELD)
                            .Object = 0
                            CalculateStats Index
                        Else
                            .Value = A
                            SendSocket Index, Chr2(119) + Chr2(20 + EQ_DUALWIELD) + QuadChar(.Value)
                        End If
                    End If
                End If
            End With
        End If
        If .Equipped(EQ_RING).Object > 0 Then
            'Has Ring
            With .Equipped(EQ_RING)
                If Object(.Object).Data(0) = 0 Then
                    Modifier = Object(.Object).Data(2)
                    damage = damage + Modifier
                    B = 0
                    If .prefix > 0 Then
                        If prefix(.prefix).ModType = 10 Then
                            B = 1
                        ElseIf prefix(.prefix).ModType = 17 Then
                            If Int(Rnd * (.prefixVal And 63)) > 0 Then
                                 B = 1
                            End If
                        End If
                    End If
                    If .suffix > 0 Then
                        If prefix(.suffix).ModType = 10 Then
                            B = 1
                        ElseIf prefix(.suffix).ModType = 17 Then
                            If Int(Rnd * (.SuffixVal And 63)) > 0 Then
                                B = 1
                            End If
                        End If
                    End If
                    If ExamineBit(Object(.Object).Flags, 1) Then B = 1
                    'SKILL: CRAFTMANSHIP
                    'If CanUseSkill(CByte(Index), SKILL_CRAFTSMANSHIP) Then
                    '    If Int(Rnd * 20) < player(Index).SkillLevel(SKILL_CRAFTSMANSHIP) Then
                    '        b = 1
                    '    End If
                    'End If
                    If B = 0 Then
                        A = .Value - 1 'Modifier
                        If A <= 0 Then
                            'Object Is Destroyedf
                            SendSocket Index, Chr2(57) + Chr2(EQ_RING)
                            .Object = 0
                            CalculateStats Index
                        Else
                            .Value = A
                           SendSocket Index, Chr2(119) + Chr2(20 + EQ_RING) + QuadChar(.Value)
                        End If
                    End If
                End If
            End With
        End If
        
        damage = damage + GetStatPerBonus(.strength, StrengthPerDamage)
        damage = damage + .AttackDamageBonus
        'SKILL: VENGEANCE
        If .SkillLevel(SKILL_VENGEANCE) Then
            damage = damage + (damage * (100 - CLng(.HP) * 100 / .MaxHP) * (.SkillLevel(SKILL_VENGEANCE) * 0.03)) / 100
        End If
        
        If GetStatusEffect(Index, SE_INVULNERABILITY) Then damage = damage / 2
        
        If GetStatusEffect(Index, SE_BERSERK) Then damage = damage * 1.23
        If player(Index).StatusData(SE_FIERYESSENCE).timer > 0 Then damage = damage * 1.1
        If player(Index).SkillLevel(SKILL_VENGEFUL) Then damage = damage * 1.1
        If player(Index).SkillLevel(SKILL_GREATFORCE) Then damage = damage * 1.15
        
        
        PlayerDamage = damage
    End With
End Function
Sub PlayerDied(Index As Long, Optional PK As Boolean = False, Optional Killer As Long = 0, Optional NoKiller As Boolean = False, Optional DeathByTrap As Boolean = False)
   ' If player(Index).deathStamp = 0 And player(Index).SkillLevel(SKILL_LOSTARCANA) > 0 And Not DeathByTrap Then
   '     player(Index).deathStamp = GetTickCount + 10000
   '     player(Index).Killer = Killer
   '     player(Index).KillerPK = PK
   '     SetStatusEffect Index, SE_LOSTARCANA
   '     player(Index).StatusData(SE_LOSTARCANA).timer = 6
   ' End If
    
   ' If player(Index).deathStamp >= GetTickCount Or player(Index).SkillLevel(SKILL_LOSTARCANA) = 0 Then
    
   '     If PK = False And Killer > 0 Then
   '         SendSocket Index, Chr2(53) + Chr2(map(mapnum).monster(Killer).monster) 'Monster Killed You
   '         SendAllBut Index, Chr2(62) + Chr2(b) + Chr2(map(mapnum).monster(Killer).monster) 'Player was killed by monster
   '         CreateFloatingEvent mapnum, player(Index).x, player(Index).y, FT_ENDED
   '     End If
    
   '     If player(Index).SkillLevel(SKILL_LOSTARCANA) = 1 Then player(Index).deathStamp = 0
   
   PrintLog player(Index).user & "(" & player(Index).Name & ")" & " died on map " & player(Index).map & "   class: " & Class(player(Index).Class).Name
   
   
        Dim A As Long, B As Long, C As Long, D As Long, E As Long, OldX As Long, OldY As Long, ST1 As String, st2 As String
        Dim mapNum As Long
        With player(Index)
            If .Trading = True Then
                CloseTrade Index
            End If
            ST1 = ""
            st2 = ""
            mapNum = .map
            
            If PK Then
                If Killer > 0 Then
                    'SKILL: NECROMANCY
                    If .Buff.Type = BUFF_NECROMANCY Then
                        A = .Buff.Data(0)
                        If .Level < player(Killer).Level Then
                            A = A - (player(Killer).Level - .Level)
                        End If
                        If A > 0 Then
                            If player(Killer).HP < player(Killer).MaxHP Then
                                player(Killer).HP = player(Killer).HP + A
                                If player(Killer).HP > player(Killer).MaxHP Then player(Killer).HP = player(Killer).MaxHP
                                SendSocket Killer, Chr2(46) + DoubleChar$(player(Killer).HP)
                            End If
                            CreateFloatingText mapNum, player(Killer).x, player(Killer).y, 10, CStr(A)
                        End If
                    End If
                    'SKILL: EVOCATION
                    If .Buff.Type = BUFF_EVOCATION Then
                        A = .Buff.Data(0)
                        If .Level < player(Killer).Level Then
                            A = A - (player(Killer).Level - .Level)
                        End If
                        If A > 0 Then
                            If player(Killer).Mana < player(Killer).MaxMana Then
                                player(Killer).Mana = player(Killer).Mana + A
                                If player(Killer).Mana > player(Killer).MaxMana Then player(Killer).Mana = player(Killer).MaxMana
                                SendSocket Killer, Chr2(48) + DoubleChar$(player(Killer).Mana)
                            End If
                            CreateFloatingText mapNum, player(Killer).x, player(Killer).y, 9, CStr(A)
                        End If
                    End If
                End If
            End If
            
            Parameter(0) = Index
            Parameter(1) = IIf(PK, 1, 0)
            If NoKiller Then Parameter(1) = 2
            Parameter(2) = Killer
            OldX = .x
            OldY = .y
            If RunScript("PLAYERDIE") = 0 Then
                CreateFloatingEvent mapNum, OldX, OldY, FT_ENDED
                If ExamineBit(map(mapNum).Flags(0), 2) = False And .Level > DeathDropItemsLevel Then
                    For A = 1 To 20
                        If .Inv(A).Object > 0 Then
                            If (Rnd <= 0.15) And RunScript("DROPOBJ" + CStr(.Inv(A).Object)) = 0 And ExamineBit(Object(.Inv(A).Object).Flags, 0) = 0 Then
                                B = FreeMapObj(mapNum)
                                If B >= 0 Then
                                    If Object(.Inv(A).Object).Type = 6 Then
                                        E = Int(Rnd * .Inv(A).Value) + 1
                                    Else
                                        E = .Inv(A).Value
                                    End If
                                    map(mapNum).Object(B).Object = .Inv(A).Object
                                    map(mapNum).Object(B).Value = E
                                    map(mapNum).Object(B).prefix = .Inv(A).prefix
                                    map(mapNum).Object(B).prefixVal = .Inv(A).prefixVal
                                    map(mapNum).Object(B).suffix = .Inv(A).suffix
                                    map(mapNum).Object(B).SuffixVal = .Inv(A).SuffixVal
                                    map(mapNum).Object(B).Affix = .Inv(A).Affix
                                    map(mapNum).Object(B).AffixVal = .Inv(A).AffixVal
                                    map(mapNum).Object(B).ObjectColor = .Inv(A).ObjectColor
                                    For C = 0 To 3
                                        map(mapNum).Object(B).Flags(C) = .Inv(A).Flags(C)
                                    Next C
                                    C = 0
                                    map(mapNum).Object(B).TimeStamp = GetTickCount + GLOBAL_DEATH_DROP_RESET_RATE
                                    map(mapNum).Object(B).deathObj = True
                                    If .Inv(A).prefix > 0 And .Inv(A).prefix < 256 Then
                                        C = prefix(.Inv(A).prefix).Light.Intensity
                                        D = prefix(.Inv(A).prefix).Light.Radius
                                    End If
                                    If .Inv(A).suffix > 0 And .Inv(A).suffix < 256 Then
                                        C = C + prefix(.Inv(A).suffix).Light.Intensity
                                        D = D + prefix(.Inv(A).suffix).Light.Radius
                                    End If
                                    If .Inv(A).Affix > 0 And .Inv(A).Affix < 256 Then
                                        C = C + prefix(.Inv(A).Affix).Light.Intensity
                                        D = D + prefix(.Inv(A).Affix).Light.Radius
                                    End If
                                    If C > 255 Then C = 255
                                    If D > 255 Then D = 255
                                    If PK = True And DeathByTrap = False Then
                                        map(mapNum).Object(B).x = player(Killer).x
                                        map(mapNum).Object(B).y = player(Killer).y
                                        ST1 = ST1 + DoubleChar(24) + Chr2(14) + Chr2(B) + DoubleChar(.Inv(A).Object) + Chr2(player(Killer).x) + Chr2(player(Killer).y) + Chr2(C) + Chr2(D) + QuadChar$(map(mapNum).Object(B).Value) + Chr2(map(mapNum).Object(B).prefix) + Chr2(map(mapNum).Object(B).prefixVal) + Chr2(map(mapNum).Object(B).suffix) + Chr2(map(mapNum).Object(B).SuffixVal) + Chr2(map(mapNum).Object(B).Affix) + Chr2(map(mapNum).Object(B).AffixVal) + QuadChar(GLOBAL_DEATH_DROP_RESET_RATE) + Chr2(Abs(map(mapNum).Object(B).deathObj)) + Chr2(map(mapNum).Object(B).ObjectColor)
                                    Else
                                        map(mapNum).Object(B).x = OldX
                                        map(mapNum).Object(B).y = OldY
                                        ST1 = ST1 + DoubleChar(24) + Chr2(14) + Chr2(B) + DoubleChar(.Inv(A).Object) + Chr2(OldX) + Chr2(OldY) + Chr2(C) + Chr2(D) + QuadChar$(map(mapNum).Object(B).Value) + Chr2(map(mapNum).Object(B).prefix) + Chr2(map(mapNum).Object(B).prefixVal) + Chr2(map(mapNum).Object(B).suffix) + Chr2(map(mapNum).Object(B).SuffixVal) + Chr2(map(mapNum).Object(B).Affix) + Chr2(map(mapNum).Object(B).AffixVal) + QuadChar(GLOBAL_DEATH_DROP_RESET_RATE) + Chr2(Abs(map(mapNum).Object(B).deathObj)) + Chr2(map(mapNum).Object(B).ObjectColor)
                                    End If
                                End If
                                If .Inv(A).Value <> E And Object(.Inv(A).Object).Type = 6 Then
                                    .Inv(A).Value = .Inv(A).Value - E
                                    st2 = st2 + DoubleChar(15) + Chr2(17) + Chr2(A) + DoubleChar(.Inv(A).Object) + QuadChar(.Inv(A).Value) + String$(7, Chr2(0)) 'Update inv obj
                                Else
                                    .Inv(A).Object = 0
                                    st2 = st2 + DoubleChar(2) + Chr2(18) + Chr2(A)
                                End If
                            End If
                        End If
                    Next A
                    E = 0
                    For A = 1 To 5
                        If .Equipped(A).Object > 0 Then
                            If Not ExamineBit(Object(.Equipped(A).Object).Flags, 0) Then
                                E = E + 1
                            End If
                        End If
                    Next A
                    If E > 0 Then
                        A = Int(Rnd * 5) + 1
                        Do Until E = 0
                            With .Equipped(A)
                                If .Object > 0 Then
                                    If Not ExamineBit(Object(.Object).Flags, 0) Then
                                        If RunScript("DROPOBJ" + CStr(.Object)) = 0 Then
                                            B = FreeMapObj(mapNum)
                                            If B >= 0 Then
                                                map(mapNum).Object(B).Object = .Object
                                                map(mapNum).Object(B).Value = .Value
                                                map(mapNum).Object(B).prefix = .prefix
                                                map(mapNum).Object(B).prefixVal = .prefixVal
                                                map(mapNum).Object(B).suffix = .suffix
                                                map(mapNum).Object(B).SuffixVal = .SuffixVal
                                                map(mapNum).Object(B).Affix = .Affix
                                                map(mapNum).Object(B).AffixVal = .AffixVal
                                                map(mapNum).Object(B).TimeStamp = GetTickCount + GLOBAL_DEATH_DROP_RESET_RATE
                                                map(mapNum).Object(B).deathObj = True
                                                map(mapNum).Object(B).ObjectColor = .ObjectColor
                                                For C = 0 To 3
                                                    map(mapNum).Object(B).Flags(C) = .Flags(C)
                                                Next C
                                                C = 0
                                                If .prefix > 0 And .prefix < 256 Then
                                                    C = prefix(.prefix).Light.Intensity
                                                    D = prefix(.prefix).Light.Radius
                                                End If
                                                If .suffix > 0 And .suffix < 256 Then
                                                    C = C + prefix(.suffix).Light.Intensity
                                                    D = D + prefix(.suffix).Light.Radius
                                                End If
                                                If .Affix > 0 And .Affix < 256 Then
                                                    C = C + prefix(.Affix).Light.Intensity
                                                    D = D + prefix(.Affix).Light.Radius
                                                End If
                                                If PK = True Then
                                                    map(mapNum).Object(B).x = player(Killer).x
                                                    map(mapNum).Object(B).y = player(Killer).y
                                                    ST1 = ST1 + DoubleChar(24) + Chr2(14) + Chr2(B) + DoubleChar(.Object) + Chr2(player(Killer).x) + Chr2(player(Killer).y) + Chr2(C) + Chr2(D) + QuadChar$(map(mapNum).Object(B).Value) + Chr2(map(mapNum).Object(B).prefix) + Chr2(map(mapNum).Object(B).prefixVal) + Chr2(map(mapNum).Object(B).suffix) + Chr2(map(mapNum).Object(B).SuffixVal) + Chr2(map(mapNum).Object(B).Affix) + Chr2(map(mapNum).Object(B).AffixVal) + QuadChar(GLOBAL_DEATH_DROP_RESET_RATE) + Chr2(Abs(map(mapNum).Object(B).deathObj)) + Chr2(map(mapNum).Object(B).ObjectColor)
                                                Else
                                                    map(mapNum).Object(B).x = OldX
                                                    map(mapNum).Object(B).y = OldY
                                                    ST1 = ST1 + DoubleChar(24) + Chr2(14) + Chr2(B) + DoubleChar(.Object) + Chr2(OldX) + Chr2(OldY) + Chr2(C) + Chr2(D) + QuadChar$(map(mapNum).Object(B).Value) + Chr2(map(mapNum).Object(B).prefix) + Chr2(map(mapNum).Object(B).prefixVal) + Chr2(map(mapNum).Object(B).suffix) + Chr2(map(mapNum).Object(B).SuffixVal) + Chr2(map(mapNum).Object(B).Affix) + Chr2(map(mapNum).Object(B).AffixVal) + QuadChar(GLOBAL_DEATH_DROP_RESET_RATE) + Chr2(Abs(map(mapNum).Object(B).deathObj)) + Chr2(map(mapNum).Object(B).ObjectColor)
                                                End If
                                            End If
                                            .Object = 0
                                            st2 = st2 + DoubleChar(2) + Chr2(18) + Chr2(20 + A)
                                            E = 0
                                        Else
                                            E = E - 1
                                        End If
                                    End If
                                End If
                                A = Int(Rnd * 5) + 1
                            End With
                        Loop
                    End If
                End If
                
                
                
                
                If ST1 <> "" Then
                    SendToMapRaw mapNum, ST1
                End If
                
    
                
                If .map = mapNum Then
                
                    If .Guild > 0 Then
                        If Guild(.Guild).Hall >= 1 Then
                            A = 1
                        Else
                            A = 0
                        End If
                    Else
                        A = 0
                    End If
                    
                    Partmap Index
                    
                    If A = 0 Then
                        'Random Start Location
                        A = Int(Rnd * 2)
                        
                        .map = World.StartLocation(A).map
                        .x = World.StartLocation(A).x
                        .y = World.StartLocation(A).y
                        
                        If World.StartLocation(A).Message <> "" Then
                            st2 = st2 + DoubleChar(2 + Len(World.StartLocation(A).Message)) + Chr2(56) + Chr2(15) + World.StartLocation(A).Message
                        End If
                    Else
                        A = Guild(.Guild).Hall
                        
                        .map = Hall(A).StartLocation.map
                        .x = Hall(A).StartLocation.x
                        .y = Hall(A).StartLocation.y
                    End If
                    
                    If .map < 1 Then .map = 1
                    If .map > 5000 Then .map = 5000
                    If .y > 11 Then .y = 11
                    If .x > 11 Then .x = 11
                    
                    JoinMap Index
                
                
                End If
                
                If .Status = 1 Then .Status = 0
                .StatusEffect = 0
                
                SendToMap mapNum, Chr2(112) + Chr(Index) + QuadChar(0)
                st2 = st2 + DoubleChar(6) + Chr2(112) + Chr2(Index) + QuadChar(0)
                If ExamineBit(map(mapNum).Flags(0), 2) = False Then
                    If .Level > 5 Then
                        .Experience = Int((4 / 5) * .Experience)
                    End If
                End If
                st2 = st2 + DoubleChar(5) + Chr2(60) + QuadChar(.Experience)
    
                .HP = .MaxHP
                .Mana = .MaxMana
                .Energy = .MaxEnergy
                
                st2 = st2 + DoubleChar(3) + Chr2(46) + DoubleChar(CLng(.HP))
    
                    SendToPartyAllBut .Party, Index, Chr2(104) + Chr2(0) + Chr2(Index) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                    SendToGuildAllBut Index, CLng(.Guild), Chr2(104) + Chr2(0) + Chr2(Index) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                
                If st2 <> "" Then
                    SendRaw Index, st2
                End If
            End If
        End With
        CalculateStats Index
  '  End If
End Sub

'Private Function GetString(St As String, Num As Long) As String
'    Dim a As Long, b As Long, c As Long
'
'    b = 1
'    If Num > 0 Then
'        For a = 1 To Num
'GetAgain:
'                c = b
'                b = InStr(b + 1, St, " ")
'                If b = c + 1 Then GoTo GetAgain
'
'            If b = 0 Then Exit For
'        Next a
'        If b <> 0 Then GetString = Mid$(St, b + 1, Len(St) - b)
'    End If
'End Function

Sub ResetMap(mapNum As Long)
    Dim A As Long, B As Long, x As Long, y As Long
    Dim NumPlayers As Long
    Dim ST1 As String
    
    With map(mapNum)
        NumPlayers = .NumPlayers
        For A = 0 To 49
            With .Object(A)
                If .Object > 0 Then
                    If map(mapNum).Tile(.x, .y).Att <> 5 Then
                        .Object = 0
                        If NumPlayers > 0 Then
                            ST1 = ST1 + DoubleChar(2) + Chr2(15) + Chr2(A)
                        End If
                    End If
                End If
            End With
        Next A
        For A = 0 To 9
            With .trap(A)
                .ActiveCounter = 0
                .CreatedTime = 0
                .player = 0
                .strength = 0
                .trapID = 0
                .Type = 0
                .x = 0
                .y = 0
            End With
        Next A
        For A = 1 To 30
            With .Projectile(A)
                .sprite = 0
            End With
        Next A

        For A = 0 To 9
            With .Door(A)
                If .Att > 0 Or .Wall > 0 Then
                    map(mapNum).Tile(.x, .y).Att = .Att
                    map(mapNum).Tile(.x, .y).WallTile = .Wall
                    If NumPlayers > 0 Then
                        ST1 = ST1 + DoubleChar(2) + Chr2(37) + Chr2(A)
                    End If
                    .Att = 0
                    .Wall = 0
                End If
            End With
        Next A
        If ExamineBit(.Flags(0), 3) = True Then
            'Create Monsters
            For A = 0 To 9
                ST1 = ST1 + NewMapMonster(mapNum, A)
            Next A
        Else
            'Clear Monsters
            For A = 0 To 9
                If .monster(A).monster > 0 Then
                    .monster(A).monster = 0
                    If NumPlayers > 0 Then
                        ST1 = ST1 + DoubleChar(2) + Chr2(39) + Chr2(A + 10)
                    End If
                End If
            Next A
        End If
        If NumPlayers > 0 Then
            SendToMapRaw mapNum, ST1
        End If
        ST1 = ""
        For y = 0 To 11
            For x = 0 To 11
                With map(mapNum).Tile(x, y)
                    If .Att = 7 Then
                        A = (.AttData(0) \ 2) * 4
                        B = (.AttData(0) And 1) * 256 + .AttData(3)
                        If A Then
                            NewMapObject mapNum, B, CLng(.AttData(1)) * 256 + CLng(.AttData(2)), x, y, True, A
                        Else
                            NewMapObject mapNum, B, CLng(.AttData(1)) * 256 + CLng(.AttData(2)), x, y, True, 0
                        End If
                    End If
                End With
            Next x
        Next y
        .ResetTimer = 0
    End With
End Sub

Sub SoftResetMap(mapNum As Long, Optional clearMonsters As Boolean = True)
    Dim A As Long, B As Long, x As Long, y As Long
    Dim NumPlayers As Long
    Dim ST1 As String
    
    With map(mapNum)
        NumPlayers = .NumPlayers
        For A = 0 To 49
            With .Object(A)
            
                If .Object > 0 Then
                    If map(mapNum).Tile(.x, .y).Att <> 5 Then
                        If .TimeStamp = 0 Then
                        .Object = 0
                            If NumPlayers > 0 Then
                                ST1 = ST1 + DoubleChar(2) + Chr2(15) + Chr2(A)
                            End If
                        End If
                    End If
                End If
            End With
        Next A
        For A = 0 To 9
            With .trap(A)
                .ActiveCounter = 0
                .CreatedTime = 0
                .player = 0
                .strength = 0
                .trapID = 0
                .Type = 0
                .x = 0
                .y = 0
            End With
        Next A
        For A = 1 To 30
            With .Projectile(A)
                .sprite = 0
            End With
        Next A

        For A = 0 To 9
            With .Door(A)
                If .Att > 0 Or .Wall > 0 Then
                    map(mapNum).Tile(.x, .y).Att = .Att
                    map(mapNum).Tile(.x, .y).WallTile = .Wall
                    If NumPlayers > 0 Then
                        ST1 = ST1 + DoubleChar(2) + Chr2(37) + Chr2(A)
                    End If
                    .Att = 0
                    .Wall = 0
                End If
            End With
        Next A
        If ExamineBit(.Flags(0), 3) = True Then
            'Create Monsters
            For A = 0 To 9
                ST1 = ST1 + NewMapMonster(mapNum, A)
            Next A
        Else
            If clearMonsters Then
                'Clear Monsters
                For A = 0 To 9
                    If .monster(A).monster > 0 Then
                        .monster(A).monster = 0
                        If NumPlayers > 0 Then
                            ST1 = ST1 + DoubleChar(2) + Chr2(39) + Chr2(A + 10)
                        End If
                    End If
                Next A
            End If
        End If
        If NumPlayers > 0 Then
            SendToMapRaw mapNum, ST1
        End If
        ST1 = ""
        For y = 0 To 11
            For x = 0 To 11
                With map(mapNum).Tile(x, y)
                    If .Att = 7 Then
                        A = (.AttData(0) \ 2) * 4
                        B = (.AttData(0) And 1) * 256 + .AttData(3)
                        If A Then
                            NewMapObject mapNum, B, CLng(.AttData(1)) * 256 + CLng(.AttData(2)), x, y, True, A
                        Else
                            NewMapObject mapNum, B, CLng(.AttData(1)) * 256 + CLng(.AttData(2)), x, y, True, 0
                        End If
                    End If
                End With
            Next x
        Next y
        .ResetTimer = 0
    End With
End Sub


Sub SendCharacterData(Index As Long)
Dim St As String, ST1 As String, st2 As String
    With player(Index)
        If .Level > 0 And .Class > 0 Then
            St = UserRS!SkillLevels
            ST1 = UserRS!SkillEXP
            If .Guild > 0 Then
                st2 = Guild(.Guild).Name
            Else
                st2 = ""
            End If
            SendSocket Index, Chr2(3) + Chr2(.Level) + Chr2(.Class) + Chr2(0) + Chr2(0) + Chr2(0) + Chr2(0) + Chr2(.Gender) + Chr2(.sprite) + DoubleChar(CLng(.HP)) + DoubleChar(CLng(.Energy)) + DoubleChar(CLng(.Mana)) + DoubleChar(CLng(.MaxHP)) + DoubleChar(CLng(.MaxEnergy)) + DoubleChar(CLng(.MaxMana)) + Chr2(.strength) + Chr2(.Agility) + Chr2(.Endurance) + Chr2(.Wisdom) + Chr2(.Constitution) + Chr2(.Intelligence) + Chr2(.Level) + Chr2(.Status) + Chr2(.Guild) + Chr2(.GuildRank) + Chr2(.Access) + Chr2(Index) + QuadChar(.Experience) + Chr2(.Squelched) + QuadChar(.StatusEffect) + DoubleChar(CLng(.StatPoints)) + DoubleChar(CLng(.SkillPoints)) + St + ST1 + .Name + Chr2(0) + "" + Chr2(0) + st2
        Else
            SendSocket Index, Chr2(3)
        End If
    End With
End Sub

Function Cryp(St As String) As String
Dim A As Long, ST1 As String
For A = 1 To Len(St)
    ST1 = ST1 & Chr2(219 Xor Asc(Mid$(St, A, 1)))
Next A
Cryp = ST1
End Function

Sub SendDataPacket(Index As Long, StartNum As Long)
    Dim A As Long, ST1 As String
    
    For A = StartNum To 255
        If Len(NPC(A).Name) > 0 Then
            With NPC(A)
                ST1 = ST1 + DoubleChar(6 + Len(.Name)) + Chr2(85) + Chr2(A) + Chr2(.Flags) + Chr2(.Portrait) + Chr2(.sprite) + Chr2(.direction) + Cryp(.Name)
            End With
        End If
        If Len(Hall(A).Name) > 0 Then
            With Hall(A)
                ST1 = ST1 + DoubleChar(2 + Len(.Name)) + Chr2(82) + Chr2(A) + Cryp(.Name)
            End With
        End If
        If Len(Guild(A).Name) <> 0 Then
            With Guild(A)
                ST1 = ST1 + DoubleChar(2 + Len(.Name)) + Chr2(70) + Chr2(A) + Cryp(.Name)
                ST1 = ST1 + DoubleChar(12) + Chr2(136) + Chr2(A) + Chr2(CountGuildMembers(A)) + Chr2(CountGuildMembersOL(A)) + QuadChar(GetGuildRenown(A) / CountGuildMembers(A)) + Chr2(.Symbol1) + Chr2(.Symbol2) + Chr2(.Symbol3) + Chr2(.Hall)
            End With
        End If
        If Len(prefix(A).Name) > 0 Then
            With prefix(A)
                ST1 = ST1 + DoubleChar(5 + Len(.Name)) + Chr2(108) + Chr2(A) + Chr2(.Light.Intensity) + Chr2(.Light.Radius) + Chr2(.ModType) + Cryp(.Name)
            End With
        End If
        If A <= 255 Then
            If Len(Lights(A).Name) > 0 Then
                With Lights(A)
                    ST1 = ST1 + DoubleChar(9 + Len(.Name)) + Chr2(129) + Chr2(A) + Chr2(.red) + Chr2(.green) + Chr2(.blue) + Chr2(.Intensity) + Chr2(.Radius) + Chr2(.MaxFlicker) + Chr2(.FlickerRate) + Cryp(.Name)
                End With
            End If
        End If
        If A <= MaxUsers Then
            If player(A).InUse And player(A).Mode = modePlaying Then
                With player(A)
                    ST1 = ST1 + DoubleChar(3) + Chr2(105) + Chr2(A) + Chr2(.Party)
                End With
            End If
        End If

        If Len(ST1) >= 1024 Then
            If A < 255 Then
                ST1 = ST1 + DoubleChar(5) + Chr2(35) + Chr2(24) + Chr2(1) + DoubleChar(A + 1)
            Else
                ST1 = ST1 + DoubleChar(5) + Chr2(35) + Chr2(24) + Chr2(2) + DoubleChar(1)
            End If
            SendRaw Index, ST1
            Exit Sub
        End If
    Next A
    
    ST1 = ST1 + DoubleChar(5) + Chr2(35) + Chr2(24) + Chr2(2) + DoubleChar(1)
    SendRaw Index, ST1
End Sub

Sub SendItemPacket(Index As Long, StartNum As Long)
    Dim A As Long, ST1 As String

    For A = StartNum To MAXITEMS
        If Object(A).Picture > 0 Then
            With Object(A)
                ST1 = ST1 + DoubleChar(21 + Len(.Name) + Len(.Description)) + Chr2(31) + DoubleChar(CInt(A)) + Chr2(.Picture) + Chr2(.Type) + Chr2(.Data(0)) + Chr2(.Data(1)) + Chr2(.Data(2)) + Chr2(.Data(3)) + Chr2(.Data(4)) + Chr2(.Data(5)) + Chr2(.Data(6)) + Chr2(.Data(7)) + Chr2(.Data(8)) + Chr2(.Data(9)) + Chr2(.MinLevel) + Chr2(.Flags) + DoubleChar$(.Class) + Chr2(.EquipmentPicture) + Cryp(.Name) + Chr2(0) + Cryp(.Description)
            End With
        End If
        If monster(A).sprite > 0 Then
            With monster(A)
                ST1 = ST1 + DoubleChar(15 + Len(.Name)) + Chr2(32) + DoubleChar(A) + Chr2(.sprite) + DoubleChar(CLng(.HP)) + Chr2(.Flags) + Chr2(.DeathSound) + Chr2(.AttackSound) + Chr2(.alpha) + Chr2(.red) + Chr2(.green) + Chr2(.blue) + Chr2(.Light) + Chr2(.Flags2) + Cryp(.Name)
            End With
        End If
        If Len(ST1) >= 1024 Then
            If A < MAXITEMS Then
                ST1 = ST1 + DoubleChar(5) + Chr2(35) + Chr2(24) + Chr2(2) + DoubleChar(A + 1)
            Else
                ST1 = ST1 + DoubleChar(2) + Chr2(35) + Chr2(23)
            End If
            SendRaw Index, ST1
            Exit Sub
        End If
    Next A
    
    ST1 = ST1 + DoubleChar(2) + Chr2(35) + Chr2(23)
    SendRaw Index, ST1
End Sub

Function SpawnMapMonster(mapNum As Long, MonsterNum As Long, MonsterType As Long, tX As Long, tY As Long)
Dim A As Long

    With map(mapNum).monster(MonsterNum)
        .monster = MonsterType
        .x = tX
        .y = tY
        .HP = monster(.monster).HP
        .Distance = monster(.monster).Sight
        .Target = 0
        .TargType = 0
        .MoveSpeed = monster(MonsterType).MoveSpeed
        .AttackSpeed = monster(MonsterType).AttackSpeed
        .Poison = 0
        .PoisonLength = 0
        .R = 0
        .G = 0
        .B = 0
        .A = 0
        For A = 0 To 4
            .Flags(A) = 0
        Next A
        .CurrentQueue = 0
        For A = 0 To 15
            .MonsterQueue(A).Action = 0
        Next A
        .D = Int(Rnd * 4)

            SpawnMapMonster = DoubleChar(9) + Chr2(38) + Chr2(MonsterNum) + DoubleChar(.monster) + Chr2(.x) + Chr2(.y) + Chr2(.D) + DoubleChar(CLng(.HP))

    End With
End Function

Function ValidName(St As String) As Boolean
    Dim A As Long, B As Long
    If Len(St) > 0 Then
        For A = 1 To Len(St)
            B = Asc(Mid$(St, A, 1))
            If (B < 48 Or B > 57) And (B < 65 Or B > 90) And (B < 97 Or B > 122) And B <> 32 And B <> 95 Then
                ValidName = False
                Exit Function
            End If
        Next
    End If
    ValidName = True
End Function
Function WindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim TempSocket As Long
    
    If uMsg >= 1029 And uMsg < 1029 + MaxUsers Then
        Select Case lParam And 255
            Case FD_CLOSE
                AddSocketQue uMsg - 1028, 10000
            Case FD_READ
                ReadClientData uMsg - 1028
        End Select
    End If
    Select Case uMsg
        Case 1025 'Listening Socket
            Select Case lParam And 255
                Case FD_ACCEPT
                    If lParam = FD_ACCEPT Then
                        Dim NewPlayer As Long, Address As sockaddr
                        Dim ClientIP As String
                        Dim A As Long, InvIP As Byte
                        
                        NewPlayer = FreePlayer()
                        If NewPlayer > 0 Then
                            With player(NewPlayer)
                                .Socket = accept(ListeningSocket, Address, sockaddr_size)
                                If Not .Socket = INVALID_SOCKET Then
                                    ClientIP = GetPeerAddress(.Socket)
                                    InvIP = 0
                                    #If CheckIPDupe = True Then
                                        For A = 1 To MaxUsers
                                            With player(A)
                                                If .InUse = True And .ip = ClientIP Then
                                                    InvIP = InvIP + 1
                                                End If
                                            End With
                                        Next A
                                    #End If
                                    
                                    'If InvIP > 3 Then
                                    '    If InvIP > 3 And InvIP < 5 Then
                                    '        'Duplicate IP
                                    '        SendData .Socket, DoubleChar(59) + Chr2(0) + Chr2(0) + "You may not log in multiple times from the same computer!"
                                    '        closesocket .Socket
                                    '    Else
                                    '        closesocket .Socket
                                    '    End If
                                    'Else
                                        If WSAAsyncSelect(.Socket, gHW, ByVal 1028 + NewPlayer, ByVal FD_READ Or FD_WRITE Or FD_CONNECT Or FD_CLOSE) = 0 Then
                                            .InUse = True
                                            .Mode = modeNotConnected
                                            .ip = ClientIP
                                            
                                            .Class = 0
                                            .Level = 0
                                            .Leaving = False
                                            .SocketData = ""
                                            .St = ""
                                            .Frozen = 0
                                            .StatusEffect = 0
                                            .combatTimer = 0
                                            .AttackSkill = 0
                                            .PacketsSent = 0                'Used in checksum
                                            .LastMsg = GetTickCount - 50
                                            .ClientVer = ""
                                            .FloodTimer = 0
                                            .SquelchTimer = 0
                                            .LoginStamp = GetTickCount
                                            .Widgets.NumWidgets = 0
                                            .Widgets.MenuVisible = False
                                            .ProjectileDamage = 0
                                            .ProjectileType = 0
                                            .ProjectileX = 0
                                            .speedhack = 0
                                            .ProjectileY = 0
                                            .alpha = 0
                                            .red = 0
                                            .green = 0
                                            .blue = 0
                                            .GlobalSpellTick = 0
                                            'SEEKHERE
                                            For A = 0 To MAX_SKILLS
                                                .LocalSpellTick(A) = 0
                                            Next A
                                            PrintLog ("Connection accepted from " + .ip)
                                            NumUsers = NumUsers + 1
                                            frmMain.mnuReset.Enabled = False
                                            frmMain.Caption = TitleString + " [" + CStr(NumUsers) + "]"
                                        Else
                                            closesocket .Socket
                                        End If
                                    'End If
                                Else
                                    closesocket .Socket
                                End If
                            End With
                        Else
                            TempSocket = accept(ListeningSocket, Address, sockaddr_size)
                            SendData TempSocket, DoubleChar(2) + Chr2(0) + Chr2(4)
                            closesocket TempSocket
                        End If
                    End If
            End Select
    End Select
    WindowProc = CallWindowProc(lpPrevWndProc, hw, uMsg, wParam, lParam)
End Function

Function ExamineBit(bytByte, Bit) As Byte
    ExamineBit = ((bytByte And (2 ^ Bit)) > 0)
End Function

Sub SetBit(bytByte, Bit)
    bytByte = bytByte Or (2 ^ Bit)
End Sub

Sub ClearBit(bytByte, Bit)
      bytByte = bytByte And Not (2 ^ Bit)
End Sub

Sub CloseClientSocket(Index As Long, Optional forceClose = False)
Dim A As Long, B As Long
    With player(Index)
        If .InUse = True Then
            If .HP >= .MaxHP Or .Access > 0 Or forceClose Then
                If .Trading = True Then
                    CloseTrade Index
                End If
                'Decrement User Num
                NumUsers = NumUsers - 1
                If NumUsers = 0 Then frmMain.mnuReset.Enabled = True
                frmMain.Caption = TitleString + " [" + CStr(NumUsers) + "]"
                
                For A = 1 To MaxPlayerTimers
                    If .ScriptTimer(A) > 0 Then
                        Parameter(0) = Index
                        .ScriptTimer(A) = 0
                        RunScript .Script(A)
                    End If
                Next A
                
                If .Mode = modePlaying Then
                    Parameter(0) = Index
                    .Mode = modeNotConnected
                    If player(Index).Guild > 0 Then
                        UpdateGuildInfo CByte(player(Index).Guild)
                    End If
                    .Mode = modePlaying
                    RunScript ("PARTGAME")
                End If
                
                'Close Socket
                If Not .Socket = INVALID_SOCKET Then
                    closesocket .Socket
                    .Socket = INVALID_SOCKET
                End If
                
                If Not .Level = 0 Then
                    If .Status = 2 Then .Status = 0
                    SavePlayerData Index
                End If
                
                PrintLog "Connection closed from " + .ip + " [" + player(Index).Name + "]"
                
                If Index = currentMaxUser Then
                    For A = 1 To MaxUsers
                        If IsPlaying(A) Then
                            If A <> Index Then currentMaxUser = A
                        End If
                    Next A
                End If
                'Clear Socket Data
                .InUse = False
                .ip = ""
                .SocketData = ""
                .Class = 0
                .Level = 0
                .user = ""
                .DeferSends = False
                .Name = ""
                .combatTimer = 0
                .Party = 0
                .IParty = 0
                .Frozen = 0
                .FloodTimer = 0
                .SquelchTimer = 0
                .StatusEffect = 0
                .ProjectileDamage = 0
                .ProjectileType = 0
                .ProjectileX = 0
                .ProjectileY = 0
                .alpha = 0
                .red = 0
                .green = 0
                .Buff.Type = 0
                .St = ""
                .blue = 0
                .speedhack = 0
                .AttackCount = 0
                For A = 1 To 37
                .StatusData(A).timer = 0
                Next A
                
                For A = 1 To STORAGEPAGES
                    For B = 1 To 20
                        With .Storage(A, B)
                            .Object = 0
                            .prefix = 0
                            .suffix = 0
                            .prefixVal = 0
                            .SuffixVal = 0
                            .Value = 0
                            .Flags(0) = 0
                            .Flags(1) = 0
                            .Flags(2) = 0
                            .Flags(3) = 0
                            .ObjectColor = 0
                        End With
                    Next B
                Next A
                'Send Quit Message
                If .Mode = modePlaying Then
                    SendAll Chr2(7) + Chr2(Index)
                    If .map > 0 Then
                        Partmap Index
                        .map = 0
                    End If
                End If
                .Mode = modeNotConnected
            Else
                AddSocketQue Index, 1
            End If
        End If
    End With
End Sub

Sub CreateDatabase()

    'Create Database
    #If PublicServer = True Then
        Set db = WS.CreateDatabase("server.dat", ";pwd=" + Chr2(100) + Chr2(114) + Chr2(97) + Chr2(99) + Chr2(111) + dbLangGeneral, dbEncrypt + dbVersion30)
    #Else
        Set db = WS.CreateDatabase("server.dat", dbLangGeneral, dbVersion30)
    #End If
    
    CreateAccountsTable db
    CreateGuildsTable db
    CreateNPCsTable
    CreateMonstersTable
    CreateObjectsTable db
    CreateDataTable
    CreateMapsTable
    CreatePreFixTable
    CreateBansTable
    CreateHallsTable db
    CreateScriptsTable
End Sub
Function DoubleChar(ByVal Num As Long) As String
    DoubleChar = Chr2(Int(Num / 256)) + Chr2(Num Mod 256)
End Function
Function negChar(ByVal Num As Long) As String
    If Num > 255 Then Num = 255
    If Num < -255 Then Num = -255
    negChar = Chr2(Abs(Num))
    If Num < 0 Then
        negChar = negChar + Chr2(1)
    Else
        negChar = negChar + Chr2(0)
    End If
End Function
Function TripleChar(ByVal Num As Long) As String
    TripleChar = Chr2(Int(Num / 65536)) + Chr2(Int((Num Mod 65536) / 256)) + Chr2(Num Mod 256)
End Function
Function QuadChar(ByVal Num As Long) As String
    QuadChar = Chr2(Int(Num / 16777216) Mod 256) + Chr2(Int(Num / 65536) Mod 256) + Chr2(Int(Num / 256) Mod 256) + Chr2(Num Mod 256)
End Function
Function Exists(fileName As String) As Boolean
     Exists = (Dir(fileName) <> "")
End Function
Function GetInt(Chars As String) As Long
    GetInt = CLng(Asc(Mid$(Chars, 1, 1))) * 256& + CLng(Asc(Mid$(Chars, 2, 1)))
End Function
Function GetLong(Chars As String) As Long
    GetLong = CLng(Asc(Mid$(Chars, 1, 1))) * 16777216 + CLng(Asc(Mid$(Chars, 2, 1))) * 65536 + CLng(Asc(Mid$(Chars, 3, 1))) * 256& + CLng(Asc(Mid$(Chars, 4, 1)))
End Function
Sub GetWords(St As String)
    Dim A As Long, B As Long, C As Long
    B = 1
    Erase Word
    For A = 1 To 50
TryAgain:
        C = InStr(B, St, " ")
        If C - B = 0 Then B = B + 1: GoTo TryAgain
        If C <> 0 Then
            Word(A) = Mid$(St, B, C - B)
        Else
            Word(A) = Mid$(St, B, Len(St) - B + 1)
            Exit For
        End If
        B = C + 1
    Next A
End Sub
Sub GetSections(St)
    Dim A As Long, B As Long, C As Long
    B = 1
    Erase Word
    For A = 1 To 10
        C = InStr(B, St, Chr2(0))
        If C - B = 0 Then
            Word(A) = ""
        ElseIf C <> 0 Then
            Word(A) = Mid$(St, B, C - B)
        Else
            Word(A) = Mid$(St, B, Len(St) - B + 1)
            Exit For
        End If
        B = C + 1
    Next A
End Sub
'Function Nick(UserHost As String) As String
'    Dim a As Long
'
'    a = InStr(UserHost, "!")
'    If a > 0 Then
'        Nick = Mid$(UserHost, 1, a - 1)
'    Else
'        Nick = UserHost
'    End If
'End Function
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
Sub SavePlayerData(Index)
    Dim A As Long, B As Long, St As String, C As Long
    'printcrashdebug 1, 1
    'Kill "debugCrash.Log"
    With player(Index)
        UserRS.Index = "User"
        UserRS.Seek "=", .user
        UserRS.Edit
        
        .Bookmark = UserRS.Bookmark
        UserRS!Access = .Access
        
        'Character Data
        UserRS!CharNum = .CharNum
        UserRS!Name = .Name
 '       printcrashdebug 15, 1

        UserRS!Class = .Class
        UserRS!Gender = .Gender
        UserRS!sprite = .sprite
        'UserRS!desc = .desc
        
  '      printcrashdebug 15, 2
        UserRS!Squelched = .Squelched
        
 '       printcrashdebug 15, 3
        'Position Data
        UserRS!map = .map
        UserRS!x = .x
        UserRS!y = .y
        UserRS!D = .D
        
    '    printcrashdebug 15, 4
        'Character Vital Stats
 '       printcrashdebug 16, 1
        UserRS!MaxHP = .OldHP
 '       printcrashdebug 16, 2
        UserRS!MaxEnergy = .OldEnergy
 '       printcrashdebug 16, 3
        UserRS!MaxMana = .OldMana
 '       printcrashdebug 16, 4
        UserRS!HP = .HP
  '      printcrashdebug 16, 5
        UserRS!Energy = .Energy
 '       printcrashdebug 16, 6
        UserRS!Mana = .Mana
 '       printcrashdebug 16, 7
 '       printcrashdebug 17, CLng(.StatusEffect)
        UserRS!StatusEffect = CStr(.StatusEffect) '.StatusEffect
 '       printcrashdebug 15, 5
        For A = 1 To MAXSTATUS
            UserRS.Fields("StatusData" + CStr(A)).Value = QuadChar(.StatusData(A).timer) + Chr2(.StatusData(A).Data(0)) + Chr2(.StatusData(A).Data(1)) + Chr2(.StatusData(A).Data(2)) + Chr2(.StatusData(A).Data(3))
        Next A
        
  '      printcrashdebug 15, 6
        'Character Physical Stats
        UserRS!strength = .OldStrength
        UserRS!Agility = .OldAgility
        UserRS!Endurance = .OldEndurance
        UserRS!Wisdom = .OldWisdom
        UserRS!Constitution = .OldConstitution
        UserRS!Intelligence = .OldIntelligence
        
    '    printcrashdebug 15, 7
        UserRS!Level = .Level
        UserRS!Experience = .Experience
        UserRS!StatPoints = .StatPoints
        
     '   printcrashdebug 15, 8
        'Misc. Data
        UserRS!Bank = .Bank
        UserRS!Status = .Status
        UserRS!SkillPoints = .SkillPoints
        UserRS!Renown = .Renown
        
     '   printcrashdebug 15, 9
        St = ""
        For A = 1 To 255
            St = St + Chr2(.SkillLevel(A))
        Next A
        
   '     printcrashdebug 15, 10
        UserRS!SkillLevels = St
        St = ""
        For A = 1 To 255
            St = St + QuadChar(.SkillEXP(A))
        Next A
   '     printcrashdebug 15, 11
        
        UserRS!SkillEXP = St
        UserRS!lastplayed = CLng(Date)
        
   '     printcrashdebug 15, 12
        
        'Inventory Data
        For A = 1 To 20
            St = ""
            With .Inv(A)
                St = DoubleChar(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal)
                For B = 0 To 3
                    St = St + QuadChar(.Flags(B))
                Next B
                St = St + Chr2(.ObjectColor)
            End With
            UserRS.Fields("InvObject" + CStr(A)).Value = St
        Next A
   '     printcrashdebug 15, 13
        
        For A = 1 To STORAGEPAGES
            St = ""
            For B = 1 To 20
                With .Storage(A, B)
                    St = St + DoubleChar(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal)
                End With
            Next B
            For B = 1 To 20
                With .Storage(A, B)
                    St = St + Chr2(.Affix) + Chr2(.AffixVal)
                End With
            Next B
            For B = 1 To 20
                With .Storage(A, B)
                    For C = 0 To 3
                        St = St + QuadChar(.Flags(C))
                    Next C
                End With
            Next B
            For B = 1 To 20
                With .Storage(A, B)
                    St = St + Chr2(.ObjectColor)
                End With
            Next B
            UserRS.Fields("StorageObjects" + CStr(A)).Value = St
        Next A
        
     '   printcrashdebug 15, 14
        
        UserRS!NumStoragePages = .NumStoragePages
        
        'Equipped Objects
        For A = 1 To 5
            St = ""
            With .Equipped(A)
                St = DoubleChar(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal)
                For B = 0 To 3
                    St = St + QuadChar(.Flags(B))
                Next B
                St = St + Chr2(.ObjectColor)
            End With
            UserRS.Fields("Equipped" + CStr(A)).Value = St
        Next A
        
   '     printcrashdebug 15, 15
        
        'Flags
        St = ""
        For A = 0 To 255
            With .Flag(A)
                St = St + QuadChar(.Value) + QuadChar(.ResetCounter)
            End With
        Next A
        UserRS!Flags = St
        
     '   printcrashdebug 15, 16
        
        UserRS.Update
        
      '  printcrashdebug 15, 17
    End With
End Sub

Sub ShutdownServer()
    Dim A As Long
    For A = 1 To currentMaxUser
        If player(A).InUse = True Then
            If IsPlaying(A) Then SavePlayerData A
            CloseClientSocket A
        End If
    Next A
    SaveFlags
    SaveObjects
    
    UserRS.Close
    GuildRS.Close
    NPCRS.Close
    MonsterRS.Close
    ObjectRS.Close
    PrefixRS.Close
    DataRS.Close
    MapRS.Close
    BanRS.Close
    db.Close
    WS.Close
    If ListeningSocket <> INVALID_SOCKET Then
        closesocket ListeningSocket
    End If
    EndWinsock
    Unhook
    Unload frmMain
    Unload frmLoading
    Unload frmOptions
End Sub
Sub SendToMapAllBut(ByVal mapNum As Long, ByVal Index As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            
            If .Mode = modePlaying And .map = mapNum And Index <> A Then
                .St = .St + DoubleChar(Len(St)) + St
                If (Not .DeferSends) Then
                    If SendData(.Socket, .St) = SOCKET_ERROR Then
                        'CloseClientSocket Index
                    End If
                    .St = ""
                End If
            End If
            
        End With
    Next A
End Sub
Sub SendSocket(ByVal Index As Long, ByVal St As String)
    With player(Index)
        If .InUse = True Then
            .St = .St + DoubleChar(Len(St)) + St
            If Not .DeferSends Then
                If SendData(.Socket, .St) = SOCKET_ERROR Then
                    'CloseClientSocket Index
                End If
                .St = ""
            End If
        End If
    End With
End Sub

Sub SendIp(ip As String, ByVal St As String)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .InUse = True Then
                If .ip = ip Then
                .St = .St + DoubleChar(Len(St)) + St
                If Not .DeferSends Then
                    If SendData(.Socket, .St) = SOCKET_ERROR Then
                        'CloseClientSocket Index
                    End If
                    .St = ""
                End If
                End If
            End If
        End With
    Next A
End Sub

Sub FlushSocket(ByVal Index As Long)
    With player(Index)
        If .InUse = True Then
            'If Len(.St) > 1024 Then
            '.St = .St
            'End If
            If Not .DeferSends Then
                If SendData(.Socket, .St) = SOCKET_ERROR Then
                    'CloseClientSocket Index
                End If
                .St = ""
            End If
        End If
    End With
End Sub

Sub SendSocket2(ByVal Index As Long, ByVal St As String)
    With player(Index)
        If .InUse = True Then
            If Len(St) > 65536 Then
                .St = .St + DoubleChar(65535) + QuadChar(Len(St)) + St
            Else
                .St = .St + DoubleChar(Len(St)) + St
            End If
        End If
    End With
End Sub
Sub SendToMapAllBut2(ByVal mapNum As Long, ByVal Index As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode = modePlaying And .map = mapNum And Index <> A Then
                .St = .St + DoubleChar(Len(St)) + St
            End If
        End With
    Next A
End Sub
Sub SendToPartyAllBut2(ByVal PartyNum As Long, ByVal Index As Long, ByVal St As String)
    Dim A As Long
    If player(Index).Party > 0 Then
        For A = 1 To currentMaxUser
            With player(A)
                If .Mode = modePlaying And .Party = PartyNum And Index <> A Then
                    .St = .St + DoubleChar(Len(St)) + St
                End If
            End With
        Next A
    End If
End Sub

Sub SendRaw(ByVal Index As Long, ByVal St As String)
    With player(Index)
        If .InUse = True Then
            .St = .St + St
            If Not .DeferSends Then
                If SendData(.Socket, .St) = SOCKET_ERROR Then
                    'CloseClientSocket Index
                End If
                .St = ""
            End If
        End If
    End With
End Sub

Sub SendRaw2(ByVal Index As Long, ByVal St As String)
    With player(Index)
        If .InUse = True Then
            .St = .St + St
        End If
    End With
End Sub

Sub SendToPartyAllBut(ByVal PartyNum As Long, ByVal Index As Long, ByVal St As String)
    If player(Index).Party > 0 Then
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode = modePlaying And .Party = PartyNum And Index <> A Then
                .St = .St + DoubleChar(Len(St)) + St
                If Not .DeferSends Then
                    If SendData(.Socket, .St) = SOCKET_ERROR Then
                        'CloseClientSocket Index
                    End If
                    .St = ""
                End If
            End If
        End With
    Next A
    End If
End Sub

Sub PrintList(St)
    With frmMain.lstLog
        .AddItem St
        If .ListCount > 200 Then .RemoveItem 0
        If .ListIndex = .ListCount - 2 Then .ListIndex = .ListCount - 1
    End With
End Sub

Sub PrintLog(St)
    With frmMain.lstLog
        .AddItem St
        If .ListCount > 200 Then .RemoveItem 0
        If .ListIndex = .ListCount - 2 Then .ListIndex = .ListCount - 1
        
        Open "debug.log" For Append As #1
        Print #1, St
        Close #1
    End With
End Sub

Function CheckBan(Index As Long) As Boolean
Dim BanNum As Long, banned As Boolean
CheckBan = False

BanNum = FindBan(Index)
If BanNum > 0 Then
    With Ban(BanNum)
        If CLng(Date) >= .UnbanDate Then
            'Unban
            .user = ""
            BanRS.Seek "=", BanNum
            If BanRS.NoMatch = False Then
                If BanRS!Number > 0 Then
                    Ban(BanRS!Number).InUse = False
                    Ban(BanRS!Number).Banner = ""
                    Ban(BanRS!Number).ip = ""
                    Ban(BanRS!Number).Name = ""
                    Ban(BanRS!Number).reason = ""
                    Ban(BanRS!Number).iniUID = ""
                    Ban(BanRS!Number).uId = ""
                    Ban(BanRS!Number).UnbanDate = 0
                    Ban(BanRS!Number).user = ""
                End If
                BanRS.Delete
            End If
            banned = False
            Exit Function
        End If
    End With
    SendSocket Index, Chr2(0) + Chr2(3) + QuadChar(Ban(BanNum).UnbanDate) + Ban(BanNum).reason
    CheckBan = True
    'CloseClientSocket Index
End If
End Function

Function AddSocketQue(Index As Long, KickDelay As Long) As Integer
Dim A As Integer
    player(Index).Leaving = True
For A = 1 To MaxUsers
    If CloseSocketQueue(A).user = Index Then
        Exit Function
    End If
Next A

For A = 1 To MaxUsers
    If CloseSocketQueue(A).user = 0 Then
        CloseSocketQueue(A).user = Index
        CloseSocketQueue(A).TimeStamp = GetTickCount + KickDelay
        Exit For
    End If
Next A
End Function

Sub GiveStartingEQ(Index As Long)
    Dim A As Long, B As Long, C As Long
    If Index >= 1 And Index <= MaxUsers Then
    
    With player(Index)
    For A = 1 To 8
        If World.StartObjects(A) > 0 Then
            B = World.StartObjects(A)
            C = World.StartObjValues(A)
            .Inv(A).Object = B

        Select Case Object(B).Type
            Case 1, 2, 3, 4 'Weapon, Shield, Armor, Helmut
                .Inv(A).Value = CLng(Object(B).Data(0)) * 10
            Case 6, 11 'Money
                .Inv(A).Value = C
            Case 8 'Ring
                .Inv(A).Value = CLng(Object(B).Data(1)) * 10
            Case Else
                .Inv(A).Value = 0
        End Select
        End If
    Next A
    End With
    End If
End Sub
Function GetRepairCost(Index As Long, Slot As Integer) As Long
Dim A As Long, B As Long, C As Long
If Index >= 1 And Index <= MaxUsers Then
    If Slot >= 0 And Slot <= 20 Then
        Select Case Object(player(Index).Inv(Slot).Object).Type
            Case 1, 2, 3, 4, 8 'Weapon, Shield, Armor, Helmet, Ring
                A = Object(player(Index).Inv(Slot).Object).Type
            Case Else
                A = 0
        End Select
    End If
        
        If A > 0 Then
            Select Case A
                Case 1, 2, 3, 4 'Weapon, Shield, Armor, Helmet
                    C = Object(player(Index).Inv(Slot).Object).Data(0) - (player(Index).Inv(Slot).Value / 10)
                    B = B + (C * Cost_Per_Durability)
                    B = B + (Object(player(Index).Inv(Slot).Object).Data(1) * Cost_Per_Strength)
                    GetRepairCost = B
                    Exit Function
                Case 8 'Ring
                    C = Object(player(Index).Inv(Slot).Object).Data(1) - (player(Index).Inv(Slot).Value / 10)
                    B = B + (C * Cost_Per_Durability)
                    B = B + (Object(player(Index).Inv(Slot).Object).Data(2) * Cost_Per_Modifier)
                    GetRepairCost = B
                    Exit Function
            End Select
        Else
            GetRepairCost = 0
        End If
End If
End Function

Sub SendAll(ByVal St As String)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode = modePlaying Then
                .St = .St + DoubleChar(Len(St)) + St
                If Not .DeferSends Then
                    If SendData(.Socket, .St) = SOCKET_ERROR Then
                        'CloseClientSocket Index
                    End If
                    .St = ""
                End If
            End If
        End With
    Next A
End Sub
Sub SendAll2(ByVal St As String)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode = modePlaying Then
                .St = .St + DoubleChar(Len(St)) + St
            End If
        End With
    Next A
End Sub
Sub SendToConnected(ByVal St As String)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode > 0 Then
                .St = .St + DoubleChar(Len(St)) + St
                If Not .DeferSends Then
                    If SendData(.Socket, .St) = SOCKET_ERROR Then
                        'CloseClientSocket Index
                    End If
                    .St = ""
                End If
            End If
        End With
    Next A
End Sub
Sub SendAllBut(ByVal Index As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode = modePlaying And A <> Index Then
                .St = .St + DoubleChar(Len(St)) + St
                If Not .DeferSends Then
                    If SendData(.Socket, .St) = SOCKET_ERROR Then
                        'CloseClientSocket Index
                    End If
                    .St = ""
                End If
            End If
        End With
    Next A
End Sub
Sub SendAllBut2(ByVal Index As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode = modePlaying And A <> Index Then
                .St = .St + DoubleChar(Len(St)) + St
            End If
        End With
    Next A
End Sub
Sub SendAllButRaw(ByVal Index As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode = modePlaying And A <> Index Then
                .St = .St + St
                If Not .DeferSends Then
                    If SendData(.Socket, .St) = SOCKET_ERROR Then
                        'CloseClientSocket Index
                    End If
                    .St = ""
                End If
            End If
        End With
    Next A
End Sub
Sub SendAllButBut(ByVal Index1 As Long, ByVal Index2 As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode = modePlaying And A <> Index1 And A <> Index2 Then
                .St = .St + DoubleChar(Len(St)) + St
                If Not .DeferSends Then
                    If SendData(.Socket, .St) = SOCKET_ERROR Then
                        'CloseClientSocket Index
                    End If
                    .St = ""
                End If
            End If
        End With
    Next A
End Sub

Sub SendToGods(ByVal St As String)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode = modePlaying And .Access > 0 Then
                .St = .St + DoubleChar(Len(St)) + St
                If Not .DeferSends Then
                    If SendData(.Socket, .St) = SOCKET_ERROR Then
                        'CloseClientSocket Index
                    End If
                    .St = ""
                End If
            End If
        End With
    Next A
End Sub

Sub SendToGods2(ByVal St As String)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode = modePlaying And .Access > 0 Then
                .St = .St + DoubleChar(Len(St)) + St
            End If
        End With
    Next A
End Sub

Sub SendToGodsAllBut(Index As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode = modePlaying And .Access > 0 And Index <> A Then
                .St = .St + DoubleChar(Len(St)) + St
                If Not .DeferSends Then
                    If SendData(.Socket, .St) = SOCKET_ERROR Then
                        'CloseClientSocket Index
                    End If
                    .St = ""
                End If
            End If
        End With
    Next A
End Sub

Sub SendToMap(ByVal mapNum As Long, ByVal St As String)
    Dim A As Long
    If map(mapNum).NumPlayers > 0 Then
        For A = 1 To currentMaxUser
            With player(A)
                If .Mode = modePlaying And .map = mapNum Then
                    .St = .St + DoubleChar(Len(St)) + St
                    If Not .DeferSends Then
                        If SendData(.Socket, .St) = SOCKET_ERROR Then
                            'CloseClientSocket Index
                        End If
                        .St = ""
                    End If
                End If
            End With
        Next A
    End If
End Sub
Sub SendToMap2(ByVal mapNum As Long, ByVal St As String)
    Dim A As Long
    If map(mapNum).NumPlayers > 0 Then
        For A = 1 To currentMaxUser
            With player(A)
                If .Mode = modePlaying And .map = mapNum Then
                    .St = .St + DoubleChar(Len(St)) + St
                End If
            End With
        Next A
    End If
End Sub
Sub SendToMapRaw(ByVal mapNum As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode = modePlaying And .map = mapNum Then
                .St = .St + St
                If Not .DeferSends Then
                    If SendData(.Socket, .St) = SOCKET_ERROR Then
                        'CloseClientSocket Index
                    End If
                    .St = ""
                End If
            End If
        End With
    Next A
End Sub
Sub SendToMapRaw2(ByVal mapNum As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode = modePlaying And .map = mapNum Then
                .St = .St + St
            End If
        End With
    Next A
End Sub
Sub SendToZone(ByVal ZoneNum As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode = modePlaying Then
                If map(.map).Zone = ZoneNum Then
                    .St = .St + DoubleChar(Len(St)) + St
                    If Not .DeferSends Then
                        If SendData(.Socket, .St) = SOCKET_ERROR Then
                            'CloseClientSocket Index
                        End If
                        .St = ""
                    End If
                End If
            End If
        End With
    Next A
End Sub
Sub SendToZone2(ByVal ZoneNum As Long, ByVal St As String)
    Dim A As Long
    For A = 1 To currentMaxUser
        With player(A)
            If .Mode = modePlaying Then
                If map(.map).Zone = ZoneNum Then
                    .St = .St + DoubleChar(Len(St)) + St
                End If
            End If
        End With
    Next A
End Sub

Public Sub AddGod(Index As Long, Access As Long)
Dim A As Long, B As Long
A = Index
B = Access
    If A >= 1 And A <= MaxUsers And B >= 0 And B <= 11 Then
        With player(A)
            If .Access <= 11 Then
                .Access = B
                SendSocket A, Chr2(65) + Chr2(B)
                If .Access > 0 Then
                    SendAllBut A, Chr2(91) + Chr2(A) + Chr2(3)
                    .Status = 3
                Else
                    SendAllBut A, Chr2(91) + Chr2(A) + Chr2(0)
                    player(A).Status = 0
                End If
            Else
                With player(A)
                    .Access = 0
                    SendSocket A, Chr2(65) + Chr2(0)
                    SendAllBut A, Chr2(91) + Chr2(A) + Chr2(0)
                    .Status = 0
                End With
            End If
        End With
    End If
End Sub

Sub MoveObject(Index As Long, Slot As Byte, ByRef Item As InvObject)
    Dim A As Byte
    With player(Index)
        With .Inv(Slot)
            .Object = Item.Object
            .prefix = Item.prefix
            .prefixVal = Item.prefix
            .suffix = Item.suffix
            .SuffixVal = Item.SuffixVal
            .Value = Item.Value
            For A = 0 To 3
                .Flags(A) = Item.Flags(A)
            Next A
            .ObjectColor = Item.ObjectColor
        End With
        SendSocket Index, Chr2(18) + Chr2(Slot)
        SendSocket Index, Chr2(17) + Chr2(Slot) + DoubleChar(Item.Object) + QuadChar(Item.Value) + Chr2(Item.prefix) + Chr2(Item.prefixVal) + Chr2(Item.suffix) + Chr2(Item.SuffixVal) + Chr2(Item.Affix) + Chr2(Item.AffixVal) + Chr2(Item.ObjectColor)
        Item.Object = 0
    End With
End Sub

Sub EquipObject(Index As Long, ItemNum As Long, EquipNum As Long)
Dim A As Long, B As Long, C As Long, D As Long, E As Long, F As Long, G As Long, H As Long, i As Long, J As Long, k As Long, L As Long, M As Long
    With player(Index).Inv(ItemNum)
        SendSocket Index, Chr2(18) + Chr2(ItemNum)
        If player(Index).Equipped(EquipNum).Object > 0 Then
            'Player already has this item equipped, so remove it
            A = player(Index).Equipped(EquipNum).Object
            B = player(Index).Equipped(EquipNum).prefix
            C = player(Index).Equipped(EquipNum).prefixVal
            D = player(Index).Equipped(EquipNum).suffix
            E = player(Index).Equipped(EquipNum).SuffixVal
            F = player(Index).Equipped(EquipNum).Value
            G = player(Index).Equipped(EquipNum).Affix
            H = player(Index).Equipped(EquipNum).AffixVal
            i = player(Index).Equipped(EquipNum).Flags(0)
            J = player(Index).Equipped(EquipNum).Flags(1)
            k = player(Index).Equipped(EquipNum).Flags(2)
            L = player(Index).Equipped(EquipNum).Flags(3)
            M = player(Index).Equipped(EquipNum).ObjectColor
            SendSocket Index, Chr2(17) + Chr2(ItemNum) + DoubleChar(CInt(A)) + QuadChar(F) + Chr2(B) + Chr2(C) + Chr2(D) + Chr2(E) + Chr2(G) + Chr2(H) + Chr2(M)
        End If

        SendSocket Index, Chr2(17) + Chr2(20 + EquipNum) + DoubleChar(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)
        player(Index).Equipped(EquipNum).Object = .Object
        player(Index).Equipped(EquipNum).prefix = .prefix
        player(Index).Equipped(EquipNum).prefixVal = .prefixVal
        player(Index).Equipped(EquipNum).suffix = .suffix
        player(Index).Equipped(EquipNum).SuffixVal = .SuffixVal
        player(Index).Equipped(EquipNum).Affix = .Affix
        player(Index).Equipped(EquipNum).AffixVal = .AffixVal
        player(Index).Equipped(EquipNum).Value = .Value
        player(Index).Equipped(EquipNum).Flags(0) = .Flags(0)
        player(Index).Equipped(EquipNum).Flags(1) = .Flags(1)
        player(Index).Equipped(EquipNum).Flags(2) = .Flags(2)
        player(Index).Equipped(EquipNum).Flags(3) = .Flags(3)
        player(Index).Equipped(EquipNum).ObjectColor = .ObjectColor
        If A > 0 Then
            .Object = A
            .prefix = B
            .prefixVal = C
            .suffix = D
            .SuffixVal = E
            .Value = F
            .Affix = G
            .AffixVal = H
            .Flags(0) = i
            .Flags(1) = J
            .Flags(2) = k
            .Flags(3) = L
            .ObjectColor = M
        Else
            .Object = 0
        End If
    End With

    CalculateStats Index
End Sub

Function RemoveObject(Index As Long, EquipNum As Long) As Long
Dim A As Long
    A = FreeInvNum(Index)
    If A > 0 Then
        With player(Index).Equipped(EquipNum)
            player(Index).Inv(A).Object = .Object
            player(Index).Inv(A).prefix = .prefix
            player(Index).Inv(A).prefixVal = .prefixVal
            player(Index).Inv(A).suffix = .suffix
            player(Index).Inv(A).SuffixVal = .SuffixVal
            player(Index).Inv(A).Affix = .Affix
            player(Index).Inv(A).AffixVal = .AffixVal
            player(Index).Inv(A).Value = .Value
            player(Index).Inv(A).Flags(0) = .Flags(0)
            player(Index).Inv(A).Flags(1) = .Flags(1)
            player(Index).Inv(A).Flags(2) = .Flags(2)
            player(Index).Inv(A).Flags(3) = .Flags(3)
            player(Index).Inv(A).ObjectColor = .ObjectColor
            .Object = 0
            SendSocket Index, Chr2(18) + Chr2(20 + EquipNum)
            SendSocket Index, Chr2(17) + Chr2(A) + DoubleChar(player(Index).Inv(A).Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)
        End With
        'MoveObject Index, CByte(A), Player(Index).Equipped(EquipNum)
        SendSocket Index, Chr2(20) + Chr2(EquipNum)
        RemoveObject = 1
    Else
        RemoveObject = 0
    End If
    
    CalculateStats Index
End Function


Sub CalculateStats(Index As Long, Optional UpdateClient As Boolean = True)
Dim A As Long, StrMod As Long, AgiMod As Long, EndMod As Long, B As Long
Dim ConMod As Long, WisMod As Long, IntMod As Long, HpReg As Long
Dim ManaReg As Long, AttMod As Long, MRMod As Double, MagicMod As Long
Dim LeechMod As Long, AttSpdMod As Long, CritMod As Long, PRes As Long

With player(Index)
    .MaxHP = .OldHP
    .MaxEnergy = .OldEnergy
    .MaxMana = .OldMana
    
    .strength = .OldStrength
    StrMod = .StrMod(0)
    .Constitution = .OldConstitution
    ConMod = .ConMod(0)
    .Wisdom = .OldWisdom
    WisMod = .WisMod(0)
    .Agility = .OldAgility
    AgiMod = .AgiMod(0)
    .Endurance = .OldEndurance
    EndMod = .EndMod(0)
    .Intelligence = .OldIntelligence
    IntMod = .IntMod(0)
    
    HpReg = 0
    ManaReg = 0
    MRMod = 0
    MagicMod = 0
    AttSpdMod = 0
    
    .AttackDamageBonus = 0
    .MagicDamageBonus = 0
    .MagicDefenseBonus = 0
    .PhysicalDefenseBonus = 0
    .PhysicalResistanceBonus = 0
    .ShieldBlockBonus = 0
    .DodgeBonus = 0
    
    If .Equipped(EQ_WEAPON).Object > 0 Then 'Attack Speed
        If Object(.Equipped(EQ_WEAPON).Object).Type = 1 Or Object(.Equipped(EQ_WEAPON).Object).Type = 10 Then
              AttMod = 10 + Object(.Equipped(EQ_WEAPON).Object).Data(4)
        End If
    Else
        AttMod = 10
    End If
    AttMod = AttMod - (.Agility / 60)
    'SKILL: FAST ATTACK
    'If CanUseSkill(CByte(Index), SKILL_FASTATTACK) Then
    '    attmod = attmod - (.SkillLevel(SKILL_FASTATTACK) / 10)
    'End If
    
    For A = 1 To 5
        If .Equipped(A).Object > 0 Then
            With .Equipped(A)
                If .prefix > 0 Then
                    Select Case prefix(.prefix).ModType
                          Case 1 'Strength
                                If (.prefixVal And 63) > 0 Then
                                    StrMod = StrMod + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 2 'Agility
                                If (.prefixVal And 63) > 0 Then
                                    AgiMod = AgiMod + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 3 'Endurance
                                If (.prefixVal And 63) > 0 Then
                                    EndMod = EndMod + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 4 'Constitution
                                If (.prefixVal And 63) > 0 Then
                                    ConMod = ConMod + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 5 'Wisdom
                                If (.prefixVal And 63) > 0 Then
                                    WisMod = WisMod + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 6 'Intelligence
                                If (.prefixVal And 63) > 0 Then
                                    IntMod = IntMod + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 7 'HP Regen
                                If (.prefixVal And 63) > 0 Then
                                    HpReg = HpReg + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 8 'Mana Regen
                                If (.prefixVal And 63) > 0 Then
                                    ManaReg = ManaReg + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 9 'Magic Resistance
                                If (.prefixVal And 63) > 0 Then
                                    MRMod = MRMod + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 11 'All Stats
                                If (.prefixVal And 63) > 0 Then
                                    StrMod = StrMod + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                    AgiMod = AgiMod + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                    EndMod = EndMod + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                    ConMod = ConMod + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                    WisMod = WisMod + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                    IntMod = IntMod + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 12 'Increased Magic Find
                                If (.prefixVal And 63) > 0 Then
                                    MagicMod = MagicMod + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 13 'Leech HP
                                If (.prefixVal And 63) > 0 Then
                                    LeechMod = LeechMod + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 14 'Increased Attack Speed
                                If (.prefixVal And 63) > 0 Then
                                    AttSpdMod = AttSpdMod + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))

                                End If
                            Case 15 'Increased Crit
                                If (.prefixVal And 63) > 0 Then
                                    CritMod = CritMod + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 16 'Resist Poison
                                If (.prefixVal And 63) > 0 Then
                                    PRes = PRes + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 18 'Attack Damage
                                If (.prefixVal And 63) > 0 Then
                                    player(Index).AttackDamageBonus = player(Index).AttackDamageBonus + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 19 'Magic Damage
                                If (.prefixVal And 63) > 0 Then
                                    player(Index).MagicDamageBonus = player(Index).MagicDamageBonus + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 20 'Magic Defense
                                If (.prefixVal And 63) > 0 Then
                                    player(Index).MagicDefenseBonus = player(Index).MagicDefenseBonus + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 21 'Physical Defense
                                If (.prefixVal And 63) > 0 Then
                                    player(Index).PhysicalDefenseBonus = player(Index).PhysicalDefenseBonus + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 22 'Physical Resistance
                                If (.prefixVal And 63) > 0 Then
                                    player(Index).PhysicalResistanceBonus = player(Index).PhysicalResistanceBonus + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 23 'visible bonus
                            Case 24 'Shield Block %
                                If (.prefixVal And 63) > 0 Then
                                    player(Index).ShieldBlockBonus = player(Index).ShieldBlockBonus + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 25 'dodge %
                                If (.prefixVal And 63) > 0 Then
                                    player(Index).DodgeBonus = player(Index).DodgeBonus + (.prefixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                                End If
                            Case 26 'no description
                            Case 27 'hidden entirely
                        End Select
                    End If
                    If .suffix > 0 Then
                        Select Case prefix(.suffix).ModType
                    Case 1 'Strength
                        If .SuffixVal > 0 Then
                            StrMod = StrMod + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 2 'Agility
                        If .SuffixVal > 0 Then
                            AgiMod = AgiMod + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 3 'Endurance
                        If .SuffixVal > 0 Then
                            EndMod = EndMod + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 4 'Constitution
                        If .SuffixVal > 0 Then
                            ConMod = ConMod + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 5 'Wisdom
                        If .SuffixVal > 0 Then
                            WisMod = WisMod + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 6 'Intelligence
                        If .SuffixVal > 0 Then
                            IntMod = IntMod + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 7 'HP Regen
                        If .SuffixVal > 0 Then
                            HpReg = HpReg + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 8 'Mana Regen
                        If .SuffixVal > 0 Then
                            ManaReg = ManaReg + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 9 'Magic Resistance
                        If .SuffixVal > 0 Then
                            MRMod = MRMod + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 11 'All Stats
                        If .SuffixVal > 0 Then
                            StrMod = StrMod + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                            AgiMod = AgiMod + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                            EndMod = EndMod + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                            ConMod = ConMod + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                            WisMod = WisMod + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                            IntMod = IntMod + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 12 'Increased Magic Find
                        If .SuffixVal > 0 Then
                            MagicMod = MagicMod + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 13 'Leech HP
                        If .SuffixVal > 0 Then
                            LeechMod = LeechMod + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 14 'Increased Hit Speed
                        If .SuffixVal > 0 Then
                            AttSpdMod = AttSpdMod + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 15 'Increased Crit
                        If .SuffixVal > 0 Then
                            CritMod = CritMod + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 16 'Resist Poison
                        If .SuffixVal > 0 Then
                            PRes = PRes + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 18 'Attack Damage
                        If .SuffixVal > 0 Then
                            player(Index).AttackDamageBonus = player(Index).AttackDamageBonus + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 19 'Magic Damage
                        If .SuffixVal > 0 Then
                            player(Index).MagicDamageBonus = player(Index).MagicDamageBonus + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 20 'Magic Defense
                        If .SuffixVal > 0 Then
                            player(Index).MagicDefenseBonus = player(Index).MagicDefenseBonus + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 21 'Physical Defense
                        If .SuffixVal > 0 Then
                            player(Index).PhysicalDefenseBonus = player(Index).PhysicalDefenseBonus + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 22 'Physical Resistance
                        If .SuffixVal > 0 Then
                            player(Index).PhysicalResistanceBonus = player(Index).PhysicalResistanceBonus + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 23 'visible bonus
                    Case 24 'Shield Block %
                        If .SuffixVal > 0 Then
                            player(Index).ShieldBlockBonus = player(Index).ShieldBlockBonus + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 25 'Dodge %
                        If .SuffixVal > 0 Then
                            player(Index).DodgeBonus = player(Index).DodgeBonus + (.SuffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 26 'no description
                    Case 27 'hidden entirely
                    End Select
                End If
                If .Affix > 0 Then
                        Select Case prefix(.Affix).ModType
                    Case 1 'Strength
                        If .AffixVal > 0 Then
                            StrMod = StrMod + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 2 'Agility
                        If .AffixVal > 0 Then
                            AgiMod = AgiMod + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 3 'Endurance
                        If .AffixVal > 0 Then
                            EndMod = EndMod + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 4 'Constitution
                        If .AffixVal > 0 Then
                            ConMod = ConMod + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 5 'Wisdom
                        If .AffixVal > 0 Then
                            WisMod = WisMod + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 6 'Intelligence
                        If .AffixVal > 0 Then
                            IntMod = IntMod + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 7 'HP Regen
                        If .AffixVal > 0 Then
                            HpReg = HpReg + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 8 'Mana Regen
                        If .AffixVal > 0 Then
                            ManaReg = ManaReg + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 9 'Magic Resistance
                        If .AffixVal > 0 Then
                            MRMod = MRMod + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 11 'All Stats
                        If .AffixVal > 0 Then
                            StrMod = StrMod + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                            AgiMod = AgiMod + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                            EndMod = EndMod + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                            ConMod = ConMod + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                            WisMod = WisMod + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                            IntMod = IntMod + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 12 'Increased Magic Find
                        If .AffixVal > 0 Then
                            MagicMod = MagicMod + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 13 'Leech HP
                        If .AffixVal > 0 Then
                            LeechMod = LeechMod + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 14 'Increased Hit Speed
                        If .AffixVal > 0 Then
                            AttSpdMod = AttSpdMod + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 15 'Increased Crit
                        If .AffixVal > 0 Then
                            CritMod = CritMod + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 16 'Resist Poison
                        If .AffixVal > 0 Then
                            PRes = PRes + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 18 'Attack Damage
                        If .AffixVal > 0 Then
                            player(Index).AttackDamageBonus = player(Index).AttackDamageBonus + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 19 'Magic Damage
                        If .AffixVal > 0 Then
                            player(Index).MagicDamageBonus = player(Index).MagicDamageBonus + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 20 'Magic Defense
                        If .AffixVal > 0 Then
                            player(Index).MagicDefenseBonus = player(Index).MagicDefenseBonus + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 21 'Physical Defense
                        If .AffixVal > 0 Then
                            player(Index).PhysicalDefenseBonus = player(Index).PhysicalDefenseBonus + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 22 'Physical Resistance
                        If .AffixVal > 0 Then
                            player(Index).PhysicalResistanceBonus = player(Index).PhysicalResistanceBonus + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 23 'visible bonus
                    Case 24 'Shield Block %
                        If .AffixVal > 0 Then
                            player(Index).ShieldBlockBonus = player(Index).ShieldBlockBonus + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 25 'Dodge %
                        If .AffixVal > 0 Then
                            player(Index).DodgeBonus = player(Index).DodgeBonus + (.AffixVal And 63) * (1 + 0.5 * player(Index).SkillLevel(SKILL_MAGICALCONDUIT))
                        End If
                    Case 26 'no description
                    Case 27 'hidden entirely
                    End Select
                End If
            End With
        End If
    Next A
    
    'If GetStatusEffect(index, SE_BERSERK) Then StrMod = StrMod + 10
    
    'BUFFPLAYER SCRIPT
        If .StatusData(SE_HPMOD).timer > 0 Then
            .MaxHP = .MaxHP + (.StatusData(SE_HPMOD).Data(0) * 256 + .StatusData(SE_HPMOD).Data(1)) * (1 - 2 * .StatusData(SE_HPMOD).Data(2))
        End If
        If .StatusData(SE_MANAMOD).timer > 0 Then
            .MaxMana = .MaxMana + (.StatusData(SE_MANAMOD).Data(0) * 256 + .StatusData(SE_MANAMOD).Data(1)) * (1 - 2 * .StatusData(SE_MANAMOD).Data(2))
        End If
        If .StatusData(SE_ENERGYMOD).timer > 0 Then
            .MaxEnergy = .MaxEnergy + (.StatusData(SE_ENERGYMOD).Data(0) * 256 + .StatusData(SE_ENERGYMOD).Data(1)) * (1 - 2 * .StatusData(SE_ENERGYMOD).Data(2))
        End If
        
        B = SE_STRENGTHMOD: If .StatusData(B).timer > 0 Then StrMod = StrMod + (.StatusData(B).Data(0) * 256 + .StatusData(B).Data(1)) * (1 - 2 * .StatusData(B).Data(2))
        B = SE_AGILITYMOD: If .StatusData(B).timer > 0 Then AgiMod = AgiMod + (.StatusData(B).Data(0) * 256 + .StatusData(B).Data(1)) * (1 - 2 * .StatusData(B).Data(2))
        B = SE_ENDURANCEMOD: If .StatusData(B).timer > 0 Then EndMod = EndMod + (.StatusData(B).Data(0) * 256 + .StatusData(B).Data(1)) * (1 - 2 * .StatusData(B).Data(2))
        B = SE_WISDOMMOD: If .StatusData(B).timer > 0 Then WisMod = WisMod + (.StatusData(B).Data(0) * 256 + .StatusData(B).Data(1)) * (1 - 2 * .StatusData(B).Data(2))
        B = SE_INTELLIGENCEMOD: If .StatusData(B).timer > 0 Then IntMod = IntMod + (.StatusData(B).Data(0) * 256 + .StatusData(B).Data(1)) * (1 - 2 * .StatusData(B).Data(2))
        B = SE_CONSTITUTIONMOD: If .StatusData(B).timer > 0 Then ConMod = ConMod + (.StatusData(B).Data(0) * 256 + .StatusData(B).Data(1)) * (1 - 2 * .StatusData(B).Data(2))
        B = SE_HPREGENMOD: If .StatusData(B).timer > 0 Then HpReg = HpReg + (.StatusData(B).Data(0) * 256 + .StatusData(B).Data(1)) * (1 - 2 * .StatusData(B).Data(2))
   '''''''''''''     B = SE_ENERGYREGENMOD: If .StatusData(B).timer > 0 Then .ener = StrMod + (.StatusData(B).Data(0) * 256 + .StatusData(B).Data(1)) * (-1 * .StatusData(B).Data(2))
        B = SE_MANAMOD: If .StatusData(B).timer > 0 Then ManaReg = ManaReg + (.StatusData(B).Data(0) * 256 + .StatusData(B).Data(1)) * (1 - 2 * .StatusData(B).Data(2))
        B = SE_MAGICRESISTMOD: If .StatusData(B).timer > 0 Then MRMod = MRMod + (.StatusData(B).Data(0) * 256 + .StatusData(B).Data(1)) * (1 - 2 * .StatusData(B).Data(2))
        B = SE_ATTACKSPEEDMOD: If .StatusData(B).timer > 0 Then AttSpdMod = AttSpdMod + (.StatusData(B).Data(0) * 256 + .StatusData(B).Data(1)) * (1 - 2 * .StatusData(B).Data(2))
        B = SE_CRITICALCHANCEMOD: If .StatusData(B).timer > 0 Then CritMod = CritMod + (.StatusData(B).Data(0) * 256 + .StatusData(B).Data(1)) * (1 - 2 * .StatusData(B).Data(2))
        B = SE_POISONRESISTMOD: If .StatusData(B).timer > 0 Then PRes = PRes + (.StatusData(B).Data(0) * 256 + .StatusData(B).Data(1)) * (1 - 2 * .StatusData(B).Data(2))

        
        
    
    'SKILL: GREAT STRENGTH
    StrMod = StrMod + .SkillLevel(SKILL_GREATSTRENGTH) * 2

    .StrMod(1) = 0
    .AgiMod(1) = 0
    .EndMod(1) = 0
    .ConMod(1) = 0
    .WisMod(1) = 0
    .IntMod(1) = 0
            
'Buffs
    Select Case .Buff.Type
        Case BUFF_EMPOWER
            StrMod = StrMod + .Buff.Data(0)
            .StrMod(1) = .Buff.Data(0)
        Case BUFF_WILLPOWER
            IntMod = IntMod + .Buff.Data(0)
            WisMod = WisMod + .Buff.Data(0)
            .IntMod(1) = .Buff.Data(0)
            .WisMod(1) = .Buff.Data(0)
        Case BUFF_ZEAL
            StrMod = StrMod + .Buff.Data(0)
            AgiMod = AgiMod + .Buff.Data(0)
            EndMod = EndMod + .Buff.Data(0)
            ConMod = ConMod + .Buff.Data(0)
            WisMod = WisMod + .Buff.Data(0)
            IntMod = IntMod + .Buff.Data(0)
            .StrMod(1) = .Buff.Data(0)
            .AgiMod(1) = .Buff.Data(0)
            .EndMod(1) = .Buff.Data(0)
            .ConMod(1) = .Buff.Data(0)
            .WisMod(1) = .Buff.Data(0)
            .IntMod(1) = .Buff.Data(0)
        Case BUFF_VITALITY
            .MaxHP = .MaxHP + .Buff.Data(0)
    End Select
    
    If player(Index).SkillLevel(SKILL_COLOSSUS) Then .MaxHP = .MaxHP + 100
    
    'SKILL: MANA RESERVES
    A = 0
    If player(Index).SkillLevel(SKILL_MANARESERVES) > 0 Then
        Parameter(0) = Index
            A = RunScript("SPELL" & SKILL_MANARESERVES)
        .MaxMana = .MaxMana + A
    End If

    'SKILL: GREAT FORTITUDE
    .MaxHP = .MaxHP + .SkillLevel(SKILL_GREATFORTITUDE) * 10
    HpReg = HpReg + .SkillLevel(SKILL_VIGOR) * 2
    .MaxEnergy = .MaxEnergy + 40 * .SkillLevel(SKILL_RESTLESS)
    
    If StrMod > 255 Then StrMod = 255
    If AgiMod > 255 Then AgiMod = 255
    If EndMod > 255 Then EndMod = 255
    If ConMod > 255 Then ConMod = 255
    If WisMod > 255 Then WisMod = 255
    If IntMod > 255 Then IntMod = 255
    If MRMod > 100 Then MRMod = 100
    If MagicMod > 100 Then MagicMod = 100
    If LeechMod > 100 Then LeechMod = 100
    'If AttSpdMod > 10 Then AttSpdMod = 10
    PRes = PRes + (ConMod / 5)
    If PRes > 100 Then PRes = 100
    If CritMod > 25 Then CritMod = 25
    
    .strength = IIf(.OldStrength + StrMod > 255, 255, IIf(.OldStrength + StrMod < 0, 0, .OldStrength + StrMod))
    If UpdateClient Then SendSocket Index, Chr2(109) + Chr2(1) + Chr2(.OldStrength) + negChar(StrMod)
    
    .Agility = IIf(.OldAgility + AgiMod > 255, 255, IIf(.OldAgility + AgiMod < 0, 0, .OldAgility + AgiMod))
   If UpdateClient Then SendSocket Index, Chr2(109) + Chr2(2) + Chr2(.OldAgility) + negChar(AgiMod)
   
    .Endurance = IIf(.OldEndurance + EndMod > 255, 255, IIf(.OldEndurance + EndMod < 0, 0, .OldEndurance + EndMod))
    If UpdateClient Then SendSocket Index, Chr2(109) + Chr2(3) + Chr2(.OldEndurance) + negChar(EndMod)
    
    .Constitution = IIf(.OldConstitution + ConMod > 255, 255, IIf(.OldConstitution + ConMod < 0, 0, .OldConstitution + ConMod))
    If UpdateClient Then SendSocket Index, Chr2(109) + Chr2(4) + Chr2(.OldConstitution) + negChar(ConMod)
    
    .Wisdom = IIf(.OldWisdom + WisMod > 255, 255, IIf(.OldWisdom + WisMod < 1, 0, .OldWisdom + WisMod))
    If UpdateClient Then SendSocket Index, Chr2(109) + Chr2(5) + Chr2(.OldWisdom) + negChar(WisMod)
    
    .Intelligence = IIf(.OldIntelligence + IntMod > 255, 255, IIf(.OldIntelligence + IntMod < 0, 0, .OldIntelligence + IntMod))
    If UpdateClient Then SendSocket Index, Chr2(109) + Chr2(6) + Chr2(.OldIntelligence) + negChar(IntMod)
    
    
    .MaxMana = .MaxMana + GetBonusPerStat(.Intelligence, ManaPerIntelligence) + GetStatPerBonusHigh(.Wisdom, PietyPerMana)
    .MaxHP = .MaxHP + GetBonusPerStat(.Constitution, HPPerConstitution) + GetStatPerBonusHigh(.Wisdom, PietyPerHP)
    .MaxEnergy = .MaxEnergy + GetStatPerBonus(.Endurance, EndurancePerEnergy)
    
    
    
    If .MaxHP < 1 Then .MaxHP = 1
    If .MaxMana < 1 Then .MaxMana = 1
    If .MaxEnergy < 1 Then .MaxEnergy = 1
    
    If .MaxMana > 9999 Then .MaxMana = 9999
    If A <> 0 And .Mana > .MaxMana Then .Mana = .MaxMana
    
    SetPlayerMana Index, .Mana
    
    If .MaxHP > 9999 Then .MaxHP = 9999
    If .HP > .MaxHP Then .HP = .MaxHP
    SetPlayerHP Index, .HP
    
    If .Energy > 9999 Then .Energy = 9999
    If .Energy > .MaxEnergy Then .Energy = .MaxEnergy
    SetPlayerEnergy Index, .Energy

    If UpdateClient Then SendSocket Index, Chr2(116) + DoubleChar(CLng(.MaxHP)) + DoubleChar(CLng(.MaxMana)) + DoubleChar(CLng(.MaxEnergy))

    .HPRegen = HpReg
    .ManaRegen = ManaReg

    If GetStatusEffect(Index, SE_BERSERK) Then
        
        If AttMod < 14 Then AttMod = AttMod - 1
        If AttMod >= 14 Then AttMod = AttMod - 2

        
        
    End If
   
    If .Access = 0 Then
        AttMod = IIf((AttMod * 100 - AttSpdMod * 30) > 299, (AttMod * 3.33334 - AttSpdMod), 3)
    Else
        AttMod = IIf((AttMod * 100 - AttSpdMod * 30) >= 0, (AttMod * 3.33334 - AttSpdMod), 0)
    End If
    .AttackSpeed = AttMod * 30
    
    If .Equipped(1).Object > 0 And .Equipped(2).Object > 0 Then
        If Object(.Equipped(1).Object).Type = 1 And Object(.Equipped(2).Object).Type = 1 Then .AttackSpeed = .AttackSpeed - 150
    End If
    
    If UpdateClient Then SendSocket Index, Chr2(109) + Chr2(7) + Chr2(AttMod)
    .MagicResist = MRMod + GetStatPerBonus(.Wisdom, PietyPerMagicResist)
    .MagicBonus = MagicMod
    .Leech = LeechMod
    .ResistPoison = IIf(PRes >= 0, PRes, 0)
    .CriticalBonus = IIf(CritMod >= 0, CritMod, 0)
End With

Exit Sub
Error_Handler:
Open App.Path + "/LOG.TXT" For Append As #1
    Print #1, "ARGLEFRASTER3" & player(Index).Name & "/" & Err.Number & "/" & Err.Description & "/" & "  modServer"
Close #1
Unhook
End
End Sub

Public Sub MonsterMove(mapNum As Long, MonsterNum As Long)
Dim A As Long, B As Long
    If mapNum >= 1 And mapNum <= 5000 Then
        For A = 0 To 9
            With map(mapNum).monster(A)
                If .monster > 0 Then
                    If A <> MonsterNum Then
                        If ExamineBit(monster(.monster).Flags, 4) Then
                            If .Target = 0 And .TargType = 0 Then
                                B = Sqr((CLng(.x) - map(mapNum).monster(MonsterNum).x) ^ 2 + (CLng(.y) - map(mapNum).monster(MonsterNum).y) ^ 2)
                                If B <= monster(.monster).Sight Then
                                    .Target = MonsterNum
                                    .TargType = TargTypeMonster
                                End If
                            End If
                        End If
                    End If
                End If
            End With
        Next A
            If map(mapNum).monster(MonsterNum).monster > 0 Then
                With map(mapNum).monster(MonsterNum)
                    For A = 0 To MaxTraps
                        If map(mapNum).trap(A).x = .x And map(mapNum).trap(A).y = .y Then
                            MonsterTriggerTrap mapNum, MonsterNum, A
                        End If
                    Next A
                End With
            End If
    End If
End Sub

Public Sub PoisonMonster(ByVal mapNum As Long, ByVal MonsterNum As Long, PoisonStrength As Long, PoisonLength As Long)
    With map(mapNum).monster(MonsterNum)
        .Poison = PoisonStrength
        .PoisonLength = PoisonLength
    End With
End Sub

Public Sub LogCommand(Index As Long, LogData As String)
On Error Resume Next
    MkDir (App.Path & "\GodLogs")
On Error GoTo 0
Open App.Path + "\GodLogs\" & player(Index).user & ".log" For Append As #1
    Print #1, DateTime.Now, "(" & player(Index).ip & ")", player(Index).Name, LogData & "  modServer"
Close #1
End Sub

Function CanAttack(fX As Long, fY As Long, tX As Long, tY As Long, CurMap As Integer, Optional walkStamp As Long = 0) As Boolean
    If CurMap > 0 Then
    If map(CurMap).Tile(fX, fY).Att = 21 Then
        CanAttack = False
    Else
        If Abs(tX - fX) = 1 And Abs(tY - fY) = 1 Then
            CanAttack = False
            If walkStamp <> 0 Then
                If GetTickCount - walkStamp < 300 Then
                    CanAttack = True
                End If
            End If
        ElseIf Abs(tX - fX) = 2 And Abs(tY - fY) = 0 Then
            CanAttack = False
            If walkStamp <> 0 Then
                If GetTickCount - walkStamp < 300 Then
                    CanAttack = True
                End If
            End If
        ElseIf Abs(tX - fX) = 0 And Abs(tY - fY) = 2 Then
            CanAttack = False
            If walkStamp <> 0 Then
                If GetTickCount - walkStamp < 300 Then
                    CanAttack = True
                End If
            End If
        ElseIf tX = fX And tY = fY Then
            CanAttack = True
        Else
            If tY - fY = -1 Then 'up
                If ExamineBit(map(CurMap).Tile(tX, tY).WallTile, 1) Or ExamineBit(map(CurMap).Tile(fX, fY).WallTile, 4) Then
                    CanAttack = False
                Else
                    CanAttack = True
                End If
            End If
            If tY - fY = 1 Then 'down
                If ExamineBit(map(CurMap).Tile(tX, tY).WallTile, 0) Or ExamineBit(map(CurMap).Tile(fX, fY).WallTile, 5) Then
                    CanAttack = False
                Else
                    CanAttack = True
                End If
            End If
            If tX - fX = -1 Then 'left
                If ExamineBit(map(CurMap).Tile(tX, tY).WallTile, 3) Or ExamineBit(map(CurMap).Tile(fX, fY).WallTile, 6) Then
                    CanAttack = False
                Else
                    CanAttack = True
                End If
            End If
            If tX - fX = 1 Then 'right
                If ExamineBit(map(CurMap).Tile(tX, tY).WallTile, 2) Or ExamineBit(map(CurMap).Tile(fX, fY).WallTile, 7) Then
                    CanAttack = False
                Else
                    CanAttack = True
                End If
            End If
        End If
    End If
    End If
End Function
Function FreeMap() As Long
    Dim A As Long
    For A = 1 To 5000
        MapRS.Seek "=", A
        If MapRS.NoMatch = True Then
            FreeMap = A
            Exit Function
        End If
    Next A
End Function
Sub EquateLight(Index As Long)
    Dim A As Long, B As Long, C As Long
    With player(Index)
        .Light.Intensity = 0
        .Light.Radius = 0
        'PREFIXES
        For B = 1 To 5
            A = .Equipped(B).Object
            If A > 0 Then
                If .Equipped(B).prefix > 0 Then
                    If prefix(.Equipped(B).prefix).Light.Intensity > 0 Then
                        C = .Light.Intensity
                        C = C + prefix(.Equipped(B).prefix).Light.Intensity
                        If C > 255 Then C = 255
                        .Light.Intensity = C
                        .Light.Radius = .Light.Radius + prefix(.Equipped(B).prefix).Light.Radius
                    End If
                End If
                If .Equipped(B).suffix > 0 Then
                    If prefix(.Equipped(B).suffix).Light.Intensity > 0 Then
                        C = .Light.Intensity
                        C = C + prefix(.Equipped(B).suffix).Light.Intensity
                        If C > 255 Then C = 255
                        .Light.Intensity = C
                        '.Light.Intensity = 0
                        .Light.Radius = .Light.Radius + prefix(.Equipped(B).suffix).Light.Radius
                        '.Light.Radius = 0
                    End If
                End If
            End If
        Next B
    End With
End Sub

Public Function IsBlocked(mapNum As Long, x As Long, y As Long) As Boolean
 If x < 0 Or x > 11 Or y < 0 Or y > 11 Then Exit Function
Dim A As Long
 A = map(mapNum).Tile(x, y).Att
 If A = 1 Or A = 2 Or A = 3 Or A = 10 Then
    IsBlocked = True
 End If
End Function

Public Function IsValidMonster(mapNum As Long, monster As Long) As Boolean
        If map(mapNum).monster(monster).monster > 0 Then
                IsValidMonster = True
        End If
End Function
Function FreeTradeSlot(Index As Long) As Byte
Dim A As Long, B As Long
B = 5
For A = 1 To 10
    If player(Index).Trade.Slot(A) = 0 Then
        If FreeInvSlotCount(player(Index).Trade.Trader) - 1 >= A Then
            B = A
            Exit For
        Else
            B = 11
            Exit For
        End If
    End If
Next A
FreeTradeSlot = B
End Function

Function FreeInvTradeSlotCount(Index As Long) As Byte
    Dim A As Long, B As Long, C As Long, D As Long
    
    D = 0
    With player(Index)
        For A = 1 To 20
            C = 0
            For B = 1 To 10
                If .Trade.Slot(B) = A Then
                    If Object(.Inv(.Trade.Slot(B)).Object).Type <> 6 And Object(.Inv(.Trade.Slot(B)).Object).Type <> 11 Then
                        D = D + 1
                        C = 1
                        Exit For
                    End If
                End If
            Next B
            If C = 0 Then
                If .Inv(A).Object = 0 Then
                    D = D + 1
                End If
            End If
        Next A
    End With
    FreeInvTradeSlotCount = D
    PrintLog player(Index).Name & ": " & D
End Function

Function FreeInvSlotCount(Index As Long) As Byte
Dim A As Long, B As Long
B = 0
For A = 1 To 20
    If player(Index).Inv(A).Object = 0 Then
        B = B + 1
    End If
Next A
FreeInvSlotCount = B
End Function
Sub CloseTrade(Index As Long)
    Dim Index2 As Long, A As Long
    Index2 = player(Index).Trade.Trader
    If Index2 > 0 And Index2 <= MaxUsers Then
        player(Index2).Trade.Trader = 0
        player(Index2).Trade.State = 0
        For A = 1 To 10
            player(Index2).Trade.Slot(A) = 0
            ClearObject player(Index2).Trade.Item(A)
        Next A
        player(Index2).Trading = False
        SendSocket Index2, Chr2(113) + Chr2(2)
    End If

    player(Index).Trade.Trader = 0
    player(Index).Trade.State = 0
    For A = 1 To 10
        player(Index).Trade.Slot(A) = 0
        ClearObject player(Index).Trade.Item(A)
    Next A
    player(Index).Trading = False
    SendSocket Index, Chr2(113) + Chr2(2)
End Sub

Sub WarpPlayer(ByVal Index As Long, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long)
    If Index >= 1 And Index <= MaxUsers And mapNum >= 1 And mapNum <= 5000 And x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
    Dim ST1 As String
        With player(Index)
        
            If .Mode = modePlaying Then
                If mapNum = .map Then
                    .x = x
                    .y = y
                    If (.alpha <> 0 Or .red <> 0 Or .blue <> 0 Or .green <> 0) Then
                        ST1 = Chr2(8) + Chr2(Index) + Chr2(x) + Chr2(y) + Chr2(.D) + Chr2(.red) + Chr2(.green) + Chr2(.blue) + Chr2(.alpha)
                    Else
                        ST1 = Chr2(8) + Chr2(Index) + Chr2(x) + Chr2(y) + Chr2(.D)
                    End If
                    
                    If .Equipped(1).Object > 0 Then
                        ST1 = ST1 + DoubleChar(.Equipped(1).Object)
                    Else
                        ST1 = ST1 + DoubleChar(0)
                    End If
                    If .Equipped(2).Object > 0 Then
                        ST1 = ST1 + DoubleChar(.Equipped(2).Object)
                    Else
                        ST1 = ST1 + DoubleChar(0)
                    End If
                    If .Equipped(3).Object > 0 Then
                        ST1 = ST1 + DoubleChar(.Equipped(3).Object)
                    Else
                        ST1 = ST1 + DoubleChar(0)
                    End If
                    If .Equipped(4).Object > 0 Then
                        ST1 = ST1 + DoubleChar(.Equipped(4).Object)
                    Else
                        ST1 = ST1 + DoubleChar(0)
                    End If

                    
                    SendToMapAllBut mapNum, Index, ST1
                    
                    SendSocket Index, Chr2(8) + Chr2((x * 16) + y)
                Else
                    Partmap (Index)
                    .map = mapNum
                    .x = x
                    .y = y
                    JoinMap (Index)
                End If
            End If
        End With
    End If
End Sub

Function ClipString(St As String) As String
    Dim A As Long
    For A = Len(St) To 1 Step -1
        If Mid$(St, A, 1) <> Chr2(32) And Mid$(St, A, 1) <> Chr2(0) Then
            ClipString = Mid$(St, 1, A)
            Exit Function
        End If
    Next A
End Function

Sub SwapObject(Index As Long, Slot1 As Long, ByRef Item1 As InvObject, Slot2 As Long, ByRef Item2 As InvObject)  'Slot1 As Byte, Slot2 As Byte)
    
    Dim A As Long
    
    Parameter(0) = Index
    Parameter(1) = Slot1
    Parameter(2) = Slot2
    A = RunScript("SWAPOBJECT")
    
    If A = 0 Then
        Dim TmpObject As InvObject, ObjLen As Long
        ObjLen = LenB(TmpObject)
        MemCopy TmpObject, Item1, ObjLen
        MemCopy Item1, Item2, ObjLen
        MemCopy Item2, TmpObject, ObjLen
        SendSocket Index, Chr2(17) + Chr2(Slot1) + DoubleChar$(Item1.Object) + QuadChar$(Item1.Value) + Chr2(Item1.prefix) + Chr2(Item1.prefixVal) + Chr2(Item1.suffix) + Chr2(Item1.SuffixVal) + Chr2(Item1.Affix) + Chr2(Item1.AffixVal) + Chr2(Item1.ObjectColor)
        SendSocket Index, Chr2(17) + Chr2(Slot2) + DoubleChar$(Item2.Object) + QuadChar$(Item2.Value) + Chr2(Item2.prefix) + Chr2(Item2.prefixVal) + Chr2(Item2.suffix) + Chr2(Item2.SuffixVal) + Chr2(Item2.Affix) + Chr2(Item2.AffixVal) + Chr2(Item2.ObjectColor)
    End If
End Sub

Sub CopyObject(ByRef source As InvObject, ByRef dest As InvObject)
    Dim ObjLen As Long
    ObjLen = LenB(source)
    MemCopy dest, source, ObjLen
End Sub
 
Sub ClearObject(ByRef Item As InvObject)
    Item.Object = 0
    Item.prefix = 0
    Item.prefixVal = 0
    Item.suffix = 0
    Item.SuffixVal = 0
    Item.Value = 0
End Sub

Function CompareObject(ByRef Item1 As InvObject, ByRef Item2 As InvObject) As Boolean
    Dim retVal As Boolean
    retVal = True
    If Item1.Object <> Item2.Object Then retVal = False
    If Item1.Value <> Item2.Value Then retVal = False
    If Item1.prefix <> Item2.prefix Then retVal = False
    If Item1.prefixVal <> Item2.prefixVal Then retVal = False
    If Item1.suffix <> Item2.suffix Then retVal = False
    If Item1.SuffixVal <> Item2.SuffixVal Then retVal = False
    
    CompareObject = retVal
End Function



Sub SendInventory(Index As Long)
    Dim A As Long, ST1 As String
    
    ST1 = ""
    With player(Index)
        For A = 1 To 20
            If .Inv(A).Object > 0 Then
                ST1 = ST1 + DoubleChar$(15) + Chr2(17) + Chr2(A) + DoubleChar(.Inv(A).Object) + QuadChar(.Inv(A).Value) + Chr2(.Inv(A).prefix) + Chr2(.Inv(A).prefixVal) + Chr2(.Inv(A).suffix) + Chr2(.Inv(A).SuffixVal) + Chr2(.Inv(A).Affix) + Chr2(.Inv(A).AffixVal) + Chr2(.Inv(A).ObjectColor)
            Else
                ST1 = ST1 + DoubleChar$(2) + Chr2(18) + Chr2(A)
            End If
        Next A
    End With
    
    SendRaw Index, ST1
End Sub
