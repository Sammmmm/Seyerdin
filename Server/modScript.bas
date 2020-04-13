Attribute VB_Name = "modScript"
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

#Const UseExperience = True
#Const CheckIPDupe = True
#Const AdminCheck = False
#Const GodChecking = False
#Const PublicServer = True
#Const DebugScripts = True

'Declare Function RunASMScript Lib "script.dll" Alias "RunScript" (Script As Byte, FunctionTable As Long, Parameters As Long) As Long
Declare Function RunASMScript Lib "script.dll" Alias "RunScript" (Script As Any, FunctionTable As Any, Parameters As Any) As Long
Declare Function SysFreeString Lib "oleaut32.dll" (ByVal StringPointer As Long) As Long
Declare Function SysAllocStringByteLen Lib "oleaut32.dll" (ByVal St As String, ByVal Length As Long) As Long

Public FunctionTable(0 To 300) As Long
Public ScriptsRunning As Long

Public Parameter(0 To 9) As Long
Public StringStack(0 To 1023) As Long
Public StringPointer As Long
Public LastScript As String
Public scriptTable As New clsHashTable

Function RunScript(Name As String, Optional scriptRunScript As Boolean = False) As Long
On Error GoTo ErrHandler
    If scriptTable.Exists(Name) Then
        Dim localParams(0 To 9) As Long
        Dim A As Long
        Dim MCODE() As Byte
        Dim t As Long
        Dim tempMap As EditMapData
                
        LogScriptStart Name
                
        For A = 0 To 9
            localParams(A) = Parameter(A)
        Next A
        
        MCODE = scriptTable.Item(Name)
        LastScript = Name
        
        t = GetTickCount

        tempMap = CurEditMap
        'DoEvents 'TODO2020: think about this, this may help with long running scripts to let networking go, but it also feels like it can make script running nondeterministic...
        'update i disabled it maybe
        ScriptsRunning = ScriptsRunning + 1
        RunScript = RunASMScript(MCODE(0), FunctionTable(0), localParams(0))
        ScriptsRunning = ScriptsRunning - 1

        If (CurEditMap.Num <> tempMap.Num) Then CurEditMap = tempMap
        
        t = GetTickCount - t
        If t > LongRunningThreshold Then
            PrintLog "Script " & Name & " takes " & t
        End If
        
        If ScriptsRunning = 0 Then
            For A = 1 To currentMaxUser
                If IsPlaying(A) Then
                    If player(A).St <> "" Then
                        FlushSocket A
                    End If
                    
                    If (player(A).ScriptUpdateMap) Then
                        Partmap A, False, False
                        JoinMap A, False
                        player(A).ScriptUpdateMap = False
                    End If
                End If
            Next A
        End If
        
        LogScriptEnd Name
    Else
        LogScriptNotExists Name
    End If
Exit Function
ErrHandler:
    Dim ST1 As String
    ST1 = "----SCRIPT ERROR: " & Err.Description & ", (" & Err.Number & ") " & "--------" & vbCrLf
    ST1 = ST1 & "Script: " & Name & vbCrLf
    For A = 0 To 9
        ST1 = ST1 & "P" & A & ": " & Parameter(A) & vbCrLf
    Next A
    ST1 = ST1 & "------------------------------------------------"
    
    LogScriptCrash Name
    LogCrash ST1
        
    'For A = 0 To currentMaxUser
    '    If IsPlaying(A) Then
    '        If player(A).Access = 11 Then
    '            Resume Next
    '        End If
    '    End If
    'Next A

    ' I think we dont need to crash if a god isn't online at this stage in our lives
    Resume Next
    
'Unhook
'End
End Function

Sub Boot_Player(ByVal Index As Long, ByVal reason As String)
    BootPlayer Index, 0, StrConv(reason, vbUnicode)
End Sub
Sub Ban_Player(ByVal Index As Long, ByVal NumDays As Long, ByVal reason As String)
    BanPlayer Index, 0, NumDays, StrConv(reason, vbUnicode), "Script Ban"
End Sub
Function Find_Player(ByVal Name As String) As Long
    Find_Player = FindPlayer(StrConv(Name, vbUnicode))
End Function

Function GetObjX(ByVal mapIndex As Long, ByVal ObjIndex As Long) As Long
    If mapIndex >= 1 And mapIndex <= 5000 And ObjIndex >= 0 And ObjIndex <= 49 Then
        With map(mapIndex).Object(ObjIndex)
            GetObjX = .x
        End With
    End If
End Function
Function GetObjY(ByVal mapIndex As Long, ByVal ObjIndex As Long) As Long
    If mapIndex >= 1 And mapIndex <= 5000 And ObjIndex >= 0 And ObjIndex <= 49 Then
        With map(mapIndex).Object(ObjIndex)
            GetObjY = .y
        End With
    End If
End Function
Function GetObjNum(ByVal mapIndex As Long, ByVal ObjIndex As Long) As Long
    If mapIndex >= 1 And mapIndex <= 5000 And ObjIndex >= 0 And ObjIndex <= 49 Then
        With map(mapIndex).Object(ObjIndex)
            GetObjNum = .Object
        End With
    End If
End Function
Function GetTileAtt(ByVal mapIndex As Long, ByVal x As Long, ByVal y As Long) As Long
    If mapIndex >= 1 And mapIndex <= 5000 And x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
        GetTileAtt = map(mapIndex).Tile(x, y).Att
    End If
End Function
Function DestroyObj(ByVal mapIndex As Long, ByVal ObjIndex As Long) As Long
    If mapIndex >= 1 And mapIndex <= 5000 And ObjIndex >= 0 And ObjIndex <= 49 Then
        With map(mapIndex).Object(ObjIndex)
            .Object = 0
            SendToMap mapIndex, Chr2(15) + Chr2(ObjIndex) 'Erase Map Obj
        End With
    End If
End Function
Function GetObjVal(ByVal mapIndex As Long, ByVal ObjIndex As Long) As Long
    If mapIndex >= 1 And mapIndex <= 5000 And ObjIndex >= 0 And ObjIndex <= 49 Then
        With map(mapIndex).Object(ObjIndex)
            GetObjVal = .Value
        End With
    End If
End Function
Function NewString(St As String) As Long
    Dim A As Long
    If StringPointer < 1024 Then
        If StringStack(StringPointer) > 0 Then
            SysFreeString StringStack(StringPointer)
        End If
    
        A = SysAllocStringByteLen(St, Len(St))
        StringStack(StringPointer) = A
        StringPointer = StringPointer + 1
        NewString = A
    Else
        StringPointer = 0
        
        If StringStack(StringPointer) > 0 Then
            SysFreeString StringStack(StringPointer)
        End If
        
        A = SysAllocStringByteLen(St, Len(St))
        StringStack(StringPointer) = A
        StringPointer = StringPointer + 1
        NewString = A
    End If
End Function
Function ScriptAttackMonster(ByVal Index As Long, ByVal monsterIndex As Long, ByVal damage As Long, ByVal Magic As Long) As Long
    If Index >= 1 And Index <= MaxUsers And monsterIndex >= 0 And monsterIndex <= 9 Then
        If player(Index).Mode = modePlaying Then
            AttackMonster Index, monsterIndex, damage, False, True
        End If
    End If
End Function
Function CanAttackMonster(ByVal Index As Long, ByVal monsterIndex As Long) As Long
    Dim mapIndex As Long
    If Index >= 1 And Index <= MaxUsers And monsterIndex >= 0 And monsterIndex <= 9 Then
        If player(Index).Mode = modePlaying Then
            mapIndex = player(Index).map
            If ExamineBit(map(mapIndex).Flags(0), 5) = False And map(mapIndex).monster(monsterIndex).monster > 0 Then
                CanAttackMonster = True
            End If
        End If
    End If
End Function
Function CanAttackPlayer(ByVal Player1 As Long, ByVal Player2 As Long, Optional ByVal CanAttackGuild = False) As Long
    If Player1 >= 1 And Player1 <= MaxUsers And Player2 >= 1 And Player2 <= MaxUsers Then
        With player(Player1)
            If ExamineBit(map(.map).Flags(0), 0) = False Then
                If .Mode = modePlaying And player(Player2).Mode = modePlaying Then
                    If .map = player(Player2).map Then
                        'If .Access = 0 And Player(Player2).Access = 0 Then
                            If map(player(Player2).map).Tile(player(Player2).x, player(Player2).y).Att <> 21 Then
                                If map(player(Player1).map).Tile(player(Player1).x, player(Player1).y).Att <> 2 Then
                                    'If .Party = 0 Or Not .Party <> Player(Player2).Party Then
                                        If (.Guild > 0) Or (ExamineBit(map(.map).Flags(0), 6) = True) Then
                                            If player(Player2).Guild > 0 Or (ExamineBit(map(.map).Flags(0), 6) = True) Then
                                                'If .Guild = 0 Or (.Guild <> player(Player2).Guild Or CanAttackGuild = True) Then
                                                    CanAttackPlayer = True
                                                'End If
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
End Function

Sub KillPlayer(ByVal Index As Long)
    If Index > 0 And Index < MaxUsers Then
        SendAll2 Chr2(61) + Chr2(Index)
        CreateFloatingEvent player(Index).map, player(Index).x, player(Index).y, FT_ENDED
        PlayerDied Index
    End If
End Sub

Function GetAbs(ByVal Value As Long) As Long
    GetAbs = Abs(Value)
End Function
Function GetMaxUsers() As Long
    GetMaxUsers = currentMaxUser
End Function
Function GetMonsterType(ByVal mapIndex As Long, ByVal monsterIndex As Long) As Long
    If mapIndex >= 1 And mapIndex <= 5000 And monsterIndex >= 0 And monsterIndex <= 9 Then
        GetMonsterType = map(mapIndex).monster(monsterIndex).monster
    End If
End Function
Function GetMonsterTarget(ByVal mapIndex As Long, ByVal monsterIndex As Long) As Long
    If mapIndex >= 1 And mapIndex <= 5000 And monsterIndex >= 0 And monsterIndex <= 9 Then
        GetMonsterTarget = map(mapIndex).monster(monsterIndex).Target
    End If
End Function

Sub NPCTell(ByVal Index As Long, ByVal NPCNum As Long, ByVal St As String)
    If Index >= 1 And Index <= MaxUsers Then
        With player(Index)
            If .Mode = modePlaying Then
                If NPCNum > 0 Then
                    SendSocket2 Index, Chr2(88) + Chr2(NPCNum) + StrConv(St, vbUnicode)
                End If
            End If
        End With
    End If
End Sub

Sub ResetPlayerFlag(ByVal FlagNum As Long)
    Dim A As Long
    If FlagNum >= 0 And FlagNum <= 255 Then
        A = World.PlayerFlagCounter(FlagNum)
        If A = 2147483647 Then
            A = 0
        Else
            A = A + 1
        End If
        World.PlayerFlagCounter(FlagNum) = A
    End If
End Sub
Sub ScriptTimer(ByVal Index As Long, ByVal Seconds As Long, ByVal Script As String)
Dim A As Long
    If Index >= 1 And Index <= MaxUsers Then
        If Seconds > 86400 Then Seconds = 86400
        If Seconds < 0 Then Seconds = 0
        With player(Index)
            If .Mode = modePlaying Then
                For A = 1 To MaxPlayerTimers
                    If .ScriptTimer(A) = 0 Then
                        .Script(A) = StrConv(Script, vbUnicode)
                        .ScriptTimer(A) = GetTickCount + Seconds * 1000 - 500
                        Exit For
                    End If
                Next A
            End If
        End With
    End If
End Sub
Sub SetFlag(ByVal FlagNum As Long, ByVal Value As Long)
    If FlagNum >= 0 And FlagNum <= 255 Then
        If Value >= 0 Then
            World.Flag(FlagNum) = Value
        End If
    End If
End Sub
Function GetFlag(ByVal FlagNum As Long) As Long
    If FlagNum >= 0 And FlagNum <= 255 Then
        GetFlag = World.Flag(FlagNum)
    End If
End Function
Function SetMonsterTarget(ByVal mapIndex As Long, ByVal monsterIndex As Long, ByVal player As Long) As Long
    If mapIndex >= 1 And mapIndex <= 5000 And monsterIndex >= 0 And monsterIndex <= 9 And player >= 1 And player <= MaxUsers Then
        With map(mapIndex).monster(monsterIndex)
            .TargType = TargTypePlayer
            .Target = player
            .Distance = 1
        End With
    Else
        If mapIndex >= 1 And mapIndex <= 5000 And monsterIndex >= 0 And monsterIndex <= 9 And player = 0 Then
            With map(mapIndex).monster(monsterIndex)
                .TargType = 0
                .Target = 0
            End With
        End If
    End If
End Function
Function SetMonsterTarget2(ByVal mapIndex As Long, ByVal monsterIndex As Long, ByVal MonsterNum As Long) As Long
    If mapIndex >= 1 And mapIndex <= 5000 And monsterIndex >= 0 And monsterIndex <= 9 And MonsterNum >= 0 And MonsterNum <= 9 Then
        With map(mapIndex).monster(monsterIndex)
            .TargType = TargTypeMonster
            .Target = MonsterNum
            .Distance = 1
        End With
    End If
End Function

Function GetMonsterX(ByVal mapIndex As Long, ByVal monsterIndex As Long) As Long
    If mapIndex >= 1 And mapIndex <= 5000 And monsterIndex >= 0 And monsterIndex <= 9 Then
        GetMonsterX = map(mapIndex).monster(monsterIndex).x
    End If
End Function
Function GetMonsterY(ByVal mapIndex As Long, ByVal monsterIndex As Long) As Long
    If mapIndex >= 1 And mapIndex <= 5000 And monsterIndex >= 0 And monsterIndex <= 9 Then
        GetMonsterY = map(mapIndex).monster(monsterIndex).y
    End If
End Function

Function GetSqr(ByVal Value As Long) As Long
    GetSqr = Sqr(Value)
End Function

Function GetTime() As Long
    GetTime = World.Hour
End Function

Function HasObj(ByVal Index As Long, ByVal ObjIndex As Long) As Long
    Dim A As Long, B As Long, C As Long
    If Index >= 1 And Index <= MaxUsers And ObjIndex >= 1 And ObjIndex <= MAXITEMS Then
        B = Object(ObjIndex).Type
        C = 0
        With player(Index)
            For A = 1 To 20
                With .Inv(A)
                    If .Object = ObjIndex Then
                        If B = 6 Or B = 11 Then
                            C = C + .Value
                        Else
                            C = C + 1
                        End If
                    End If
                End With
            Next A
            For A = 1 To 5
                With .Equipped(A)
                    If .Object = ObjIndex Then
                        If B = 6 Or B = 11 Then
                            C = C + .Value
                        Else
                            C = C + 1
                        End If
                    End If
                End With
            Next A
        End With
        HasObj = C
    End If
End Function
Function GetPlayerName(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        With player(Index)
            GetPlayerName = NewString(.Name)
        End With
    End If
End Function
Function GetPlayerIP(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        With player(Index)
            GetPlayerIP = NewString(.ip)
        End With
    End If
End Function

Function GetPlayerUser(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        With player(Index)
            GetPlayerUser = NewString(.user)
        End With
    End If
End Function
'Function GetPlayerDesc(ByVal Index As Long) As Long
'    If Index >= 1 And Index <= MaxUsers Then
'        With player(Index)
'            GetPlayerDesc = NewString(.desc)
'        End With
'    End If
'End Function
Function GetGuildName(ByVal Index As Long) As Long
    If Index >= 1 And Index <= 255 Then
        With Guild(Index)
            GetGuildName = NewString(.Name)
        End With
    End If
End Function

Function GiveObj(ByVal Index As Long, ByVal ObjIndex As Long, ByVal amount As Long) As Long
    Dim stackSize As Long, invSlot As Long, temp As Long, toGive As Long
    Dim remaining As Long

    If amount = 0 Then amount = 1
    remaining = amount
    If ObjIndex > 0 And ObjIndex <= 1000 Then
        With player(Index)
            If Object(ObjIndex).Type = 6 Or Object(ObjIndex).Type = 11 Then
                stackSize = Object(ObjIndex).Data(0)
                invSlot = FindInvObject(Index, ObjIndex, False)
                If invSlot = 0 Then
                    invSlot = FreeInvNum(Index)
                    If invSlot > 0 Then .Inv(invSlot).Value = 0
                Else
                    While (player(Index).Inv(invSlot).Value >= stackSize And stackSize <> 0 And invSlot <> 20)
                        temp = FindInvObject(Index, ObjIndex, False, invSlot + 1)
                        If temp = invSlot Then
                            invSlot = FreeInvNum(Index)
                            If invSlot > 0 Then .Inv(invSlot).Value = 0
                        Else
                            invSlot = temp
                        End If
                    Wend
                End If
            Else
                invSlot = FreeInvNum(Index)
                If invSlot > 0 Then .Inv(invSlot).Value = 0
            End If
        
getAnother:
            If invSlot > 0 Then
                With .Inv(invSlot)
                    .Object = ObjIndex 'c
                    If Object(ObjIndex).Type = 6 Or Object(ObjIndex).Type = 11 Then 'Money or ammo
                        If CDbl(.Value) + remaining > 2147483647# Then
                            toGive = 2147483647
                        Else
                            If stackSize > 0 Then
                                If .Value + remaining > stackSize Then
                                    temp = stackSize - .Value
                                    .Value = stackSize
                                    remaining = remaining - temp
                                    SendSocket Index, Chr2(17) + Chr2(invSlot) + DoubleChar(CInt(ObjIndex)) + QuadChar(stackSize) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)  'New Inv Obj
                                    temp = FindInvObject(Index, ObjIndex, False, invSlot + 1)
                                    If temp = invSlot Then
                                        invSlot = FreeInvNum(Index)
                                        If invSlot > 0 Then player(Index).Inv(invSlot).Value = 0
                                    Else
                                        invSlot = temp
                                    End If
                                    
                                    GoTo getAnother
                                Else
                                    toGive = .Value + remaining
                                    remaining = 0
                                End If
                            Else
                                toGive = .Value + remaining
                                remaining = 0
                            End If
                        End If
                    Else
                        toGive = 1
                        remaining = remaining - 1
                    End If
    
                    .Object = ObjIndex
                    .prefix = 0
                    .prefixVal = 0
                    .suffix = 0
                    .SuffixVal = 0
                    .ObjectColor = 0
                    .Affix = 0
                    .AffixVal = 0
                    For temp = 0 To 3
                      .Flags(temp) = 0
                    Next temp
                    Select Case Object(ObjIndex).Type
                        Case 1, 2, 3, 4, 10 'Weapon, Shield, Armor, Helmut
                            .Value = CLng(Object(ObjIndex).Data(0)) * 10
                        Case 6, 11 'Money
                            .Value = toGive
                        Case 8 'Ring
                            .Value = CLng(Object(ObjIndex).Data(1)) * 10
                        Case Else
                            .Value = 0
                    End Select
                    
                    SendSocket2 Index, Chr2(17) + Chr2(invSlot) + DoubleChar(CInt(ObjIndex)) + QuadChar(.Value) + String$(7, Chr2(0)) 'New Inv Obj
                End With
            Else
                'SendSocket Index, Chr2(16) + Chr2(1) 'Inv Full
            End If
            GiveObj = amount - remaining
        End With
    End If
End Function

Sub GlobalMessage(ByVal Message As String, ByVal MsgColor As Long)
    MsgColor = MsgColor Mod 16
    SendAll2 Chr2(56) + Chr2(MsgColor) + StrConv(Message, vbUnicode)
End Sub

Function IsPlaying(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        IsPlaying = player(Index).Mode = modePlaying
    End If
End Function

Sub MapMessage(ByVal Index As Long, ByVal Message As String, ByVal MsgColor As Long)
    If Index >= 1 And Index <= 5000 Then
        MsgColor = MsgColor Mod 16
        SendToMap2 Index, Chr2(56) + Chr2(MsgColor) + StrConv(Message, vbUnicode)
    End If
End Sub
Sub MapMessageAllBut(ByVal mapIndex As Long, ByVal playerIndex As Long, ByVal Message As String, ByVal MsgColor As Long)
    If mapIndex >= 1 And mapIndex <= 5000 And playerIndex >= 1 And playerIndex <= MaxUsers Then
        MsgColor = MsgColor Mod 16
        SendToMapAllBut2 mapIndex, playerIndex, Chr2(56) + Chr2(MsgColor) + StrConv(Message, vbUnicode)
    End If
End Sub
Function OpenDoor(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, ByVal Flags As Long) As Long
    Dim A As Long
    If mapNum >= 1 And mapNum <= 5000 And x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
        If Flags <= 0 Or Flags > 3 Then Flags = 3
    
        A = FreeMapDoorNum(mapNum)
        If A >= 0 Then
            With map(mapNum).Door(A)
                .Att = map(mapNum).Tile(x, y).Att
                .Wall = map(mapNum).Tile(x, y).WallTile
                .x = x
                .y = y
                .t = GetTickCount
            End With
            
            If ExamineBit(Flags, 1) Then map(mapNum).Tile(x, y).Att = 0
            If ExamineBit(Flags, 0) Then map(mapNum).Tile(x, y).WallTile = 0
            
            SendToMap2 mapNum, Chr2(36) + Chr2(A) + Chr2(x) + Chr2(y) + Chr2(Flags)
            OpenDoor = 1
        End If
    End If
End Function
Sub PlayerMessage(ByVal Index As Long, ByVal Message As String, ByVal MsgColor As Long)
    If Index >= 1 And Index <= MaxUsers Then
        MsgColor = MsgColor Mod 16
        SendSocket2 Index, Chr2(56) + Chr2(MsgColor) + StrConv(Message, vbUnicode)
    End If
End Sub
Function RunScript0(ByVal Script As String) As Long
    RunScript0 = RunScript(StrConv(Script, vbUnicode), True)
End Function
Function RunScript1(ByVal Script As String, ByVal Parm1 As Long) As Long
    Parameter(0) = Parm1
    RunScript1 = RunScript(StrConv(Script, vbUnicode), True)
End Function
Function RunScript2(ByVal Script As String, ByVal Parm1 As Long, ByVal Parm2 As Long) As Long
    Parameter(0) = Parm1
    Parameter(1) = Parm2
    RunScript2 = RunScript(StrConv(Script, vbUnicode), True)
End Function
Function RunScript3(ByVal Script As String, ByVal Parm1 As Long, ByVal Parm2 As Long, ByVal Parm3 As Long) As Long
    Parameter(0) = Parm1
    Parameter(1) = Parm2
    Parameter(2) = Parm3
    RunScript3 = RunScript(StrConv(Script, vbUnicode), True)
End Function
Function RunScript4(ByVal Script As String, ByVal Parm1 As Long, ByVal Parm2 As Long, ByVal Parm3 As Long, ByVal Parm4 As Long) As Long
    Parameter(0) = Parm1
    Parameter(1) = Parm2
    Parameter(2) = Parm3
    Parameter(3) = Parm4
    RunScript4 = RunScript(StrConv(Script, vbUnicode), True)
End Function

Function RunScript5(ByVal Script As String, ByVal Parm1 As Long, ByVal Parm2 As Long, ByVal Parm3 As Long, ByVal Parm4 As Long, ByVal Parm5 As Long) As Long
    Parameter(0) = Parm1
    Parameter(1) = Parm2
    Parameter(2) = Parm3
    Parameter(3) = Parm4
    Parameter(4) = Parm5
    RunScript5 = RunScript(StrConv(Script, vbUnicode), True)
End Function

Function RunScript6(ByVal Script As String, ByVal Parm1 As Long, ByVal Parm2 As Long, ByVal Parm3 As Long, ByVal Parm4 As Long, ByVal Parm5 As Long, ByVal Parm6 As Long) As Long
    Parameter(0) = Parm1
    Parameter(1) = Parm2
    Parameter(2) = Parm3
    Parameter(3) = Parm4
    Parameter(4) = Parm5
    Parameter(5) = Parm6
    RunScript6 = RunScript(StrConv(Script, vbUnicode), True)
End Function

Function RunScript10(ByVal Script As String, ByVal Parm1 As Long, ByVal Parm2 As Long, ByVal Parm3 As Long, ByVal Parm4 As Long, ByVal Parm5 As Long, ByVal Parm6 As Long, ByVal Parm7 As Long, ByVal Parm8 As Long, ByVal Parm9 As Long, ByVal Parm10 As Long) As Long
    Parameter(0) = Parm1
    Parameter(1) = Parm2
    Parameter(2) = Parm3
    Parameter(3) = Parm4
    Parameter(4) = Parm5
    Parameter(5) = Parm6
    Parameter(6) = Parm7
    Parameter(7) = Parm8
    Parameter(8) = Parm9
    Parameter(9) = Parm10
    RunScript10 = RunScript(StrConv(Script, vbUnicode), True)
End Function



Sub SetPlayerSprite(ByVal Index As Long, ByVal sprite As Long)
    If Index >= 1 And Index <= MaxUsers And sprite >= 0 And sprite <= 255 Then
        With player(Index)
            If sprite = 0 Then
                If .Guild > 0 Then
                    If Guild(.Guild).sprite > 0 Then
                        .sprite = Guild(.Guild).sprite
                    Else
                        .sprite = .Class * 2 + .Gender - 1
                    End If
                Else
                    .sprite = .Class * 2 + .Gender - 1
                End If
            Else
                .sprite = sprite
            End If
            SendAll2 Chr2(63) + Chr2(Index) + Chr2(.sprite)
        End With
    End If
End Sub

Sub SetPLayerClass(ByVal Index As Long, ByVal Class As Long)


        If Index >= 1 And Index <= MaxUsers And Class <= 10 And Class >= 1 Then
            With player(Index)
                        .Class = Class
                        ResetStats (Index)
                        ResetSkills (Index)
                        SavePlayerData (Index)
                        BootPlayer Index, Index, "Class Change"
            End With
        End If

End Sub

Function SpawnMonster(ByVal mapIndex As Long, ByVal monster As Long, ByVal x As Long, ByVal y As Long, ByVal Frozen As Long) As Long
    Dim A As Long, B As Long
    
    If mapIndex >= 1 And mapIndex <= 5000 And monster >= 1 And monster <= MAXITEMS And x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
        With map(mapIndex)
            For A = 0 To 9
                With .monster(A)
                    If .monster = 0 Then
                        SendToMapRaw2 mapIndex, SpawnMapMonster(mapIndex, A, monster, x, y)
                        If Frozen > 0 Then
                            With map(mapIndex).monster(A)
                                For B = 0 To Frozen
                                    If B < 15 Then
                                        .MonsterQueue(.CurrentQueue).Action = QUEUE_PAUSE
                                        .CurrentQueue = .CurrentQueue + 1
                                    Else
                                        Exit For
                                    End If
                                Next B
                            End With
                        End If
                        SpawnMonster = A + 1
                        Exit Function
                    End If
                End With
            Next A
        End With
    End If
    
    SpawnMonster = 0
End Function

Function SpawnMonsterOnMap(ByVal mapIndex As Long, ByVal monster As Long) As Long
    Dim A As Long
    If mapIndex >= 1 And mapIndex <= 5000 And monster >= 1 And monster <= MAXITEMS Then
        With map(mapIndex)
            For A = 0 To 9
                With .monster(A)
                    If .monster = 0 Then
                        SpawnMonsterOnMap = A
                        SendToMapRaw2 mapIndex, NewMapMonster2(mapIndex, monster, A)
                        Exit For
                    End If
                End With
            Next A
        End With
    End If
End Function

Sub WarpMonster(ByVal mapNum As Long, ByVal MonsterNum As Long, ByVal x As Long, ByVal y As Long)
    If mapNum >= 1 And mapNum <= 5000 And MonsterNum < 10 And x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
        With map(mapNum).monster(MonsterNum)
            If .monster > 0 Then
                .x = x
                .y = y
                SendToMap2 mapNum, Chr2(38) + Chr2(MonsterNum) + DoubleChar(.monster) + Chr2(.x) + Chr2(.y) + Chr2(.D) + DoubleChar(.HP)
            End If
        End With
    End If
End Sub
Function SpawnObject(ByVal mapIndex As Long, ByVal Object As Long, ByVal Value As Long, ByVal x As Long, ByVal y As Long, ByVal magical As Long) As Long
    If mapIndex >= 1 And mapIndex <= 5000 And Object >= 1 And Object <= MAXITEMS And Value >= 0 And x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
        If magical < 0 Then magical = 0
        If magical > 100 Then magical = 100
        SpawnObject = NewMapObject(mapIndex, Object, Value, x, y, False, magical)
    End If
End Function
Function SpawnObject2(ByVal mapIndex As Long, ByVal Object As Long, ByVal Value As Long, ByVal x As Long, ByVal y As Long, ByVal magical As Long) As Long
    If mapIndex >= 1 And mapIndex <= 5000 And Object >= 1 And Object <= MAXITEMS And Value >= 0 And x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
        If magical < 0 Then magical = 0
        If magical > 100 Then magical = 100
        SpawnObject2 = NewMapObject(mapIndex, Object, Value, x, y, True, magical)
    End If
End Function

Function SpawnObject3(ByVal mapIndex As Long, ByVal Object As Long, ByVal Value As Long, ByVal x As Long, ByVal y As Long, ByVal magical As Long, ByVal unburnable As Long, ByVal durabilityPercent As Long, ByVal flag0 As Long, ByVal flag1 As Long, ByVal flag2 As Long, ByVal flag3 As Long, ByVal ObjectColor As Long, ByVal prefix As Long, ByVal prefixVal As Long, ByVal suffix As Long, ByVal SuffixVal As Long, ByVal Affix As Long, ByVal AffixVal As Long) As Long
    If mapIndex >= 1 And mapIndex <= 5000 And Object >= 1 And Object <= MAXITEMS And Value >= 0 And x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
        If magical < 0 Then magical = 0
        If magical > 100 Then magical = 100
        SpawnObject3 = NewMapObject(mapIndex, Object, Value, x, y, (unburnable > 0), magical, durabilityPercent, flag0, flag1, flag2, flag3, ObjectColor, prefix, prefixVal, suffix, SuffixVal, Affix, AffixVal)
    End If
End Function

Function Str(ByVal Value As Long) As Long
    Str = NewString(CStr(Value))
End Function
Function StrCat(ByVal String1 As String, ByVal String2 As String) As Long
    StrCat = NewString(StrConv(String1, vbUnicode) + StrConv(String2, vbUnicode))
End Function

Function StrCmp(ByVal String1 As String, ByVal String2 As String) As Long
    StrCmp = UCase$(StrConv(String1, vbUnicode)) = UCase$(StrConv(String2, vbUnicode))
End Function
Function GetInStr(ByVal String1 As String, ByVal String2 As String) As Long
    GetInStr = InStr(UCase$(StrConv(String1, vbUnicode)), UCase$(StrConv(String2, vbUnicode)))
End Function

Function StrFormat(ByVal String1 As String, ByVal String2 As String) As Long
    Dim St As String, ST1 As String, st2 As String
    ST1 = StrConv(String1, vbUnicode)
    st2 = StrConv(String2, vbUnicode)
    
    Dim A As Long, B As Byte
    For A = 1 To Len(ST1)
        B = Asc(Mid$(ST1, A, 1))
        If B = 42 Then
            St = St + st2
        Else
            St = St + Chr2(B)
        End If
    Next A
    
    StrFormat = NewString(St)
End Function
Function GetGuildHall(ByVal Index As Long) As Long
    If Index >= 1 And Index <= 255 Then
        GetGuildHall = Guild(Index).Hall
    End If
End Function
Function GetGuildBank(ByVal Index As Long) As Long
    If Index >= 1 And Index <= 255 Then
        GetGuildBank = Guild(Index).Bank
    End If
End Function
Function GetGuildSprite(ByVal Index As Long) As Long
    If Index >= 1 And Index <= 255 Then
        GetGuildSprite = Guild(Index).sprite
    End If
End Function
Function GetGuildMemberCount(ByVal Index As Long) As Long
    If Index >= 1 And Index <= 255 Then
        GetGuildMemberCount = CountGuildMembers(Index)
    End If
End Function
Function GetMapPlayerCount(ByVal Index As Long) As Long
    If Index >= 1 And Index <= 5000 Then
        GetMapPlayerCount = map(Index).NumPlayers
    End If
End Function

Function GetPlayerAccess(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerAccess = player(Index).Access
    End If
End Function
Function GetPlayerAgility(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerAgility = player(Index).Agility
    End If
End Function

Function GetPlayerConstitution(ByVal Index As Long) As Long
    If Index > 0 And Index <= MaxUsers Then
        GetPlayerConstitution = player(Index).Constitution
    End If
End Function

Function GetPlayerWisdom(ByVal Index As Long) As Long
    If Index > 0 And Index <= MaxUsers Then
        GetPlayerWisdom = player(Index).Wisdom
    End If
End Function

Function GetPlayerBank(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerBank = player(Index).Bank
    End If
End Function
Function GetPlayerClass(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerClass = player(Index).Class
    End If
End Function
Function GetPlayerEndurance(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerEndurance = player(Index).Endurance
    End If
End Function
Function GetPlayerEnergy(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerEnergy = player(Index).Energy
    End If
End Function
Function GetPlayerEquipped(ByVal Index As Long, ByVal EquippedIndex As Long) As Long
    If Index >= 1 And Index <= MaxUsers And EquippedIndex >= 1 And EquippedIndex <= 6 Then
        GetPlayerEquipped = player(Index).Equipped(EquippedIndex).Object
    End If
End Function

Function GetPlayerExperience(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerExperience = player(Index).Experience
    End If
End Function
Function GetPlayerGender(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerGender = player(Index).Gender
    End If
End Function
Function GetPlayerGuild(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerGuild = player(Index).Guild
    End If
End Function
Function GetPlayerHP(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerHP = player(Index).HP
    End If
End Function
Function GetPlayerMagicFind(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerMagicFind = player(Index).MagicBonus
    End If
End Function
Function GetPlayerIntelligence(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerIntelligence = player(Index).Intelligence
    End If
End Function
Function GetStatBonus(ByVal statValue As Long, ByVal piety As Long) As Long
    If statValue >= 0 And statValue <= 255 Then
        If piety >= 0 And piety <= 255 Then
            GetStatBonus = GetGenericStatBonus(statValue, piety)
        End If
    End If
End Function

Function GetPlayerInvObject(ByVal Index As Long, ByVal InvIndex As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        If InvIndex >= 1 And InvIndex <= 20 Then
            GetPlayerInvObject = player(Index).Inv(InvIndex).Object
        ElseIf InvIndex >= 21 And InvIndex <= 25 Then
            GetPlayerInvObject = player(Index).Equipped(InvIndex - 20).Object
        End If
    End If
End Function
Function GetPlayerInvValue(ByVal Index As Long, ByVal InvIndex As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        If InvIndex >= 1 And InvIndex <= 20 Then
            GetPlayerInvValue = player(Index).Inv(InvIndex).Value
        ElseIf InvIndex >= 21 And InvIndex <= 25 Then
            GetPlayerInvValue = player(Index).Equipped(InvIndex - 20).Value
        End If
    End If
End Function

Function GetPlayerLevel(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerLevel = player(Index).Level
    End If
End Function
Function GetPlayerMana(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerMana = player(Index).Mana
    End If
End Function
Function GetPlayerMap(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        With player(Index)
            If .Mode = modePlaying Then
                GetPlayerMap = .map
            End If
        End With
    End If
End Function

Function GetPlayerMaxEnergy(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerMaxEnergy = player(Index).MaxEnergy
    End If
End Function
Function GetPlayerMaxHP(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerMaxHP = player(Index).MaxHP
    End If
End Function
Function GetPlayerMaxMana(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerMaxMana = player(Index).MaxMana
    End If
End Function
Function GetPlayerSprite(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerSprite = player(Index).sprite
    End If
End Function
Function GetPlayerFlag(ByVal Index As Long, ByVal FlagNum As Long) As Long
    If Index >= 1 And Index <= MaxUsers And FlagNum >= 0 And FlagNum <= 255 Then
        With player(Index).Flag(FlagNum)
            If .ResetCounter = World.PlayerFlagCounter(FlagNum) Then
                GetPlayerFlag = .Value
            Else
                .Value = 0
                .ResetCounter = World.PlayerFlagCounter(FlagNum)
            End If
        End With
    End If
End Function
Sub SetPlayerFlag(ByVal Index As Long, ByVal FlagNum As Long, ByVal Value As Long)
    If Index >= 1 And Index <= MaxUsers And FlagNum >= 0 And FlagNum <= 255 Then
        If Value >= 0 Then
        With player(Index).Flag(FlagNum)
            .Value = Value
            .ResetCounter = World.PlayerFlagCounter(FlagNum)
        End With
        End If
    End If
End Sub

Function GetPlayerStatus(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerStatus = player(Index).Status
    End If
End Function
Function GetPlayerStrength(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerStrength = player(Index).strength
    End If
End Function
Function GetPlayerX(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerX = player(Index).x
    End If
End Function
Function GetPlayerY(ByVal Index As Long) As Long
    If Index >= 1 And Index <= MaxUsers Then
        GetPlayerY = player(Index).y
    End If
End Function
Function GetValue(Value As Long) As Long
    GetValue = Value
End Function

Function Random(ByVal Max As Long) As Long
    Random = Int(Rnd * Max)
End Function

Sub SetPlayerEnergy(ByVal Index As Long, ByVal Energy As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With player(Index)
            If .Mode = modePlaying Then
                If Energy > 999 Then Energy = 999
                If Energy < 0 Then Energy = 0
                .Energy = Energy
                SendSocket2 Index, Chr2(47) + DoubleChar(CInt(Energy))
            End If
        End With
    End If
End Sub


Sub SetPlayerMana(ByVal Index As Long, ByVal Mana As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With player(Index)
            If .Mode = modePlaying Then
                If Mana > 9999 Then Mana = 9999
                If Mana < 0 Then Mana = 0
                .Mana = Mana
                SendSocket2 Index, Chr2(48) + DoubleChar(CInt(Mana))
            End If
        End With
    End If
End Sub


Sub ScriptSetPlayerHP(ByVal Index As Long, ByVal HP As Long)
    If Index >= 1 And Index <= MaxUsers Then
        SetPlayerHP Index, HP
    End If
End Sub
Sub SetPlayerBank(ByVal Index As Long, ByVal Bank As Long)
    If Index >= 1 And Index <= MaxUsers Then
        With player(Index)
            If .Mode = modePlaying Then
                .Bank = Bank
            End If
        End With
    End If
End Sub

Sub SetPlayerStatus(ByVal Index As Long, ByVal Status As Long)
    If Index >= 1 And Index <= MaxUsers And Status >= 0 And Status <= 100 Then
        With player(Index)
            If .Mode = modePlaying Then
                .Status = Status
                SendAll2 Chr2(91) + Chr2(Index) + Chr2(Status)
            End If
        End With
    End If
End Sub
Sub SetPlayerGuild(ByVal Index As Long, ByVal GuildIndex As Long)
    If Index >= 1 And Index <= MaxUsers And GuildIndex >= 0 And GuildIndex <= 255 Then
        With player(Index)
            If GuildIndex > 0 Then
                If .Guild <> GuildIndex Then
                    .Guild = GuildIndex
                    .GuildRank = 1
                    If Guild(GuildIndex).sprite > 0 Then
                        .sprite = Guild(GuildIndex).sprite
                        SendAll2 Chr2(63) + Chr2(Index) + Chr2(.sprite)
                    End If
                    SendSocket2 Index, Chr2(72) + Chr2(GuildIndex) 'Change guild
                    SendAllBut2 Index, Chr2(73) + Chr2(Index) + Chr2(GuildIndex) 'Player changed guild
                End If
            Else
                If .Guild > 0 Then
                    If Guild(.Guild).sprite > 0 Then
                        .sprite = .Class * 2 + .Gender - 1
                        SendAll2 Chr2(63) + Chr2(Index) + Chr2(.sprite)
                    End If
                    .Guild = 0
                    SendSocket2 Index, Chr2(72) + Chr2(0) 'Change guild
                    SendAllBut2 Index, Chr2(73) + Chr2(Index) + Chr2(0) 'Player changed guild
                End If
            End If
        End With
    End If
End Sub

Sub SetGuildBank(ByVal Index As Long, ByVal Bank As Long)
    If Index >= 1 And Index <= 255 Then
        With Guild(Index)
            If .Name <> "" Then
                .Bank = Bank
                GuildRS.Bookmark = .Bookmark
                GuildRS.Edit
                GuildRS!Bank = Bank
                GuildRS.Update
            End If
        End With
    End If
End Sub
Sub PlayerWarp(ByVal Index As Long, ByVal map As Long, ByVal x As Long, ByVal y As Long)
    If Index >= 1 And Index <= MaxUsers And map >= 1 And map <= 5000 And x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
        With player(Index)
            If .Mode = modePlaying Then
                WarpPlayer Index, map, x, y
            End If
        End With
    End If
End Sub
Sub DeleteString(ByVal StPointer As Long)
    PrintLog StPointer & "    sp: " & StringPointer
End Sub
Function StrVal(ByVal String1 As String) As Long
    On Error Resume Next
    StrVal = Int(Val(StrConv(String1, vbUnicode)))
    On Error GoTo 0
End Function

Function TakeObj(ByVal Index As Long, ByVal ObjIndex As Long, ByVal amount As Long) As Long
    Dim A As Long, stackSize As Long, remaining As Long
    If Index >= 1 And Index <= MaxUsers And ObjIndex >= 1 And ObjIndex <= MAXITEMS Then
        With player(Index)
            If .Mode = modePlaying Then
                remaining = amount
                A = FindInvObject(Index, ObjIndex, False)
                Do
                    If A > 0 Then
                        With .Inv(A)
                            If Object(ObjIndex).Type = 6 Or Object(ObjIndex).Type = 11 Then
                                If .Value > remaining Then
                                    .Value = .Value - remaining
                                    remaining = 0
                                    SendSocket2 Index, Chr2(17) + Chr2(A) + DoubleChar(CLng(ObjIndex)) + QuadChar(.Value) + String$(7, Chr2(0)) 'New Inv Obj
                                Else
                                    remaining = remaining - .Value
                                    .Value = 0
                                    .Object = 0
                                    SendSocket2 Index, Chr2(18) + Chr2(A)
                                End If
                            Else
                                .Object = 0
                                .prefix = 0
                                .suffix = 0
                                TakeObj = 1
                                SendSocket2 Index, Chr2(18) + Chr2(A)
                                remaining = remaining - 1
                            End If
                        End With
                    Else
                        Exit Do
                    End If
                    A = FindInvObject(Index, ObjIndex, False, A + 1)
                Loop While remaining > 0 And .Inv(A).Object = ObjIndex
                TakeObj = amount - remaining
            End If
        End With
    End If
End Function

Sub GivePlayerExp(ByVal Index As Long, ByVal Experience As Long)
If Index >= 1 And Index <= MaxUsers And Experience <= 50000 Then
    GainExp Index, Experience, False, False
    'SendSocket Index, chr2(60) + QuadChar(Player(Index).Experience)
End If
End Sub
Function GetObjectName(ByVal ObjectNum As Long) As Long
If ObjectNum >= 1 And ObjectNum <= MAXITEMS Then
    GetObjectName = NewString(Object(ObjectNum).Name)
End If
End Function

Function GetObjectData(ByVal ObjectNum As Long, ByVal DataNum As Long) As Long
If ObjectNum >= 1 And ObjectNum <= MAXITEMS Then
    Select Case DataNum
        Case Is <= 4
            GetObjectData = Object(ObjectNum).Data(DataNum)
        Case 5
            GetObjectData = Object(ObjectNum).Class
        Case 6
            If ExamineBit(Object(ObjectNum).Flags, 6) Then
                GetObjectData = Object(ObjectNum).Picture + 255
            Else
                GetObjectData = Object(ObjectNum).Picture
            End If
        Case 7
            GetObjectData = Object(ObjectNum).Flags
        Case 8
            GetObjectData = Object(ObjectNum).MinLevel
        Case 9
            GetObjectData = Object(ObjectNum).Data(9)
    End Select
End If
End Function

Function GetObjectType(ByVal ObjectNum As Long) As Long
If ObjectNum >= 1 And ObjectNum <= MAXITEMS Then
    GetObjectType = Object(ObjectNum).Type
End If
End Function

Sub DisplayObjDur(ByVal Index As Long, ByVal ObjectNum As Long)
Dim Percent As Single, St As String, MsgColor As Long
Dim Display As Boolean
    Select Case Object(player(Index).Inv(ObjectNum).Object).Type
    Case 1, 2, 3, 4
        Display = True
    Case Else
        Display = False
    End Select
If Display = True Then
    'Percent = Int((Player(Index).Inv(ObjectNum).Value * 5) / Object(Player(Index).Inv(ObjectNum).Object).Data(0))
    Percent = player(Index).Inv(ObjectNum).Value / (Object(player(Index).Inv(ObjectNum).Object).Data(0) * 10)
    Percent = Int(Percent * 100)
    If Percent > 100 Then Percent = 100
    If Percent <= 5 Then
        St = "Your " + Object(player(Index).Inv(ObjectNum).Object).Name + " is about to break!"
        MsgColor = 2
    Else
        St = "Your " + Object(player(Index).Inv(ObjectNum).Object).Name + " is at " + CStr(Percent) + "% durability."
        MsgColor = 14
    End If
    SendSocket2 Index, Chr2(56) + Chr2(MsgColor Mod 16) + St
Else
    St = "This is an invalid object or no object."
    MsgColor = 2
    SendSocket2 Index, Chr2(56) + Chr2(MsgColor Mod 16) + St
End If
End Sub

Function GetObjectDur(ByVal Index As Long, ByVal Slot As Long) As Long
Dim Percent As Single
Dim Display As Boolean

    If Index > 0 And Index <= MaxUsers Then
        If Slot > 0 And Slot <= 20 Then
            Select Case Object(player(Index).Inv(Slot).Object).Type
                Case 1, 2, 3, 4, 8
                    Display = True
                Case Else
                    Display = False
            End Select
            If Display = True Then
                'Percent = Int((Player(Index).Inv(ObjectNum).Value * 5) / Object(Player(Index).Inv(ObjectNum).Object).Data(0))
                Select Case Object(player(Index).Inv(Slot).Object).Type
                    Case 1, 2, 3, 4 'Weapon, armor, etc.
                        Percent = player(Index).Inv(Slot).Value / (Object(player(Index).Inv(Slot).Object).Data(0) * 10)
                    Case 8 'Ring
                        Percent = player(Index).Inv(Slot).Value / (Object(player(Index).Inv(Slot).Object).Data(1) * 10)
                End Select
                Percent = Int(Percent * 100)
                'If Percent > 100 Then Percent = 100
                GetObjectDur = Percent
            Else
                GetObjectDur = 0
            End If
        ElseIf Slot >= 21 And Slot <= 25 Then
            Slot = Slot - 20
            Select Case Object(player(Index).Equipped(Slot).Object).Type
                Case 1, 2, 3, 4, 8
                    Display = True
                Case Else
                    Display = False
            End Select
            If Display = True Then
                'Percent = Int((Player(Index).Inv(ObjectNum).Value * 5) / Object(Player(Index).Inv(ObjectNum).Object).Data(0))
                Select Case Object(player(Index).Equipped(Slot).Object).Type
                    Case 1, 2, 3, 4 'Weapon, armor, etc.
                        Percent = player(Index).Equipped(Slot).Value / (Object(player(Index).Equipped(Slot).Object).Data(0) * 10)
                    Case 8 'Ring
                        Percent = player(Index).Equipped(Slot).Value / (Object(player(Index).Equipped(Slot).Object).Data(1) * 10)
                End Select
                Percent = Int(Percent * 100)
                If Percent > 100 Then Percent = 100
                GetObjectDur = Percent
            Else
                GetObjectDur = 0
            End If
    
        End If
    End If
End Function


Function GetMapObjectDur(ByVal mapNum As Long, ByVal ObjectNum As Long) As Long
Dim Percent As Single
Dim Display As Boolean
    Select Case Object(map(mapNum).Object(ObjectNum).Object).Type
    Case 1, 2, 3, 4, 8
        Display = True
    Case Else
        Display = False
    End Select
If Display = True Then
    'Percent = Int((Player(Index).Inv(ObjectNum).Value * 5) / Object(Player(Index).Inv(ObjectNum).Object).Data(0))
    Select Case Object(map(mapNum).Object(ObjectNum).Object).Type
        Case 1, 2, 3, 4 'Weapon, armor, etc.
            Percent = map(mapNum).Object(ObjectNum).Value / (Object(map(mapNum).Object(ObjectNum).Object).Data(0) * 10)
        Case 8 'Ring
            Percent = map(mapNum).Object(ObjectNum).Value / (Object(map(mapNum).Object(ObjectNum).Object).Data(1) * 10)
    End Select
    Percent = Int(Percent * 100)
    If Percent > 100 Then Percent = 100
    GetMapObjectDur = Percent
Else
    GetMapObjectDur = 0
End If
End Function

Sub SetInvObjectVal(ByVal Index As Long, ByVal invSlot As Long, ByVal NewVal As Long)
    If Index >= 1 And Index <= MaxUsers Then
        If invSlot >= 1 And invSlot <= 20 Then
            player(Index).Inv(invSlot).Value = NewVal
        ElseIf invSlot >= 21 And invSlot <= 25 Then
            player(Index).Equipped(invSlot - 20).Value = NewVal
        End If
        
        If NewVal > 0 Then
            SendSocket2 Index, Chr2(119) + Chr2(invSlot) + QuadChar(NewVal)
        Else
            If invSlot > 20 Then
                SendSocket Index, Chr2(57) + Chr2(invSlot - 20)
                player(Index).Equipped(invSlot - 20).Object = 0
                CalculateStats Index
            Else
                player(Index).Inv(invSlot).Object = 0
                SendSocket2 Index, Chr2(18) + Chr2(invSlot)
            End If
        End If
    End If
End Sub

Sub PlayCustomWav(ByVal Index As Long, ByVal SoundNum As Long)
    If Index >= 1 And Index <= MaxUsers And SoundNum <= 255 And SoundNum >= 1 Then
        SendSocket2 Index, Chr2(96) + Chr2(0) + Chr2(SoundNum)
    End If
End Sub

Sub PlayMapWav(ByVal mapNum As Long, ByVal SoundNum As Long)
    If mapNum >= 1 And mapNum <= 5000 Then
        If SoundNum >= 1 And SoundNum <= 255 Then
            SendToMap2 mapNum, Chr2(96) + Chr2(0) + Chr2(SoundNum)
        End If
    End If
End Sub

Sub PlayMusic(ByVal Index As Long, ByVal MusicNum As Long)
    If Index >= 1 And Index <= MaxUsers And MusicNum <= 255 And MusicNum >= 0 Then
        SendSocket2 Index, Chr2(96) + Chr2(1) + Chr2(MusicNum)
    End If
End Sub

Function GetPlayerArmor(ByVal Index As Long, ByVal damage As Long) As Long
If Index >= 1 And Index <= MaxUsers And damage <= 9999 And damage >= 1 Then
    GetPlayerArmor = PlayerArmor(Index, damage)
End If
End Function

Sub CreateMonsterEffect(ByVal mapNum As Long, ByVal monster As Long, ByVal sprite As Long, ByVal speed As Long, ByVal TotalFrames As Long, ByVal LoopCount As Long, ByVal EndSound As Long)
    If mapNum >= 1 And mapNum <= 5000 Then
            SendToMap2 mapNum, Chr2(99) + Chr2(TT_MONSTER) + Chr2(monster) + Chr2(sprite) + DoubleChar(CInt(speed)) + Chr2(TotalFrames) + Chr2(LoopCount) + Chr2(EndSound)
    End If
End Sub
Sub CreatePlayerEffect(ByVal Index As Long, ByVal sprite As Long, ByVal speed As Long, ByVal TotalFrames As Long, ByVal LoopCount As Long, ByVal EndSound As Long)
    If Index > 0 And Index <= MaxUsers Then
        SendToMap2 player(Index).map, Chr2(99) + Chr2(TT_PLAYER) + Chr2(Index) + Chr2(sprite) + DoubleChar(CInt(speed)) + Chr2(TotalFrames) + Chr2(LoopCount) + Chr2(EndSound)
    End If
End Sub
Sub CreateTileEffect(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, ByVal sprite As Long, ByVal speed As Long, ByVal TotalFrames As Long, ByVal LoopCount As Long, ByVal EndSound As Long)
    If mapNum >= 1 And mapNum <= 5000 Then
        If x < 0 Then x = 0
        If x > 11 Then x = 11
        If y < 0 Then y = 0
        If y > 11 Then y = 11
        SendToMap2 mapNum, Chr2(99) + Chr2(TT_TILE) + Chr2(x) + Chr2(y) + Chr2(sprite) + DoubleChar(CInt(speed)) + Chr2(TotalFrames) + Chr2(LoopCount) + Chr2(EndSound)
    End If
End Sub
Sub ScriptCreateFloatingTextAllBut(ByVal Index As Long, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, ByVal Color As Byte, ByVal Text As String)
    If mapNum >= 1 And mapNum <= 5000 Then
        If Len(Text) >= 1 Then
            SendToMapAllBut2 mapNum, Index, Chr2(100) + Chr2(x * 16 + y) + Chr2(Color) + StrConv(Text, vbUnicode)
        End If
    End If
End Sub
Sub CreateFloatingEvent(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, ByVal EventType As Long)
    If mapNum >= 1 And mapNum <= 5000 Then
        If EventType >= 0 Then
            SendToMap2 mapNum, Chr2(100) + Chr2(x * 16 + y) + Chr2(EventType * 16)
        End If
    End If
End Sub
Sub CreateStaticText(ByVal Index As Long, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, ByVal Color As Long, ByVal Time As Long, ByVal ID As Long, ByVal Text As String)
    If mapNum > 0 And mapNum <= 5000 Then
        SendSocket2 Index, Chr2(103) + DoubleChar(CInt(x)) + DoubleChar(CInt(y)) + Chr2(Color) + Chr2(Time) + Chr2(ID) + Chr2(10) + StrConv(Text, vbUnicode)
    End If
End Sub

Sub CreateSizedStaticText(ByVal Index As Long, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, ByVal Color As Long, ByVal Time As Long, ByVal ID As Long, ByVal Text As String, ByVal Size As Long)
    If mapNum > 0 And mapNum <= 5000 Then
        SendSocket2 Index, Chr2(103) + DoubleChar(CInt(x)) + DoubleChar(CInt(y)) + Chr2(Color) + Chr2(Time) + Chr2(ID) + Chr2(Size) + StrConv(Text, vbUnicode)
    End If
End Sub

Function sExamineBit(ByVal bytByte As Long, ByVal Bit As Long) As Long
    sExamineBit = ((bytByte And (2 ^ Bit)) > 0)
End Function

Sub sSetBit(bytByte As Long, Bit As Long)
    bytByte = bytByte Or (2 ^ Bit)
End Sub

Sub sClearBit(bytByte As Long, Bit As Long)
      bytByte = bytByte And Not (2 ^ Bit)
End Sub
 
Function GetClassLevel(ByVal Index As Long, ByVal Class As Long) As Long
    If Index > 0 And Index <= MaxUsers Then
        If Class > 0 And Class < 5 Then
            GetClassLevel = 0 'Player(Index).Class(Class)
        End If
    End If
End Function

Function LearnSkill(ByVal Index As Long, ByVal SKILL As Long) As Long
    If Index > 0 And Index <= MaxUsers Then
        If SKILL > SKILL_INVALID And SKILL < MAX_SKILLS Then
            If player(Index).SkillLevel(SKILL) = 0 Then
                player(Index).SkillLevel(SKILL) = 1
                SendSocket2 Index, Chr2(123) + Chr2(1) + Chr2(SKILL) + Chr2(player(Index).SkillLevel(SKILL))
            End If
        End If
    End If
End Function


Sub SetStringFlag(ByVal Flag As Long, ByVal String1 As String)
    If Flag >= 0 And Flag <= 255 Then
        If Len(String1) <= 2048 Then
            World.StringFlag(Flag) = StrConv(String1, vbUnicode)
        End If
    End If
End Sub

Function SetPlayerName(ByVal Index As Long, ByVal Data As String) As Long
Dim B As Long, C As Long, D As Long
Dim Name As String
Name = StrConv(Data, vbUnicode)
D = 0
For B = 1 To Len(Name)
    C = Asc(Mid$(Name, B, 1))
    If C = 8 Or (C >= 48 And C <= 57) Or (C >= 65 And C <= 90) Or (C >= 97 And C <= 122) Or C = 95 Then
    Else
        SetPlayerName = 0
        D = 1
        Exit For
    End If
Next B
If D = 0 Then
    If InStr(1, Name, " ") = 0 Then
        If InStr(1, Name, "_") = 0 Then
            If Len(Name) >= 3 And Len(Name) <= 16 And ValidName(Name) Then
                UserRS.Index = "Name"
                UserRS.Seek "=", Name
                If (UserRS.NoMatch) And GuildNum(Name) = 0 And NPCNum(Name) = 0 Then
                    If Index >= 1 And Index <= MaxUsers Then
                        With player(Index)
                            If .Mode = modePlaying Then
                                .Name = Name
                                SetPlayerName = 1
                                SendAll2 Chr2(64) + Chr2(Index) + .Name
                            End If
                        End With
                    End If
                Else
                    SetPlayerName = 0
                Exit Function
            End If
            Else
                SetPlayerName = 0
            End If
        Else
            SetPlayerName = 0
        End If
    Else
        SetPlayerName = 0
    End If
End If

End Function

Function GetStringFlag(ByVal Flag As Long) As Long
    If Flag >= 0 And Flag <= 255 Then
        GetStringFlag = NewString(World.StringFlag(Flag))
    End If
End Function

Function GetMonsterName(ByVal Index As Long) As Long
    If Index > 0 And Index <= MAXITEMS Then
        GetMonsterName = NewString(monster(Index).Name)
    End If
End Function

Function GetMonsterHP(ByVal Index As Long) As Long
    If Index > 0 And Index <= 255 Then
        GetMonsterHP = monster(Index).HP
    End If
End Function

Function GetMonsterDescription(ByVal Index As Long) As Long
    If Index > 0 And Index <= 255 Then
        GetMonsterDescription = NewString(monster(Index).Description)
    End If
End Function

Function GetMonsterSprite(ByVal Index As Long) As Long
    If Index > 0 And Index <= 255 Then
        GetMonsterSprite = monster(Index).sprite
    End If
End Function

Function GetMonsterExperience(ByVal Index As Long) As Long
    If Index > 0 And Index <= 255 Then
        GetMonsterExperience = monster(Index).Experience
    End If
End Function

Function GetMonsterLevel(ByVal Index As Long) As Long
    If Index > 0 And Index <= 255 Then
        GetMonsterLevel = monster(Index).Level
    End If
End Function

Function GetMonsterArmor(ByVal Index As Long) As Long
    If Index > 0 And Index <= 255 Then
        GetMonsterArmor = monster(Index).Armor
    End If
End Function

Sub DestroyMonster(ByVal mapNum As Long, ByVal monster As Long)
    If mapNum > 0 And mapNum <= 5000 Then
        If monster >= 0 And monster <= 9 Then
            If map(mapNum).monster(monster).HP > 0 Then
                map(mapNum).monster(monster).monster = 0
                SendToMap2 mapNum, Chr2(39) + Chr2(monster)
            End If
        End If
    End If
End Sub

Sub SetMapMonsterHP(ByVal mapNum As Long, ByVal monsterIndex As Long, ByVal HP As Long)
     If mapNum > 0 And mapNum <= 5000 Then
        If monsterIndex >= 0 And monsterIndex <= 9 Then
            With map(mapNum).monster(monsterIndex)
                If .monster > 0 Then
                    If HP > monster(.monster).HP Then
                        .HP = monster(.monster).HP
                    Else
                        .HP = HP
                    End If
                    SendToMap mapNum, Chr2(132) + Chr2(monsterIndex) + DoubleChar(.HP)
                End If
            End With
        End If
    End If
End Sub

Function GetMapMonsterHP(ByVal mapNum As Long, ByVal monsterIndex As Long) As Long
    If mapNum > 0 And mapNum <= 5000 Then
        If monsterIndex >= 0 And monsterIndex <= 9 Then
            If map(mapNum).monster(monsterIndex).monster > 0 Then
                GetMapMonsterHP = map(mapNum).monster(monsterIndex).HP
            End If
        End If
    End If
End Function

Sub SetMapMonsterFlag(ByVal mapNum As Long, ByVal monsterIndex As Long, ByVal FlagNum As Long, ByVal Value As Long)
    If mapNum > 0 And mapNum <= 5000 Then
        If monsterIndex >= 0 And monsterIndex <= 9 Then
            If FlagNum >= 0 And FlagNum <= 4 Then
                With map(mapNum).monster(monsterIndex)
                    .Flags(FlagNum) = Value
                End With
            End If
        End If
    End If
End Sub

Function GetMapMonsterFlag(ByVal mapNum As Long, ByVal monsterIndex As Long, ByVal FlagNum As Long) As Long
    If mapNum > 0 And mapNum <= 5000 Then
        If monsterIndex >= 0 And monsterIndex <= 9 Then
            If FlagNum >= 0 And FlagNum <= 4 Then
                GetMapMonsterFlag = map(mapNum).monster(monsterIndex).Flags(FlagNum)
            End If
        End If
    End If
End Function

Sub SetTile(ByVal x As Long, ByVal y As Long, ByVal Layer As Long, ByVal NewTile As Long)
    If CurEditMap.Num > 0 And CurEditMap.Num <= 5000 Then
        If Layer > 0 And Layer <= 5 Then
            If x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
                Select Case Layer
                    Case 1 'Ground
                        CurEditMap.Tile(x, y).Ground = NewTile
                    Case 2 'Ground2
                        CurEditMap.Tile(x, y).Ground2 = NewTile
                    Case 3 'BGTile1
                        CurEditMap.Tile(x, y).BGTile1 = NewTile
                    Case 5 'FGTile
                        CurEditMap.Tile(x, y).FGTile = NewTile
                End Select
            End If
        End If
    End If
End Sub

Sub SetTileAtt(ByVal x As Long, ByVal y As Long, ByVal Att As Long, ByVal Data1 As Long, ByVal Data2 As Long, ByVal Data3 As Long, ByVal Data4 As Long)
    If CurEditMap.Num > 0 And CurEditMap.Num <= 5000 Then
        If x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
            With CurEditMap.Tile(x, y)
                .Att = Att
                .AttData(0) = Data1
                .AttData(1) = Data2
                .AttData(2) = Data3
                .AttData(3) = Data4
            End With
        End If
    End If
End Sub

Sub SetMapExitDirection(ByVal direction As Long, ByVal map As Long)
    If map >= 0 And map <= 5000 Then
        Select Case direction
            Case 0 'Up
                CurEditMap.ExitUp = map
            Case 1 'Down
                CurEditMap.ExitDown = map
            Case 2 'Left
                CurEditMap.ExitLeft = map
            Case 3 'Right
                CurEditMap.ExitRight = map
        End Select
    End If
End Sub

Sub SetWall(ByVal x As Long, ByVal y As Long, ByVal Wall As Long)
    If CurEditMap.Num > 0 And CurEditMap.Num <= 5000 Then
        If x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
            CurEditMap.Tile(x, y).WallTile = Wall
        End If
    End If
End Sub

Function GetWall(ByVal x As Long, ByVal y As Long) As Long
    If CurEditMap.Num > 0 And CurEditMap.Num <= 5000 Then
        If x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
            GetWall = CurEditMap.Tile(x, y).WallTile
        End If
    End If
End Function

Function GetMapWall(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long) As Long
    If mapNum > 0 And mapNum <= 5000 Then
        If x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
            GetMapWall = map(mapNum).Tile(x, y).WallTile
        End If
    End If
End Function

Sub SetAnim(ByVal x As Long, ByVal y As Long, ByVal NumFrames As Long, ByVal FrameDelay As Long, ByVal Layer As Long, ByVal AnimDelay As Long, ByVal Flags As Long)
Dim A(1 To 2) As Byte
    If CurEditMap.Num > 0 And CurEditMap.Num <= 5000 Then
        If x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
            FrameDelay = FrameDelay \ 4
            A(1) = ((NumFrames * 16) Or FrameDelay)
            A(2) = ((AnimDelay * 4) Or Layer)
            CurEditMap.Tile(x, y).Anim(1) = A(1)
            CurEditMap.Tile(x, y).Anim(2) = A(2)
        End If
    End If
End Sub

Sub SetAnimBinary(ByVal x As Long, ByVal y As Long, ByVal part1 As Long, ByVal part2 As Long)
Dim A(1 To 2) As Byte
    If CurEditMap.Num > 0 And CurEditMap.Num <= 5000 Then
        If x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
            CurEditMap.Tile(x, y).Anim(1) = part1
            CurEditMap.Tile(x, y).Anim(2) = part2
        End If
    End If
End Sub

Function GetMapAnimBinary(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, ByVal part As Long) As Long
    If mapNum > 0 And mapNum <= 5000 Then
        If x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
            If part >= 1 And part <= 2 Then
                GetMapAnimBinary = map(mapNum).Tile(x, y).Anim(part)
            End If
        End If
    End If
End Function

Function GetTile(ByVal x As Long, ByVal y As Long, ByVal Layer As Long) As Long
    If CurEditMap.Num > 0 And CurEditMap.Num <= 5000 Then
        If Layer > 0 And Layer <= 5 Then
            If x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
                Select Case Layer
                    Case 1 'Ground
                        GetTile = CurEditMap.Tile(x, y).Ground
                    Case 2 'Ground2
                        GetTile = CurEditMap.Tile(x, y).Ground2
                    Case 3 'BGTile1
                        GetTile = CurEditMap.Tile(x, y).BGTile1
                    Case 5 'FGTile
                        GetTile = CurEditMap.Tile(x, y).FGTile
                End Select
            End If
        End If
    End If
End Function

Function GetMapTile(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, ByVal Layer As Long) As Long
    If mapNum > 0 And mapNum <= 5000 Then
        If Layer > 0 And Layer <= 5 Then
            If x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
                Select Case Layer
                    Case 1 'Ground
                        GetMapTile = map(mapNum).Tile(x, y).Ground
                    Case 2 'Ground2
                        GetMapTile = map(mapNum).Tile(x, y).Ground2
                    Case 3 'BGTile1
                        GetMapTile = map(mapNum).Tile(x, y).BGTile1
                    Case 5 'FGTile
                        GetMapTile = map(mapNum).Tile(x, y).FGTile
                End Select
            End If
        End If
    End If
End Function

Function ScriptLoadMap(ByVal mapNum As Long) As Long
    Dim MapData As String
    Dim x As Long, y As Long, A As Long

    If mapNum > 0 And mapNum <= 5000 Then
        MapRS.Seek "=", mapNum
        If Not MapRS.NoMatch Then
            MapData = MapRS!Data
        Else
            MapData = String(2379, 0)
        End If
        If Len(MapData) = 2374 Or Len(MapData) = 2379 Then
            With CurEditMap
                .Num = mapNum
                .Name = (Mid$(MapData, 1, 30))
                .Version = Asc(Mid$(MapData, 31, 1)) * 16777216 + Asc(Mid$(MapData, 32, 1)) * 65536 + Asc(Mid$(MapData, 33, 1)) * 256& + Asc(Mid$(MapData, 34, 1))
                .NPC = Asc(Mid$(MapData, 35, 1))
                .MIDI = Asc(Mid$(MapData, 36, 1))
                .ExitUp = Asc(Mid$(MapData, 37, 1)) * 256 + Asc(Mid$(MapData, 38, 1))
                .ExitDown = Asc(Mid$(MapData, 39, 1)) * 256 + Asc(Mid$(MapData, 40, 1))
                .ExitLeft = Asc(Mid$(MapData, 41, 1)) * 256 + Asc(Mid$(MapData, 42, 1))
                .ExitRight = Asc(Mid$(MapData, 43, 1)) * 256 + Asc(Mid$(MapData, 44, 1))
                .BootLocation.map = Asc(Mid$(MapData, 45, 1)) * 256 + Asc(Mid$(MapData, 46, 1))
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
                .RainColor = Asc(Mid$(MapData, 59, 1))
                For A = 0 To 4
                    .MonsterSpawn(A).monster = Asc(Mid$(MapData, 61 + A * 2))
                    .MonsterSpawn(A).Rate = Asc(Mid$(MapData, 62 + A * 2))
                Next A
                For y = 0 To 11
                    For x = 0 To 11
                        With .Tile(x, y)
                            A = 71 + y * 192 + x * 16
                            .Ground = Asc(Mid$(MapData, A, 1)) * 256 + Asc(Mid$(MapData, A + 1, 1))
                            .Ground2 = Asc(Mid$(MapData, A + 2, 1)) * 256 + Asc(Mid$(MapData, A + 3, 1))
                            .BGTile1 = Asc(Mid$(MapData, A + 4, 1)) * 256 + Asc(Mid$(MapData, A + 5, 1))
                            .Anim(1) = Asc(Mid$(MapData, A + 6, 1))
                            .Anim(2) = Asc(Mid$(MapData, A + 7, 1))
                            .FGTile = Asc(Mid$(MapData, A + 8, 1)) * 256 + Asc(Mid$(MapData, A + 9, 1))
                            .Att = Asc(Mid$(MapData, A + 10, 1))
                            .AttData(0) = Asc(Mid$(MapData, A + 11, 1))
                            .AttData(1) = Asc(Mid$(MapData, A + 12, 1))
                            .AttData(2) = Asc(Mid$(MapData, A + 13, 1))
                            .AttData(3) = Asc(Mid$(MapData, A + 14, 1))
                            .WallTile = Asc(Mid$(MapData, A + 15, 1))
                        End With
                    Next x
                Next y
            If Len(MapData) = 2379 Then
                For A = 0 To 4 '61 - 70
                    .MonsterSpawn(A).monster = .MonsterSpawn(A).monster + Asc(Mid$(MapData, A + 2375)) * 256
                Next A
            End If
            End With
            CurEditMap.Num = mapNum
            ScriptLoadMap = 1
            Exit Function
        End If
    End If
    ScriptLoadMap = 0
End Function

Sub ScriptSaveMap()
Dim mapNum As Long
mapNum = CurEditMap.Num
If mapNum > 0 And mapNum <= 5000 Then
    Dim MapData As String, ST1 As String * 30
    Dim x As Long, y As Long, A As Long
    With CurEditMap
        If .Version < 2147483647 Then
            .Version = .Version + 1
        Else
            .Version = 1
        End If
        ST1 = .Name
        MapData = ST1 + QuadChar(.Version) + Chr2(.NPC) + Chr2(.MIDI) + _
        DoubleChar$(CLng(.ExitUp)) + DoubleChar$(CLng(.ExitDown)) + _
        DoubleChar$(CLng(.ExitLeft)) + DoubleChar$(CLng(.ExitRight)) + _
        DoubleChar(CLng(.BootLocation.map)) + Chr2(.BootLocation.x) + _
        Chr2(.BootLocation.y) + Chr2(.Flags(0)) + Chr2(.Intensity) + _
        Chr2(.Flags(1)) + DoubleChar$(.Raining) + DoubleChar$(.Snowing) + Chr2(.Zone) + _
        Chr2(.Fog) + Chr2(.SnowColor) + Chr2(.RainColor) + " " + _
        Chr2(.MonsterSpawn(0).monster Mod 256) + Chr2(.MonsterSpawn(0).Rate) + _
        Chr2(.MonsterSpawn(1).monster Mod 256) + Chr2(.MonsterSpawn(1).Rate) + _
        Chr2(.MonsterSpawn(2).monster Mod 256) + Chr2(.MonsterSpawn(2).Rate) + _
        Chr2(.MonsterSpawn(3).monster Mod 256) + Chr2(.MonsterSpawn(3).Rate) + _
        Chr2(.MonsterSpawn(4).monster Mod 256) + Chr2(.MonsterSpawn(4).Rate)
        For y = 0 To 11
            For x = 0 To 11
                With .Tile(x, y)
                    If .Att = 24 Then
                        MapData = MapData + DoubleChar(CLng(.Ground)) + DoubleChar(CLng(.Ground2)) + DoubleChar(CLng(.BGTile1)) + Chr2(.Anim(1)) + Chr2(.Anim(2)) + DoubleChar(CLng(.FGTile)) + Chr2(.Att) + Chr2(.AttData(0)) + Chr2(.AttData(1)) + Chr2(.AttData(2)) + Chr2(.AttData(1)) + Chr2(.WallTile)
                    Else
                        MapData = MapData + DoubleChar(CLng(.Ground)) + DoubleChar(CLng(.Ground2)) + DoubleChar(CLng(.BGTile1)) + Chr2(.Anim(1)) + Chr2(.Anim(2)) + DoubleChar(CLng(.FGTile)) + Chr2(.Att) + Chr2(.AttData(0)) + Chr2(.AttData(1)) + Chr2(.AttData(2)) + Chr2(.AttData(3)) + Chr2(.WallTile)
                    End If
                End With
            Next x
        Next y
        MapData = MapData + Chr2(Int(.MonsterSpawn(0).monster / 256))
        MapData = MapData + Chr2(Int(.MonsterSpawn(1).monster / 256))
        MapData = MapData + Chr2(Int(.MonsterSpawn(2).monster / 256))
        MapData = MapData + Chr2(Int(.MonsterSpawn(3).monster / 256))
        MapData = MapData + Chr2(Int(.MonsterSpawn(4).monster / 256))
    End With
    MapRS.Seek "=", mapNum
     If MapRS.NoMatch Then
         MapRS.AddNew
         MapRS!Number = mapNum
     Else
         MapRS.Edit
     End If
     MapRS!Data = MapData
     MapRS.Update
     LoadMap mapNum, MapData
     For A = 0 To 9
         map(mapNum).Door(A).Att = 0
         map(mapNum).Door(A).Wall = 0
     Next A
     For A = 1 To currentMaxUser
         With player(A)
             If .Mode = modePlaying And .map = mapNum Then
                .ScriptUpdateMap = True
             End If
         End With
     Next A
End If
End Sub

Function GetMapAttData(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, ByVal AttData As Long) As Long
    If mapNum > 0 And mapNum <= 5000 Then
        If x >= 0 And x <= 11 And y >= 0 And y <= 11 Then
            If AttData >= 0 And AttData <= 3 Then
                GetMapAttData = map(mapNum).Tile(x, y).AttData(AttData)
            End If
        End If
    End If
End Function


Function GetMapExitDirection(ByVal mapNum As Long, ByVal direction As Long) As Long
    If mapNum > 0 And mapNum <= 5000 Then
        If direction >= 0 And direction <= 3 Then
            Select Case direction
                Case 0
                    GetMapExitDirection = map(mapNum).ExitUp
                Case 1
                    GetMapExitDirection = map(mapNum).ExitDown
                Case 2
                    GetMapExitDirection = map(mapNum).ExitLeft
                Case 3
                    GetMapExitDirection = map(mapNum).ExitRight
            End Select
        End If
    End If
End Function

Function GetPlayerDirection(ByVal Index As Long) As Long
    If Index > 0 And Index <= MaxUsers Then
        GetPlayerDirection = player(Index).D
    End If
End Function

Sub AddMapMonsterQueueMove(ByVal mapNum As Long, ByVal MonsterNum As Long, ByVal direction As Long)
    If mapNum > 0 And mapNum <= 5000 Then
        If MonsterNum >= 0 And MonsterNum <= 9 Then
            If map(mapNum).monster(MonsterNum).monster > 0 Then
                If direction >= 0 And direction < 4 Then
                    With map(mapNum).monster(MonsterNum)
                        If .CurrentQueue < 15 Then
                            .MonsterQueue(.CurrentQueue).Action = QUEUE_MOVE
                            .MonsterQueue(.CurrentQueue).lngData = direction
                            .CurrentQueue = .CurrentQueue + 1
                        End If
                    End With
                End If
            End If
        End If
    End If
End Sub
Sub AddMapMonsterQueueShift(ByVal mapNum As Long, ByVal MonsterNum As Long, ByVal direction As Long)
    If mapNum > 0 And mapNum <= 5000 Then
        If MonsterNum >= 0 And MonsterNum <= 9 Then
            If map(mapNum).monster(MonsterNum).monster > 0 Then
                If direction >= 0 And direction < 4 Then
                    With map(mapNum).monster(MonsterNum)
                        If .CurrentQueue < 15 Then
                            .MonsterQueue(.CurrentQueue).Action = QUEUE_SHIFT
                            .MonsterQueue(.CurrentQueue).lngData = direction
                            .CurrentQueue = .CurrentQueue + 1
                        End If
                    End With
                End If
            End If
        End If
    End If
End Sub
Sub setMonsterDirection(ByVal mapNum As Long, ByVal MonsterNum As Long, ByVal direction As Long)
    If mapNum > 0 And mapNum <= 5000 Then
        If MonsterNum >= 0 And MonsterNum <= 9 Then
            If map(mapNum).monster(MonsterNum).monster > 0 Then
                If direction >= 0 And direction < 4 Then
                    With map(mapNum).monster(MonsterNum)
                        If .CurrentQueue < 15 Then
                            .MonsterQueue(.CurrentQueue).Action = QUEUE_TURN
                            .MonsterQueue(.CurrentQueue).lngData = direction
                            .CurrentQueue = .CurrentQueue + 1
                        End If
                    End With
                End If
            End If
        End If
    End If
End Sub

Sub AddMapMonsterQueueScript(ByVal mapNum As Long, ByVal MonsterNum As Long, ByVal Script As String, ByVal Param1 As Long, ByVal Param2 As Long)
    If mapNum > 0 And mapNum <= 5000 Then
        If MonsterNum >= 0 And MonsterNum <= 9 Then
            If map(mapNum).monster(MonsterNum).monster > 0 Then
                With map(mapNum).monster(MonsterNum)
                    If .CurrentQueue < 15 Then
                        .MonsterQueue(.CurrentQueue).Action = QUEUE_SCRIPT
                        .MonsterQueue(.CurrentQueue).lngData = Param1
                        .MonsterQueue(.CurrentQueue).lngData1 = Param2
                        .MonsterQueue(.CurrentQueue).strData = StrConv(Script, vbUnicode)
                        .CurrentQueue = .CurrentQueue + 1
                    End If
                End With
            End If
        End If
    End If
End Sub

Sub AddMapMonsterQueuePause(ByVal mapNum As Long, ByVal MonsterNum As Long, ByVal Length As Long)
    Dim A As Long
    If mapNum > 0 And mapNum <= 5000 Then
        If MonsterNum >= 0 And MonsterNum <= 9 Then
            If map(mapNum).monster(MonsterNum).monster > 0 Then
                With map(mapNum).monster(MonsterNum)
                    If .CurrentQueue < 15 Then
                        For A = .CurrentQueue To .CurrentQueue + Length
                            If A < 15 Then
                                .MonsterQueue(.CurrentQueue).Action = QUEUE_PAUSE
                                .CurrentQueue = .CurrentQueue + 1
                            Else
                                Exit For
                            End If
                        Next A
                    End If
                End With
            End If
        End If
    End If
End Sub

Sub ClearMapMonsterQueue(ByVal mapNum As Long, ByVal MonsterNum As Long)
    If mapNum > 0 And mapNum <= 5000 Then
        If MonsterNum >= 0 And MonsterNum <= 9 Then
            If map(mapNum).monster(MonsterNum).monster > 0 Then
                With map(mapNum).monster(MonsterNum)
                    'For A = 0 To 15
                        .MonsterQueue(0).Action = QUEUE_EMPTY
                        .CurrentQueue = 0
                    'Next A
                End With
            End If
        End If
    End If
End Sub


'-------------------------- Widget Functions -----------------------------------
Sub StartWidgetMenu(ByVal Index As Long, ByVal Key As String, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal ReturnScript As String)
    Dim St As String
    If Index > 0 And Index <= MaxUsers Then
        player(Index).Widgets.NumWidgets = 0
        With player(Index).Widgets
            If .NumWidgets = 0 Then
                '3 Denotes its a widget menu
                St = Chr2(3) + DoubleChar(CInt(x)) + DoubleChar(CInt(y)) + DoubleChar(CInt(Width)) + DoubleChar(CInt(Height)) + StrConv(Key, vbUnicode)
                St = Chr2(Len(St)) + St
                .WidgetString = St
                .WidgetScript = StrConv(ReturnScript, vbUnicode)
                
                .NumWidgets = 1
                .MenuVisible = True
                ReDim .Widgets(1 To 1)
                With .Widgets(1)
                    .Key = StrConv(Key, vbUnicode)
                    .Type = WIDGET_FRAME
                End With
            Else
                PrintLog player(Index).Name & " had a menu created with existing widgets..."
            End If
        End With
    End If
End Sub

Sub AddWidgetButton(ByVal Index As Long, ByVal Key As String, ByVal x As Long, ByVal y As Long, ByVal Caption As String, ByVal Flags As Long)
    Dim St As String, ST1 As String
    If Index > 0 And Index <= MaxUsers Then
        With player(Index).Widgets
            If .NumWidgets > 0 And .MenuVisible Then
                ST1 = StrConv(Key, vbUnicode)
                St = Chr2(WIDGET_BUTTON) + DoubleChar(CInt(x)) + DoubleChar(CInt(y)) + QuadChar(Flags) + Chr2(Len(ST1)) + ST1 + StrConv(Caption, vbUnicode)
                St = Chr2(Len(St)) + St
                player(Index).Widgets.WidgetString = player(Index).Widgets.WidgetString + St
                
                .NumWidgets = .NumWidgets + 1
                ReDim Preserve .Widgets(1 To .NumWidgets)
                With .Widgets(.NumWidgets)
                    .Type = WIDGET_BUTTON
                    .Key = StrConv(Key, vbUnicode)
                End With
            Else
                PrintLog player(Index).Name & " had widgets created with no menu..."
            End If
        End With
    End If
End Sub

Sub AddWidgetLabel(ByVal Index As Long, ByVal Key As String, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Caption As String, ByVal Flags As Long)
    Dim St As String, ST1 As String
    If Index > 0 And Index <= MaxUsers Then
        With player(Index).Widgets
            If .NumWidgets > 0 And .MenuVisible Then
                ST1 = StrConv(Key, vbUnicode)
                St = Chr2(WIDGET_LABEL) + DoubleChar(CInt(x)) + DoubleChar(CInt(y)) + DoubleChar(CInt(Width)) + DoubleChar(CInt(Height)) + QuadChar(Flags) + Chr2(Len(ST1)) + ST1 + StrConv(Caption, vbUnicode)
                St = Chr2(Len(St)) + St
                player(Index).Widgets.WidgetString = player(Index).Widgets.WidgetString + St
        
                .NumWidgets = .NumWidgets + 1
                ReDim Preserve .Widgets(1 To .NumWidgets)
                With .Widgets(.NumWidgets)
                    .Type = WIDGET_LABEL
                    .Key = StrConv(Key, vbUnicode)
                End With
            Else
                PrintLog player(Index).Name & " had widgets created with no menu..."
            End If
        End With
    End If
End Sub

Sub AddWidgetTextBox(ByVal Index As Long, ByVal Key As String, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal Caption As String, ByVal Flags As Long)
    Dim St As String, ST1 As String
    If Index > 0 And Index <= MaxUsers Then
        With player(Index).Widgets
            If .NumWidgets > 0 And .MenuVisible Then
                ST1 = StrConv(Key, vbUnicode)
                St = Chr2(WIDGET_TEXTBOX) + DoubleChar(CInt(x)) + DoubleChar(CInt(y)) + DoubleChar(CInt(Width)) + DoubleChar(CInt(Height)) + QuadChar(Flags) + Chr2(Len(ST1)) + ST1 + StrConv(Caption, vbUnicode)
                St = Chr2(Len(St)) + St
                player(Index).Widgets.WidgetString = player(Index).Widgets.WidgetString + St
        
                .NumWidgets = .NumWidgets + 1
                ReDim Preserve .Widgets(1 To .NumWidgets)
                With .Widgets(.NumWidgets)
                    .Type = WIDGET_TEXTBOX
                    .Key = StrConv(Key, vbUnicode)
                    .Data(0) = Flags
                End With
            Else
                PrintLog player(Index).Name & " had widgets created with no menu..."
            End If
        End With
    End If
End Sub

Sub AddWidgetImage(ByVal Index As Long, ByVal Key As String, ByVal fileName As String, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Flags As Long)
    Dim St As String, ST1 As String, st2 As String
    With player(Index).Widgets
        If .NumWidgets > 0 And .MenuVisible Then
            ST1 = StrConv(Key, vbUnicode)
            st2 = StrConv(fileName, vbUnicode)
            St = Chr2(WIDGET_IMAGE) + DoubleChar$(CInt(x)) + DoubleChar$(CInt(y)) + DoubleChar$(CInt(Width)) + DoubleChar$(CInt(Height)) + DoubleChar$(CInt(SrcX)) + DoubleChar$(CInt(SrcY)) + QuadChar$(Flags) + Chr2(Len(ST1)) + ST1 + Chr2(Len(st2)) + st2
            St = Chr2(Len(St)) + St
            player(Index).Widgets.WidgetString = player(Index).Widgets.WidgetString + St
            
            .NumWidgets = .NumWidgets + 1
            ReDim Preserve .Widgets(1 To .NumWidgets)
            With .Widgets(.NumWidgets)
                .Type = WIDGET_IMAGE
                .Key = ST1
                .Data(0) = Flags
            End With
        Else
            PrintLog player(Index).Name & " had widgets created with no menu..."
        End If
    End With
End Sub

Sub SendWidgetString(ByVal Index As Long)
    If Index > 0 And Index <= MaxUsers Then
        With player(Index)
            'If Len(player(Index).Widgets.WidgetString) <= 65536 Then
                SendSocket2 Index, Chr2(124) + player(Index).Widgets.WidgetString
            'Else
                'SendSocket2
            'End If
            player(Index).Widgets.WidgetString = ""
            player(Index).Widgets.MenuVisible = True
        End With
    End If
End Sub

Function GetWidgetValueLong(ByVal Index As Long, ByVal Key As String) As Long
    Dim St As String
    St = StrConv(Key, vbUnicode)
    If Index > 0 And Index <= MaxUsers Then
        Dim A As Long
        With player(Index).Widgets
            If .NumWidgets > 0 And .MenuVisible Then
                For A = 1 To .NumWidgets
                    If .Widgets(A).Key = St Then
                        GetWidgetValueLong = .Widgets(A).lngData
                    End If
                Next A
            End If
        End With
    End If
End Function

Function GetWidgetValueString(ByVal Index As Long, ByVal Key As String) As Long
    Dim St As String
    St = StrConv(Key, vbUnicode)
    If Index > 0 And Index <= MaxUsers Then
        Dim A As Long
        With player(Index).Widgets
            If .NumWidgets > 0 And .MenuVisible Then
                For A = 1 To .NumWidgets
                    If .Widgets(A).Key = St Then
                        GetWidgetValueString = NewString(.Widgets(A).strData)
                    End If
                Next A
            End If
        End With
    End If
End Function
'-------------------------------------------------------------------------------

Sub MapReset(ByVal mapNum As Long)
    If mapNum > 0 And mapNum <= 5000 Then
        SoftResetMap mapNum
    End If
End Sub

Function GetMapMonsterCount(ByVal mapNum As Long) As Long
    Dim A As Long, B As Long
    If mapNum > 0 And mapNum <= 5000 Then
        For A = 0 To 9
            If map(mapNum).monster(A).monster > 0 Then
                B = B + 1
            End If
        Next A
    End If
    GetMapMonsterCount = B
End Function

Sub SetPlayerFrozen(ByVal Index As Long, ByVal Frozen As Long)
    SendSocket2 Index, Chr2(101) + Chr2(Frozen) + Chr2(player(Index).x * 16 + player(Index).y)
End Sub

Sub FadeMap(ByVal Index As Long, ByVal StartValue As Long, ByVal EndValue As Long, ByVal StepSize As Long)
    SendSocket2 Index, Chr2(102) + Chr2(StartValue) + Chr2(StartValue) + Chr2(EndValue) + Chr2(StepSize)
End Sub

Function GetSkillLevel(ByVal Index As Long, ByVal SkillNum As Long) As Long
    If Index > 0 And Index <= MaxUsers And SkillNum > 0 And SkillNum < MAX_SKILLS Then
        GetSkillLevel = player(Index).SkillLevel(SkillNum)
    End If
End Function

Sub SetSkillLevel(ByVal Index As Long, ByVal SkillNum As Long, ByVal Level As Long)
    If Index > 0 And Index <= MaxUsers And SkillNum > 0 And SkillNum < MAX_SKILLS Then
        player(Index).SkillLevel(SkillNum) = Level
        SendSocket2 Index, Chr2(123) + Chr2(1) + Chr2(SkillNum) + Chr2(player(Index).SkillLevel(SkillNum))
        SendSocket2 Index, Chr2(123) + Chr(SkillNum) + QuadChar(0)
    End If
End Sub

Sub SetPlayerStatusEffect(ByVal Index As Long, ByVal StatusEffect As Long, ByVal timer As Long, ByVal Data0 As Long, ByVal Data1 As Long, ByVal Data2 As Long, ByVal Data3 As Long)
    If Index > 0 And Index <= MaxUsers Then
        If StatusEffect > 0 And StatusEffect < 32 Then
            With player(Index)
                .StatusEffect = (.StatusEffect Or (2 ^ StatusEffect))
                .StatusData(StatusEffect).timer = timer
                .StatusData(StatusEffect).Data(0) = Data0
                .StatusData(StatusEffect).Data(1) = Data1
                .StatusData(StatusEffect).Data(2) = Data2
                .StatusData(StatusEffect).Data(3) = Data3
                SendToMap2 .map, Chr2(112) + Chr2(Index) + QuadChar(.StatusEffect)
            End With
        End If
    End If
End Sub

Sub SetStatMod(ByVal Index As Long, ByVal stat As Long, ByVal amount As Long)
    If Index > 0 And Index <= MaxUsers Then
        If stat > 0 And stat < 7 Then
            If amount > 255 Then amount = 255
            Select Case stat
                Case 1
                    player(Index).StrMod(0) = amount
                Case 2
                    player(Index).EndMod(0) = amount
                Case 3
                    player(Index).AgiMod(0) = amount
                Case 4
                    player(Index).WisMod(0) = amount
                Case 5
                    player(Index).ConMod(0) = amount
                Case 6
                    player(Index).IntMod(0) = amount
            End Select
            
            CalculateStats Index
        End If
    End If
End Sub

Function CreatePlayerProjectile(ByVal playerNum As Long, ByVal direction As Long, ByVal EffectNum As Long, ByVal damage As Long, ByVal ProjType As Long, ByVal speed As Long) As Long
CreatePlayerProjectile = 0
    If playerNum > 0 And playerNum <= MaxUsers Then
        If player(playerNum).ProjectileDamage = 0 Then
            If direction >= 0 And direction <= 3 Then
                If EffectNum >= 0 Then
                    If speed >= 0 And speed <= 255 Then
                        player(playerNum).ProjectileDamage = damage
                        player(playerNum).ProjectileType = ProjType
                        SendToMap2 player(playerNum).map, Chr2(125) + Chr2(playerNum) + DoubleChar$(EffectNum) + Chr2(player(playerNum).x * 16 + player(playerNum).y) + Chr2(direction) + Chr2(speed)
                        CreatePlayerProjectile = 1
                    End If
                End If
            End If
        End If
    End If
End Function

Function CreatePlayerLitProjectile(ByVal playerNum As Long, ByVal direction As Long, ByVal EffectNum As Long, ByVal damage As Long, ByVal ProjType As Long, ByVal speed As Long, ByVal Radius As Long, ByVal Intensity As Long, ByVal red As Long, ByVal green As Long, ByVal blue As Long) As Long
CreatePlayerLitProjectile = 0
    If playerNum > 0 And playerNum <= MaxUsers Then
        If player(playerNum).ProjectileDamage = 0 Then
            If direction >= 0 And direction <= 3 Then
                If EffectNum >= 0 Then
                    If speed >= 0 And speed <= 255 Then
                        If Intensity <= 255 Then
                            If Radius <= 255 Then
                                If red <= 255 And green <= 255 And blue <= 255 Then
                                    player(playerNum).ProjectileDamage = damage
                                    player(playerNum).ProjectileType = ProjType
                                    SendToMap2 player(playerNum).map, Chr2(125) + Chr2(playerNum) + DoubleChar$(EffectNum) + Chr2(player(playerNum).x * 16 + player(playerNum).y) + Chr2(direction) + Chr2(speed) + Chr2(Radius) + Chr2(Intensity) + Chr2(red) + Chr2(green) + Chr2(blue)
                                    CreatePlayerLitProjectile = 1
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Function GetPrefix(ByVal Index As Long, ByVal InvIndex As Long, ByVal pOrS As Long) As Long

    GetPrefix = 0

    If Index > 0 And Index <= MaxUsers Then
        If InvIndex > 0 And InvIndex <= 25 Then
            If InvIndex <= 20 Then
                If player(Index).Inv(InvIndex).Object > 0 Then
                    If pOrS = 0 Then
                        GetPrefix = player(Index).Inv(InvIndex).prefix
                    ElseIf pOrS = 1 Then
                        GetPrefix = player(Index).Inv(InvIndex).suffix
                    ElseIf pOrS = 2 Then
                        GetPrefix = player(Index).Inv(InvIndex).Affix
                    End If
                End If
            Else
                If player(Index).Equipped(InvIndex - 20).Object > 0 Then
                    If pOrS = 0 Then
                        GetPrefix = player(Index).Equipped(InvIndex - 20).prefix
                    ElseIf pOrS = 1 Then
                        GetPrefix = player(Index).Equipped(InvIndex - 20).suffix
                    Else
                        GetPrefix = player(Index).Equipped(InvIndex - 20).Affix
                    End If
                End If
            End If
        End If
    End If
End Function

Function GetObjectFlag(ByVal Index As Long, ByVal InvIndex As Long, ByVal FlagNum As Long) As Long

    GetObjectFlag = 0

    If Index > 0 And Index <= MaxUsers Then
        If InvIndex > 0 And InvIndex <= 25 Then
            If InvIndex <= 20 Then
                If player(Index).Inv(InvIndex).Object > 0 Then
                    If FlagNum >= 0 And FlagNum < 4 Then
                        GetObjectFlag = player(Index).Inv(InvIndex).Flags(FlagNum)
                    End If
                End If
            Else
                If player(Index).Equipped(InvIndex - 20).Object > 0 Then
                    If FlagNum >= 0 And FlagNum < 4 Then
                        GetObjectFlag = player(Index).Equipped(InvIndex - 20).Flags(FlagNum)
                    End If
                End If
            End If
        End If
    End If
End Function

Sub SetObjectFlag(ByVal Index As Long, ByVal InvIndex As Long, ByVal FlagNum As Long, ByVal Value As Long)
If Index > 0 And Index <= MaxUsers Then
    If InvIndex > 0 And InvIndex <= 25 Then
        If InvIndex <= 20 Then
            If player(Index).Inv(InvIndex).Object > 0 Then
                If FlagNum >= 0 And FlagNum < 4 Then
                    player(Index).Inv(InvIndex).Flags(FlagNum) = Value
                End If
            End If
        Else
            If player(Index).Equipped(InvIndex - 20).Object > 0 Then
                If FlagNum >= 0 And FlagNum < 4 Then
                    player(Index).Equipped(InvIndex - 20).Flags(FlagNum) = Value
                End If
            End If
        End If
    End If
End If
End Sub

Function GetObjectColor(ByVal Index As Long, ByVal InvIndex As Long) As Long

    GetObjectColor = 0

    If Index > 0 And Index <= MaxUsers Then
        If InvIndex > 0 And InvIndex <= 25 Then
            If InvIndex <= 20 Then
                If player(Index).Inv(InvIndex).Object > 0 Then
                        GetObjectColor = player(Index).Inv(InvIndex).ObjectColor
                End If
            Else
                If player(Index).Equipped(InvIndex - 20).Object > 0 Then
                        GetObjectColor = player(Index).Equipped(InvIndex - 20).ObjectColor
                End If
            End If
        End If
    End If
End Function

Sub SetObjectColor(ByVal Index As Long, ByVal InvIndex As Long, ByVal ObjectColor As Long)
If Index > 0 And Index <= MaxUsers Then
    If InvIndex > 0 And InvIndex <= 25 Then
        If ObjectColor >= 0 And ObjectColor <= 255 Then
        If InvIndex <= 20 Then
            If player(Index).Inv(InvIndex).Object > 0 Then
                player(Index).Inv(InvIndex).ObjectColor = ObjectColor
            End If
        Else
            If player(Index).Equipped(InvIndex - 20).Object > 0 Then
                player(Index).Equipped(InvIndex - 20).ObjectColor = ObjectColor
            End If
        End If
        End If
    End If
End If
If InvIndex <= 20 Then
    With player(Index).Inv(InvIndex)
        SendSocket2 Index, Chr2(17) + Chr2(InvIndex) + DoubleChar(.Object) + QuadChar$(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)
    End With
Else
    With player(Index).Equipped(InvIndex - 20)
        SendSocket2 Index, Chr2(17) + Chr2(InvIndex) + DoubleChar(.Object) + QuadChar$(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)
    End With
End If
End Sub


Function GetPrefixVal(ByVal Index As Long, ByVal InvIndex As Long, ByVal pOrS As Long) As Long

    GetPrefixVal = 0
    
    If Index > 0 And Index <= MaxUsers Then
        If InvIndex > 0 And InvIndex <= 25 Then
            If InvIndex <= 20 Then
                If player(Index).Inv(InvIndex).Object > 0 Then
                    If pOrS = 0 Then
                        If player(Index).Inv(InvIndex).prefix > 0 Then
                            GetPrefixVal = (player(Index).Inv(InvIndex).prefixVal)
                        End If
                    ElseIf pOrS = 1 Then
                        If player(Index).Inv(InvIndex).suffix > 0 Then
                            GetPrefixVal = player(Index).Inv(InvIndex).SuffixVal
                        End If
                    ElseIf pOrS = 2 Then
                        If player(Index).Inv(InvIndex).Affix > 0 Then
                            GetPrefixVal = player(Index).Inv(InvIndex).AffixVal
                        End If
                    End If
                End If
            Else
                If player(Index).Equipped(InvIndex - 20).Object > 0 Then
                    If pOrS = 0 Then
                        If player(Index).Equipped(InvIndex - 20).prefix > 0 Then
                            GetPrefixVal = (player(Index).Equipped(InvIndex - 20).prefixVal)
                        End If
                    ElseIf pOrS = 1 Then
                        If player(Index).Equipped(InvIndex - 20).suffix > 0 Then
                            GetPrefixVal = player(Index).Equipped(InvIndex - 20).SuffixVal
                        End If
                    ElseIf pOrS = 2 Then
                        If player(Index).Equipped(InvIndex - 20).Affix > 0 Then
                            GetPrefixVal = player(Index).Equipped(InvIndex - 20).AffixVal
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Function GetPrefixProperty(ByVal Index As Long, ByVal Property As String) As Long
    Property = StrConv(Property, vbUnicode)
    With prefix(Index)
        Select Case UCase$(Property)
            Case "NAME"
                GetPrefixProperty = NewString(.Name)
            Case "SUFFIX", "FLAGS"
                GetPrefixProperty = .Flags
            Case "MIN", "MINIMUM"
                GetPrefixProperty = .Min
            Case "MAX", "MAXIMUM"
                GetPrefixProperty = .Max
            Case "MOD", "MODTYPE", "MODIFIER"
                GetPrefixProperty = .ModType
            Case "RARITY"
                GetPrefixProperty = .Rarity
        End Select
    End With
End Function
Sub SetPrefix(ByVal Index As Long, ByVal InvIndex As Long, ByVal prefixNum As Long, ByVal Value As Long, ByVal pOrSorA As Long)
    If Index > 0 And Index <= MaxUsers Then
        If prefixNum >= 0 And prefixNum < 256 Then
            If Value = 256 Then
                Value = (Rnd * ((prefix(prefixNum).Max + 1) - prefix(prefixNum).Min)) + prefix(prefixNum).Min
            End If
            If prefixNum = 0 Then Value = 0
            If InvIndex > 0 And InvIndex <= 25 Then
                If InvIndex <= 20 Then
                    If player(Index).Inv(InvIndex).Object > 0 Then
                        If pOrSorA = 0 Then
                            player(Index).Inv(InvIndex).prefix = prefixNum
                            player(Index).Inv(InvIndex).prefixVal = Value
                        ElseIf pOrSorA = 1 Then
                            player(Index).Inv(InvIndex).suffix = prefixNum
                            player(Index).Inv(InvIndex).SuffixVal = Value
                        ElseIf pOrSorA = 2 Then
                            player(Index).Inv(InvIndex).Affix = prefixNum
                            player(Index).Inv(InvIndex).AffixVal = Value
                        End If
                    End If
                Else
                    If player(Index).Equipped(InvIndex - 20).Object > 0 Then
                        If pOrSorA = 0 Then
                            player(Index).Equipped(InvIndex - 20).prefix = prefixNum
                            player(Index).Equipped(InvIndex - 20).prefixVal = Value
                        ElseIf pOrSorA = 1 Then
                            player(Index).Equipped(InvIndex - 20).suffix = prefixNum
                            player(Index).Equipped(InvIndex - 20).SuffixVal = Value
                        ElseIf pOrSorA = 2 Then
                            player(Index).Equipped(InvIndex - 20).Affix = prefixNum
                            player(Index).Equipped(InvIndex - 20).AffixVal = Value
                        End If
                    End If
                End If
            End If
            If InvIndex <= 20 Then
                With player(Index).Inv(InvIndex)
                    SendSocket2 Index, Chr2(17) + Chr2(InvIndex) + DoubleChar(.Object) + QuadChar$(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)
                End With
            Else
                With player(Index).Equipped(InvIndex - 20)
                    SendSocket2 Index, Chr2(17) + Chr2(InvIndex) + DoubleChar(.Object) + QuadChar$(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)
                End With
            End If
        End If
    End If
End Sub

Sub DisplayUnzFont(ByVal Index As Long, ByVal x As Long, ByVal y As Long, ByVal Text As String, ByVal Life As Long)
    If Index > 0 And Index < MaxUsers Then
        If x > 0 And x < 384 Then
            If y > 0 And y < 384 Then
                SendSocket2 Index, Chr2(103) + DoubleChar(CInt(x)) + DoubleChar(CInt(y)) + Chr2(16) + Chr2(Life) + Chr2(0) + StrConv(Text, vbUnicode)
            End If
        End If
    End If
End Sub

Sub ScriptCreateFloatingText(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, ByVal Text As String, ByVal Color As Long)
    If mapNum > 0 And mapNum <= 5000 Then
        If x < 11 And y < 11 Then
            If Len(Text) > 0 Then
                If Color >= 0 And Color <= 15 Then
                    CreateFloatingText mapNum, x, y, Color, StrConv(Text, vbUnicode)
                End If
            End If
        End If
    End If
End Sub

Sub ScriptCreatePlayerFloatingText(ByVal Index As Long, ByVal Text As String, ByVal Color As Long)
    If Index > 0 And Index <= MaxUsers Then
        If Len(Text) > 0 Then
            If Color >= 0 And Color <= 15 Then
                CreateFloatingText player(Index).map, player(Index).x, player(Index).y, Color, StrConv(Text, vbUnicode)
            End If
        End If
    End If
End Sub

Sub KillPLayerSounds(ByVal Index As Long)
    SendSocket2 Index, Chr2(144)
End Sub

Sub SetMapWeather(ByVal mapNum As Long, ByVal Weather As Long, ByVal Intensity As Long)
    If mapNum > 0 And mapNum <= 5000 Then
        If Weather < 5 And Intensity < 10000 Then
            Select Case Weather
                Case 0
                    map(mapNum).Raining.Raining = Intensity
                    If map(mapNum).NumPlayers Then SendToMap2 mapNum, Chr2(120) + Chr2(1) + DoubleChar$(map(mapNum).Raining.Raining)
                Case 1
                    map(mapNum).Snowing = Intensity
                    If map(mapNum).NumPlayers Then SendToMap2 mapNum, Chr2(120) + Chr2(2) + DoubleChar$(map(mapNum).Snowing)
                Case 2
                    If Intensity <= 30 Then
                        map(mapNum).Fog = Intensity
                    End If
                    If map(mapNum).NumPlayers Then SendToMap2 mapNum, Chr2(120) + Chr2(3) + Chr2(map(mapNum).Fog)
                Case 3
                    map(mapNum).FlickerDark = Intensity
                    If map(mapNum).NumPlayers Then SendToMap2 mapNum, Chr2(120) + Chr2(4) + Chr2(map(mapNum).FlickerDark)
                Case 4
                    map(mapNum).FlickerLength = Intensity
                    If map(mapNum).NumPlayers Then SendToMap2 mapNum, Chr2(120) + Chr2(5) + Chr2(map(mapNum).FlickerLength)
            End Select
        End If
    End If
End Sub

Sub SetZoneWeather(ByVal ZoneNum As Long, ByVal Weather As Long, ByVal Intensity As Long)
    Dim mapNum As Long
    If ZoneNum > 0 And ZoneNum <= 255 Then
        If Weather < 5 And Intensity < 10000 Then
            For mapNum = 1 To 5000
                If map(mapNum).Zone = ZoneNum Then
                    Select Case Weather
                        Case 0
                            map(mapNum).Raining.Raining = Intensity
                            If map(mapNum).NumPlayers Then SendToMap2 mapNum, Chr2(120) + Chr2(1) + DoubleChar$(map(mapNum).Raining.Raining)
                        Case 1
                            map(mapNum).Snowing = Intensity
                            If map(mapNum).NumPlayers Then SendToMap2 mapNum, Chr2(120) + Chr2(2) + DoubleChar$(map(mapNum).Snowing)
                        Case 2
                            If Intensity <= 30 Then
                                map(mapNum).Fog = Intensity
                            End If
                            If map(mapNum).NumPlayers Then SendToMap2 mapNum, Chr2(120) + Chr2(3) + Chr2(map(mapNum).Fog)
                        Case 3
                            map(mapNum).FlickerDark = Intensity
                            If map(mapNum).NumPlayers Then SendToMap2 mapNum, Chr2(120) + Chr2(4) + Chr2(map(mapNum).FlickerDark)
                        Case 4
                            map(mapNum).FlickerLength = Intensity
                            If map(mapNum).NumPlayers Then SendToMap2 mapNum, Chr2(120) + Chr2(5) + Chr2(map(mapNum).FlickerLength)
                        Case 5
                            
                    End Select
                End If
            Next mapNum
        End If
    End If
End Sub

Function GetMapWeather(ByVal mapNum As Long, ByVal Weather As Long) As Long
    GetMapWeather = 0
    If mapNum > 0 And mapNum <= 5000 Then
        If Weather < 3 Then
            Select Case Weather
                Case 0
                    GetMapWeather = map(mapNum).Raining.Raining
                Case 1
                    GetMapWeather = map(mapNum).Snowing
                Case 2
                    GetMapWeather = map(mapNum).Fog
            End Select
        End If
    End If
End Function

Sub SetWeatherVariable(ByVal mapNum As Long, ByVal Weather As Long, ByVal Var As Long, ByVal Val As Long)
    
    If mapNum > 0 And mapNum <= 500 Then
        Select Case Weather
            Case 0 'Rain
                Select Case Var
                    Case 0 'Red
                        If Val >= 0 And Val <= 255 Then
                            map(mapNum).Raining.R = Val
                            SendToMap2 mapNum, Chr2(120) + Chr2(1) + Chr2(Var) + Chr2(Val)
                        End If
                    Case 1 'Green
                        If Val >= 0 And Val <= 255 Then
                            map(mapNum).Raining.G = Val
                            SendToMap2 mapNum, Chr2(120) + Chr2(1) + Chr2(Var) + Chr2(Val)
                        End If
                    Case 2 'Blue
                        If Val >= 0 And Val <= 255 Then
                            map(mapNum).Raining.B = Val
                            SendToMap2 mapNum, Chr2(120) + Chr2(1) + Chr2(Var) + Chr2(Val)
                        End If
                    Case 3 'Flash Chance
                        If Val >= 0 And Val <= 255 Then
                            map(mapNum).Raining.FlashChance = Val
                            SendToMap2 mapNum, Chr2(120) + Chr2(1) + Chr2(Var) + Chr2(Val)
                        End If
                    Case 4 'Flash Length
                        If Val >= 0 And Val <= 255 Then
                            map(mapNum).Raining.FlashLength = Val
                            SendToMap2 mapNum, Chr2(120) + Chr2(1) + Chr2(Var) + Chr2(Val)
                        End If
                End Select
            Case 0 'Snow
                
        End Select
    End If
End Sub

Sub SetPlayerRenown(ByVal Index As Long, ByVal Renown As Long)
Dim A As Long
    If Index > 0 And Index <= MaxUsers Then
        player(Index).Renown = Renown
        If player(Index).Guild > 0 Then
            With Guild(player(Index).Guild)
                A = FindGuildMember(player(Index).Name, CLng(player(Index).Guild))
                If A >= 0 Then
                    With Guild(player(Index).Guild).Member(A)
                        .Renown = Renown
                    End With
                    GuildRS.Bookmark = Guild(player(Index).Guild).Bookmark
                    GuildRS.Edit
                    GuildRS("MemberRenown" + CStr(A)) = Renown
                    GuildRS.Update
                End If
                
            End With
            UpdateGuildInfo CByte(player(Index).Guild)
        End If
    End If
End Sub

Function GetPlayerRenown(ByVal Index As Long) As Long
    If Index > 0 And Index <= MaxUsers Then
        GetPlayerRenown = player(Index).Renown
    End If
End Function

Sub SetPlayerSkillPoints(ByVal Index As Long, ByVal SkillPoints As Long)
    If Index > 0 And Index <= MaxUsers Then
        player(Index).SkillPoints = SkillPoints
        SendSocket2 Index, Chr2(59) + DoubleChar(CInt(player(Index).StatPoints)) + DoubleChar(CInt(player(Index).SkillPoints))
    End If
End Sub

Function GetPlayerSkillPoints(ByVal Index As Long) As Long
    If Index > 0 And Index <= MaxUsers Then
        GetPlayerSkillPoints = player(Index).SkillPoints
    End If
End Function

Sub RollItemPrefix(ByVal Index As Long, ByVal ObjectNum As Long, ByVal pOrS As Long, ByVal MagicChance As Long)
Dim B As Long, C As Long, D As Long, Mul As Single
    If Index > 0 And Index <= MaxUsers Then
        If ObjectNum > 0 And ObjectNum <= 20 Then
            With player(Index).Inv(ObjectNum)
                If .Object > 0 Then
                    If (Object(.Object).Type >= 0 And Object(.Object).Type <= 4) Or (Object(.Object).Type = 8) Then
                            If Int(Rnd * 100) < MagicChance And pOrS And 1 Then
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
                            If Int(Rnd * 100) < MagicChance And pOrS And 2 Then
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
                                .SuffixVal = (.SuffixVal Or (D * 64))
                            End If
                        End If
                    SendSocket2 Index, Chr2(17) + Chr2(ObjectNum) + DoubleChar(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)  'New Inv Obj
                End If
            End With
        End If
    End If
End Sub

Sub ResetStats(ByVal Index As Long)
    Dim A As Long
    
    If Index > 0 And Index <= MaxUsers Then
        With player(Index)
            A = (.Level) * StatsPerLevel
            .OldStrength = Class(.Class).StartStrength
            .OldAgility = Class(.Class).StartAgility
            .OldConstitution = Class(.Class).StartConstitution
            .OldEndurance = Class(.Class).StartEndurance
            .OldWisdom = Class(.Class).StartWisdom
            .OldIntelligence = Class(.Class).StartIntelligence
            .StatPoints = A
            
            SendSocket2 Index, Chr2(59) + DoubleChar$(.StatPoints) + DoubleChar$(.SkillPoints)
            
            CalculateStats Index
        End With
    End If
End Sub

Sub ResetSkills(ByVal Index As Long)
    Dim A As Long
    
    If Index > 0 And Index <= MaxUsers Then
        With player(Index)
            For A = 1 To MAX_SKILLS
                .SkillLevel(A) = 0
            Next A
            .SkillPoints = (.Level) * SkillsPerLevel
            
            '.StatusData(SE_MANARESERVES).Data(0) = 0
            '.StatusData(SE_MANARESERVES).Data(1) = 0
            CalculateStats (Index)
            
            SendSocket2 Index, Chr2(123)
            
        End With
    End If
    

End Sub

Sub uTimer(ByVal UNUSED As Long, ByVal Script As String, ByVal Wait As Long, ByVal Parm1 As Long, ByVal Parm2 As Long, ByVal Parm3 As Long, ByVal Parm4 As Long)
    Dim ST1 As String, A As Long
    ST1 = StrConv(Script, vbUnicode)
    Wait = Wait \ 1000
    If Len(ST1) > 0 Then
        If Wait < 1 Then Wait = 1
        For A = 1 To 100
            If uTimers(A).InUse = False Then
                With uTimers(A)
                    .Script = ST1
                    .Parm(0) = Parm1
                    .Parm(1) = Parm2
                    .Parm(2) = Parm3
                    .Parm(3) = Parm4
                    .timer = Wait
                    .InUse = True
                    Exit For
                End With
            End If
        Next A
    End If
End Sub

Sub PlayZoneSound(ByVal ZoneNum As Long, ByVal SoundNum As Long)
    If ZoneNum > 0 And ZoneNum <= 255 Then
        If SoundNum >= 1 And SoundNum <= 255 Then
            SendToZone2 ZoneNum, Chr2(96) + Chr2(0) + Chr2(SoundNum)
        End If
    End If
End Sub

Sub PlayZoneMusic(ByVal ZoneNum As Long, ByVal MusicNum As Long)
    If ZoneNum > 0 And ZoneNum <= 255 Then
        If MusicNum >= 1 And MusicNum <= 255 Then
            SendToZone2 ZoneNum, Chr2(96) + Chr2(1) + Chr2(MusicNum)
        End If
    End If
End Sub

Sub ZoneMessage(ByVal ZoneNum As Long, ByVal Message As String, ByVal MsgColor As Long)
    If ZoneNum >= 1 And ZoneNum <= 5000 Then
        MsgColor = MsgColor Mod 16
        SendToZone2 ZoneNum, Chr2(56) + Chr2(MsgColor) + StrConv(Message, vbUnicode)
    End If
End Sub

Sub SetPlayerNPCNameColor(ByVal playerNum As Long, ByVal NPCNum As Long, ByVal Color As Long)
    If playerNum > 0 And playerNum <= 5000 Then
        If NPCNum > 0 And NPCNum <= 255 Then
            If Color > 0 And Color <= 255 Then
                SendSocket2 playerNum, Chr2(126) + Chr2(NPCNum) + Chr2(Color)
            End If
        End If
    End If
End Sub

Function GetNPCName(ByVal NPCNum As Long) As Long
    If NPCNum > 0 And NPCNum <= 255 Then
        GetNPCName = NewString(NPC(NPCNum).Name)
    End If
End Function

Function GetNPCSprite(ByVal NPCNum As Long) As Long
    If NPCNum > 0 And NPCNum < 255 Then
        GetNPCSprite = NPC(NPCNum).sprite
    End If
End Function

Sub CreateTileParticleEffect(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, ByVal SourceType As Long, ByVal Particle As Long, ByVal red As Long, ByVal green As Long, ByVal blue As Long, ByVal Life As Long, ByVal NumParticles As Long, ByVal Size1 As Long)
    If mapNum > 0 And mapNum <= 5000 Then
        If x >= 0 And x <= 384 Then
            If y >= 0 And y <= 384 Then
                SendToMap2 mapNum, Chr2(127) + DoubleChar$(x) + DoubleChar$(y) + Chr2(SourceType) + Chr2(Particle) + Chr2(red) + Chr2(green) + Chr2(blue) + DoubleChar$(Life) + DoubleChar$(NumParticles) + DoubleChar$(Size1)
            End If
        End If
    End If
End Sub

Sub CreatePlayerParticleEffect(ByVal Index As Long, ByVal x As Long, ByVal y As Long, ByVal SourceType As Long, ByVal Particle As Long, ByVal red As Long, ByVal green As Long, ByVal blue As Long, ByVal Life As Long, ByVal NumParticles As Long, ByVal Size1 As Long)
    If Index > 0 And Index <= 5000 Then
        If x >= 0 And x <= 384 Then
            If y >= 0 And y <= 384 Then
                SendSocket2 Index, Chr2(127) + DoubleChar$(x) + DoubleChar$(y) + Chr2(SourceType) + Chr2(Particle) + Chr2(red) + Chr2(green) + Chr2(blue) + DoubleChar$(Life) + DoubleChar$(NumParticles) + DoubleChar$(Size1)
            End If
        End If
    End If
End Sub

Sub CancelParticleEffect(ByVal Index As Long, ByVal particleX As Long, ByVal particleY As Long, ByVal cancelTime As Long)
    If particleX > 0 And particleY > 0 Then
        SendSocket2 Index, Chr2(142) + DoubleChar(particleX) + DoubleChar(particleY) + DoubleChar(cancelTime)
    End If
End Sub

Function GetPlayerVaultSize(ByVal Index As Long) As Long
    If Index > 0 And Index <= MaxUsers Then
        GetPlayerVaultSize = player(Index).NumStoragePages
    End If
End Function

Sub SetPlayerVaultSize(ByVal Index As Long, ByVal Size As Long)
    If Index > 0 And Index <= MaxUsers Then
        If Size > 0 And Size <= 10 Then
            player(Index).NumStoragePages = Size
        End If
    End If
End Sub

Function GetItemRarity(ByVal Index As Long, ByVal ObjectNum As Long, ByVal pOrS As Long) As Long
    Dim A As Long
    If Index > 0 And Index <= MaxUsers Then
        If ObjectNum > 0 And ObjectNum <= 20 Then
            If player(Index).Inv(ObjectNum).prefix > 0 Or player(Index).Inv(ObjectNum).suffix > 0 Or player(Index).Inv(ObjectNum).Affix > 0 Then
                Dim B As Long
                If pOrS = 0 Then B = player(Index).Inv(ObjectNum).prefixVal \ 64
                If pOrS = 1 Then B = player(Index).Inv(ObjectNum).SuffixVal \ 64
                If pOrS = 2 Then B = player(Index).Inv(ObjectNum).AffixVal \ 64
                Select Case B
                    Case 2
                        A = 3
                    Case 1
                        A = 2
                    Case Else
                        A = 1
                End Select
            End If
        ElseIf ObjectNum >= 21 And ObjectNum <= 25 Then
            If player(Index).Equipped(ObjectNum).prefix > 0 Or player(Index).Equipped(ObjectNum).suffix > 0 Or player(Index).Inv(ObjectNum).Affix > 0 Then
                If pOrS = 0 Then B = player(Index).Equipped(ObjectNum).prefixVal \ 64
                If pOrS = 1 Then B = player(Index).Equipped(ObjectNum).SuffixVal \ 64
                If pOrS = 2 Then B = player(Index).Equipped(ObjectNum).AffixVal \ 64
                Select Case B
                    Case 2
                        A = 3
                    Case 1
                        A = 2
                    Case Else
                        A = 1
                End Select
            End If
        End If
    End If
    GetItemRarity = A
End Function

Sub CurInvCallBack(ByVal Index As Long, ByVal Script As String)
    If Index > 0 And Index <= MaxUsers Then
        Script = StrConv(Script, vbUnicode)
        player(Index).ScriptCallback = Script
        SendSocket2 Index, Chr2(128)
    End If
End Sub


Function GetMonsterDirection(ByVal mapNum As Long, ByVal MonsterNum As Long) As Long
    If mapNum > 0 And mapNum <= 5000 Then
        If MonsterNum >= 0 And MonsterNum <= 9 Then
            GetMonsterDirection = map(mapNum).monster(MonsterNum).D
        End If
    End If
End Function

Function GetMonsterAttackSpeed(ByVal mapNum As Long, ByVal MonsterNum As Long) As Long
    If mapNum > 0 And mapNum <= 5000 Then
        If MonsterNum >= 0 And MonsterNum <= 9 Then
            GetMonsterAttackSpeed = map(mapNum).monster(MonsterNum).AttackSpeed
        End If
    End If
End Function

Function GetMonsterMoveSpeed(ByVal mapNum As Long, ByVal MonsterNum As Long) As Long
    If mapNum > 0 And mapNum <= 5000 Then
        If MonsterNum >= 0 And MonsterNum <= 9 Then
            GetMonsterMoveSpeed = map(mapNum).monster(MonsterNum).MoveSpeed
        End If
    End If
End Function

Sub SetMonsterAttackSpeed(ByVal mapNum As Long, ByVal MonsterNum As Long, ByVal speed As Long)
    If mapNum > 0 And mapNum <= 5000 Then
        If MonsterNum >= 0 And MonsterNum <= 9 Then
            map(mapNum).monster(MonsterNum).AttackSpeed = speed
        End If
    End If
End Sub

Sub SetMonsterMoveSpeed(ByVal mapNum As Long, ByVal MonsterNum As Long, ByVal speed As Long)
    If mapNum > 0 And mapNum <= 5000 Then
        If MonsterNum >= 0 And MonsterNum <= 9 Then
            map(mapNum).monster(MonsterNum).MoveSpeed = speed
        End If
    End If
End Sub

Sub DamagePlayer(ByVal playerIndex As Long, ByVal damage As Long, ByVal damageType As Long, ByVal Damager As String)
    Dim A As Long, B As Boolean
    With player(playerIndex)
        If playerIndex > 0 Then
            If damageType = PT_MAGIC Then damage = PlayerMagicArmor(playerIndex, damage)
            If damageType = PT_PHYSICAL Then damage = PlayerArmor(playerIndex, damage)
        
            SendToMap2 .map, Chr2(100) + Chr2(.x * 16 + .y) + Chr2(12) + StrConv(damage, vbUnicode)
            If damage >= .HP Then
                'Player Died
                SendSocket2 playerIndex, Chr2(54) + StrConv(Damager, vbUnicode)
                SendAllBut2 playerIndex, Chr2(55) + Chr2(playerIndex) + StrConv(Damager, vbUnicode)
                CreateFloatingEvent .map, .x, .y, FT_ENDED
                PlayerDied playerIndex, False, A, True
            Else
                .HP = .HP - damage
                SendSocket2 playerIndex, Chr2(46) + DoubleChar(CInt(.HP))
            End If
        
        End If
    End With
    
End Sub

Function scriptIsVacant(ByVal map As Long, ByVal x As Long, ByVal y As Long, ByVal FromDir As Long) As Long
scriptIsVacant = IsVacant(map, CByte(x), CByte(y), CByte(FromDir))
End Function

Function PlaceInventoryObject(ByVal playerNum As Long, ByVal invSlot As Long, ByVal mapNum As Long, ByVal x As Long, ByVal y As Long) As Long
    Dim A As Long, B As Long, C As Long

    If playerNum > 0 And playerNum <= MaxUsers Then
        If mapNum > 0 And mapNum <= 5000 Then
            If x > 0 And x <= 11 Then
                If y > 0 And y <= 11 Then
                    If invSlot > 0 And invSlot <= 20 Then
                        For A = 0 To 49
                            If map(mapNum).Object(A).Object = 0 Then
                                With player(playerNum).Inv(invSlot)
                                    map(mapNum).Object(A).Object = .Object
                                    map(mapNum).Object(A).prefix = .prefix
                                    map(mapNum).Object(A).prefixVal = .prefixVal
                                    map(mapNum).Object(A).suffix = .suffix
                                    map(mapNum).Object(A).SuffixVal = .SuffixVal
                                    map(mapNum).Object(A).Affix = .Affix
                                    map(mapNum).Object(A).AffixVal = .AffixVal
                                    map(mapNum).Object(A).Value = .Value
                                    map(mapNum).Object(A).ObjectColor = .ObjectColor
                                    map(mapNum).Object(A).x = x
                                    map(mapNum).Object(A).y = y
                                    map(mapNum).Object(A).TimeStamp = 0
                                    map(mapNum).Object(A).Flags(0) = .Flags(0)
                                    map(mapNum).Object(A).Flags(1) = .Flags(1)
                                    map(mapNum).Object(A).Flags(2) = .Flags(2)
                                    map(mapNum).Object(A).Flags(3) = .Flags(3)
                                    If .prefix > 0 And .prefix < 256 Then
                                        B = prefix(.prefix).Light.Intensity
                                        C = prefix(.prefix).Light.Radius
                                    End If
                                    If .suffix > 0 And .suffix < 256 Then
                                        B = B + prefix(.suffix).Light.Intensity
                                        C = C + prefix(.suffix).Light.Radius
                                    End If
                                    If .Affix > 0 And .Affix < 256 Then
                                        B = B + prefix(.Affix).Light.Intensity
                                        C = C + prefix(.Affix).Light.Radius
                                    End If
                                    If B > 255 Then B = 255
                                    If C > 255 Then C = 255
                                    SendToMap2 mapNum, Chr2(14) + Chr2(A) + DoubleChar(.Object) + Chr2(x) + Chr2(y) + Chr2(B) + Chr2(C) + QuadChar$(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(0) + Chr2(0) + Chr2(0) + Chr2(0) + Chr2(0) + Chr2(.ObjectColor)
                                    .Object = 0
                                    .Value = 0
                                    SendSocket2 playerNum, Chr2(18) + Chr2(invSlot)
                                    PlaceInventoryObject = 1
                                    Exit Function
                                End With
                            End If
                        Next A
                    ElseIf invSlot >= 21 And invSlot <= 25 Then
                        For A = 0 To 49
                            If map(mapNum).Object(A).Object = 0 Then
                                With player(playerNum).Equipped(invSlot - 20)
                                    map(mapNum).Object(A).Object = .Object
                                    map(mapNum).Object(A).prefix = .prefix
                                    map(mapNum).Object(A).prefixVal = .prefixVal
                                    map(mapNum).Object(A).suffix = .suffix
                                    map(mapNum).Object(A).SuffixVal = .SuffixVal
                                    map(mapNum).Object(A).Affix = .Affix
                                    map(mapNum).Object(A).AffixVal = .AffixVal
                                    map(mapNum).Object(A).ObjectColor = .ObjectColor
                                    map(mapNum).Object(A).Value = .Value
                                    map(mapNum).Object(A).x = x
                                    map(mapNum).Object(A).y = y
                                    map(mapNum).Object(A).TimeStamp = 0
                                    map(mapNum).Object(A).Flags(0) = .Flags(0)
                                    map(mapNum).Object(A).Flags(1) = .Flags(1)
                                    map(mapNum).Object(A).Flags(2) = .Flags(2)
                                    map(mapNum).Object(A).Flags(3) = .Flags(3)
                                    If .prefix > 0 And .prefix < 256 Then
                                        B = prefix(.prefix).Light.Intensity
                                        C = prefix(.prefix).Light.Radius
                                    End If
                                    If .suffix > 0 And .suffix < 256 Then
                                        B = B + prefix(.suffix).Light.Intensity
                                        C = C + prefix(.suffix).Light.Radius
                                    End If
                                    If .Affix > 0 And .Affix < 256 Then
                                        B = B + prefix(.Affix).Light.Intensity
                                        C = C + prefix(.Affix).Light.Radius
                                    End If
                                    If B > 255 Then B = 255
                                    If C > 255 Then C = 255
                                    SendToMap2 mapNum, Chr2(14) + Chr2(A) + DoubleChar(.Object) + Chr2(x) + Chr2(y) + Chr2(B) + Chr2(C) + QuadChar$(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(0) + Chr2(0) + Chr2(0) + Chr2(0) + Chr2(0) + Chr2(.ObjectColor)
                                    
                                    .Object = 0
                                    .Value = 0
                                    SendSocket2 playerNum, Chr2(18) + Chr2(invSlot)
                                    PlaceInventoryObject = 1
                                    Exit Function
                                End With
                            End If
                        Next A
                    End If
                End If
            End If
        End If
    End If
End Function

Function PlaceMapObject(ByVal playerNum As Long, ByVal mapNum As Long, ByVal ObjectNum As Long) As Long
    Dim A As Long
    If playerNum > 0 And playerNum <= MaxUsers Then
        If mapNum > 0 And mapNum <= 5000 Then
            If ObjectNum >= 0 And ObjectNum <= 49 Then
                For A = 1 To 20
                    With player(playerNum).Inv(A)
                        If player(playerNum).Inv(A).Object = 0 Then
                            .Object = map(mapNum).Object(ObjectNum).Object
                            .prefix = map(mapNum).Object(ObjectNum).prefix
                            .prefixVal = map(mapNum).Object(ObjectNum).prefixVal
                            .suffix = map(mapNum).Object(ObjectNum).suffix
                            .SuffixVal = map(mapNum).Object(ObjectNum).SuffixVal
                            .Affix = map(mapNum).Object(ObjectNum).Affix
                            .AffixVal = map(mapNum).Object(ObjectNum).AffixVal
                            .ObjectColor = map(mapNum).Object(ObjectNum).ObjectColor
                            .Flags(0) = map(mapNum).Object(ObjectNum).Flags(0)
                            .Flags(1) = map(mapNum).Object(ObjectNum).Flags(1)
                            .Flags(2) = map(mapNum).Object(ObjectNum).Flags(2)
                            .Flags(3) = map(mapNum).Object(ObjectNum).Flags(3)
                            .Value = map(mapNum).Object(ObjectNum).Value
                            
                            SendToMap2 mapNum, Chr2(15) + Chr2(ObjectNum) 'Erase Map Obj
                            SendSocket2 playerNum, Chr2(17) + Chr2(A) + DoubleChar(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor) 'New Inv Obj
                            
                            map(mapNum).Object(ObjectNum).Object = 0
                            map(mapNum).Object(ObjectNum).Value = 0
                            
                            PlaceMapObject = 1
                        End If
                    End With
                Next A
            End If
        End If
    End If
End Function

Function GetMapObjectRarity(ByVal mapNum As Long, ByVal ObjectNum As Long, ByVal pOrS As Long) As Long
    If mapNum > 0 And mapNum <= 5000 Then
        If ObjectNum >= 0 And ObjectNum < 50 Then
            With map(mapNum).Object(ObjectNum)
                If .Object > 0 Then
                    Dim B As Long
                    If .prefix > 0 Or .suffix > 0 Or .Affix > 0 Then
                        If pOrS = 0 Then B = .prefixVal \ 64
                        If pOrS = 1 Then B = .SuffixVal \ 64
                        If pOrS = 2 Then B = .AffixVal \ 64
                        Select Case B
                            Case 2
                                GetMapObjectRarity = 3
                            Case 1
                                GetMapObjectRarity = 2
                            Case Else
                                GetMapObjectRarity = 1
                        End Select
                    Else
                        GetMapObjectRarity = 0
                    End If
                End If
            End With
        End If
    End If
End Function

Sub SetMapProperty(ByVal MapProperty As String, ByVal Data1 As String, ByVal Data2 As String, ByVal Data3 As String)
    MapProperty = StrConv(MapProperty, vbUnicode)
    Data1 = StrConv(Data1, vbUnicode)
    Data2 = StrConv(Data2, vbUnicode)
    Data3 = StrConv(Data3, vbUnicode)
    If CurEditMap.Num > 0 Then
        With CurEditMap
            Select Case UCase$(MapProperty)
                Case "MONSTER1"
                    .MonsterSpawn(0).monster = Val(Data1)
                    .MonsterSpawn(0).Rate = Val(Data2)
                Case "MONSTER2"
                    .MonsterSpawn(1).monster = Val(Data1)
                    .MonsterSpawn(1).Rate = Val(Data2)
                Case "MONSTER3"
                    .MonsterSpawn(2).monster = Val(Data1)
                    .MonsterSpawn(2).Rate = Val(Data2)
                Case "MONSTER4"
                    .MonsterSpawn(3).monster = Val(Data1)
                    .MonsterSpawn(3).Rate = Val(Data2)
                Case "MONSTER5"
                    .MonsterSpawn(4).monster = Val(Data1)
                    .MonsterSpawn(4).Rate = Val(Data2)
                Case "NAME"
                    .Name = Data1
                Case "MIDI"
                    .MIDI = Val(Data1)
                Case "NPC"
                    .NPC = Val(Data1)
                Case "DARKNESS"
                    .Intensity = Val(Data1)
                Case "FOG"
                    .Fog = Val(Data1)
                Case "FLAGS"
                    .Flags(0) = (Val(Data1) And 255)
                    .Flags(1) = (Val(Data1) \ 256)
                Case "BOOTLOCATION"
                    If Val(Data1) > 0 And Val(Data2) > 0 And Val(Data3) > 0 Then
                        .BootLocation.map = Val(Data1)
                        .BootLocation.x = Val(Data2)
                        .BootLocation.y = Val(Data3)
                    End If
                Case "ZONE"
                    .Zone = Val(Data1)
                Case "RAINCOLOR"
                    .RainColor = Val(Data1)
                Case "SNOWCOLOR"
                    .SnowColor = Val(Data1)
                Case "RAININTENSITY"
                    .Raining = Val(Data1)
                Case "SNOWINTENSITY"
                    .Snowing = Val(Data1)
                Case "EXITUP"
                    .ExitUp = Val(Data1)
                Case "EXITDOWN"
                    .ExitDown = Val(Data1)
                Case "EXITLEFT"
                    .ExitLeft = Val(Data1)
                Case "EXITRIGHT"
                    .ExitRight = Val(Data1)
            End Select
        End With
    End If
End Sub

Function GetMapProperty(ByVal MapProperty As String, ByVal Data1 As String) As Long
    MapProperty = StrConv(MapProperty, vbUnicode)
    Data1 = StrConv(Data1, vbUnicode)
    If CurEditMap.Num > 0 Then
        With CurEditMap
            Select Case UCase$(MapProperty)
                Case "MONSTER1"
                    Select Case UCase$(Data1)
                        Case "MONSTER"
                            GetMapProperty = .MonsterSpawn(0).monster
                        Case "RATE"
                            GetMapProperty = .MonsterSpawn(0).Rate
                    End Select
                Case "MONSTER2"
                    Select Case UCase$(Data1)
                        Case "MONSTER"
                            GetMapProperty = .MonsterSpawn(1).monster
                        Case "RATE"
                            GetMapProperty = .MonsterSpawn(1).Rate
                    End Select
                Case "MONSTER3"
                    Select Case UCase$(Data1)
                        Case "MONSTER"
                            GetMapProperty = .MonsterSpawn(2).monster
                        Case "RATE"
                            GetMapProperty = .MonsterSpawn(2).Rate
                    End Select
                Case "MONSTER4"
                    Select Case UCase$(Data1)
                        Case "MONSTER"
                            GetMapProperty = .MonsterSpawn(3).monster
                        Case "RATE"
                            GetMapProperty = .MonsterSpawn(3).Rate
                    End Select
                Case "MONSTER5"
                    Select Case UCase$(Data1)
                        Case "MONSTER"
                            GetMapProperty = .MonsterSpawn(4).monster
                        Case "RATE"
                            GetMapProperty = .MonsterSpawn(4).Rate
                    End Select
                Case "NAME"
                    GetMapProperty = NewString(.Name)
                Case "MIDI"
                    GetMapProperty = .MIDI
                Case "NPC"
                    GetMapProperty = .NPC
                Case "DARKNESS"
                    GetMapProperty = .Intensity
                Case "FOG"
                    GetMapProperty = .Fog
                Case "FLAGS"
                    GetMapProperty = .Flags(1) * 256 + .Flags(0)
                    
                Case "BOOTLOCATION"
                    Select Case UCase$(Data1)
                        Case "MAP"
                            GetMapProperty = .BootLocation.map
                        Case "X"
                            GetMapProperty = .BootLocation.x
                        Case "Y"
                            GetMapProperty = .BootLocation.y
                    End Select
                Case "ZONE"
                    GetMapProperty = .Zone
                Case "RAINCOLOR"
                    GetMapProperty = .RainColor
                Case "SNOWCOLOR"
                    GetMapProperty = .SnowColor
                Case "RAININTENSITY"
                    GetMapProperty = CLng(.Raining)
                Case "SNOWINTENSITY"
                    GetMapProperty = CLng(.Snowing)
                Case "EXITUP"
                    GetMapProperty = CLng(.ExitUp)
                Case "EXITDOWN"
                    GetMapProperty = CLng(.ExitDown)
                Case "EXITLEFT"
                    GetMapProperty = CLng(.ExitLeft)
                Case "EXITRIGHT"
                    GetMapProperty = CLng(.ExitRight)
            End Select
        End With
    End If
End Function
Function GetGuildProperty(ByVal guildn As Long, ByVal GuildProperty As String, ByVal Data1 As String) As Long
    GuildProperty = StrConv(GuildProperty, vbUnicode)
    Data1 = StrConv(Data1, vbUnicode)
    If guildn > 0 Then
        With Guild(guildn)
            Select Case UCase$(GuildProperty)
                Case "HALL"
                    Select Case UCase$(Data1)
                        Case ""
                            GetGuildProperty = .Hall
                        Case "NAME"
                            GetGuildProperty = NewString(Hall(.Hall).Name)
                        Case "UPKEEP"
                            GetGuildProperty = Hall(.Hall).Upkeep
                        Case "PRICE"
                            GetGuildProperty = Hall(.Hall).Price
                        Case "STARTX"
                            GetGuildProperty = Hall(.Hall).StartLocation.x
                        Case "STARTY"
                            GetGuildProperty = Hall(.Hall).StartLocation.y
                        Case "STARTMAP"
                            GetGuildProperty = Hall(.Hall).StartLocation.map
                    End Select
                Case "NAME"
                    GetGuildProperty = NewString(.Name)
                Case "SPRITE"
                    GetGuildProperty = .sprite
                Case "SYMBOL1"
                    GetGuildProperty = .Symbol1
                Case "SYMBOL2"
                    GetGuildProperty = .Symbol2
                Case "SYMBOL3"
                    GetGuildProperty = .Symbol3
                Case "BALANCE"
                    GetGuildProperty = .Bank
                Case "FOUNDED"
                    GetGuildProperty = .founded
                Case "MOTD"
                    GetGuildProperty = NewString(.MOTD)
                        
            End Select
        End With
    End If
End Function

Function GetServerTime(ByVal TimeType As String) As Long
    TimeType = StrConv(TimeType, vbUnicode)

            Select Case UCase$(TimeType)
                Case "HOUR", "HOURS", "H"
                    GetServerTime = Hour(Time)
                Case "MINUTE", "M", "MIN", "MINUTES"
                    GetServerTime = Minute(Time)
                Case "SECOND", "S", "SEC", "SECONDS"
                    GetServerTime = Second(Time)
                Case "DAY", "DAYS", "D", "DD"
                    GetServerTime = Day(Date)
                Case "MONTH", "MONTHS", "MON", "MM"
                    GetServerTime = Month(Date)
                Case "YEAR", "YEARS", "Y", "YY"
                    GetServerTime = Year(Date)
            End Select

    
End Function


Function GetMapName(ByVal mapNum As Long) As Long
    If mapNum > 0 And mapNum <= 5000 Then
        GetMapName = NewString(map(mapNum).Name)
    End If
End Function

Function GetMapZone(ByVal mapNum As Long) As Long
    If mapNum > 0 And mapNum <= 5000 Then
        GetMapZone = map(mapNum).Zone
    End If
End Function

Function SubString(ByVal Message As String, ByVal Start As Long, ByVal Length As Long) As Long
    If Start = 0 Then Start = 1
    If (Start >= 1) Then
        If (Start + Length) >= Len(Message) Then Length = Len(Message) - Start
        Message = StrConv(Message, vbUnicode)
        SubString = NewString(Mid(Message, Start, Length))
    End If
End Function
Function Length(ByVal Message As String) As Long
    Message = StrConv(Message, vbUnicode)
    Length = Len(Message)
End Function
Function InString(ByVal Message As String, ByVal FindVal As String, ByVal Start As Long) As Long

    Message = StrConv(Message, vbUnicode)
    FindVal = StrConv(FindVal, vbUnicode)
    If Start = 0 Then Start = 1
    If Start <= Len(Message) Then
        InString = InStr(Start, Message, FindVal, 1)
    Else
        InString = 0
    End If
    
End Function
Function ReplaceString(ByVal Message As String, ByVal FindVal As String, ByVal ReplaceVal As String) As Long
    Message = StrConv(Message, vbUnicode)
    FindVal = StrConv(FindVal, vbUnicode)
    ReplaceVal = StrConv(ReplaceVal, vbUnicode)
    ReplaceString = NewString(Replace$(Message, FindVal, ReplaceVal, 1, -1, 1))
End Function

Sub PoisonMapMonster(ByVal mapNum As Long, ByVal monsterIndex As Long, ByVal PoisonStr As Long, ByVal PoisonDur As Long)
    If mapNum > 0 And mapNum <= 5000 Then
        If monsterIndex >= 0 And monsterIndex <= 9 Then
            With map(mapNum).monster(monsterIndex)
                If .monster > 0 Then
                    If PoisonStr >= 155 Then PoisonStr = 155
                    If PoisonStr <= -100 Then PoisonStr = -100
                    PoisonMonster mapNum, monsterIndex, PoisonStr + 100, PoisonDur * 4
                End If
            End With
        End If
    End If
End Sub
Function GetPoisonStrength(ByVal mapNum As Long, ByVal monsterIndex As Long) As Long
    If mapNum > 0 And mapNum <= 5000 Then
        If monsterIndex >= 0 And monsterIndex <= 9 Then
            With map(mapNum).monster(monsterIndex)
                If .monster > 0 Then
                    GetPoisonStrength = .Poison - 100
                End If
            End With
        End If
    End If
End Function
Function GetPoisonLength(ByVal mapNum As Long, ByVal monsterIndex As Long) As Long
    If mapNum > 0 And mapNum <= 5000 Then
        If monsterIndex >= 0 And monsterIndex <= 9 Then
            With map(mapNum).monster(monsterIndex)
                If .monster > 0 Then
                    GetPoisonLength = .PoisonLength / 4
                End If
            End With
        End If
    End If
    
End Function
Sub SetMapMonsterTint(ByVal mapNum As Long, ByVal monsterIndex As Long, ByVal red As Long, ByVal green As Long, ByVal blue As Long, ByVal alpha As Long)
    If mapNum > 0 And mapNum <= 5000 Then
        If monsterIndex >= 0 And monsterIndex <= 9 Then
            With map(mapNum).monster(monsterIndex)
                If .monster > 0 Then
                    If monster(.monster).alpha - alpha < 0 Then alpha = monster(.monster).alpha
                    If monster(.monster).red - red < 0 Then red = monster(.monster).red
                    If monster(.monster).green - green < 0 Then green = monster(.monster).green
                    If monster(.monster).blue - blue < 0 Then blue = monster(.monster).blue
                    
                    
                    .R = red
                    .G = green
                    .B = blue
                    .A = alpha
                    
                    
                    SendToMap2 mapNum, Chr2(130) + Chr2(monsterIndex) + Chr2(red) + Chr2(green) + Chr2(blue) + Chr2(alpha)
                End If
            End With
        End If
     End If
End Sub
Sub SetPlayerTint(ByVal playerIndex As Long, ByVal red As Long, ByVal green As Long, ByVal blue As Long, ByVal alpha As Long)
        If playerIndex >= 0 And playerIndex <= 81 Then
            With player(playerIndex)
                    
                    If red > 255 Then red = 255
                    If green > 255 Then green = 255
                    If blue > 255 Then blue = 255
                    If alpha > 255 Then alpha = 255
                    If red < 0 Then red = 0
                    If green < 0 Then green = 0
                    If blue < 0 Then blue = 0
                    If alpha < 0 Then alpha = 0
                    
                    .red = red
                    .green = green
                    .blue = blue
                    .alpha = alpha
                    
                    
                    SendToMap .map, Chr2(133) + Chr2(playerIndex) + Chr2(red) + Chr2(green) + Chr2(blue) + Chr2(alpha)

            End With
        End If
End Sub

Sub ScriptSizedFloatingText(ByVal mapNum As Long, ByVal x As Long, ByVal y As Long, ByVal Text As String, ByVal Color As Long, ByVal MultTenths As Long, ByVal Life As Long)
    If mapNum > 0 And mapNum <= 5000 Then
        If x < 11 And y < 11 Then
            If Len(Text) > 0 Then
                If Color >= 0 And Color <= 15 Then
                     CreateSizedFloatingText mapNum, x, y, Color, StrConv(Text, vbUnicode), MultTenths, Life
                End If
            End If
        End If
    End If
End Sub

Sub BuffPlayer(ByVal playerIndex As Long, ByVal stat As Long, ByVal statMod As Long, ByVal timerVal As Long)
    Dim A As Long, B As Long, C As Long, D As Long
    If playerIndex > 0 Then
        Select Case stat
            Case 0: A = SE_HPMOD
            Case 1: A = SE_ENERGYMOD
            Case 2: A = SE_MANAMOD
            Case 3: A = SE_STRENGTHMOD
            Case 4: A = SE_AGILITYMOD
            Case 5: A = SE_ENDURANCEMOD
            Case 6: A = SE_WISDOMMOD
            Case 7: A = SE_INTELLIGENCEMOD
            Case 8: A = SE_HPREGENMOD
            Case 9: A = SE_MPREGENMOD
            Case 10: A = SE_CONSTITUTIONMOD
            Case 11: A = SE_MAGICRESISTMOD
            Case 12: A = SE_ATTACKSPEEDMOD
            Case 13: A = SE_CRITICALCHANCEMOD
            Case 14: A = SE_POISONRESISTMOD
        End Select
        
        If statMod > 32765 Then statMod = 32765
        If statMod < -32765 Then statMod = -32765
        
        B = Abs(statMod) / 256
        C = Abs(statMod) Mod 256
        D = IIf(statMod <> Abs(statMod), 1, 0)
        
        If timerVal > 0 Then
            player(playerIndex).StatusData(A).timer = timerVal
            If A <= 32 Then SetStatusEffect playerIndex, A
            player(playerIndex).StatusData(A).Data(0) = B
            player(playerIndex).StatusData(A).Data(1) = C
            player(playerIndex).StatusData(A).Data(2) = D
            
            player(playerIndex).StatusData(A).timer = timerVal
        
            CalculateStats (playerIndex)
            
        End If
    End If
End Sub
Sub ShowObjectInformation(ByVal playerIndex As Long, ByVal ObjNum As Long, ByVal objVal As Long)
    SendSocket2 playerIndex, Chr2(141) + DoubleChar(ObjNum) + QuadChar(objVal)
End Sub


Sub ScriptPrintDebug(ByVal Message As String)
    PrintDebug StrConv(Message, vbUnicode)
End Sub

Sub ScriptPrintReport(ByVal Message As String, ByVal fileName As String)
    Open StrConv(fileName, vbUnicode) & ".log" For Append As #1
    Print #1, StrConv(Message, vbUnicode)
    Close #1
End Sub

Sub createMapProjectile(ByVal mapIndex As Long, ByVal x As Long, ByVal y As Long, ByVal direction As Long, ByVal sprite As Long, ByVal speed As Long, ByVal damage As Long, ByVal damageType As Long, ByVal Message As String)
    Dim A As Long
    With map(mapIndex)
        For A = 1 To 30
            With .Projectile(A)
                If .sprite = 0 Then
                    .startTime = GetTickCount
                    .startX = x
                    .startY = y
                    .direction = direction
                    .sprite = sprite
                    .speed = speed
                    .magical = damageType
                    .damage = damage
                    .damageString = Message
                    SendToMap2 mapIndex, Chr2(125) + Chr2(255) + DoubleChar$(sprite) + Chr2(x * 16 + y) + Chr2(direction) + Chr2(speed) + Chr2(0) + Chr2(A) + Chr2(0) + Chr2(0) + Chr2(0)
                    'SendToMap mapIndex, Chr2(125) + Chr2(255) + DoubleChar$(sprite) + Chr2(x * 16 + y) + Chr2(direction) + Chr2(speed) + Chr2(Radius) + Chr2(Intensity) + Chr2(red) + Chr2(green) + Chr2(blue)
                    Exit Sub
                End If
            End With
        Next A
        
        Dim B As Long, C As Long
        B = GetTickCount
        For A = 1 To 30
            With .Projectile(A)
                If .startTime < B Then
                    B = .startTime
                    C = A
                End If
            End With
        Next A
        
        With .Projectile(C)
            .startTime = GetTickCount
            .startX = x
            .startY = y
            .direction = direction
            .sprite = sprite
            .speed = speed
            .magical = damageType
            .damage = damage
            .damageString = Message
            SendToMap2 mapIndex, Chr2(125) + Chr2(255) + DoubleChar$(sprite) + Chr2(x * 16 + y) + Chr2(direction) + Chr2(speed) + Chr2(0) + Chr2(C) + Chr2(1) + Chr2(0) + Chr2(0)
            Exit Sub
        End With
        
    End With
End Sub

Sub ScriptWriteIniString(ByVal file As String, ByVal header As String, ByVal Key As String, ByVal Value As String)
    WriteString StrConv(file, vbUnicode), StrConv(header, vbUnicode), StrConv(Key, vbUnicode), StrConv(Value, vbUnicode)
End Sub
Function ScriptReadiniString(ByVal file As String, ByVal header As String, ByVal Key As String) As Long
    ScriptReadiniString = NewString(ReadString(StrConv(file, vbUnicode), StrConv(header, vbUnicode), StrConv(Key, vbUnicode)))
End Function
Function ScriptReadiniInt(ByVal file As String, ByVal header As String, ByVal Key As String) As Long
    ScriptReadiniInt = ReadInt(StrConv(file, vbUnicode), StrConv(header, vbUnicode), StrConv(Key, vbUnicode))
End Function

Sub NoSuchFunction()
    'Yea... This will run if they do something really dumb to their mbsc.inc which is the most likely cause for a script.dll crash
End Sub
Sub InitFunctionTable()
    Dim A As Long
    FunctionTable(0) = GetValue(AddressOf DeleteString)
    FunctionTable(1) = GetValue(AddressOf StrCat)
    FunctionTable(2) = GetValue(AddressOf StrCmp)
    FunctionTable(3) = GetValue(AddressOf StrFormat)
    FunctionTable(4) = GetValue(AddressOf Random)
    FunctionTable(5) = GetValue(AddressOf GetPlayerAccess)
    FunctionTable(6) = GetValue(AddressOf GetPlayerMap)
    FunctionTable(7) = GetValue(AddressOf GetPlayerX)
    FunctionTable(8) = GetValue(AddressOf GetPlayerY)
    FunctionTable(9) = GetValue(AddressOf GetPlayerSprite)
    FunctionTable(10) = GetValue(AddressOf GetPlayerClass)
    FunctionTable(11) = GetValue(AddressOf GetPlayerGender)
    FunctionTable(12) = GetValue(AddressOf GetPlayerHP)
    FunctionTable(13) = GetValue(AddressOf GetPlayerEnergy)
    FunctionTable(14) = GetValue(AddressOf GetPlayerMana)
    FunctionTable(15) = GetValue(AddressOf GetPlayerMaxHP)
    FunctionTable(16) = GetValue(AddressOf GetPlayerMaxEnergy)
    FunctionTable(17) = GetValue(AddressOf GetPlayerMaxMana)
    FunctionTable(18) = GetValue(AddressOf GetPlayerStrength)
    FunctionTable(19) = GetValue(AddressOf GetPlayerEndurance)
    FunctionTable(20) = GetValue(AddressOf GetPlayerIntelligence)
    FunctionTable(21) = GetValue(AddressOf GetPlayerAgility)
    FunctionTable(22) = GetValue(AddressOf GetPlayerBank)
    FunctionTable(23) = GetValue(AddressOf GetPlayerExperience)
    FunctionTable(24) = GetValue(AddressOf GetPlayerLevel)
    FunctionTable(25) = GetValue(AddressOf GetPlayerStatus)
    FunctionTable(26) = GetValue(AddressOf GetPlayerGuild)
    FunctionTable(27) = GetValue(AddressOf GetPlayerInvObject)
    FunctionTable(28) = GetValue(AddressOf GetPlayerInvValue)
    FunctionTable(29) = GetValue(AddressOf GetPlayerEquipped)
    FunctionTable(30) = GetValue(AddressOf GetPlayerName)
    FunctionTable(31) = GetValue(AddressOf GetPlayerUser)
    FunctionTable(32) = GetValue(AddressOf GetMapWall)
    FunctionTable(33) = GetValue(AddressOf ScriptSetPlayerHP)
    FunctionTable(34) = GetValue(AddressOf SetPlayerEnergy)
    FunctionTable(35) = GetValue(AddressOf SetPlayerMana)
    FunctionTable(36) = GetValue(AddressOf PlayerMessage)
    FunctionTable(37) = GetValue(AddressOf PlayerWarp)
    FunctionTable(38) = GetValue(AddressOf MapMessage)
    FunctionTable(39) = GetValue(AddressOf GlobalMessage)
    FunctionTable(40) = GetValue(AddressOf GetGuildHall)
    FunctionTable(41) = GetValue(AddressOf GetGuildBank)
    FunctionTable(42) = GetValue(AddressOf GetGuildMemberCount)
    FunctionTable(43) = GetValue(AddressOf GetGuildName)
    FunctionTable(44) = GetValue(AddressOf GetMapPlayerCount)
    FunctionTable(45) = GetValue(AddressOf MapMessageAllBut)
    FunctionTable(46) = GetValue(AddressOf HasObj)
    FunctionTable(47) = GetValue(AddressOf TakeObj)
    FunctionTable(48) = GetValue(AddressOf GiveObj)
    FunctionTable(49) = GetValue(AddressOf GetTime)
    FunctionTable(50) = GetValue(AddressOf GetMaxUsers)
    FunctionTable(51) = GetValue(AddressOf RunScript0)
    FunctionTable(52) = GetValue(AddressOf RunScript1)
    FunctionTable(53) = GetValue(AddressOf RunScript2)
    FunctionTable(54) = GetValue(AddressOf RunScript3)
    FunctionTable(55) = GetValue(AddressOf OpenDoor)
    FunctionTable(56) = GetValue(AddressOf Str)
    FunctionTable(57) = GetValue(AddressOf SetPlayerSprite)
    FunctionTable(58) = GetValue(AddressOf GetAbs)
    FunctionTable(59) = GetValue(AddressOf GetSqr)
    FunctionTable(60) = GetValue(AddressOf CanAttackPlayer)
    FunctionTable(61) = GetValue(AddressOf IsPlaying)
    FunctionTable(63) = GetValue(AddressOf ScriptAttackMonster)
    FunctionTable(64) = GetValue(AddressOf CanAttackMonster)
    FunctionTable(65) = GetValue(AddressOf GetMonsterType)
    FunctionTable(66) = GetValue(AddressOf GetMonsterX)
    FunctionTable(67) = GetValue(AddressOf GetMonsterY)
    FunctionTable(68) = GetValue(AddressOf GetMonsterTarget)
    FunctionTable(69) = GetValue(AddressOf SetMonsterTarget)
    FunctionTable(70) = GetValue(AddressOf GetInStr)
    FunctionTable(71) = GetValue(AddressOf SpawnObject)
    FunctionTable(74) = GetValue(AddressOf GetGuildSprite)
    FunctionTable(75) = GetValue(AddressOf ScriptTimer)
    FunctionTable(76) = GetValue(AddressOf SetPlayerGuild)
    FunctionTable(77) = GetValue(AddressOf GetFlag)
    FunctionTable(78) = GetValue(AddressOf SetFlag)
    FunctionTable(79) = GetValue(AddressOf GetPlayerFlag)
    FunctionTable(80) = GetValue(AddressOf SetPlayerFlag)
    FunctionTable(81) = GetValue(AddressOf ResetPlayerFlag)
    FunctionTable(83) = GetValue(AddressOf GetObjX)
    FunctionTable(84) = GetValue(AddressOf GetObjY)
    FunctionTable(85) = GetValue(AddressOf GetObjNum)
    FunctionTable(86) = GetValue(AddressOf GetObjVal)
    FunctionTable(87) = GetValue(AddressOf DestroyObj)
    FunctionTable(88) = GetValue(AddressOf Boot_Player)
    FunctionTable(89) = GetValue(AddressOf Ban_Player)
    FunctionTable(90) = GetValue(AddressOf SetPlayerName)
    FunctionTable(91) = GetValue(AddressOf SetPlayerBank)
    FunctionTable(92) = GetValue(AddressOf SetGuildBank)
    FunctionTable(93) = GetValue(AddressOf Find_Player)
    FunctionTable(94) = GetValue(AddressOf StrVal)
    FunctionTable(95) = GetValue(AddressOf GetTileAtt)
    FunctionTable(96) = GetValue(AddressOf GetPlayerIP)
    FunctionTable(97) = GetValue(AddressOf RunScript4)
    FunctionTable(98) = GetValue(AddressOf SpawnMonster)
    FunctionTable(99) = GetValue(AddressOf SetPlayerStatus)
    FunctionTable(100) = GetValue(AddressOf GivePlayerExp)
    FunctionTable(101) = GetValue(AddressOf GetMapAnimBinary)
    FunctionTable(102) = GetValue(AddressOf SetAnimBinary)
    FunctionTable(103) = GetValue(AddressOf GetObjectName)
    FunctionTable(104) = GetValue(AddressOf GetObjectData)
    FunctionTable(105) = GetValue(AddressOf GetObjectType)
    FunctionTable(106) = GetValue(AddressOf GetObjectDur)
    FunctionTable(107) = GetValue(AddressOf SetInvObjectVal)
    FunctionTable(108) = GetValue(AddressOf PlayCustomWav)
    FunctionTable(109) = GetValue(AddressOf GetPlayerArmor)
    FunctionTable(110) = GetValue(AddressOf CreateTileEffect)
    FunctionTable(112) = GetValue(AddressOf CreateMonsterEffect)
    FunctionTable(113) = GetValue(AddressOf CreatePlayerEffect)
    FunctionTable(114) = GetValue(AddressOf sSetBit)
    FunctionTable(115) = GetValue(AddressOf sExamineBit)
    FunctionTable(116) = GetValue(AddressOf sClearBit)
    FunctionTable(118) = GetValue(AddressOf GetClassLevel)
    FunctionTable(119) = GetValue(AddressOf LearnSkill)
    FunctionTable(120) = GetValue(AddressOf SetStringFlag)
    FunctionTable(121) = GetValue(AddressOf GetStringFlag)
    FunctionTable(122) = GetValue(AddressOf GetMonsterName)
    FunctionTable(123) = GetValue(AddressOf GetMonsterHP)
    FunctionTable(124) = GetValue(AddressOf GetMonsterDescription)
    FunctionTable(125) = GetValue(AddressOf GetMonsterSprite)
    FunctionTable(126) = GetValue(AddressOf GetMonsterExperience)
    FunctionTable(127) = GetValue(AddressOf GetMonsterLevel)
    FunctionTable(128) = GetValue(AddressOf GetMonsterArmor)
    FunctionTable(129) = GetValue(AddressOf DestroyMonster)
    FunctionTable(130) = GetValue(AddressOf SetMapMonsterHP)
    FunctionTable(131) = GetValue(AddressOf GetMapMonsterHP)
    FunctionTable(132) = GetValue(AddressOf SetMapMonsterFlag)
    FunctionTable(133) = GetValue(AddressOf GetMapMonsterFlag)
    FunctionTable(134) = GetValue(AddressOf ScriptLoadMap)
    FunctionTable(135) = GetValue(AddressOf ScriptSaveMap)
    FunctionTable(136) = GetValue(AddressOf SetTile)
    FunctionTable(137) = GetValue(AddressOf GetTile)
    FunctionTable(138) = GetValue(AddressOf GetMapExitDirection)
    FunctionTable(139) = GetValue(AddressOf CreateStaticText)
    FunctionTable(140) = GetValue(AddressOf GetPlayerDirection)
    FunctionTable(141) = GetValue(AddressOf AddMapMonsterQueueMove)
    FunctionTable(142) = GetValue(AddressOf AddMapMonsterQueueScript)
    FunctionTable(143) = GetValue(AddressOf StartWidgetMenu)
    FunctionTable(144) = GetValue(AddressOf SendWidgetString)
    FunctionTable(145) = GetValue(AddressOf AddWidgetButton)
    FunctionTable(146) = GetValue(AddressOf AddWidgetLabel)
    FunctionTable(147) = GetValue(AddressOf AddWidgetTextBox)
    FunctionTable(148) = GetValue(AddressOf GetWidgetValueLong)
    FunctionTable(149) = GetValue(AddressOf GetWidgetValueString)
    FunctionTable(150) = GetValue(AddressOf MapReset)
    FunctionTable(151) = GetValue(AddressOf GetMapMonsterCount)
    FunctionTable(152) = GetValue(AddressOf SetPlayerFrozen)
    FunctionTable(153) = GetValue(AddressOf FadeMap)
    FunctionTable(154) = GetValue(AddressOf KillPlayer)
    FunctionTable(155) = GetValue(AddressOf WarpMonster)
    FunctionTable(156) = GetValue(AddressOf GetSkillLevel)
    FunctionTable(157) = GetValue(AddressOf SetSkillLevel)
    FunctionTable(158) = GetValue(AddressOf SetPlayerStatusEffect)
    FunctionTable(159) = GetValue(AddressOf GetPlayerConstitution)
    FunctionTable(160) = GetValue(AddressOf GetPlayerWisdom)
    FunctionTable(161) = GetValue(AddressOf SetStatMod)
    FunctionTable(162) = GetValue(AddressOf CreatePlayerProjectile)
    FunctionTable(163) = GetValue(AddressOf AddWidgetImage)
    FunctionTable(164) = GetValue(AddressOf GetPrefix)
    FunctionTable(165) = GetValue(AddressOf GetPrefixVal)
    FunctionTable(166) = GetValue(AddressOf SetPrefix)
    FunctionTable(167) = GetValue(AddressOf DisplayUnzFont)
    FunctionTable(168) = GetValue(AddressOf ScriptCreateFloatingText)
    FunctionTable(169) = GetValue(AddressOf ScriptCreatePlayerFloatingText)
    FunctionTable(170) = GetValue(AddressOf PlayMusic)
    FunctionTable(171) = GetValue(AddressOf ResetStats)
    FunctionTable(172) = GetValue(AddressOf ClearMapMonsterQueue)
    FunctionTable(173) = GetValue(AddressOf SetMapWeather)
    FunctionTable(174) = GetValue(AddressOf PlayMapWav)
    FunctionTable(175) = GetValue(AddressOf SetWeatherVariable)
    FunctionTable(176) = GetValue(AddressOf uTimer)
    FunctionTable(177) = GetValue(AddressOf GetMapTile)
    FunctionTable(178) = GetValue(AddressOf PlayZoneSound)
    FunctionTable(179) = GetValue(AddressOf PlayZoneMusic)
    FunctionTable(180) = GetValue(AddressOf ZoneMessage)
    FunctionTable(181) = GetValue(AddressOf SetZoneWeather)
    FunctionTable(182) = GetValue(AddressOf GetMapWeather)
    FunctionTable(183) = GetValue(AddressOf AddMapMonsterQueuePause)
    FunctionTable(184) = GetValue(AddressOf SetPlayerNPCNameColor)
    FunctionTable(185) = GetValue(AddressOf GetNPCName)
    FunctionTable(186) = GetValue(AddressOf GetNPCSprite)
    FunctionTable(187) = GetValue(AddressOf SetAnim)
    FunctionTable(188) = GetValue(AddressOf SetTileAtt)
    FunctionTable(189) = GetValue(AddressOf SetMapExitDirection)
    FunctionTable(190) = GetValue(AddressOf SetWall)
    FunctionTable(191) = GetValue(AddressOf CreateTileParticleEffect)
    FunctionTable(192) = GetValue(AddressOf SetPlayerRenown)
    FunctionTable(193) = GetValue(AddressOf GetPlayerRenown)
    FunctionTable(194) = GetValue(AddressOf RunScript5)
    FunctionTable(195) = GetValue(AddressOf SetPlayerSkillPoints)
    FunctionTable(196) = GetValue(AddressOf GetPlayerSkillPoints)
    FunctionTable(197) = GetValue(AddressOf RollItemPrefix)
    FunctionTable(198) = GetValue(AddressOf GetWall)
    FunctionTable(199) = GetValue(AddressOf SetPlayerVaultSize)
    FunctionTable(200) = GetValue(AddressOf GetPlayerVaultSize)
    FunctionTable(201) = GetValue(AddressOf GetItemRarity)
    FunctionTable(202) = GetValue(AddressOf CurInvCallBack)
    FunctionTable(203) = GetValue(AddressOf SetMapProperty)
    FunctionTable(204) = GetValue(AddressOf GetMonsterDirection)
    FunctionTable(205) = GetValue(AddressOf GetMonsterAttackSpeed)
    FunctionTable(206) = GetValue(AddressOf GetMonsterMoveSpeed)
    FunctionTable(207) = GetValue(AddressOf SetMonsterAttackSpeed)
    FunctionTable(208) = GetValue(AddressOf SetMonsterMoveSpeed)
    FunctionTable(209) = GetValue(AddressOf ResetSkills)
    FunctionTable(210) = GetValue(AddressOf GetMapObjectDur)
    FunctionTable(211) = GetValue(AddressOf PlaceInventoryObject)
    FunctionTable(212) = GetValue(AddressOf PlaceMapObject)
    FunctionTable(213) = GetValue(AddressOf GetMapObjectRarity)
    FunctionTable(214) = GetValue(AddressOf RunScript6)
    FunctionTable(215) = GetValue(AddressOf CreatePlayerParticleEffect)
    FunctionTable(216) = GetValue(AddressOf CreatePlayerLitProjectile)
    FunctionTable(217) = GetValue(AddressOf SpawnMonsterOnMap)
    FunctionTable(218) = GetValue(AddressOf GetMapAttData)
    FunctionTable(219) = GetValue(AddressOf GetMapName)
    FunctionTable(220) = GetValue(AddressOf ScriptPrintDebug)
    FunctionTable(223) = GetValue(AddressOf DamagePlayer)
    FunctionTable(224) = GetValue(AddressOf BuffPlayer)
    FunctionTable(229) = GetValue(AddressOf GetMapProperty)
    FunctionTable(230) = GetValue(AddressOf SubString)
    FunctionTable(231) = GetValue(AddressOf Length)
    FunctionTable(232) = GetValue(AddressOf InString)
    FunctionTable(233) = GetValue(AddressOf ReplaceString)
    FunctionTable(234) = GetValue(AddressOf PoisonMapMonster)
    FunctionTable(235) = GetValue(AddressOf GetPoisonStrength)
    FunctionTable(236) = GetValue(AddressOf GetPoisonLength)
    FunctionTable(237) = GetValue(AddressOf SetMapMonsterTint)
    FunctionTable(238) = GetValue(AddressOf ScriptSizedFloatingText)
    FunctionTable(239) = GetValue(AddressOf SetPlayerTint)
    FunctionTable(240) = GetValue(AddressOf SetMonsterTarget2)
    FunctionTable(241) = GetValue(AddressOf ScriptPrintReport)
    FunctionTable(242) = GetValue(AddressOf createMapProjectile)
    FunctionTable(243) = GetValue(AddressOf setMonsterDirection)
    FunctionTable(244) = GetValue(AddressOf AddMapMonsterQueueShift)
    FunctionTable(245) = GetValue(AddressOf scriptIsVacant)
    FunctionTable(246) = GetValue(AddressOf SpawnObject2)
    FunctionTable(247) = GetValue(AddressOf ShowObjectInformation)
    FunctionTable(248) = GetValue(AddressOf GetGuildProperty)
    FunctionTable(249) = GetValue(AddressOf GetStatBonus)
    FunctionTable(250) = GetValue(AddressOf ScriptWriteIniString)
    FunctionTable(251) = GetValue(AddressOf ScriptReadiniString)
    FunctionTable(252) = GetValue(AddressOf ScriptReadiniInt)
    FunctionTable(253) = GetValue(AddressOf GetPlayerMagicFind)
    FunctionTable(254) = GetValue(AddressOf GetServerTime)
    FunctionTable(255) = GetValue(AddressOf SetPLayerClass)
    FunctionTable(256) = GetValue(AddressOf CancelParticleEffect)
    FunctionTable(257) = GetValue(AddressOf GetPrefixProperty)
    FunctionTable(258) = GetValue(AddressOf RunScript10)
    FunctionTable(259) = GetValue(AddressOf GetObjectFlag)
    FunctionTable(260) = GetValue(AddressOf SetObjectFlag)
    FunctionTable(261) = GetValue(AddressOf GetObjectColor)
    FunctionTable(262) = GetValue(AddressOf SetObjectColor)
    FunctionTable(263) = GetValue(AddressOf KillPLayerSounds)
    FunctionTable(264) = GetValue(AddressOf GetMapZone)
    FunctionTable(265) = GetValue(AddressOf SpawnObject3)
    FunctionTable(266) = GetValue(AddressOf CreateSizedStaticText)
   ' FunctionTable(258) = GetValue(AddressOf CancelParticleEffect)
   ' FunctionTable(259) = GetValue(AddressOf CancelParticleEffect)
For A = 1 To 300
    If FunctionTable(A) = 0 Then
        FunctionTable(A) = GetValue(AddressOf NoSuchFunction)
    End If
Next A

End Sub
