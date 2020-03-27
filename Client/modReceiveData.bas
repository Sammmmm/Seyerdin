Attribute VB_Name = "modReceiveData"
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

Public PingTime As Long

'Dim st2 As String
Sub ReceiveData()
'On Error GoTo Error_Handler
    Dim PacketLength As Long, PacketID As Integer
    Dim St As String, St1 As String
    Dim A As Long, b As Long, C As Long, D As Long, E As Long, F As Long
    SocketData = SocketData + Receive(ClientSocket)
LoopRead:
    If Len(SocketData) >= 3 Then
        PacketLength = GetInt(Mid$(SocketData, 1, 2))
        
        If PacketLength = 65535 Then
            PacketLength = GetLong(Mid$(SocketData, 3, 4))
        End If
        
        If Len(SocketData) - 2 >= PacketLength Or (PacketLength > 65536 And Len(SocketData) - 6 >= PacketLength) Then
            If (PacketLength > 65535 And Len(SocketData) - 6 >= PacketLength) Then
                SocketData = Mid$(SocketData, 5)
            End If
            St = Mid$(SocketData, 3, PacketLength)
            SocketData = Mid$(SocketData, PacketLength + 3)
            If PacketLength > 0 Then
                PacketID = Asc(Mid$(St, 1, 1))
                If Len(St) > 1 Then
                    St = Mid$(St, 2)
                Else
                    St = ""
                End If
      '         'PrintChat "Packet " & Str(PacketID), 5, 10
                Select Case PacketID
                    Case 0 'Error Logging On
                        If Len(St) >= 1 Then
                            Select Case Asc(Mid$(St, 1, 1))
                                Case 0 'Custom Message
                                    If Len(St) >= 2 Then
                                        frmMenu.SetStatusText Mid$(St, 2), 2
                                    End If
                                Case 1 'Invalid User/Pass
                                    frmMenu.SetStatusText "Invalid username/password!", 2
                                Case 2 'Account already in use
                                    frmMenu.SetStatusText "Account is already connected.", 2
                                Case 3 'Banned
                                    If Len(St) >= 5 Then
                                        A = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                                        If Len(St) > 5 Then
                                            BanString = "Banned until " + CStr(CDate(A)) + " (" + Mid$(St, 6) + ")!"
                                            'frmMenu.SetStatusText "Banned until " + CStr(CDate(A)) + " (" + Mid$(St, 6) + ")!", 2
                                        Else
                                            BanString = "Banned until " + CStr(CDate(A)) + "!"
                                            'frmMenu.SetStatusText "Banned until " + CStr(CDate(A)) + "!", 2
                                        End If
                                    End If
                                Case 4 'Server Full
                                    frmMenu.SetStatusText "Server is full, please try again later.", 2
                                Case 6 'Out of date
                                    frmMenu.SetStatusText "Version Outdated.  Please run the updater.", 2
                            End Select
                        End If
                        frmMenu.Hide
                        closesocket (ClientSocket)
                        ClientSocket = INVALID_SOCKET
                        frmMenu.Show
                        
                    Case 1 'Error Creating New Account
                        If Len(St) >= 1 Then
                            Select Case Asc(Mid$(St, 1, 1))
                                Case 0 'Custom Message
                                    If Len(St) >= 2 Then
                                    End If
                                Case 1 'User name already in use
                                    frmMenu.SetStatusText "Username already in use.", 2
                            End Select
                        End If
                        closesocket ClientSocket
                        ClientSocket = INVALID_SOCKET
                    Case 2 'Account Created
                        frmMenu.SetMenu 1 'MENU_MENU
                        frmMenu.SetStatusText "Account Created Successfully", 1
                        closesocket ClientSocket
                        ClientSocket = INVALID_SOCKET
                    Case 3 'Logged On / Character Data
                        SetGUIWindow WINDOW_INVALID
                        Character.Trading = False
                        TradeData.Tradestate(0) = 0
                        TradeData.Tradestate(1) = 0
                        TradeData.player = 0
                        For A = 1 To 10
                            TradeData.YourObjects(A).Object = 0
                            TradeData.TheirObjects(A).Object = 0
                        Next A
                        frmMenu.SetMenu 3 'Character
                        If Len(St) >= 37 Then
                            With Character
                                .Name = ""
                                .Level = Asc(Mid$(St, 1, 1))
                                .Class = Asc(Mid$(St, 2, 1))
                                .Gender = Asc(Mid$(St, 7, 1))
                                .Sprite = Asc(Mid$(St, 8, 1))
                                .HP = Asc(Mid$(St, 9, 1)) * 256 + Asc(Mid$(St, 10, 1))
                                .Energy = Asc(Mid$(St, 11, 1)) * 256 + Asc(Mid$(St, 12, 1))
                                .Mana = Asc(Mid$(St, 13, 1)) * 256 + Asc(Mid$(St, 14, 1))
                                .MaxHP = Asc(Mid$(St, 15, 1)) * 256 + Asc(Mid$(St, 16, 1))
                                .MaxEnergy = Asc(Mid$(St, 17, 1)) * 256 + Asc(Mid$(St, 18, 1))
                                .MaxMana = Asc(Mid$(St, 19, 1)) * 256 + Asc(Mid$(St, 20, 1))
                                .strength = Asc(Mid$(St, 21, 1))
                                .Agility = Asc(Mid$(St, 22, 1))
                                .Endurance = Asc(Mid$(St, 23, 1))
                                .Wisdom = Asc(Mid$(St, 24, 1))
                                .Constitution = Asc(Mid$(St, 25, 1))
                                .Intelligence = Asc(Mid$(St, 26, 1))
                                .Level = Asc(Mid$(St, 27, 1))
                                .Status = Asc(Mid$(St, 28, 1))
                                .Guild = Asc(Mid$(St, 29, 1))
                                .GuildRank = Asc(Mid$(St, 30, 1))
                                .Access = Asc(Mid$(St, 31, 1))
                                .Index = Asc(Mid$(St, 32, 1))
                                .Experience = Asc(Mid$(St, 33, 1)) * 16777216 + Asc(Mid$(St, 34, 1)) * 65536 + Asc(Mid$(St, 35, 1)) * 256& + Asc(Mid$(St, 36, 1))
                                .Squelched = Asc(Mid$(St, 37, 1))
                                .StatusEffect = Asc(Mid$(St, 38, 1)) * 16777216 + Asc(Mid$(St, 39)) * 65536 + Asc(Mid$(St, 40, 1)) * 256& + Asc(Mid$(St, 41))
                                .StatPoints = Asc(Mid$(St, 42, 1)) * 256& + Asc(Mid$(St, 43, 1))
                                TempVar5 = .StatPoints
                                Tstr = 0: TInt = 0: TAgi = 0: TEnd = 0: TWis = 0: TCon = 0
                                .skillPoints = Asc(Mid$(St, 44, 1)) * 256& + Asc(Mid$(St, 45, 1))
                                St = Mid$(St, 46)
                                For A = 1 To 255
                                    .SkillLevels(A) = Asc(Mid$(St, A, 1))
                                Next A
                                St = Mid$(St, 256)
                                For A = 0 To 254
                                    .SkillEXP(A + 1) = Asc(Mid$(St, A * 4 + 1, 1)) * 16777216 + Asc(Mid$(St, A * 4 + 2, 1)) * 65536 + Asc(Mid$(St, A * 4 + 3, 1)) * 256& + Asc(Mid$(St, A * 4 + 4, 1))
                                Next A
                                St = Mid$(St, 1021)
                                GetSections2 (St)
                                .Name = Section(1)
                                .desc = Section(2)
                                If .Guild > 0 Then
                                    Guild(.Guild).Name = Section(3)
                                End If
                                DrawTrainBars
                                DrawExperience
                                
                            End With
                            frmMenu.DrawCharacterData True
                        Else
                            Character.Class = 0
                            frmMenu.DrawCharacterData False
                        End If
                        LoadMapStringData (UCase$(User))
                        LoadSkillMacros
                        
                    Case 4 'Motd
                        
                    Case 5 'Password Changed
                        frmMenu.SetMenu 3 'Character
                        frmMenu.SetStatusText "Password changed successfully.", 1
                        
                    Case 6 'Player Joined Game
                        If Len(St) >= 7 Then
                            A = Asc(Mid$(St, 1, 1))
                            With player(A)
                                .Ignore = False
                                .Sprite = Asc(Mid$(St, 2, 1))
                                .Status = Asc(Mid$(St, 3, 1))
                                .Guild = Asc(Mid$(St, 4, 1))
                                .Light.Intensity = Asc(Mid$(St, 5, 1))
                                .Light.Radius = Asc(Mid$(St, 6, 1))
                                .Light.Type = LT_PLAYER
                                .Name = Mid$(St, 7)
                                .alpha = 0
                                .Red = 0
                                .Green = 0
                                .Blue = 0
                                .Mana = 100
                                .HP = 100
                                If CMap > 0 Then
                                    If .Status = 2 Then
                                        PrintChat "All hail " + .Name + ", a new adventurer in this land!", 3, Options.FontSize, 15
                                    Else
                                        PrintChat .Name + " has joined the game!", 3, Options.FontSize, 15
                                    End If
                                End If
                                UpdatePlayerColor A
                                ReDoLightSources
                            End With
                        End If
                        
                    Case 7 'Player Left Game
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                With player(A)
                                    PlayerLeftMap A
                                    .Sprite = 0
                                    If .Party = Character.Party Then
                                        .Party = 0
                                        DrawPartyNames
                                    End If
                                    .Light.Intensity = 0
                                    .Light.Radius = 0
                                    .Light.Type = LT_NONE
                                    ReDoLightSources
                                    If CurrentTarget.TargetType = TT_PLAYER And CurrentTarget.Target = A Then CurrentTarget.TargetType = 0
                                    PrintChat .Name + " has left the game!", 3, Options.FontSize, 15
                                End With
                            End If
                        End If
                    
                    Case 8 'Player joined map
                        If Len(St) = 16 Then
                            A = Asc(Mid$(St, 1, 1))
                             With player(A)
                                 .map = CMap
                                 .x = Asc(Mid$(St, 2, 1))
                                 .y = Asc(Mid$(St, 3, 1))
                                 .D = Asc(Mid$(St, 4, 1))
                                 .XO = .x * 32
                                 .YO = .y * 32
                                 .A = 0
                                 .Red = Asc(Mid$(St, 5, 1))
                                 .Green = Asc(Mid$(St, 6, 1))
                                 .Blue = Asc(Mid$(St, 7, 1))
                                 .alpha = Asc(Mid$(St, 8, 1))
                                 .equippedPicture(1) = Asc(Mid$(St, 9, 1)) * 256 + Asc(Mid$(St, 10, 1))
                                 .equippedPicture(2) = Asc(Mid$(St, 11, 1)) * 256 + Asc(Mid$(St, 12, 1))
                                 .equippedPicture(3) = Asc(Mid$(St, 13, 1)) * 256 + Asc(Mid$(St, 14, 1))
                                 .equippedPicture(4) = Asc(Mid$(St, 15, 1)) * 256 + Asc(Mid$(St, 16, 1))
                                 .buff = 0
                             End With
                             ReDoLightSources
                        Else
                            If Len(St) = 12 Then
                                A = Asc(Mid$(St, 1, 1))
                                With player(A)
                                    .map = CMap
                                    .x = Asc(Mid$(St, 2, 1))
                                    .y = Asc(Mid$(St, 3, 1))
                                    .D = Asc(Mid$(St, 4, 1))
                                    .XO = .x * 32
                                    .YO = .y * 32
                                    .A = 0
                                    .Red = 0
                                    .Green = 0
                                    .Blue = 0
                                    .alpha = 0
                                    .equippedPicture(1) = Asc(Mid$(St, 5, 1)) * 256 + Asc(Mid$(St, 6, 1))
                                    .equippedPicture(2) = Asc(Mid$(St, 7, 1)) * 256 + Asc(Mid$(St, 8, 1))
                                    .equippedPicture(3) = Asc(Mid$(St, 9, 1)) * 256 + Asc(Mid$(St, 10, 1))
                                    .equippedPicture(4) = Asc(Mid$(St, 11, 1)) * 256 + Asc(Mid$(St, 12, 1))
                                End With
                                ReDoLightSources
                            ElseIf Len(St) = 1 Then
                                With Character
                                    cX = (Asc(Mid$(St, 1, 1)) \ 16)
                                    CX2 = cX ^ 2 + 5
                                    Cxo = cX * 32
                                    cY = (Asc(Mid$(St, 1, 1)) And 15)
                                    CY2 = cY ^ 2 + 5
                                    CYO = cY * 32
                                    CWalk = 0
                                    Freeze = False
                                    Transition 1, 0, 0, 0, 1
                                    RedoOwnLight
                                End With
                            End If
                        End If
                        
                    Case 9 'Player left map
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                PlayerLeftMap (A)
                                ReDoLightSources
                            End If
                        End If
                        
                    Case 10 'Player moved
                        If Len(St) = 3 Then
                            A = Asc(Mid$(St, 3, 1))
                            With player(Asc(Mid$(St, 1, 1)))
                                'If .X * 32 = .XO And .Y * 32 = .YO Then
                                '    .X = Asc(Mid$(St, 2, 1))
                                '    .Y = .X And 15
                                '    .X = .X \ 16
                                'Else
                                    .XO = .x * 32
                                    .YO = .y * 32
                                    .x = Asc(Mid$(St, 2, 1))
                                    .y = .x And 15
                                    .x = .x \ 16
                                'End If
                                .D = (A And 7)
                                .WalkStep = (A \ 8)
                                .WalkStart = GetTickCount
                                If .LightSourceNumber > 0 Then
                                    LightSource(.LightSourceNumber).x = .XO + 16
                                    LightSource(.LightSourceNumber).y = .YO
                                End If
                            End With
                        ElseIf Len(St) = 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            b = Asc(Mid$(St, 2, 1))
                            player(A).D = b
                        ElseIf Len(St) = 1 Then
                            cX = (Asc(Mid$(St, 1, 1)) \ 16)
                            CX2 = cX ^ 2 + 5
                            Cxo = cX * 32
                            cY = (Asc(Mid$(St, 1, 1)) And 15)
                            CY2 = cY ^ 2 + 5
                            CYO = cY * 32
                        End If
                        
                    Case 11 'Say
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If player(A).Ignore = False Then PrintChat player(A).Name + " says, " + Chr$(34) + SwearFilter(Mid$(St, 3)) + Chr$(34), 7, Options.FontSize, 1
                        End If
                        
                    Case 12 'You joined map
                        If Len(St) = 14 Then
                            If MapEdit = True Then CloseMapEdit
                            Effects.Clear 'Destroy Effects
                            Projectiles.Clear 'Destroy Projectiles
                            FloatingText.Clear 'Destroy Floating Text
                            ParticleEngineB.Clear 'Destroy all particle sources
                            ParticleEngineF.Clear 'Destroy all particle sources
                            DestroyClickWindow
                            Character.CurProjectile.Key = ""
                            StorageOpen = False
                            SetGUIWindow WINDOW_INVALID 'Close all GUI windows
                            
                            If CMap = 0 Then
                                St1 = ""
                                b = 0
                                DrawInv
                                UpdateSkills
                                For A = 1 To MAXUSERS
                                    With player(A)
                                        If .Sprite > 0 And A <> Character.Index Then
                                            b = b + 1
                                            St1 = St1 + ", " + .Name
                                        End If
                                    End With
                                Next A
                                PrintChat " ", 1, 10
                                PrintChat " ", 1, 10
                                
                                If b > 0 Then
                                    St1 = Mid$(St1, 2)
                                    PrintChat "Welcome to Seyerdin Online!  There are " + CStr(b) + " other players online:" + St1, 15, Options.FontSize, 15
                                Else
                                    PrintChat "Welcome to Seyerdin Online!  There are no other users currently online.", 15, Options.FontSize, 15
                                End If
                                'tickCountMod = 0
                                Load frmMain
                            End If
                            
                            If RemoteWidgetParent <> "" Then
                                UnloadWidgets
                            End If
                            
                            
                            TargetMonster = 10
                            If CurrentTarget.TargetType <> TT_CHARACTER Then
                                CurrentTarget.Target = 0
                                CurrentTarget.TargetType = 0
                            End If
                            STOPMAPCHECK = False
                            ''''''''''''''''''''''''''''''''''''''''
                                CMap = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                                CMap2 = CMap ^ 2 + 5
                                cX = Asc(Mid$(St, 3, 1))
                                cY = Asc(Mid$(St, 4, 1))
                                CX2 = cX ^ 2 + 5
                                CY2 = cY ^ 2 + 5
                                CDir = Asc(Mid$(St, 5, 1))
                                CWalkCode = Asc(Mid$(St, 6, 1))
                                Cxo = cX * 32
                                CYO = cY * 32
                            For A = 0 To 9
                                map.Monster(A).Monster = 0
                                map.Monster(A).HP = 0
                                map.DeadBody(A).Sprite = 0
                                map.DeadBody(A).Counter = 0
                                map.DeadBody(A).Event = 0
                                map.DeadBody(A).BodyType = 0
                                map.DeadBody(A).Frame = 0
                            Next A
                            For A = 0 To 49
                                map.Object(A).Object = 0
                            Next A
                            For A = 0 To 9
                                map.Door(A).BGTile1 = 0
                                map.Door(A).Att = 0
                                map.Door(A).WallTile = 0
                            Next A
                            For A = 0 To MaxTraps
                                map.Trap(A).Counter = 0
                                map.Trap(A).Created = 0
                                map.Trap(A).x = 0
                                map.Trap(A).y = 0
                            Next A
                            For A = 1 To 10
                                map.Fish(A).TimeStamp = 0
                            Next A
                            For A = 1 To MAXUSERS
                                If (player(A).Party = 0 Or player(A).Party <> Character.Party) And (player(A).Guild = 0 Or player(A).Guild <> Character.Guild) Then player(A).map = 0
                            Next A
                            Freeze = True
                            Open MapCacheFile For Random As #1 Len = 2379
                            Get #1, CMap, MapData
                            Close #1
                            MapData = RC4_DecryptString(MapData, MapKey)
                            If HasBeenThere(CMap) = False Then
                                SetMapBit CMap, UCase$(User)
                            End If
                     
                            Dim a1, a2, a3, a4 As Long
                            a1 = Asc(Mid$(MapData, 31, 1)) * 16777216
                            a2 = Asc(Mid$(MapData, 32, 1)) * 65536
                            a3 = Asc(Mid$(MapData, 33, 1)) * 256&
                            a4 = Asc(Mid$(MapData, 34, 1))
                     

                            A = a1 + a2 + a3 + a4
                            b = Asc(Mid$(St, 7, 1)) * 16777216 + Asc(Mid$(St, 8, 1)) * 65536 + Asc(Mid$(St, 9, 1)) * 256& + Asc(Mid$(St, 10, 1))
                            C = CheckSum(MapData)
                            D = Asc(Mid$(St, 11, 1)) * 16777216 + Asc(Mid$(St, 12, 1)) * 65536 + Asc(Mid$(St, 13, 1)) * 256& + Asc(Mid$(St, 14, 1))
                            
                            If A <> b Or C <> D Then
                                SendSocket Chr$(45)
                                RequestedMap = True
                            Else
                                LoadMap MapData
                                RequestedMap = False
                                CalculateAmbientAlpha
                                If MiniMapTab = tsMap Then
                                    CreateMiniMap
                                End If
                            End If
                        End If
                        
                    Case 13 'Error creating character
                        frmMenu.SetStatusText "That name is already in use, please try another.", 2
                        
                    Case 14 'New Map Object
                        If Len(St) = 23 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A <= 49 Then
                                With map.Object(A)
                                    .Object = Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1))
                                    .x = Asc(Mid$(St, 4, 1))
                                    .y = Asc(Mid$(St, 5, 1))
                                    .Light.Radius = Asc(Mid$(St, 6, 1))
                                    .Light.Intensity = Asc(Mid$(St, 7, 1))
                                    .Value = Asc(Mid$(St, 8, 1)) * 16777216 + Asc(Mid$(St, 9, 1)) * 65536 + Asc(Mid$(St, 10, 1)) * 256& + Asc(Mid$(St, 11, 1))
                                    .Prefix = Asc(Mid$(St, 12, 1))
                                    .PrefixVal = Asc(Mid$(St, 13, 1))
                                    .Suffix = Asc(Mid$(St, 14, 1))
                                    .SuffixVal = Asc(Mid$(St, 15, 1))
                                    .Affix = Asc(Mid$(St, 16, 1))
                                    .AffixVal = Asc(Mid$(St, 17, 1))
                                    .TimeStamp = Asc(Mid$(St, 18, 1)) * 16777216 + Asc(Mid$(St, 19, 1)) * 65536 + Asc(Mid$(St, 20, 1)) * 256& + Asc(Mid$(St, 21, 1))
                                    
                                    .DeathObj = IIf(Asc(Mid$(St, 22, 1)) = 1, True, False)
                                    .ObjectColor = Asc(Mid$(St, 23, 1))
                                    
                                    .XOffset = 0
                                    .YOffset = 0
                                    If NotLegalPath(.x, .y, .x - 1, .y) = False Then .XOffset = .XOffset - Int(Rnd * 5)
                                    If NotLegalPath(.x, .y, .x + 1, .y) = False Then .XOffset = .XOffset + Int(Rnd * 5)
                                    If NotLegalPath(.x, .y, .x, .y - 1) = False Then .YOffset = .YOffset - Int(Rnd * 5)
                                    If NotLegalPath(.x, .y, .x, .y + 1) = False Then .YOffset = .YOffset + Int(Rnd * 5)
                                    ReDoLightSources
                                End With
                            End If
                        End If
                        
                    Case 15 'Erase Map Object
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A <= 49 Then
                                With map.Object(A)
                                    .Object = 0
                                    .Light.Intensity = 0
                                    .Light.Type = LT_NONE
                                    .Prefix = 0
                                    .PrefixVal = 0
                                    .Suffix = 0
                                    .SuffixVal = 0
                                    If CurInvObj = 56 + A Then
                                        CurInvObj = 0
                                        DrawCurInvObj
                                    End If
                                    ReDoLightSources
                                End With
                            End If
                        End If
                        
                    Case 16 'Error messages
                        If Len(St) >= 1 Then
                            Select Case Asc(Mid$(St, 1, 1))
                                Case 0 'Custom
                                    If Len(St) > 2 Then
                                        PrintChat Mid$(St, 2), 7, Options.FontSize
                                    End If
                                Case 1 'Inv full
                                    If spamdelay = 0 Then
                                        spamdelay = 1
                                        PrintChat "Your inventory is full!", 7, Options.FontSize, 15
                                    End If
                                Case 2 'Map full
                                    PrintChat "There is too much already on the ground here to drop that.", 7, Options.FontSize, 15
                                'Case 3 'No such object
                                '    PrintChat "No such object.", 7, Options.FontSize, 15
                                Case 4 'No such player
                                    PrintChat "No such player.", 7, Options.FontSize, 15
                                Case 5 'No such monster
                                    PrintChat "No such monster.", 7, Options.FontSize, 15
                                Case 6 'Player is too far away
                                    PrintChat "Player is too far away.", 7, Options.FontSize, 15
                                Case 7 'Monster is too far away
                                    PrintChat "Monster is too far away.", 7, Options.FontSize, 15
                                Case 8 'You cannot use that
                                    PrintChat "You cannot use that object.", 7, Options.FontSize, 15
                                Case 9 'Friendly Zone - can't attack
                                    PrintChat "This is a friendly area, you cannot attack here!", 7, Options.FontSize, 15
                                Case 10 'Cannot attack immortal
                                    PrintChat "You may not attack an immortal!", 7, Options.FontSize, 15
                                Case 11 'You are an immortal
                                    PrintChat "Immortals may not attack other players!", 7, Options.FontSize, 15
                                Case 12 'Can't attack monsters here
                                    PrintChat "You cannot attack these monsters!", 7, Options.FontSize, 15
                                Case 13 'Ban list full
                                    PrintChat "The ban list is full!", 7, Options.FontSize, 15
                                Case 14 'Not invited to join
                                    PrintChat "You have not been invited to join any guild.", 7, Options.FontSize, 15
                                Case 15 'Not enough cash
                                    PrintChat "You do not have enough gold to do that!", 7, Options.FontSize, 15
                                Case 16 'Guild name in use
                                    PrintChat "That name is already used either by another player or guild.  Please try another.", 7, Options.FontSize, 15
                                Case 17 'Guild full
                                    PrintChat "That guild is full!", 7, Options.FontSize, 15
                                Case 18 'too many guilds
                                    PrintChat "Too many guilds already exist.  You may join another guild or try again later.", 7, Options.FontSize, 15
                                Case 19 'cannot attack player -- he is not in guild
                                    PrintChat "That player is not in a guild -- you may not attack non-guild players.", 7, Options.FontSize, 15
                                Case 20 'cannot attack player -- you are not in guild
                                    PrintChat "You must be a member of a guild to attack other players.", 7, Options.FontSize, 15
                                Case 21 'not in a hall
                                    PrintChat "You are not in a guild hall!", 7, Options.FontSize, 15
                                Case 22 'hall already owned
                                    PrintChat "This hall is already owned by another guild.", 7, Options.FontSize, 15
                                Case 23 'already have hall
                                    PrintChat "Your guild already owns a hall.  You must move out of your old hall before you may purchase a new one.", 7, Options.FontSize, 15
                                Case 24 'don't have enough money to buy hall
                                    PrintChat "Your guild does not have enough money in its bank account to buy this hall.  Type /guild hallinfo for the price information of this hall.", 7, Options.FontSize, 15
                                Case 25 'do not own a guild hall
                                    PrintChat "Your guild does not own a hall.", 7, Options.FontSize, 15
                                Case 26 'need 5 members
                                    PrintChat "You must have atleast 5 members in your guild before you may do that.", 7, Options.FontSize, 15
                                Case 27 'Can't afford that
                                    PrintChat "You do not have the items required to purchase that!", 7, Options.FontSize, 15
                                Case 28 'Not in a bank
                                    PrintChat "You are not in a bank!", 7, Options.FontSize, 15
                                Case 29 'too far away
                                    PrintChat "That player is too far away to hit!", 7, Options.FontSize, 15
                                Case 30 'must be level 15 to join guild
                                    PrintChat "You must be at least level 15 to join a guild!", 7, Options.FontSize, 15
                                Case 31 'Invalid Admin Password
                                    PrintChat "Incorrect Administration Password to perform action!", 7, Options.FontSize, 15
                                Case 32 'Must be in a smithy shop
                                    PrintChat "You are not in a blacksmithy shop!", 7, Options.FontSize, 15
                                Case 33 'Do not have enough money
                                    PrintChat "You do not have enough money to repair this object!", 7, Options.FontSize, 15
                                Case 34 'Do not have specified object
                                    PrintChat "You do not have the object to be repaired!", 7, Options.FontSize, 15
                                Case 35 'Party created Successfully
                                    PrintChat "You have created a party.", 7, Options.FontSize, 15
                                Case 36 'leave party
                                    PrintChat "You have left the party.", 7, Options.FontSize, 15
                                Case 37 'Not enough gold
                                    PrintChat "You do not have enough gold to stay here.", 7, Options.FontSize, 15
                                Case 38 'Class cannot use this
                                    PrintChat "Your class cannot use this item.", 7, Options.FontSize, 15
                                Case 39 'Not high enough level
                                    PrintChat "You are too low of a level to use this item.", 7, Options.FontSize, 15
                                Case 40 'only attack pk
                                    PrintChat "You can only attack evil with this spell.", 7, Options.FontSize, 15
                                Case 41
                                    PrintChat "You have died from poison!", 12, Options.FontSize, 15
                                Case 42
                                    PrintChat "Your spell has failed! [Vex]", 15, Options.FontSize, 15
                                Case 43
                                    PrintChat "Your spell has failed! [Mute]", 15, Options.FontSize, 15
                                Case 44
                                    PrintChat "This object cannot be repaired.", 7, Options.FontSize, 15
                                Case 45
                                    PrintChat "You must have both hands free.", 7, Options.FontSize, 15
                                Case 46
                                    PrintChat "Your shield hand is occupied.", 7, Options.FontSize, 15
                                Case 47
                                    PrintChat "You cannot deposit this item.", 7, Options.FontSize, 15
                                Case 48
                                    PrintChat "You do not have any ammo.", 7, Options.FontSize, 15
                                Case 49
                                    PrintChat "You cannot use this ammo with this weapon.", 7, Options.FontSize, 15
                                Case 50
                                    PrintChat "You have joined the party!", 15, Options.FontSize, 15
                                Case 51
                                    PrintChat "You have not been invited to join a party!", 15, Options.FontSize, 15
                                Case 52
                                    PrintChat "You do not have enough gold to stay here!", 7, Options.FontSize, 15
                                Case 53
                                    PrintChat "You cannot mine here!", 7, Options.FontSize, 15
                                Case 54
                                    PrintChat "You mine some ore.", 7, Options.FontSize, 15
                                Case 55
                                    PrintChat "You fail to mine some ore.", 7, Options.FontSize, 15
                                Case 56 'Already in a trade
                                    PrintChat "You are already trading.", 7, Options.FontSize, 15
                                Case 57 'Player already in trade
                                    PrintChat "That player is already in a trade.", 7, Options.FontSize, 15
                                Case 58 'Trading Player doesn't have enough inv slots
                                    PrintChat player(TradeData.player).Name & " does not have enough free inventory slots to trade.", 7, Options.FontSize, 15
                                Case 59 'You do not have enough free inv slots
                                    PrintChat "You do not have enough free inventory slots to trade.", 7, Options.FontSize, 15
                                Case 60
                                    PrintChat "There is an error with this trade.", 7, Options.FontSize, 15
                                Case 61
                                    PrintChat "This object does not need repaired.", 7, Options.FontSize, 15
                                Case 62
                                    PrintChat "You can only trade with people on the same map.", 7, Options.FontSize, 15
                                Case 63
                                    PrintChat "Your guild does not have enough money to buy this symbol (100,000 per part).", 7, Options.FontSize, 15
                                Case 64
                                    PrintChat "Another guild already owns a symbol too similar or matching this one.", 7, Options.FontSize, 15
                                Case 65
                                    PrintChat "Sliders that are not set to a symbol must be set to the far left (No picking unused symbol slots).", 7, Options.FontSize, 15
                                Case 66
                                    PrintChat "Each map can only have one guildmaster door.", 7, Options.FontSize, 15
                            End Select
                        End If
                        
                    Case 17 'New Inv Object
                        If Len(St) = 14 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 And A <= 25 Then
                                If A <= 20 Then
                                    With Character.Inv(A)
                                        .Object = Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1))
                                        .Value = Asc(Mid$(St, 4, 1)) * 16777216 + Asc(Mid$(St, 5, 1)) * 65536 + Asc(Mid$(St, 6, 1)) * 256& + Asc(Mid$(St, 7, 1))
                                        .Prefix = Asc(Mid$(St, 8, 1))
                                        .PrefixValue = Asc(Mid$(St, 9, 1))
                                        .Suffix = Asc(Mid$(St, 10, 1))
                                        .SuffixValue = Asc(Mid$(St, 11, 1))
                                        .Affix = Asc(Mid$(St, 12, 1))
                                        .AffixValue = Asc(Mid$(St, 13, 1))
                                        .ObjectColor = Asc(Mid$(St, 14, 1))
                                    End With
                                Else
                                    With Character.Equipped(A - 20)
                                        .Object = Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1))
                                        .Value = Asc(Mid$(St, 4, 1)) * 16777216 + Asc(Mid$(St, 5, 1)) * 65536 + Asc(Mid$(St, 6, 1)) * 256& + Asc(Mid$(St, 7, 1))
                                        .Prefix = Asc(Mid$(St, 8, 1))
                                        .PrefixValue = Asc(Mid$(St, 9, 1))
                                        .Suffix = Asc(Mid$(St, 10, 1))
                                        .SuffixValue = Asc(Mid$(St, 11, 1))
                                        .Affix = Asc(Mid$(St, 12, 1))
                                        .AffixValue = Asc(Mid$(St, 13, 1))
                                        .ObjectColor = Asc(Mid$(St, 14, 1))
                                        If CurrentTab = tsStats2 Then SetTab tsStats2
                                    End With
                                End If
                                DrawInv
                            End If
                        ElseIf Len(St) = 5 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A <= 20 Then
                                Character.Inv(A).Value = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                            ElseIf A > 20 And A <= 25 Then
                                Character.Equipped(A - 20).Value = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                            End If
                        End If
                        
                    Case 18 'Erase Inv Object
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 And A <= 25 Then
                                If A <= 20 Then
                                    With Character.Inv(A)
                                        .Object = 0
                                        .Prefix = 0
                                        .PrefixValue = 0
                                        .Suffix = 0
                                        .SuffixValue = 0
                                    End With
                                Else
                                    With Character.Equipped(A - 20)
                                        .Object = 0
                                        .Prefix = 0
                                        .PrefixValue = 0
                                        .Suffix = 0
                                        .SuffixValue = 0
                                    End With
                                End If
                                DrawInv
                            End If
                        End If
                        
                    Case 19 'Use Object
                        If Len(St) = 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            b = Asc(Mid$(St, 2, 1))
                            If A >= 1 And A <= 20 Then
                                With Character.Equipped(b)
                                    .Object = Character.Inv(A).Object
                                    .Prefix = Character.Inv(A).Prefix
                                    .PrefixValue = Character.Inv(A).PrefixValue
                                    .Suffix = Character.Inv(A).Suffix
                                    .SuffixValue = Character.Inv(A).SuffixValue
                                    .Value = Character.Inv(A).Value
                                    
                                End With
                                RedoOwnLight
                            End If
                            DrawInv
                        End If
                        
                    Case 20 'Stop using object
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 And A <= 5 Then
                                Character.Equipped(A).Object = 0
                                If CurrentTab = tsStats2 And combatCounter = 10 Then
                                    calculateHpRegen
                                    calculateManaRegen
                                    DrawMoreStats
                                End If
                            End If
                            DrawInv
                        End If
                        
                    Case 21 'Map Data
                        If Len(St) = 2374 Or Len(St) = 2379 Then
                            MapData = St
                            Open MapCacheFile For Random As #1 Len = 2379
                                MapData = RC4_EncryptString(MapData, MapKey)
                                Put #1, CMap, MapData
                            Close #1
                            MapData = RC4_DecryptString(MapData, MapKey)
                            If Len(St) = 2374 Then ' Or Len(St) = 2379 Then
                                LoadMap (St)
                            Else
                                LoadMap MapData
                            End If
                            CalculateAmbientAlpha
                            ShowMap
                            If MiniMapTab = tsMap Then
                                CreateMiniMap
                            End If
                        End If
                        
                    Case 22 'Done Sending Map
                        If RequestedMap = False Then
                            ShowMap
                        End If
                        
                    Case 23 'Squelch
                        Character.Squelched = Asc(Mid$(St, 1, 1))
                        If Character.Squelched = 0 Then PrintChat "You have been unsquelched", 15, Options.FontSize, 15
                        
                    Case 24 'Joined Game

                        frmMenu.SetStatusText "Loading Game Data . . .", 1
                        For A = 1 To MAXUSERS
                            player(A).map = 0
                        Next A
                        Character.Red = 0
                        Character.Green = 0
                        Character.Blue = 0
                        Character.alpha = 0

                        Load frmMain
                        DrawHP
                        DrawEnergy
                        DrawMana
                        ReDoLightSources
                        blnPlaying = True
                        frmMenu.clickedPlay = False
                        UpdateSkills

                        SendSocket Chr(92) + Registry_Read("HKEY_LOCAL_MACHINE\Software\Classes\OddKeys\", "UID2")
                    Case 25 'Tell
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                LastPlayerTellNum = A
                                LastPlayerTellName = player(A).Name
                                If Options.Away = True And Len(Options.AwayMsg) > 0 And Len(Options.AwayMsg) < 513 Then
                                    SendSocket Chr$(63) + Chr$(A) + Options.AwayMsg
                                    AddLog "<" + player(A).Name + "> " + Chr$(34) + SwearFilter(Mid$(St, 2)) + Chr$(34) + " Time: " + CStr(Time) + "/" + CStr(Date)
                                End If
                                    If Options.ForwardUser <> "" Then
                                        If FindPlayer(Options.ForwardUser) <> 0 Then
                                            b = FindPlayer(Options.ForwardUser)
                                            SendSocket Chr$(14) + Chr$(b) + ("Forward Message from " + Character.Name + ": (" + player(A).Name + ") " + Chr$(34) + SwearFilter(Mid$(St, 2)) + Chr$(34))
                                            SendSocket Chr$(14) + Chr$(A) + "Your message has been forwarded to a Helper or a god. They will be with you shortly!"
                                        Else
                                            PrintChat player(FindPlayer(Options.ForwardUser)).Name + " does not exist!", 2, Options.FontSize, 2
                                            Options.ForwardUser = ""
                                        End If
                                    Else
                                        If Options.Away = False Then
                                            With player(A)
                                                If .Ignore = False Then
                                                    PrintChat .Name + " tells you, " + Chr$(34) + SwearFilter(Mid$(St, 2)) + Chr$(34), 10, Options.FontSize, 2
                                                End If
                                            End With
                                        End If
                                    End If
                             End If
                        End If
                        
                    Case 26 'Broadcast
                        If Len(St) >= 2 Then 'And Options.Broadcasts = True Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                If A = Character.Index Then
                                    If Asc(Mid$(St, 2)) = 40 Then
                                        ChatString = "/BROADCAST " + LastBroadcast
                                        Chat.Enabled = True
                                        DrawChatString
                                    ElseIf Asc(Mid$(St, 2)) = 20 Then 'god chat
                                        PrintChat Character.Name + ": " + SwearFilter(Mid$(St, 3)), RGB(220, 0, 150), Options.FontSize, 0, True
                                    Else
                                        PrintChat Character.Name + ": " + SwearFilter(Mid$(St, 3)), CHannelColors(Asc(Mid$(St, 2))), Options.FontSize, 0, True
                                    End If
                                    LastBroadcast = vbNullString
                                Else
                                    If Asc(Mid$(St, 2)) = 20 Then 'god chat
                                      PrintChat player(A).Name + ": " + SwearFilter(Mid$(St, 3)), RGB(220, 0, 150), Options.FontSize, 0, True
                                    Else
                                      If player(A).Ignore = False Then PrintChat player(A).Name + ": " + SwearFilter(Mid$(St, 3)), CHannelColors(Asc(Mid$(St, 2))), Options.FontSize, 0, True
                                    End If
                                End If
                            End If
                        End If
                        
                    Case 27 'Emote
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If player(A).Ignore = False Then PrintChat player(A).Name + " " + SwearFilter(Mid$(St, 2)), 11, Options.FontSize, 3
                        End If
                        
                    Case 28 'Yell
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If player(A).Ignore = False Then PrintChat player(A).Name + " yells, " + Chr$(34) + SwearFilter(Mid$(St, 3)) + Chr$(34), 7, Options.FontSize, 4
                        End If
                        
                    Case 29 'Map Changed
                        If Len(St) = 1 Then
                            PrintChat "This map has been altered by " + player(A).Name + ".", 14, Options.FontSize
                        End If
                        
                    Case 30 'Server Message
                        If Len(St) > 0 Then
                            PrintChat "Server Message: " + St, 9, Options.FontSize, 15
                        End If
                        
                    Case 31 'Object Data
                        If Len(St) >= 14 Then
                            A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                            If A >= 1 Then
                                With Object(A)
                                    .Picture = Asc(Mid$(St, 3, 1))
                                    .Type = Asc(Mid$(St, 4, 1))

                                    For b = 0 To 9
                                        .ObjData(b) = Asc(Mid$(St, 5 + b, 1))
                                    Next b
                                    .MinLevel = Asc(Mid$(St, 15, 1))
                                    .Flags = Asc(Mid$(St, 16, 1))
                                    .Class = GetInt(Mid$(St, 17, 2))
                                    .EquipmentPicture = Asc(Mid$(St, 19, 1))
                                    GetSections2 Mid$(St, 20)
                                    .Name = Section(1)
                                    .Name = Crypt(.Name)
                                    .Description = Section(2)
                                    .Description = Crypt(.Description)
                                    If frmMonster_Loaded = True And frmList.txtContaining.Text = "" Then
                                        frmMonster.cmbObject(0).List(A) = CStr(A) + ": " + .Name
                                        frmMonster.cmbObject(1).List(A) = CStr(A) + ": " + .Name
                                        frmMonster.cmbObject(2).List(A) = CStr(A) + ": " + .Name
                                    End If
                                    If frmNPC_Loaded = True And frmList.txtContaining.Text = "" Then
                                        frmNPC.cmbGiveObject.List(A) = CStr(A) + ": " + .Name
                                        frmNPC.cmbTakeObject.List(A) = CStr(A) + ": " + .Name
                                    End If
                                    If frmList_Loaded = True And frmList.txtContaining.Text = "" Then
                                        frmList.lstObjects.List(A - 1) = CStr(A) + ": " + .Name
                                    End If
                                End With
                            End If
                        End If
                        
                    Case 32 'Monster Data
                        If Len(St) >= 14 Then
                            A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                            St = Mid$(St, 2)
                            If A >= 1 Then
                                With Monster(A)
                                    .Sprite = Asc(Mid$(St, 2, 1))
                                    .MaxHP = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                                    .Flags = Asc(Mid$(St, 5, 1))
                                    .DeathSound = Asc(Mid$(St, 6, 1))
                                    .AttackSound = Asc(Mid$(St, 7, 1))
                                    .alpha = Asc(Mid$(St, 8, 1))
                                    .Red = Asc(Mid$(St, 9, 1))
                                    .Green = Asc(Mid$(St, 10, 1))
                                    .Blue = Asc(Mid$(St, 11, 1))
                                    .Light = Asc(Mid$(St, 12, 1))
                                    .Flags2 = Asc(Mid$(St, 13, 1))
                                    If Len(St) >= 14 Then
                                        .Name = Mid$(St, 14)
                                        .Name = Crypt(.Name)
                                    Else
                                        .Name = ""
                                    End If
                                    If frmList_Loaded = True Then
                                        'frmList.lstMonsters.List(A - 1) = CStr(A) + ": " + .Name
                                    End If
                                    If frmMapProperties_Loaded = True Then
                                        frmMapProperties.cmbMonster(0).List(A) = CStr(A) + ": " + .Name
                                        frmMapProperties.cmbMonster(1).List(A) = CStr(A) + ": " + .Name
                                        frmMapProperties.cmbMonster(2).List(A) = CStr(A) + ": " + .Name
                                        frmMapProperties.cmbMonster(3).List(A) = CStr(A) + ": " + .Name
                                        frmMapProperties.cmbMonster(4).List(A) = CStr(A) + ": " + .Name
                                    End If
                                End With
                            End If
                        End If
                        
                    Case 33 'Edit Object Data
                        If Len(St) = 18 Then
                            A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                            If frmObject_Loaded = False Then Load frmObject
                            With frmObject
                                .lblNumber = A
                                .txtName = Object(A).Name
                                .txtDescription = Object(A).Description
                                b = Asc(Mid$(St, 3, 1))
                                If Object(A).Picture > 0 Then
                                    .sclPicture = Object(A).Picture
                                    If ExamineBit(b, 6) Then .sclPicture = .sclPicture + 255
                                Else
                                    .sclPicture = 1
                                End If
                                
                                For C = 0 To 7
                                    If ExamineBit(b, C) Then
                                        frmObject.chkFlags(C).Value = 1
                                    Else
                                        frmObject.chkFlags(C).Value = 0
                                    End If
                                Next C
                                If Object(A).Type < .cmbType.ListCount Then
                                    .cmbType.ListIndex = 0
                                    .cmbType.ListIndex = Object(A).Type
                                Else
                                    .cmbType.ListIndex = 0
                                End If
                                Dim H As Long, i As Long, j As Long, G As Long
                                b = Asc(Mid$(St, 4, 1))
                                C = Asc(Mid$(St, 5, 1))
                                D = Asc(Mid$(St, 6, 1))
                                E = Asc(Mid$(St, 7, 1))
                                F = Asc(Mid$(St, 8, 1))
                                G = Asc(Mid$(St, 9, 1))
                                H = Asc(Mid$(St, 10, 1))
                                i = Asc(Mid$(St, 11, 1))
                                j = Asc(Mid$(St, 12, 1))
                                
                                With frmObject
                                    Select Case Object(A).Type
                                        Case 1 'Weapon
                                            .sclWeaponhp = IIf(b > 0, b, 1)
                                            .sclWeaponMin = C
                                            .sclWeaponMax = IIf(D * 256# + E > 10000, 10000, D * 256 + E)
                                            .sclASpeed = IIf(F > 10 Or F < 0, 1, F)
                                        Case 2 'Shield
                                            .sclShieldHP = IIf(b > 0, b, 1)
                                            .sclShieldDefense = C
                                            .sclShieldDamagePercent = D
                                            .sclshieldmagicchance = E
                                            .sclshieldmagicpercent = F
                                        Case 3 'Armor
                                            .sclArmorHP = IIf(b > 0, b, 1)
                                            .sclArmorDefense = IIf(C * 256 + D <= 1000, C * 256 + D, 1)
                                            .sclArmorResist = E
                                            .sclArmorMDefense = F
                                            .sclArmorMResist = G
                                        Case 4 'Helmet
                                            .sclHelmetHP = IIf(b > 0, b, 1)
                                            .sclHelmetDefense = IIf(C * 256 + D <= 1000, C * 256 + D, 1)
                                            .sclHelmResist = E
                                            .sclHelmMDefense = F
                                            .sclHelmMResist = G
                                        Case 5 'Potion
                                            .cmbPotionType.ListIndex = b
                                            .sclPotionValue = IIf(C * 256 + D <= 1000, C * 256 + D, 1)
                                        Case 6 'Money
                                            .sclStackSize = b
                                        Case 7 'Key
                                            .chkKeyUnlim.Value = b
                                        Case 8 'Ring
                                            .cmbRingType.ListIndex = b
                                            .sclRingHP = IIf(C > 0, C, 1)
                                            .sclRingAmount = D
                                        Case 10 'Projectile
                                            .sclProjectileHP = IIf(b > 0, b, 1)
                                            .sclProjectileRange = IIf(C > 0, C, 1)
                                            .sclProjectilePlus = D
                                            .sclProjectileAmmoType = E
                                            .sclProjectileSpeed = F
                                        Case 11 'Ammo
                                            .sclAmmoLimit = IIf(b > 0, b, 1)
                                            .sclAmmoAnimation = C
                                            .sclAmmoMin = D
                                            .sclAmmoMax = E
                                            .sclAmmoType = F
                                        Case Else
                                    End Select
                                End With

                                .sclMagicAmp = Asc(Mid$(St, 13, 1))

                                If Asc(Mid$(St, 14, 1)) = 0 Then
                                    .sclMinLevel = 1
                                Else
                                    .sclMinLevel = Asc(Mid$(St, 14, 1))
                                End If
                                A = GetInt(Mid$(St, 15, 2))
                                For C = 0 To MAX_CLASS - 1
                                    .lstClass.Selected(C) = IIf(A And (2 ^ C), True, False)
                                Next C
                                If Asc(Mid$(St, 17, 1)) = 0 Then
                                    .sclLevel = 1
                                Else
                                    .sclLevel = Asc(Mid$(St, 17, 1))
                                End If
                                    .sclEquipmentPicture = Asc(Mid$(St, 18, 1))
                                    .lblEquipmentPicture.Caption = Asc(Mid$(St, 18, 1))
                                .Show 1
                            End With
                        End If
                        
                    Case 34 'Edit Monster Data
                        If Len(St) = 46 Then
                            A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                            St = Mid$(St, 2)
                            
                            If frmMonster_Loaded = False Then Load frmMonster
                            With frmMonster
                                .lblNumber = A
                                .txtName = Monster(A).Name
 
                                b = Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1))
                                If b > 0 Then .sclHP = b Else .sclHP = 1
                                b = Asc(Mid$(St, 4, 1))
                                If b > 0 Then .sclMin = b Else .sclMin = 1
                                b = GetInt(Mid$(St, 5, 2))
                                If b > 0 Then .sclMax = b Else .sclMax = 1
                                b = Asc(Mid$(St, 7, 1))
                                .sclArmor = b
                                b = Asc(Mid$(St, 8, 1))
                                If b > 0 Then .sclSight = b Else .sclSight = 1
                                b = Asc(Mid$(St, 9, 1))
                                If b <= 100 Then .sclAgility = b Else .sclAgility = 100
                                b = Asc(Mid$(St, 10, 1))
                                For C = 0 To 7
                                    If ExamineBit(CByte(b), CByte(C)) = True Then
                                        .chkFlag(C) = 1
                                    Else
                                        .chkFlag(C) = 0
                                    End If
                                Next C
                                .cmbObject(0).ListIndex = Asc(Mid$(St, 11, 1)) * 256 + Asc(Mid$(St, 12, 1))
                                .txtValue(0) = Asc(Mid$(St, 13, 1))
                                .cmbObject(1).ListIndex = Asc(Mid$(St, 14, 1)) * 256 + Asc(Mid$(St, 15, 1))
                                .txtValue(1) = Asc(Mid$(St, 16, 1))
                                .cmbObject(2).ListIndex = Asc(Mid$(St, 17, 1)) * 256 + Asc(Mid$(St, 18, 1))
                                .txtValue(2) = Asc(Mid$(St, 19, 1))
                                .sclEXP = Asc(Mid$(St, 20, 1)) * 256 + Asc(Mid$(St, 21, 1))
                                .txtChance(0) = Asc(Mid$(St, 22, 1))
                                .txtChance(1) = Asc(Mid$(St, 23, 1))
                                .txtChance(2) = Asc(Mid$(St, 24, 1))
                                .sclLevel = Asc(Mid$(St, 25, 1))
                                .sclMagicResist = Asc(Mid$(St, 26, 1))
                                A = Asc(Mid$(St, 27, 1)) * 16777216 + Asc(Mid$(St, 28, 1)) * 65536 + Asc(Mid$(St, 29, 1)) * 256& + Asc(Mid$(St, 30, 1))
                                For b = 0 To frmMonster.lstStatusEffect.ListCount - 1
                                    .lstStatusEffect.Selected(b) = ExamineBit(A, b)
                                Next b
                                A = Asc(Mid$(St, 31, 1)) * 16777216 + Asc(Mid$(St, 32, 1)) * 65536 + Asc(Mid$(St, 33, 1)) * 256& + Asc(Mid$(St, 34, 1))
                                For b = 0 To frmMonster.lstMonsterType.ListCount - 1
                                    .lstMonsterType.Selected(b) = ExamineBit(A, b)
                                Next b
                                .sclDeathSound = Asc(Mid$(St, 35, 1))
                                .sclAttackSound = Asc(Mid$(St, 36, 1))
                                .sclMoveSpeed = Asc(Mid$(St, 37, 1))
                                .sclAttackSpeed = Asc(Mid$(St, 38, 1))
                                .sclWander = Asc(Mid$(St, 39, 1))
                                .sclAlpha = Asc(Mid$(St, 40, 1))
                                .sclRed = Asc(Mid$(St, 41, 1))
                                .sclGreen = Asc(Mid$(St, 42, 1))
                                .sclBlue = Asc(Mid$(St, 43, 1))
                                .sclLight = Asc(Mid$(St, 44, 1))
                                b = Asc(Mid$(St, 45, 1))
                                For C = 0 To 7
                                    If ExamineBit(CByte(b), CByte(C)) = True Then
                                        .chkFlag2(C) = 1
                                    Else
                                        .chkFlag2(C) = 0
                                    End If
                                Next C
                                A = Asc(Mid$(St, 1, 1))
                                If Monster(A).Sprite > 0 Then
                                    .sclSprite = Monster(A).Sprite
                                Else
                                    .sclSprite = 1
                                End If
                                
                                If .chkFlag2(1) Then .sclSprite = .sclSprite + 255
                                If .chkFlag2(0) Or .chkFlag2(2) Then .sclSprite.max = 100
                                .Enabled = True
                                .Show vbModal
                            End With
                        End If
                    Case 35 'Repeat
                        If Len(St) >= 1 Then
                            SendSocket St
                        End If
                        
                    Case 36 'Door Open
                        If Len(St) = 4 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A <= 9 Then
                                With map.Door(A)
                                    .x = Asc(Mid$(St, 2, 1))
                                    .y = Asc(Mid$(St, 3, 1))
                                    .Att = map.Tile(.x, .y).Att
                                    
                                    C = Asc(Mid$(St, 4, 1))
                                    
                                    For b = 0 To 3
                                        .AttData(b) = map.Tile(.x, .y).AttData(b)
                                        If ExamineBit(C, 1) Then map.Tile(.x, .y).AttData(b) = 0
                                    Next b
                                    .BGTile1 = map.Tile(.x, .y).BGTile1
                                    .WallTile = map.Tile(.x, .y).WallTile
                                    
                                    
                                    If ExamineBit(C, 1) Then map.Tile(.x, .y).Att = 18
                                    
                                    map.Tile(.x, .y).BGTile1 = 0
                                    If ExamineBit(C, 0) Then map.Tile(.x, .y).WallTile = 0
                                    mapChangedBg(.x, .y) = True
                                End With
                            End If
                        End If
                        
                    Case 37 'Close Door
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A <= 9 Then
                                With map.Door(A)
                                    map.Tile(.x, .y).Att = .Att
                                    For b = 0 To 3
                                        map.Tile(.x, .y).AttData(b) = .AttData(b)
                                    Next b
                                    map.Tile(.x, .y).BGTile1 = .BGTile1
                                    map.Tile(.x, .y).WallTile = .WallTile
                                    .BGTile1 = 0
                                    .Att = 0
                                    .WallTile = 0
                                    mapChangedBg(.x, .y) = True
                                End With
                            End If
                        End If
                        
                    Case 38 'New Map Monster
                        If Len(St) = 8 Then
                            A = Asc(Mid$(St, 1, 1))
                            
                            If A <= 9 Then
                                With map.Monster(A)
                                    .Monster = Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1))
                                    St = Mid$(St, 2)
                                    .x = Asc(Mid$(St, 3, 1))
                                    .OX = .x
                                    .y = Asc(Mid$(St, 4, 1))
                                    .OY = .y
                                    .D = Asc(Mid$(St, 5, 1))
                                    .HP = Asc(Mid$(St, 6, 1)) * 256 + Asc(Mid$(St, 7, 1))
                                    .XO = .x * 32
                                    .YO = .y * 32
                                    .A = 0
                                    .R = 0
                                    .G = 0
                                    .b = 0
                                    If Monster(.Monster).Light > 0 Then
                                        ReDoLightSources
                                    End If
                                End With
                            End If
                        Else
                            If Len(St) = 12 Then
                                A = Asc(Mid$(St, 1, 1))
                                If A <= 9 Then
                                    With map.Monster(A)
                                        .Monster = Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1))
                                        St = Mid$(St, 2)
                                        .x = Asc(Mid$(St, 3, 1))
                                        .OX = .x
                                        .y = Asc(Mid$(St, 4, 1))
                                        .OY = .y
                                        .D = Asc(Mid$(St, 5, 1))
                                        .HP = Asc(Mid$(St, 6, 1)) * 256 + Asc(Mid$(St, 7, 1))
                                        .XO = .x * 32
                                        .YO = .y * 32
                                        .R = Asc(Mid$(St, 8, 1))
                                        .G = Asc(Mid$(St, 9, 1))
                                        .b = Asc(Mid$(St, 10, 1))
                                        .alpha = Asc(Mid$(St, 11, 1))
                                        
                                        If .R > Monster(.Monster).Red Then .R = Monster(.Monster).Red
                                        If .G > Monster(.Monster).Green Then .G = Monster(.Monster).Green
                                        If .b > Monster(.Monster).Blue Then .b = Monster(.Monster).Blue
                                        If .alpha > Monster(.Monster).alpha Then .alpha = Monster(.Monster).alpha
                                        
                                        
                                        If Monster(.Monster).Light > 0 Then
                                            ReDoLightSources
                                        End If
                                    End With
                                End If
                            End If
                        
                        

                        End If
                        
                    Case 39 'Monster Die
                        If Len(St) = 1 Or Len(St) = 3 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 10 Then
                                C = 1
                                A = A - 10
                            Else
                                C = 0
                            End If
                            If Len(St) = 3 Then
                                D = Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1))
                                FloatingText.Add map.Monster(A).x * 32, map.Monster(A).y * 32 - 16, CStr(D), &HFFFF0000
                            End If
                            If A <= 9 Then
                                If map.Monster(A).Monster > 0 Then
                                    If Monster(map.Monster(A).Monster).DeathSound > 0 Then
                                        PlayWav Monster(map.Monster(A).Monster).DeathSound
                                    End If
                                    If TargetMonster = A Then TargetMonster = 10
                                    If C = 0 Then
                                        b = GetFreeDeadBody
                                        map.DeadBody(b).MonNum = map.Monster(A).Monster
                                        map.DeadBody(b).Sprite = Monster(map.Monster(A).Monster).Sprite
                                        map.DeadBody(b).Counter = 450
                                        map.DeadBody(b).x = map.Monster(A).x
                                        map.DeadBody(b).y = map.Monster(A).y
                                        map.DeadBody(b).BodyType = TT_MONSTER
                                        If C Then
                                            map.DeadBody(b).Frame = map.Monster(A).D
                                        Else
                                            map.DeadBody(b).Frame = 13
                                        End If
                                        map.DeadBody(b).Event = 0
                                        map.DeadBody(b).Name = Monster(map.Monster(A).Monster).Name
    
                                    End If
                                    map.Monster(A).Monster = 0
                                    map.Monster(A).HP = 0
                                    If map.Monster(A).LightSourceNumber > 0 Then
                                        map.Monster(A).LightSourceNumber = 0
                                        ReDoLightSources
                                    End If
                                    MonsterDied A
                                End If
                            End If
                        End If
                        
                    Case 40 'Monster Move
                        If Len(St) = 3 Then
                            A = (Asc(Mid$(St, 1, 1)) And 240) / 16
                            If A <= 9 Then
                                With map.Monster(A)
                                    'If CLng(.X) * 32 <> .XO Or CLng(.Y) * 32 <> .YO Then
                                    '    .X = (Asc(Mid$(St, 2, 1)) And 240) / 16
                                    '    .Y = Asc(Mid$(St, 2, 1)) And 15
                                    '    .XO = .X * 32
                                    '    .YO = .Y * 32
                                    'Else
                                        .XO = .x * 32
                                        .YO = .y * 32
                                        .OX = .x
                                        .OY = .y
                                        .x = (Asc(Mid$(St, 2, 1)) And 240) / 16
                                        .y = Asc(Mid$(St, 2, 1)) And 15
                                        .W = 1
                                    'End If
                                    .D = (Asc(Mid$(St, 1, 1)) And 14) / 2
                                    b = Asc(Mid$(St, 3, 1))
                                    .StartTick = GetTickCount
                                    .EndTick = .StartTick + ((b + 1) * 190)
                                    If b < 1 Then
                                        .WalkStep = 4
                                    ElseIf b >= 1 And b < 4 Then
                                        .WalkStep = 2
                                    Else
                                        .WalkStep = 1
                                    End If
                                End With
                            End If
                        ElseIf Len(St) = 1 Then
                            A = (Asc(Mid$(St, 1, 1)) And 240) / 16
                            If A <= 9 Then
                                map.Monster(A).D = (Asc(Mid$(St, 1, 1)) And 15)
                            End If
                        End If
                        
                    Case 41 'Monster Attack
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A <= 9 Then
                                map.Monster(A).A = 10
                                
                                b = Sqr((cX - map.Monster(A).x) * (cX - map.Monster(A).x) + (cY - map.Monster(A).y) * (cY - map.Monster(A).y))
                                
                                If Monster(map.Monster(A).Monster).AttackSound > 0 Then
                                    PlayWav Monster(map.Monster(A).Monster).AttackSound, b
                                Else
                                    PlayWav 2, b
                                End If
                                
                                
                            End If
                        End If
                        
                    Case 42 'Player Attack
                        If Len(St) = 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A = Character.Index Then
                                CAttack = 5
                                PlayWav (14)
                            Else
                                If Asc(Mid$(St, 2, 1)) = 1 Then
                                    player(A).A = 10
                                    b = Sqr((cX - player(A).x) * (cX - player(A).x) + (cY - player(A).y) * (cY - player(A).y))
                                    PlayWav 14, b
                                End If
                            End If
                        End If
                        
                    Case 43 'You hit player
                        If Len(St) = 4 Then
                            A = Asc(Mid$(St, 2, 1))
                            If A >= 1 Then
                                b = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                                Select Case Asc(Mid$(St, 1, 1))
                                    Case 0
                                        PlayWav 11 + Int(Rnd * 3)
                                        CAttack = 5
                                    Case 1
                                        PlayWav 14
                                End Select
                            End If
                        End If
                        
                    Case 44 'You hit monster
                        If Len(St) = 5 Then
                            C = Asc(Mid$(St, 1, 1))
                            D = Asc(Mid$(St, 2, 1))
                            A = Asc(Mid$(St, 3, 1))
                            If A >= 10 Then
                                C = 1
                                A = A - 10
                            End If
                            If A >= 0 And A <= 9 Then
                                If map.Monster(A).Monster > 0 Then
                                    b = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1))
                                    map.Monster(A).HP = map.Monster(A).HP - b
                                    If C = 0 And D = 0 Then FloatingText.Add map.Monster(A).x * 32, map.Monster(A).y * 32 - 16, CStr(b), &HFFFF0000
                                    CAttack = 5
                                    Select Case C
                                        Case 0
                                            PlayWav 11 + Int(Rnd * 3)
                                        Case 1
                                            PlayWav 14
                                    End Select
                                End If
                            End If
                        ElseIf Len(St) = 4 Then
                            C = Asc(Mid$(St, 1, 1)) 'hit/miss
                            A = Asc(Mid$(St, 2, 1))
                            b = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                            With map.Monster(A)
                                If C = 0 Then FloatingText.Add .x * 32, .y * 32 - 32, CStr(b), &HFFFF0000
                                .HP = .HP - b
                           
                                b = Sqr((cX - .x) * (cX - .x) + (cY - .y) * (cY - .y))
                                PlayWav 11 + Int(Rnd * 3), b
                            End With
                        End If
                        
                    Case 45 'You killed player
                        If Len(St) = 5 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                PlayWav 8
                                With player(A)
                                    If .Status = 1 Then
                                        PrintChat "You have put the evil murderer " + .Name + " to justice!", 12, Options.FontSize
                                        .Status = 0
                                    Else
                                        PrintChat "You have murdered " + .Name + " in cold blood!", 12, Options.FontSize
                                        If ExamineBit(map.Flags(0), 2) = False Then Character.Status = 1
                                    End If
                                    CLastKilled = .Name
                                    b = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                                End With
                            End If
                        End If
                        
                    Case 46 'Change HP
                        If Len(St) = 2 Then
                            A = Character.HP
                            Character.HP = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                            If Character.HP < A Then
                                If combatCounter < 5 Then
                                    combatCounter = 5
                                    drawTopBar
                                End If
                            End If
                            DrawHP
                        End If
                        
                    Case 47 'Change Energy
                        If Len(St) = 2 Then
                            Character.Energy = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                            DrawEnergy
                        End If
                        
                    Case 48 'Change Mana
                        If Len(St) = 2 Then
                            Character.Mana = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                            DrawMana
                        End If
                        
                    Case 49 'Player Hit You
                        If Len(St) = 6 Then
                            If combatCounter < 30 Then combatCounter = 30
                            If CurrentTab = tsStats2 And combatCounter = 30 Then
                                calculateHpRegen
                                calculateManaRegen
                                DrawMoreStats
                            End If
                            drawTopBar
                            A = Asc(Mid$(St, 2, 1))
                            If A >= 1 Then
                                b = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                                C = Asc(Mid$(St, 5, 1)) * 256 + Asc(Mid$(St, 6, 1))
                                If b > 0 Then
                                    If Character.HP > b Then
                                        Character.HP = Character.HP - b
                                    Else
                                        Character.HP = 0
                                    End If
                                    DrawHP
                                End If
                                Character.Energy = C
                                DrawEnergy
                                player(A).A = 10
                                Select Case Asc(Mid$(St, 1, 1))
                                    Case 0
                                        PlayWav 11 + Int(Rnd * 3)
                                    Case 1
                                        PlayWav 14
                                End Select
                            End If
                        End If
                        
                    Case 50 'Monster Hit You
                        If Len(St) = 4 Then
                            If combatCounter < 10 Then combatCounter = 10
                            If CurrentTab = tsStats2 And combatCounter = 10 Then
                                calculateHpRegen
                                calculateManaRegen
                                DrawMoreStats
                            End If
                            drawTopBar
                            A = Asc(Mid$(St, 2, 1))
                            If A <= 9 Then
                                b = map.Monster(A).Monster
                                If b > 0 Then
                                    C = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                                    Select Case Asc(Mid$(St, 1, 1))
                                        Case 0
                                            If Monster(b).AttackSound > 0 Then
                                                PlayWav Monster(b).AttackSound
                                            Else
                                                PlayWav 2
                                            End If
                                            If C > 0 Then
                                                If Character.HP > C Then
                                                    Character.HP = Character.HP - C
                                                Else
                                                    Character.HP = 0
                                                End If
                                                DrawHP
                                            End If
                                            map.Monster(A).A = 10
                                        Case 1
                                            PlayWav 3
                                    End Select
                                End If
                            End If
                        End If
                        
                    Case 51 'You killed the monster
                        If Len(St) = 5 Or Len(St) = 7 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 10 Then
                                C = 1
                                A = A - 10
                            Else
                                C = 0
                            End If
                            If A <= 9 Then
                                If TargetMonster = A Then TargetMonster = 10
                                With map.Monster(A)
                                    D = map.Monster(A).Monster
                                    If D > 0 Then
                                        If Monster(D).DeathSound > 0 Then
                                            PlayWav Monster(D).DeathSound
                                        End If
                                        If Len(St) = 7 Then
                                            E = Asc(Mid$(St, 6, 1)) * 256 + Asc(Mid$(St, 7, 1))
                                            FloatingText.Add .x * 32, .y * 32 - 32, CStr(E), &HFFFF0000
                                        End If
                                        b = GetFreeDeadBody
                                        map.DeadBody(b).MonNum = D
                                        map.DeadBody(b).Counter = 450
                                        map.DeadBody(b).Sprite = Monster(D).Sprite
                                        map.DeadBody(b).x = .x
                                        map.DeadBody(b).y = .y
                                        map.DeadBody(b).Name = Monster(D).Name
                                        map.DeadBody(b).BodyType = TT_MONSTER
                                        If C Then
                                            map.DeadBody(b).Frame = .D
                                        Else
                                            map.DeadBody(b).Frame = 13
                                        End If
                                        .Monster = 0
                                        .HP = 0
                                        If .LightSourceNumber > 0 Then
                                            .LightSourceNumber = 0
                                            ReDoLightSources
                                        End If
                                    End If
                                End With
                                b = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                                MonsterDied (A)
                            End If
                        End If
                        
                    Case 52 'Player Killed You
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                With player(A)
                                    If Character.Status = 1 Then
                                        PrintChat .Name & " has put you to justice!  You have lost 1/5 of your experience!", 12, Options.FontSize
                                        Character.Status = 0
                                    Else
                                        PrintChat .Name & " has murdered you in cold blood!  You have lost 1/5 of your experience!", 12, Options.FontSize, 15
                                        If ExamineBit(map.Flags(1), 2) = False Then .Status = 1
                                    End If
                                    CLastKiller = .Name
                                End With
                                YouDied
                            End If
                        End If
                    
                    Case 53 'Monster Killed You
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(2, 1))
                            If A >= 1 Then
                                PlayWav 8
                                With Monster(A)
                                    PrintChat "The " + .Name + " has killed you!  You have lost 1/5 of your experience!", 12, Options.FontSize, 15
                                End With
                                If Character.Status = 1 Then Character.Status = 0
                                YouDied
                            End If
                        End If
                        
                    Case 54 'You died
                        PlayWav 8
                        If Character.Status = 1 Then Character.Status = 0
                        YouDied
                        PrintChat "You were killed by " + Mid$(St, 1) + ".  You have lost 1/5 of your experience!", 12, Options.FontSize, 15
                        
                    Case 55 'PLayer died
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                With player(A)
                                        PrintChat .Name + " has been killed by " + Mid$(St, 2) + "!", 12, Options.FontSize, 15
                                        If CurrentTarget.TargetType = TT_PLAYER Then
                                            If CurrentTarget.Target = A Then
                                                CurrentTarget.TargetType = 0
                                            End If
                                        End If
                                    If .Status = 1 Then .Status = 0
                                End With
                            End If
                        End If
                    Case 56 'Text
                        If Len(St) >= 2 Then
                            PrintChat Mid$(St, 2), Asc(Mid$(St, 1, 1)), Options.FontSize, 15
                        End If
                    
                    Case 57 'Object Breaks
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 And A <= 5 Then
                                With Character.Equipped(A)
                                    If .Object > 0 Then
                                        PrintChat "Your " & Object(.Object).Name & " breaks!", 12, Options.FontSize
                                    End If
                                    .Object = 0
                                    DrawInv
                                End With
                            End If
                        End If
                    
                    Case 58 'Ping
                        'SendSocket Chr$(29) + QuadChar(GetTickCount) 'Pong
                        
                    Case 59 'Level Up
                        If Len(St) = 10 Then
                            With Character
                                .Level = .Level + 1
                                .MaxHP = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                                .MaxEnergy = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                                .MaxMana = Asc(Mid$(St, 5, 1)) * 256 + Asc(Mid$(St, 6, 1))
                                .StatPoints = Asc(Mid$(St, 7, 1)) * 256 + Asc(Mid$(St, 8, 1))
                                .skillPoints = Asc(Mid$(St, 9, 1)) * 256 + Asc(Mid$(St, 10, 1))
                                .Experience = 0
                                PrintChat "Level Up!  You are now level " + CStr(.Level) + ".  You have " + CStr(.StatPoints) + " stat points.  Click on the Stats tab to spend them!", 12, Options.FontSize, 15
                                DrawHP
                                DrawEnergy
                                DrawMana
                                UpdateSkills
                            End With
                            A = 1
                        End If
                        If Len(St) = 4 Then
                            Character.StatPoints = GetInt(Mid$(St, 1, 2))
                            Character.skillPoints = GetInt(Mid$(St, 3, 2))
                        End If
                            Tstr = 0
                            TAgi = 0
                            TEnd = 0
                            TWis = 0
                            TempVar5 = Character.StatPoints
                            TCon = 0
                            TInt = 0
                            DrawTrainBars
                            DrawExperience
                        
                    Case 60 'Experience Change
                        If Len(St) = 4 Then
                            Character.Experience = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                            DrawTrainBars
                            DrawExperience
                        End If
                        
                    Case 61 'Player killed by player
                        If Len(St) = 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            b = Asc(Mid$(St, 2, 1))
                            If A >= 1 And b >= 1 Then
                                With player(A)
                                    If .Status = 1 Then
                                        PrintChat player(b).Name + " has put " + .Name + " to justice!", 12, Options.FontSize, 15
                                        .Status = 0
                                    Else
                                        PrintChat player(b).Name + " has murdered " + .Name + " in cold blood!", 12, Options.FontSize, 15
                                        If ExamineBit(map.Flags(0), 2) = False Then player(b).Status = 1
                                    End If
                                    If .map = CMap Then
                                        PlayWav 8
                                        If CurrentTarget.TargetType = TT_PLAYER Then
                                            If CurrentTarget.Target = A Then
                                                CurrentTarget.TargetType = 0
                                            End If
                                        End If
                                    End If
                                End With
                            End If
                        ElseIf Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                If A = Character.Index Then
                                    PrintChat "You have died!", 12, Options.FontSize, 15
                                Else
                                    PrintChat player(A).Name + " has died!", 12, Options.FontSize, 15
                                    If player(A).map = CMap Then
                                        PlayWav 8
                                    End If
                                End If
                            End If
                        End If
                        
                    Case 62 'Player killed by monster
                        If Len(St) = 3 Then
                            A = Asc(Mid$(St, 1, 1))
                            'b = Asc(Mid$(St, 2, 1))*256 + asc(mid$(st,3,1))
                            If A >= 1 And b >= 1 Then
                                With player(A)
                                    If .map = CMap Then
                                        PlayWav 8
                                        'PrintChat .Name + " has been killed by a " + Monster(B).Name + "!", 12, Options.FontSize, 15
                                        If CurrentTarget.TargetType = TT_PLAYER Then
                                            If CurrentTarget.Target = A Then
                                                CurrentTarget.TargetType = 0
                                            End If
                                        End If
                                    End If
                                    If .Status = 1 Then .Status = 0
                                End With
                            End If
                        End If
                    
                    Case 63 'Player Sprite Changed
                        If Len(St) = 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            b = Asc(Mid$(St, 2, 1))
                            If A >= 1 And b >= 1 Then
                                If A = Character.Index Then
                                    Character.Sprite = b
                                Else
                                    If player(A).Sprite > 0 Then
                                        player(A).Sprite = b
                                    End If
                                End If
                            End If
                        End If
                        
                    Case 64 'Player Name Change
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                If A = Character.Index Then
                                    Character.Name = Mid$(St, 2)
                                    DrawTrainBars
                                Else
                                    If player(A).Sprite > 0 Then
                                        player(A).Name = Mid$(St, 2)
                                    End If
                                End If
                            End If
                        End If
                        
                    Case 65 'Changed access
                        If Len(St) = 1 Then
                            Character.Access = Asc(Mid$(St, 1, 1))
                            If Character.Access > 0 Then
                                Character.Status = 3
                            Else
                                If Character.Status = 0 Then
                                    Character.Status = 0
                                End If
                            End If
                        End If
                        
                    Case 66 'Player banned
                        If Len(St) >= 2 Then
                            b = Asc(Mid$(St, 1, 1))
                            C = Asc(Mid$(St, 2, 1))
                            If b >= 1 And C >= 1 Then
                                If C >= 1 Then
                                    If Len(St) > 2 Then
                                        PrintChat player(b).Name + " has been banned by " + player(C).Name + ": " + Mid$(St, 3), 15, Options.FontSize, 15
                                    Else
                                        PrintChat player(b).Name + " has been banned by " + player(C).Name + "!", 15, Options.FontSize, 15
                                    End If
                                Else
                                    If Len(St) > 2 Then
                                        PrintChat player(b).Name + " has been banned: " + Mid$(St, 3), 15, Options.FontSize, 15
                                    Else
                                        PrintChat player(b).Name + " has been banned!", 15, Options.FontSize, 15
                                    End If
                                End If
                            End If
                        End If
                        
                    Case 67 'Booted
                        If Len(St) >= 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                If Len(St) > 1 Then
                                    MessageBox frmMenu.hwnd, "You have been booted from Seyerdin Online by " + player(A).Name + ": " + Mid$(St, 2), TitleString, vbOKOnly + vbExclamation
                                Else
                                    MessageBox frmMenu.hwnd, "You have been booted from Seyerdin Online by " + player(A).Name + "!", TitleString, vbOKOnly + vbExclamation
                                End If
                            Else
                                If Len(St) > 1 Then
                                    MessageBox frmMenu.hwnd, "You have been booted from Seyerdin Online: " + Mid$(St, 2), TitleString, vbOKOnly + vbExclamation
                                Else
                                    MessageBox frmMenu.hwnd, "You have been booted from Seyerdin Online", TitleString, vbOKOnly + vbExclamation
                                End If
                            End If
                            CloseClientSocket 0, False
                        End If
                        
                    Case 68 'Player Booted
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            b = Asc(Mid$(St, 2, 1))
                            If A >= 1 Then
                                If b >= 1 Then
                                    If Len(St) > 2 Then
                                        PrintChat player(A).Name + " has been booted by " + player(b).Name + ": " + Mid$(St, 3), 15, Options.FontSize, 15
                                    Else
                                        PrintChat player(A).Name + " has been booted by " + player(b).Name + "!", 15, Options.FontSize, 15
                                    End If
                                Else
                                    If Len(St) > 2 Then
                                        PrintChat player(A).Name + " has been booted: " + Mid$(St, 3), 15, Options.FontSize, 15
                                    Else
                                        PrintChat player(A).Name + " has been booted!", 15, Options.FontSize, 15
                                    End If
                                End If
                            End If
                        End If
                        
                    Case 69 'Ban List
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            With frmList.lstBans
                                .AddItem CStr(A) + ": " + Mid$(St, 2)
                                .ItemData(.ListCount - 1) = A
                                .Visible = True
                            End With
                        Else
                            frmList.Show
                        End If
                        
                    Case 70 'Guild Data
                        If Len(St) >= 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                If Len(St) > 1 Then
                                    Guild(A).Name = Mid$(St, 2)
                                    Guild(A).Name = Crypt(Guild(A).Name)
                                Else
                                    Guild(A).Name = ""
                                End If
                            End If
                        End If
                        
                    Case 71 'Guild Dec. Data
                        If Len(St) = 3 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A <= 4 Then
                                With Character.GuildDeclaration(A)
                                    .Guild = Asc(Mid$(St, 2, 1))
                                    .Type = Asc(Mid$(St, 3, 1))
                                End With
                                UpdatePlayersColors
                            End If
                        End If
                        
                    Case 72 'Guild Change
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A > 0 Then
                                PrintChat "You are now a member of " + Chr$(34) + Guild(A).Name + Chr$(34), 15, Options.FontSize, 15
                            Else
                                If Character.Guild > 0 Then
                                    PrintChat "You are no longer a member of " + Chr$(34) + Guild(Character.Guild).Name + Chr$(34), 15, Options.FontSize, 15
                                End If
                            End If
                            DrawChat
                            frmMain.Refresh
                            Character.Guild = A
                            Character.GuildRank = 0
                            UpdatePlayersColors
                        End If
                        
                    Case 73 'Player Changed Guild
                        If Len(St) = 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            b = Asc(Mid$(St, 2, 1))
                            If A >= 1 Then
                                If player(A).Guild = Character.Guild And Character.Guild > 0 Then
                                    PrintChat player(A).Name + " is no longer a member of your guild.", 15, Options.FontSize, 5
                                End If
                                player(A).Guild = b
                                If b > 0 And b = Character.Guild Then
                                    PrintChat player(A).Name + " is now a member of your guild.", 15, Options.FontSize, 5
                                End If
                            End If
                            UpdatePlayerColor A
                        End If
                        
                    Case 74 'Guild Account Status
                        If Len(St) = 4 Then
                            A = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                            PrintChat "Your guild has " + CStr(A) + " gold in the bank.", 15, Options.FontSize, 5
                        ElseIf Len(St) = 8 Then
                            A = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                            b = Asc(Mid$(St, 5, 1)) * 16777216 + Asc(Mid$(St, 6, 1)) * 65536 + Asc(Mid$(St, 7, 1)) * 256& + Asc(Mid$(St, 8, 1))
                            PrintChat "Your guild owes " + CStr(A) + " gold.  This must be payed before " + CStr(CDate(b)) + " or your guild will be disbanded.  Type '/guild pay <amount>' to pay toward the debt.", 15, Options.FontSize, 5
                        End If
                        
                    Case 75 'Guild Deleted
                        If Len(St) = 1 Then
                            Select Case Asc(Mid$(St, 1, 1))
                                Case 0
                                    PrintChat "Your guild has failed to pay its debt in time and has been disbanded!", 15, Options.FontSize, 5
                                Case 1
                                    PrintChat "Your guild member count has fallen below three -- your guild has been disbanded!", 15, Options.FontSize, 5
                                Case 2
                                    PrintChat "Your guild has been disbanded!", 15, Options.FontSize, 5
                                Case 3
                                    PrintChat "Your guild has been disbanded by a god!", 15, Options.FontSize, 5
                            End Select
                            Character.Guild = 0
                            Character.GuildRank = 0
                            UpdatePlayersColors
                        End If
                        
                    Case 76 'Rank Changed
                        If Len(St) = 1 Then
                            Character.GuildRank = Asc(Mid$(St, 1, 1))
                            PrintChat "Your guild rank has been changed to " + Chr$(34) + Choose(Character.GuildRank + 1, "Initiate", "Member", "Officer", "Guildmaster") + Chr$(34) + ".", 15, Options.FontSize, 5
                        End If
                        
                    Case 77 'Invited to join guild
                        If Len(St) = 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            b = Asc(Mid$(St, 2, 1))
                            If A >= 1 And b >= 1 And player(b).Ignore = False Then
                                PrintChat "You have been invited to join the guild " + Chr$(34) + Guild(A).Name + Chr$(34) + " by " + player(b).Name + ".  If you wish to join, type /guild join.  It will cost 1000 gold to join this guild.", 15, Options.FontSize, 15
                            End If
                        End If
                        
                    Case 78 'View Guild Data
                        If Len(St) >= 12 Then
                            frmMain.picDrop.Visible = False
                            TempVar1 = Asc(Mid$(St, 1, 1))
                            If frmGuilds_Loaded = False Then Load frmGuilds
                            With frmGuilds
                                A = Asc(Mid$(St, 2, 1))
                                If A > 0 Then
                                    .lblHall = Hall(A).Name
                                Else
                                    .lblHall = "<none>"
                                End If
                                .lstDeclarations.Clear
                                For A = 0 To 9
                                    b = Asc(Mid$(St, 3 + 2 * A))
                                    If b > 0 Then
                                        If Asc(Mid$(St, 4 + 2 * A)) = 0 Then
                                            .lstDeclarations.AddItem "Declaration of Alliance with " + Guild(b).Name
                                        Else
                                            .lstDeclarations.AddItem "Declaration of War with " + Guild(b).Name
                                        End If
                                        .lstDeclarations.ItemData(.lstDeclarations.ListCount - 1) = A
                                    End If
                                Next A
                                .lstDeclarations2.Clear
                                For A = 0 To 9
                                    b = Asc(Mid$(St, 3 + 2 * A))
                                    If b > 0 Then
                                        If Asc(Mid$(St, 4 + 2 * A)) = 0 Then
                                            .lstDeclarations2.AddItem "Declaration of Alliance with " + Guild(b).Name
                                        Else
                                            .lstDeclarations2.AddItem "Declaration of War with " + Guild(b).Name
                                        End If
                                        .lstDeclarations2.ItemData(.lstDeclarations2.ListCount - 1) = A
                                    End If
                                Next A
                                .sclSymbol1.Value = Guild(TempVar1).Symbol1
                                .sclSymbol2.Value = Guild(TempVar1).Symbol2 ' = Asc(Mid$(St, 22, 1)) * 256 + Asc(Mid$(St, 23, 1))
                                .sclSymbol3.Value = Guild(TempVar1).Symbol3 ' ' = Asc(Mid$(St, 24, 1)) * 256 + Asc(Mid$(St, 25, 1))
                                Dim len1 As Long
                                Dim len2 As Long, renown As Long
                                
                                len1 = Asc(Mid$(St, 27, 1)) * 256 + Asc(Mid$(St, 28, 1))
                                len2 = Asc(Mid$(St, 29, 1)) * 256 + Asc(Mid$(St, 30, 1))
                                
                                renown = Asc(Mid$(St, 31, 1)) * 16777216 + Asc(Mid$(St, 32, 1)) * 65536 + Asc(Mid$(St, 33, 1)) * 256& + Asc(Mid$(St, 34, 1))
                                .lblFounded = CDate(renown)
                                renown = 0
                                .lblmotd = Mid$(St, 35, len1)
                                .txtInfo = Mid$(St, 35 + len1, len2)
                                .txtMOTD = Mid$(St, 35, len1)
                                
                                Dim item As cListItem
                                If Len(St) >= 35 + len1 + len2 Then
                                    GetSections3 Mid$(St, 35 + len1 + len2)
                                    .lblName = Guild(TempVar1).Name
                                    .lstMembers.ListItems.Clear
                                    .lstMembers2.Clear
                                    For A = 0 To 19
                                        If Len(Section(A + 1)) >= 14 Then
                                            b = Asc(Mid$(Section(A + 1), 1, 1)) - 1
                                            If b <= 3 Then
                                                '2, 3, and 4 are deaths, kills and renown
                                                Set item = .lstMembers.ListItems.Add(, , Mid$(Section(A + 1), 14))
                                                item.SubItems(1).Caption = Choose(b + 1, "Initiate", "Member", "Officer", "Guildmaster")
                                                item.SubItems(2).Caption = b
                                                len1 = Asc(Mid$(Section(A + 1), 2, 1)) * 256 + Asc(Mid$(Section(A + 1), 3, 1)) - 257
                                                len2 = Asc(Mid$(Section(A + 1), 4, 1)) * 256 + Asc(Mid$(Section(A + 1), 5, 1)) - 257
                                                renown = renown + (Asc(Mid$(Section(A + 1), 6, 1)) * 16777216 + Asc(Mid$(Section(A + 1), 7, 1)) * 65536 + Asc(Mid$(Section(A + 1), 8, 1)) * 256& + Asc(Mid$(Section(A + 1), 9, 1)) - 16843010)
                                                item.SubItems(3).Caption = Asc(Mid$(Section(A + 1), 6, 1)) * 16777216 + Asc(Mid$(Section(A + 1), 7, 1)) * 65536 + Asc(Mid$(Section(A + 1), 8, 1)) * 256& + Asc(Mid$(Section(A + 1), 9, 1)) - 16843010
                                                
                                        
                                                item.SubItems(4).Caption = CDate(Asc(Mid$(Section(A + 1), 10, 1)) * 16777216 + Asc(Mid$(Section(A + 1), 11, 1)) * 65536 + Asc(Mid$(Section(A + 1), 12, 1)) * 256& + Asc(Mid$(Section(A + 1), 13, 1)) - 16843010)
                                                .lstMembers2.AddItem Mid$(Section(A + 1), 14) + " - " + Choose(b + 1, "Initiate", "Member", "Officer", "Guildmaster")
                                                .lstMembers2.ItemData(.lstMembers2.ListCount - 1) = A
                                                
                                                If (Character.Name = Mid$(Section(A + 1), 14)) Then
                                                    .meIndex = A
                                                End If
                                                
                                            End If
                                        End If
                                    Next A
                                    .lstMembers.Columns(3).SortOrder = eSortOrderDescending
                                    .lstMembers.Columns(3).SortType = eLVSortString
                                    .lstMembers.ListItems.SortItems
                                    .lblRenown = renown
                                    .lblAverageRenown = Guild(TempVar1).AverageRenown
                                    .lblMembersTotal = Str(Guild(TempVar1).MembersOnline) & "/" & Str(Guild(TempVar1).members)
                                End If
                                If Character.Guild = TempVar1 And Character.GuildRank >= 2 Then
                                    'If Character.GuildRank = 3 Then
                                    '    .btnDisband.Enabled = True
                                    'Else
                                    '    .btnDisband.Enabled = False
                                    'End If
                                    'If .lstDeclarations.ListCount < 5 Then
                                    '    .btnAddDeclaration.Enabled = True
                                    'Else
                                    '    .btnAddDeclaration.Enabled = False
                                    'End If
                                    'If .lblHall = "<none>" Then
                                    '    .btnMoveOut.Enabled = False
                                    'Else
                                    '    .btnMoveOut.Enabled = True
                                    'End If
                                Else
                                    '.btnDisband.Enabled = False
                                    '.btnAddDeclaration.Enabled = False
                                    '.btnMoveOut.Enabled = False
                                End If
                                '.btnRemoveMember.Enabled = False
                                '.btnRemoveDeclaration.Enabled = False
                                '.btnRank(0).Enabled = False
                                '.btnRank(1).Enabled = False
                                '.btnRank(2).Enabled = False
                                '.btnOk.Enabled = True
                                If .nextTab = 0 Then .nextTab = 1
                                .GuildRank = Character.GuildRank
                                .playerGuild = Character.Guild
                                .ctab = 0
                                .setGuildTab
                                .lstEnabled = True
                                If .lstMembers2.ListCount >= .promoteidx Then .lstMembers2.ListIndex = .promoteidx
                                .Show
                            End With
                        End If
                        
                    Case 79 'Guild Chat
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                PrintChat player(A).Name + " -> Guild: " + Mid$(St, 2), 15, Options.FontSize, 5
                            End If
                        End If
                    
                    Case 80 'Created Guild
                        If Len(St) = 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            Character.Guild = A
                            Character.GuildRank = 3
                            If A > 0 Then
                                PrintChat "You have created a new guild called " + Chr$(34) + Guild(A).Name + Chr$(34) + ".  To invite other players to your guild, type '/guild invite <player>'.  You must get atleast two other players to join your guild today or your guild will be disbanded.", 15, Options.FontSize, 5
                            End If
                            DrawChat
                            frmMain.Refresh
                        End If
                        
                    Case 81 'Guild hall change
                        If Len(St) = 1 Then
                            If Asc(Mid$(St, 1, 1)) = 0 Then
                                PrintChat "Your guild now owns a hall!", 15, Options.FontSize, 5
                            Else
                                PrintChat "Your guild no longer owns a hall!", 15, Options.FontSize, 5
                            End If
                        End If
                    
                    Case 82 'Guild hall data
                        If Len(St) >= 1 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                If Len(St) >= 2 Then
                                    Hall(A).Name = Mid$(St, 2)
                                    Hall(A).Name = Crypt(Hall(A).Name)
                                Else
                                    Hall(A).Name = ""
                                End If
                                If frmList_Loaded = True Then
                                    frmList.lstHalls.List(A - 1) = CStr(A) + ": " + Hall(A).Name
                                End If
                            End If
                        End If
                        
                    Case 83 'Guild Hall Edit Data
                        If Len(St) = 13 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                If frmHall_Loaded = False Then Load frmHall
                                With frmHall
                                    .lblNumber = A
                                    .txtName = Hall(A).Name
                                    .txtPrice = CStr(Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1)))
                                    .txtUpkeep = CStr(Asc(Mid$(St, 6, 1)) * 16777216 + Asc(Mid$(St, 7, 1)) * 65536 + Asc(Mid$(St, 8, 1)) * 256& + Asc(Mid$(St, 9, 1)))
                                    b = Asc(Mid$(St, 10, 1)) * 256 + Asc(Mid$(St, 11, 1))
                                    If b < 1 Then b = 1
                                    If b > 5000 Then b = 5000
                                    .sclStartMap = b
                                    b = Asc(Mid$(St, 12, 1))
                                    If b > 11 Then b = 11
                                    .sclStartX = b
                                    b = Asc(Mid$(St, 13, 1))
                                    If b > 11 Then b = 11
                                    .sclStartY = b
                                    .Show
                                End With
                            End If
                        End If
                        
                    Case 84 'Guild Hall Info
                        If Len(St) = 10 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                b = Asc(Mid$(St, 2, 1))
                                If b > 0 Then
                                    PrintChat "Owned By: " + Guild(b).Name, 15, Options.FontSize, 15
                                Else
                                    PrintChat "This guild hall is not yet owned!", 15, Options.FontSize, 15
                                End If
                                A = Asc(Mid$(St, 3, 1)) * 16777216 + Asc(Mid$(St, 4, 1)) * 65536 + Asc(Mid$(St, 5, 1)) * 256& + Asc(Mid$(St, 6, 1))
                                PrintChat "Cost: " + CStr(A) + " gold coins", 15, Options.FontSize, 15
                                A = Asc(Mid$(St, 7, 1)) * 16777216 + Asc(Mid$(St, 8, 1)) * 65536 + Asc(Mid$(St, 9, 1)) * 256& + Asc(Mid$(St, 10, 1))
                                PrintChat "Upkeep: " + CStr(A) + " gold coins per day", 15, Options.FontSize, 15
                            End If
                        End If
                        
                    Case 85 'NPC Data
                        If Len(St) >= 5 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                With NPC(A)
                                    .Flags = Asc(Mid$(St, 2, 1))
                                    .Portrait = Asc(Mid$(St, 3, 1))
                                    .Sprite = Asc(Mid$(St, 4, 1))
                                    .Direction = Asc(Mid$(St, 5, 1))
                                    If Len(St) >= 2 Then
                                        .Name = Mid$(St, 6)
                                        .Name = Crypt(.Name)
                                    Else
                                        .Name = ""
                                    End If
                                    If frmList_Loaded = True Then
                                        'frmList.lstNPCs.List(A - 1) = CStr(A) + ": " + .Name
                                    End If
                                    If frmMapProperties_Loaded = True Then
                                        'frmMapProperties.cmbNPC.List(A) = CStr(A) + ": " + .Name
                                    End If
                                End With
                            End If
                        End If
                        
                    Case 86 'Buy Data
                        If Len(St) = 120 Then
                            NumShopItems = 0
                            For A = 0 To 9
                                With SaleItem(A)
                                    .GiveObject = GetInt(Mid$(St, 1 + A * 12, 2))
                                    .GiveValue = Asc(Mid$(St, 3 + A * 12, 1)) * 16777216 + Asc(Mid$(St, 4 + A * 12, 1)) * 65536 + Asc(Mid$(St, 5 + A * 12, 1)) * 256& + Asc(Mid$(St, 6 + A * 12, 1))
                                    .TakeObject = GetInt(Mid$(St, 7 + A * 12, 2))
                                    .TakeValue = Asc(Mid$(St, 9 + A * 12, 1)) * 16777216 + Asc(Mid$(St, 10 + A * 12, 1)) * 65536 + Asc(Mid$(St, 11 + A * 12, 1)) * 256& + Asc(Mid$(St, 12 + A * 12, 1))
                                    If .GiveObject >= 1 And .TakeObject >= 1 Then
                                        NumShopItems = NumShopItems + 1
                                    End If
                                End With
                            Next A
                            If NumShopItems > 0 Then
                                CurShopPage = 1
                                SetGUIWindow WINDOW_SHOP
                            End If
                            frmMain.picDrop.Visible = False
                        End If
                    
                    Case 87 'Edit NPC Data
                        If Len(St) >= 127 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                If frmNPC_Loaded = False Then Load frmNPC
                                b = NPC(A).Flags
                                For C = 0 To 2
                                    If ExamineBit(CByte(b), CByte(C)) = True Then
                                        frmNPC.chkFlag(C) = 1
                                    Else
                                        frmNPC.chkFlag(C) = 0
                                    End If
                                Next C
                                For b = 0 To 9
                                    With SaleItem(b)
                                        .GiveObject = Asc(Mid$(St, 2 + b * 12, 1)) * 256 + Asc(Mid$(St, 3 + b * 12, 1))
                                        .GiveValue = Asc(Mid$(St, 4 + b * 12, 1)) * 16777216 + Asc(Mid$(St, 5 + b * 12, 1)) * 65536 + Asc(Mid$(St, 6 + b * 12, 1)) * 256& + Asc(Mid$(St, 7 + b * 12, 1))
                                        .TakeObject = Asc(Mid$(St, 8 + b * 12, 1)) * 256 + Asc(Mid$(St, 9 + b * 12, 1))
                                        .TakeValue = Asc(Mid$(St, 10 + b * 12, 1)) * 16777216 + Asc(Mid$(St, 11 + b * 12, 1)) * 65536 + Asc(Mid$(St, 12 + b * 12, 1)) * 256& + Asc(Mid$(St, 13 + b * 12, 1))
                                    End With
                                    UpdateSaleItem b
                                Next b
                                '103
                                GetSections2 Mid$(St, 122)
                                With frmNPC
                                    .lblNumber = A
                                    .txtName = NPC(A).Name
                                    .sclPortrait = NPC(A).Portrait
                                    .sclSprite = NPC(A).Sprite
                                    .sclDirection = NPC(A).Direction
                                    .txtJoinText = Section(1)
                                    .txtLeaveText = Section(2)
                                    .txtSayText1 = Section(3)
                                    .txtSayText2 = Section(4)
                                    .txtSayText3 = Section(5)
                                    .txtSayText4 = Section(6)
                                    .txtSayText5 = Section(7)
                                    .Show 1
                                End With
                            End If
                        End If
                    
                    Case 88 'NPC Talks
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                PrintChat NPC(A).Name + " says, " + Chr$(34) + Mid$(St, 2) + Chr$(34), 7, Options.FontSize, 15
                            End If
                        End If
                        
                    Case 89 'Bank Balance
                        If Len(St) = 4 Then
                            'If Map.NPC >= 1 Then
                                'A = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                                'PrintChat NPC(Map.NPC).Name + " tells you, " + Chr$(34) + "You have " + CStr(A) + " gold coins in the bank." + Chr$(34), 7, Options.FontSize
                            'End If
                        End If
                        
                    Case 90 'God Chat
                        If Len(St) >= 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                PrintChat "<" + player(A).Name + ">: " + Mid$(St, 2), 11, Options.FontSize, 15
                            End If
                        End If
                        
                    Case 91 'Status Change
                        If Len(St) = 2 Then
                            A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                If A = Character.Index Then
                                    Character.Status = Asc(Mid$(St, 2, 1))
                                Else
                                    player(A).Status = Asc(Mid$(St, 2, 1))
                                End If
                                UpdatePlayerColor A
                            End If
                            DrawPartyNames
                        End If
                        
                    Case 92 'Edit Ban Data
                        If Len(St) >= 4 Then
                            If frmBan_Loaded = False Then Load frmBan
                            GetSections2 Mid$(St, 3)
                            With frmBan
                                .lblNumber = Asc(Mid$(St, 1, 1))
                                .sclUnban = Asc(Mid$(St, 2, 1))
                                .txtName = Section(1)
                                .txtBanner = Section(2)
                                .txtIP = Section(3)
                                .txtReason = Section(4)
                                .Show
                            End With
                        End If
                        
                    Case 93 'Gained exp
                        If Len(St) = 4 Then
                            b = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                            PrintChat "You have gained " + CStr(b) + " exp.", 12, Options.FontSize + 15
                        End If
                        
                    Case 94 'Edit Script Data
                        If Len(St) >= 3 Then
                            A = InStr(St, Chr$(0))
                            If A >= 1 Then
                                Load frmScript
                                With frmScript
                                    .lblName = Left$(St, A - 1)
                                    '.txtCode = Mid$(St, A + 1)
                                    .sciScript.Text = Mid$(St, A + 1)

                                   If Len(Mid$(St, A + 1)) = 0 Then
                                       St1 = .lblName
                                       If St1 Like "MAPSAY*" Then
                                           .sciScript.Text = "FUNCTION Main(Player AS LONG, Message AS STRING) AS LONG" + Chr$(13) + Chr$(10) + "   Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                                       ElseIf St1 Like "MAP*" Or St1 Like "MONSTERDIE*" Or St1 Like "JOINMAP*" Or St1 Like "PARTMAP*" Or St1 = "JOINGAME" Or St1 = "PARTGAME" Then
                                           .sciScript.Text = "SUB Main(Player AS LONG)" + Chr$(13) + Chr$(10) + "END SUB"
                                       ElseIf St1 Like "USEOBJ*" Or St1 Like "GETOBJ*" Or St1 Like "DROPOBJ*" Or St1 Like "BUYOBJ*" Or St1 Like "WITHDRAWOBJ*" Or St1 Like "MONSTERSEE*" Then
                                           .sciScript.Text = "FUNCTION Main(Player AS LONG) AS LONG" + Chr$(13) + Chr$(10) + "   Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                                        ElseIf St1 = "BROADCAST" Then
                                           .sciScript.Text = "FUNCTION Main(Player AS LONG, Message AS STRING) AS LONG" + Chr$(13) + Chr$(10) + "   Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                                       ElseIf St1 = "COMMAND" Then
                                           .sciScript.Text = "FUNCTION Main(Player as LONG, Command as STRING, Parm1 as STRING, Parm2 as STRING, Parm3 as STRING) AS LONG" + Chr$(13) + Chr$(10) + "Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                                       ElseIf St1 = "PLAYERDIE" Then
                                           .sciScript.Text = "FUNCTION Main(Player AS LONG, PK AS LONG, Killer AS LONG) AS LONG" + Chr$(13) + Chr$(10) + "   Main = Continue" + Chr$(13) + Chr$(10) + "END FUNCTION"
                                       ElseIf St1 Like "SPELL*" Then
                                           .sciScript.Text = "FUNCTION Main(Player AS LONG, TargetType AS LONG, Target AS LONG, X AS LONG, Y AS LONG) AS LONG" + Chr$(13) + Chr$(10) + "Main = 0" + Chr$(13) + Chr$(10) + "END FUNCTION"
                                        ElseIf St1 = "CLICKPLAYER" Or St1 = "CLICKMONSTER" Then
                                            .sciScript.Text = "SUB Main(Player AS LONG, Clicked AS LONG)" + Chr$(13) + Chr$(10) + "END SUB"
                                        End If
                                    End If
                                    frmScript.Show
                                End With
                            End If
                        End If
                        
                    Case 95 'User is Away
                    If Len(St) >= 1 Then
                        A = Asc(Mid$(St, 1, 1))
                            If A >= 1 Then
                                PrintChat player(A).Name + " is currently away! ' " + Mid$(St, 2) + "'", 14, Options.FontSize, 15
                            End If
                    End If
                    
                    Case 96 'Custom Sound/Music
                    If Len(St) >= 2 Then
                        A = Asc(Mid$(St, 1, 1))
                        b = Asc(Mid$(St, 2, 1))
                        Select Case A
                            Case 0
                                If Exists("Data/Sound/Sound" + CStr(b) + ".wav") Then
                                    PlayWav b
                                End If
                            Case 1
                                If b > 0 Then
                                    If Exists("Data/Music/" & CStr(b) & ".mp3") Then
                                        Sound_PlayStream b
                                    End If
                                Else
                                    Sound_StopStream
                                End If
                        End Select
                    End If
                    
                    Case 97 'New Guild Info
                        Select Case Asc(Mid$(St, 1, 1))
                            Case 0 'Already in a guild.
                                PrintChat "You are already in a guild.  If you would like to create a new guild, you must first leave this guild by typing '/guild leave'.", 14, Options.FontSize, 15
                            Case 1 'New Guild Show
                                frmNewGuild.Show
                            Case 2 'Guilds are Disabled
                                PrintChat "Guilds have been disabled.", 14, Options.FontSize, 15
                            Case 3 'Need to be at least level 10
                                PrintChat "You must be at least level 10 to join a guild!", 14, Options.FontSize, 15
                            Case 4 'Need to be at least level 15
                                PrintChat "You must be at least level 15 to create a guild!", 14, Options.FontSize, 15
                        End Select
                    
                    Case 98 'Repairing
                        Select Case Asc(Mid$(St, 1, 1))
                        
                            Case 1 'NPC Repair Display
                                A = Asc(Mid$(St, 5, 1)) * 16777216 + Asc(Mid$(St, 6, 1)) * 65536 + Asc(Mid$(St, 7, 1)) * 256& + Asc(Mid$(St, 8, 1)) 'Repair Ammount
                                b = Asc(Mid$(St, 2, 1)) 'Dur
                                C = QBColor(1)
                                D = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                                RepairCost = A
                                SetGUIWindow WINDOW_REPAIR
                        Case 2 'Done Repairing Object
                            'A = Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1))
                            PrintChat "Your " + Object(RepairObj.Object).Name + " is now at 100% durability. You repaired it successfully.", 14, Options.FontSize, 15
                            SetGUIWindow WINDOW_INVALID
                        End Select
                    
                    Case 99 'Projectiles
                        Select Case Asc(Mid$(St, 1, 1))
                            Case TT_TILE 'Tile Effect
                                CreateTileEffect Asc(Mid$(St, 2, 1)), Asc(Mid$(St, 3, 1)), Asc(Mid$(St, 4, 1)), Asc(Mid$(St, 5, 1)) * 256 + Asc(Mid$(St, 6, 1)), Asc(Mid$(St, 7, 1)), Asc(Mid$(St, 8, 1))
                                PlayWav Asc(Mid$(St, 9, 1))
                            Case TT_MONSTER 'Monster Effect
                                CreateMonsterEffect Asc(Mid$(St, 2, 1)), Asc(Mid$(St, 3, 1)), Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)), Asc(Mid$(St, 6, 1)), Asc(Mid$(St, 7, 1))
                                PlayWav Asc(Mid$(St, 8, 1))
                            Case TT_PLAYER 'Player Effect
                                If Asc(Mid$(St, 2, 1)) = Character.Index Then
                                    CreateCharacterEffect Asc(Mid$(St, 3, 1)), Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)), Asc(Mid$(St, 6, 1)), Asc(Mid$(St, 7, 1))
                                    PlayWav Asc(Mid$(St, 8, 1))
                                Else
                                    CreatePlayerEffect Asc(Mid$(St, 2, 1)), Asc(Mid$(St, 3, 1)), Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1)), Asc(Mid$(St, 6, 1)), Asc(Mid$(St, 7, 1))
                                    PlayWav Asc(Mid$(St, 8, 1))
                                End If
                        End Select
                    Case 100 'Floating Text
                        If Len(St) >= 2 Then
                            A = Int(Asc(Mid$(St, 1, 1)) / 16) * 32
                            b = (Asc(Mid$(St, 1, 1)) And 15) * 32 - 16
                            C = Asc(Mid$(St, 2, 1))
                            If C <= 15 Then
                                St1 = Mid$(St, 3)
                                C = QBColor(C)
                                C = D3DColorARGB(&HFF, (C And &HFF), (C \ &H100) And &HFF, (C \ &H10000) And &HFF)
                                Call FloatingText.Add(A, b, St1, C)
                            ElseIf C >= 16 Then
                                Select Case Int(C / 16)
                                    Case 1 'Miss
                                        Call FloatingText.Add(A, b, "Miss!", &HFFFFFFFF)
                                    Case 2 'Ineffective
                                        Call FloatingText.Add(A, b, "Ineffective!", &HFFFFFFFF)
                                    Case 3 'Ended
                                        Call FloatingText.Add(A, b, "Ended!", &HFFFF0000)
                                    Case 4 'Poison
                                        Call FloatingText.Add(A, b, "Poison!", &HFF00FF00)
                                    Case 5 'Exhaust
                                        Call FloatingText.Add(A, b, "Exhaust!", &HFF7F7F7F)
                                    Case 7 'Block
                                        Call FloatingText.Add(A, b, "Block!", &HFFFFFFFF)
                                End Select
                            End If
                        End If
                    Case Is > 100
                        ReceiveData2 PacketID, St
                End Select
            End If
            GoTo LoopRead
        End If
    End If

Exit Sub
Error_Handler:
Open App.Path + "/LOG.TXT" For Append As #1
    St1 = ""
    If Len(St) > 0 Then
        b = Len(St)
        For A = 1 To b
            St1 = St1 & Asc(Mid$(St, A, 1)) & "-"
        Next A
    End If
    Print #1, Err.Number & "/" & Err.Description & "/" & PacketID & "/" & Len(St) & "/" & St1
Close #1
Unhook
EndWinsock
End
End Sub

Public Sub ReceiveData2(ByVal PacketID As Long, St As String)

Dim A As Long, b As Long, C As Long, D As Long, E As Long, F As Long, G As Long, j As Long

Dim St1 As String

On Error GoTo Error_Handler
    Select Case PacketID
        Case 101 'FreezePlayer
            If Len(St) = 2 Then
                If Asc(Mid$(St, 1, 1)) = 1 Then
                    Character.Frozen = True
                    cX = Int(Asc(Mid$(St, 2, 1)) / 16)
                    cY = Asc(Mid$(St, 2, 1)) And 15
                    CX2 = cX ^ 2 + 5
                    CY2 = cY ^ 2 + 5
                Else
                    Character.Frozen = False
                End If
            End If
        Case 102 'Force Transition
            If Len(St) = 4 Then
                Select Case Asc(Mid$(St, 1, 1))
                    Case 0 'Fade Light
                        Transition 0, 0, 0, 0, 10
                    Case 1
                        Transition 1, 0, 0, 0, 10
                End Select
            End If
        Case 103 'Static Text
            If Len(St) >= 9 Then
                A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                b = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                C = Asc(Mid$(St, 5, 1))
                D = Asc(Mid$(St, 6, 1))
                E = Asc(Mid$(St, 7, 1))
                F = Asc(Mid$(St, 8, 1))
                St1 = Mid$(St, 9)
                If C < 16 Then
                    C = QBColor(C)
                    C = D3DColorARGB(&HFF, (C And &HFF), (C \ &H100) And &HFF, (C \ &H10000) And &HFF)
                    Call FloatingText.Add(A, b, St1, C, False, D, CStr(E), F)
                Else
                    NumUnzText = NumUnzText + 1
                    ReDim Preserve UnzText(NumUnzText - 1)
                    UnzText(NumUnzText - 1).x = A
                    UnzText(NumUnzText - 1).y = b
                    UnzText(NumUnzText - 1).Lifetime = D * 10
                    UnzText(NumUnzText - 1).Fade = 0
                    UnzText(NumUnzText - 1).Text = St1
                End If
            End If
        Case 104 'Update Party Status
            If Len(St) > 1 Then
                A = Asc(Mid$(St, 1, 1))
                Select Case A
                    Case 0 'Update Party Member HP
                        If Len(St) = 6 Then
                            b = Asc(Mid$(St, 2, 1))
                            With player(b)
                                .HP = GetInt(Mid$(St, 3, 2))
                                .Mana = GetInt(Mid$(St, 5, 2))
                            End With
                        End If
                        DrawPartyNames
                    Case 1 'Update Party Member Level
                        If Len(St) = 3 Then
                            b = Asc(Mid$(St, 2, 1))
                            player(b).Level = Asc(Mid$(St, 3, 1))
                        End If
                        DrawPartyNames
                    Case 2 'Update Player Location
                        If Len(St) = 4 Then
                            b = Asc(Mid$(St, 2, 1))
                            player(b).map = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                        Else
                            b = Asc(Mid$(St, 2, 1))
                            player(b).map = 0
                        End If
                        'DrawPartyNames
                    Case 3
                        If Len(St) = 3 Then
                            b = Asc(Mid$(St, 2, 1))
                            player(b).x = Asc(Mid$(St, 3, 1))
                            player(b).y = player(b).x And 15
                            player(b).x = player(b).x \ 16
                            doMiniMapDraw = True
                        End If
                End Select
            End If
        Case 105 'Change Party
            If Len(St) = 2 Then
                A = Asc(Mid$(St, 1, 1))
                b = Asc(Mid$(St, 2, 1))
                If Character.Party = b And Character.Party > 0 Then
                    PrintChat player(A).Name + " has joined the party!", 8, Options.FontSize, 6
                ElseIf Character.Party > 0 And Character.Party = player(A).Party And Character.Party = b Then
                    PrintChat player(A).Name + " has left the party!", 8, Options.FontSize, 6
                End If
                If A = Character.Index Then
                    Character.Party = b
                    frmMain.Refresh
                Else
                    player(A).Party = b
                End If
                DrawPartyNames
            End If
        Case 106 'Party Chat
            A = Asc(Mid$(St, 1, 1))
            PrintChat "[Party Chat] " & player(A).Name & " -> " & Mid$(St, 2), CHannelColors(3), Options.FontSize, 3, True
        Case 107 'Edit Prefix
            If Len(St) <= 27 Then
                A = Asc(Mid$(St, 1, 1))
                With frmPrefix
                    Load frmPrefix
                    .lblNumber = CStr(A)
                    .cmbModify.ListIndex = Asc(Mid$(St, 2, 1))
                    .sclMin.Value = Asc(Mid$(St, 3, 1))
                    .sclMax.Value = Asc(Mid$(St, 4, 1))
                    A = Asc(Mid$(St, 5, 1))
                    If ExamineBit(CByte(A), 1) Then
                        .chkType.Value = 1
                    Else
                        .chkType.Value = 0
                    End If
                    .cmbStrength1.ListIndex = Asc(Mid$(St, 6, 1))
                    .cmbStrength2.ListIndex = Asc(Mid$(St, 7, 1))
                    .cmbWeakness1.ListIndex = Asc(Mid$(St, 8, 1))
                    .cmbWeakness2.ListIndex = Asc(Mid$(St, 9, 1))
                    If Asc(Mid$(St, 10, 1)) > 127 Then
                        .sclIntensity = -(Asc(Mid$(St, 10, 1)) - 128)
                    Else
                        .sclIntensity = Asc(Mid$(St, 10, 1))
                    End If
                    If Asc(Mid$(St, 11, 1)) = 0 Then
                        .sclRadius = 1
                    Else
                        .sclRadius = Asc(Mid$(St, 11, 1))
                    End If
                    If Asc(Mid$(St, 12, 1)) = 0 Then
                        .sclRarity = 1
                    Else
                        .sclRarity = Asc(Mid$(St, 12, 1))
                    End If
                    If Len(St) >= 13 Then
                        .txtName = Mid$(St, 13)
                    Else
                        .txtName = ""
                    End If
                    .Show vbModal
                End With
            End If
        Case 108 'New Prefix
            If Len(St) >= 4 Then
                A = Asc(Mid$(St, 1, 1))
                If A > 0 Then
                    Prefix(A).Light.Intensity = Asc(Mid$(St, 2, 1))
                    Prefix(A).Light.Radius = Asc(Mid$(St, 3, 1))
                    Prefix(A).ModType = Asc(Mid$(St, 4, 1))
                    If Len(St) >= 5 Then
                        Prefix(A).Name = Trim$(Mid$(St, 5))
                    Else
                        Prefix(A).Name = ""
                    End If
                    Prefix(A).Name = Crypt(Prefix(A).Name)
                    If frmList_Loaded = True Then
                        With frmList.lstPrefix
                            .List(A - 1) = CStr(A) + ": " + Prefix(A).Name & " (" & ModString(Prefix(A).ModType) & ")"
                        End With
                    End If
                End If
            End If
        Case 109 'Set Stat
            If Len(St) >= 2 Then
                A = Asc(Mid$(St, 2, 1))
                If Asc(Mid$(St, 1, 1)) < 7 Then
                    b = Asc(Mid$(St, 3, 1))
                    If Asc(Mid$(St, 4, 1)) = 1 Then b = b * -1
                End If
                Select Case Asc(Mid$(St, 1, 1))
                    Case 1 'Strength
                        Character.strength = A
                        Character.StrengthMod = b
                    Case 2 'Agility
                        Character.Agility = A
                        Character.AgilityMod = b
                    Case 3 'Endurance
                        Character.Endurance = A
                        Character.EnduranceMod = b
                    Case 4 'Constitution
                        Character.Constitution = A
                        Character.ConstitutionMod = b
                    Case 5 'Wisdom
                        Character.Wisdom = A
                        Character.WisdomMod = b
                    Case 6 'Intelligence
                        Character.Intelligence = A
                        Character.IntelligenceMod = b
                    Case 7 'Attack Speed
                        Character.AttackSpeed = A * 30
                        If (Character.SkillLevels(SKILL_OPPORTUNIST) > 0) Then
                            If Character.Equipped(1).Object > 0 And Character.Equipped(2).Object > 0 Then
                                If Object(Character.Equipped(1).Object).Type = 1 And Object(Character.Equipped(2).Object).Type = 1 Then Character.AttackSpeed = Character.AttackSpeed - 150
                            End If
                        End If
                        If CurrentTab = tsStats2 Then SetTab tsStats2
                End Select
                DrawTrainBars
            End If
        Case 110 'hour
            If Len(St) = 1 Then
                A = Asc(Mid$(St, 1, 1))
                If A >= 1 And A <= 24 Then
                    World.Hour = A
                    World.Minute = 240
                    CalculateAmbientAlpha
                    'DrawSunDial
                    
                    If World.Hour = 7 And blnNight Then
                        blnNight = False
                        If CMap > 0 Then
                            PrintChat "It is now day time...", 13, Options.FontSize, 15
                            If ExamineBit(map.Flags(0), 1) = False Then ReDoLightSources
                        End If
                    End If
                    If World.Hour = 22 And blnNight = False Then
                        blnNight = True
                        If CMap > 0 Then
                            PrintChat "It is now night time...", 13, Options.FontSize, 15
                            If ExamineBit(map.Flags(0), 1) = False Then ReDoLightSources
                        End If
                    End If
                        
                    
                End If
            End If
        Case 111 'Skill
            If Len(St) > 0 Then
                A = Asc(Mid$(St, 1, 1))
                b = Asc(Mid$(St, 2, 1))
                St = Mid$(St, 3)
                RunSkill CByte(A), CByte(b), St
            End If
        Case 112 'status effect
            If Len(St) = 5 Then
                A = Asc(Mid$(St, 1, 1))
                b = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                If A = Character.Index Then
                    If Not (Character.StatusEffect And SE_BLIND) Then
                        If (b And SE_BLIND) Then
                            combatCounter = 30
                            drawTopBar
                        End If
                    End If
                    
                    Character.StatusEffect = b
                    If CurrentTab = tsStats2 Then SetTab tsStats2
                ElseIf A > 0 Then
                    player(A).StatusEffect = b
                End If
            End If
        Case 113 'Player Trade System
            Select Case Asc(Mid$(St, 1, 1))
                Case 0 'Invited to trade
                    If Character.Trading = False Then
                        PrintChat "You have been invited to trade with " & player(Asc(Mid$(St, 2, 1))).Name & ".  Type /PLAYERTRADE ACCEPT to accept this offer.", 14, Options.FontSize, 15
                        TradeData.Tradestate(0) = TRADE_STATE_INVITED
                        TradeData.player = Asc(Mid$(St, 2, 1))
                    End If
                Case 1 'Accepted trade
                    If Len(St) = 2 Then
                        If Character.Trading = False Then
                            A = Asc(Mid$(St, 2, 1))
                            If A > 0 And A <= MAXUSERS Then
                                TradeData.player = Asc(Mid$(St, 2, 1))
                                Character.Trading = True
                                TradeData.Tradestate(0) = TRADE_STATE_OPEN
                                TradeData.Tradestate(1) = TRADE_STATE_OPEN
                                For A = 1 To 10
                                    TradeData.Slot(A) = 0
                                    TradeData.YourObjects(A).Object = 0
                                    TradeData.TheirObjects(A).Object = 0
                                Next A
                                SetGUIWindow WINDOW_TRADE
                            End If
                        Else
                            PrintChat "You are already currently trading.", 7, Options.FontSize, 15
                            'Character.Trading = False
                        End If
                    End If
                Case 2 'Other Player Cancels Trade
                    If Character.Trading Then
                        Character.Trading = False
                        TradeData.player = 0
                        For A = 1 To 10
                            TradeData.Slot(A) = 0
                            TradeData.YourObjects(A).Object = 0
                            TradeData.TheirObjects(A).Object = 0
                        Next A
                        TradeData.Tradestate(0) = 0
                        SetGUIWindow WINDOW_INVALID
                    End If
                Case 3 'Add/Remove Other Players Trade Object
                    If Character.Trading = True Then
                        If TradeData.player > 0 And TradeData.player <= MAXUSERS Then
                            If Len(St) = 15 Then
                                A = Asc(Mid$(St, 2, 1)) 'Item Index
                                If A > 0 And A <= 10 Then
                                    With TradeData.TheirObjects(A)
                                        .Object = GetInt(Mid$(St, 3, 2))
                                        .Value = GetLong(Mid$(St, 5, 4))
                                        .Prefix = Asc(Mid$(St, 9, 1))
                                        .PrefixValue = Asc(Mid$(St, 10, 1))
                                        .Suffix = Asc(Mid$(St, 11, 1))
                                        .SuffixValue = Asc(Mid$(St, 12, 1))
                                        .Affix = Asc(Mid$(St, 13, 1))
                                        .AffixValue = Asc(Mid$(St, 14, 1))
                                        .ObjectColor = Asc(Mid$(St, 15, 1))
                                    End With
                                End If
                            End If
                        End If
                    Else
                        PrintChat "You are not currently trading.", 7, Options.FontSize, 15
                    End If
                Case 4 'Add/Remove YOUR trade object
                    If Character.Trading Then
                        If Len(St) = 16 Then
                            A = Asc(Mid$(St, 2, 1)) 'Slot
                            b = Asc(Mid$(St, 3, 1)) 'Inv Obj
                            If A > 0 And A <= 10 Then
                                If b >= 0 And b <= 20 Then
                                    TradeData.Slot(A) = b
                                    With TradeData.YourObjects(A)
                                        .Object = GetInt(Mid$(St, 4, 2))
                                        .Value = GetLong(Mid$(St, 6, 4))
                                        .Prefix = Asc(Mid$(St, 10, 1))
                                        .PrefixValue = Asc(Mid$(St, 11, 1))
                                        .Suffix = Asc(Mid$(St, 12, 1))
                                        .SuffixValue = Asc(Mid$(St, 13, 1))
                                        .Affix = Asc(Mid$(St, 14, 1))
                                        .AffixValue = Asc(Mid$(St, 15, 1))
                                        .ObjectColor = Asc(Mid$(St, 16, 1))
                                    End With
                                End If
                            End If
                        End If
                    End If
                Case 5 'Other player presses "accept"
                    If Character.Trading Then
                        TradeData.Tradestate(1) = TRADE_STATE_ACCEPTED
                    End If
                Case 6 'Re-Open trade
                    If Character.Trading Then
                        TradeData.Tradestate(1) = TRADE_STATE_OPEN
                        TradeData.Tradestate(0) = TRADE_STATE_OPEN
                    End If
            End Select
        Case 114 'Scan
            frmMain.lblScanData(0).Caption = Mid$(St, 1, InStr(1, St, Chr$(0) + Chr$(0), vbBinaryCompare))
            St = Mid$(St, InStr(1, St, Chr$(0)) + 1)
            frmMain.lblScanData(14).Caption = Mid$(St, 1, InStr(1, St, Chr$(1), vbBinaryCompare) - 1)
            St = Mid$(St, InStr(1, St, Chr$(1), vbBinaryCompare) + 1)
            frmMain.lblScanData(1).Caption = player(Asc(Mid$(St, 1, 1))).Name
            frmMain.lblScanData(2).Caption = Asc(Mid$(St, 2, 1)) & "/" & Asc(Mid$(St, 3, 1)) & "/" & Asc(Mid$(St, 4, 1)) & "/" & Asc(Mid$(St, 5, 1))
            frmMain.lblScanData(3).Caption = Asc(Mid$(St, 6, 1))
            frmMain.lblScanData(4).Caption = CStr(Asc(Mid$(St, 8, 1)) * 256 + Asc(Mid$(St, 9, 1))) + "/" + CStr(Asc(Mid$(St, 10, 1)) * 256 + Asc(Mid$(St, 11, 1)))
            frmMain.lblScanData(5).Caption = CStr(Asc(Mid$(St, 12, 1)) * 256 + Asc(Mid$(St, 13, 1))) + "/" + CStr(Asc(Mid$(St, 14, 1)) * 256 + Asc(Mid$(St, 15, 1)))
            frmMain.lblScanData(6).Caption = CStr(Asc(Mid$(St, 16, 1)) * 256 + Asc(Mid$(St, 17, 1))) + "/" + CStr(Asc(Mid$(St, 18, 1)) * 256 + Asc(Mid$(St, 19, 1)))
            frmMain.lblScanData(7).Caption = Asc(Mid$(St, 20, 1)) 'str
            frmMain.lblScanData(8).Caption = Asc(Mid$(St, 21, 1)) 'agility
            frmMain.lblScanData(9).Caption = Asc(Mid$(St, 22, 1)) 'Endurance
            frmMain.lblScanData(10).Caption = Asc(Mid$(St, 23, 1)) 'Wisdom
            frmMain.lblScanData(11).Caption = Asc(Mid$(St, 24, 1)) 'Constitution
            frmMain.lblScanData(12).Caption = Asc(Mid$(St, 25, 1)) 'Intelligence
            frmMain.lblScanData(13).Caption = Asc(Mid$(St, 26, 1)) * 100 'Attack Speed
            If Len(St) > 26 Then
                St = Mid$(St, 27)
                For b = 1 To 25
                    PlayerScanItem(b).objNum = Asc(Mid$(St, (b - 1) * 10 + 1)) * 256 + Asc(Mid$(St, (b - 1) * 10 + 1))
                    If PlayerScanItem(b).objNum > 0 Then
                        PlayerScanItem(b).ObjValue = (Asc(Mid$(St, (b - 1) * 10 + 3, 1)) * 16777216 + Asc(Mid$(St, (b - 1) * 10 + 4, 1)) * 65536 + Asc(Mid$(St, (b - 1) * 10 + 5, 1)) * 256& + Asc(Mid$(St, (b - 1) * 10 + 6, 1)))
                        PlayerScanItem(b).Prefix = Asc(Mid$(St, (b - 1) * 10 + 7, 1))
                        PlayerScanItem(b).PrefixVal = Asc(Mid$(St, (b - 1) * 10 + 8, 1))
                        PlayerScanItem(b).Suffix = Asc(Mid$(St, (b - 1) * 10 + 9, 1))
                        PlayerScanItem(b).SuffixVal = Asc(Mid$(St, (b - 1) * 10 + 10, 1))
                    Else
                        frmMain.lblScannedItem(b - 1).Caption = "<empty>"
                        PlayerScanItem(b).ObjValue = 0
                        PlayerScanItem(b).Prefix = 0
                        PlayerScanItem(b).PrefixVal = 0
                        PlayerScanItem(b).Suffix = 0
                        PlayerScanItem(b).SuffixVal = 0
                        A = A + 1
                    End If
                Next b
            End If
            
            frmMain.picPlayerScan.Visible = True
        Case 115 'Buff
            If Len(St) = 2 Then
                A = Asc(Mid$(St, 1, 1))
                If A > 0 And A <= MAXUSERS Then
                    If A = Character.Index Then
                        Character.buff = Asc(Mid$(St, 2, 1))
                    Else
                        player(A).buff = Asc(Mid$(St, 2, 1))
                    End If
                End If
            End If
        Case 116 'Update MaxHP and MaxMana and MaxEnergy
            If Len(St) = 6 Then
                With Character
                    .MaxHP = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                    DrawHP
                    .MaxMana = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                    DrawMana
                    .MaxEnergy = Asc(Mid$(St, 5, 1)) * 256 + Asc(Mid$(St, 6, 1))
                    DrawEnergy
                End With
            End If
        Case 117 'Storage
            A = Asc(Mid$(St, 1, 1))
            Select Case A
                Case 0
                    If Len(St) = 2603 Then
                        St = Mid$(St, 2)
                        For A = 1 To 10
                            For C = 1 To 20
                                With Character.Storage(A, C)
                                    b = ((A - 1) * 260) + ((C - 1) * 13)
                                    .Object = Asc(Mid$(St, b + 1, 1)) * 256 + Asc(Mid$(St, b + 2, 1))
                                    .Value = Asc(Mid$(St, b + 3, 1)) * 16777216 + Asc(Mid$(St, b + 4, 1)) * 65536 + Asc(Mid$(St, b + 5, 1)) * 256& + Asc(Mid$(St, b + 6, 1))
                                    .Prefix = Asc(Mid$(St, b + 7, 1))
                                    .PrefixValue = Asc(Mid$(St, b + 8, 1))
                                    .Suffix = Asc(Mid$(St, b + 9, 1))
                                    .SuffixValue = Asc(Mid$(St, b + 10, 1))
                                    .Affix = Asc(Mid$(St, b + 11, 1))
                                    .AffixValue = Asc(Mid$(St, b + 12, 1))
                                    .ObjectColor = Asc(Mid$(St, b + 13, 1))
                                End With
                            Next C
                        Next A
                        b = Asc(Mid$(St, 2601, 1))
                        C = Asc(Mid$(St, 2602, 2))
                        If C > 0 And C <= 10 Then Character.NumStoragePages = C
                        If b > 0 And b <= 10 Then
                            Character.CurStoragePage = b
                            SetGUIWindow WINDOW_STORAGE
                        End If
                    End If
                Case 1
                    A = Asc(Mid$(St, 2, 1))
                    With Character.Storage(Character.CurStoragePage, A)
                        .Object = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                        .Value = Asc(Mid$(St, 5, 1)) * 16777216 + Asc(Mid$(St, 6, 1)) * 65536 + Asc(Mid$(St, 7, 1)) * 256& + Asc(Mid$(St, 8, 1))
                        .Prefix = Asc(Mid$(St, 9, 1))
                        .PrefixValue = Asc(Mid$(St, 10, 1))
                        .Suffix = Asc(Mid$(St, 11, 1))
                        .SuffixValue = Asc(Mid$(St, b + 12, 1))
                        .Affix = Asc(Mid$(St, 13, 1))
                        .AffixValue = Asc(Mid$(St, 14, 1))
                        .ObjectColor = Asc(Mid$(St, 15, 1))
                    End With
                Case 2
                    If Len(St) = 2 Then
                        A = Asc(Mid$(St, 2, 1))
                        With Character.Storage(Character.CurStoragePage, A)
                            .Object = 0
                            .Value = 0
                            .Prefix = 0
                            .PrefixValue = 0
                            .Suffix = 0
                            .SuffixValue = 0
                            .Affix = 0
                            .AffixValue = 0
                            .ObjectColor = 0
                        End With
                    End If
                Case 3
                    StorageOpen = False
                    Character.CurStoragePage = 0
                    If CurrentWindow = WINDOW_STORAGE Then
                        SetGUIWindow WINDOW_INVALID
                    End If
                Case 4
                    If Len(St) = 2 Then
                        Character.CurStoragePage = Asc(Mid$(St, 2, 1))
                    End If
            End Select
        Case 118 'PlayerList
            If Len(St) > 0 Then
                For A = 1 To Len(St)
                    PrintChat player(Asc(Mid$(St, A, 1))).Name, 15, Options.FontSize
                Next A
            End If
        Case 119 'Change Dur
            If Len(St) = 5 Then
                A = Asc(Mid$(St, 1, 1))
                If A >= 1 And A <= 25 Then
                    If A <= 20 Then
                        With Character.Inv(A)
                            .Value = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                            If CurInvObj = A Then DrawCurInvObj
                        End With
                    Else
                        With Character.Equipped(A - 20)
                            .Value = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                            If CurInvObj = A Then DrawCurInvObj
                        End With
                    End If
                End If
            End If
        Case 120 'Weather
            If Len(St) >= 1 Then
                Select Case Asc(Mid$(St, 1, 1))
                    Case 0 'All
                        World.Rain = GetInt(Mid$(St, 2, 2))
                        If map.Raining = 0 Or ((map.Flags(1) And MAP_RAINING) = 0) Then
                            InitRain World.Rain
                        End If
                        World.Snow = GetInt(Mid$(St, 4, 2))
                        If map.Snowing = 0 Or ((map.Flags(1) And MAP_SNOWING) = 0) Then
                            InitSnow World.Snow
                        End If
                        World.Fog = Asc(Mid$(St, 6, 1))
                        If map.Fog < 1 Or map.Fog > 31 Then
                            If World.Fog > 0 Then
                                InitFog World.Fog, True
                            Else
                                ClearFog False
                            End If
                        End If
                        World.FlickerDark = Asc(Mid$(St, 7, 1))
                        World.FlickerLength = Asc(Mid$(St, 8, 1))
                    Case 1 'Rain
                        World.Rain = GetInt(Mid$(St, 2, 2))
                        If map.Raining = 0 Or ((map.Flags(1) And MAP_RAINING) = 0) Then
                            InitRain World.Rain
                        End If
                    Case 2 'Snow
                        World.Snow = GetInt(Mid$(St, 2, 2))
                        If map.Snowing = 0 Or ((map.Flags(1) And MAP_SNOWING) = 0) Then
                            InitSnow World.Snow
                        End If
                    Case 3 'Fog
                        World.Fog = Asc(Mid$(St, 2, 1))
                        If map.Fog < 1 Or map.Fog > 31 Then
                            If World.Fog > 0 Then
                                InitFog World.Fog, True
                            ElseIf CurFog > 0 Then
                                ClearFog False
                            End If
                        End If
                    Case 4 'Flicker Freq
                        World.FlickerDark = Asc(Mid$(St, 2, 2))
                    Case 5 'Flicker Length
                        World.FlickerLength = Asc(Mid$(St, 2, 2))
                        
                End Select
            End If
        Case 121 'Party Message
            If Len(St) = 1 Then
                A = Asc(Mid$(St, 1, 1))
                If player(A).Ignore = False Then
                    PrintChat "You have been invited to join a party by " + player(A).Name + ".  Type '/party join' to join.", 15, Options.FontSize, 15
                End If
            End If
        Case 122 'Freemap Message
            If Len(St) = 2 Then
                PrintChat "The lowest avalible map number is: " + CStr(Asc(Mid$(St, 1, 1)) * 256& + Asc(Mid$(St, 2, 1))), 15, Options.FontSize, 15
            End If
        Case 123 'Gain Skill EXP
            If Len(St) = 5 Then
                A = Asc(Mid$(St, 1, 1))
                If A <> SKILL_INVALID And A <= MAX_SKILLS Then
                    Character.SkillEXP(A) = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                End If
            ElseIf Len(St) = 3 Then
                A = Asc(Mid$(St, 2, 1))
                If A <> SKILL_INVALID And A <= MAX_SKILLS Then
                    Character.SkillLevels(A) = Asc(Mid$(St, 3, 1))
                    UpdateSkills
                End If
            ElseIf Len(St) = 0 Then
                For A = 1 To MAX_SKILLS
                    Character.SkillLevels(A) = 0
                    Character.skillPoints = (Character.Level) * 3
                    UpdateSkills
                Next A
            End If
        Case 124 'Widget String
            A = Len(St)
            If Len(St) > 4 Then
                CreateMenuFromString (St)
            End If
        Case 125 'Shoot Projectile
            If Len(St) = 6 Then
                A = Asc(Mid$(St, 1, 1)) 'Index
                b = Asc(Mid$(St, 2, 1)) * 256& + Asc(Mid$(St, 3, 1)) 'Effect
                C = Asc(Mid$(St, 4, 1)) 'X/Y
                D = Asc(Mid$(St, 5, 1)) 'Direction
                E = Asc(Mid$(St, 6, 1)) 'Speed
                If D >= 0 And D <= 3 Then
                    If Projectiles.Count = 31 Then
                        Dim PJ As clsProjectile
                        For Each PJ In Projectiles
                            If PJ.Owner = 255 Then
                                Projectiles.Remove PJ.Key
                                Exit For
                            End If
                        Next PJ
                    End If
                    If A <> Character.Index Then
                        Projectiles.Add b, CByte(D), E / 255, C \ 16, C And 15, CByte(A), 0, 0, 0, 0, 0
                    Else
                        If Len(Character.CurProjectile.Key) = 0 Then
                            Character.CurProjectile.Key = Projectiles.Add(b, CByte(D), E / 255, C \ 16, C And 15, CByte(A), 0, 0, 0, 0, 0).Key
                        End If
                    End If
                End If
            ElseIf Len(St) = 11 Then
                A = Asc(Mid$(St, 1, 1)) 'Index
                b = Asc(Mid$(St, 2, 1)) * 256& + Asc(Mid$(St, 3, 1)) 'Effect
                C = Asc(Mid$(St, 4, 1)) 'X/Y
                D = Asc(Mid$(St, 5, 1)) 'Direction
                E = Asc(Mid$(St, 6, 1)) 'Speed
                If D >= 0 And D <= 3 Then

                    If A = 255 Then ' map proj
                        If Asc(Mid$(St, 7, 1)) = 0 And Asc(Mid$(St, 9, 1)) = 1 Then
                            F = Asc(Mid$(St, 8, 1))
                            
                            For Each PJ In Projectiles
                                If PJ.Intensity = F Then Projectiles.Remove PJ.Key
                            Next PJ
                        End If
                        Projectiles.Add b, CByte(D), E / 255, C \ 16, C And 15, CByte(A), Asc(Mid$(St, 7, 1)), Asc(Mid$(St, 8, 1)), Asc(Mid$(St, 9, 1)), Asc(Mid$(St, 10, 1)), Asc(Mid$(St, 11, 1))
                    Else
                        If Projectiles.Count = 31 Then

                            For Each PJ In Projectiles
                                If PJ.Owner = 255 Then
                                    Projectiles.Remove PJ.Key
                                    Exit For
                                End If
                            Next PJ
                        End If
                        If A <> Character.Index Then
                            Projectiles.Add b, CByte(D), E / 255, C \ 16, C And 15, CByte(A), Asc(Mid$(St, 7, 1)), Asc(Mid$(St, 8, 1)), Asc(Mid$(St, 9, 1)), Asc(Mid$(St, 10, 1)), Asc(Mid$(St, 11, 1))
                        Else
                            If Len(Character.CurProjectile.Key) = 0 Then
                                Character.CurProjectile.Key = Projectiles.Add(b, CByte(D), E / 255, C \ 16, C And 15, CByte(A), Asc(Mid$(St, 7, 1)), Asc(Mid$(St, 8, 1)), Asc(Mid$(St, 9, 1)), Asc(Mid$(St, 10, 1)), Asc(Mid$(St, 11, 1))).Key
                            End If
                        End If
                    End If
                End If
                ReDoLightSources
            End If
        Case 126 'Set Player NPC Name Color
            If Len(St) = 2 Then
                A = Asc(Mid$(St, 1, 1)) 'NPC
                b = Asc(Mid$(St, 2, 1)) 'Color
                FloatingText.Add -1, -1, NPC(A).Name, NPCStatusColors(b), False, 255, NPC(A).Name
            End If
        Case 127 'Particle Effect
            A = GetInt(Mid$(St, 1, 2)) 'X
            b = GetInt(Mid$(St, 3, 2)) 'Y
            C = Asc(Mid$(St, 5, 1))    'Type
            D = Asc(Mid$(St, 6, 1)) 'Particle
            E = Asc(Mid$(St, 7, 1)) 'Red
            F = Asc(Mid$(St, 8, 1)) 'Green
            G = Asc(Mid$(St, 9, 1)) 'Blue
            j = GetInt(Mid$(St, 10, 2))
            Select Case C
                Case 3 'Explosion
                    ParticleEngineF.Add A, b, 3, E, F, G, j, 7.5, 5, GetInt(Mid$(St, 12, 2)), 8, D, 0, 0, 0
                Case 6 'Sucking Effect
                    ParticleEngineF.Add A, b, 6, E, F, G, j, 0.05, 0, GetInt(Mid$(St, 12, 2)), 8, D, GetInt(Mid$(St, 14, 2)), TT_NO_TARGET, 0
                Case 7 'Smoke
                    ParticleEngineF.Add A, b, 7, E, F, G, j, 0, -1, GetInt(Mid$(St, 12, 2)), 10, D, GetInt(Mid$(St, 14, 2)), 0, 0
                Case 8 'Fire
                    ParticleEngineF.Add A, b, 8, E, F, G, j, 0, -1, GetInt(Mid$(St, 12, 2)), 10, D, GetInt(Mid$(St, 14, 2)), 0, 0
            End Select
        Case 128 'Cur Inv Callback
            If CurInvObj > 0 And CurInvObj <= 20 Then
                SendSocket Chr$(78) + Chr$(CurInvObj)
            Else
                SendSocket Chr$(78) + Chr$(0)
            End If
        Case 129 'Light Data
            If Len(St) >= 8 Then
                A = Asc(Mid$(St, 1, 1))
                If A >= 1 And A <= 255 Then
                    With Lights(A)
                        .Red = Asc(Mid$(St, 2, 1))
                        .Green = Asc(Mid$(St, 3, 1))
                        .Blue = Asc(Mid$(St, 4, 1))
                        .Intensity = Asc(Mid$(St, 5, 1))
                        .Radius = Asc(Mid$(St, 6, 1))
                        .MaxFlicker = Asc(Mid$(St, 7, 1))
                        .FlickerRate = Asc(Mid$(St, 8, 1))
                        If Len(St) > 8 Then
                            .Name = Crypt(Mid$(St, 9))
                        End If
                    End With
                    ReDoLightSources
                End If
            End If
        Case 130 'modified monster tint
            If Len(St) = 5 Then
                A = Asc(Mid$(St, 1, 1))
                If A <= 9 Then
                    With map.Monster(A)
                        .R = Asc(Mid$(St, 2, 1))
                        .G = Asc(Mid$(St, 3, 1))
                        .b = Asc(Mid$(St, 4, 1))
                        .alpha = Asc(Mid$(St, 5, 1))
                        
                        If .R > Monster(.Monster).Red Then .R = Monster(.Monster).Red
                        If .G > Monster(.Monster).Green Then .G = Monster(.Monster).Green
                        If .b > Monster(.Monster).Blue Then .b = Monster(.Monster).Blue
                        If .alpha > Monster(.Monster).alpha Then .alpha = Monster(.Monster).alpha
                        
                        
                        
                    End With
                End If
            End If
         Case 131 'Variable Sized Floating Text
            If Len(St) >= 2 Then
                A = Int(Asc(Mid$(St, 1, 1)) / 16) * 32
                b = (Asc(Mid$(St, 1, 1)) And 15) * 32 - 16
                C = Asc(Mid$(St, 2, 1))
                D = Asc(Mid$(St, 3, 1))
                E = Asc(Mid$(St, 4, 1))
                If C <= 15 Then
                    St1 = Mid$(St, 5)
                    C = QBColor(C)
                    C = D3DColorARGB(&HFF, (C And &HFF), (C \ &H100) And &HFF, (C \ &H10000) And &HFF)
                    Call FloatingText.Add(A, b, St1, C, True, E, "", D)
                End If
            End If
        Case 132 'Update Monster HP Bar
            If Len(St) = 3 Then
                
                A = Asc(Mid$(St, 1, 1)) 'Monster
                b = Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1)) 'hp
                map.Monster(A).HP = b
            End If
        Case 133 'modified player tint
            If Len(St) = 5 Then
                A = Asc(Mid$(St, 1, 1))
                If Character.Index <> A Then
                    With player(A)
                        .Red = Asc(Mid$(St, 2, 1))
                        .Green = Asc(Mid$(St, 3, 1))
                        .Blue = Asc(Mid$(St, 4, 1))
                        .alpha = Asc(Mid$(St, 5, 1))
                    End With
                Else
                        Character.Red = Asc(Mid$(St, 2, 1))
                        Character.Green = Asc(Mid$(St, 3, 1))
                        Character.Blue = Asc(Mid$(St, 4, 1))
                        Character.alpha = Asc(Mid$(St, 5, 1))
                End If
            End If
        Case 134 'ping retrieved
            If Len(St) = 0 Then
                PrintChat "Ping(ms): " & Str(GetTickCount - PingTime), 7, Options.FontSize, 15
            End If
        Case 135 'Guild Text With Returns
            If Len(St) >= 2 Then
                b = Asc(Mid$(St, 1, 1))
                A = b
                St = Mid$(St, 2)
                While (InStr(1, St, Chr$(10)))
                 
                
                PrintChat Mid$(St, 1, InStr(2, St, Chr$(10)) - 2), A, Options.FontSize, 4
                'PrintChat Mid$(St, 1, InStr(2, St, Chr$(10))), CHannelColors(4), Options.FontSize, 4, True
                St = Mid$(St, InStr(1, St, Chr$(10)) + 1)
                A = b
                Wend
                PrintChat Mid$(St, 1), A, Options.FontSize, 4
            End If
        Case 136 'guild list data
            If Len(St) >= 11 Then
                A = Asc(Mid$(St, 1, 1))
                With (Guild(A))
                    .members = Asc(Mid$(St, 2, 1))
                    .MembersOnline = Asc(Mid$(St, 3, 1))
                    .AverageRenown = Asc(Mid$(St, 4, 1)) * 16777216 + Asc(Mid$(St, 5, 1)) * 65536 + Asc(Mid$(St, 6, 1)) * 256& + Asc(Mid$(St, 7, 1))
                    .Symbol1 = Asc(Mid$(St, 8, 1))
                    .Symbol2 = Asc(Mid$(St, 9, 1))
                    .Symbol3 = Asc(Mid$(St, 10, 1))
                    .hallNum = Asc(Mid$(St, 11, 1))
                    
                End With
            End If
        Case 137 'map trap set
            If Len(St) >= 3 Then
                With map.Trap(Asc(Mid$(St, 3, 1)))
                    .Counter = (Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1))) / 1000
                    .Created = GetTickCount
                    .x = Asc(Mid$(St, 1, 1))
                    .y = Asc(Mid$(St, 2, 1))
                End With
            End If
        Case 138 'You hit monster ALTERNATE
            If Len(St) = 5 Then
                C = Asc(Mid$(St, 1, 1))
                D = Asc(Mid$(St, 2, 1))
                A = Asc(Mid$(St, 3, 1))
                If A >= 10 Then
                    C = 1
                    A = A - 10
                End If
                If A >= 0 And A <= 9 Then
                    If map.Monster(A).Monster > 0 Then
                        b = Asc(Mid$(St, 4, 1)) * 256 + Asc(Mid$(St, 5, 1))
                        map.Monster(A).HP = map.Monster(A).HP - b
                        If C = 0 And D = 0 Then FloatingText.Add map.Monster(A).x * 32, map.Monster(A).y * 32 - 16, CStr(b), &HFFFF0000
                        'CAttack = 5
                        Select Case C
                            Case 0
                                PlayWav 11 + Int(Rnd * 3)
                            Case 1
                                PlayWav 14
                        End Select
                    End If
                End If
            ElseIf Len(St) = 4 Then
                C = Asc(Mid$(St, 1, 1)) 'hit/miss
                A = Asc(Mid$(St, 2, 1))
                b = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                With map.Monster(A)
                    If C = 0 Then FloatingText.Add .x * 32, .y * 32 - 32, CStr(b), &HFFFF0000
                    .HP = .HP - b
                    End With
            End If
        Case 139 'Fish!
            A = Asc(Mid$(St, 1, 1))
                
            b = Asc(Mid$(St, 2, 1))
            For C = 1 To 10
                If map.Fish(C).TimeStamp <= GetTickCount Then
                    map.Fish(C).TimeStamp = GetTickCount + 2100
                    map.Fish(C).x = A
                    map.Fish(C).y = b
                    Exit For
                End If
            Next C
            ParticleEngineF.Add A * 32 + 16, b * 32 + 16, 5, 19, 26, 53, 90, 0.4, 0, 40, 2, 4, 5, TT_NO_TARGET, 0
            'CreateTileEffect A, b, 99, 110, 6, 0
        Case 140 'Mine out of stuff
            A = Asc(Mid$(St, 1, 1))
            b = Asc(Mid$(St, 2, 1))
            mapFGChanged = True
            
            mapChangedBg(A, b) = True
        
            With map.Tile(A, b)
                .AttData(3) = 0
            End With
            
            
            
            
        Case 141 'show object stats
            A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
            b = Asc(Mid$(St, 3, 1)) * 16777216 + Asc(Mid$(St, 4, 1)) * 65536 + Asc(Mid$(St, 5, 1)) * 256& + Asc(Mid$(St, 6, 1))
            SetTab tsInventory
            DrawCurInvObj (A), b
            
        Case 142 'cancel particle effect
            A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
            b = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
            C = Asc(Mid$(St, 5, 1)) * 256 + Asc(Mid$(St, 6, 1))
                Dim tmpPS As clsParticleSource
                For Each tmpPS In ParticleEngineB
                    With tmpPS
                        If .x = A And .y = b Then
                            .LifeLeft = C
                        End If
                    End With
                Next
            
                For Each tmpPS In ParticleEngineF
                With tmpPS
                    If .x = A And .y = b Then
                        .LifeLeft = C
                    End If
                End With
                Next
        Case 143 'edit scripts
            If frmList_Loaded = False Then Load frmList
                With frmList
                    A = 0
                    .lstScripts.Clear
                    While Len(St) > 1
                    scripts(A) = (Mid$(St, 1, InStr(1, St, " ")))
                    A = A + 1
                    .lstScripts.AddItem (Mid$(St, 1, InStr(1, St, " ")))
                    St = Mid$(St, InStr(1, St, " ") + 1)
                    Wend
                
                
                    .lstObjects.Visible = False
                    .lstMonsters.Visible = False
                    .lstNPCs.Visible = False
                    .lstBans.Visible = False
                    .lstHalls.Visible = False
                    .lstPrefix.Visible = False
                    .lstLights.Visible = False
                    .lstScripts.Visible = True
                    .btnOk.Caption = "Edit"
                    .Show
                 End With
        Case 144 'kill sounds
            For A = 0 To 31
                FSOUND_StopSound A
            Next A
    End Select
    
Exit Sub
Error_Handler:
Open App.Path + "/LOG.TXT" For Append As #1
    St1 = ""
    If Len(St) > 0 Then
        b = Len(St)
        For A = 1 To b
            St1 = St1 & Asc(Mid$(St, A, 1)) & "-"
        Next A
    End If
    Print #1, Err.Number & "/RecClieData/" & Err.Description & "/" & PacketID & "/" & Len(St) & "/" & St1
Close #1
Unhook
EndWinsock
End
End Sub

