Attribute VB_Name = "modReadClientData"
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

Sub ReadClientData(Index As Long)
On Error GoTo Error_Handler
    Dim St As String, SocketData As String, PacketLength As Long, PacketID As Long, Crypt As Long, Crypt2 As Byte
    Dim A As Long, B As Long, C As Long, D As Long, E As Long, F As Long, G As Long, H As Long, i As Long, J As Long, k As Long, L As Byte, M As Long
    Dim ST1 As String
    Dim mapNum As Long
    St = "UNDEFINED"
    If player(Index).Leaving = True Then Exit Sub
    With player(Index)
        mapNum = .map
        SocketData = .SocketData + Receive(.Socket)
        .LastMsg = GetTickCount
LoopRead:

    .FloodTimer = .FloodTimer + 1
    If .FloodTimer > 150 And player(Index).Access = 0 Then
        BootPlayer Index, 0, "Flooding"
    End If
        If Len(SocketData) >= 3 Then
            PacketLength = GetInt(Mid$(SocketData, 1, 2))
            If PacketLength >= 3072 And .Access < 10 Then
                Hacker Index, "C.1"
                Exit Sub
            End If
            If Len(SocketData) - 2 >= PacketLength Then
                St = Mid$(SocketData, 3, PacketLength)
                SocketData = Mid$(SocketData, PacketLength + 3)
                If PacketLength > 0 Then
                    PacketID = Asc(Mid$(St, 1, 1))
                    'packetTypeUsage(PacketID) = packetTypeUsage(PacketID) + 1
                    If Len(St) > 1 Then
                        St = Mid$(St, 2)
                        Crypt = 0: A = 0: B = 0: ST1 = "": C = 0
                        B = Asc(Right$(St, 1))
                        St = Mid$(St, 1, Len(St) - 1)
                        If .PacketsSent = 255 Then .PacketsSent = 0
                        If PacketLength < 50 Then                       'No need to perform check on very large packets
                            ST1 = St
                            ST1 = Chr2(PacketID) + St
                            For A = 1 To Len(ST1)
                                Crypt = Crypt + Asc(Mid$(ST1, A, 1)) + 7
                            Next A
                            Crypt = Crypt Mod 256
                            Crypt2 = Crypt Xor (.PacketsSent + 1)
                            Crypt2 = Not Crypt2
                            If Not (B = Crypt2) Then
                                'Hacker Index, "Invalid Packet " & PacketID
                                Dim st2 As String
                                For A = 1 To Len(ST1)
                                    st2 = st2 & Asc(Mid$(ST1, A, 1)) & " "
                                Next A
                                
                                SendToGods Chr2(56) + Chr2(15) + .Name + " has sent an invalid packet!|||" & ST1 & "|||" & St + "|||" & st2
                                .PacketsSent = .PacketsSent + 1
                                GoTo LoopRead
                                'Exit Sub
                            End If
                            'Close #1
                        End If
                    Else
                        Hacker Index, "C'MON DO THE FUNKY CHICKEN C'MON!"
                        Exit Sub
                    End If
                    .PacketsSent = .PacketsSent + 1
                    Select Case .Mode
                        Case modePlaying
                            Select Case PacketID
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                              Case 7 'Move
                                  AddToMovePlayerMoveQueue Index, St
                              Case Is >= 26
                                  Call ReceiveData2(Index, PacketID, St)
                              
                              Case 13 'Exit Map
                                  If Len(St) = 1 Then
                                      Select Case Asc(Mid$(St, 1, 1))
                                          Case 0
                                              If map(mapNum).ExitUp > 0 Then
                                                  If .y = 0 Then
                                                      Partmap Index
                                                      .map = map(mapNum).ExitUp
                                                      .y = 11
                                                      JoinMap Index
                                                  Else
                                                      'Hacker Index, "A.23.1"
                                                      Partmap Index
                                                      .map = mapNum
                                                      JoinMap Index
                                                  End If
                                              Else
                                                  'Hacker Index, "A.23.1"
                                                  Partmap Index
                                                  .map = mapNum
                                                  JoinMap Index
                                              End If
                                          Case 1
                                              If map(mapNum).ExitDown > 0 Then
                                                  If .y = 11 Then
                                                      Partmap Index
                                                      .map = map(mapNum).ExitDown
                                                      .y = 0
                                                      JoinMap Index
                                                  Else
                                                      'Hacker Index, "A.23.1"
                                                      Partmap Index
                                                      .map = mapNum
                                                      JoinMap Index
                                                  End If
                                              Else
                                                  'Hacker Index, "A.23.1"
                                                  Partmap Index
                                                  .map = mapNum
                                                  JoinMap Index
                                              End If
                                          Case 2
                                              If map(mapNum).ExitLeft > 0 Then
                                                  If .x = 0 Then
                                                      Partmap Index
                                                      .map = map(mapNum).ExitLeft
                                                      .x = 11
                                                      JoinMap Index
                                                  Else
                                                      'Hacker Index, "A.23.1"
                                                      Partmap Index
                                                      .map = mapNum
                                                      JoinMap Index
                                                  End If
                                              Else
                                                  'Hacker Index, "A.23.1"
                                                  Partmap Index
                                                  .map = mapNum
                                                  JoinMap Index
                                              End If
                                          Case 3
                                              If map(mapNum).ExitRight > 0 Then
                                                  If .x = 11 Then
                                                      Partmap Index
                                                      .map = map(mapNum).ExitRight
                                                      .x = 0
                                                      JoinMap Index
                                                  Else
                                                      'Hacker Index, "A.23.1"
                                                      Partmap Index
                                                      .map = mapNum
                                                      JoinMap Index
                                                  End If
                                              Else
                                                  'Hacker Index, "A.23.1"
                                                  Partmap Index
                                                  .map = mapNum
                                                  JoinMap Index
                                              End If
                                      End Select
                                  Else
                                      Hacker Index, "A.23"
                                  End If
                              
                              Case 24 'Shoot Projectile
                                  If Len(St) = 2 Then
                                      If .ProjectileDamage = 0 And .ProjectileType = 0 Then
                                          A = Asc(Mid$(St, 2, 1))
                                          If A >= 0 And A <= 3 Then
                                              If .ProjectileDamage = 0 Then
                                                  If .Equipped(EQ_WEAPON).Object > 0 Then
                                                      If Object(.Equipped(EQ_WEAPON).Object).Type = 10 Then
                                                          If .Equipped(EQ_AMMO).Object > 0 Then
                                                              If Object(.Equipped(EQ_AMMO).Object).Type = 11 Then
                                                                  If .Equipped(EQ_AMMO).Value > 0 Then
                                                                      If Object(.Equipped(EQ_WEAPON).Object).Data(3) = Object(.Equipped(EQ_AMMO).Object).Data(4) Then
                                                                          .ProjectileDamage = PlayerDamage(Index)
                                                                          .ProjectileType = PT_PHYSICAL
                                                                          .ProjectileX = .x
                                                                          .ProjectileY = .y
                                                                          .Equipped(EQ_AMMO).Value = .Equipped(EQ_AMMO).Value - 1
                                                                          If .Equipped(EQ_AMMO).Value > 0 Then
                                                                              SendSocket Index, Chr2(17) + Chr2(20 + EQ_AMMO) + QuadChar$(.Equipped(EQ_AMMO).Value)
                                                                          Else
                                                                              .Equipped(EQ_AMMO).Object = 0
                                                                              SendSocket Index, Chr2(18) + Chr2(20 + EQ_AMMO)
                                                                          End If
                                                                          SendToMapAllBut mapNum, Index, Chr2(125) + Chr2(Index) + DoubleChar$(256) + Chr2(.x * 16 + .y) + Chr2(A) + Chr2(127)
                                                                      End If
                                                                  End If
                                                              End If
                                                          End If
                                                      End If
                                                  End If
                                              End If
                                          End If
                                      End If
                                  End If
                                  
                              Case 8 'Pick up map object
                                  If player(Index).Trading = False Then
                                      If Len(St) = 0 Then
                                          For A = 0 To 49
                                              C = map(mapNum).Object(A).Object
                                              If C > 0 Then
                                                  If map(mapNum).Object(A).x = .x And map(mapNum).Object(A).y = .y Then
                                                      If ExamineBit(Object(map(mapNum).Object(A).Object).Flags, 7) Then GoTo nexta
                                                      Parameter(0) = Index
                                                      Parameter(1) = map(mapNum).Object(A).Value
                                                      Parameter(2) = A
                                                      If RunScript("GETOBJ" + CStr(C)) = 0 Then
                                                          If Object(C).Type = 6 Or Object(C).Type = 11 Then
                                                              'Money or ammo
                                                              J = Object(C).Data(0)
                                                              B = FindInvObject(Index, C, False)
                                                              If B = 0 Then
                                                                  B = FreeInvNum(Index)
                                                                  If B > 0 Then .Inv(B).Value = 0
                                                                  E = 2
                                                              Else
                                                                  While (.Inv(B).Value = J And B <> 20)
                                                                      H = FindInvObject(Index, C, False, B + 1)
                                                                      If H = B Then
                                                                          B = FreeInvNum(Index)
                                                                          If B > 0 Then .Inv(B).Value = 0
                                                                      Else
                                                                          B = H
                                                                      End If
                                                                  Wend
                                                                  E = 1
                                                              End If
                                                          Else
                                                              B = FreeInvNum(Index)
                                                              If B > 0 Then .Inv(B).Value = 0
                                                              E = 0
                                                          End If
                                                          Parameter(1) = A
                                                          Parameter(2) = map(mapNum).Object(A).Value
                                                          If RunScript("GETOBJ") = 0 Then
                                                            If map(mapNum).Object(A).Object > 0 Then
getAnother:
                                                                If B > 0 Then
                                                                    With .Inv(B)
                                                                        .Object = C
                                                                        If E = 1 Then
                                                                            If CDbl(.Value) + CDbl(map(mapNum).Object(A).Value) > 2147483647# Then
                                                                                D = 2147483647
                                                                            Else
                                                                                
                                                                                If J > 0 Then
                                                                                    If .Value + map(mapNum).Object(A).Value > J Then
                                                                                        i = J - .Value
                                                                                        .Value = J
                                                                                        map(mapNum).Object(A).Value = map(mapNum).Object(A).Value - i
                                                                                        D = .Value
                                                                                        SendSocket Index, Chr2(17) + Chr2(B) + DoubleChar(CInt(C)) + QuadChar(D) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)  'New Inv Obj
                                                                                        B = FreeInvNum(Index)
                                                                                        If B > 0 Then .Value = 0
                                                                                        If B = 0 Then
                                                                                            With map(mapNum).Object(A)
                                                                                                SendToMap player(Index).map, Chr2(14) + Chr2(A) + DoubleChar(CInt(.Object)) + Chr2(.x) + Chr2(.y) + Chr2(0) + Chr2(0) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + QuadChar(IIf(.TimeStamp - GetTickCount > 0, .TimeStamp - GetTickCount, 0)) + Chr2(Abs(.deathObj)) + Chr2(.ObjectColor)
                                                                                            End With
                                                                                        End If
                                                                                        If map(mapNum).Object(A).Value > J Then
                                                                                            E = 2
                                                                                        Else
                                                                                            E = 0
                                                                                        End If
                                                                                        GoTo getAnother
                                                                                    Else
                                                                                        D = .Value + map(mapNum).Object(A).Value
                                                                                    End If
                                                                                Else
                                                                                    D = .Value + map(mapNum).Object(A).Value
                                                                                End If
                                                                            End If
                                                                        Else
                                                                            If E = 2 Then
                                                                                 If J > 0 Then
                                                                                    If .Value > 0 Then .Value = 0
                                                                                    If .Value + map(mapNum).Object(A).Value > J Then
                                                                                        i = J - .Value
                                                                                        .Value = J
                                                                                        D = .Value
                                                                                        map(mapNum).Object(A).Value = map(mapNum).Object(A).Value - i
                                                                                        SendSocket Index, Chr2(17) + Chr2(B) + DoubleChar(CInt(C)) + QuadChar(D) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)  'New Inv Obj
                                                                                        B = FreeInvNum(Index)
                                                                                        If B > 0 Then .Value = 0
                                                                                        If B = 0 Then
                                                                                            With map(mapNum).Object(A)
                                                                                                SendToMap player(Index).map, Chr2(14) + Chr2(A) + DoubleChar(CInt(.Object)) + Chr2(.x) + Chr2(.y) + Chr2(0) + Chr2(0) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + QuadChar(IIf(.TimeStamp - GetTickCount > 0, .TimeStamp - GetTickCount, 0)) + Chr2(Abs(.deathObj)) + Chr2(.ObjectColor)
                                                                                            End With
                                                                                        End If
                                                                                        If map(mapNum).Object(A).Value > J Then
                                                                                            E = 2
                                                                                        Else
                                                                                            E = 0
                                                                                        End If
                                                                                        GoTo getAnother
                                                                                    Else
                                                                                        D = .Value + map(mapNum).Object(A).Value
                                                                                    End If
                                                                                Else
                                                                                    D = .Value + map(mapNum).Object(A).Value
                                                                                End If
                                                                            Else
                                                                                D = map(mapNum).Object(A).Value
                                                                            End If
                                                                        End If
                                                                        .Value = D
                                                                        .prefix = map(mapNum).Object(A).prefix
                                                                        .prefixVal = map(mapNum).Object(A).prefixVal
                                                                        .suffix = map(mapNum).Object(A).suffix
                                                                        .SuffixVal = map(mapNum).Object(A).SuffixVal
                                                                        .Affix = map(mapNum).Object(A).Affix
                                                                        .AffixVal = map(mapNum).Object(A).AffixVal
                                                                        .ObjectColor = map(mapNum).Object(A).ObjectColor
                                                                        For D = 0 To 3
                                                                          .Flags(D) = map(mapNum).Object(A).Flags(D)
                                                                        Next D
                                                                    End With
                                                                    map(mapNum).Object(A).Object = 0
                                                                    SendToMap mapNum, Chr2(15) + Chr2(A) 'Erase Map Obj
                                                                    With .Inv(B)
                                                                        
                                                                        SendSocket Index, Chr2(17) + Chr2(B) + DoubleChar(CInt(C)) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)  'New Inv Obj
                                                                    End With
                                                                Else
                                                                    SendSocket Index, Chr2(16) + Chr2(1) 'Inv Full
                                                                    Exit For
                                                                End If
                                                              End If
                                                            End If
                                                      End If
                                                  End If
                                              End If
nexta:
                                          Next A
                                      Else
                                          Hacker Index, "A.17"
                                      End If
                                  Else
                                      SendSocket Index, Chr2(113) + Chr2(6) + Chr2(4) 'can't do that while trading
                                  End If
                              
                              Case 23 'Projectile Collision
                                  If Len(St) >= 1 Then
                                      A = Asc(Mid$(St, 1, 1))
                                      Select Case A
                                          Case TT_PLAYER
                                              If Len(St) = 2 Then
                                                  If .ProjectileDamage > 0 Then
                                                      B = Asc(Mid$(St, 2, 1))
                                                      If B > 0 And B <= MaxUsers Then
                                                          If (.ProjectileX - player(B).x) <= 0 Or (.ProjectileY - player(B).y) <= 0 Then
                                                              'If LOS(.Map, .ProjectileX, .ProjectileY, Player(B).X, Player(B).Y, 0) Then
                                                                  If CanAttackPlayer(Index, B) Then
                                                                      C = .ProjectileDamage
                                                                      If C < 0 Then C = 0
                                                                      AttackPlayer Index, B, C, True, IIf(.ProjectileType = PT_MAGIC, True, False), True, True, True, False, False
                                                                  End If
                                                              'End If
                                                          End If
                                                      End If
                                                  End If
                                              End If
                                              .ProjectileDamage = 0
                                              .ProjectileType = 0
                                          Case TT_MONSTER
                                              If Len(St) = 2 Then
                                                  B = Asc(Mid$(St, 2, 1))
                                                  If B >= 0 And B <= 9 Then
                                                      If map(mapNum).monster(B).monster > 0 Then
                                                          If (.ProjectileX - map(mapNum).monster(B).x) <= 1 Or (.ProjectileY - map(mapNum).monster(B).y) <= 1 Then
                                                              If ExamineBit(map(mapNum).Flags(0), 5) = False Then
                                                                  If map(mapNum).Tile(player(Index).x, player(Index).y).Att <> 21 Then
                                                                      C = .ProjectileDamage
                                                                      If .ProjectileType = PT_PHYSICAL Then
                                                                          C = C - monster(map(mapNum).monster(B).monster).Armor
                                                                      ElseIf .ProjectileType = PT_MAGIC Then
                                                                          C = C - monster(map(mapNum).monster(B).monster).MagicResist
                                                                      End If
                                                                      If C < 0 Then C = 0
                                                                      AttackMonster Index, B, C, IIf(.ProjectileType = PT_MAGIC, True, False), True
                                                                  End If
                                                              End If
                                                          End If
                                                      End If
                                                  End If
                                              End If
                                              .ProjectileDamage = 0
                                              .ProjectileType = 0
                                          Case TT_TILE
                                              If Len(St) = 3 Then
                                                  A = Asc(Mid$(St, 2, 1))
                                                  B = Asc(Mid$(St, 3, 1))
                                                  Parameter(0) = Index
                                                  RunScript ("PROJ" & mapNum & "_" & A & "_" & B)
                                                  .ProjectileDamage = 0
                                                  .ProjectileType = 0
                                              End If
                                      End Select
                                      .ProjectileDamage = 0
                                      .ProjectileType = 0
                                  End If
                              
                              Case 10 'Use Object
                                  If player(Index).Trading = False Then
                                      If Len(St) = 1 Then
                                          A = Asc(Mid$(St, 1, 1))
                                          If A >= 1 And A <= 20 Then
                                              If .Inv(A).Object > 0 Then
                                                  B = .Inv(A).Object
                                                  k = 0
                                                  Parameter(0) = Index
                                                  Parameter(1) = A
                                                  Parameter(2) = .Inv(A).Object
                                                  Parameter(3) = .Inv(A).Value
                                                  k = RunScript("USEOBJ")
                                                  If k = 0 Then
                                                      If (Object(.Inv(A).Object).Class = 0) Or ((2 ^ (.Class - 1)) And Object(.Inv(A).Object).Class) Then
                                                          If .Level >= Object(.Inv(A).Object).MinLevel Then
                                                                Parameter(0) = Index
                                                                Parameter(1) = A
                                                                Parameter(2) = .Inv(A).Object
                                                                k = RunScript("USEOBJ" + CStr(.Inv(A).Object))
                                                              'If K = 0 Then
                                                                'Parameter(0) = Index
                                                                'Parameter(1) = A
                                                                'Parameter(2) = .Inv(A).Object
                                                                'Parameter(3) = .Inv(A).Value
                                                                'K = RunScript("USEOBJ")
                                                              'End If
                                                              If k = 0 Then
                                                                  If .Inv(A).Object > 0 Then
                                                                      Select Case Object(.Inv(A).Object).Type
                                                                          Case 1, 10 'Weapon
                                                                              If .Equipped(EQ_DUALWIELD).Object > 0 Then
                                                                                  If Object(.Equipped(EQ_DUALWIELD).Object).Type = 1 Then
                                                                                      If Not ExamineBit(Object(.Inv(A).Object).Type, 5) Then 'Both are dual wield arr!!
                                                                                          'unequip this item
                                                                                          'EquipObject Index, 0, EQ_DUALWIELD, False
                                                                                          RemoveObject Index, EQ_DUALWIELD
                                                                                          SendToMapAllBut mapNum, Index, Chr2(8) + Chr2(Index) + Chr2(.x) + Chr2(.y) + Chr2(.D) + Chr2(.red) + Chr2(.green) + Chr2(.blue) + Chr2(.alpha) + DoubleChar(.Equipped(1).Object) + DoubleChar(.Equipped(2).Object) + DoubleChar(.Equipped(3).Object) + DoubleChar(.Equipped(4).Object)
                                                                                      End If
                                                                                  End If
                                                                              End If
                                                                              If ExamineBit(Object(.Inv(A).Object).Flags, 3) Then 'two handed
                                                                                  If .Equipped(EQ_SHIELD).Object > 0 Then
                                                                                       SendSocket Index, Chr2(16) + Chr2(45)
                                                                                      B = 0
                                                                                  Else
                                                                                      B = EQ_WEAPON
                                                                                  End If
                                                                              ElseIf .Equipped(EQ_WEAPON).Object > 0 And ExamineBit(Object(.Inv(A).Object).Flags, 5) Then 'dual wield
                                                                                  If ExamineBit(Object(.Equipped(EQ_WEAPON).Object).Flags, 5) Then 'Weapon is duel wield and equipped weapon is duel wield
                                                                                      B = EQ_DUALWIELD
                                                                                  Else
                                                                                      B = EQ_WEAPON
                                                                                  End If
                                                                              Else
                                                                                  B = EQ_WEAPON
                                                                              End If
                                                                          Case 2, 11 'Shield
                                                                              If .Equipped(EQ_WEAPON).Object > 0 Then
                                                                                  If ExamineBit(Object(.Equipped(EQ_WEAPON).Object).Flags, 3) Then 'Two-Handed
                                                                                      SendSocket Index, Chr2(16) + Chr2(46)
                                                                                      B = 0
                                                                                  Else
                                                                                      B = EQ_SHIELD
                                                                                  End If
                                                                              Else
                                                                                  B = EQ_SHIELD
                                                                              End If
                                                                              C = 0
                                                                          Case 3 'Armor
                                                                              B = EQ_ARMOR
                                                                              C = 0
                                                                          Case 4 'Helmet
                                                                              B = EQ_HELMET
                                                                              C = 0
                                                                          Case 5 'Potion
                                                                              B = Object(.Inv(A).Object).Data(1) * 256& + Object(.Inv(A).Object).Data(2)
                                                                              C = Object(.Inv(A).Object).Data(3) * 256& + Object(.Inv(A).Object).Data(4)
                                                                              If B > 1000 Then B = 1000
                                                                              If C > 1000 Then C = 1000
                                                                              Select Case Object(.Inv(A).Object).Data(0)
                                                                                  Case 0 'Gives HP
                                                                                      If CLng(.HP) + B < .MaxHP Then
                                                                                          .HP = .HP + B
                                                                                      Else
                                                                                          .HP = .MaxHP
                                                                                      End If
                                                                                      SendSocket Index, Chr2(46) + DoubleChar(CLng(.HP))
       
                                                                                          SendToPartyAllBut .Party, Index, Chr2(104) + Chr2(0) + Chr2(Index) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                                                                                           SendToGods Chr2(104) + Chr2(0) + Chr2(Index) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                                                                                          SendToGuildAllBut Index, CLng(.Guild), Chr2(104) + Chr2(0) + Chr2(Index) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
         
                                                                                      CreateFloatingText .map, .x, .y, 10, CStr(B)
                                                                                  Case 1 'Takes HP
                                                                                      If CLng(.HP) - B > 0 Then
                                                                                          .HP = .HP - B
                                                                                      Else
                                                                                          .HP = 0
                                                                                      End If
                                                                                      SendSocket Index, Chr2(46) + DoubleChar(CLng(.HP))
    
                                                                                          SendToPartyAllBut .Party, Index, Chr2(104) + Chr2(0) + Chr2(Index) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                                                                                          SendToGods Chr2(104) + Chr2(0) + Chr2(Index) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                                                                                          SendToGuildAllBut Index, CLng(.Guild), Chr2(104) + Chr2(0) + Chr2(Index) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
                    
                                                                                      CreateFloatingText .map, .x, .y, 12, CStr(B)
                                                                                  Case 2 'Gives Mana
                                                                                      If CLng(.Mana) + B < .MaxMana Then
                                                                                          .Mana = .Mana + B
                                                                                      Else
                                                                                          .Mana = .MaxMana
                                                                                      End If
                                                                                      SendSocket Index, Chr2(48) + DoubleChar(CLng(.Mana))
                                                                                      CreateFloatingText .map, .x, .y, 3, CStr(B)
                                                                                  Case 3 'Takes Mana
                                                                                      If CLng(.Mana) - B > 0 Then
                                                                                          .Mana = .Mana - B
                                                                                      Else
                                                                                          .Mana = 0
                                                                                      End If
                                                                                      SendSocket Index, Chr2(48) + DoubleChar(CLng(.Mana))
                                                                                      CreateFloatingText .map, .x, .y, 2, CStr(B)
                                                                                  Case 4 'Gives Energy
                                                                                      If CLng(.Energy) + B < .MaxEnergy Then
                                                                                          .Energy = .Energy + B
                                                                                      Else
                                                                                          .Energy = .MaxEnergy
                                                                                      End If
                                                                                      SendSocket Index, Chr2(47) + DoubleChar(CInt(.Energy))
                                                                                      CreateFloatingText .map, .x, .y, 7, CStr(B)
                                                                                  Case 5 'Takes Energy
                                                                                      If CLng(.Energy) - B > 0 Then
                                                                                          .Energy = .Energy - B
                                                                                      Else
                                                                                          .Energy = 0
                                                                                      End If
                                                                                      SendSocket Index, Chr2(47) + DoubleChar(CInt(.Energy))
                                                                                      CreateFloatingText .map, .x, .y, 8, CStr(B)
                                                                                  Case 6 'Cures Poison
                                                                                      RemoveStatusEffect Index, SE_POISON
                                                                                  Case 7 'Causes Poison
                                                                                      SetStatusEffect Index, SE_POISON
                                                                                      .StatusData(SE_POISON).Data(0) = B
                                                                                      .StatusData(SE_POISON).timer = C
                                                                                  Case 8 'Causes Regen
                                                                                      SetStatusEffect Index, SE_REGENERATION
                                                                                      .StatusData(SE_REGENERATION).Data(1) = B
                                                                                      .StatusData(SE_REGENERATION).timer = C
                                                                                  Case 9 'Cures Mute
                                                                                      RemoveStatusEffect Index, SE_MUTE
                                                                                  Case 10 'Causes Mute
                                                                                      SetStatusEffect Index, SE_MUTE
                                                                                      .StatusData(SE_MUTE).timer = C
                                                                              End Select
                                                                              B = 0
                                                                              C = 1
                                                                          Case 7 'Key
                                                                              Select Case .D
                                                                                  Case 0 'Up
                                                                                      C = .x
                                                                                      D = CLng(.y) - 1
                                                                                  Case 1 'Down
                                                                                      C = .x
                                                                                      D = .y + 1
                                                                                  Case 2 'Left
                                                                                      C = CLng(.x) - 1
                                                                                      D = .y
                                                                                  Case 3 'Right
                                                                                      C = .x + 1
                                                                                      D = .y
                                                                              End Select
                                                                              If C >= 0 And C <= 11 And D >= 0 And D <= 11 Then
                                                                                  If map(mapNum).Tile(C, D).Att = 3 And (map(mapNum).Tile(C, D).AttData(3) * 256& + map(mapNum).Tile(C, D).AttData(0)) = .Inv(A).Object Then
                                                                                      E = FreeMapDoorNum(mapNum)
                                                                                      If E >= 0 Then
                                                                                          With map(mapNum).Door(E)
                                                                                              .Att = 3
                                                                                              .x = C
                                                                                              .y = D
                                                                                              .t = GetTickCount
                                                                                              .Wall = map(mapNum).Tile(C, D).WallTile
                                                                                          End With
                                                                                          map(mapNum).Tile(C, D).Att = 0
                                                                                          map(mapNum).Tile(C, D).WallTile = 0
                                                                                          SendToMap mapNum, Chr2(36) + Chr2(E) + Chr2(C) + Chr2(D)
                                                                                          If Object(.Inv(A).Object).Data(0) = 0 Then
                                                                                              C = 1
                                                                                          Else
                                                                                              C = 0
                                                                                          End If
                                                                                      End If
                                                                                  Else
                                                                                      C = 0
                                                                                  End If
                                                                              Else
                                                                                  C = 0
                                                                              End If
                                                                              B = 0
                                                                            
                                                                          Case 8 'Ring
                                                                              B = EQ_RING
                                                                              C = 0
                                                                          
                                                                          Case 9 'Guild Deed
                                                                              If .Guild = 0 Then
                                                                                  If .Level >= 15 Then
                                                                                      B = 0
                                                                                      C = 1
                                                                                      SendSocket Index, Chr2(97) + Chr2(1)
                                                                                  Else
                                                                                      SendSocket Index, Chr2(97) + Chr2(4)
                                                                                  End If
                                                                              Else
                                                                                  B = 0
                                                                                  C = 0
                                                                                  SendSocket Index, Chr2(97) + Chr2(0)
                                                                              End If
                                                                          
                                                                          Case Else
                                                                              B = 0
                                                                              C = 0
                                                                              If k <> 0 Then SendSocket Index, Chr2(16) + Chr2(8) 'You cannot use that
                                                                      End Select
                                                                      If B > 0 And B <= 5 Then
                                                                          'Equip Item
                                                                          EquipObject Index, A, B
                                                                            With player(Index)
                                                                                SendToMapAllBut mapNum, Index, Chr2(8) + Chr2(Index) + Chr2(.x) + Chr2(.y) + Chr2(.D) + Chr2(.red) + Chr2(.green) + Chr2(.blue) + Chr2(.alpha) + DoubleChar(.Equipped(1).Object) + DoubleChar(.Equipped(2).Object) + DoubleChar(.Equipped(3).Object) + DoubleChar(.Equipped(4).Object)
                                                                            End With
                                                                      End If
                                                                      If C > 0 Then
                                                                          'Destroy Item
                                                                          .Inv(A).Object = 0
                                                                          SendSocket Index, Chr2(18) + Chr2(A) 'Remove inv object
                                                                      End If
                                                                  End If
                                                              Else
                                                                  SendSocket Index, Chr2(16) + Chr2(3) 'No such object
                                                              End If
                                                          Else
                                                              SendSocket Index, Chr2(16) + Chr2(39) 'invalid level
                                                          End If
                                                      Else
                                                          SendSocket Index, Chr2(16) + Chr2(38) 'invalid class
                                                      End If
                                                  End If
                                              End If
                                          Else
                                              Hacker Index, "A.19"
                                          End If
                                      Else
                                          Hacker Index, "A.20"
                                      End If
                                  Else
                                      SendSocket Index, Chr2(113) + Chr2(6) + Chr2(4) 'can't do that while trading
                                  End If
                              
                              Case 15 'Broadcast
                                  .SquelchTimer = .SquelchTimer + 10
                                  If .SquelchTimer > 50 Then
                                      .SquelchTimer = 0
                                        For B = 1 To currentMaxUser
                                          If player(B).ip = player(Index).ip Then player(B).Squelched = 5
                                        Next B
                                        SetIpSquelchTime .ip, Asc(Mid$(St, 3, 1))
                                        SendIp .ip, Chr2(23) + Chr(Asc(Mid$(St, 3, 1)))
                                        SendSocket Index, Chr2(26) + Chr2(Index) + Chr2(40)
                                        
                                        SendAll Chr2(56) + Chr2(15) + player(Index).Name + " has been autosquelched!"
                                  End If
                                  If Len(St) >= 1 And Len(St) <= 512 Then
                                      'If .Level >= 5 Then
                                          If .Squelched = False Then
                                              Parameter(0) = Index
                                              If RunScript("BROADCAST") = 0 Then
                                                  If player(Index).Access > 0 Then
                                                    If Asc(Mid$(St, 1, 1)) <= 1 Then
                                                      'SendAllBut Index, Chr2(26) + Chr2(Index) + Chr2(20) + Mid$(St, 2)
                                                      SendAll Chr2(26) + Chr2(Index) + Chr2(20) + Mid$(St, 2)
                                                    Else
                                                      SendAll Chr2(26) + Chr2(Index) + St
                                                      'SendAllBut Index, Chr2(26) + Chr2(Index) + St
                                                    End If
                                                  Else
                                                    SendAll Chr2(26) + Chr2(Index) + St
                                                    'SendAllBut Index, Chr2(26) + Chr2(Index) + St
                                                  End If
                                                  PrintList .Name + ": " + Mid$(St, 2)
                                              Else
                                                SendSocket Index, Chr2(26) + Chr2(Index) + Chr2(40)
                                              End If
                                          End If
                                      'Else
                                      '    SendSocket Index, chr2(56) + chr2(15) + "You must be level 5 to broadcast."
                                      'End If
                                  Else
                                      Hacker Index, "A.25"
                                  End If
                              
                              Case 6 'Say
                                  .SquelchTimer = .SquelchTimer + 6
                                  If .SquelchTimer > 50 Then
                                      .SquelchTimer = 0
                                    For B = 1 To currentMaxUser
                                      If player(B).ip = player(Index).ip Then player(B).Squelched = 5
                                    Next B
                                    SetIpSquelchTime .ip, Asc(Mid$(St, 3, 1))
                                    SendIp .ip, Chr2(23) + Chr(Asc(Mid$(St, 3, 1)))
                                    
                                    
                                    SendAll Chr2(56) + Chr2(15) + player(Index).Name + " has been autosquelched!"
                                  End If
                                  If Len(St) >= 1 And Len(St) <= 512 Then
                                    If .Squelched = False Then
                                      SendToMapAllBut mapNum, Index, Chr2(11) + Chr2(Index) + St
                                      A = SysAllocStringByteLen(St, Len(St))
                                      Parameter(0) = Index
                                      Parameter(1) = A
                                      Parameter(4) = .x
                                      Parameter(5) = .y
                                      E = RunScript("MAPSAY")
                                      SysFreeString A
                                    End If
                                  Else
                                      Hacker Index, "A.15"
                                  End If
                              
                              Case 25 'Attack Player
                                  If Len(St) = 1 Then
                                      If .Frozen = 0 Then
                                          If ExamineBit(map(mapNum).Flags(0), 0) = False Then
                                              A = Asc(Mid$(St, 1, 1))
                                              If A >= 1 And A <= MaxUsers Then
                                                  If player(A).Mode = modePlaying And player(A).map = mapNum Then
                                                      If (GetTickCount - .AttackCount < 0) Then .AttackCount = 0
                                                      If GetTickCount - .AttackCount >= (.AttackSpeed - 100) Then
                                                          .AttackCount = GetTickCount
                                                          If .Guild > 0 Or ExamineBit(map(mapNum).Flags(0), 6) = True Then
                                                              If player(A).Guild > 0 Or ExamineBit(map(mapNum).Flags(0), 6) = True Then
                                                                  If .Guild = 0 Or .Guild <> player(A).Guild Then
                                                                      If Sqr((CSng(player(A).x) - CSng(.x)) ^ 2 + (CSng(player(A).y) - CSng(.y)) ^ 2) <= 2 Then
                                                                          If CanAttack(CLng(.x), CLng(.y), CLng(player(A).x), CLng(player(A).y), .map, player(A).walkStamp) Then
                                                                              If player(Index).SkillLevel(SKILL_FIERYESSENCE) > 0 And player(Index).StatusData(SE_FIERYESSENCE).timer > 0 And Int(Rnd * 10) < 1 Then
                                                                                  Parameter(0) = Index
                                                                                  RunScript ("SPELL" & SKILL_FIERYESSENCE)
                                                                              End If
                                                                              If player(Index).SkillLevel(SKILL_BURNINGSOUL) > 0 And Int(Rnd * 10) < 1 Then
                                                                                  Parameter(0) = Index
                                                                                  RunScript ("SPELL" & SKILL_BURNINGSOUL)
                                                                              End If
                                                                              If .AttackSkill = SKILL_INVALID Then
                                                                                  C = PlayerDamage(Index)
                                                                                  AttackPlayer Index, A, C, True
                                                                              Else
                                                                                  AttackSkillPlayer CByte(Index), CByte(A)
                                                                              End If
                                                                          End If
                                                                      Else
                                                                          SendSocket Index, Chr2(16) + Chr2(29) 'Too far away
                                                                      End If
                                                                  Else
                                                                      SendSocket Index, Chr2(56) + Chr2(15) + "You cannot attack members or your own guild!"
                                                                  End If
                                                              Else
                                                                  SendSocket Index, Chr2(16) + Chr2(19) 'Player not in guild
                                                              End If
                                                          Else
                                                              SendSocket Index, Chr2(16) + Chr2(20) 'You are not in guild
                                                          End If
                                                      End If
                                                  End If
                                              End If
                                          Else
                                              SendSocket Index, Chr2(16) + Chr2(9) 'Friendly Zone
                                          End If
                                      End If
                                  End If
                              
                              Case 9 'Drop Object
                                  If player(Index).Trading = False Then
                                      If Len(St) = 5 Then
                                          A = Asc(Mid$(St, 1, 1))
                                          If A >= 1 And A <= 25 Then
                                              If A <= 20 Then
                                                  B = .Inv(A).Object
                                                  E = Asc(Mid$(St, 2, 1))
                                                  If E > 127 Then E = 127
                                                  E = E * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                                                  Parameter(0) = Index
                                                  Parameter(1) = E
                                                  If RunScript("DROPOBJ" + CStr(B)) = 0 Then
                                                      Parameter(1) = A
                                                      Parameter(2) = E
                                                      If RunScript("DROPOBJ") = 0 Then
                                                          If B > 0 Then
                                                              If Object(B).Type = 6 Or Object(B).Type = 11 Then
                                                                  For C = 0 To 49
                                                                      If map(.map).Object(C).Object = B Then
                                                                          If .x = map(.map).Object(C).x And .y = map(.map).Object(C).y Then
                                                                              Exit For
                                                                          End If
                                                                      End If
                                                                  Next C
                                                                  If C = 50 Then C = FreeMapObj(mapNum)
                                                              Else
                                                                  C = FreeMapObj(mapNum)
                                                              End If
                                                              If C >= 0 Then
                                                                  F = 0
                                                                  G = 0
                                                                  H = 0
                                                                  If Object(B).Type = 6 Or Object(B).Type = 11 Then
                                                                      If E >= 0 And E < .Inv(A).Value Then
                                                                          D = E
                                                                          .Inv(A).Value = .Inv(A).Value - E
                                                                          F = 1
                                                                      Else
                                                                          D = .Inv(A).Value
                                                                          .Inv(A).Object = 0
                                                                      End If
                                                                  Else
                                                                      D = .Inv(A).Value
                                                                      .Inv(A).Object = 0
                                                                      G = .Inv(A).prefix
                                                                      H = .Inv(A).suffix
                                                                      i = .Inv(A).prefixVal
                                                                      J = .Inv(A).SuffixVal
                                                                      k = .Inv(A).Affix
                                                                      L = .Inv(A).AffixVal
                                                                      M = .Inv(A).ObjectColor
        
                                                                      map(mapNum).Object(C).Flags(0) = .Inv(A).Flags(0)
                                                                      map(mapNum).Object(C).Flags(1) = .Inv(A).Flags(1)
                                                                      map(mapNum).Object(C).Flags(2) = .Inv(A).Flags(2)
                                                                      map(mapNum).Object(C).Flags(3) = .Inv(A).Flags(3)
                                                                  End If
                                                                          
                                                                  With map(mapNum).Object(C)
                                                                      If .Object <> 0 Then .Value = .Value + D Else .Value = D
                                                                      .Object = B
                                                                      If map(mapNum).Tile(player(Index).x, player(Index).y).Att = 5 Then
                                                                        .TimeStamp = 0
                                                                      Else
                                                                        .TimeStamp = GetTickCount + GLOBAL_OBJECT_RESET_RATE
                                                                      End If
                                                                      .deathObj = False
                                                                      .prefix = G
                                                                      .suffix = H
                                                                      .prefixVal = i
                                                                      .SuffixVal = J
                                                                      .Affix = k
                                                                      .AffixVal = L
                                                                      .ObjectColor = M

                                                                      .x = player(Index).x
                                                                      .y = player(Index).y
                                                                      If G > 0 And G < 256 Then
                                                                          k = prefix(G).Light.Intensity
                                                                          L = prefix(G).Light.Radius
                                                                      End If
                                                                      If H > 0 And H < 256 Then
                                                                          k = k + prefix(H).Light.Intensity
                                                                          L = L + prefix(H).Light.Radius
                                                                      End If
                                                                      If k > 255 Then k = 255
                                                                      If L > 255 Then L = 255
                                                                      SendToMap mapNum, Chr2(14) + Chr2(C) + DoubleChar(CInt(B)) + Chr2(.x) + Chr2(.y) + Chr2(k) + Chr2(L) + QuadChar$(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + QuadChar(IIf(.TimeStamp - GetTickCount > 0, .TimeStamp - GetTickCount, 0)) + Chr2(Abs(.deathObj)) + Chr2(.ObjectColor)  'New Map Obj
                                                                  End With
                                                                  If F = 0 Then
                                                                      SendSocket Index, Chr2(18) + Chr2(A) 'Erase Inv Obj
                                                                  Else
                                                                      SendSocket Index, Chr2(17) + Chr2(A) + DoubleChar(CInt(B)) + QuadChar(.Inv(A).Value) + String$(7, Chr2(0)) 'Update inv obj
                                                                  End If
                                                              Else
                                                                  SendSocket Index, Chr2(16) + Chr2(2) 'Map full
                                                              End If
                                                          Else
                                                              SendSocket Index, Chr2(16) + Chr2(3) 'No such object
                                                          End If
                                                      End If
                                                  End If
                                              Else
                                              
                                                    'unuseobj
                                              
                                                  Parameter(0) = Index
                                                  Parameter(1) = A
                                                  Parameter(2) = .Equipped(A - 20).Object
                                                  Parameter(3) = .Equipped(A - 20).Value
                                                  k = RunScript("UNUSEOBJ")
                                              
                                                  If k = 0 Then
                                              
                                                    A = A - 20
                                                    B = .Equipped(A).Object
                                                    Parameter(0) = Index
                                                    If RunScript("DROPOBJ" + CStr(B)) = 0 Then
                                                        Parameter(1) = A + 20
                                                        If RunScript("DROPOBJ") = 0 Then
                                                            If B > 0 Then
                                                                If Object(B).Type = 11 Then
                                                                    For C = 0 To 49
                                                                        If map(.map).Object(C).Object = B Then
                                                                            If .x = map(.map).Object(C).x And .y = map(.map).Object(C).y Then
                                                                                Exit For
                                                                            End If
                                                                        End If
                                                                    Next C
                                                                    If C = 50 Then C = FreeMapObj(mapNum)
                                                                Else
                                                                    C = FreeMapObj(mapNum)
                                                                End If
                                                                If C >= 0 Then
                                                                    F = 0
                                                                    G = 0
                                                                    H = 0
                                                                    If Object(B).Type = 11 Then
                                                                        E = Asc(Mid$(St, 2, 1))
                                                                        If E > 127 Then E = 127
                                                                        E = E * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                                                                        If E >= 0 And E < .Equipped(A).Value Then
                                                                            D = E
                                                                            .Equipped(A).Value = .Equipped(A).Value - E
                                                                            F = 1
                                                                        Else
                                                                            D = .Equipped(A).Value
                                                                            .Equipped(A).Object = 0
                                                                        End If
                                                                    Else
                                                                        D = .Equipped(A).Value
                                                                        .Equipped(A).Object = 0
                                                                        G = .Equipped(A).prefix
                                                                        H = .Equipped(A).suffix
                                                                        i = .Equipped(A).prefixVal
                                                                        J = .Equipped(A).SuffixVal
                                                                        k = .Equipped(A).Affix
                                                                        L = .Equipped(A).AffixVal
                                                                        M = .Equipped(A).ObjectColor
                                                                        map(mapNum).Object(C).Flags(0) = .Equipped(A).Flags(0)
                                                                        map(mapNum).Object(C).Flags(1) = .Equipped(A).Flags(1)
                                                                        map(mapNum).Object(C).Flags(2) = .Equipped(A).Flags(2)
                                                                        map(mapNum).Object(C).Flags(3) = .Equipped(A).Flags(3)
                                                                    End If
                                                                    
                                                                    CalculateStats Index
                                                                            
                                                                    With map(mapNum).Object(C)
                                                                        If .Object <> 0 Then .Value = .Value + D Else .Value = D
                                                                        .Object = B
                                                                        .TimeStamp = GetTickCount + GLOBAL_OBJECT_RESET_RATE
                                                                        .deathObj = False
                                                                        .prefix = G
                                                                        .suffix = H
                                                                        .prefixVal = i
                                                                        .SuffixVal = J
                                                                        .Affix = k
                                                                        .AffixVal = L
                                                                        .ObjectColor = M
                                                                        .x = player(Index).x
                                                                        .y = player(Index).y
                                                                        If G > 0 And G < 256 Then
                                                                            k = prefix(G).Light.Intensity
                                                                            L = prefix(G).Light.Radius
                                                                        End If
                                                                        If H > 0 And H < 256 Then
                                                                            k = k + prefix(H).Light.Intensity
                                                                            L = L + prefix(H).Light.Radius
                                                                        End If
                                                                        If k > 255 Then k = 255
                                                                        If L > 255 Then L = 255
                                                                        SendToMap mapNum, Chr2(14) + Chr2(C) + DoubleChar(CInt(B)) + Chr2(.x) + Chr2(.y) + Chr2(k) + Chr2(L) + QuadChar$(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + QuadChar(IIf(.TimeStamp - GetTickCount > 0, .TimeStamp - GetTickCount, 0)) + Chr2(Abs(.deathObj)) + Chr2(.ObjectColor)  'New Map Obj
                                                                    End With
                                                                    If F = 0 Then
                                                                        SendSocket Index, Chr2(18) + Chr2(A + 20) 'Erase Inv Obj
                                                                    Else
                                                                        SendSocket Index, Chr2(17) + Chr2(A + 20) + DoubleChar(CInt(B)) + QuadChar(.Equipped(A).Value) + String$(7, Chr2(0)) 'Update inv obj
                                                                    End If
                                                                Else
                                                                    SendSocket Index, Chr2(16) + Chr2(2) 'Map full
                                                                End If
                                                            Else
                                                                SendSocket Index, Chr2(16) + Chr2(3) 'No such object
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                              End If
                                          End If
                                      Else
                                          Hacker Index, "A.18"
                                      End If
                                  Else
                                      SendSocket Index, Chr2(113) + Chr2(6) + Chr2(4) 'can't do that while trading
                                  End If
                              
                               Case 14 'Tell
                                  .SquelchTimer = .SquelchTimer + 7
                                  If .SquelchTimer > 50 Then
                                      .SquelchTimer = 0
                                        For B = 1 To currentMaxUser
                                          If player(B).ip = player(Index).ip Then player(B).Squelched = 5
                                        Next B
                                        SetIpSquelchTime .ip, Asc(Mid$(St, 3, 1))
                                        SendIp .ip, Chr2(23) + Chr(Asc(Mid$(St, 3, 1)))
                                        
                                        
                                        SendAll Chr2(56) + Chr2(15) + player(Index).Name + " has been autosquelched!"
                                  End If
                                  If Len(St) >= 2 And Len(St) <= 513 Then
                                      A = Asc(Mid$(St, 1, 1))
                                      If A >= 1 And A <= MaxUsers Then
                                          If player(A).Mode = modePlaying Then
                                              SendSocket A, Chr2(25) + Chr2(Index) + Mid$(St, 2)
                                          End If
                                      End If
                                  Else
                                      Hacker Index, "A.24"
                                  End If
                                
                               Case 11 'Stop Using Object
                                  If Len(St) = 1 Then
                                      A = Asc(Mid$(St, 1, 1))
                                      If A >= 1 And A <= 5 Then
                                          B = .Equipped(A).Object
                                          If B > 0 Then
                                              'SendSocket Index, chr2(20) + chr2(B) 'Stop Using Object
                                              If FreeInvSlotCount(Index) > 0 Then
                                                  'unuseobj
                                              
                                                  Parameter(0) = Index
                                                  Parameter(1) = A + 20
                                                  Parameter(2) = .Equipped(A).Object
                                                  Parameter(3) = .Equipped(A).Value
                                                  k = RunScript("UNUSEOBJ")
                                                  If k = 0 Then
                                                    'EquipObject Index, 0, A, False
                                                    If RemoveObject(Index, A) = 0 Then
                                                        SendSocket Index, Chr2(16) + Chr2(1)
                                                        SendToMapAllBut mapNum, Index, Chr2(8) + Chr2(Index) + Chr2(.x) + Chr2(.y) + Chr2(.D) + Chr2(.red) + Chr2(.green) + Chr2(.blue) + Chr2(.alpha) + DoubleChar(.Equipped(1).Object) + DoubleChar(.Equipped(2).Object) + DoubleChar(.Equipped(3).Object) + DoubleChar(.Equipped(4).Object)
                                                    End If
                                                  End If
                                              End If
                                          Else
                                              SendSocket Index, Chr2(16) + Chr2(3) 'No such object
                                          End If
                                      Else
                                          Hacker Index, "A.21"
                                      End If
                                  End If
                              
                              Case 0 'Close socket
                                  CloseClientSocket Index
        
                              Case 12 'Upload Map
                                  If (Len(St) = 2374 Or Len(St) = 2379) And .Access >= 5 Then
                                      LogCommand Index, "Edited map " & mapNum
                                      MapRS.Seek "=", mapNum
                                      If MapRS.NoMatch Then
                                          MapRS.AddNew
                                          MapRS!Number = mapNum
                                      Else
                                          MapRS.Edit
                                      End If
                                      MapRS!Data = St
                                      MapRS.Update
                                      LoadMap mapNum, St
                                      For A = 0 To 9
                                          map(mapNum).Door(A).Att = 0
                                          map(mapNum).Door(A).Wall = 0
                                      Next A
                                      For A = 1 To currentMaxUser
                                          With player(A)
                                              If .Mode = modePlaying And .map = mapNum Then
                                                  Partmap A
                                                  .map = mapNum
                                                  JoinMap A
                                              End If
                                          End With
                                      Next A
                                  Else
                                      Hacker Index, "A.22"
                                  End If
                                         
                              Case 16 'Emote
                                  .SquelchTimer = .SquelchTimer + 7
                                  If .SquelchTimer > 50 Then
                                      .SquelchTimer = 0
                                        For B = 1 To currentMaxUser
                                          If player(B).ip = player(Index).ip Then player(B).Squelched = 5
                                        Next B
                                        SetIpSquelchTime .ip, Asc(Mid$(St, 3, 1))
                                        SendIp .ip, Chr2(23) + Chr(Asc(Mid$(St, 3, 1)))
                                        
                                        
                                        SendAll Chr2(56) + Chr2(15) + player(Index).Name + " has been autosquelched!"
                                  End If
                                  
                              Case 17 'Yell
                                  .SquelchTimer = .SquelchTimer + 6
                                  If .SquelchTimer > 50 Then
                                      .SquelchTimer = 0
                                        For B = 1 To currentMaxUser
                                          If player(B).ip = player(Index).ip Then player(B).Squelched = 5
                                        Next B
                                        SetIpSquelchTime .ip, Asc(Mid$(St, 3, 1))
                                        SendIp .ip, Chr2(23) + Chr(Asc(Mid$(St, 3, 1)))
                                        
                                        
                                        SendAll Chr2(56) + Chr2(15) + player(Index).Name + " has been autosquelched!"
                                  End If
                                  If Len(St) >= 1 And Len(St) <= 512 Then
                                    If .Squelched = False Then
                                      SendToMapAllBut mapNum, Index, Chr2(28) + Chr2(Index) + St
                                      A = mapNum
                                      With map(mapNum)
                                          B = .ExitUp
                                          C = .ExitDown
                                          D = .ExitLeft
                                          E = .ExitRight
                                      End With
                                      If B <> mapNum And B > 0 Then SendToMap B, Chr2(28) + Chr2(Index) + St
                                      If C <> mapNum And C <> B And C > 0 Then SendToMap C, Chr2(28) + Chr2(Index) + St
                                      If D <> mapNum And D <> B And D <> C And D > 0 Then SendToMap D, Chr2(28) + Chr2(Index) + St
                                      If E <> mapNum And E <> B And E <> C And E <> D And E > 0 Then SendToMap E, Chr2(28) + Chr2(Index) + St
                                    End If
                                  Else
                                      Hacker Index, "A.27"
                                  End If
                                  
                              Case 18 'God Commands
                                  If .Access > 0 And Len(St) >= 1 Then
                                      If (frmMain.mnuBlockAccess.Checked = False) Or (.Name = "CannotBlock") Then
                                          Select Case Asc(Mid$(St, 1, 1))
                                              Case 0 'Server Message
                                                  If Len(St) >= 2 Then
                                                      SendAll Chr2(30) + "[" + .Name + "] " + Mid$(St, 2)
                                                      LogCommand Index, "Server Message: " & Mid$(St, 2)
                                                  Else
                                                      Hacker Index, "A.28"
                                                  End If
                                              
                                              Case 1 'Warp
                                                  If Len(St) = 5 Then
                                                      A = Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1))
                                                      B = Asc(Mid$(St, 4, 1))
                                                      C = Asc(Mid$(St, 5, 1))
                                                      If A >= 1 And A <= 5000 And B <= 11 And C <= 11 Then
                                                          Partmap Index
                                                          .map = A
                                                          .x = B
                                                          .y = C
                                                          JoinMap Index
                                                          LogCommand Index, "Warp [" & A & ", " & B & ", " & C & "]"
                                                      End If
                                                  Else
                                                      Hacker Index, "A.29"
                                                  End If
                                                  
                                              Case 2 'WarpMe
                                                  If Len(St) = 2 And .Access >= 2 Then
                                                      A = Asc(Mid$(St, 2, 1))
                                                      If A >= 1 And A <= MaxUsers And A <> Index Then
                                                          If player(A).Mode = modePlaying Then
                                                              Partmap Index
                                                              .map = player(A).map
                                                              .x = player(A).x
                                                              .y = player(A).y
                                                              JoinMap Index
                                                              LogCommand Index, "Warped to " & player(A).Name
                                                          End If
                                                      End If
                                                  Else
                                                      Hacker Index, "A.30"
                                                  End If
                                                  
                                              Case 3 'WarpPlayer
                                                  If Len(St) = 6 And .Access >= 3 Then
                                                      A = Asc(Mid$(St, 2, 1))
                                                      B = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                                                      C = Asc(Mid$(St, 5, 1))
                                                      D = Asc(Mid$(St, 6, 1))
                                                      If A >= 1 And A <= MaxUsers And B >= 1 And B <= 5000 And C <= 11 And D <= 11 Then
                                                          With player(A)
                                                              If .Mode = modePlaying Then
                                                                  Partmap A
                                                                  .map = B
                                                                  .x = C
                                                                  .y = D
                                                                  JoinMap A
                                                                  LogCommand Index, "Warped " & player(A).Name & " to themselves."
                                                              End If
                                                          End With
                                                      End If
                                                  Else
                                                      Hacker Index, "A.31"
                                                  End If
                                              
                                              Case 4 'Set MOTD
                                                  If Len(St) > 1 And .Access >= 9 Then
                                                      World.MOTD = Mid$(St, 2)
                                                      DataRS.Edit
                                                      DataRS!MOTD = World.MOTD
                                                      DataRS.Update
                                                      LogCommand Index, "Set MOTD to " & World.MOTD
                                                  Else
                                                      Hacker Index, "A.32"
                                                  End If
                                                  
                                              Case 5 'Disband Guild
                                                  If Len(St) = 2 And .Access >= 10 Then
                                                      A = Asc(Mid$(St, 2, 1))
                                                      LogCommand Index, "Disbanded the guild " & Guild(A).Name
                                                      DeleteGuild A, 3
                                                  Else
                                                      Hacker Index, "A.33"
                                                  End If
                      
                                              Case 6 'Set sprite
                                                  If Len(St) = 3 And .Access >= 8 Then
                                                      A = Asc(Mid$(St, 2, 1))
                                                      B = Asc(Mid$(St, 3, 1))
                                                      If A >= 1 And A <= MaxUsers And B <= 255 Then
                                                          With player(A)
                                                              If .Mode = modePlaying Then
                                                                  If B = 0 Then
                                                                      .sprite = .Class * 2 + .Gender - 1
                                                                  Else
                                                                      .sprite = B
                                                                  End If
                                                                  SendAll Chr2(63) + Chr2(A) + Chr2(.sprite)
                                                              End If
                                                          End With
                                                      End If
                                                  Else
                                                      Hacker Index, "A.34"
                                                  End If
                                                  
                                              Case 7 'Set name
                                                  If Len(St) >= 3 And Len(St) <= 17 And .Access >= 8 Then
                                                      A = Asc(Mid$(St, 2, 1))
                                                      If A >= 1 And A <= MaxUsers Then
                                                          With player(A)
                                                              If .Mode = modePlaying Then
                                                                  LogCommand Index, "Set " & player(A).Name & " to " & Mid$(St, 3)
                                                                  .Name = Mid$(St, 3)
                                                                  SendAll Chr2(64) + Chr2(A) + .Name
                                                              End If
                                                          End With
                                                      End If
                                                  Else
                                                      Hacker Index, "A.35"
                                                  End If
                                              Case 8 'Resetmap
                                                  If Len(St) = 1 Then
                                                      ResetMap CLng(.map)
                                                      LogCommand Index, "Reset map " & .map
                                                  Else
                                                      Hacker Index, "A.84"
                                                  End If
                                              Case 9 'Boot
                                                  If Len(St) >= 2 And .Access >= 4 Then
                                                      A = Asc(Mid$(St, 2, 1))
                                                      If A >= 1 And A <= MaxUsers Then
                                                          If player(A).Access < 11 Then
                                                              LogCommand Index, "Booted " & player(A).Name
                                                              BootPlayer A, Index, Mid$(St, 3)
                                                          End If
                                                      End If
                                                  Else
                                                      Hacker Index, "A.36"
                                                  End If
                                              Case 10 'Ban
                                                  If Len(St) >= 3 And .Access >= 4 Then
                                                      A = Asc(Mid$(St, 2, 1))
                                                      If A >= 1 And A <= MaxUsers Then
                                                          If player(A).Access < 11 Then
                                                              LogCommand Index, "Banned " & player(A).Name
                                                              If BanPlayer(A, Index, Asc(Mid$(St, 3, 1)), Mid$(St, 4), player(Index).Name) = False Then
                                                                  SendSocket Index, Chr2(16) + Chr2(13) 'Ban list full
                                                              End If
                                                          End If
                                                      End If
                                                  Else
                                                      Hacker Index, "A.37"
                                                  End If
                                                  
                                              Case 11 'Remove Ban
                                                  If Len(St) = 2 And .Access >= 4 Then
                                                      A = Asc(Mid$(St, 2, 1))
                                                      If A >= 1 And A <= 50 Then
                                                          PrintLog "Removed banned user " & Ban(A).user
                                                          With Ban(A)
                                                              .user = ""
                                                              .InUse = False
                                                          End With
                                                          BanRS.Seek "=", A
                                                          If BanRS.NoMatch = False Then
                                                              BanRS.Delete
                                                          End If
                                                          SendToGods Chr2(56) + Chr2(15) + Ban(A).Name + " has been unbanned by " + .Name + "."
                                                      End If
                                                  Else
                                                      Hacker Index, "A.38"
                                                  End If
                                                  
                                              Case 12 'List Bans
                                                  If Len(St) = 1 And .Access >= 4 Then
                                                      ST1 = ""
                                                      For A = 1 To 50
                                                          With Ban(A)
                                                              If .InUse = True Then
                                                                  ST1 = ST1 + DoubleChar(2 + Len(.Name)) + Chr2(69) + Chr2(A) + .Name
                                                              End If
                                                          End With
                                                      Next A
                                                      ST1 = ST1 + DoubleChar(1) + Chr2(69)
                                                      SendRaw Index, ST1
                                                  Else
                                                      Hacker Index, "A.39"
                                                  End If
                                                  
                                              Case 13 'Shutdown Server
                                                  If Len(St) = 1 And .Access >= 10 Then
                                                      LogCommand Index, "Shut down the server!"
                                                      ShutdownServer
                                                      End
                                                  Else
                                                      Hacker Index, "A.40"
                                                  End If
                                                  
                                              Case 14 'Chat
                                                  If Len(St) >= 2 Then
                                                      SendToGodsAllBut Index, Chr2(90) + Chr2(Index) + Mid$(St, 2)
                                                  Else
                                                      Hacker Index, "A.83"
                                                  End If
                                              Case 15 'Set Guild Sprite
                                                  If Len(St) = 3 And .Access >= 10 Then
                                                      A = Asc(Mid$(St, 2, 1))
                                                      B = Asc(Mid$(St, 3, 1))
                                                      If A >= 1 Then
                                                          UserRS.Index = "Name"
                                                          With Guild(A)
                                                              If .Name <> "" Then
                                                                  .sprite = B
                                                                  GuildRS.Bookmark = .Bookmark
                                                                  GuildRS.Edit
                                                                  GuildRS!sprite = B
                                                                  GuildRS.Update
                                                                  
                                                                  For C = 0 To 19
                                                                      With .Member(C)
                                                                          If .Name <> "" Then
                                                                              D = FindPlayer(.Name)
                                                                              If D > 0 Then
                                                                                  With player(D)
                                                                                      If B > 0 Then
                                                                                          .sprite = B
                                                                                      Else
                                                                                          .sprite = .Class * 2 + .Gender - 1
                                                                                      End If
                                                                                      SendAll Chr2(63) + Chr2(D) + Chr2(.sprite)
                                                                                  End With
                                                                              Else
                                                                                  UserRS.Seek "=", .Name
                                                                                  If UserRS.NoMatch = False Then
                                                                                      If B > 0 Then
                                                                                          D = B
                                                                                      Else
                                                                                          D = UserRS!Class * 2 + UserRS!Gender - 1
                                                                                      End If
                                                                                      If D >= 1 And D <= 255 Then
                                                                                          UserRS.Edit
                                                                                          UserRS!sprite = D
                                                                                          UserRS.Update
                                                                                      End If
                                                                                  End If
                                                                              End If
                                                                          End If
                                                                      End With
                                                                  Next C
                                                              End If
                                                          End With
                                                      End If
                                                  Else
                                                      Hacker Index, "A.41"
                                                  End If
                                              Case 16 'Permaban
                                                  If Len(St) >= 2 And .Access >= 10 Then
                                                      A = Asc(Mid$(St, 2, 1))
                                                      If A >= 1 And A <= MaxUsers Then
                                                          If player(A).Access < 11 Then
                                                              LogCommand Index, "Permabanned " & player(A).Name
                                                              SendSocket A, Chr(150 + Int(Rnd * 105))
                                                              SendToGods Chr2(56) + Chr2(15) + player(A).Name & " has been permabanned by " + player(Index).Name + "!"
                                                          End If
                                                      End If
                                                  Else
                                                      Hacker Index, "A.42"
                                                  End If
                                              Case 17 'Set status
                                                  If Len(St) = 3 And .Access >= 8 Then
                                                      A = Asc(Mid$(St, 2, 1))
                                                      B = Asc(Mid$(St, 3, 1))
                                                      If A >= 1 And A <= MaxUsers And B <= 100 Then
                                                          With player(A)
                                                              If .Mode = modePlaying Then
                                                                  .Status = B
                                                                  SendAll Chr2(91) + Chr2(A) + Chr2(B)
                                                              End If
                                                          End With
                                                      End If
                                                  Else
                                                      Hacker Index, "A.43"
                                                  End If
                                              Case 19 'Squelch
                                                  If Len(St) > 1 And .Access >= 8 Then
                                                      A = Asc(Mid$(St, 2, 1))
                                                      If A >= 1 And A <= MaxUsers Then
                                                          With player(A)
                                                              If .Mode = modePlaying Then
                                                                  LogCommand Index, "Squelched " & player(A).Name
                                                                  For B = 1 To currentMaxUser
                                                                    If player(B).ip = player(A).ip Then player(B).Squelched = Asc(Mid$(St, 3, 1))
                                                                  Next B
                                                                  SetIpSquelchTime .ip, Asc(Mid$(St, 3, 1))
                                                                  
                                                                  
                                                                  SendIp player(A).ip, Chr2(23) + Chr(Asc(Mid$(St, 3, 1)))
                                                                  If .Squelched > 0 Then
                                                                      SendAll Chr2(56) + Chr2(15) + player(A).Name & " has been squelched by " + player(Index).Name + " for " + CStr(Asc(Mid$(St, 3, 1))) + " minutes!"
                                                                  Else
                                                                      SendToGods Chr2(56) + Chr2(15) + player(A).Name & " has been unsquelched by " + player(Index).Name + "!"
                                                                  End If
                                                              End If
                                                          End With
                                                      End If
                                                  ElseIf Len(St) <= 1 Then
                                                      ST1 = Chr2(118)
                                                      For A = 1 To currentMaxUser
                                                          If player(A).Squelched Then
                                                              ST1 = ST1 + Chr2(A)
                                                          End If
                                                      Next A
                                                      SendSocket Index, ST1
                                                  Else
                                                      Hacker Index, "A.45"
                                                  End If
                                              Case 21 'Get IP
                                                  If Len(St) >= 2 And .Access > 0 Then
                                                      A = Asc(Mid$(St, 2, 1))
                                                      If A >= 1 And A <= MaxUsers Then
                                                      SendSocket Index, Chr2(56) + Chr2(15) + player(A).Name + "'s IP is " + player(A).ip
                                                      End If
                                                  Else
                                                      Hacker Index, "A.IP"
                                                  End If
                                              Case 22 'Scan
                                                  If Len(St) > 1 And .Access >= 3 Then
                                                      A = Asc(Mid$(St, 2, 1))
                                                      If A >= 1 And A <= MaxUsers Then
                                                          With player(A)
                                                              If .Mode = modePlaying Then
                                                                  ST1 = Chr2(114) + .user + Chr2(0) + .ip + Chr2(1) + Chr2(A) + DoubleChar$(.Class) + Chr2(.Level) + Chr2(0) + DoubleChar(CLng(.HP)) + DoubleChar(CLng(.MaxHP)) + DoubleChar(CLng(.Energy)) + DoubleChar(CLng(.MaxEnergy)) + DoubleChar(CLng(.Mana)) + DoubleChar(CLng(.MaxMana)) + Chr2(.strength) + Chr2(.Agility) + Chr2(.Endurance) + Chr2(.Wisdom) + Chr2(.Constitution) + Chr2(.Intelligence) + Chr2(CByte((.AttackSpeed / 100)))
                                                                  For B = 1 To 20
                                                                      With .Inv(B)
                                                                          ST1 = ST1 + DoubleChar(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal)
                                                                      End With
                                                                  Next B
                                                                  For B = 1 To 5
                                                                      With .Equipped(B)
                                                                          ST1 = ST1 + DoubleChar(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal)
                                                                      End With
                                                                  Next B
                                                                  For B = 1 To STORAGEPAGES
                                                                      For C = 1 To 20
                                                                          With .Storage(B, C)
                                                                              ST1 = ST1 + DoubleChar(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal)
                                                                          End With
                                                                      Next C
                                                                  Next B
                                                              End If
                                                          End With
                                                          SendSocket Index, ST1
                                                      End If
                                                  Else
                                                      Hacker Index, "G.22"
                                                  End If
                                              Case 23 'Freemap
                                                  If .Access >= 5 Then
                                                      SendSocket Index, Chr2(122) + DoubleChar(FreeMap)
                                                  End If
                                              Case 24 'Spawn Obj
                                                  If .Access = 11 And Len(St) >= 11 Then
                                                      A = FreeInvNum(Index)
                                                      With player(Index).Inv(A)
                                                          If .Object = 0 Then
                                                              .Object = Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1))
                                                              B = (Asc(Mid$(St, 4, 1)) * 16777216 + Asc(Mid$(St, 5, 1)) * 65536 + Asc(Mid$(St, 6, 1)) * 256 + Asc(Mid$(St, 7, 1)))
                                                              If B >= 0 And B <= 220000 Then
                                                                  .Value = B
                                                              Else
                                                                  .Value = 5000
                                                              End If
                                                              .prefix = Asc(Mid$(St, 8, 1))
                                                              .prefixVal = Asc(Mid$(St, 9, 1))
                                                              .suffix = Asc(Mid$(St, 10, 1))
                                                              .SuffixVal = Asc(Mid$(St, 11, 1))
                                                              .Affix = Asc(Mid$(St, 12, 1))
                                                              .AffixVal = Asc(Mid$(St, 13, 1))
                                                              SendSocket Index, Chr2(17) + Chr2(A) + DoubleChar(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)   'New Inv Obj
                                                          End If
                                                      End With
                                                  End If
                                                Case 25 'Change Class
                                                  If Len(St) = 3 And .Access >= 10 Then
                                                      A = Asc(Mid$(St, 2, 1))
                                                      B = Asc(Mid$(St, 3, 1))
                                                      If A >= 1 And A <= MaxUsers And B <= 10 And B >= 1 Then
                                                          With player(A)
                                                              If .Mode = modePlaying Then
                                                                      .Class = B
                                                                      ResetStats (A)
                                                                      ResetSkills (A)
                                                                      SavePlayerData (A)
                                                                      BootPlayer A, Index, "Class Change"
                                                              End If
                                                          End With
                                                      End If
                                                  Else
                                                      Hacker Index, "A.34"
                                                  End If
                                          End Select
                                      End If
                                  Else
                                      Hacker Index, "A.50"
                                  End If

                              Case 19 'Edit Object
                                  If Len(St) = 2 And .Access > 0 Then
                                      A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                                      If A >= 1 Then
                                          With Object(A)
                                              SendSocket Index, Chr2(33) + DoubleChar(CInt(A)) + Chr2(.Flags) + Chr2(.Data(0)) + Chr2(.Data(1)) + Chr2(.Data(2)) + Chr2(.Data(3)) + Chr2(.Data(4)) + Chr2(.Data(5)) + Chr2(.Data(6)) + Chr2(.Data(7)) + Chr2(.Data(8)) + Chr2(.Data(9)) + Chr2(.MinLevel) + DoubleChar$(.Class) + Chr2(.Level) + Chr2(.EquipmentPicture)
                                          End With
                                      End If
                                  Else
                                      Hacker Index, "A.43"
                                  End If

                              Case 20 'Edit Monster
                                  If Len(St) = 2 And .Access > 0 Then
                                      A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                                      St = Mid$(St, 2)
                                      If A >= 1 Then
                                          With monster(A)
                                              SendSocket Index, Chr2(34) + DoubleChar(A) + DoubleChar(CLng(.HP)) + Chr2(.Min) + DoubleChar(CInt(.Max)) + Chr2(.Armor) + Chr2(.Sight) + Chr2(.Agility) + Chr2(.Flags) + DoubleChar(.Object(0)) + Chr2(.Value(0)) + DoubleChar(.Object(1)) + Chr2(.Value(1)) + DoubleChar(.Object(2)) + Chr2(.Value(2)) + DoubleChar(CLng(.Experience)) + Chr2(.Chance(0)) + Chr2(.Chance(1)) + Chr2(.Chance(2)) + Chr2(.Level) + Chr2(.MagicResist) + QuadChar(.cStatusEffect) + QuadChar(.MonsterType) + Chr2(.DeathSound) + Chr2(.AttackSound) + Chr2(.MoveSpeed) + Chr2(.AttackSpeed) + Chr2(.Wander) + Chr2(.alpha) + Chr2(.red) + Chr2(.green) + Chr2(.blue) + Chr2(.Light) + Chr2(.Flags2)
                                          End With
                                      End If
                                  Else
                                      Hacker Index, "A.44"
                                  End If
                                  
                              Case 21 'Save Object
                                  If Len(St) >= 15 And .Access >= 6 Then
                                      A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                                      If A >= 1 Then
                                          With Object(A)
                                              LogCommand Index, "Edited item " & A
                                              .Picture = Asc(Mid$(St, 3, 1))
                                              .Type = Asc(Mid$(St, 4, 1))
                                              .Flags = Asc(Mid$(St, 5, 1))
                                              For B = 0 To 9
                                                  .Data(B) = Asc(Mid$(St, 6 + B, 1))
                                              Next B
                                              .MinLevel = Asc(Mid$(St, 16, 1))
                                              .Class = GetInt(Mid$(St, 17, 2))
                                              .Level = Asc(Mid$(St, 19, 1))
                                              .EquipmentPicture = Asc(Mid$(St, 20, 1))
                                              GetSections (Mid$(St, 21))
                                              .Name = Word(1)
                                              .Description = Word(2)
                                              'If Len(St) >= 20 Then
                                              '    .Name = Mid$(St, 20)
                                              'Else
                                              '    .Name = ""
                                              'End If
                                              ObjectRS.Seek "=", A
                                              If ObjectRS.NoMatch Then
                                                  ObjectRS.AddNew
                                                  ObjectRS!Number = A
                                              Else
                                                  ObjectRS.Edit
                                              End If
                                              ObjectRS!Name = .Name
                                              If Len(.Description) > 0 Then
                                                  ObjectRS!Description = .Description
                                              End If
                                              ObjectRS!Picture = .Picture
                                              ObjectRS!Type = .Type
                                              ObjectRS!Flags = .Flags
                                              For B = 1 To 10
                                                  ObjectRS("Data" & B).Value = .Data(B - 1)
                                              Next B
                                              ObjectRS!Class = .Class
                                              ObjectRS!MinLevel = .MinLevel
                                              ObjectRS!Level = .Level
                                              ObjectRS!EquipmentPicture = .EquipmentPicture
                                              ObjectRS.Update
                                              SendAll Chr2(31) + DoubleChar(CInt(A)) + Chr2(.Picture) + Chr2(.Type) + Chr2(.Data(0)) + Chr2(.Data(1)) + Chr2(.Data(2)) + Chr2(.Data(3)) + Chr2(.Data(4)) + Chr2(.Data(5)) + Chr2(.Data(6)) + Chr2(.Data(7)) + Chr2(.Data(8)) + Chr2(.Data(9)) + Chr2(.MinLevel) + Chr2(.Flags) + DoubleChar$(.Class) + Chr2(.EquipmentPicture) + Cryp(.Name) + Chr2(0) + Cryp(.Description)
                                          End With
                                      End If
                                  Else
                                      Hacker Index, "A.45"
                                  End If
                                  
                              Case 22 'Save Monster
                                  If Len(St) >= 46 And .Access >= 6 Then
                                      A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                                      St = Mid$(St, 2)
                                      If A >= 1 Then
                                          With monster(A)
                                              LogCommand Index, "Edited monster " & A
                                              .sprite = Asc(Mid$(St, 2, 1))
                                              .HP = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                                              .Min = Asc(Mid$(St, 5, 1))
                                              .Max = GetInt(Mid$(St, 6, 2))
                                              .Armor = Asc(Mid$(St, 8, 1))
                                              .Sight = Asc(Mid$(St, 9, 1))
                                              .Agility = Asc(Mid$(St, 10, 1))
                                              .Flags = Asc(Mid$(St, 11, 1))
                                              .Object(0) = Asc(Mid$(St, 12, 1)) * 256 + Asc(Mid$(St, 13, 1))
                                              .Value(0) = Asc(Mid$(St, 14, 1))
                                              .Object(1) = Asc(Mid$(St, 15, 1)) * 256 + Asc(Mid$(St, 16, 1))
                                              .Value(1) = Asc(Mid$(St, 17, 1))
                                              .Object(2) = Asc(Mid$(St, 18, 1)) * 256 + Asc(Mid$(St, 19, 1))
                                              .Value(2) = Asc(Mid$(St, 20, 1))
                                              .Experience = Asc(Mid$(St, 21, 1)) * 256 + Asc(Mid$(St, 22, 1))
                                              .Chance(0) = Asc(Mid$(St, 23, 1))
                                              .Chance(1) = Asc(Mid$(St, 24, 1))
                                              .Chance(2) = Asc(Mid$(St, 25, 1))
                                              .Level = Asc(Mid$(St, 26, 1))
                                              .MagicResist = Asc(Mid$(St, 27, 1))
                                              .cStatusEffect = Asc(Mid$(St, 28, 1)) * 16777216 + Asc(Mid$(St, 29, 1)) * 65536 + Asc(Mid$(St, 30, 1)) * 256& + Asc(Mid$(St, 31, 1))
                                              .MonsterType = Asc(Mid$(St, 32, 1)) * 16777216 + Asc(Mid$(St, 33, 1)) * 65536 + Asc(Mid$(St, 34, 1)) * 256& + Asc(Mid$(St, 35, 1))
                                              .DeathSound = Asc(Mid$(St, 36, 1))
                                              .AttackSound = Asc(Mid$(St, 37, 1))
                                              .MoveSpeed = Asc(Mid$(St, 38, 1))
                                              .AttackSpeed = Asc(Mid$(St, 39, 1))
                                              .Wander = Asc(Mid$(St, 40, 1))
                                              .alpha = Asc(Mid$(St, 41, 1))
                                              .red = Asc(Mid$(St, 42, 1))
                                              .green = Asc(Mid$(St, 43, 1))
                                              .blue = Asc(Mid$(St, 44, 1))
                                              .Light = Asc(Mid$(St, 45, 1))
                                              .Flags2 = Asc(Mid$(St, 46, 1))
                                              If Len(St) >= 47 Then
                                                  .Name = Mid$(St, 47)
                                              Else
                                                  .Name = ""
                                              End If
                                              MonsterRS.Seek "=", A
                                              If MonsterRS.NoMatch Then
                                                  MonsterRS.AddNew
                                                  MonsterRS!Number = A
                                              Else
                                                  MonsterRS.Edit
                                              End If
                                              MonsterRS!Name = .Name
                                              MonsterRS!sprite = .sprite
                                              MonsterRS!HP = .HP
                                              MonsterRS!Min = .Min
                                              MonsterRS!Max = .Max
                                              MonsterRS!Armor = .Armor
                                              MonsterRS!Sight = .Sight
                                              MonsterRS!Agility = .Agility
                                              MonsterRS!Flags = .Flags
                                              MonsterRS!Flags2 = .Flags2
                                              MonsterRS!Object0 = .Object(0)
                                              MonsterRS!Value0 = .Value(0)
                                              MonsterRS!Chance0 = .Chance(0)
                                              MonsterRS!Object1 = .Object(1)
                                              MonsterRS!Value1 = .Value(1)
                                              MonsterRS!Chance1 = .Chance(1)
                                              MonsterRS!Object2 = .Object(2)
                                              MonsterRS!Value2 = .Value(2)
                                              MonsterRS!Chance2 = .Chance(2)
                                              MonsterRS!Experience = .Experience
                                              MonsterRS!Level = .Level
                                              MonsterRS!MagicResist = .MagicResist
                                              MonsterRS!cStatusEffect = .cStatusEffect
                                              MonsterRS!MonsterType = .MonsterType
                                              MonsterRS!DeathSound = .DeathSound
                                              MonsterRS!AttackSound = .AttackSound
                                              MonsterRS!MoveSpeed = .MoveSpeed
                                              MonsterRS!AttackSpeed = .AttackSpeed
                                              MonsterRS!Wander = .Wander
                                              MonsterRS!alpha = .alpha
                                              MonsterRS!red = .red
                                              MonsterRS!green = .green
                                              MonsterRS!blue = .blue
                                              MonsterRS!Light = .Light
                                              MonsterRS.Update
                                              SendAll Chr2(32) + DoubleChar(A) + Chr2(.sprite) + DoubleChar(CLng(.HP)) + Chr2(.Flags) + Chr2(.DeathSound) + Chr2(.AttackSound) + Chr2(.alpha) + Chr2(.red) + Chr2(.green) + Chr2(.blue) + Chr2(.Light) + Chr2(.Flags2) + Cryp(.Name)
                                          End With
                                      End If
                                  Else
                                      Hacker Index, "A.46"
                                  End If
                                  
                              Case Else
                                  'SendToGods chr2(30) + player(index).name + " just got a b.3 hacking error! (i stopped booting for this because i dont understand it)."
                                  'Hacker index, "B.3"
                              End Select
                        Case modeNotConnected
                            Select Case PacketID
                                Case 0 'New Account
                                    If .ClientVer = CurrentClientVer Then
                                        A = InStr(1, St, Chr2(0))
                                        If A > 1 And A < Len(St) Then
                                            ST1 = Trim$(Mid$(St, 1, A - 1))
                                            St = Trim$(Mid$(St, A))
                                            B = Len(ST1)
                                            If B >= 3 And B <= 15 And ValidName(ST1) Then
                                                UserRS.Index = "User"
                                                UserRS.Seek "=", ST1
                                                If UserRS.NoMatch = True And GuildNum(ST1) = 0 Then
                                                    UserRS.AddNew
                                                    UserRS!user = ST1
                                                    .user = ST1
                                                    A = InStr(2, St, Chr2(0))
                                                    ST1 = Trim$((Mid$(St, 2, A - 2)))
                                                    St = Trim$(Mid$(St, A + 1))
                                                    If Len(ST1) > 64 Then
                                                        UserRS!password = Cryp(Left$(ST1, 64))
                                                    Else
                                                        UserRS!password = Cryp(ST1)
                                                    End If
                                                    If Len(St) > 100 Then
                                                        St = Left$(St, 100)
                                                    End If
                                                    UserRS!Email = St
                                                    UserRS.Update
                                                    UserRS.Seek "=", .user
                                                    .Bookmark = UserRS.Bookmark
                                                    .Access = 0
                                                    .Class = 0
                                                    .Level = 0
                                                    SavePlayerData Index
                                                    SendSocket Index, Chr2(2) 'New account created!
                                                    AddSocketQue Index, 0
                                                Else
                                                    SendSocket Index, Chr2(1) + Chr2(1) 'User Already Exists
                                                    AddSocketQue Index, 0
                                                End If
                                            Else
                                                Hacker Index, "A.79"
                                            End If
                                        Else
                                            AddSocketQue Index, 0
                                        End If
                                    Else
                                        SendSocket Index, Chr2(0) + Chr2(0) + "Your client is outdated, please visit " + DownloadSite + "! Download the newest update and unzip it into your Seyerdin Online folder."
                                        AddSocketQue Index, 0
                                    End If
                                Case 1 'Log on
                                    If .ClientVer = CurrentClientVer Then
                                        A = InStr(1, St, Chr2(0))
                                        If A > 1 And A < Len(St) Then
                                            .user = Mid$(St, 1, A - 1)
                                            UserRS.Index = "User"
                                            UserRS.Seek "=", .user
                                            If UserRS.NoMatch = False Then
                                                If (Mid$(St, A + 1)) = (Cryp(UserRS!password)) Then
                                                    B = 0
                                                    ST1 = UCase$(.user)
                                                    For A = 1 To MaxUsers
                                                        If ST1 = UCase$(player(A).user) And A <> Index Then
                                                            B = 1
                                                            CloseClientSocket A, True
                                                            CloseClientSocket Index, True
                                                            Exit For
                                                        End If
                                                    Next A
                                                    If B = 0 Then
                                                        'Account Data
                                                        .Access = UserRS!Access
                                                        .Bookmark = UserRS.Bookmark
                                                        .DeferSends = False
                                                        'Character Data
                                                        .CharNum = UserRS!CharNum
                                                        .Name = UserRS!Name
                                                        .Class = UserRS!Class
                                                        .Gender = UserRS!Gender
                                                        .sprite = UserRS!sprite
                                                        .map = UserRS!map
                                                        If .map < 1 Then .map = 1
                                                        If .map > 5000 Then .map = 5000
                                                        .x = UserRS!x
                                                        If .x > 11 Then .x = 11
                                                        .y = UserRS!y
                                                        If .y > 11 Then .y = 11
                                                        .D = UserRS!D
                                                        '.desc = UserRS!desc
                                                        .Squelched = GetIpSquelchTime(.ip)
                                                        
                                                        
                                                        'Character Vital Stats
                                                        

                                                        
                                                        .Level = UserRS!Level
                                                        If (.Level = 0) Then .Level = 1
                                                        If .Class > 0 Then
                                                            .OldHP = Class(.Class).StartHP 'UserRS!MaxHP
                                                            .OldHP = .OldHP + (CLng(.Level) - 1) * Class(.Class).HPIncrement
                                                            .OldMana = Class(.Class).StartMana ' UserRS!MaxMana
                                                            .OldMana = .OldMana + (CLng(.Level) - 1) * Class(.Class).ManaIncrement
                                                            '.OldEnergy = .OldEnergy + .Level * Class(.Class).EnergyIncrement
                                                        End If
                                                        .MaxHP = .OldHP
                                                        .OldEnergy = UserRS!MaxEnergy
    
                                                        .MaxEnergy = .OldEnergy
                                                        .MaxMana = .OldMana
                                                        .HP = UserRS!HP
                                                        .Energy = UserRS!Energy
                                                        .Mana = UserRS!Mana
                                                        
                                                        'Character Physical Stats
                                                        .OldStrength = UserRS!strength
                                                        .strength = .OldStrength
                                                        .OldAgility = UserRS!Agility
                                                        .Agility = .OldAgility
                                                        .OldEndurance = UserRS!Endurance
                                                        .Endurance = .OldEndurance
                                                        .OldWisdom = UserRS!Wisdom
                                                        .Wisdom = .OldWisdom
                                                        .OldConstitution = UserRS!Constitution
                                                        .Constitution = .OldConstitution
                                                        .OldIntelligence = UserRS!Intelligence
                                                        .Intelligence = .OldIntelligence
                                                        
                                                        
                                                        .Experience = UserRS!Experience
                                                        .StatPoints = UserRS!StatPoints
                                                        .SkillPoints = UserRS!SkillPoints
                                                        .Renown = UserRS!Renown
                                                        ST1 = UserRS!SkillLevels
                                                        If Len(ST1) = 255 Then
                                                            For A = 1 To 255
                                                                .SkillLevel(A) = Asc(Mid$(ST1, A, 1))
                                                            Next A
                                                        End If
                                                        ST1 = UserRS!SkillEXP
                                                        If Len(ST1) = 1020 Then
                                                            For A = 0 To 254
                                                                .SkillEXP(A + 1) = Asc(Mid$(ST1, A * 4 + 1, 1)) * 16777216 + Asc(Mid$(ST1, A * 4 + 2, 1)) * 65536 + Asc(Mid$(ST1, A * 4 + 3, 1)) * 256& + Asc(Mid$(ST1, A * 4 + 4, 1))
                                                            Next A
                                                        End If
                                                        .Party = 0
                                                        .IParty = 0
                                                        .StatusEffect = UserRS!StatusEffect
                                                        For A = 1 To MAXSTATUS
                                                            If A <= 32 Then
                                                                ST1 = UserRS.Fields("StatusData" + CStr(A)).Value
                                                                .StatusData(A).timer = Asc(Mid$(ST1, 1, 1)) * 16777216 + Asc(Mid$(ST1, 2, 1)) * 65536 + Asc(Mid$(ST1, 3, 1)) * 256& + Asc(Mid$(ST1, 4, 1))
                                                                .StatusData(A).Data(0) = Asc(Mid$(ST1, 5, 1))
                                                                .StatusData(A).Data(1) = Asc(Mid$(ST1, 6, 1))
                                                                .StatusData(A).Data(2) = Asc(Mid$(ST1, 7, 1))
                                                                .StatusData(A).Data(3) = Asc(Mid$(ST1, 8, 1))
                                                            Else
                                                                If IsNull(UserRS.Fields("StatusData" + CStr(A)).Value) Then
                                                                    ST1 = QuadChar(0) + Chr2(0) + Chr2(0) + Chr2(0) + Chr2(0)
                                                                Else
                                                                    ST1 = UserRS.Fields("StatusData" + CStr(A)).Value
                                                                End If
                                                                .StatusData(A).timer = Asc(Mid$(ST1, 1, 1)) * 16777216 + Asc(Mid$(ST1, 2, 1)) * 65536 + Asc(Mid$(ST1, 3, 1)) * 256& + Asc(Mid$(ST1, 4, 1))
                                                                .StatusData(A).Data(0) = Asc(Mid$(ST1, 5, 1))
                                                                .StatusData(A).Data(1) = Asc(Mid$(ST1, 6, 1))
                                                                .StatusData(A).Data(2) = Asc(Mid$(ST1, 7, 1))
                                                                .StatusData(A).Data(3) = Asc(Mid$(ST1, 8, 1))
                                                            End If
                                                        Next A
                                                        
                                                        'God Checking
                                                        #If GodChecking = True Then
                                                            If .Access > 0 Then
                                                                A = FindGodAccount(.user)
                                                                If A = 0 Then 'Doesn't Exist
                                                                    SendSocket Index, Chr2(0) + Chr2(5)
                                                                    CloseClientSocket Index
                                                                End If
                                                            End If
                                                        #End If
                                                        
                                                        'Inventory Data
                                                        For A = 1 To 20
                                                            ST1 = UserRS.Fields("InvObject" + CStr(A))
                                                            If Len(ST1) >= 10 Then
                                                                .Inv(A).Object = Asc(Mid$(ST1, 1, 1)) * 256 + Asc(Mid$(ST1, 2, 1)) 'UserRS.Fields("InvObject" + CStr(A))
                                                                .Inv(A).Value = Asc(Mid$(ST1, 3, 1)) * 16777216 + Asc(Mid$(ST1, 4, 1)) * 65536 + Asc(Mid$(ST1, 5, 1)) * 256& + Asc(Mid$(ST1, 6, 1)) 'UserRS.Fields("InvValue" + CStr(A))
                                                                .Inv(A).prefix = Asc(Mid$(ST1, 7, 1)) 'UserRS.Fields("Prefix" + CStr(A))
                                                                .Inv(A).prefixVal = Asc(Mid$(ST1, 8, 1)) 'UserRS.Fields("PrefixVal" + CStr(A))
                                                                .Inv(A).suffix = Asc(Mid$(ST1, 9, 1)) 'UserRS.Fields("Suffix" + CStr(A))
                                                                .Inv(A).SuffixVal = Asc(Mid$(ST1, 10, 1)) 'UserRS.Fields("SuffixVal" + CStr(A))
                                                                If Len(ST1) >= 12 Then
                                                                    .Inv(A).Affix = Asc(Mid$(ST1, 11, 1)) 'UserRS.Fields("Suffix" + CStr(A))
                                                                    .Inv(A).AffixVal = Asc(Mid$(ST1, 12, 1)) 'UserRS.Fields("SuffixVal" + CStr(A))
                                                                    If Len(ST1) > 12 Then
                                                                        For B = 0 To 3
                                                                            .Inv(A).Flags(B) = Asc(Mid$(ST1, 13 + B * 4, 1)) * 16777216 + Asc(Mid$(ST1, 14 + B * 4, 1)) * 65536 + Asc(Mid$(ST1, 15 + B * 4, 1)) * 256& + Asc(Mid$(ST1, 16 + B * 4, 1))
                                                                        Next B
                                                                        If Len(ST1) > 28 Then
                                                                            .Inv(A).ObjectColor = Asc(Mid$(ST1, 29, 1))
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        Next A
                                                        
                                                        For A = 1 To STORAGEPAGES
                                                            ST1 = IIf(IsNull(UserRS.Fields("StorageObjects" + CStr(A))), "", UserRS.Fields("StorageObjects" + CStr(A)))
                                                            For B = 1 To 20
                                                                If Len(ST1) >= 200 Then
                                                                    C = ((B - 1) * 10)
                                                                    .Storage(A, B).Object = Asc(Mid$(ST1, C + 1, 1)) * 256 + Asc(Mid$(ST1, C + 2, 1))
                                                                    .Storage(A, B).Value = Asc(Mid$(ST1, C + 3, 1)) * 16777216 + Asc(Mid$(ST1, C + 4, 1)) * 65536 + Asc(Mid$(ST1, C + 5, 1)) * 256& + Asc(Mid$(ST1, C + 6, 1))
                                                                    .Storage(A, B).prefix = Asc(Mid$(ST1, C + 7, 1))
                                                                    .Storage(A, B).prefixVal = Asc(Mid$(ST1, C + 8, 1))
                                                                    .Storage(A, B).suffix = Asc(Mid$(ST1, C + 9, 1))
                                                                    .Storage(A, B).SuffixVal = Asc(Mid$(ST1, C + 10, 1))
                                                                    If Len(ST1) >= 240 Then
                                                                        C = 200 + ((B - 1) * 2)
                                                                        .Storage(A, B).Affix = Asc(Mid$(ST1, C + 1, 1))
                                                                        .Storage(A, B).AffixVal = Asc(Mid$(ST1, C + 2, 1))
                                                                        If Len(ST1) > 240 Then
                                                                            C = 240 + ((B - 1) * 4)
                                                                            For D = 0 To 3
                                                                                .Storage(A, B).Flags(D) = Asc(Mid$(ST1, C + 1, 1)) * 16777216 + Asc(Mid$(ST1, C + 2, 1)) * 65536 + Asc(Mid$(ST1, C + 3, 1)) * 256& + Asc(Mid$(ST1, C + 4, 1))
                                                                            Next D
                                                                            If Len(ST1) > 560 Then
                                                                                C = 560 + ((B - 1))
                                                                                .Storage(A, B).ObjectColor = Asc(Mid$(ST1, C + 1, 1))
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                                
                                                            Next B
                                                        Next A
                                                        .NumStoragePages = IIf(IsNull(UserRS!NumStoragePages), 1, UserRS!NumStoragePages)
                                                        If .NumStoragePages = 0 Then .NumStoragePages = 1
                                                        
                                                        For A = 1 To 5
                                                            ST1 = UserRS.Fields("Equipped" + CStr(A))
                                                            If Len(ST1) >= 10 Then
                                                                .Equipped(A).Object = Asc(Mid$(ST1, 1, 1)) * 256 + Asc(Mid$(ST1, 2, 1))
                                                                .Equipped(A).Value = Asc(Mid$(ST1, 3, 1)) * 16777216 + Asc(Mid$(ST1, 4, 1)) * 65536 + Asc(Mid$(ST1, 5, 1)) * 256& + Asc(Mid$(ST1, 6, 1))
                                                                .Equipped(A).prefix = Asc(Mid$(ST1, 7, 1))
                                                                .Equipped(A).prefixVal = Asc(Mid$(ST1, 8, 1))
                                                                .Equipped(A).suffix = Asc(Mid$(ST1, 9, 1))
                                                                .Equipped(A).SuffixVal = Asc(Mid$(ST1, 10, 1))
                                                                If Len(ST1) >= 12 Then
                                                                    .Equipped(A).Affix = Asc(Mid$(ST1, 11, 1))
                                                                    .Equipped(A).AffixVal = Asc(Mid$(ST1, 12, 1))
                                                                    If Len(ST1) > 12 Then
                                                                        For B = 0 To 3
                                                                            .Equipped(A).Flags(B) = Asc(Mid$(ST1, 13 + B * 4, 1)) * 16777216 + Asc(Mid$(ST1, 14 + B * 4, 1)) * 65536 + Asc(Mid$(ST1, 15 + B * 4, 1)) * 256& + Asc(Mid$(ST1, 16 + B * 4, 1))
                                                                        Next B
                                                                        If Len(ST1) > 28 Then
                                                                            .Equipped(A).ObjectColor = Asc(Mid$(ST1, 29, 1))
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        Next A
                                                        
                                                        'Flags
                                                        ST1 = UserRS!Flags
                                                        For A = 0 To IIf(Len(ST1) = 1024, 127, 255)
                                                            With .Flag(A)
                                                                .Value = Asc(Mid$(ST1, A * 8 + 1, 1)) * 16777216 + Asc(Mid$(ST1, A * 8 + 2, 1)) * 65536 + Asc(Mid$(ST1, A * 8 + 3, 1)) * 256& + Asc(Mid$(ST1, A * 8 + 4, 1))
                                                                .ResetCounter = Asc(Mid$(ST1, A * 8 + 5, 1)) * 16777216 + Asc(Mid$(ST1, A * 8 + 6, 1)) * 65536 + Asc(Mid$(ST1, A * 8 + 7, 1)) * 256& + Asc(Mid$(ST1, A * 8 + 8, 1))
                                                            End With
                                                        Next A
                                                        
                                                        'Misc Data
                                                        .Bank = UserRS!Bank
                                                        .Status = UserRS!Status
                                                        
                                                        .Guild = 0
                                                        .GuildRank = 0
                                                        
                                                            'Find Guild
                                                            ST1 = .Name
                                                            For A = 1 To 255
                                                                With Guild(A)
                                                                    If .Name <> "" Then
                                                                        For B = 0 To 19
                                                                            If .Member(B).Name = ST1 Then
                                                                                player(Index).Guild = A
                                                                                player(Index).GuildRank = .Member(B).Rank
                                                                                Exit For
                                                                            End If
                                                                        Next B
                                                                    End If
                                                                End With
                                                            Next A
                                                        
                                                        'CalculateStats Index, False, True
                                                        '.Mana = .Mana
                                                        '.MaxMana = .MaxMana
                                                        For A = 1 To MaxPlayerTimers
                                                            .JoinRequest = 0
                                                            .ScriptTimer(A) = 0
                                                        Next A
                                                        
                                                        If CheckBan(Index) Then
                                                            CloseClientSocket Index, True
                                                        Else
                                                            .Mode = modeConnected
                                                            
                                                            'SendSocket Index, chr2(23) + chr2(Index) + chr2(.Access) + chr2(.Squelched) 'Send Misc Data
                                                            SendCharacterData Index
                                                            If World.MOTD <> "" Then
                                                                SendSocket Index, Chr2(4) + World.MOTD
                                                            End If
                                                        End If
                                                    Else
                                                        SendSocket Index, Chr2(0) + Chr2(2) 'Account already in use
                                                        AddSocketQue Index, 0
                                                    End If
                                                Else
                                                    SendSocket Index, Chr2(0) + Chr2(1) 'Invalid User/Password
                                                    AddSocketQue Index, 0
                                                End If
                                            Else
                                                SendSocket Index, Chr2(0) + Chr2(1) 'Invalid User/Password
                                                AddSocketQue Index, 0
                                            End If
                                        Else
                                            Hacker Index, "A.1"
                                        End If
                                    Else
                                        SendSocket Index, Chr2(0) + Chr2(0) + "Your client is outdated, please visit " + DownloadSite + "! Download the newest update and unzip it into your Seyerdin Online folder."
                                        AddSocketQue Index, 0
                                    End If
                                    
                                Case 29 'Pong
                                    If Len(St) > 0 Then
                                        Hacker Index, "A.2"
                                    End If
                                    
                                Case 61 'Version
                                    If Len(St) = 5 Then
                                        .ClientVer = Asc(Mid$(St, 1, 1))
                                        If .ClientVer = CurrentClientVer Then
                                            'Right Version
                                            'A = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                                            'A = Abs(A - GetTickCount)
                                            '.TickCount = A
                                        Else
                                            SendSocket Index, Chr2(0) + Chr2(6)
                                            AddSocketQue Index, 0
                                            'CloseClientSocket Index
                                        End If
                                    End If
                                Case 92 'UID
                                    .uId = St
                                Case 93 'iniUID
                                    .iniUID = St
                                Case 255 'eek
                                    If NumUsers > 1 Then
                                        SendRaw Index, CStr(NumUsers - 1) + Chr2(0)
                                    Else
                                        SendRaw Index, "0" + Chr2(0)
                                    End If
                                    CloseClientSocket Index
                                Case Else
                                    Hacker Index, "B.1"
                            End Select
                        Case modeConnected
                            Select Case PacketID
                               
                                Case 2 'Create New Character
                                    If Len(St) >= 4 Then
                                        A = InStr(4, St, Chr2(0))
                                        If A > 1 Then
                                            ST1 = Trim$(Mid$(St, 4, A - 4))
                                            For B = 1 To Len(ST1)
                                                C = Asc(Mid$(ST1, B, 1))
                                                If C = 8 Or (C >= 48 And C <= 57) Or (C >= 65 And C <= 90) Or (C >= 97 And C <= 122) Or C = 95 Then
                                                Else
                                                    CloseClientSocket Index
                                                    Exit Sub
                                                End If
                                            Next B
                                            If Len(ST1) >= 3 And Len(ST1) <= 16 And ValidName(ST1) Then
                                                UserRS.Index = "Name"
                                                UserRS.Seek "=", ST1
                                                If (UserRS.NoMatch Or UCase$(.Name) = UCase$(ST1)) And GuildNum(ST1) = 0 And NPCNum(ST1) = 0 Then
                                                    If .Level > 0 Then
                                                        UserRS.Bookmark = .Bookmark
                                                        DeleteCharacter
                                                    End If
                                                    .Class = Asc(Mid$(St, 1, 1))
                                                    If .Class = 0 Or .Class > MAX_CLASS Then
                                                        .Class = 1
                                                    End If
                                                    If Class(.Class).Enabled <> 1 Then
                                                        For A = 1 To 10
                                                            If Class(A).Enabled = 1 Then
                                                                .Class = A
                                                                A = 11
                                                            End If
                                                        Next A
                                                    End If
                                                    .Gender = Asc(Mid$(St, 3, 1))
                                                    If .Gender > 1 Then .Gender = 1
                                                    .sprite = 181 + (.Class - 1) * 8 + (.Gender * 4)
                                                    .Name = ST1
                                                    A = InStr(4, St, Chr2(0))
                                                    If Len(St) > A Then
                                                        ST1 = Mid$(St, A + 1)
                                                        If Len(ST1) > 255 Then
                                                            '.desc = Left$(st1, 255)
                                                        Else
                                                            '.desc = st1
                                                        End If
                                                    Else
                                                        '.desc = ""
                                                    End If
                                                                                                       
                                                    .Squelched = 0
                                                    .Level = 1
                                                    .Bank = 0
                                                    .Status = 2
                                                    .MaxHP = Class(.Class).StartHP
                                                    .HP = .MaxHP
                                                    .OldHP = .MaxHP
                                                    .MaxEnergy = Class(.Class).StartEnergy
                                                    .Energy = .MaxEnergy
                                                    .OldEnergy = .MaxEnergy
                                                    .MaxMana = Class(.Class).StartMana
                                                    .Mana = .MaxMana
                                                    .OldMana = .MaxMana
                                                    .strength = Class(.Class).StartStrength
                                                    .OldStrength = .strength
                                                    .Agility = Class(.Class).StartAgility
                                                    .OldAgility = .Agility
                                                    .Endurance = Class(.Class).StartEndurance
                                                    .OldEndurance = .Endurance
                                                    .Wisdom = Class(.Class).StartWisdom
                                                    .OldWisdom = .Wisdom
                                                    .Constitution = Class(.Class).StartConstitution
                                                    .OldConstitution = .Constitution
                                                    .Intelligence = Class(.Class).StartIntelligence
                                                    .OldIntelligence = .Intelligence
                                                    .Experience = 0
                                                    .StatPoints = 3
                                                    .SkillPoints = 3
                                                    .Renown = 0
                                                    For A = 1 To 255
                                                        .SkillLevel(A) = 0
                                                        .SkillEXP(A) = 0
                                                    Next A
                                                    For A = 1 To 20
                                                        .Inv(A).Object = 0
                                                    Next A
                                                    .NumStoragePages = 1
                                                    For A = 1 To STORAGEPAGES
                                                        For B = 1 To 20
                                                            .Storage(A, B).Object = 0
                                                        Next B
                                                    Next A
                                                    For A = 1 To 5
                                                        .Equipped(A).Object = 0
                                                    Next A
                                                    For A = 0 To 255
                                                        With .Flag(A)
                                                            .Value = 0
                                                            .ResetCounter = 0
                                                        End With
                                                    Next A
                                                    ClearBuff Index
                                                    .StatusEffect = 0
                                                    For A = 1 To MAXSTATUS
                                                        .StatusData(A).timer = 0
                                                    Next A
                                                    .map = World.StartLocation(0).map
                                                    If .map < 1 Then .map = 1
                                                    If .map > 5000 Then .map = 5000
                                                    .x = World.StartLocation(0).x
                                                    If .x > 11 Then .x = 11
                                                    .y = World.StartLocation(0).y
                                                    If .y > 11 Then .y = 11
                                                    .Guild = 0
                                                    .GuildRank = 0
                                                    GiveStartingEQ Index
                                                    SavePlayerData Index
                                                    SendCharacterData Index
                                                Else
                                                    SendSocket Index, Chr2(13) 'Name already in use
                                                End If
                                            Else
                                                Hacker Index, "A.78"
                                            End If
                                        Else
                                            Hacker Index, "A.4"
                                        End If
                                    Else
                                        Hacker Index, "A.5"
                                    End If
                                    
                                Case 3 'Change Password
                                    If Len(St) > 0 Then
                                        UserRS.Bookmark = .Bookmark
                                        UserRS.Edit
                                        UserRS!password = Cryp((St))
                                        UserRS.Update
                                        SendSocket Index, Chr2(5) 'Password Changed
                                    Else
                                        Hacker Index, "A.6"
                                    End If
                                    
                                Case 4 'Delete Account
                                    If Len(St) = 0 Then
                                        .Class = 0
                                        .Level = 0
                                        UserRS.Bookmark = .Bookmark
                                        DeleteAccount
                                        CloseClientSocket Index
                                    Else
                                        Hacker Index, "A.7"
                                    End If
                                    
                                Case 5 'Play
                                    If .Level > 0 Then
                                        If mapNum > 0 Then
                                            SendDataPacket Index, 1
                                        Else
                                            Hacker Index, "A.8"
                                        End If
                                    Else
                                        Hacker Index, "A.9"
                                    End If
                                    
                                Case 23 'Done receiving Data
                                    If Len(St) = 0 Then
                                        JoinGame Index
                                    Else
                                        Hacker Index, "A.10"
                                    End If
                                    
                                Case 24 'Send Next Packet
                                    If Len(St) = 3 Then
                                        A = Asc(Mid$(St, 1, 1))
                                        B = Asc(Mid$(St, 2, 1)) * 256 + Asc(Mid$(St, 3, 1))
                                        Select Case A
                                            Case 1
                                                SendDataPacket Index, B
                                            Case 2
                                                SendItemPacket Index, B
                                        End Select
                                    Else
                                        Hacker Index, "A.11"
                                    End If
                                    
                                Case 45 'Request Map
                                    If Len(St) = 2 And .Access > 0 Then
                                        A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                                        MapRS.Seek "=", A
                                        If MapRS.NoMatch Then
                                            SendSocket Index, Chr2(21) + String$(2379, Chr2(0))
                                        Else
                                            SendSocket Index, Chr2(21) + MapRS!Data
                                        End If
                                    Else
                                        Hacker Index, "A.12"
                                    End If
                                Case 46 'Edit Map
                                    If Len(St) = 2376 And .Access >= 10 Then
                                        A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                                        If A >= 1 And A <= 5000 Then
                                            mapNum = A
                                            St = Mid$(St, 2)
                                            LogCommand Index, "Edited map " & mapNum
                                            MapRS.Seek "=", mapNum
                                            If MapRS.NoMatch Then
                                                MapRS.AddNew
                                                MapRS!Number = mapNum
                                            Else
                                                MapRS.Edit
                                            End If
                                            MapRS!Data = St
                                            MapRS.Update
                                            LoadMap mapNum, St
                                            For A = 0 To 9
                                                map(mapNum).Door(A).Att = 0
                                                map(mapNum).Door(A).Wall = 0
                                            Next A
                                            For A = 1 To currentMaxUser
                                                With player(A)
                                                    If .Mode = modePlaying And .map = mapNum Then
                                                        Partmap A
                                                        .map = mapNum
                                                        JoinMap A
                                                    End If
                                                End With
                                            Next A
                                        End If
                                    Else
                                        Hacker Index, "A.22"
                                    End If
                                Case 47 'Free Map
                                    If Len(St) = 3 And .Access >= 10 Then
                                        ST1 = ""
                                        B = 0
                                        For A = GetInt(Mid$(St, 1, 2)) To 5000
                                            MapRS.Seek "=", A
                                            If MapRS.NoMatch Then
                                                ST1 = ST1 + DoubleChar(CInt(A))
                                                B = B + 1
                                                If B >= Asc(Mid$(St, 3, 1)) Then
                                                    Exit For
                                                End If
                                            End If
                                        Next A
                                    End If
                                    
                                    SendSocket Index, Chr2(255) + ST1
                                    
                                Case 29 'Pong
                                    If Len(St) > 4 Then
                                        Hacker Index, "A.13"
                                    End If
                                    
                                Case 30 'Close Connection
                                    CloseClientSocket Index
                                    
                                Case 255 'Users online
                                    SendSocket Index, NumUsers
                                    CloseClientSocket Index
                                    
                                Case Else
                                    Hacker Index, "B.2"
                            End Select
                        End Select
                End If
                GoTo LoopRead
            End If
        End If
        .SocketData = SocketData
    End With
    

Exit Sub
Error_Handler:
Open App.Path + "/LOG.TXT" For Append As #1
    ST1 = ""
    If Len(St) > 0 Then
        B = Len(St)
        For A = 1 To B
            ST1 = ST1 & Asc(Mid$(St, A, 1)) & "-"
        Next A
    End If
    Print #1, player(Index).Name & "/" & Err.Number & "/" & Err.Description & "/" & PacketID & "/" & Len(St) & "/" & St & "/" & ST1 & "/" & player(Index).Mode & "/" & "/" & player(Index).AttackSkill & "/ modreadclient1 - "
Close #1
Unhook
End
End Sub
Public Sub ReceiveData2(Index As Long, header As Long, St As String)
Dim A As Long, B As Long, C As Long, D As Long, E As Long, F As Long, G As Long, H As Long, i As Long, J As Long, ST1 As String, st2 As String, L As Long, M As Long
Dim mapNum As Long
mapNum = player(Index).map

On Error GoTo Error_Handler
With player(Index)
    Select Case header
        Case 26 'Attack Monster
            If GetTickCount + 10000 > .combatTimer Then .combatTimer = GetTickCount + 10000
            If Len(St) = 1 Then
                If .Frozen = 0 Then
                    If ExamineBit(map(mapNum).Flags(0), 5) = False Then
                        If (GetTickCount - .AttackCount < 0) Then .AttackCount = 0
                        If GetTickCount - .AttackCount >= (.AttackSpeed - 75) Then
                            .AttackCount = GetTickCount
                            A = Asc(Mid$(St, 1, 1))
                            If A <= 9 Then
                                If map(mapNum).monster(A).monster > 0 Then
                                    If Sqr((CSng(map(mapNum).monster(A).x) - CSng(.x)) ^ 2 + (CSng(map(mapNum).monster(A).y) - CSng(.y)) ^ 2) <= 2 * ((monster(map(mapNum).monster(A).monster).Flags2 And MONSTER_LARGE) + 1) Then
                                        If .AttackSkill = SKILL_INVALID Then
                                        
                                            'STATUS: INVISIBLE
                                            If GetStatusEffect(Index, SE_INVISIBLE) Then
                                                RemoveStatusEffect Index, SE_INVISIBLE
                                            End If
                                            map(mapNum).monster(A).Target = Index
                                            map(mapNum).monster(A).TargType = TargTypePlayer
                                            With monster(map(mapNum).monster(A).monster)
                                                If Int(Rnd * 255) > .Agility / 2 Then
                                                    'Hit Target
                                                    B = 0
                                                    C = PlayerDamage(Index)
                                                    
                                                    D = PlayerCritDamage(Index, C)
                                                    If C <> D Then
                                                        C = D
                                                        D = 1
                                                    Else
                                                        D = 0
                                                    End If
                                                    
                                                    C = C - .Armor
                                                    If C < 0 Then C = 0
                                                    Parameter(0) = Index
                                                    Parameter(1) = A
                                                    Parameter(2) = C
                                                    Parameter(3) = AT_MELEE
                                                    Parameter(4) = D
                                                    C = RunScript("ATTACKMONSTER")
                                                    If C = -1 Then Exit Sub
                                                    'SKILL: DESOLATION
                                                    If C > 9999 Then C = 9999
                                                Else
                                                    'Missed
                                                    B = 1
                                                    C = 0
                                                    D = 0
                                                End If
                                            End With

                                            SendRaw Index, DoubleChar(6) + Chr2(44) + Chr2(B) + Chr2(D) + Chr2(A) + DoubleChar(CInt(C)) + DoubleChar(3) + Chr2(47) + DoubleChar$(CInt(.Energy))
                                            SendToMapAllBut mapNum, Index, Chr2(44) + Chr2(B + D) + Chr2(A) + DoubleChar(CInt(C))
                                            SendToMapAllBut mapNum, Index, Chr2(42) + Chr2(Index) + Chr(0)
                                            With map(mapNum).monster(A)
                                                If B = 0 Then
                                                    If D Then
                                                        CreateSizedFloatingText mapNum, .x, .y, 12, CStr(C), critsize, 0
                                                        'CreateFloatingText mapNum, .x, .y, 12, "Critical Hit! - " & CStr(C)
                                                    End If
                                                Else
                                                    CreateFloatingEvent mapNum, .x, .y, FT_MISS
                                                End If
                                                If .HP > C Then
                                                    .HP = .HP - C
                                                    'If player(Index).StatusData(SE_FIERYESSENCE).timer > 0 And Int(Rnd * 10) < 1 Then
                                                    '    Parameter(0) = Index
                                                    '    RunScript ("SPELL" & SKILL_FIERYESSENCE)
                                                    'End If
                                                    'If player(Index).SkillLevel(SKILL_BURNINGSOUL) > 0 And Int(Rnd * 2) < 1 Then
                                                    '    Parameter(0) = Index
                                                    '    RunScript ("SPELL" & SKILL_BURNINGSOUL)
                                                    'End If
                                                    If player(Index).Leech > 0 Then
                                                        G = Int((C * player(Index).Leech) / 100)
                                                        If G > 0 Then
                                                            If Not player(Index).HP = player(Index).MaxHP Then
                                                                If player(Index).HP + G > player(Index).MaxHP Then
                                                                    SetPlayerHP Index, player(Index).MaxHP
                                                                Else
                                                                    SetPlayerHP Index, player(Index).HP + G
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    If player(Index).Leech > 0 Then
                                                        G = Int((.HP * player(Index).Leech) / 100)
                                                        If G > 0 Then
                                                            If Not player(Index).HP = player(Index).MaxHP Then
                                                                If player(Index).HP + G > player(Index).MaxHP Then
                                                                    SetPlayerHP Index, player(Index).MaxHP
                                                                Else
                                                                    SetPlayerHP Index, player(Index).HP + G
                                                                End If
                                                            End If
                                                        End If
                                                    End If
    
                                                        If player(Index).Buff.Type = BUFF_NECROMANCY Then
                                                            D = player(Index).Buff.Data(0)
                                                            If monster(.monster).Level < player(Index).Level Then
                                                                D = D - (player(Index).Level - monster(.monster).Level)
                                                            End If
                                                            If D > 0 Then
                                                                If player(Index).HP < player(Index).MaxHP Then
                                                                    player(Index).HP = player(Index).HP + D
                                                                    If player(Index).HP > player(Index).MaxHP Then player(Index).HP = player(Index).MaxHP
                                                                    SendSocket Index, Chr2(46) + DoubleChar$(player(Index).HP)
                                                                End If
                                                                SendToMap mapNum, Chr2(111) + Chr2(Index) + Chr2(SKILL_NECROMANCY) + Chr2(A)
                                                            End If
                                                       
    
                                                            'SKILL: Evocation
                                                            If player(Index).SkillLevel(SKILL_EVOCATION) > 0 Then
                                                                D = 1 + player(Index).SkillLevel(SKILL_EVOCATION) + ((player(Index).Wisdom - player(Index).WisMod(1)) \ 20)
                                                                If monster(.monster).Level < player(Index).Level Then
                                                                    D = D - (player(Index).Level - monster(.monster).Level)
                                                                End If
                                                                If D > 0 Then
                                                                    If player(Index).Mana < player(Index).MaxMana Then
                                                                        player(Index).Mana = player(Index).Mana + D
                                                                        If player(Index).Mana > player(Index).MaxMana Then player(Index).Mana = player(Index).MaxMana
                                                                        SendSocket Index, Chr2(48) + DoubleChar$(player(Index).Mana)
                                                                    End If
                                                                    SendToMap mapNum, Chr2(111) + Chr2(Index) + Chr2(SKILL_EVOCATION) + Chr2(A)
                                                                End If
                                                            End If
                                                        End If
                                                        
'
                                                        GainExp Index, CLng(monster(.monster).Experience), True, False, monster(.monster).Level
                                                        
                                                        If F > 0 Then
                                                            CreateFloatingTextAllBut Index, mapNum, .x, .y, 12, CStr(C)
                                                        Else
                                                            CreateFloatingTextAllBut Index, player(Index).map, player(Index).x, player(Index).y, 14, CStr(0)
                                                        End If

                                                    For D = 0 To 2
                                                        If Int(Rnd * 100 + 1) <= monster(.monster).Chance(D) Then
                                                            E = monster(.monster).Object(D)
                                                            If E > 0 Then
                                                                NewMapObject mapNum, E, monster(.monster).Value(D), CLng(.x), CLng(.y), False, GLOBAL_MAGIC_DROP_RATE + player(Index).MagicBonus
                                                            End If
                                                        End If
                                                    Next D
                                                    Parameter(0) = Index
                                                    Parameter(1) = 0
                                                    Parameter(2) = A
                                                    Parameter(3) = mapNum
                                                    RunScript "MONSTERDIE" + CStr(.monster)
                                                    RunScript "MONSTERDIE"
                                                    'Monster Died
                                                    SendToMapAllBut mapNum, Index, Chr2(39) + Chr2(A) 'Monster Died
                                                    SendSocket Index, Chr2(51) + Chr2(A) + QuadChar(F) 'You killed monster
                                                    .monster = 0
                                                End If
                                                    If player(Index).SkillLevel(SKILL_FIERYESSENCE) > 0 And player(Index).StatusData(SE_FIERYESSENCE).timer > 0 And Int(Rnd * 10) < 1 Then
                                                        Parameter(0) = Index
                                                        RunScript ("SPELL" & SKILL_FIERYESSENCE)
                                                    End If
                                                    If player(Index).SkillLevel(SKILL_BURNINGSOUL) > 0 And Int(Rnd * 10) < 1 Then
                                                        Parameter(0) = Index
                                                        RunScript ("SPELL" & SKILL_BURNINGSOUL)
                                                    End If
                                            End With
                                        Else
                                            AttackSkillMonster CByte(Index), CByte(A)
                                        End If
                                    Else
                                        SendSocket Index, Chr2(16) + Chr2(7) 'Too far away
                                    End If
                                Else
                                    SendSocket Index, Chr2(16) + Chr2(5) 'No such monster
                                End If
                            End If
                        End If
                    Else
                        SendSocket Index, Chr2(16) + Chr2(12) 'Can't attack monsters here
                    End If
                End If
            Else
                Hacker Index, "A.48"
            End If
        Case 72 'Cast Spell
            If Len(St) > 0 Then
                A = Asc(Mid$(St, 1, 1))
                PrintDebug player(Index).user & "(" & player(Index).Name & ")" & " cast spell " & A & "   class: " & Class(.Class).Name
                If Len(St) > 1 Then
                    St = Mid$(St, 2)
                Else
                    St = ""
                End If
                UseSkill CByte(Index), CByte(A), St
            End If
            If GetTickCount + 10000 > .combatTimer Then .combatTimer = GetTickCount + 10000
        Case 84 'attack swing weapon
            
            Parameter(0) = Index
            Parameter(1) = .map
            Parameter(2) = .x
            Parameter(3) = .y
            Parameter(4) = .D
            A = RunScript("SWINGWEAPON")
            If player(Index).SkillLevel(SKILL_FIERYESSENCE) > 0 And player(Index).StatusData(SE_FIERYESSENCE).timer > 0 And Int(Rnd * 10) < 1 Then
                Parameter(0) = Index
                RunScript ("SPELL" & SKILL_FIERYESSENCE)
            End If
            If player(Index).SkillLevel(SKILL_BURNINGSOUL) > 0 And Int(Rnd * 10) < 1 Then
                Parameter(0) = Index
                RunScript ("SPELL" & SKILL_BURNINGSOUL)
            End If
            If Not A Then SendToMapAllBut .map, Index, Chr2(42) + Chr2(Index) + Chr2(1)
            If GetTickCount + 5000 > .combatTimer Then .combatTimer = GetTickCount + 5000
        Case 29 'Pong
            If Len(St) = 4 Then
                'A = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                'A = Abs(A - GetTickCount)
                'If Abs(.TickCount - A) > 30000 Then
                '    Hacker Index, "S.H"
                'End If
            Else
                Hacker Index, "A.51"
            End If
        Case 45 'Request Map
            If Len(St) = 0 Then
                MapRS.Seek "=", .map
                If MapRS.NoMatch Then
                    SendSocket Index, Chr2(21) + String$(2379, Chr2(0))
                Else
                    A = Len(MapRS!Data)
                    SendSocket Index, Chr2(21) + MapRS!Data
                End If
            Else
                Hacker Index, "A.66"
            End If
            
        Case 28 'Describe
            If Len(St) >= 1 And Len(St) <= 255 Then
                '.desc = St
            Else
                Hacker Index, "A.50"
            End If
            

        Case 30 'Train
            If Len(St) = 6 Then
                A = Asc(Mid$(St, 1, 1))
                B = Asc(Mid$(St, 2, 1))
                C = Asc(Mid$(St, 3, 1))
                D = Asc(Mid$(St, 4, 1))
                E = Asc(Mid$(St, 5, 1))
                F = Asc(Mid$(St, 6, 1))
                If A + B + C + D + E + F <= .StatPoints Then
                    If CLng(.OldStrength) + A <= 50 Then
                        .OldStrength = .OldStrength + A
                        .StatPoints = .StatPoints - A
                    End If
                    If CLng(.OldAgility) + B <= 50 Then
                        .OldAgility = .OldAgility + B
                        .StatPoints = .StatPoints - B
                    End If
                    If CLng(.OldEndurance) + C <= 50 Then
                        .OldEndurance = .OldEndurance + C
                        .StatPoints = .StatPoints - C
                    End If
                    If CLng(.OldIntelligence) + D <= 50 Then
                        .OldIntelligence = .OldIntelligence + D
                        .StatPoints = .StatPoints - D
                    End If
                    If CLng(.OldWisdom) + E <= 50 Then
                        .OldWisdom = .OldWisdom + E
                        .StatPoints = .StatPoints - E
                    End If
                    If CLng(.OldConstitution) + F <= 50 Then
                        .OldConstitution = .OldConstitution + F
                        .StatPoints = .StatPoints - F
                    End If
                    CalculateStats Index
                    SendSocket Index, Chr2(116) + DoubleChar(CLng(.MaxHP)) + DoubleChar(CLng(.MaxMana)) + DoubleChar(CLng(.MaxEnergy))
                Else
                    Hacker Index, "A.52"
                End If
            ElseIf Len(St) = 2 Then
                Select Case Asc(Mid$(St, 1, 1))
                    Case 1 'Skill Points
                        A = Asc(Mid$(St, 2, 1))
                        If A <> SKILL_INVALID And A <= MAX_SKILLS Then
                            If .SkillLevel(A) < Skills(A).MaxLevel Then
                                If .SkillPoints > 0 Then
                                    .SkillPoints = .SkillPoints - 1
                                    .SkillLevel(A) = .SkillLevel(A) + 1
                                    
                                    'If a = SKILL_MANARESERVES Then
                                    '    UseSkill Index, SKILL_MANARESERVES, ""
                                    '    CalculateStats Index
                                    'End If
                                    'If a = SKILL_COLUSSUS Then CalculateStats Index
                                    'If a = SKILL_VIGOR Then CalculateStats Index
                                    'If a = SKILL_RESTLESS Then CalculateStats Index
                                End If
                            End If
                        End If
                        CalculateStats Index
                End Select
            Else
                Hacker Index, "A.52"
            End If
            
        Case 31 'Join Guild
            If Len(St) = 0 Then
                    If .Guild = 0 Then
                        If .Level >= 10 Then
                            If .JoinRequest > 0 Then
                                If Guild(.JoinRequest).Name <> "" Then
                                    A = FreeGuildMemberNum(CLng(.JoinRequest))
                                    If A >= 0 Then
                                        B = FindInvObject(Index, CLng(World.MoneyObj), False)
                                        If B > 0 Then
                                            If .Inv(B).Value >= 1000 Then
                                                With .Inv(B)
                                                    .Value = .Value - 1000
                                                    If .Value = 0 Then .Object = 0
                                                    SendSocket Index, Chr2(17) + Chr2(B) + DoubleChar(.Object) + QuadChar(.Value) + String$(7, Chr2(0)) 'Change inv object
                                                End With
                                                .Guild = .JoinRequest
                                                .JoinRequest = 0
                                                If Guild(.Guild).sprite > 0 Then
                                                    .sprite = Guild(.Guild).sprite
                                                    SendAll Chr2(63) + Chr2(Index) + Chr2(.sprite)
                                                End If
                                                Guild(.Guild).Member(A).Name = .Name
                                                Guild(.Guild).Member(A).Rank = 0
                                                Guild(.Guild).Member(A).Deaths = 0
                                                Guild(.Guild).Member(A).Kills = 0
                                                Guild(.Guild).Member(A).Renown = .Renown
                                                Guild(.Guild).Member(A).Joined = DateValue(Now)
                                                GuildRS.Bookmark = Guild(.Guild).Bookmark
                                                GuildRS.Edit
                                                GuildRS("MemberName" + CStr(A)) = .Name
                                                GuildRS("MemberRank" + CStr(A)) = 0
                                                GuildRS("MemberRenown" + CStr(A)) = .Renown
                                                GuildRS("MemberKills" + CStr(A)) = 0
                                                GuildRS("MemberDeaths" + CStr(A)) = 0
                                                GuildRS("Joined" + CStr(A)) = DateValue(Now)
                                                GuildRS.Update
                                                
                                                SendSocket Index, Chr2(72) + Chr2(.Guild) 'Change guild
                                                SendAllBut Index, Chr2(73) + Chr2(Index) + Chr2(.Guild) 'Player changed guild
                                                For A = 0 To 9
                                                    With Guild(.Guild).Declaration(A)
                                                        SendSocket Index, Chr2(71) + Chr2(A) + Chr2(.Guild) + Chr2(.Type)
                                                    End With
                                                Next A
                                                
                                                UpdateGuildInfo CByte(.Guild)
                                                
                                            Else
                                                SendSocket Index, Chr2(16) + Chr2(15) 'Not enough money
                                            End If
                                        Else
                                            SendSocket Index, Chr2(16) + Chr2(15) 'Not enough money
                                        End If
                                    Else
                                        SendSocket Index, Chr2(16) + Chr2(17) 'Guild is full
                                    End If
                                Else
                                    .JoinRequest = 0
                                    SendSocket Index, Chr2(16) + Chr2(14) 'You have not been invited
                                End If
                            Else
                                SendSocket Index, Chr2(16) + Chr2(14) 'You have not been invited
                            End If
                        Else
                            SendSocket Index, Chr2(97) + Chr2(3) 'You must be level 10 to join
                        End If
                    Else
                        SendSocket Index, Chr2(16) + Chr2(16) 'You are already in a guild
                    End If
            Else
                Hacker Index, "A.53"
            End If
            
        Case 32 'Leave Guild
            If Len(St) = 0 Then
                    If .Guild > 0 Then
                        A = FindGuildMember(.Name, CLng(.Guild))
                        If A >= 0 Then
                            With Guild(.Guild).Member(A)
                                .Name = ""
                                .Rank = 0
                            End With
                            GuildRS.Bookmark = Guild(.Guild).Bookmark
                            GuildRS.Edit
                            GuildRS("MemberName" + CStr(A)) = ""
                            GuildRS("MemberRank" + CStr(A)) = 0
                            GuildRS.Update
                        End If
                        SendSocket Index, Chr2(72) + Chr2(0)
                        SendAllBut Index, Chr2(73) + Chr2(Index) + Chr2(0)
                        If Guild(.Guild).sprite > 0 Then
                            .sprite = .Class * 2 + .Gender - 1
                            SendAll Chr2(63) + Chr2(Index) + Chr2(.sprite)
                        End If
                        CheckGuild CLng(.Guild)
                        If Guild(.Guild).Name <> "" Then UpdateGuildInfo CByte(.Guild)
                        .Guild = 0
                    End If
            Else
                Hacker Index, "A.54"
            End If
        Case 33 'Start New Guild
            If Len(St) >= 1 And Len(St) <= 15 And ValidName(St) Then
                    A = FreeGuildNum
                    If A > 0 Then
                        UserRS.Index = "Name"
                        UserRS.Seek "=", St
                        If UserRS.NoMatch = True And GuildNum(St) = 0 And NPCNum(St) = 0 Then
                            GuildRS.AddNew
                            GuildRS!Number = A
                            With Guild(A)
                                .Name = St
                                GuildRS!Name = St
                                .Bank = 0
                                GuildRS!Bank = 0
                                .DueDate = 0
                                GuildRS!DueDate = 0
                                .Hall = 0
                                GuildRS!Hall = 0
                                .sprite = 0
                                GuildRS!sprite = 0
                                .MOTD = ""
                                GuildRS!MOTD = ""
                                .Info = ""
                                GuildRS!Info = ""
                                .Symbol1 = 0
                                GuildRS!GuildDeaths = 0
                                .Symbol2 = 0
                                GuildRS!GuildKills = 0
                                .Symbol3 = 0
                                GuildRS!Symbol = 0
                                .founded = DateValue(Now)
                                GuildRS!founded = .founded
                                For B = 0 To 9
                                    .Declaration(B).Guild = 0
                                    .Declaration(B).Type = 0
                                    GuildRS("DeclarationGuild" + CStr(B)) = 0
                                    GuildRS("DeclarationType" + CStr(B)) = 0
                                Next B
                                .Member(0).Name = player(Index).Name
                                .Member(0).Rank = 3
                                .Member(0).Renown = player(Index).Renown
                                .Member(0).Kills = 0
                                .Member(0).Deaths = 0
                                .Member(0).Joined = DateValue(Now)
                                GuildRS!MemberName0 = player(Index).Name
                                GuildRS!MemberRank0 = 3
                                GuildRS!MemberKills0 = 0
                                GuildRS!MemberDeaths0 = 0
                                GuildRS!MemberRenown0 = player(Index).Renown
                                GuildRS!Joined0 = .founded
                                    For B = 1 To 19
                                        .Member(B).Name = ""
                                        .Member(B).Rank = 0
                                        .Member(B).Kills = 0
                                        .Member(B).Deaths = 0
                                        .Member(B).Renown = 0
                                        GuildRS("MemberName" + CStr(B)) = ""
                                        GuildRS("MemberRank" + CStr(B)) = 0
                                        GuildRS("MemberKills" + CStr(B)) = 0
                                        GuildRS("MemberDeaths" + CStr(B)) = 0
                                        GuildRS("MemberRenown" + CStr(B)) = 0
                                        GuildRS("Joined" + CStr(B)) = TimeValue(Now)
                                    Next B
                                GuildRS.Update
                                GuildRS.Seek "=", A
                                Guild(A).Bookmark = GuildRS.Bookmark
                                
                                player(Index).Guild = A
                                player(Index).GuildRank = 3
                                    
                                SendAll Chr2(70) + Chr2(A) + Cryp(St) 'Guild Data
                                SendAll Chr2(136) + Chr2(A) + Chr2(1) + Chr2(1) + QuadChar(player(Index).Renown) + QuadChar(0) + QuadChar(0)
                                SendSocket Index, Chr2(80) + Chr2(A) 'Guild Created
                                SendAllBut Index, Chr2(73) + Chr2(Index) + Chr2(A) 'Player changed guild
                            End With
                        Else
                            SendSocket Index, Chr2(16) + Chr2(16) 'Name in use
                        End If
                    Else
                        SendSocket Index, Chr2(16) + Chr2(18) 'Too many guilds
                    End If
            Else
                Hacker Index, "A.82"
            End If
        Case 34 'Invite Player to Guild
            If Len(St) = 1 Then
                    .SquelchTimer = .SquelchTimer + 7
                    If .SquelchTimer > 50 Then
                        .SquelchTimer = 0

                        
                        For B = 1 To currentMaxUser
                          If player(B).ip = player(Index).ip Then player(B).Squelched = 10
                        Next B
                        SetIpSquelchTime .ip, Asc(Mid$(St, 3, 1))
                        SendIp .ip, Chr2(23) + Chr(Asc(Mid$(St, 3, 1)))
                        
                        
                        SendAll Chr2(56) + Chr2(15) + player(Index).Name + " has been autosquelched!"
                    End If
                    If .Guild > 0 And .GuildRank >= 2 And .Squelched = 0 Then
                        A = Asc(Mid$(St, 1, 1))
                        If A >= 1 And A <= MaxUsers Then
                            If player(A).Mode = modePlaying Then
                                player(A).JoinRequest = .Guild
                                SendSocket A, Chr2(77) + Chr2(.Guild) + Chr2(Index) 'Invited to join guild
                            End If
                        End If
                    End If
            Else
                Hacker Index, "A.55"
            End If
            
        Case 35 'Kick player from guild
            If Len(St) = 1 Then
                    If .Guild > 0 And .GuildRank >= 2 Then
                        A = Asc(Mid$(St, 1, 1))
                        If A <= 19 Then
                            B = .Guild
                            If Guild(B).Member(A).Rank <= .GuildRank Then
                                With Guild(B).Member(A)
                                    ST1 = .Name
                                    .Name = ""
                                    .Rank = 0
                                End With
                                GuildRS.Bookmark = Guild(B).Bookmark
                                GuildRS.Edit
                                GuildRS("MemberName" + CStr(A)) = ""
                                GuildRS("MemberRank" + CStr(A)) = 0
                                GuildRS.Update
                                A = FindPlayer(ST1)
                                If A > 0 Then
                                    With player(A)
                                        .Guild = 0
                                        .GuildRank = 0
                                        SendSocket A, Chr2(72) + Chr2(0)
                                        SendAllBut A, Chr2(73) + Chr2(A) + Chr2(0)
                                        If Guild(B).sprite > 0 Then
                                            .sprite = .Class * 2 + .Gender - 1
                                            SendAll Chr2(63) + Chr2(A) + Chr2(.sprite)
                                        End If
                                    End With
                                ElseIf Guild(B).sprite > 0 Then
                                    UserRS.Index = "Name"
                                    UserRS.Seek "=", ST1
                                    If UserRS.NoMatch = False Then
                                        A = UserRS!Class * 2 + UserRS!Gender - 1
                                        If A >= 1 And A <= 255 Then
                                            UserRS.Edit
                                            UserRS!sprite = A
                                            UserRS.Update
                                        End If
                                    End If
                                End If
                                CheckGuild B
                            End If
                        End If
                    End If
            Else
                Hacker Index, "A.56"
            End If
        
        Case 36 'Change player's rank
            If Len(St) = 2 Then
                    If .Guild > 0 And .GuildRank >= 2 Then
                        A = Asc(Mid$(St, 1, 1))
                        B = Asc(Mid$(St, 2, 1))
                        D = .Guild
                        If A <= 19 And B <= .GuildRank Then
                            If Guild(D).Member(A).Rank <= .GuildRank Then
                                With Guild(D).Member(A)
                                    If .Name <> "" Then
                                        If B = 1 Then
                                            B = .Rank + 1
                                        Else
                                            B = .Rank - 1
                                        End If
                                            If B < 0 Then B = 0
                                            If B > player(Index).GuildRank Then B = player(Index).GuildRank
                                        .Rank = B
                                        C = FindPlayer(.Name)
                                        If C > 0 Then
                                            player(C).GuildRank = B
                                            SendSocket C, Chr2(76) + Chr2(B) 'Rank Changed
                                        End If
                                    End If
                                End With
                                GuildRS.Bookmark = Guild(D).Bookmark
                                GuildRS.Edit
                                GuildRS("MemberRank" + CStr(A)) = B
                                GuildRS.Update
                            End If
                        End If
                    End If
            Else
                Hacker Index, "A.57"
            End If
            
        Case 37 'Add Declaration
            If Len(St) = 2 Then
                    If .Guild > 0 And .GuildRank >= 2 Then
                        A = Asc(Mid$(St, 1, 1))
                        B = Asc(Mid$(St, 2, 1))
                        If A >= 1 And (B = 0 Or B = 1) Then
                            D = .Guild
                            C = FreeGuildDeclarationNum(D)
                            If C >= 0 Then
                                With Guild(D).Declaration(C)
                                    .Guild = A
                                    .Type = B
                                End With
                                SendToGuild D, Chr2(71) + Chr2(C) + Chr2(A) + Chr2(B)
                                
                                GuildRS.Bookmark = Guild(D).Bookmark
                                GuildRS.Edit
                                GuildRS("DeclarationGuild" + CStr(C)) = A
                                GuildRS("DeclarationType" + CStr(C)) = B
                                GuildRS.Update
                            End If
                        End If
                    End If
            Else
                Hacker Index, "A.58"
            End If

        Case 38 'Remove Declaration
            If Len(St) = 1 Then
                    If .Guild > 0 And .GuildRank >= 2 Then
                        A = Asc(Mid$(St, 1, 1))
                        If A <= 4 Then
                            B = .Guild
                            With Guild(B).Declaration(A)
                                .Guild = 0
                                .Type = 0
                            End With
                            
                            SendToGuild B, Chr2(71) + Chr2(A) + Chr2(0) + Chr2(0)
                            
                            GuildRS.Bookmark = Guild(B).Bookmark
                            GuildRS.Edit
                            GuildRS("DeclarationGuild" + CStr(A)) = 0
                            GuildRS("DeclarationType" + CStr(A)) = 0
                            GuildRS.Update
                        End If
                    End If
            Else
                Hacker Index, "A.59"
            End If
        Case 39 'View Guild Data
            If Len(St) = 1 Then
                    A = Asc(Mid$(St, 1, 1))
                    If A >= 1 Then
                        With Guild(A)
                            ST1 = Chr2(78) + Chr2(A) + Chr2(.Hall)
                            For B = 0 To 9
                                ST1 = ST1 + Chr2(.Declaration(B).Guild) + Chr2(.Declaration(B).Type)
                            Next B
                            ST1 = ST1 + DoubleChar(0)
                            ST1 = ST1 + DoubleChar(0)
                            ST1 = ST1 + DoubleChar(Len(.MOTD))
                            ST1 = ST1 + DoubleChar(Len(.Info)) + QuadChar(CLng(.founded)) + .MOTD + .Info
                            For B = 0 To 19
                                If B > 0 Then
                                    ST1 = ST1 + Chr2(0)
                                End If
                                ST1 = ST1 + Chr2(.Member(B).Rank + 1) + DoubleChar(.Member(B).Kills + 257) + DoubleChar(.Member(B).Deaths + 257) + QuadChar(.Member(B).Renown + 16843010) + QuadChar(CLng(.Member(B).Joined) + 16843010) + .Member(B).Name
                            Next B
                        End With
                        SendSocket Index, ST1
                    End If
            Else
                Hacker Index, "A.60"
            End If
        Case 40 'Pay guild balance
            If player(Index).Trading = False Then
                If Len(St) = 4 Then
                        If .Guild > 0 Then
                            'F = Map(MapNum).NPC
                            'If F > 0 Then
                                'If ExamineBit(NPC(F).Flags, 0) = True Then
                                    C = .Guild
                                    A = Asc(Mid$(St, 1, 1)) * 16777216 + Asc(Mid$(St, 2, 1)) * 65536 + Asc(Mid$(St, 3, 1)) * 256& + Asc(Mid$(St, 4, 1))
                                    If A > 0 Then
                                        B = FindInvObject(Index, CLng(World.MoneyObj), False)
                                        If B > 0 Then
                                            If .Inv(B).Value >= A Then
                                                With .Inv(B)
                                                    .Value = .Value - A
                                                    If .Value = 0 Then .Object = 0
                                                    SendSocket Index, Chr2(17) + Chr2(B) + DoubleChar(.Object) + QuadChar(.Value) + String$(7, Chr2(0)) 'Change inv object
                                                End With
                                                With Guild(C)
                                                    If CSng(.Bank) + CSng(A) >= 2147483647 Then
                                                        .Bank = 2147483647
                                                    Else
                                                        .Bank = .Bank + A
                                                    End If
                                                    
                                                    GuildRS.Bookmark = .Bookmark
                                                    GuildRS.Edit
                                                    GuildRS!Bank = .Bank
                                                    GuildRS.Update
                                                    
                                                    If .Bank >= 0 Then
                                                        SendToGuild C, Chr2(74) + QuadChar(.Bank)
                                                    Else
                                                        SendToGuild C, Chr2(74) + QuadChar(Abs(.Bank)) + QuadChar(.DueDate)
                                                    End If
                                                End With
                                            Else
                                                SendSocket Index, Chr2(16) + Chr2(15) 'Not enough money
                                            End If
                                        Else
                                            SendSocket Index, Chr2(16) + Chr2(15) 'Not enough money
                                        End If
                                    End If
                                'Else
                                '    SendSocket Index, chr2(16) + chr2(28) 'Not in a bank
                                'End If
                            'End If
                        End If
                Else
                    Hacker Index, "A.61"
                End If
            Else
                SendSocket Index, Chr2(113) + Chr2(6) + Chr2(4) 'can't do that while trading
            End If
            
        Case 41 'Guild Chat
            If Len(St) >= 1 Then
                If .Guild > 0 Then
                    SendToGuildAllBut Index, CLng(.Guild), Chr2(79) + Chr2(Index) + St
                End If
            Else
                Hacker Index, "A.62"
            End If
            
        Case 42 'Disband Guild
            If Len(St) = 0 Then
                    If .Guild > 0 And .GuildRank = 3 Then
                        DeleteGuild CLng(.Guild), 2
                    End If
            Else
                Hacker Index, "A.63"
            End If
            
        Case 43 'Buy guild hall
            If Len(St) = 0 Then
                    If .Guild > 0 And .GuildRank >= 2 Then
                        D = .Guild
                        If Guild(D).Hall = 0 Then
                            A = map(mapNum).Hall
                            If A > 0 Then
                                C = 0
                                For B = 1 To 255
                                    With Guild(B)
                                        If .Name <> "" And .Hall = A Then
                                            C = 1
                                            Exit For
                                        End If
                                    End With
                                Next B
                                If C = 0 Then
                                    With Guild(D)
                                        If CountGuildMembers(D) >= 5 Then
                                            If .Bank >= Hall(A).Price Then
                                                .Bank = .Bank - Hall(A).Price
                                                SendToGuild D, Chr2(74) + QuadChar(.Bank)
                                                .Hall = map(mapNum).Hall
                                                GuildRS.Bookmark = .Bookmark
                                                GuildRS.Edit
                                                GuildRS!Bank = .Bank
                                                GuildRS!Hall = .Hall
                                                GuildRS.Update
                                                SendToGuild D, Chr2(81) + Chr2(0)
                                            Else
                                                SendSocket Index, Chr2(16) + Chr2(24) 'Cost 20k to buy hall
                                            End If
                                        Else
                                            SendSocket Index, Chr2(16) + Chr2(26) 'Need 3 members
                                        End If
                                    End With
                                Else
                                    SendSocket Index, Chr2(16) + Chr2(22) 'Hall already owned
                                End If
                            Else
                                SendSocket Index, Chr2(16) + Chr2(21) 'Not in a hall
                            End If
                        Else
                            SendSocket Index, Chr2(16) + Chr2(23) 'Already have a hall
                        End If
                    End If
            Else
                Hacker Index, "A.64"
            End If
            
        Case 44 'Leave guild hall
            If Len(St) = 0 Then
                    If .Guild > 0 And .GuildRank >= 2 Then
                        A = .Guild
                        With Guild(A)
                            If .Hall > 0 Then
                                .Hall = 0
                                GuildRS.Bookmark = .Bookmark
                                GuildRS.Edit
                                GuildRS!Hall = 0
                                GuildRS.Update
                                SendToGuild A, Chr2(81) + Chr2(1)
                            End If
                        End With
                    End If
            Else
                Hacker Index, "A.65"
            End If
            

        Case 46 'Guild Balance
            If Len(St) = 0 Then
                    If .Guild > 0 Then
                        With Guild(.Guild)
                            If .Bank >= 0 Then
                                SendSocket Index, Chr2(74) + QuadChar(.Bank)
                            Else
                                SendSocket Index, Chr2(74) + QuadChar(Abs(.Bank)) + QuadChar(.DueDate)
                            End If
                        End With
                    End If
            Else
                Hacker Index, "A.67"
            End If
            
        Case 47 'Guild Hall Info
            If Len(St) = 0 Then
                    A = map(mapNum).Hall
                    If A >= 1 Then
                        With Hall(A)
                            C = 0
                            For B = 1 To 255
                                With Guild(B)
                                    If .Name <> "" And .Hall = A Then
                                        C = B
                                        Exit For
                                    End If
                                End With
                            Next B
                            SendSocket Index, Chr2(84) + Chr2(A) + Chr2(C) + QuadChar(Hall(A).Price) + QuadChar(Hall(A).Upkeep)
                        End With
                    Else
                        SendSocket Index, Chr2(16) + Chr2(21) 'Not in a hall
                    End If
            Else
                Hacker Index, "A.68"
            End If
        
        Case 48 'Edit Guild Hall
            If Len(St) = 1 And .Access >= 10 Then
                A = Asc(Mid$(St, 1, 1))
                If A >= 1 Then
                    With Hall(A)
                        SendSocket Index, Chr2(83) + Chr2(A) + QuadChar(.Price) + QuadChar(.Upkeep) + DoubleChar(CLng(.StartLocation.map)) + Chr2(.StartLocation.x) + Chr2(.StartLocation.y)
                    End With
                End If
            Else
                Hacker Index, "A.69"
            End If
            
        Case 49 'Upload Guild hall data
            If Len(St) >= 13 And Len(St) <= 28 And .Access >= 10 Then
                A = Asc(Mid$(St, 1, 1))
                If A >= 1 Then
                    With Hall(A)
                        .Price = Asc(Mid$(St, 2, 1)) * 16777216 + Asc(Mid$(St, 3, 1)) * 65536 + Asc(Mid$(St, 4, 1)) * 256& + Asc(Mid$(St, 5, 1))
                        .Upkeep = Asc(Mid$(St, 6, 1)) * 16777216 + Asc(Mid$(St, 7, 1)) * 65536 + Asc(Mid$(St, 8, 1)) * 256& + Asc(Mid$(St, 9, 1))
                        With .StartLocation
                            .map = Asc(Mid$(St, 10, 1)) * 256 + Asc(Mid$(St, 11, 1))
                            .x = Asc(Mid$(St, 12, 1))
                            .y = Asc(Mid$(St, 13, 1))
                        End With
                        If Len(St) >= 14 Then
                            .Name = Mid$(St, 14)
                        Else
                            .Name = ""
                        End If
                        HallRS.Seek "=", A
                        If HallRS.NoMatch = True Then
                            HallRS.AddNew
                            HallRS!Number = A
                        Else
                            HallRS.Edit
                        End If
                        HallRS!Name = .Name
                        HallRS!Price = .Price
                        HallRS!Upkeep = .Upkeep
                        HallRS!StartLocationMap = .StartLocation.map
                        HallRS!StartLocationX = .StartLocation.x
                        HallRS!StartLocationY = .StartLocation.y
                        HallRS.Update
                        
                        SendAll Chr2(82) + Chr2(A) + Cryp(.Name)
                    End With
                End If
            Else
                Hacker Index, "A.70"
            End If
            
        Case 50 'Edit NPC Data
            If Len(St) = 1 And .Access >= 6 Then
                A = Asc(Mid$(St, 1, 1))
                If A >= 1 Then
                    With NPC(A)
                        ST1 = Chr2(87) + Chr2(A)
                        For B = 0 To 9
                            With .SaleItem(B)
                                ST1 = ST1 + DoubleChar(.GiveObject) + QuadChar(.GiveValue) + DoubleChar(.TakeObject) + QuadChar(.TakeValue)
                            End With
                        Next B
                        ST1 = ST1 + .JoinText + Chr2(0) + .LeaveText + Chr2(0) + .SayText(0) + Chr2(0) + .SayText(1) + Chr2(0) + .SayText(2) + Chr2(0) + .SayText(3) + Chr2(0) + .SayText(4)
                        SendSocket Index, ST1
                    End With
                End If
            Else
                Hacker Index, "A.71"
            End If
            
        Case 51 'Upload NPC Data
            If Len(St) >= 132 And .Access >= 6 Then
                A = Asc(Mid$(St, 1, 1))
                If A >= 1 Then
                    With NPC(A)
                        .Flags = Asc(Mid$(St, 2, 1))
                        .Portrait = Asc(Mid$(St, 3, 1))
                        .sprite = Asc(Mid$(St, 4, 1))
                        .direction = Asc(Mid$(St, 5, 1))
                        For B = 0 To 9
                            With .SaleItem(B)
                                .GiveObject = Asc(Mid$(St, 6 + B * 12, 1)) * 256 + Asc(Mid$(St, 7 + B * 12, 1))
                                .GiveValue = Asc(Mid$(St, 8 + B * 12, 1)) * 16777216 + Asc(Mid$(St, 9 + B * 12, 1)) * 65536 + Asc(Mid$(St, 10 + B * 12, 1)) * 256& + Asc(Mid$(St, 11 + B * 12, 1))
                                .TakeObject = Asc(Mid$(St, 12 + B * 12, 1)) * 256 + Asc(Mid$(St, 13 + B * 12, 1))
                                .TakeValue = Asc(Mid$(St, 14 + B * 12, 1)) * 16777216 + Asc(Mid$(St, 15 + B * 12, 1)) * 65536 + Asc(Mid$(St, 16 + B * 12, 1)) * 256& + Asc(Mid$(St, 17 + B * 12, 1))
                                
                                If .GiveObject > 0 Then
                                    If Object(.GiveObject).Type = 6 Or Object(.GiveObject).Type = 11 Then
                                        If Object(.GiveObject).Data(0) > 0 Then
                                            If .GiveValue > Object(.GiveObject).Data(0) Then
                                                .GiveValue = Object(.GiveObject).Data(0)
                                            End If
                                        End If
                                    End If
                                End If
                            
                            End With
                        Next B
                        '103
                        GetSections Mid$(St, 126)
                        .Name = Word(1)
                        .JoinText = Word(2)
                        .LeaveText = Word(3)
                        .SayText(0) = Word(4)
                        .SayText(1) = Word(5)
                        .SayText(2) = Word(6)
                        .SayText(3) = Word(7)
                        .SayText(4) = Word(8)
                        NPCRS.Seek "=", A
                        If NPCRS.NoMatch = True Then
                            NPCRS.AddNew
                            NPCRS!Number = A
                        Else
                            NPCRS.Edit
                        End If
                        NPCRS!Name = .Name
                        NPCRS!Flags = .Flags
                        NPCRS!Portrait = .Portrait
                        NPCRS!sprite = .sprite
                        NPCRS!direction = .direction
                        NPCRS!JoinText = .JoinText
                        NPCRS!LeaveText = .LeaveText
                        NPCRS!SayText0 = .SayText(0)
                        NPCRS!SayText1 = .SayText(1)
                        NPCRS!SayText2 = .SayText(2)
                        NPCRS!SayText3 = .SayText(3)
                        NPCRS!SayText4 = .SayText(4)
                        For B = 0 To 9
                            With .SaleItem(B)
                                NPCRS("GiveObject" + CStr(B)) = .GiveObject
                                NPCRS("GiveValue" + CStr(B)) = .GiveValue
                                NPCRS("TakeObject" + CStr(B)) = .TakeObject
                                NPCRS("TakeValue" + CStr(B)) = .TakeValue
                            End With
                        Next B
                        NPCRS.Update
                        SendAll Chr2(85) + Chr2(A) + Chr2(.Flags) + Chr2(.Portrait) + Chr2(.sprite) + Chr2(.direction) + Cryp(.Name)
                    End With
                End If
            Else
                Hacker Index, "A.72"
            End If
            
        Case 52 '/trade
            If player(Index).Trading = False Then
                If Len(St) = 2 Then
                    A = Asc(Mid$(St, 1, 1))
                    B = Asc(Mid$(St, 2, 1))
                    If A >= 0 And A <= 11 And B >= 0 And B <= 11 Then
                        With map(mapNum).Tile(A, B)
                            C = .AttData(0)
                            If .Att = 25 Then
                                If C > 0 And C <= 255 Then
                                    Parameter(0) = Index
                                    If RunScript("NPCTRADE" & C) = 0 Then
                                        A = C
                                        With NPC(A)
                                            ST1 = Chr2(86)
                                            C = 0
                                            For B = 0 To 9
                                                With .SaleItem(B)
                                                    If Not .GiveObject = 0 Or Not .TakeObject = 0 Then C = 1
                                                    ST1 = ST1 + DoubleChar(.GiveObject) + QuadChar(.GiveValue) + DoubleChar(.TakeObject) + QuadChar(.TakeValue)
                                                End With
                                            Next B
                                            If C = 1 Then SendSocket Index, ST1
                                        End With
                                    End If
                                End If
                            End If
                        End With
                    End If
                Else
                    Hacker Index, "A.73"
                End If
            Else
                SendSocket Index, Chr2(113) + Chr2(6) + Chr2(4) 'can't do that while trading
            End If
        Case 53 'trade item(s)
            If .Trading = False Then
                If Len(St) = 3 Then
                    A = Asc(Mid$(St, 1, 1))
                    B = Asc(Mid$(St, 2, 1))
                    If A >= 0 And A <= 11 And B >= 0 And B <= 11 Then
                        If map(mapNum).Tile(A, B).Att = 25 Then
                            C = map(mapNum).Tile(A, B).AttData(0)
                            If C > 0 And C <= 255 Then
                                A = C
                                B = Asc(Mid$(St, 3, 1))
                                If B <= 9 Then
                                    With NPC(A).SaleItem(B)
                                        C = .GiveObject
                                        D = .GiveValue
                                        E = .TakeObject
                                        F = .TakeValue
                                    End With
                                    If C >= 1 And E >= 1 Then
                                        G = FindInvObject(Index, E, False)
                                        If G > 0 Then
                                            If Object(E).Type = 6 Or Object(E).Type = 11 Then
                                                If .Inv(G).Value >= F Then
                                                    H = 1
                                                Else
                                                    H = 0
                                                End If
                                            Else
                                                H = 1
                                            End If
                                            If H = 1 Then
                                                If Object(C).Type = 6 Or Object(C).Type = 11 Then
                                                 'money or ammolo
                                                    i = FindInvObject(Index, C, False)
                                                    If i = 0 Then
                                                        i = FreeInvNum(Index)
                                                        If i > 0 Then
                                                            .Inv(i).Value = 0
                                                        End If
                                                    End If
                                                Else
                                                    i = FreeInvNum(Index)
                                                End If
                                                If i > 0 Then
                                                    Parameter(0) = Index
                                                    If RunScript("BUYOBJ" & C) = 0 Then
                                                        With .Inv(G)
                                                            If Object(E).Type = 6 Or Object(E).Type = 11 Then
                                                            'money or ammo
                                                                .Value = .Value - F
                                                                If .Value = 0 Then .Object = 0
                                                            Else
                                                                .Object = 0
                                                            End If
                                                        End With
getAnother3:
                                                        With .Inv(i)
                                                            .Object = C
                                                            .prefix = 0
                                                            .prefixVal = 0
                                                            .suffix = 0
                                                            .SuffixVal = 0
                                                            .Affix = 0
                                                            .AffixVal = 0
                                                            .ObjectColor = 0
                                                            .Flags(0) = 0
                                                            .Flags(1) = 0
                                                            .Flags(2) = 0
                                                            .Flags(3) = 0
                                                            Select Case Object(C).Type
                                                                Case 1, 2, 3, 4 'Weapon, Shield, Armor, Helmut
                                                                    .Value = CLng(Object(C).Data(0)) * 10
                                                                Case 6, 11 'Money
                                                                    If CDbl(.Value) + CDbl(D) >= 2147483647# Then
                                                                        .Value = 2147483647
                                                                    Else
                                                                        J = Object(C).Data(0)

                                                                        If J > 0 Then
                                                                            If .Value = J Then
                                                                                If i <> 20 Then
                                                                                    M = FindInvObject(Index, C, False, i + 1)
                                                                                    If M = i Then
                                                                                        i = FreeInvNum(Index)
                                                                                        player(Index).Inv(i).Value = 0
                                                                                    Else
                                                                                        i = M
                                                                                    End If
                                                                                    M = 0
                                                                                Else
                                                                                    i = 0
                                                                                End If
                                                                            If i <> 0 Then
                                                                                GoTo getAnother3
                                                                            Else
                                                                                i = FindInvObject(Index, C, False)
                                                                            End If
                                                                            End If
                                                                        
                                                                            If .Value + D > J Then
                                                                                
                                                                                L = J - .Value
                                                                                .Value = J
                                                                                D = D - L
                                                                                M = FreeInvNum(Index)
                                                                                If M = 0 Then
                                                                                    If FreeInvNum(Index) = 0 Then
                                                                                        .Value = .Value - L
                                                                                        SendSocket Index, Chr2(16) + Chr2(1) 'Inventory Full
                                                                                        With player(Index).Inv(G)
                                                                                                .Value = .Value + F
                                                                                                .Object = E
                                                                                        End With
                                                                                    Else
                                                                                        i = FreeInvNum(Index)
                                                                                        GoTo getAnother3
                                                                                    End If
                                                                                Else
                                                                                    With player(Index).Inv(M)
                                                                                        .Object = C
                                                                                        .prefix = 0
                                                                                        .prefixVal = 0
                                                                                        .suffix = 0
                                                                                        .SuffixVal = 0
                                                                                        .Affix = 0
                                                                                        .AffixVal = 0
                                                                                        .ObjectColor = 0
                                                                                        .Flags(0) = 0
                                                                                        .Flags(1) = 0
                                                                                        .Flags(2) = 0
                                                                                        .Flags(3) = 0
                                                                                        .Value = D
                                                                                    End With
                                                                                End If
                                                                            Else
                                                                                .Value = .Value + D
                                                                                If .Value = 0 Then .Value = 1
                                                                            End If
                                                                        Else
                                                                            .Value = .Value + D
                                                                            If .Value = 0 Then .Value = 1
                                                                        End If
                                                                    End If
                                                                Case Else
                                                                    .Value = 0
                                                            End Select
                                                        End With
                                                        With .Inv(G)
                                                            If .Object > 0 Then
                                                                St = DoubleChar(15) + Chr2(17) + Chr2(G) + DoubleChar(.Object) + QuadChar(.Value) + String$(7, Chr2(0))
                                                            Else
                                                                St = DoubleChar(15) + Chr2(17) + Chr2(G) + DoubleChar(.Object) + QuadChar(.Value) + String$(7, Chr2(0))
                                                            End If
                                                        End With
                                                        With .Inv(i)
                                                            If .Object > 0 Then
                                                                St = St + DoubleChar(15) + Chr2(17) + Chr2(i) + DoubleChar(.Object) + QuadChar(.Value) + String$(7, Chr2(0)) 'Change inv objects
                                                            Else
                                                                St = St + DoubleChar(15) + Chr2(17) + Chr2(i) + DoubleChar(.Object) + QuadChar(.Value) + String$(7, Chr2(0)) 'Change inv objects
                                                            End If
                                                        End With
                                                        If M > 0 Then
                                                            With .Inv(M)
                                                                If .Object > 0 Then
                                                                    St = St + DoubleChar(15) + Chr2(17) + Chr2(M) + DoubleChar(.Object) + QuadChar(.Value) + String$(7, Chr2(0)) 'Change inv objects
                                                                Else
                                                                    St = St + DoubleChar(15) + Chr2(17) + Chr2(M) + DoubleChar(.Object) + QuadChar(.Value) + String$(7, Chr2(0)) 'Change inv objects
                                                                End If
                                                            End With
                                                        End If
                                                        St = St + DoubleChar(3) + Chr2(96) + Chr2(0) + Chr2(8) 'Play Sound
                                                        SendRaw Index, St
                                                        CalculateStats Index
                                                    End If
                                                Else
                                                    SendSocket Index, Chr2(16) + Chr2(1) 'Inventory Full
                                                End If
                                            Else
                                                SendSocket Index, Chr2(16) + Chr2(27) 'Can't afford that
                                            End If
                                        Else
                                            SendSocket Index, Chr2(16) + Chr2(27) 'Can't afford that
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    Hacker Index, "A.74"
                End If
            Else
                SendSocket Index, Chr2(113) + Chr2(6) + Chr2(4) 'can't do that while trading
            End If
            
        Case 57 'Edit Ban
            If Len(St) = 1 And .Access >= 4 Then
                A = Asc(Mid$(St, 1, 1))
                If A >= 1 And A <= 50 Then
                    With Ban(A)
                        B = .UnbanDate - CLng(Date)
                        If B < 0 Then B = 0
                        If B > 100 Then B = 100
                        SendSocket Index, Chr2(92) + Chr2(A) + Chr2(B) + .Name + Chr2(0) + .Banner + Chr2(0) + .ip + Chr2(0) + .reason
                    End With
                End If
            Else
                Hacker Index, "A.80"
            End If
            
        Case 58 'Change Ban
            If Len(St) >= 4 And .Access >= 4 Then
                A = Asc(Mid$(St, 1, 1))
                If A >= 1 And A <= 50 Then
                    With Ban(A)
                        GetSections Mid$(St, 3)
                        .UnbanDate = CLng(Date) + Asc(Mid$(St, 2, 1))
                        .Name = Word(1)
                        .InUse = True
                        .user = Word(2)
                        .reason = Word(3)
                        If .Name = "" Then .InUse = False
                        BanRS.Seek "=", A
                        If BanRS.NoMatch Then
                            BanRS.AddNew
                            BanRS!Number = A
                        Else
                            BanRS.Edit
                        End If
                        BanRS!Name = .Name
                        BanRS!user = .user
                        BanRS!reason = .reason
                        BanRS!UnbanDate = .UnbanDate
                        BanRS.Update
                    End With
                End If
            Else
                Hacker Index, "A.81"
            End If
        
        Case 59 'Edit Script
            If Len(St) >= 1 And .Access >= 10 Then
                ScriptRS.Seek "=", St
                If ScriptRS.NoMatch = False Then
                    SendSocket Index, Chr2(94) + St + Chr2(0) + ScriptRS!source
                Else
                    SendSocket Index, Chr2(94) + St + Chr2(0)
                End If
            Else
                Hacker Index, "A.87"
            End If
        
        Case 60 'Change Script
            If Len(St) >= 3 And .Access >= 10 Then
                A = InStr(St, Chr2(0))
                If A >= 2 Then
                    B = InStr(A + 1, St, Chr2(0))
                    If B > 0 Then
                        ScriptRS.Seek "=", Left$(St, A - 1)
                        ST1 = Mid$(St, A + 1, B - A - 1)
                        st2 = Mid$(St, B + 1)
                        If ST1 = "" And st2 = "" Then
                            If ScriptRS.NoMatch = False Then
                                ScriptRS.Delete
                            End If
                        Else
                            If ScriptRS.NoMatch Then
                                ScriptRS.AddNew
                                ScriptRS!Name = Left$(St, A - 1)
                            Else
                                ScriptRS.Edit
                            End If
                            ScriptRS!source = ST1
                            ScriptRS!Data = st2
                            ScriptRS.Update
                        End If
                        
                        scriptTable.RemoveAll
                        InitScriptTable
                        
                    End If
                End If
            Else
                Hacker Index, "A.88"
            End If
        Case 61 'Talk to NPC
            If Len(St) = 2 Then
                A = Asc(Mid$(St, 1, 1))
                B = Asc(Mid$(St, 2, 1))
                If A >= 0 And A <= 11 And B >= 0 And B <= 11 Then
                    With map(mapNum).Tile(A, B)
                        C = .AttData(0)
                        If .Att = 25 Then
                            If C > 0 And C <= 255 Then
                                Parameter(0) = Index
                                Parameter(1) = C
                                If RunScript("NPCSAY" & C) = 0 Then
                                    With NPC(C)
                                        If Len(.SayText(.LastSay)) > 0 Then
                                            SendSocket Index, Chr2(88) + Chr2(C) + .SayText(.LastSay)
                                            .LastSay = .LastSay + 1
                                            If .LastSay = 5 Then .LastSay = 0
                                        Else
                                            .LastSay = 0
                                            If Len(.LastSay) > 0 Then
                                                SendSocket Index, Chr2(88) + Chr2(C) + .SayText(.LastSay)
                                            End If
                                        End If
                                    End With
                                End If
                            End If
                        End If
                    End With
                End If
            End If
        Case 62 'Command
            If Len(St) >= 1 Then
                GetSections St
                A = SysAllocStringByteLen(Word(1), Len(Word(1)))
                B = SysAllocStringByteLen(Word(2), Len(Word(2)))
                C = SysAllocStringByteLen(Word(3), Len(Word(3)))
                D = SysAllocStringByteLen(Word(4), Len(Word(4)))
                Parameter(0) = Index
                Parameter(1) = A
                Parameter(2) = B
                Parameter(3) = C
                Parameter(4) = D
                E = RunScript("COMMAND")
                SysFreeString D
                SysFreeString C
                SysFreeString B
                SysFreeString A
                If E = 0 Then
                    SendSocket Index, Chr2(56) + Chr2(14) + "Invalid command."
                End If
            Else
                Hacker Index, "Illegal Command Attempt"
            End If
        Case 63 'user is Away
             If Len(St) >= 2 And Len(St) <= 513 Then
                A = Index
                B = Asc(Mid$(St, 1, 1))
                If A >= 1 And B >= 1 Then
                    SendSocket B, Chr2(95) + Chr2(A) + Mid$(St, 2)
                End If
             Else
                Hacker Index, "A.89"
             End If

        Case 64 'Admin Commands
            If Len(St) > 0 Then
                Select Case Asc(Mid$(St, 1, 1))
                    Case 1 'Access Change
                        If Len(St) >= 4 Then
                            If Mid$(St, 4) = ServerAdminPass Then
                                A = Asc(Mid$(St, 2, 1))
                                If A >= 1 And A <= MaxUsers Then
                                    If player(A).Access < 11 Then
                                        AddGod A, Asc(Mid$(St, 3, 1))
                                    Else
                                        .Access = 0
                                        SendSocket Index, Chr2(65) + Chr2(0)
                                        SendAllBut Index, Chr2(91) + Chr2(Index) + Chr2(0)
                                        .Status = 0
                                    End If
                                End If
                            Else
                                Hacker Index, "Admin Command Attempt"
                            End If
                        Else
                            Hacker Index, "Illegal Admin Command Attempt"
                        End If
                End Select
            Else
                
            End If
        Case 65 'Repairing
            If .Trading = False Then
                If Len(St) >= 3 Then
                    Select Case Asc(Mid$(St, 1, 1))
                        Case 1 'NPC Repair Display
                            A = Asc(Mid$(St, 2, 1))
                            B = Asc(Mid$(St, 3, 1))
                            If A >= 0 And A <= 11 And B >= 0 And B <= 11 Then
                                If map(mapNum).Tile(A, B).Att = 25 Then
                                    C = map(mapNum).Tile(A, B).AttData(0)
                                    If C > 0 Then
                                        If NPC(C).Flags And NPC_REPAIR Then
                                            If Len(St) = 4 Then
                                                .CurrentRepairTar = Asc(Mid$(St, 4, 1))
                                                If .CurrentRepairTar >= 1 And .CurrentRepairTar <= 20 Then
                                                    If .Inv(.CurrentRepairTar).Object > 0 Then
                                                        If Not ExamineBit(Object(.Inv(.CurrentRepairTar).Object).Flags, 2) Then 'Repairable
                                                            A = GetRepairCost(Index, .CurrentRepairTar)
                                                            If A > 0 Then
                                                                ST1 = "This object doesn't need to be repaired."
                                                                B = GetObjectDur(Index, .CurrentRepairTar)
                                                                If B >= 100 Then '100% Free repair
                                                                    A = 0
                                                                    SendSocket Index, Chr2(88) + Chr2(C) + ST1
                                                                Else
                                                                    SendSocket Index, Chr2(98) + Chr2(1) + Chr2(B) + DoubleChar(.Inv(.CurrentRepairTar).Object) + QuadChar(A)
                                                                End If
                                                            Else
                                                                SendSocket Index, Chr2(88) + Chr2(C) + "I cannot repair this object!"
                                                            End If
                                                        Else
                                                            SendSocket Index, Chr2(16) + Chr2(44)
                                                        End If
                                                    End If
                                                Else
                                                    SendSocket Index, Chr2(16) + Chr2(34)
                                                End If
                                            End If
                                        Else
                                            SendSocket Index, Chr2(16) + Chr2(32)
                                        End If
                                    End If
                                End If
                            End If
                        Case 2 'NPC Repair the Object
                            A = Asc(Mid$(St, 2, 1))
                            B = Asc(Mid$(St, 3, 1))
                            If A >= 0 And A <= 11 And B >= 0 And B <= 11 Then
                                If map(mapNum).Tile(A, B).Att = 25 Then
                                    If map(mapNum).Tile(A, B).AttData(0) > 0 Then
                                        If NPC(map(mapNum).Tile(A, B).AttData(0)).Flags And NPC_REPAIR Then
                                            'If Len(St) = 4 Then
                                                If .CurrentRepairTar > 0 Then
                                                    B = .Inv(.CurrentRepairTar).Object 'Object
                                                    If B > 0 Then 'Slot isn't empty
                                                        If Not ExamineBit(Object(B).Flags, 2) Then
                                                            A = GetRepairCost(Index, .CurrentRepairTar) 'Cost
                                                            D = World.MoneyObj
                                                            C = FindInvObject(Index, D, False) 'Money Slot
                                                            If C > 0 Then 'Has money
                                                                If .Inv(C).Value >= A Then 'Has the Cash
                                                                    TakeObj Index, D, A 'Take Cash
                                                                    E = .Inv(.CurrentRepairTar).Object
                                                                    Select Case Object(E).Type
                                                                        Case 1, 2, 3, 4 'Weapon, Shield, Armor, Helmut
                                                                            If Not E > 0 Then Exit Sub
                                                                            .Inv(.CurrentRepairTar).Value = CLng(Object(E).Data(0)) * 10
                                                                        Case 8 'Ring
                                                                            If Not E > 0 Then Exit Sub
                                                                            .Inv(.CurrentRepairTar).Value = CLng(Object(E).Data(1)) * 10
                                                                    End Select
                                                                    With .Inv(.CurrentRepairTar)
                                                                        SendSocket Index, Chr2(17) + Chr2(player(Index).CurrentRepairTar) + DoubleChar(.Object) + QuadChar(.Value) + String$(7, Chr2(0))
                                                                        SendSocket Index, Chr2(17) + Chr2(player(Index).CurrentRepairTar) + DoubleChar(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor) 'New Inv Obj
                                                                    End With
                                                                    SendSocket Index, Chr2(98) + Chr2(2) + DoubleChar(B)
                                                                Else
                                                                    SendSocket Index, Chr2(16) + Chr2(33)
                                                                End If
                                                            Else
                                                                SendSocket Index, Chr2(16) + Chr2(33)
                                                            End If
                                                        End If
                                                    Else
                                                        SendSocket Index, Chr2(16) + Chr2(34)
                                                    End If
                                                End If
                                            'End If
                                        End If
                                    End If
                                End If
                            End If
                    End Select
                Else
                    Hacker Index, "R.1"
                End If
            Else
                SendSocket Index, Chr2(113) + Chr2(6) + Chr2(4) 'can't do that while trading
            End If
        Case 66 'Switch Item
            If Len(St) = 2 Then
                A = Asc(Mid$(St, 1, 1))
                B = Asc(Mid$(St, 2, 1))
                If A > 0 And A <= 20 Then
                    If B > 0 And B <= 20 Then
                        SwapObject Index, A, .Inv(A), B, .Inv(B)
                    Else
                        Hacker Index, "66.B"
                    End If
                Else
                    Hacker Index, "66.A"
                End If
            End If
            
        Case 69 'Partys
            If Len(St) >= 1 Then
                Select Case Asc(Mid$(St, 1, 1))
                    Case 1 'Creating a Party
                        If Len(St) = 2 Then
                            If .Party = 0 Then
                                C = 0
                                For A = 1 To 255
                                    B = 0
                                    For D = 1 To currentMaxUser
                                        If player(D).Party = A Then
                                            B = B + 1
                                        End If
                                    Next D
                                    If B = 0 Then
                                        C = A
                                        Exit For
                                    End If
                                Next A
                                If C > 0 Then
                                    If Asc(Mid$(St, 2, 1)) = 1 Then SendSocket Index, Chr2(16) + Chr2(35)
                                    .Party = C
                                    SendAll Chr2(105) + Chr2(Index) + Chr2(C)
                                End If
                            End If
                        End If
                    Case 2 'Leaving Party
                        If .Party <> 0 Then
                            player(Index).Party = 0
                            SendAll Chr2(105) + Chr2(Index) + Chr2(0)
                            SendSocket Index, Chr2(16) + Chr2(36)
                        End If
                    Case 3 'Invite player to party
                        If Len(St) = 2 Then
                            .SquelchTimer = .SquelchTimer + 7
                            If .SquelchTimer > 50 Then
                                .SquelchTimer = 0
                                    For B = 1 To currentMaxUser
                                      If player(B).ip = player(Index).ip Then player(B).Squelched = 5
                                    Next B
                                    SetIpSquelchTime .ip, Asc(Mid$(St, 3, 1))
                                    SendIp .ip, Chr2(23) + Chr(Asc(Mid$(St, 3, 1)))
                                    
                                    
                                    SendAll Chr2(56) + Chr2(15) + player(Index).Name + " has been autosquelched!"
                                            
                            End If
                            A = Asc(Mid$(St, 2, 1))
                            If .Party <> 0 And .Squelched = 0 And A > 0 And A < 82 Then
                                player(A).IParty = player(Index).Party
                                SendSocket A, Chr2(121) + Chr2(Index)
                            End If
                        End If
                    Case 4 'Join a party
                        If .IParty <> 0 Then
                            .Party = .IParty
                            .IParty = 0
                            SendAll Chr2(105) + Chr2(Index) + Chr2(.Party)
                            SendToPartyAllBut .Party, Index, Chr2(104) + Chr2(0) + Chr2(Index) + DoubleChar$(((CLng(.HP) * 100) \ CLng(.MaxHP))) + DoubleChar$(((CLng(.Mana) * 100) \ CLng(.MaxMana)))
'                            For A = 1 To MaxUsers
'                                If Player(A).Party = .Party Then
'                                    SendSocket Index, chr2(104) + chr2(0) + chr2(A) + DoubleChar(((CLng(Player(A).HP) * 100) \ CLng(Player(A).MaxHP)))
'                                End If
'                            Next A
                            SendSocket Index, Chr2(16) + Chr2(50)
                        Else
                            SendSocket Index, Chr2(16) + Chr2(51)
                        End If
                    Case 6 'Party Chat
                        If Len(St) > 1 Then
                            If .Party > 0 Then
                                .SquelchTimer = .SquelchTimer + 4
                                If .SquelchTimer > 50 Then
                                    .SquelchTimer = 0
                                    For B = 1 To currentMaxUser
                                      If player(B).ip = player(Index).ip Then player(B).Squelched = 5
                                    Next B
                                    SetIpSquelchTime .ip, Asc(Mid$(St, 3, 1))
                                    SendIp .ip, Chr2(23) + Chr(Asc(Mid$(St, 3, 1)))
                                    
                                    
                                    SendAll Chr2(56) + Chr2(15) + player(Index).Name + " has been autosquelched!"
                                Else
                                    For A = 1 To currentMaxUser
                                        With player(A)
                                            If .Mode = modePlaying And .Party = player(Index).Party And Index <> A Then
                                                 SendSocket A, Chr2(106) + Chr2(Index) + Mid$(St, 2)
                                            End If
                                        End With
                                    Next A
                                End If
                            End If
                        End If
                End Select
            'Else
            '    Hacker Index, "P.1"
            End If
        Case 70 'Edit Prefix
            If Len(St) = 1 Then
                A = Asc(St)
                If .Access >= 5 Then
                    If A >= 1 And A <= 255 Then
                        ST1 = ""
                        With prefix(A)
                            SendSocket Index, Chr2(107) + Chr2(A) + Chr2(.ModType) + Chr2(.Min) + Chr2(.Max) + Chr2(.Flags) + Chr2(.Strength1) + Chr2(.Strength2) + Chr2(.Weakness1) + Chr2(.Weakness2) + Chr2(.Light.Intensity) + Chr2(.Light.Radius) + Chr2(.Rarity) + .Name
                        End With
                    End If
                Else
                    Hacker Index, "Pr.2"
                End If
            Else
                Hacker Index, "Pr.1"
            End If
        Case 71 'Upload Prefix
            If Len(St) <= 27 Then
                If .Access >= 5 Then
                    A = Asc(Mid$(St, 1, 1))
                    If A >= 1 And A <= 255 Then
                        With prefix(A)
                            .ModType = Asc(Mid$(St, 2, 1))
                            .Min = Asc(Mid$(St, 3, 1))
                            .Max = Asc(Mid$(St, 4, 1))
                            .Flags = Asc(Mid$(St, 5, 1))
                            .Strength1 = Asc(Mid$(St, 6, 1))
                            .Strength2 = Asc(Mid$(St, 7, 1))
                            .Weakness1 = Asc(Mid$(St, 8, 1))
                            .Weakness2 = Asc(Mid$(St, 9, 1))
                            .Light.Intensity = Asc(Mid$(St, 10, 1))
                            .Light.Radius = Asc(Mid$(St, 11, 1))
                            .Rarity = Asc(Mid$(St, 12, 1))
                            If Len(St) >= 13 Then
                                .Name = Mid$(St, 13)
                            Else
                                .Name = ""
                            End If
                            PrefixRS.Seek "=", A
                            If PrefixRS.NoMatch Then
                                PrefixRS.AddNew
                                PrefixRS!Number = A
                            Else
                                PrefixRS.Edit
                            End If
                            PrefixRS!Name = .Name
                            PrefixRS!Data = Mid$(St, 2, 11)
                            PrefixRS.Update
                            SendAll Chr2(108) + Chr2(A) + Chr2(.Light.Intensity) + Chr2(.Light.Radius) + Chr2(.ModType) + Cryp(.Name)
                        End With
                        GeneratePrefixList
                        GenerateSuffixList
                    End If
                End If
            Else
                
            End If

        Case 73 'Player Trade System
            If Len(St) > 0 Then
                Select Case Asc(Mid$(St, 1, 1))
                    Case 1 'Request trade
                        If Len(St) = 2 Then
                            If .Trading = False Then
                                A = Asc(Mid$(St, 2, 1))
                                If A > 0 And A <= MaxUsers Then
                                    If player(A).Trading = False And player(A).Mode = modePlaying Then
                                        If player(A).map = .map Then
                                            .Trade.Trader = A
                                            .Trade.State = TRADE_STATE_INVITER
                                            player(A).Trade.State = TRADE_STATE_INVITED
                                            player(A).Trade.Trader = Index
                                            SendSocket A, Chr2(113) + Chr2(0) + Chr(Index)
                                        Else
                                            SendSocket Index, Chr2(16) + Chr2(62) 'must be on same map
                                        End If
                                    Else
                                        SendSocket Index, Chr2(16) + Chr2(57) 'Player is already trading!
                                    End If
                                Else
                                    SendSocket Index, Chr2(16) + Chr2(4)
                                End If
                            Else
                                SendSocket Index, Chr2(16) + Chr2(56) 'You are already in a trade
                            End If
                        End If
                    Case 2 'Accept Invitation
                        If Len(St) = 1 Then
                            If .Trading = False Then
                                If .Trade.State = TRADE_STATE_INVITED Then
                                    A = .Trade.Trader
                                    If A > 0 And A <= MaxUsers Then
                                        If player(A).Trade.State = TRADE_STATE_INVITER Then
                                            If player(A).Trade.Trader = Index Then
                                                If player(A).Trading = False And player(A).Mode = modePlaying Then
                                                    If player(A).map = .map Then
                                                        .Trade.State = TRADE_STATE_OPEN
                                                        player(A).Trade.State = TRADE_STATE_OPEN
                                                        .Trading = True
                                                        player(A).Trading = True
                                                        SendSocket Index, Chr2(113) + Chr2(1) + Chr2(A)
                                                        SendSocket A, Chr2(113) + Chr2(1) + Chr2(Index)
                                                    Else
                                                        SendSocket Index, Chr2(16) + Chr2(62) 'must be on same map
                                                    End If
                                                Else
                                                    SendSocket Index, Chr2(16) + Chr2(57) 'Player in trade
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                SendSocket Index, Chr2(16) + Chr2(56) 'Already in trade
                            End If
                        End If
                    Case 3 'Cancel Trade
                        If Len(St) = 1 Then
                            If .Trading Then
                                CloseTrade Index
                            End If
                        End If
                    Case 4 'Add Item to Trade
                        If Len(St) = 7 Then
                            If .Trading Then
                                If .Trade.Trader > 0 And .Trade.Trader <= MaxUsers Then
                                    If .Trade.State = TRADE_STATE_OPEN And player(.Trade.Trader).Trade.State = TRADE_STATE_OPEN Then
                                        If Index = player(.Trade.Trader).Trade.Trader Then
                                            A = Asc(Mid$(St, 2, 1))
                                            B = Asc(Mid$(St, 3, 1))
                                            If B > 0 And B <= 20 Then
                                                If .Inv(B).Object > 0 Then
                                                    If Object(.Inv(B).Object).Type = 6 Or Object(.Inv(B).Object).Type = 11 Then
                                                        If A > 0 And A <= 10 Then
                                                            If .Trade.Slot(A) = 0 Or .Trade.Slot(A) = B Then
                                                                C = GetLong(Mid$(St, 4, 4))
                                                                If C > 0 And C <= .Inv(B).Value Then
                                                                    .Trade.Slot(A) = B
                                                                    If (Object(.Inv(B).Object).Type = 6 Or Object(.Inv(B).Object).Type = 11) Then
                                                                        If Object(.Inv(B).Object).Data(0) <> 0 Then
                                                                            If C > Object(.Inv(B).Object).Data(0) Then
                                                                                C = Object(.Inv(B).Object).Data(0)
                                                                            End If
                                                                        End If
                                                                    End If
                                                                        
                                                                        
                                                                    
                                                                    .Trade.Item(A).Object = .Inv(B).Object
                                                                    .Trade.Item(A).Value = C
                                                                    With .Trade.Item(A)
                                                                        SendSocket player(Index).Trade.Trader, Chr2(113) + Chr2(3) + Chr2(A) + DoubleChar$(.Object) + QuadChar(.Value) + String$(6, 0)
                                                                        SendSocket Index, Chr2(113) + Chr2(4) + Chr2(A) + Chr2(B) + DoubleChar$(.Object) + QuadChar(.Value) + String$(6, 0)
                                                                    End With
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        For C = 1 To 10
                                                            If .Trade.Slot(C) = B Then A = 0 'Set A to invalid object (trying to trade object again)
                                                        Next C
                                                        If A > 0 And A <= 10 Then
                                                            If .Trade.Slot(A) = 0 Then
                                                                .Trade.Slot(A) = B
                                                                CopyObject .Inv(B), .Trade.Item(A)
                                                                With .Trade.Item(A)
                                                                    SendSocket player(Index).Trade.Trader, Chr2(113) + Chr2(3) + Chr2(A) + DoubleChar$(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)
                                                                    SendSocket Index, Chr2(113) + Chr2(4) + Chr2(A) + Chr2(B) + DoubleChar$(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)
                                                                End With
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Case 5 'Remove Object from Trade
                        If Len(St) = 2 Then
                            If .Trading = True Then
                                If .Trade.Trader > 0 And .Trade.Trader <= MaxUsers Then
                                    If .Trade.State = TRADE_STATE_OPEN And player(.Trade.Trader).Trade.State = TRADE_STATE_OPEN Then
                                        If Index = player(.Trade.Trader).Trade.Trader Then
                                            A = Asc(Mid$(St, 2, 1))
                                            If A > 0 And A <= 10 Then
                                                .Trade.Slot(A) = 0
                                                ClearObject .Trade.Item(A)
                                                SendSocket .Trade.Trader, Chr2(113) + Chr2(3) + Chr2(A) + String$(12, vbNullChar)
                                                SendSocket Index, Chr2(113) + Chr2(4) + Chr2(A) + Chr2(0) + String$(12, vbNullChar)
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                SendSocket Index, Chr2(16) + Chr2(56)
                            End If
                        End If
                    Case 6 'Accept Trade
                        If .Trading = True Then
                            If .Trade.Trader > 0 And .Trade.Trader <= MaxUsers Then
                                If .Trade.State = TRADE_STATE_OPEN Then
                                    If Index = player(.Trade.Trader).Trade.Trader Then
                                        .Trade.State = TRADE_STATE_ACCEPTED
                                        SendSocket .Trade.Trader, Chr2(113) + Chr2(5)
                                    End If
                                End If
                            End If
                        Else
                            SendSocket Index, Chr2(16) + Chr2(56)
                        End If
                    Case 7 'Re-Open trade
                        If .Trading = True Then
                            If .Trade.Trader > 0 And .Trade.Trader <= MaxUsers Then
                                If .Trade.State = TRADE_STATE_ACCEPTED Then
                                    If Index = player(.Trade.Trader).Trade.Trader Then
                                        .Trade.State = TRADE_STATE_OPEN
                                        player(.Trade.Trader).Trade.State = TRADE_STATE_OPEN
                                        SendSocket .Trade.Trader, Chr2(113) + Chr2(6)
                                        SendSocket Index, Chr2(113) + Chr2(6)
                                    End If
                                End If
                            End If
                        End If
                    Case 8 'Finalize Trade
                        If .Trading = True Then
                            If .Trade.State = TRADE_STATE_ACCEPTED Then
                                If .Trade.Trader > 0 And .Trade.Trader <= MaxUsers Then
                                    If Index = player(.Trade.Trader).Trade.Trader Then
                                        If player(.Trade.Trader).Trade.State = TRADE_STATE_FINISHED Then

                                            'Both sides finished, time to see if this trade is A-OK
                                            D = 0 'Trade possible
                                            B = 0 'Number of inv slots needed
                                            'Check Player 1
                                            For A = 1 To 10
                                                If .Trade.Slot(A) > 0 And .Trade.Slot(A) <= 20 Then
                                                    If Object(.Inv(.Trade.Slot(A)).Object).Type = 6 Or Object(.Inv(.Trade.Slot(A)).Object).Type = 11 Then
                                                        If FindInvObject(.Trade.Trader, .Inv(.Trade.Slot(A)).Object, False) = 0 Then
                                                            B = B + 1
                                                        Else
                                                            J = Object(.Inv(.Trade.Slot(A)).Object).Data(0)
                                                            If J > 0 Then
                                                                If .Trade.Item(A).Value + .Inv(FindInvObject(.Trade.Trader, .Inv(.Trade.Slot(A)).Object, False)).Value > J Then
                                                                    B = B + 2
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        If .Inv(.Trade.Slot(A)).Object > 0 Then
                                                            B = B + 1
                                                        End If
                                                    End If
                                                End If
                                            Next A
                                            If FreeInvTradeSlotCount(.Trade.Trader) < B Then
                                                D = 1 'Not Enough free inv slots
                                            End If
                                            
                                            If D = 0 Then
                                                B = 0 'Number of inv slots needed
                                                'Check Player 2
                                                For A = 1 To 10
                                                    If player(.Trade.Trader).Trade.Slot(A) > 0 Then
                                                    If player(.Trade.Trader).Trade.Slot(A) <= 20 Then '''
                                                        If Object(player(.Trade.Trader).Inv(player(.Trade.Trader).Trade.Slot(A)).Object).Type = 6 Or Object(player(.Trade.Trader).Inv(player(.Trade.Trader).Trade.Slot(A)).Object).Type = 11 Then
                                                            If FindInvObject(Index, player(.Trade.Trader).Inv(player(.Trade.Trader).Trade.Slot(A)).Object, False) = 0 Then
                                                                B = B + 1
                                                            Else
                                                                J = Object(player(.Trade.Trader).Inv(player(.Trade.Trader).Trade.Slot(A)).Object).Data(0)
                                                                If J > 0 Then
                                                                    If player(.Trade.Trader).Trade.Item(A).Value + player(.Trade.Trader).Inv(FindInvObject(player(.Trade.Trader).Trade.Trader, player(.Trade.Trader).Inv(player(.Trade.Trader).Trade.Slot(A)).Object, False)).Value > J Then
                                                                        B = B + 2
                                                                    End If
                                                                End If
                                                            End If
                                                        Else
                                                            If player(.Trade.Trader).Inv(player(.Trade.Trader).Trade.Slot(A)).Object > 0 Then
                                                                B = B + 1
                                                            End If
                                                        End If
                                                    End If
                                                    End If ''''
                                                Next A
                                                If FreeInvTradeSlotCount(Index) < B Then
                                                    D = 1 'Not Enough free inv slots
                                                End If
                                                
                                                If D = 0 Then
                                                    'Both players have enough free inventory slots ... now check if items are acceptable
                                                    'Player1
                                                    For A = 1 To 10
                                                        If .Trade.Slot(A) > 0 And .Trade.Slot(A) <= 20 Then
                                                            If .Inv(.Trade.Slot(A)).Object > 0 Then
                                                                If Object(.Inv(.Trade.Slot(A)).Object).Type = 6 Or Object(.Inv(.Trade.Slot(A)).Object).Type = 11 Then
                                                                    If .Trade.Item(A).Value > .Inv(.Trade.Slot(A)).Value Then
                                                                        D = 1
                                                                    End If
                                                                Else
                                                                    If CompareObject(.Trade.Item(A), .Inv(.Trade.Slot(A))) = False Then
                                                                        D = 1
                                                                    End If
                                                                End If
                                                            Else
                                                                D = 1
                                                            End If
                                                        End If
                                                    Next A
                                                    
                                                    If D = 0 Then
                                                        'Player2
                                                        With player(.Trade.Trader)
                                                            For A = 1 To 10
                                                                If .Trade.Slot(A) > 0 And .Trade.Slot(A) <= 20 Then
                                                                    If .Inv(.Trade.Slot(A)).Object > 0 Then
                                                                        If Object(.Inv(.Trade.Slot(A)).Object).Type = 6 Or Object(.Inv(.Trade.Slot(A)).Object).Type = 11 Then
                                                                            If .Trade.Item(A).Value > .Inv(.Trade.Slot(A)).Value Then
                                                                                D = 1
                                                                            End If
                                                                        Else
                                                                            If CompareObject(.Trade.Item(A), .Inv(.Trade.Slot(A))) = False Then
                                                                                D = 1
                                                                            End If
                                                                        End If
                                                                    Else
                                                                        D = 1
                                                                    End If
                                                                End If
                                                            Next A
                                                        End With
                                                        
                                                        If D = 0 Then
                                                            'Clear inventory of trade objects
                                                            For A = 1 To 10
                                                                If .Trade.Slot(A) > 0 And .Trade.Slot(A) <= 20 Then
                                                                    If .Trade.Item(A).Object > 0 Then
                                                                        If Object(.Trade.Item(A).Object).Type = 6 Or Object(.Trade.Item(A).Object).Type = 11 Then
                                                                            .Inv(.Trade.Slot(A)).Value = .Inv(.Trade.Slot(A)).Value - .Trade.Item(A).Value
                                                                            If .Inv(.Trade.Slot(A)).Value = 0 Then .Inv(.Trade.Slot(A)).Object = 0
                                                                        Else
                                                                            ClearObject .Inv(.Trade.Slot(A))
                                                                        End If
                                                                    End If
                                                                End If
                                                                With player(.Trade.Trader)
                                                                    If .Trade.Slot(A) > 0 And .Trade.Slot(A) <= 20 Then
                                                                        If .Trade.Item(A).Object > 0 Then
                                                                            If Object(.Trade.Item(A).Object).Type = 6 Or Object(.Trade.Item(A).Object).Type = 11 Then
                                                                                .Inv(.Trade.Slot(A)).Value = .Inv(.Trade.Slot(A)).Value - .Trade.Item(A).Value
                                                                                If .Inv(.Trade.Slot(A)).Value = 0 Then .Inv(.Trade.Slot(A)).Object = 0
                                                                            Else
                                                                                ClearObject .Inv(.Trade.Slot(A))
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With
                                                            Next A
                                                            For A = 1 To 10 'Give objects
                                                                If .Trade.Item(A).Object > 0 And .Trade.Slot(A) > 0 Then
                                                                    If Object(.Trade.Item(A).Object).Type = 6 Or Object(.Trade.Item(A).Object).Type = 11 Then
                                                                        B = FindInvObject(.Trade.Trader, .Trade.Item(A).Object, False)
                                                                        If B = 0 Then
                                                                            B = FreeInvNum(.Trade.Trader)
                                                                            If B > 0 Then
                                                                                player(.Trade.Trader).Inv(B).Object = .Trade.Item(A).Object
                                                                                player(.Trade.Trader).Inv(B).Value = .Trade.Item(A).Value
                                                                            End If
                                                                        Else
                                                                                J = Object(.Trade.Item(A).Object).Data(0)
                                                                                If J > 0 Then
                                                                                    B = FreeInvNum(.Trade.Trader)
                                                                                    player(.Trade.Trader).Inv(B).Object = .Trade.Item(A).Object
                                                                                    player(.Trade.Trader).Inv(B).Value = .Trade.Item(A).Value
                                                                                End If
                                                                            player(.Trade.Trader).Inv(B).Value = player(.Trade.Trader).Inv(B).Value + .Trade.Item(A).Value
                                                                        End If
                                                                    Else
                                                                        B = FreeInvNum(.Trade.Trader)
                                                                        If B > 0 Then
                                                                            CopyObject .Trade.Item(A), player(.Trade.Trader).Inv(B)
                                                                        End If
                                                                    End If
                                                                End If
                                                                
                                                                With player(.Trade.Trader)
                                                                    If .Trade.Item(A).Object > 0 And .Trade.Slot(A) > 0 Then
                                                                        If Object(.Trade.Item(A).Object).Type = 6 Or Object(.Trade.Item(A).Object).Type = 11 Then
                                                                            B = FindInvObject(Index, .Trade.Item(A).Object, False)
                                                                            If B = 0 Then
                                                                                B = FreeInvNum(Index)
                                                                                If B > 0 Then
                                                                                    player(Index).Inv(B).Object = .Trade.Item(A).Object
                                                                                    player(Index).Inv(B).Value = .Trade.Item(A).Value
                                                                                End If
                                                                            Else
                                                                            '''''''
                                                                                J = Object(.Trade.Item(A).Object).Data(0)
                                                                                If J > 0 Then
                                                                                    B = FreeInvNum(Index)
                                                                                    player(Index).Inv(B).Object = .Trade.Item(A).Object
                                                                                    player(Index).Inv(B).Value = .Trade.Item(A).Value
                                                                                Else
                                                                                    player(Index).Inv(B).Value = player(Index).Inv(B).Value + .Trade.Item(A).Value
                                                                                End If
                                                                            End If
                                                                        Else
                                                                            B = FreeInvNum(Index)
                                                                            If B > 0 Then
                                                                                CopyObject .Trade.Item(A), player(Index).Inv(B)
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End With
                                                            Next A
                                                            
                                                            SendInventory Index
                                                            SendInventory .Trade.Trader
                                                            CloseTrade Index
                                                        Else
                                                            SendSocket Index, Chr2(16) + Chr2(60)
                                                            SendSocket .Trade.Trader, Chr2(16) + Chr2(60)
                                                        End If
                                                    Else
                                                        SendSocket Index, Chr2(16) + Chr2(60)
                                                        SendSocket .Trade.Trader, Chr2(16) + Chr2(60)
                                                    End If
                                                Else
                                                    SendSocket Index, Chr2(16) + Chr2(59)
                                                    SendSocket .Trade.Trader, Chr2(16) + Chr2(58)
                                                End If
                                            Else
                                                SendSocket Index, Chr2(16) + Chr2(58)
                                                SendSocket .Trade.Trader, Chr2(16) + Chr2(59)
                                            End If
                                        Else
                                            .Trade.State = TRADE_STATE_FINISHED
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            SendSocket Index, Chr2(16) + Chr2(56)
                        End If
                End Select
            End If
        Case 74 'edit scripts
            If .Access >= 10 Then
                ST1 = ""
                If ScriptRS.BOF = False Then
                    ScriptRS.MoveFirst
                End If
                While ScriptRS.EOF = False
                    ST1 = ST1 + ScriptRS!Name + " "
                    ScriptRS.MoveNext
                Wend
                SendSocket Index, Chr2(143) + St + ST1
            Else
                Hacker Index, "A.87"
            End If
        
        
        Case 75 'Storage
            If .Trading = False Then
                If Len(St) >= 3 Then
                    A = Asc(Mid$(St, 2, 1))
                    B = Asc(Mid$(St, 3, 1))
                    If A >= 0 And A <= 11 And B >= 0 And B <= 11 Then
                        If map(mapNum).Tile(A, B).Att = 25 Then
                            If map(mapNum).Tile(A, B).AttData(0) > 0 Then
                                If NPC(map(mapNum).Tile(A, B).AttData(0)).Flags And NPC_BANK Then
                                    A = Asc(Mid$(St, 1, 1))
                                    Select Case A
                                        Case 0 'Open
                                            If Len(St) = 3 Then
                                                ST1 = vbNullString
                                                For A = 1 To STORAGEPAGES
                                                    For B = 1 To 20
                                                        With .Storage(A, B)
                                                            ST1 = ST1 + DoubleChar(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)
                                                        End With
                                                    Next B
                                                Next A
                                                .CurStoragePage = 1
                                                SendSocket Index, Chr2(117) + Chr2(0) + ST1 + Chr2(.CurStoragePage) + Chr2(.NumStoragePages)
                                            End If
                                        Case 1 'Add Item
                                            If Len(St) = 8 Then
                                                A = Asc(Mid$(St, 4, 1)) 'Item to Add
                                                If A >= 1 And A <= 20 Then
                                                    D = Asc(Mid$(St, 5, 1)) * 16777216 + Asc(Mid$(St, 6, 1)) * 65536 + Asc(Mid$(St, 7, 1)) * 256& + Asc(Mid$(St, 8, 1))
                                                    C = .Inv(A).Object
                                                    If D <= .Inv(A).Value Then
                                                        If C > 0 Then
                                                            If ExamineBit(Object(C).Flags, 4) = False Then
                                                                If Object(C).Type = 6 Or Object(C).Type = 11 Then
                                                                    H = 1
                                                                    B = FindStackableStorageObject(Index, C) 'This finds the object AND sets the CurStoragePage AND sets it
                                                                    If B = 0 Then
                                                                        B = FreeStorageNum(Index)
                                                                        E = 0
                                                                    Else
                                                                        E = 1
                                                                    End If
                                                                Else
                                                                    B = FreeStorageNum(Index)
                                                                    E = 0
                                                                    H = 0
                                                                End If
                                                                If B > 0 Then
                                                                    With .Storage(.CurStoragePage, B)
                                                                        .Object = C
                                                                        player(Index).Inv(A).Value = player(Index).Inv(A).Value - D
                                                                        If E = 1 Then
                                                                            If CDbl(.Value) + CDbl(D) > 2147483647# Then
                                                                                D = 2147483647
                                                                            Else
                                                                                D = .Value + D
                                                                            End If
                                                                        End If
                                                                        If H = 1 Then
                                                                            .Value = D
                                                                        Else
                                                                            .Value = player(Index).Inv(A).Value
                                                                            player(Index).Inv(A).Object = 0
                                                                        End If
                                                                        
                                                                        .prefix = player(Index).Inv(A).prefix
                                                                        .prefixVal = player(Index).Inv(A).prefixVal
                                                                        .suffix = player(Index).Inv(A).suffix
                                                                        .SuffixVal = player(Index).Inv(A).SuffixVal
                                                                        .Affix = player(Index).Inv(A).Affix
                                                                        .AffixVal = player(Index).Inv(A).AffixVal
                                                                        .ObjectColor = player(Index).Inv(A).ObjectColor
                                                                        For i = 0 To 3
                                                                            .Flags(i) = player(Index).Inv(A).Flags(i)
                                                                        Next i
                                                                        
                    '                                                    If Not (Object(C).Type = 6 Or Object(C).Type = 11) Then
                    '                                                        Player(Index).Inv(A).Object = 0
                    '                                                    End If
                                                                        If player(Index).Inv(A).Value < 0 Then
                                                                            Hacker Index, "75.1"
                                                                        ElseIf player(Index).Inv(A).Value = 0 Then
                                                                            player(Index).Inv(A).Object = 0
                                                                        End If
                                                                        
                                                                        SendSocket Index, Chr2(117) + Chr2(1) + Chr2(B) + DoubleChar(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor) 'Update Storage Object
                                                                        With player(Index).Inv(A)
                                                                            SendSocket Index, Chr2(17) + Chr2(A) + DoubleChar(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)  'Update inventory
                                                                        End With
                                                                    End With
                                                                Else
                                                                    SendSocket Index, Chr2(16) + Chr2(47)
                                                                End If
                                                            Else
                                                                SendSocket Index, Chr2(16) + Chr2(47)
                                                            End If
                                                        End If
                                                    End If
                                                Else
                                                    Hacker Index, "A.75-1"
                                                End If
                                            End If
                                        Case 2 'Remove Item
                                            If Len(St) = 8 Then
                                                A = Asc(Mid$(St, 4, 1))
                                                If A >= 1 And A <= 20 Then
                                                    D = Asc(Mid$(St, 5, 1)) * 16777216 + Asc(Mid$(St, 6, 1)) * 65536 + Asc(Mid$(St, 7, 1)) * 256& + Asc(Mid$(St, 8, 1))
                                                    C = .Storage(.CurStoragePage, A).Object
                                                    If C > 0 Then
                                                        If D <= .Storage(.CurStoragePage, A).Value Then
                                                            If Object(C).Type = 6 Or Object(C).Type = 11 Then 'money or Ammo
                                                                H = 1
                                                                J = Object(C).Data(0)
                                                                B = FindInvObject(Index, C, False)
                                                                If B = 0 Then
                                                                    B = FreeInvNum(Index)
                                                                    E = 2
                                                                Else
                                                                    While (.Inv(B).Value = J And B <> 20)
                                                                        L = FindInvObject(Index, C, False, B + 1)
                                                                        If L = B Then
                                                                            B = FreeInvNum(Index)
                                                                            If B > 0 Then .Inv(B).Value = 0
                                                                        Else
                                                                            B = L
                                                                        End If
                                                                    Wend
                                                                    E = 1
                                                                End If
                                                            Else
                                                                D = .Storage(.CurStoragePage, A).Value
                                                                B = FreeInvNum(Index)
                                                                E = 0
                                                                H = 0
                                                            End If
getAnother2:
                                                            If B > 0 Then
                                                                Parameter(0) = Index
                                                                If RunScript("WITHDRAWOBJ" + CStr(C)) = 0 Then
                                                                    With .Inv(B)
                                                                        If E = 1 Then
                                                                            If .Object <> player(Index).Storage(player(Index).CurStoragePage, A).Object Then .Value = 0
                                                                        End If
                                                                        .Object = C
                                                                        If E = 1 Then
                                                                            If CDbl(.Value) + CDbl(D) > 2147483647# Then
                                                                                .Value = 2147483647
                                                                            Else ''''''''''''''''''''''
                                                                                If J > 0 Then
                                                                                    If .Value + D > J Then
                                                                                        i = J - .Value
                                                                                        .Value = J
                                                                                        D = D - i
                                                                                        player(Index).Storage(player(Index).CurStoragePage, A).Value = player(Index).Storage(player(Index).CurStoragePage, A).Value - i
                                                                                        SendSocket Index, Chr2(17) + Chr2(B) + DoubleChar(CInt(C)) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)
                                                                                        B = FreeInvNum(Index)
                                                                                        
                                                                                        'If B = 0 Then
                                                                                        '    With Map(MapNum).Object(A)
                                                                                                'SendToMap Player(Index).Map, chr2(14) + chr2(A) + DoubleChar(CInt(.Object)) + chr2(.X) + chr2(.Y) + chr2(0) + chr2(0) + QuadChar(.Value) + chr2(.Prefix) + chr2(.PrefixVal) + chr2(.Suffix) + chr2(.SuffixVal)
                                                                                        '    End With
                                                                                        'End If
                                                                                        If D > J Then
                                                                                            E = 2
                                                                                        Else
                                                                                            E = 0
                                                                                        End If
                                                                                        GoTo getAnother2
                                                                                    Else
                                                                                        .Value = .Value + D
                                                                                    End If
                                                                                Else
                                                                                    .Value = .Value + D
                                                                                End If
                                                                            End If
                                                                        Else
                                                                            If E = 2 Then '''''''''''''''
                                                                                If J > 0 Then
                                                                                    If .Value > 0 Then .Value = 0
                                                                                    If .Value + D > J Then
                                                                                        i = J - .Value
                                                                                        .Value = J
                                                                                        D = D - i
                                                                                        player(Index).Storage(player(Index).CurStoragePage, A).Value = player(Index).Storage(player(Index).CurStoragePage, A).Value - i
                                                                                        SendSocket Index, Chr2(17) + Chr2(B) + DoubleChar(CInt(C)) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)
                                                                                        B = FreeInvNum(Index)
                                                                                        
                                                                                        'If B = 0 Then
                                                                                        '    With Map(MapNum).Object(A)
                                                                                                'SendToMap Player(Index).Map, chr2(14) + chr2(A) + DoubleChar(CInt(.Object)) + chr2(.X) + chr2(.Y) + chr2(0) + chr2(0) + QuadChar(.Value) + chr2(.Prefix) + chr2(.PrefixVal) + chr2(.Suffix) + chr2(.SuffixVal)
                                                                                        '    End With
                                                                                        'End If
                                                                                        If D > J Then
                                                                                            E = 2
                                                                                        Else
                                                                                            E = 0
                                                                                        End If
                                                                                        GoTo getAnother2
                                                                                    Else
                                                                                        .Value = .Value + D
                                                                                    End If
                                                                                Else
                                                                                    If .Value > 0 Then .Value = 0
                                                                                    .Value = .Value + D
                                                                                End If
                                                                            Else
                                                                                .Value = D
                                                                            End If
                                                                        End If
                                                                    End With
                                                                    
                                                                    With player(Index).Storage(player(Index).CurStoragePage, A)
                                                                        player(Index).Inv(B).prefix = .prefix
                                                                        player(Index).Inv(B).prefixVal = .prefixVal
                                                                        player(Index).Inv(B).suffix = .suffix
                                                                        player(Index).Inv(B).SuffixVal = .SuffixVal
                                                                        player(Index).Inv(B).Affix = .Affix
                                                                        player(Index).Inv(B).AffixVal = .AffixVal
                                                                        player(Index).Inv(B).Flags(0) = .Flags(0)
                                                                        player(Index).Inv(B).Flags(1) = .Flags(1)
                                                                        player(Index).Inv(B).Flags(2) = .Flags(2)
                                                                        player(Index).Inv(B).Flags(3) = .Flags(3)
                                                                        player(Index).Inv(B).ObjectColor = .ObjectColor
                                                                        .Value = .Value - D
                                                                        If .Value < 0 Then
                                                                            Hacker Index, "75.2"
                                                                        ElseIf .Value = 0 Then
                                                                            .Object = 0
                                                                        End If
                                                                        
                                                                        With player(Index).Inv(B)
                                                                            SendSocket Index, Chr2(17) + Chr2(B) + DoubleChar(CInt(C)) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor) 'New Inv Obj
                                                                        End With
                                                                        If .Object = 0 Then
                                                                            SendSocket Index, Chr2(117) + Chr2(2) + Chr2(A)
                                                                        Else
                                                                            SendSocket Index, Chr2(117) + Chr2(1) + Chr2(A) + DoubleChar(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor)  'Update Bank Object
                                                                        End If
                                                                    End With
                                                                End If
                                                            Else
                                                                With player(Index).Storage(player(Index).CurStoragePage, A)
                                                                    SendSocket Index, Chr2(117) + Chr2(1) + Chr2(A) + DoubleChar(.Object) + QuadChar(.Value) + Chr2(.prefix) + Chr2(.prefixVal) + Chr2(.suffix) + Chr2(.SuffixVal) + Chr2(.Affix) + Chr2(.AffixVal) + Chr2(.ObjectColor) 'Update Bank Object
                                                                End With
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        Case 3 'Next Page
                                            If player(Index).CurStoragePage < player(Index).NumStoragePages And player(Index).CurStoragePage < STORAGEPAGES Then
                                                player(Index).CurStoragePage = player(Index).CurStoragePage + 1
                                            End If
                                            SendSocket Index, Chr2(117) + Chr2(4) + Chr2(player(Index).CurStoragePage)
                                        Case 4 'Previous Page
                                            If player(Index).CurStoragePage > 1 Then
                                                player(Index).CurStoragePage = player(Index).CurStoragePage - 1
                                            End If
                                            SendSocket Index, Chr2(117) + Chr2(4) + Chr2(player(Index).CurStoragePage)
                                    End Select
                                    CalculateStats Index
                                End If
                            End If
                        End If
                    End If
                Else
                    Hacker Index, "75.L"
                End If
            Else
                SendSocket Index, Chr2(113) + Chr2(6) + Chr2(4) 'can't do that while trading
            End If
        Case 76 'Click
            If Len(St) > 0 Then
                A = Asc(Mid$(St, 1, 1))
                Select Case A
                    Case 0 'Click Tile / news tile
                        If Len(St) = 3 Then
                            B = Asc(Mid$(St, 2, 1))
                            C = Asc(Mid$(St, 3, 1))
                            If B >= 0 And B <= 11 And C >= 0 And C <= 11 Then
                                If map(.map).Tile(B, C).Att = 14 Then
                                    If ExamineBit(map(.map).Tile(B, C).AttData(0), 1) Then
                                        Parameter(0) = Index
                                        Parameter(1) = MC_CLICK
                                        Parameter(4) = B
                                        Parameter(5) = C
                                        RunScript "MAP" & CStr(.map) & "_" & CStr(B) & "_" & CStr(C)
                                    Else
                                        'Hacker Index, "C.5"
                                    End If
                                ElseIf map(.map).Tile(B, C).Att = 19 Then
                                    If map(.map).Tile(B, C).AttData(3) = 1 Then
                                        Parameter(0) = Index
                                        Parameter(1) = MC_CLICK
                                        Parameter(4) = B
                                        Parameter(5) = C
                                        RunScript "MAP" & CStr(.map) & "_" & CStr(B) & "_" & CStr(C)
                                    Else
                                        'Hacker Index, "C.5"
                                    End If
                                ElseIf map(.map).Tile(B, C).Att = 6 Then
                                    If ExamineBit(map(.map).Tile(B, C).AttData(3), 0) Then
                                        Parameter(0) = Index
                                        Parameter(1) = map(.map).Tile(B, C).AttData(1)
                                        Parameter(2) = map(.map).Tile(B, C).AttData(2)
                                        Parameter(3) = map(.map).Tile(B, C).AttData(3)
                                        Parameter(4) = B
                                        Parameter(5) = C
                                        RunScript ("NEWS" & map(.map).Tile(B, C).AttData(0))
                                    Else
                                        'Hacker Index, "C.5"
                                    End If
                                        
                                Else
                                    'Hacker Index, "C.1"
                                End If
                            Else
                                'Hacker Index, "C.2"
                            End If
                        Else
                            Hacker Index, "C.3"
                        End If
                    Case 1 'Attack Tile
                        If Len(St) = 3 Then
                            B = Asc(Mid$(St, 2, 1))
                            C = Asc(Mid$(St, 3, 1))
                            If ((.x - B) ^ 2 + (.y - C) ^ 2) = 1 Then
                                If B >= 0 And B <= 11 And C >= 0 And C <= 11 Then
                                    If map(.map).Tile(B, C).Att = 14 Then
                                        If ExamineBit(map(.map).Tile(B, C).AttData(0), 2) Then
                                            Parameter(0) = Index
                                            Parameter(1) = MC_ATTACK
                                            Parameter(2) = PlayerDamage(Index)
                                            Parameter(4) = B
                                            Parameter(5) = C
                                            RunScript "MAP" & CStr(.map) & "_" & CStr(B) & "_" & CStr(C)
                                        Else
                                            Hacker Index, "C.6"
                                        End If
                                    Else
                                        Hacker Index, "C.7"
                                    End If
                                Else
                                    Hacker Index, "C.8"
                                End If
                            Else
                                Hacker Index, "C.10"
                            End If
                        Else
                            Hacker Index, "C.9"
                        End If
                    Case 2 'Player
                        If Len(St) = 2 Then
                            A = Asc(Mid$(St, 2, 1))
                            If A > 0 And A <= MaxUsers Then
                                If player(A).Mode = modePlaying Then
                                    If player(A).map = .map Then
                                        Parameter(0) = Index
                                        Parameter(1) = A
                                        RunScript ("CLICKPLAYER")
                                    End If
                                End If
                            End If
                        End If
                    Case 3 'Monster
                        If Len(St) = 2 Then
                            A = Asc(Mid$(St, 2, 1))
                            If A >= 0 And A <= 9 Then
                                If map(.map).monster(A).monster > 0 Then
                                    Parameter(0) = Index
                                    Parameter(1) = A
                                    RunScript ("CLICKMONSTER")
                                End If
                            End If
                        End If
                End Select
            Else
                Hacker Index, "C.4"
            End If
        Case 77
            If Len(St) > 0 Then
                With .Widgets
                    If .MenuVisible Then
                        If .WidgetScript <> "" Then
                            'For A = 1 To .NumWidgets - 1
                            '    F = Asc(Mid$(St, 1, 1))
                            '    ST1 = Mid$(St, 2, F)
                             '   St = Mid$(St, F + 2)
                             '   B = Asc(Mid$(ST1, 1, 1)) 'Type of Widget
                             '   C = Asc(Mid$(ST1, 2, 1)) 'Length of Key
                             '   st2 = Mid$(ST1, 3, C)  'Key
                            '
                            '    If B = WIDGET_TEXTBOX Then
                            '        E = Asc(Mid$(ST1, C + 3, 1))
                            '        .Widgets(A).strData = Mid$(ST1, C + 4, E)
                            '    End If
                            '    For D = 1 To .NumWidgets
                            '        If .Widgets(D).Key = st2 Then
                            '            Select Case B
                            '                Case WIDGET_TEXTBOX
                            '                    E = Asc(Mid$(ST1, C + 3, 1))
                            '                    .Widgets(D).strData = Mid$(ST1, C + 4, E)
                            '                Case Else
                            '
                            '            End Select
                            '        End If
                            '    Next D
                            'Next A
                            'D = 1
                            For A = 1 To .NumWidgets - 1
                                If .Widgets(A).Type = WIDGET_TEXTBOX Then
                                    F = Asc(Mid$(St, 1, 1))
                                    ST1 = Mid$(St, 2, F)
                                    St = Mid$(St, F + 2)
                                    B = Asc(Mid$(ST1, 1, 1)) 'Type of Widget
                                    C = Asc(Mid$(ST1, 2, 1)) 'Length of Key
                                    st2 = Mid$(ST1, 3, C)  'Key
                                    
                                    If B = WIDGET_TEXTBOX Then
                                        E = Asc(Mid$(ST1, C + 3, 1))
                                        .Widgets(A).strData = Mid$(ST1, C + 4, E)
                                    End If
                                    'D = D + 1
                                End If
                            Next A
                            
                            
                            Parameter(0) = Index
                            A = SysAllocStringByteLen(St, Len(St))
                            Parameter(1) = A
                            RunScript .WidgetScript
                            SysFreeString (A)
                            'ReDim .Widgets(0)
                            '.NumWidgets = 0
                            '.WidgetScript = ""
                            .WidgetString = ""
                        End If
                    Else
                        'Hacker index, "W.1" 'Widget Hack 1
                    End If
                End With
            End If
        Case 78 'Script Callback
            If Len(St) = 1 Then
                If Len(.ScriptCallback) > 0 Then
                    A = Asc(Mid$(St, 1, 1))
                    Parameter(0) = Index
                    Parameter(1) = A
                    RunScript .ScriptCallback
                End If
            End If
        Case 79 'Upload Light
            If Len(St) >= 8 Then
                A = Asc(Mid$(St, 1, 1))
                If A >= 1 And A <= 255 Then
                    With Lights(A)
                        .red = Asc(Mid$(St, 2, 1))
                        .green = Asc(Mid$(St, 3, 1))
                        .blue = Asc(Mid$(St, 4, 1))
                        .Intensity = Asc(Mid$(St, 5, 1))
                        .Radius = Asc(Mid$(St, 6, 1))
                        .MaxFlicker = Asc(Mid$(St, 7, 1))
                        .FlickerRate = Asc(Mid$(St, 8, 1))
                        If Len(St) > 8 Then
                            .Name = Mid$(St, 9)
                        Else
                            .Name = ""
                        End If
                        LightsRS.Seek "=", A
                        If LightsRS.NoMatch Then
                            LightsRS.AddNew
                            LightsRS!Number = A
                        Else
                            LightsRS.Edit
                        End If
                        LightsRS!Name = .Name
                        LightsRS!Data = Mid$(St, 2, 7)
                        LightsRS.Update
                        SendAll Chr2(129) + Chr2(A) + Chr2(.red) + Chr2(.green) + Chr2(.blue) + Chr2(.Intensity) + Chr2(.Radius) + Chr2(.MaxFlicker) + Chr2(.FlickerRate) + Cryp(.Name)
                    End With
                End If
            End If
        Case 80 'request edit
            If Len(St) = 1 Then
                If player(Index).Access = 0 Then Hacker Index, "Hack: Invalid .ME"
            End If
        Case 81 'ping
            If Len(St) = 0 Then
                SendSocket Index, Chr2(134)
            End If
        Case 82 'Guild MOTD
            If player(Index).Guild > 0 Then
                Guild(player(Index).Guild).MOTD = St
                SendToGuild CInt(player(Index).Guild), Chr2(56) + Chr2(15) + "Guild MOTD Update:"
                SendToGuild CInt(player(Index).Guild), Chr2(135) + Chr2(15) + St
                
                GuildRS.Bookmark = Guild(player(Index).Guild).Bookmark
                GuildRS.Edit
                GuildRS!MOTD = Guild(player(Index).Guild).MOTD
                GuildRS.Update
            End If
        Case 83 'Guild Info
            If player(Index).Guild > 0 Then
                Guild(player(Index).Guild).Info = St
                SendToGuild CInt(player(Index).Guild), Chr2(135) + Chr2(15) + "Guild Info Updated"
                GuildRS.Bookmark = Guild(player(Index).Guild).Bookmark
                GuildRS.Edit
                GuildRS("Info") = St
                GuildRS.Update
            End If

        Case 85 'edit guild door data
        If Len(St) >= 5 Then
            A = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
            ScriptLoadMap A
            With (CurEditMap)
                A = Asc(Mid$(St, 3, 1))
                B = Asc(Mid$(St, 4, 1))
                C = Asc(Mid$(St, 5, 1))
                F = 0
                If player(Index).GuildRank = 3 Then
                    For D = 0 To 11
                        For E = 0 To 11
                            If .Tile(D, E).Att = 8 And .Tile(D, E).AttData(2) > 0 Then
                                If (C = 3 And ExamineBit(.Tile(D, E).AttData(3), 5)) Then
                                    SendSocket Index, Chr2(16) + Chr2(66)
                                    Exit Sub
                                Else
                                    If .Tile(D, E).AttData(0) = A And .Tile(D, E).AttData(1) = B And .Tile(D, E).AttData(2) = Guild(player(Index).Guild).Hall Then
                                        If ExamineBit(.Tile(D, E).AttData(3), 0) Then L = 1
                                        If ExamineBit(.Tile(D, E).AttData(3), 1) Then L = L + 2
                                        L = L + 2 ^ (C + 2)
                                        .Tile(D, E).AttData(3) = L
                                        F = 1
                                    End If
                                End If
                            End If
                        Next E
                    Next D
                    If F = 1 Then
                        ScriptSaveMap
                    End If
                End If
            End With
        Else
            Hacker Index, "Invalid Guild Door Edit"
        End If
        Case 86
            
            If Asc(Mid$(St, 1, 1)) = 2 Then
                SendToGods Chr2(30) + player(Index).Name + "'s client is registering over 2000 ticks per second for 3 or more seconds."
            Else
                player(Index).speedhack = player(Index).speedhack + 1
                SendToGods Chr2(30) + player(Index).Name + " speedhack detect! [" + player(Index).ip + "]"
                PrintLog player(Index).Name & " speedhack detect! [" & player(Index).ip & "]"
                ScriptPrintReport player(Index).Name + " speedhack detect! [" + player(Index).ip + "]", "speedhacking"
                
                If player(Index).speedhack > 1 Then BootPlayer Index, 0, "Speedhacking"
            End If
        Case 87 ' i was hit by a map projectile!
            With map(player(Index).map).Projectile(Asc(Mid$(St, 1, 1)))
                If .sprite > 0 Then
                    If IsPlaying(Index) Then
                        If player(Index).map = mapNum Then
                                .sprite = 0
                                DamagePlayer Index, .damage, .magical, .damageString
                        End If
                    End If
                End If
            End With
        
        Case 88 'buy guild symbol
            If Len(St) = 3 Then
                A = Asc(Mid$(St, 1, 1))
                B = Asc(Mid$(St, 2, 1))
                C = Asc(Mid$(St, 3, 1))
                E = 0
                With player(Index)
                    If (.Guild > 0) And .GuildRank >= 2 Then
                        With Guild(.Guild)
                            If (.Bank) >= IIf(A > 0, 100000, 0) + IIf(B > 0, 100000, 0) + IIf(C > 0, 100000, 0) Then
                                For D = 1 To 255
                                    If D <> player(Index).Guild Then
                                        With Guild(D)
                                            If A > GetFlag(221) Or B > GetFlag(222) Or C > GetFlag(223) Then
                                                SendSocket Index, Chr2(16) + Chr2(65)
                                                Exit Sub
                                            Else
                                                If .Symbol1 = A And .Symbol2 = B And .Symbol3 = C Then E = 1
                                                If .Symbol1 = A And .Symbol2 = 0 And .Symbol3 = 0 And B = 0 And C = 0 Then E = 1
                                                If .Symbol1 = 0 And .Symbol2 = B And .Symbol3 = 0 And A = 0 And C = 0 Then E = 1
                                                If .Symbol1 = 0 And .Symbol2 = 0 And .Symbol3 = C And B = 0 And A = 0 Then E = 1
                                                If .Symbol1 = A And .Symbol3 = C Then E = 1
                                                If .Symbol1 = A And .Symbol2 = B Then E = 1
                                                If .Symbol1 = A And .Symbol2 = B Then E = 1
                                            End If
                                        End With
                                    End If
                                Next D
                                If E = 0 Then
                                    If A > 0 Then .Bank = .Bank - 100000
                                    If B > 0 Then .Bank = .Bank - 100000
                                    If C > 0 Then .Bank = .Bank - 100000
                                    .Symbol1 = A
                                    .Symbol2 = B
                                    .Symbol3 = C
                                    GuildRS.Bookmark = .Bookmark
                                    GuildRS.Edit
                                    GuildRS("Symbol") = A
                                    GuildRS("GuildKills") = B
                                    GuildRS("GuildDeaths") = C
                                    GuildRS.Update
                                    SendToGuild CByte(player(Index).Guild), Chr2(74) + QuadChar(.Bank)
                                    UpdateGuildInfo CByte(player(Index).Guild)
                                Else
                                    SendSocket Index, Chr2(16) + Chr2(64)
                                End If
                            Else
                                SendSocket Index, Chr2(16) + Chr2(63)
                            End If
                        End With
                    End If
                End With
            End If
            
        Case 89 'mine
            If Len(St) = 2 Then
                A = Asc(Mid$(St, 1, 1))
                B = Asc(Mid$(St, 2, 1))
                With map(player(Index).map).Tile(A, B)
                    If .Att = 24 Then
                        If Abs(player(Index).x - A) > 1 Or Abs(player(Index).y - B) > 1 Then
                            Hacker Index, "MINE2"
                        Else
                  
                            
                            If .AttData(3) <> 0 Or .AttData(1) = 0 Then
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
                                
                                Parameter(0) = Index
                                Parameter(1) = .AttData(0) 'mine type
                                Parameter(2) = .AttData(1) 'max amount
                                Parameter(3) = .AttData(2) 'disposeable
                                Parameter(4) = .AttData(3) 'current amount
                                C = RunScript("MINE" & .AttData(0))
                                
                                
                                With player(Index).Equipped(1)
                                    If .Value > 0 Then .Value = .Value - 1 '(Damage / 2)
                                    If .Value <= 0 Then
                                        'Object Is Destroyed
                                        SendSocket Index, Chr2(57) + Chr2(EQ_WEAPON)
                                        .Object = 0
                                        CalculateStats Index
                                    Else
                                        SendSocket Index, Chr2(119) + Chr2(20 + EQ_WEAPON) + QuadChar(.Value)
                                    End If
                                End With
                                
                                If C = 0 Then 'successful mine
                                    If .AttData(3) > 0 Then .AttData(3) = .AttData(3) - 1
                                End If
                                If .AttData(3) = 0 Then
                                    SendToMap player(Index).map, Chr2(140) + Chr2(A) + Chr2(B)
                                    
                                    If ExamineBit(.AttData(2), 3) Then
                                        If A > 0 Then
                                            map(player(Index).map).Tile(A - 1, B).AttData(3) = 0
                                            SendToMap player(Index).map, Chr2(140) + Chr2(A - 1) + Chr2(B)
                                        End If
                                    End If
                                    If ExamineBit(.AttData(2), 4) Then
                                        If A < 11 Then
                                            map(player(Index).map).Tile(A + 1, B).AttData(3) = 0
                                            SendToMap player(Index).map, Chr2(140) + Chr2(A + 1) + Chr2(B)
                                        End If
                                    End If
                                    
                                    
                                End If
                            End If
                        End If
                    Else
                        Hacker Index, "MINE"
                    End If
                End With
            End If
        
        Case 90 'fish
            If Len(St) = 2 Then
                C = 0
                D = Asc(Mid$(St, 1, 1))
                E = Asc(Mid$(St, 2, 1))
                With map(player(Index).map)
                    For A = 1 To 10
                        If .Fish(A).x = D And .Fish(A).y = E And .Fish(A).TimeStamp > GetTickCount Then
                            Parameter(0) = Index
                            Parameter(1) = .Zone
                            Parameter(2) = D
                            Parameter(3) = E
                            B = RunScript("FISH")
                            If B = 0 Then 'successful fish
                                .Fish(A).TimeStamp = 0
                                .Fish(A).x = 0
                                .Fish(A).y = 0
                            End If
                            
                            
                            With player(Index).Equipped(1)
                                If .Value > 0 Then .Value = .Value - 1 '(Damage / 2)
                                If .Value <= 0 Then
                                    'Object Is Destroyed
                                    SendSocket Index, Chr2(57) + Chr2(EQ_WEAPON)
                                    .Object = 0
                                    CalculateStats Index
                                Else
                                    SendSocket Index, Chr2(119) + Chr2(20 + EQ_WEAPON) + QuadChar(.Value)
                                End If
                            End With
                            
                            Exit For
                        End If
                    Next A
                End With
            Else
                Hacker Index, "Fish"
            End If
        Case 91 'widget click
                A = SysAllocStringByteLen(.Widgets.WidgetScript, Len(.Widgets.WidgetScript))
                Parameter(0) = Index
                Parameter(1) = A
                Parameter(2) = Asc(Mid$(St, 1, 1)) * 256 + Asc(Mid$(St, 2, 1))
                Parameter(3) = Asc(Mid$(St, 3, 1)) * 256 + Asc(Mid$(St, 4, 1))
                RunScript ("CLICKWIDGET")
                SysFreeString A
        Case Else
            'SendToGods chr2(30) + player(index).name + " just got a b.3 hacking error! (i stopped booting for this because i dont understand it)."
            'Hacker index, "B.3"
    End Select
End With
Exit Sub
Error_Handler:
Open App.Path + "/LOG.TXT" For Append As #1
    ST1 = ""
    If Len(St) > 0 Then
        B = Len(St)
        For A = 1 To B
            ST1 = ST1 & Asc(Mid$(St, A, 1)) & "-"
        Next A
    End If
    Print #1, player(Index).Name & "/" & Err.Number & "/" & Err.Description & "/" & header & "/" & Len(St) & "/" & St & "/" & ST1 & "/" & player(Index).Mode & "  modReadClient2" & " - "
Close #1
Unhook
End
End Sub
