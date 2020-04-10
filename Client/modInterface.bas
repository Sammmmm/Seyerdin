Attribute VB_Name = "modInterface"
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


'Ok so I made this module mainly because I change the inventory
'every now and then and its really a bitch, so I decided to make it
'just a TINY bit easier to change but in the end its still a bitch
'so here are some constants to make it a little easier

'Update: its still a fucking bitch


'Inventory
'Public INVSrcX As Long
'Public InvSrcY As Long
Public INVDestX As Long
Public INVDestY As Long
Public INVWIDTH As Long
Public INVHEIGHT As Long

'HP
'Public HPSrcX As Long
'Public HPSrcY As Long
Public HPDestX As Long
Public HPDestY As Long
Public HPWIDTH As Long
Public HPHEIGHT As Long

'Energy
'Public ENERGYSrcX As Long
'Public ENERGYSrcY As Long
Public ENERGYDestX As Long
Public ENERGYDestY As Long
Public ENERGYWIDTH As Long
Public ENERGYHEIGHT As Long

'Mana
'Public MANASrcX As Long
'Public MANASrcY As Long
Public MANADestX As Long
Public MANADestY As Long
Public MANAWIDTH As Long
Public MANAHEIGHT As Long

'EXP Bar
'Public EXPSrcX As Long
'Public EXPSrcY As Long
Public EXPDestX As Long
Public EXPDestY As Long
Public EXPWIDTH As Long
Public EXPHEIGHT As Long


Public CurrentTab As Long
Public Const tsInventory = 0
Public Const tsParty = 1
Public Const tsStats = 2
Public Const tsSkills = 3
Public Const tsStats2 = 4

Public MiniMapTab As Long
Public Const tsButtons = 0
Public Const tsMap = 1

'###########################################################################################
'Game Window Constants 'I will eventually try to move all windows in to the viewport
Public CurrentWindow As Long
Public Const WINDOW_INVALID = 0
Public Const WINDOW_TRADE = 1
Public Const WINDOW_PLAYERCLICK = 2
Public Const WINDOW_CHARACTERCLICK = 3
Public Const WINDOW_NPCCLICK = 4
Public Const WINDOW_REPAIR = 5
Public Const WINDOW_SHOP = 6
Public Const WINDOW_STORAGE = 7

    'Player Click Window Stuff
    Type ClickIconData
        IconType As Byte
        IconData As Byte 'Will hold Spell for now .. can't think of anything else
        Picture As Byte
        ToolTip As String
    End Type
    Type ClickWindowData
        Icons() As ClickIconData
        x As Long
        y As Long
        Width As Long
        Height As Long
        NumIcons As Long
        Portrait As Byte
        Target As TargetData
        Loaded As Boolean
    End Type
    Public ClickWindow As ClickWindowData
    
'Click Window Constants
    Public Const CW_PARTY = 1
    Public Const CW_TRADE = 2
    Public Const CW_INFO = 3
    Public Const CW_GUILD = 4
    Public Const CW_SKILL = 10

'Current window flags
Public CurrentWindowFlags As Long
Public Const WINDOW_FLAG_INVISIBLE = 1


'Party Stuff
Public PartyIndex(1 To 10) As Long

'############################################################################################
'^ Textures for windows ^
Public texTrade As TextureType
Public texRepair As TextureType
Public texShop As TextureType
Public texStorage As TextureType
'############################################################################################
'Trade Window Constants
Public Const TRADEWINDOWWIDTH = 370
Public Const TRADEWINDOWHEIGHT = 144

'Repair Window Constants
Public Const REPAIRWINDOWWIDTH = 240
Public Const REPAIRWINDOWHEIGHT = 162
'Repair Window Variables
Public RepairCost As Long
Public RepairObj As InvObjData

'Shop Window Variables
Public CurShopPage As Long
Public NumShopItems As Long

'Storage Window Constants
Public Const STORAGEWINDOWWIDTH = 208
Public Const STORAGEWINDOWHEIGHT = 250
'############################################################################################

'Widget Constants
Public Const DIALOG_TELL = 1
Public CurrentDialog As Long

    'Widget Types
        Public Const WIDGET_BUTTON = 1
    'Widget Style Flags
        Public Const STYLE_SMALL = 1
        Public Const STYLE_MEDIUM = 2
        Public Const STYLE_LARGE = 4
        Public Const STYLE_DYNAMIC = 8


Sub DrawHP()
    Dim Percent As Single, St As String

    If Character.MaxHP > 0 Then
        If Character.HP > Character.MaxHP Then
            Percent = 1
        Else
            Percent = Character.HP / Character.MaxHP
        End If
    Else
        Percent = 0
    End If
    St = CStr(Int(Percent * 100)) + "% " + CStr(Character.HP) + "/" + CStr(Character.MaxHP)
    With frmMain
        Dim R1 As RECT, r2 As RECT
        R1.Top = cHPSrcY
        R1.Left = cHPSrcX
        R1.Bottom = R1.Top + cHPHEIGHT
        R1.Right = R1.Left + cHPWIDTH
        r2.Top = HPDestY
        r2.Left = HPDestX
        r2.Right = HPDestX + HPWIDTH
        r2.Bottom = HPDestY + HPHEIGHT
        sfcInventory(0).Surface.BltToDC .hdc, R1, r2
        R1.Top = 372
        R1.Left = 2 + (cHPWIDTH * Percent)
        R1.Bottom = 372 + cHPHEIGHT
        R1.Right = 2 + cHPWIDTH
        r2.Top = HPDestY
        r2.Left = HPDestX + (HPWIDTH * Percent)
        r2.Right = HPDestX + HPWIDTH
        r2.Bottom = HPDestY + HPHEIGHT
        sfcInventory(0).Surface.BltToDC .hdc, R1, r2
        .FontName = "Arial"
        .FontSize = 8 * WindowScaleY
        .ForeColor = vbWhite
        TextOut .hdc, HPDestX + (HPWIDTH - .TextWidth(St)) / 2, HPDestY + (HPHEIGHT - .textHeight(St)) / 2, St, Len(St)
        .Refresh
    End With
End Sub

Sub DrawEnergy()
    Dim Percent As Single, St As String
    
    If Character.MaxEnergy > 0 Then
        If Character.Energy > Character.MaxEnergy Then
            Percent = 1
        Else
            Percent = Character.Energy / Character.MaxEnergy
        End If
    Else
        Percent = 0
    End If
    St = CStr(Int(Percent * 100)) + "% " + CStr(Character.Energy) + "/" + CStr(Character.MaxEnergy)
    With frmMain
        Dim R1 As RECT, r2 As RECT
        R1.Top = cENERGYSrcY
        R1.Left = cENERGYSrcX
        R1.Bottom = R1.Top + cENERGYHEIGHT
        R1.Right = R1.Left + cENERGYWIDTH
        r2.Top = ENERGYDestY
        r2.Left = ENERGYDestX
        r2.Right = ENERGYDestX + ENERGYWIDTH
        r2.Bottom = ENERGYDestY + ENERGYHEIGHT
        sfcInventory(0).Surface.BltToDC .hdc, R1, r2
        R1.Top = 372
        R1.Left = 2 + (cENERGYWIDTH * Percent)
        R1.Bottom = 372 + cENERGYHEIGHT
        R1.Right = 2 + cENERGYWIDTH
        r2.Top = ENERGYDestY
        r2.Left = ENERGYDestX + (ENERGYWIDTH * Percent)
        r2.Right = ENERGYDestX + ENERGYWIDTH
        r2.Bottom = ENERGYDestY + ENERGYHEIGHT
        sfcInventory(0).Surface.BltToDC .hdc, R1, r2
        .FontName = "Arial"
        .FontSize = 8 * WindowScaleY
        .ForeColor = vbWhite
        TextOut .hdc, ENERGYDestX + (ENERGYWIDTH - .TextWidth(St)) / 2, ENERGYDestY + (ENERGYHEIGHT - .textHeight(St)) / 2, St, Len(St)
        .Refresh
    End With
End Sub


Sub DrawMana()
    Dim Percent As Single, St As String
    
    If Character.MaxMana > 0 Then
        If Character.Mana > Character.MaxMana Then
            Percent = 1
        Else
            Percent = Character.Mana / Character.MaxMana
        End If
    Else
        Percent = 0
    End If
    St = CStr(Int(Percent * 100)) + "% " + CStr(Character.Mana) + "/" + CStr(Character.MaxMana)
    With frmMain
        Dim R1 As RECT, r2 As RECT
        R1.Top = cMANASrcY
        R1.Left = cMANASrcX
        R1.Bottom = R1.Top + cMANAHEIGHT
        R1.Right = R1.Left + cMANAWIDTH
        r2.Top = MANADestY
        r2.Left = MANADestX
        r2.Right = MANADestX + MANAWIDTH
        r2.Bottom = MANADestY + MANAHEIGHT
        sfcInventory(0).Surface.BltToDC .hdc, R1, r2
        R1.Top = 372
        R1.Left = 2 + (cMANAWIDTH * Percent)
        R1.Bottom = 372 + cMANAHEIGHT
        R1.Right = 2 + cMANAWIDTH
        r2.Top = MANADestY
        r2.Left = MANADestX + (MANAWIDTH * Percent)
        r2.Right = MANADestX + MANAWIDTH
        r2.Bottom = MANADestY + MANAHEIGHT
        sfcInventory(0).Surface.BltToDC .hdc, R1, r2
        .FontName = "Arial"
        .FontSize = 8 * WindowScaleY

        .ForeColor = vbWhite
        TextOut .hdc, MANADestX + (MANAWIDTH - .TextWidth(St)) / 2, MANADestY + (MANAHEIGHT - .textHeight(St)) / 2, St, Len(St)
        .Refresh
    End With

End Sub

Sub DrawExperience()
    Dim MaxExp As Single, EXP As Single
    
    EXP = Character.Experience
    'MaxExp = Int(1000 * Character.Level ^ 1.4)
    MaxExp = EXPLevel(Character.Level)

    Dim Percent As Single, St As String
    If MaxExp = 0 Then
        Percent = 0
    Else
        Percent = EXP / MaxExp
    End If
    If Percent > 1 Then Percent = 1
    If Character.Level < 75 Then
        St = CStr(Int(Percent * 100)) + "% to Level " & (Character.Level + 1)
    Else
        St = Character.Experience
    End If
    With frmMain
        Dim R1 As RECT, r2 As RECT
        R1.Top = cEXPSrcY
        R1.Left = cEXPSrcX
        R1.Bottom = R1.Top + cEXPHEIGHT
        R1.Right = R1.Left + cEXPWIDTH
        r2.Top = EXPDestY
        r2.Left = EXPDestX
        r2.Right = EXPDestX + EXPWIDTH
        r2.Bottom = EXPDestY + EXPHEIGHT
        sfcInventory(0).Surface.BltToDC .hdc, R1, r2
        R1.Top = 372
        R1.Left = 2 + (cEXPWIDTH * Percent)
        R1.Bottom = 372 + cEXPHEIGHT
        R1.Right = 2 + cEXPWIDTH
        r2.Top = EXPDestY
        r2.Left = EXPDestX + (EXPWIDTH * Percent)
        r2.Right = EXPDestX + EXPWIDTH
        r2.Bottom = EXPDestY + EXPHEIGHT
        sfcInventory(0).Surface.BltToDC .hdc, R1, r2
        .FontName = "Arial"
        .FontSize = 8 * WindowScaleY
        .ForeColor = vbWhite
        TextOut .hdc, EXPDestX + (EXPWIDTH - .TextWidth(St)) / 2, EXPDestY + (EXPHEIGHT - .textHeight(St)) / 2, St, Len(St)
        .Refresh
    End With
End Sub

Sub DrawInv()
Dim A As Long, b As Long, C As Long, D As Long, x As Long, y As Long, InvNum As Long, r2 As RECT
    Draw sfcInventory2, 0, 0, cINVWIDTH, cINVHEIGHT, sfcInventory(0), cINVSrcX, cInvSrcY
    
    
    For InvNum = 1 To 20
        Dim R1 As RECT
        x = 4 + 34 * ((InvNum - 1) Mod 5)
        y = 3 + 35 * Int((InvNum - 1) / 5)
        With Character.Inv(InvNum)
            If .Object > 0 Then
                A = Object(.Object).Picture - 1
                If ExamineBit(Object(.Object).Flags, 6) Then A = A + 255
                b = (A \ 64) + 1
                C = 0
                If .Prefix > 0 Or .Suffix > 0 Or .Affix > 0 Then
                    D = ((.PrefixValue \ 64))
                    If ((.SuffixValue \ 64)) > D Then D = ((.SuffixValue \ 64))
                    If ((.AffixValue \ 64)) > D Then D = ((.AffixValue \ 64))
                    
                    Select Case D
                        Case 3
                            C = 5
                        Case 2
                            C = 4
                        Case 1
                            C = 3
                        Case Else
                            C = 2
                    End Select
                End If
                If (Object(.Object).Class > 0 And ((2 ^ (Character.Class - 1)) And Object(.Object).Class) = 0) Or Character.Level < Object(.Object).MinLevel Then
                    C = 1
                End If
                If A >= 0 And b > 0 And b <= NumObjects Then
                    If CurInvObj = InvNum Then
                        Draw sfcInventory2, x, y, 32, 32, sfcInventory(0), 420, 315
                        Select Case C
                            Case 1
                                Draw sfcInventory2, x + 2, y + 2, 28, 28, sfcInventory(0), 488, 317
                            Case 2
                                Draw sfcInventory2, x + 2, y + 2, 28, 28, sfcInventory(0), 519, 317
                            Case 3
                                Draw sfcInventory2, x + 2, y + 2, 28, 28, sfcInventory(0), 488, 348
                            Case 4
                                Draw sfcInventory2, x + 2, y + 2, 28, 28, sfcInventory(0), 519, 348
                            Case 5
                                Draw sfcInventory2, x + 2, y + 2, 28, 28, sfcInventory(0), 488, 380
                        End Select
                        Draw sfcInventory2, x, y, 32, 32, sfcObjects(b), ((A Mod 64) Mod 8) * 32, ((A Mod 64) \ 8) * 32, True, 0
                    Else
                        'If C = 1 Then Draw sfcInventory2, X, Y, 31, 31, sfcInventory(0), 486, 315
                        Select Case C
                            Case 1
                                Draw sfcInventory2, x, y, 31, 31, sfcInventory(0), 486, 315
                            Case 2
                                Draw sfcInventory2, x, y, 31, 31, sfcInventory(0), 517, 315
                            Case 3
                                Draw sfcInventory2, x, y, 31, 31, sfcInventory(0), 486, 346
                            Case 4
                                Draw sfcInventory2, x, y, 31, 31, sfcInventory(0), 517, 346
                            Case 5
                                Draw sfcInventory2, x, y, 31, 31, sfcInventory(0), 486, 378
                        End Select
                        Draw sfcInventory2, x, y, 32, 32, sfcObjects(b), ((A Mod 64) Mod 8) * 32, ((A Mod 64) \ 8) * 32, True, 0
                    End If
                Else
                    R1.Top = y - 1: R1.Bottom = R1.Top + 34: R1.Left = x - 1: R1.Right = R1.Left + 34
                    If CurInvObj = InvNum Then
                        Draw sfcInventory2, x, y, 32, 32, sfcInventory(0), 420, 315
                    End If
                End If
            Else
                R1.Top = y - 1: R1.Bottom = R1.Top + 34: R1.Left = x - 1: R1.Right = R1.Left + 34
                If CurInvObj = InvNum Then
                    Draw sfcInventory2, x, y, 32, 32, sfcInventory(0), 420, 315
                End If
            End If
        End With
    Next InvNum
    For InvNum = 1 To 5
        x = 4 + 34 * (InvNum - 1)
        y = 143
        With Character.Equipped(InvNum)
            If .Object > 0 Then
                A = Object(.Object).Picture - 1
                If ExamineBit(Object(.Object).Flags, 6) Then A = A + 255
                b = (A \ 64) + 1
                C = 0
                If .Prefix > 0 Or .Suffix > 0 Or .Affix > 0 Then
                    D = ((.PrefixValue \ 64))
                    If ((.SuffixValue \ 64)) > D Then D = ((.SuffixValue \ 64))
                    If ((.AffixValue \ 64)) > D Then D = ((.AffixValue \ 64))
                
                    Select Case (D)
                        Case 3
                            C = 5
                        Case 2
                            C = 4
                        Case 1
                            C = 3
                        Case Else
                            C = 2
                    End Select
                End If
                If A >= 0 And b > 0 And b <= NumObjects Then
                    If CurInvObj - 20 = InvNum And CurInvObj > 20 Then
                        'Draw sfcInventory2, 3 + (34 * (CurInvObj - 21)), 142, 34, 34, sfcInventory(0), 316 + (34 * (CurInvObj - 21)), 349
                        Select Case C
                            Case 2
                                Draw sfcInventory2, x + 2, y + 2, 28, 28, sfcInventory(0), 519, 317
                            Case 3
                                Draw sfcInventory2, x + 2, y + 2, 28, 28, sfcInventory(0), 488, 348
                            Case 4
                                Draw sfcInventory2, x + 2, y + 2, 28, 28, sfcInventory(0), 519, 348
                        End Select
                        Draw sfcInventory2, x + 1, y + 1, 32, 32, sfcObjects(b), ((A Mod 64) Mod 8) * 32, ((A Mod 64) \ 8) * 32, True, 0
                    Else
                        'Draw sfcInventory2, X - 2, Y - 1, 34, 34, sfcInventory(0), 386, 315
                        Select Case C
                            Case 2
                                Draw sfcInventory2, x, y + 1, 30, 30, sfcInventory(0), 517, 315
                            Case 3
                                Draw sfcInventory2, x, y + 1, 30, 30, sfcInventory(0), 486, 346
                            Case 4
                                Draw sfcInventory2, x, y + 1, 30, 30, sfcInventory(0), 517, 346
                        End Select
                        Draw sfcInventory2, x + 1, y + 1, 32, 32, sfcObjects(b), ((A Mod 64) Mod 8) * 32, ((A Mod 64) \ 8) * 32, True, 0
                    End If
                Else
                    R1.Top = y - 1: R1.Bottom = R1.Top + 34: R1.Left = x - 1: R1.Right = R1.Left + 34
                    If CurInvObj > 20 Then
                        If CurInvObj - 20 = InvNum Then
                            Draw sfcInventory2, 3 + (34 * (CurInvObj - 21)), 142, 34, 34, sfcInventory(0), 316 + (34 * (CurInvObj - 21)), 349
                        End If
                    End If
                End If
            Else
                R1.Top = y - 1: R1.Bottom = R1.Top + 34: R1.Left = x - 1: R1.Right = R1.Left + 34
                If CurInvObj > 20 Then
                    If CurInvObj - 20 = InvNum Then
                        Draw sfcInventory2, 3 + (34 * (CurInvObj - 21)), 142, 34, 34, sfcInventory(0), 316 + (34 * (CurInvObj - 21)), 349
                    End If
                End If
            End If
        End With
    Next InvNum
    
    

    r2.Top = 0: r2.Left = 0: r2.Bottom = cINVHEIGHT: r2.Right = cINVWIDTH
    R1.Top = INVDestY: R1.Left = INVDestX: R1.Bottom = INVDestY + INVHEIGHT: R1.Right = INVDestX + INVWIDTH
    
    sfcInventory2.Surface.BltToDC frmMain.hdc, r2, R1
    
    DrawCurInvObj
    frmMain.Refresh
End Sub

Sub DrawCurInvObj(Optional objNum As Long = 0, Optional objVal As Long = 0)
    Dim ObjName As String, ObjDesc As String, A As Long, b As Byte, St As String, prefixName As String, suffixName As String, affixName As String
    Dim tmpObject As InvObjData
    Dim R1 As RECT, r2 As RECT
    
    If CurrentTab = tsInventory Then
        Dim iHDC As Long
        If CurInvObj > 0 Then
            If CurInvObj <= 20 Then
                tmpObject = Character.Inv(CurInvObj)
            ElseIf CurInvObj >= 21 And CurInvObj <= 25 Then 'Equipped Items
                tmpObject = Character.Equipped(CurInvObj - 20)
            ElseIf CurInvObj >= 26 And CurInvObj <= 35 Then
                tmpObject.Object = SaleItem(CurInvObj - 26).GiveObject
                If Object(SaleItem(CurInvObj - 26).GiveObject).Type = 8 Then
                    tmpObject.Value = Object(SaleItem(CurInvObj - 26).GiveObject).ObjData(1) * 10
                ElseIf Object(SaleItem(CurInvObj - 26).GiveObject).Type >= 1 And Object(SaleItem(CurInvObj - 26).GiveObject).Type <= 4 Then
                    tmpObject.Value = Object(SaleItem(CurInvObj - 26).GiveObject).ObjData(0) * 10
                Else
                    tmpObject.Value = SaleItem(CurInvObj - 26).GiveValue
                End If
            ElseIf CurInvObj >= 36 And CurInvObj <= 55 Then 'Storage Items
                tmpObject = Character.Storage(Character.CurStoragePage, CurInvObj - 35)
            ElseIf CurInvObj >= 56 And CurInvObj <= 105 Then 'Map Object
                tmpObject.Object = map.Object(CurInvObj - 56).Object
                tmpObject.Prefix = map.Object(CurInvObj - 56).Prefix
                tmpObject.PrefixValue = map.Object(CurInvObj - 56).PrefixVal
                tmpObject.Suffix = map.Object(CurInvObj - 56).Suffix
                tmpObject.SuffixValue = map.Object(CurInvObj - 56).SuffixVal
                tmpObject.Affix = map.Object(CurInvObj - 56).Affix
                tmpObject.AffixValue = map.Object(CurInvObj - 56).AffixVal
                tmpObject.Value = map.Object(CurInvObj - 56).Value
                tmpObject.ObjectColor = map.Object(CurInvObj - 56).ObjectColor
            ElseIf CurInvObj >= 106 And CurInvObj <= 115 Then
                tmpObject = TradeData.YourObjects(CurInvObj - 105)
            ElseIf CurInvObj >= 116 And CurInvObj <= 125 Then
                tmpObject = TradeData.TheirObjects(CurInvObj - 115)
            End If
            If objNum > 0 Then
                tmpObject.Object = objNum
                tmpObject.Prefix = 0
                tmpObject.Suffix = 0
                tmpObject.Affix = 0
                tmpObject.ObjectColor = 0
                If objVal <> 0 Then
                    tmpObject.Value = objVal
                Else
                    tmpObject.Value = Object(objNum).ObjData(0) * 10
                End If
            End If
            
            If True Then
                With tmpObject
                    Draw sfcInventory2, 0, 0, 190, 368, sfcInventory(0), 0, 0
                    Draw sfcInventory2, 10, 10, 32, 32, sfcInventory(0), 420, 315
                    If .Object > 0 Then
                        A = Object(.Object).Picture - 1
                        If ExamineBit(Object(.Object).Flags, 6) Then A = A + 255
                        b = (A \ 64) + 1
                        If A >= 0 And b > 0 And b <= NumObjects Then
                            Draw sfcInventory2, 10, 10, 32, 32, sfcObjects(b), ((A Mod 64) Mod 8) * 32, ((A Mod 64) \ 8) * 32, True, 0
                        End If
                        R1.Top = 45: R1.Left = 5: R1.Right = 180: R1.Bottom = 390
                        ObjName = Object(.Object).Name
                        ObjDesc = " "
                        If Object(.Object).Type = 6 Then
                            'Money
                            ObjName = ObjName + " [" + CStr(.Value) + "]"
                            If Len(Object(.Object).Description) > 0 Then
                                ObjDesc = ObjDesc + "||" + "||" + Object(.Object).Description
                            End If
                        Else
                            If Object(.Object).Type > 0 Then
                                If .Prefix > 0 Then
                                    prefixName = Prefix(.Prefix).Name
                                End If
                                If .Suffix > 0 Then
                                    suffixName = Prefix(.Suffix).Name
                                End If
                                If .Affix > 0 Then
                                    affixName = Prefix(.Affix).Name
                                End If
                                If .Prefix > 0 Then
                                    If Prefix(.Prefix).ModType <> 26 And Prefix(.Prefix).ModType <> 27 Then
                                        If Prefix(.Prefix).ModType = 17 Then
                                            ObjDesc = "||" + ModString(Prefix(.Prefix).ModType) + " X " + CStr((.PrefixValue And 63))
                                        ElseIf Prefix(.Prefix).ModType = 10 Then
                                            ObjDesc = "||" + ModString(Prefix(.Prefix).ModType)
                                        ElseIf Prefix(.Prefix).ModType = 28 Or Prefix(.Prefix).ModType = 29 Then
                                            ObjDesc = "||" + Prefix(.Prefix).Name + ": " + CStr((.PrefixValue And 63))
                                        Else
                                            ObjDesc = "||" + ModString(Prefix(.Prefix).ModType) + " + " + CStr((.PrefixValue And 63))
                                        End If
                                    End If
                                End If
                                If .Suffix > 0 Then
                                    If Prefix(.Suffix).ModType <> 26 And Prefix(.Suffix).ModType <> 27 Then
                                        If Prefix(.Suffix).ModType = 17 Then
                                            ObjDesc = ObjDesc + "||" + ModString(Prefix(.Suffix).ModType) + " X " + CStr((.SuffixValue And 63))
                                        ElseIf Prefix(.Suffix).ModType = 10 Then
                                            ObjDesc = ObjDesc + "||" + ModString(Prefix(.Suffix).ModType)
                                        ElseIf Prefix(.Prefix).ModType = 28 Or Prefix(.Prefix).ModType = 29 Then
                                            ObjDesc = "||" + Prefix(.Suffix).Name + ": " + CStr((.SuffixValue And 63))
                                        Else
                                            ObjDesc = ObjDesc + "||" + ModString(Prefix(.Suffix).ModType) + " + " + CStr((.SuffixValue And 63))
                                        End If
                                    End If
                                End If
                                If .Affix > 0 Then
                                    If Prefix(.Affix).ModType <> 26 And Prefix(.Affix).ModType <> 27 Then
                                        If Prefix(.Affix).ModType = 17 Then
                                            ObjDesc = ObjDesc + "||" + ModString(Prefix(.Affix).ModType) + " X " + CStr((.AffixValue And 63))
                                        ElseIf Prefix(.Affix).ModType = 10 Then
                                            ObjDesc = ObjDesc + "||" + ModString(Prefix(.Affix).ModType)
                                        ElseIf Prefix(.Prefix).ModType = 28 Or Prefix(.Prefix).ModType = 29 Then
                                            ObjDesc = "||" + Prefix(.Affix).Name + ": " + CStr((.AffixValue And 63))
                                        Else
                                            ObjDesc = ObjDesc + "||" + ModString(Prefix(.Affix).ModType) + " + " + CStr((.AffixValue And 63))
                                        End If
                                    End If
                                End If
                                Select Case Object(.Object).Type
                                    Case 1, 10 'Weapon, Projectile
                                        If Object(.Object).ObjData(0) > 0 Then
                                            ObjDesc = ObjDesc + "||" + "Durability: " & Round(((.Value * 10) / (Object(.Object).ObjData(0))), 2) & "%"
                                        Else
                                            ObjDesc = ObjDesc + "||" + "Durability: " & "0%"
                                        End If
                                        If Object(.Object).Type = 1 Then
                                            ObjDesc = ObjDesc + "||" + "Damage: " & Object(.Object).ObjData(1) & "-" & (CLng(Object(.Object).ObjData(2)) * 256# + CLng(Object(.Object).ObjData(3)))
                                        Else
                                            ObjDesc = ObjDesc + "||" + "Plus: " & Object(.Object).ObjData(2)
                                        End If
                                        
                                        If Object(.Object).ObjData(9) <> 100 Then
                                            If Object(.Object).ObjData(9) < 100 Then
                                                ObjDesc = ObjDesc + "||" + "Magic Amplification: -" & (100 - Object(.Object).ObjData(9)) & "%"
                                            Else
                                                ObjDesc = ObjDesc + "||" + "Magic Amplification: +" & (Object(.Object).ObjData(9) - 100) & "%"
                                            End If
                                        End If
                                        
                                    Case 2 'Shield
                                        ObjDesc = ObjDesc + "||" + "Durability: " & Round((.Value * 10) / (Object(.Object).ObjData(0)), 2) & "%"
                                        
                                        If Object(.Object).ObjData(2) <> 100 And (Object(.Object).ObjData(1)) <> 0 Then
                                            ObjDesc = ObjDesc + "|&"
                                            If (Object(.Object).ObjData(1)) <> 255 Then
                                                If Object(.Object).ObjData(2) < 100 Then
                                                    ObjDesc = ObjDesc + "||" + "Physical Block Chance: " & Object(.Object).ObjData(1) * 100 \ 255 & "%"
                                                Else
                                                    ObjDesc = ObjDesc + "||" + "Physical Block Chance: " & Object(.Object).ObjData(1) & "%"
                                                End If
                                            End If
                                            If Object(.Object).ObjData(2) <> 0 Then
                                                If Object(.Object).ObjData(2) < 100 Then
                                                    ObjDesc = ObjDesc + "||" + "Physical Resistance: +" & (100 - Object(.Object).ObjData(2)) & "%"
                                                Else
                                                    ObjDesc = ObjDesc + "||" + "Physical Resistance: -" & (Object(.Object).ObjData(2) - 100) & "%"
                                                End If
                                            End If
                                        End If

                                        If Object(.Object).ObjData(4) <> 100 And (Object(.Object).ObjData(3)) <> 0 Then
                                            ObjDesc = ObjDesc + "|&"
                                            If (Object(.Object).ObjData(3)) <> 255 Then
                                                If Object(.Object).ObjData(4) < 100 Then
                                                    ObjDesc = ObjDesc + "||" + "Magical Block Chance: " & Object(.Object).ObjData(3) * 100 \ 255 & "%"
                                                Else
                                                    ObjDesc = ObjDesc + "||" + "Magical Block Chance: " & Object(.Object).ObjData(3) & "%"
                                                End If
                                            End If
                                            If Object(.Object).ObjData(4) <> 0 Then
                                                If Object(.Object).ObjData(4) < 100 Then
                                                    ObjDesc = ObjDesc + "||" + "Magic Resistance: +" & (100 - Object(.Object).ObjData(4)) & "%"
                                                Else
                                                    ObjDesc = ObjDesc + "||" + "Magic Resistance: -" & (Object(.Object).ObjData(4) - 100) & "%"
                                                End If
                                            End If
                                        End If


                                        
                                        If Object(.Object).ObjData(9) <> 100 Then
                                            If Object(.Object).ObjData(9) < 100 Then
                                                ObjDesc = ObjDesc + "||" + "Magic Amplification: -" & (100 - Object(.Object).ObjData(9)) & "%"
                                            Else
                                                ObjDesc = ObjDesc + "||" + "Magic Amplification: +" & (Object(.Object).ObjData(9) - 100) & "%"
                                            End If
                                        End If
                                        
                                    Case 3, 4 'Armor, Helmet
                                        ObjDesc = ObjDesc + "||" + "Durability: " & Round((.Value * 10) / (Object(.Object).ObjData(0)), 2) & "%"
                                        
                                        If (CLng(Object(.Object).ObjData(1)) * 256 + CLng(Object(.Object).ObjData(2))) > 0 Or (Object(.Object).ObjData(3)) > 0 Then
                                            ObjDesc = ObjDesc + "|&"
                                        End If
                                        If (CLng(Object(.Object).ObjData(1)) * 256 + CLng(Object(.Object).ObjData(2))) > 0 Then ObjDesc = ObjDesc + "||" + "Physical Defense: " & ((CLng(Object(.Object).ObjData(1)) * 256 + CLng(Object(.Object).ObjData(2))))
                                        If (Object(.Object).ObjData(3)) > 0 And (Object(.Object).ObjData(3)) <= 100 Then ObjDesc = ObjDesc + "||" + "Physical Resistance: +" & Object(.Object).ObjData(3) & "%"
                                        
                                        If (Object(.Object).ObjData(4)) > 0 Or (Object(.Object).ObjData(5)) > 0 Then
                                            ObjDesc = ObjDesc + "|&"
                                        End If
                                        If (Object(.Object).ObjData(4)) > 0 Then ObjDesc = ObjDesc + "||" + "Magical Defense: " & Object(.Object).ObjData(4)
                                        If (Object(.Object).ObjData(5)) > 0 And (Object(.Object).ObjData(5)) <= 100 Then ObjDesc = ObjDesc + "||" + "Magical Resistance: +" & Object(.Object).ObjData(5) & "%"
                                        
                                        
                                        If Object(.Object).ObjData(9) <> 100 Then
                                            If Object(.Object).ObjData(9) < 100 Then
                                                ObjDesc = ObjDesc + "||" + "Magic Amplification: -" & (100 - Object(.Object).ObjData(9)) & "%"
                                            Else
                                                ObjDesc = ObjDesc + "||" + "Magic Amplification: +" & (Object(.Object).ObjData(9) - 100) & "%"
                                            End If
                                        End If
                                        
                                    Case 8 'Ring
                                        Select Case Object(.Object).ObjData(0)
                                            Case 0: ObjDesc = ObjDesc + "||" + "Type: " + "+" + CStr(Object(.Object).ObjData(2)) + " Attack"
                                            Case 1: ObjDesc = ObjDesc + "||" + "Type: " + "+" + CStr(Object(.Object).ObjData(2)) + " Defense"
                                        End Select
                                        ObjDesc = ObjDesc + "||" + "Durability: " & Round((.Value * 10) / (Object(.Object).ObjData(1)), 2) & "%"
                                        
                                        If Object(.Object).ObjData(9) <> 100 Then
                                            If Object(.Object).ObjData(9) < 100 Then
                                                ObjDesc = ObjDesc + "||" + "Magic Amplification: -" & (100 - Object(.Object).ObjData(9)) & "%"
                                            Else
                                                ObjDesc = ObjDesc + "||" + "Magic Amplification: +" & (Object(.Object).ObjData(9) - 100) & "%"
                                            End If
                                        End If
                                        
                                    Case 11 'Ammo
                                        ObjDesc = ObjDesc + " [" + CStr(.Value) + "]"
                                        ObjDesc = ObjDesc + "||" + "Damage: " & Object(.Object).ObjData(2) & "-" & Object(.Object).ObjData(3)
                                End Select
                            End If
                            If Object(.Object).MinLevel > 1 Then
                                ObjDesc = ObjDesc + "|&"
                                ObjDesc = ObjDesc + "||" + "Minimum Level: " & Object(.Object).MinLevel
                            End If
                            b = 0
                            St = ""
                            If Object(.Object).Class > 0 Then
                                For A = 1 To MAX_CLASS
                                    If (2 ^ (A - 1) And Object(.Object).Class) Then
                                        If Len(St) > 0 Then
                                            St = St + ", "
                                        Else
                                            St = St + "Usable by: "
                                        End If
                                        St = St + Class(A).Name
                                    End If
                                Next A
                            End If
                            
                            If Object(.Object).Flags > 0 Then
                                ObjDesc = ObjDesc + "|&"
                                If ExamineBit(Object(.Object).Flags, 0) Then
                                    ObjDesc = ObjDesc + "||" + "Undroppable"
                                End If
                                If ExamineBit(Object(.Object).Flags, 1) Then
                                    ObjDesc = ObjDesc + "||" + "Unbreakable"
                                End If
                                If ExamineBit(Object(.Object).Flags, 2) Then
                                    ObjDesc = ObjDesc + "||" + "Unrepairable"
                                End If
                                If ExamineBit(Object(.Object).Flags, 3) Then
                                    ObjDesc = ObjDesc + "||" + "Two Handed"
                                End If
                                If ExamineBit(Object(.Object).Flags, 4) Then
                                    ObjDesc = ObjDesc + "||" + "Undepositable"
                                End If
                                If ExamineBit(Object(.Object).Flags, 5) Then
                                    ObjDesc = ObjDesc + "||" + "Dual Wield"
                                End If
                            End If
                            
                            If Len(St) > 0 Then
                                ObjDesc = ObjDesc + "|&" + "||" + St
                                b = 1
                                St = ""
                            End If
                            
                            If Len(Object(.Object).Description) > 0 Then
                                ObjDesc = ObjDesc + vbCrLf + vbCrLf + Object(.Object).Description
                            End If
                            b = 0
                            St = ""
                        End If

                        frmMain.Font.Name = "System"
                        frmMain.Font.Bold = True
                        iHDC = sfcInventory2.Surface.GetDC
                        SetBkMode iHDC, vbTransparent
                        sfcInventory2.Surface.ReleaseDC iHDC
                        
                        frmMain.Font.Name = "Arial"
                        frmMain.Font.Bold = False
                        
                        R1.Top = 0: R1.Left = 0: R1.Bottom = 368: R1.Right = 190
                        r2.Top = 45: r2.Left = 8: r2.Bottom = 368 + r2.Top: r2.Right = 190 + r2.Left
                        r2.Top = r2.Top * WindowScaleY: r2.Bottom = r2.Bottom * WindowScaleY: r2.Left = r2.Left * WindowScaleX: r2.Right = r2.Right * WindowScaleX
                        sfcInventory2.Surface.BltToDC frmMain.hdc, R1, r2
                                                
                        frmMain.FontBold = True
                        frmMain.FontSize = 10 * WindowScaleY
                        R1.Top = 46 '- frmMain.textHeight(ObjName)
                        R1.Bottom = R1.Bottom + 50
                        R1.Top = R1.Top * WindowScaleY: R1.Bottom = R1.Bottom * WindowScaleY: R1.Left = R1.Left * WindowScaleX: R1.Right = R1.Right * WindowScaleX
                    
                        R1.Left = R1.Left + 51 * WindowScaleX
                        
                        A = 0
                        If .Prefix > 0 Then If Prefix(.Prefix).ModType <> 23 And Prefix(.Prefix).ModType <> 27 And Prefix(.Prefix).ModType <> 28 Then A = A + 1
                        If .Suffix > 0 Then If Prefix(.Suffix).ModType <> 23 And Prefix(.Suffix).ModType <> 27 And Prefix(.Suffix).ModType <> 28 Then A = A + 1
                        If .Affix > 0 Then If Prefix(.Affix).ModType <> 23 And Prefix(.Affix).ModType <> 27 And Prefix(.Affix).ModType <> 28 Then A = A + 1
                        If A = 1 Then R1.Top = R1.Top + frmMain.textHeight(ObjName) * 0.5
                        If A = 0 Then R1.Top = R1.Top + frmMain.textHeight(ObjName) * 1
                                                
                                                
                        If .Affix > 0 And Not ExamineBit(Prefix(.Affix).Flags, 1) Then
                            If Prefix(.Affix).ModType <> 23 And Prefix(.Affix).ModType <> 27 And Prefix(.Affix).ModType <> 28 Then
                                Select Case ((.AffixValue \ 64))
                                    Case 3
                                        frmMain.ForeColor = &HD31CFB
                                    Case 2
                                        frmMain.ForeColor = &H6FADB8
                                    Case 1
                                        frmMain.ForeColor = &H7EF466
                                    Case Else
                                        frmMain.ForeColor = &HD8DB95
                                End Select
                                DrawText frmMain.hdc, affixName, Len(affixName), R1, DT_WORDBREAK Or DT_CENTER
                                R1.Top = R1.Top + frmMain.textHeight(ObjName) * 1
                            End If
                        End If
                                                
                        If .Prefix > 0 Then
                            If Prefix(.Prefix).ModType <> 23 And Prefix(.Prefix).ModType <> 27 And Prefix(.Prefix).ModType <> 28 Then
                                Select Case ((.PrefixValue \ 64))
                                    Case 3
                                        frmMain.ForeColor = &HD31CFB
                                    Case 2
                                        frmMain.ForeColor = &H6FADB8
                                    Case 1
                                        frmMain.ForeColor = &H7EF466
                                    Case Else
                                        frmMain.ForeColor = &HD8DB95
                                End Select
                                DrawText frmMain.hdc, prefixName, Len(prefixName), R1, DT_WORDBREAK Or DT_CENTER
                                R1.Top = R1.Top + frmMain.textHeight(ObjName) * 1
                            End If
                        End If
                        
                        If (.ObjectColor > 0) Then
                            frmMain.ForeColor = RGB(Lights(.ObjectColor).Red, Lights(.ObjectColor).Green, Lights(.ObjectColor).Blue)
                        Else
                            frmMain.ForeColor = vbWhite
                        End If
                        DrawText frmMain.hdc, ObjName, Len(ObjName), R1, DT_WORDBREAK Or DT_CENTER
                        R1.Top = R1.Top + frmMain.textHeight(ObjName) * 1
                        
                        If .Suffix > 0 Then
                            If Prefix(.Suffix).ModType <> 23 And Prefix(.Suffix).ModType <> 27 And Prefix(.Suffix).ModType <> 28 Then
                                Select Case ((.SuffixValue \ 64))
                                    Case 3
                                        frmMain.ForeColor = &HD31CFB
                                    Case 2
                                        frmMain.ForeColor = &H6FADB8
                                    Case 1
                                        frmMain.ForeColor = &H7EF466
                                    Case Else
                                        frmMain.ForeColor = &HD8DB95
                                End Select
                                DrawText frmMain.hdc, suffixName, Len(suffixName), R1, DT_WORDBREAK Or DT_CENTER
                                R1.Top = R1.Top + frmMain.textHeight(ObjName) * 1
                            End If
                        End If
                        
                        If .Affix > 0 And ExamineBit(Prefix(.Affix).Flags, 1) Then
                            If Prefix(.Affix).ModType <> 23 And Prefix(.Affix).ModType <> 27 And Prefix(.Affix).ModType <> 28 Then
                                Select Case ((.AffixValue \ 64))
                                    Case 3
                                        frmMain.ForeColor = &HD31CFB
                                    Case 2
                                        frmMain.ForeColor = &H6FADB8
                                    Case 1
                                        frmMain.ForeColor = &H7EF466
                                    Case Else
                                        frmMain.ForeColor = &HD8DB95
                                End Select
                                DrawText frmMain.hdc, affixName, Len(affixName), R1, DT_WORDBREAK Or DT_CENTER
                                R1.Top = R1.Top + frmMain.textHeight(ObjName) * 1
                            End If
                        End If
                        
                        
                        R1.Left = R1.Left - 40 * WindowScaleX
                    Else
                        R1.Top = 0: R1.Left = 0: R1.Bottom = 368: R1.Right = 190
                        r2.Top = 45: r2.Left = 8: r2.Bottom = 368 + r2.Top: r2.Right = 190 + r2.Left
                        r2.Top = r2.Top * WindowScaleY: r2.Bottom = r2.Bottom * WindowScaleY: r2.Left = r2.Left * WindowScaleX: r2.Right = r2.Right * WindowScaleX
                        sfcInventory2.Surface.BltToDC frmMain.hdc, R1, r2
                    End If
                End With
          
            End If

            Dim St1 As String


            
            frmMain.FontBold = False
            frmMain.FontSize = 9 * WindowScaleY
            frmMain.ForeColor = vbWhite
            While (InStr(1, ObjDesc, "|&", vbTextCompare) Or InStr(1, ObjDesc, "||", vbTextCompare))
                If (InStr(1, ObjDesc, "|&", vbTextCompare) < InStr(1, ObjDesc, "||", vbTextCompare) And InStr(1, ObjDesc, "|&", vbTextCompare) <> False) Or InStr(1, ObjDesc, "||", vbTextCompare) = False Then
                    St1 = Mid$(ObjDesc, 1, InStr(1, ObjDesc, "|&", vbTextCompare) - 1)
                    ObjDesc = Mid$(ObjDesc, InStr(1, ObjDesc, "|&", vbTextCompare) + 2)
                    DrawText frmMain.hdc, St1, Len(St1), R1, DT_WORDBREAK Or DT_LEFT
                    R1.Top = R1.Top + frmMain.textHeight(ObjName) * 0.25
                ElseIf InStr(1, ObjDesc, "||", vbTextCompare) <> 0 Then
                    St1 = Mid$(ObjDesc, 1, InStr(1, ObjDesc, "||", vbTextCompare) - 1)
                    ObjDesc = Mid$(ObjDesc, InStr(1, ObjDesc, "||", vbTextCompare) + 2)
                    DrawText frmMain.hdc, St1, Len(St1), R1, DT_WORDBREAK Or DT_LEFT
                    R1.Top = R1.Top + frmMain.textHeight(ObjName)
                End If
            Wend
            St1 = ObjDesc
            DrawText frmMain.hdc, ObjDesc, Len(ObjDesc), R1, DT_WORDBREAK Or DT_LEFT
            

            

            
        Else
        'TODO - remove?
            R1.Top = 0: R1.Left = 0: R1.Bottom = 368: R1.Right = 190
            r2.Top = 45: r2.Left = 8: r2.Bottom = 368 + r2.Top: r2.Right = 190 + r2.Left
            'sfcInventory2.Surface.BltToDC frmMain.hdc, R1, R2
        End If
        'frmMain.picInvbar.Refresh
        frmMain.Refresh
    End If
End Sub

Sub DrawStats()
    If CurrentTab <> tsStats2 Then
        DrawTrainBars
    Else
        calculateEvasion
        calculateAttackDamage
        calculateBlock
        calculatePoisonResist
        calculateHpRegen
        calculateManaRegen
        calculateCriticalChance
        calculateMagicResist
        calculateAttackSpeed
        calculateLeachHp
        calculateMagicFind
        calculateEnergyRegen
        calculateBodyArmor
        calculateHeadArmor
        DrawMoreStats
    End If
End Sub
Sub DrawMoreStats()
    With frmMain
        Dim R1 As RECT, r2 As RECT
        R1.Top = 0: R1.Left = 0: R1.Bottom = 368: R1.Right = 190
        r2.Top = 45: r2.Left = 8: r2.Bottom = 368 + r2.Top: r2.Right = 190 + r2.Left
        r2.Top = r2.Top * WindowScaleY: r2.Bottom = r2.Bottom * WindowScaleY: r2.Left = r2.Left * WindowScaleX: r2.Right = r2.Right * WindowScaleX
        
        sfcInventory2.Surface.SetForeColor vbWhite
        Draw sfcInventory2, 0, 0, 190, 368, sfcInventory(0), 0, 0

        'more
        Draw sfcInventory2, 130, 342, 54, 23, sfcInventory(0), 415, 547
        sfcInventory2.Surface.BltToDC frmMain.hdc, R1, r2
        frmMain.ForeColor = RGB(221, 219, 189)
        frmMain.FontName = "Arial"
        frmMain.FontSize = 10 * WindowScaleY
                    
        TextOut .hdc, 151 * WindowScaleX, 390 * WindowScaleY, "More", 4
        
        frmMain.ForeColor = RGB(221, 219, 189)
        frmMain.FontSize = 12 * WindowScaleY
        frmMain.FontUnderline = True
        TextOut .hdc, 55 * WindowScaleX, 50 * WindowScaleY, "Derived Stats" + Str(Character.statBaseAttack), Len("Derived Stats")
        
        frmMain.ForeColor = vbWhite
        frmMain.FontSize = 10 * WindowScaleY
        frmMain.FontUnderline = False
        TextOut .hdc, 17 * WindowScaleX, 70 * WindowScaleY, "Average Damage: " + Str(Character.statBaseAttack), Len("Average Damage: " + Str(Character.statBaseAttack))
        TextOut .hdc, 17 * WindowScaleX, 85 * WindowScaleY, "Attack Speed(sec): " + Str(CSng(Character.statAttackSpeed / 1000)), Len("Attack Speed(sec): " + Str(CSng(Character.statAttackSpeed / 1000)))
        If Character.statAttackSpeed = 0 Then Character.statAttackSpeed = 990
        TextOut .hdc, 17 * WindowScaleX, 100 * WindowScaleY, "Melee DPS: " + Str(CDec(Character.statBaseAttack / (Character.statAttackSpeed / 1000))), IIf(Len("Melee DPS: ") + 6 < Len("Melee DPS: " + Str(CDec(Character.statBaseAttack / (Character.statAttackSpeed / 1000)))), Len("Melee DPS: ") + 6, Len("Melee DPS: " + Str(CDec(Character.statBaseAttack / (Character.statAttackSpeed / 1000)))))
        
        If Character.statHpRegenLow <> Character.statHpRegenHigh Then
            TextOut .hdc, 17 * WindowScaleX, 120 * WindowScaleY, "HP Regeneration: " + Str(Character.statHpRegenLow) + " - " + Str(Character.statHpRegenHigh), Len("HP Regeneration: " + Str(Character.statHpRegenLow) + " - " + Str(Character.statHpRegenHigh))
        Else
            TextOut .hdc, 17 * WindowScaleX, 120 * WindowScaleY, "HP Regeneration: " + Str(Character.statHpRegenLow), Len("HP Regeneration: " + Str(Character.statHpRegenLow))
        End If
        If Character.statEnergyRegenLow <> Character.statEnergyRegenHigh Then
            TextOut .hdc, 17 * WindowScaleX, 135 * WindowScaleY, "Energy Regeneration: " + Str(Character.statEnergyRegenLow) + " - " + Str(Character.statEnergyRegenHigh), Len("Energy Regeneration: " + Str(Character.statEnergyRegenLow) + " - " + Str(Character.statEnergyRegenHigh))
        Else
            TextOut .hdc, 17 * WindowScaleX, 135 * WindowScaleY, "Energy Regeneration: " + Str(Character.statEnergyRegenLow), Len("Energy Regeneration: " + Str(Character.statEnergyRegenLow))
        End If
        If Character.statManaRegenLow <> Character.statManaRegenHigh Then
            TextOut .hdc, 17 * WindowScaleX, 150 * WindowScaleY, "Mana Regeneration: " + Str(Character.statManaRegenLow) + " - " + Str(Character.statManaRegenHigh), Len("Mana Regeneration: " + Str(Character.statManaRegenLow) + " - " + Str(Character.statManaRegenHigh))
        Else
            TextOut .hdc, 17 * WindowScaleX, 150 * WindowScaleY, "Mana Regeneration: " + Str(Character.statManaRegenLow), Len("Mana Regeneration: " + Str(Character.statManaRegenLow))
        End If
        
        TextOut .hdc, 17 * WindowScaleX, 170 * WindowScaleY, "Block Chance: " + Str(CInt(Character.statBlock * 100 / 255)) + "%", Len("Block Chance: " + Str(CInt(Character.statBlock * 100 / 255)) + "%")
        TextOut .hdc, 17 * WindowScaleX, 185 * WindowScaleY, "Dodge Chance: " + Str(CInt(Character.statEvasion * 100 / 255)) + "%", Len("Dodge Chance: " + Str(CInt(Character.statEvasion * 100 / 255)) + "%")
        TextOut .hdc, 17 * WindowScaleX, 200 * WindowScaleY, "Critical Chance: " + Str((Character.statCritical)) + "%", Len("Critical Chance: " + Str((Character.statCritical)) + "%")
        TextOut .hdc, 17 * WindowScaleX, 215 * WindowScaleY, "Magic Resist: " + Str((Character.statMagicDefense)) + " + " + Str((Character.statMagicResist)) + "%", Len("Magic Resist: " + Str((Character.statMagicDefense)) + " + " + Str((Character.statMagicResist)) + "%")
        TextOut .hdc, 17 * WindowScaleX, 230 * WindowScaleY, "Physical Resist: " + Str((Character.statPhysicalDefense)) + " + " + Str((Character.statPhysicalResist)) + "%", Len("Physical Resist: " + Str((Character.statPhysicalDefense)) + " + " + Str((Character.statPhysicalResist)) + "%")
        
        TextOut .hdc, 17 * WindowScaleX, 245 * WindowScaleY, "Poison Resist: " + Str((Character.statPoisonResist)) + "%", Len("Poison Resist: " + Str((Character.statPoisonResist)) + "%")
        TextOut .hdc, 17 * WindowScaleX, 260 * WindowScaleY, "Leach HP: " + Str((Character.statLeachHP)) + "%", Len("Leach HP: " + Str((Character.statLeachHP)) + "%")
        TextOut .hdc, 17 * WindowScaleX, 275 * WindowScaleY, "Magic Find Bonus: " + Str(Character.statMagicFind) + "%", Len("Magic Find Bonus: " + Str(Character.statMagicFind) + "%")
        
        TextOut .hdc, 17 * WindowScaleX, 295 * WindowScaleY, "Body Armor: " + Str(Character.statBodyArmor), Len("Body Armor: " + Str(Character.statBodyArmor))
        TextOut .hdc, 17 * WindowScaleX, 310 * WindowScaleY, "Head Armor: " + Str(Character.statHeadArmor), Len("Head Armor: " + Str(Character.statHeadArmor))
        
        
        .Refresh
    End With
End Sub

Sub DrawTrainBars()
    Dim St As String
    If CurrentTab = tsStats Then
        With frmMain
            
            Draw sfcInventory2, 0, 0, 190, 368, sfcInventory(0), 0, 0
            Draw sfcInventory2, 20, 107, 145, 253, sfcInventory(0), 0, 389
            'more
            Draw sfcInventory2, 130, 342, 54, 23, sfcInventory(0), 415, 547
            
            
            sfcInventory2.Surface.SetForeColor vbWhite
            
            Draw sfcInventory2, 20, 121, (Character.strength) * 117 / 50, 16, sfcInventory(0), 386, 91
            Draw sfcInventory2, 20, 153, (Character.Agility) * 117 / 50, 17, sfcInventory(0), 386, 91
            Draw sfcInventory2, 20, 184, (Character.Endurance) * 117 / 50, 16, sfcInventory(0), 386, 91
            Draw sfcInventory2, 20, 216, (Character.Wisdom) * 117 / 50, 16, sfcInventory(0), 386, 91
            Draw sfcInventory2, 20, 249, (Character.Constitution) * 117 / 50, 16, sfcInventory(0), 386, 91
            Draw sfcInventory2, 20, 281, (Character.Intelligence) * 117 / 50, 16, sfcInventory(0), 386, 91

            If Tstr + TAgi + TEnd + TWis + TCon + TInt > 0 Then
                Draw sfcInventory2, 120, 90, 52, 21, sfcInventory(0), 242, 368
            End If
            
            Draw sfcInventory2, 140, 15, 32, 32, sfcInventory(0), 420, 315
            Draw sfcInventory2, 140, 15, 32, 32, sfcSprites, ((Character.Sprite - 1) Mod 16) * 32, ((Character.Sprite - 1) \ 16) * 32, True
            
            Dim R1 As RECT, r2 As RECT
            R1.Top = 0: R1.Left = 0: R1.Bottom = 368: R1.Right = 190
            r2.Top = 45: r2.Left = 8: r2.Bottom = 368 + r2.Top: r2.Right = 190 + r2.Left
            r2.Top = r2.Top * WindowScaleY: r2.Bottom = r2.Bottom * WindowScaleY: r2.Left = r2.Left * WindowScaleX: r2.Right = r2.Right * WindowScaleX
            sfcInventory2.Surface.BltToDC frmMain.hdc, R1, r2
            
            frmMain.FontName = "Arial"
            frmMain.FontSize = 10 * WindowScaleY
            frmMain.ForeColor = vbBlue

            St = Character.strength
            If Character.StrengthMod > 0 Then St = St & " +" & Character.StrengthMod
            If Character.StrengthMod < 0 Then St = St & " -" & Abs(Character.StrengthMod)
            If Tstr > 0 Then St = St + " +" + CStr(Tstr)
            TextOut .hdc, 50 * WindowScaleX, 29 * WindowScaleY + 138 * WindowScaleY, St, Len(St)
            St = Character.Agility
            If Character.AgilityMod > 0 Then St = St & " +" & Character.AgilityMod
            If Character.AgilityMod < 0 Then St = St & " -" & Abs(Character.AgilityMod)
            If TAgi > 0 Then St = St + " +" + CStr(TAgi)
            TextOut .hdc, 50 * WindowScaleX, 29 * WindowScaleY + 170 * WindowScaleY, St, Len(St)
            St = Character.Endurance
            If Character.EnduranceMod > 0 Then St = St & " +" & Character.EnduranceMod
            If Character.EnduranceMod < 0 Then St = St & " -" & Abs(Character.EnduranceMod)
            If TEnd > 0 Then St = St + " +" + CStr(TEnd)
            TextOut .hdc, 50 * WindowScaleX, 29 * WindowScaleY + 201 * WindowScaleY, St, Len(St)
            St = Character.Wisdom
            If Character.WisdomMod > 0 Then St = St & " +" & Character.WisdomMod
            If Character.WisdomMod < 0 Then St = St & " -" & Abs(Character.WisdomMod)
            If TWis > 0 Then St = St + " +" + CStr(TWis)
            TextOut .hdc, 50 * WindowScaleX, 29 * WindowScaleY + 233 * WindowScaleY, St, Len(St)
            St = Character.Constitution
            If Character.ConstitutionMod > 0 Then St = St & " +" & Character.ConstitutionMod
            If Character.ConstitutionMod < 0 Then St = St & " -" & Abs(Character.ConstitutionMod)
            If TCon > 0 Then St = St + " +" + CStr(TCon)
            TextOut .hdc, 50 * WindowScaleX, 31 * WindowScaleY + 264 * WindowScaleY, St, Len(St)
            St = Character.Intelligence
            If Character.IntelligenceMod > 0 Then St = St & " +" & Character.IntelligenceMod
            If Character.IntelligenceMod < 0 Then St = St & " -" & Abs(Character.IntelligenceMod)
            If TInt > 0 Then St = St + " +" + CStr(TInt)
            TextOut .hdc, 50 * WindowScaleX, 31 * WindowScaleY + 296 * WindowScaleY, St, Len(St)
            
            frmMain.ForeColor = vbWhite
            TextOut .hdc, 43 * WindowScaleX, 11 * WindowScaleY + 138 * WindowScaleY, "Strength:", 8
            TextOut .hdc, 43 * WindowScaleX, 11 * WindowScaleY + 170 * WindowScaleY, "Agility:", 7
            TextOut .hdc, 43 * WindowScaleX, 12 * WindowScaleY + 201 * WindowScaleY, "Endurance:", 10
            TextOut .hdc, 43 * WindowScaleX, 12 * WindowScaleY + 233 * WindowScaleY, "Piety:", 6
            TextOut .hdc, 43 * WindowScaleX, 13 * WindowScaleY + 264 * WindowScaleY, "Constitution:", 13
            TextOut .hdc, 43 * WindowScaleX, 13 * WindowScaleY + 296 * WindowScaleY, "Intelligence:", 13
            
            frmMain.ForeColor = RGB(221, 219, 189)
            TextOut .hdc, 151 * WindowScaleX, 390 * WindowScaleY, "More", 4

            
            frmMain.ForeColor = vbWhite
            
            St = Character.Name
            TextOut .hdc, 8 * WindowScaleX + 80 * WindowScaleX - frmMain.TextWidth(St) / 2, 55 * WindowScaleY, St, Len(St)
            St = "Level " & Character.Level & " " & Class(Character.Class).Name
            TextOut .hdc, 8 * WindowScaleX + 80 * WindowScaleX - frmMain.TextWidth(St) / 2, 70 * WindowScaleY, St, Len(St)
            If Character.Guild Then
                St = """" & Guild(Character.Guild).Name & """"
                TextOut .hdc, 8 * WindowScaleX + 80 * WindowScaleX - frmMain.TextWidth(St) * WindowScaleX / 2, 85 * WindowScaleY, St, Len(St)
            End If
            St = "Free Stat Points: " + CStr(TempVar5)
            TextOut .hdc, 32 * WindowScaleX, 357 * WindowScaleY, St, Len(St)
            St = "Exp: " + CStr(Character.Experience)
            TextOut .hdc, 32 * WindowScaleX, 389 * WindowScaleY, St, Len(St)
            .Refresh
        End With
    End If
End Sub

Sub DrawMapTitle(Title As String)
Dim St As String
Dim Rsrc As RECT, Rdest As RECT

    St = "[" + Title + "]"

    If ExamineBit(map.Flags(0), 0) = True Then
        frmMain.ForeColor = QBColor(11)
    ElseIf ExamineBit(map.Flags(0), 6) = True Then
        frmMain.ForeColor = QBColor(12)
    Else
        frmMain.ForeColor = QBColor(15)
    End If

    Rsrc.Top = 368
    Rsrc.Left = 0
    Rsrc.Bottom = 388
    Rsrc.Right = 243
    Rdest.Top = 7 * WindowScaleY
    Rdest.Left = 279 * WindowScaleX
    Rdest.Bottom = 27 * WindowScaleY
    Rdest.Right = 521 * WindowScaleX
    
    sfcInventory(0).Surface.BltToDC frmMain.hdc, Rsrc, Rdest
    frmMain.FontSize = 10 * WindowScaleY
    TextOut frmMain.hdc, 279 * WindowScaleX + (242 * WindowScaleX - frmMain.TextWidth(St)) / 2, 7 * WindowScaleY + (20 * WindowScaleY - frmMain.textHeight(St)) / 2, St, Len(St)
End Sub

Sub DrawPartyNames()

    Dim A As Long, St1 As String, CurIndex As Long
    Dim R1 As RECT, r2 As RECT
    Dim CurY As Long
    
    frmMain.Font.Name = "Arial"
    frmMain.Font.Size = 10 * WindowScaleY
    sfcInventory2.Surface.SetFont frmMain.Font
    sfcInventory2.Surface.SetFillStyle 0
    sfcInventory2.Surface.SetForeColor vbWhite
    
    CurY = 10
    Draw sfcInventory2, 0, 0, 190, 368, sfcInventory(0), 0, 0
    If CurrentTab = tsParty Then
        If Character.Party > 0 Then
            For A = 1 To MAXUSERS
                If player(A).Party = Character.Party Then
                    CurIndex = CurIndex + 1
                    PartyIndex(CurIndex) = A
                    Draw sfcInventory2, 5, CurY, 180, 43, sfcInventory(0), 301, 386
                    Draw sfcInventory2, 12, CurY + 6, 32, 32, sfcSprites, ((player(A).Sprite - 1) Mod 16) * 32, ((player(A).Sprite - 1) \ 16) * 32, True
                    Draw sfcInventory2, 45, CurY + 18, 134 * player(A).HP \ 100, 9, sfcInventory(0), 222, 324
                    Draw sfcInventory2, 45, CurY + 28, 134 * player(A).Mana \ 100, 8, sfcInventory(0), 222, 333
                    CurY = CurY + 43
                End If
            Next A
        End If
        R1.Top = 0: R1.Left = 0: R1.Bottom = 368: R1.Right = 190
        r2.Top = 45: r2.Left = 8: r2.Bottom = 368 + r2.Top: r2.Right = 190 + r2.Left
        r2.Top = r2.Top * WindowScaleY: r2.Bottom = r2.Bottom * WindowScaleY: r2.Left = r2.Left * WindowScaleX: r2.Right = r2.Right * WindowScaleX
        sfcInventory2.Surface.BltToDC frmMain.hdc, R1, r2
        
        
        If (Character.Party = 0) Then
            St1 = "You are not in a party!"
            TextOut frmMain.hdc, 26 * WindowScaleX, 115 * WindowScaleY, St1, Len(St1)
        Else
            CurY = 54
            For A = 1 To MAXUSERS
                If player(A).Party = Character.Party Then
                     TextOut frmMain.hdc, 53 * WindowScaleX, (CurY + 4) * WindowScaleY, player(A).Name, Len(player(A).Name)
                    CurY = CurY + 43
                End If
            Next A
        End If
        
        
        frmMain.Refresh
    End If
End Sub

Public Sub UpdateSkills()
    Dim A As Byte
    
    For A = 0 To 254
        SkillListBox.Caption(A + 1) = vbNullString
        SkillListBox.Data(A + 1) = 0
    Next A
    
    For A = 1 To MAX_SKILLS
        If CanUseSkill(A, 0, 1) = 0 And Len(Skills(A).Name) > 0 Then
            frmMain.LstAddItem Skills(A).Name, A
        End If
    Next A
    
    If GetTab = tsSkills Then drawSkills
End Sub

Public Sub drawSkills()
    Dim R1 As RECT, r2 As RECT, St As String, A As Long, b As Long
    Draw sfcInventory2, 0, 0, 190, 368, sfcInventory(0), 0, 0
    'Draw sfcInventory2, 0, 328, 190, 38, sfcInventory(0), 146, 390
    
    Draw sfcInventory2, 0, 21, 190, 16, sfcInventory(0), 386, 123, True
    Draw sfcInventory2, 0, 213, 190, 16, sfcInventory(0), 386, 107, True
    
    'Draw sfcInventory2, 3, 250, 113, 13, sfcInventory(0), 380, 335
    
    R1.Top = 0: R1.Left = 0: R1.Bottom = 368: R1.Right = 190
    r2.Top = 45: r2.Left = 8: r2.Bottom = 368 + r2.Top: r2.Right = 190 + r2.Left
    r2.Top = r2.Top * WindowScaleY: r2.Bottom = r2.Bottom * WindowScaleY: r2.Left = r2.Left * WindowScaleX: r2.Right = r2.Right * WindowScaleX
    sfcInventory2.Surface.BltToDC frmMain.hdc, R1, r2
    
    frmMain.FontName = "Arial"
    frmMain.FontSize = 10
    frmMain.ForeColor = vbWhite
    
    Draw sfcInventory2, 130, 342, 54, 23, sfcInventory(0), 415, 547
        sfcInventory2.Surface.BltToDC frmMain.hdc, R1, r2
        frmMain.ForeColor = RGB(221, 219, 189)
        frmMain.FontName = "Arial"
        frmMain.FontSize = 10 * WindowScaleY
  
        TextOut frmMain.hdc, 152 * WindowScaleX, 390 * WindowScaleY, "Tree", 4
    
    
    St = "Free Skill Points: " & Character.skillPoints
    TextOut frmMain.hdc, 12 * WindowScaleX, 390 * WindowScaleY, St, Len(St)
    R1.Left = 8: R1.Top = 277: R1.Right = 198: R1.Bottom = 406
    If SkillListBox.Selected > 0 Then
        If SkillListBox.Data(SkillListBox.Selected) > 0 Then
            A = SkillListBox.Data(SkillListBox.Selected)
            St = Skills(A).Description + vbCrLf
            b = (Skills(A).CostPerLevel * Character.Level) + (Skills(A).CostPerSLevel * Character.SkillLevels(A)) + Skills(A).CostConstant
            If b > 0 Then
                If Character.SkillLevels(SKILL_WIZARDRY) > 0 And ExamineBit(Skills(A).Flags, SF_HOSTILE) Then
                    b = b * 0.66
                End If
                St = St & "Mana: " & Str(b)
            Else
                St = St & "Passive"
            End If
            R1.Top = R1.Top * WindowScaleY: R1.Bottom = R1.Bottom * WindowScaleY: R1.Left = R1.Left * WindowScaleX: R1.Right = R1.Right * WindowScaleX
            DrawText frmMain.hdc, St, Len(St), R1, DT_WORDBREAK
        End If
    End If
    
    frmMain.Refresh
End Sub

Public Sub SetTab(CurTab As Long)
    CurrentTab = CurTab
    Dim R1 As RECT, r2 As RECT
    R1.Top = 248: R1.Left = 190: R1.Bottom = 270: R1.Right = 379
    r2.Top = 23: r2.Left = 8: r2.Bottom = 45: r2.Right = 197
    r2.Top = r2.Top * WindowScaleY: r2.Bottom = r2.Bottom * WindowScaleY: r2.Left = r2.Left * WindowScaleX: r2.Right = r2.Right * WindowScaleX
    sfcInventory(0).Surface.BltToDC frmMain.hdc, R1, r2
    Select Case CurTab
        Case tsStats
            frmMain.lstSkills.Visible = False
            R1.Top = 270: R1.Left = 190: R1.Bottom = 292: R1.Right = 249
            r2.Top = 23: r2.Left = 8: r2.Bottom = 44: r2.Right = 68
            r2.Top = r2.Top * WindowScaleY: r2.Bottom = r2.Bottom * WindowScaleY: r2.Left = r2.Left * WindowScaleX: r2.Right = r2.Right * WindowScaleX
            sfcInventory(0).Surface.BltToDC frmMain.hdc, R1, r2
            DrawStats
        Case tsStats2
            frmMain.lstSkills.Visible = False
            R1.Top = 270: R1.Left = 190: R1.Bottom = 292: R1.Right = 249
            r2.Top = 23: r2.Left = 8: r2.Bottom = 44: r2.Right = 68
            r2.Top = r2.Top * WindowScaleY: r2.Bottom = r2.Bottom * WindowScaleY: r2.Left = r2.Left * WindowScaleX: r2.Right = r2.Right * WindowScaleX
            sfcInventory(0).Surface.BltToDC frmMain.hdc, R1, r2
            DrawStats
        Case tsSkills
            frmMain.lstSkills.Visible = True
            R1.Top = 270: R1.Left = 255: R1.Bottom = 292: R1.Right = 314
            r2.Top = 23: r2.Left = 73: r2.Bottom = 45: r2.Right = 132
            r2.Top = r2.Top * WindowScaleY: r2.Bottom = r2.Bottom * WindowScaleY: r2.Left = r2.Left * WindowScaleX: r2.Right = r2.Right * WindowScaleX
            sfcInventory(0).Surface.BltToDC frmMain.hdc, R1, r2
            drawSkills
        Case tsParty
            frmMain.lstSkills.Visible = False
            R1.Top = 270: R1.Left = 320: R1.Bottom = 292: R1.Right = 379
            r2.Top = 23: r2.Left = 138: r2.Bottom = 45: r2.Right = 196
            r2.Top = r2.Top * WindowScaleY: r2.Bottom = r2.Bottom * WindowScaleY: r2.Left = r2.Left * WindowScaleX: r2.Right = r2.Right * WindowScaleX
            sfcInventory(0).Surface.BltToDC frmMain.hdc, R1, r2
            DrawPartyNames
        Case tsInventory
            frmMain.lstSkills.Visible = False
            DrawInv
    End Select
    
    frmMain.Refresh
End Sub

Public Sub SetMiniMapTab(CurTab As Long)
Dim R1 As RECT, r2 As RECT
    
    MiniMapTab = CurTab
    If CurTab = tsButtons Then
        R1.Top = 0: R1.Left = 190: R1.Bottom = 124: R1.Right = 386
        r2.Top = 297: r2.Left = 599: r2.Bottom = 421: r2.Right = 795
        r2.Top = r2.Top * WindowScaleY: r2.Bottom = r2.Bottom * WindowScaleY: r2.Left = r2.Left * WindowScaleX: r2.Right = r2.Right * WindowScaleX
        
        frmMain.picMiniMap.Width = picMiniMapWidth * WindowScaleX
        frmMain.picMiniMap.Height = picMiniMapHeight * WindowScaleY
        frmMain.picMiniMap.Left = picMiniMapLeft * WindowScaleX
        frmMain.picMiniMap.Top = picMiniMapTop * WindowScaleY
        
        
        sfcInventory(0).Surface.BltToDC frmMain.hdc, R1, r2
        
        If (Options.VsyncEnabled) Then
            frmMain.picMiniMap.Width = 0
            frmMain.picMiniMap.Height = 0
            frmMain.picMiniMap.Visible = True
        Else
            frmMain.picMiniMap.Visible = False
        End If
        
    ElseIf CurTab = tsMap Then
        R1.Top = 124: R1.Left = 365: R1.Bottom = 248: R1.Right = 386
        r2.Top = 297: r2.Left = 774: r2.Bottom = 421: r2.Right = 795
        r2.Top = r2.Top * WindowScaleY: r2.Bottom = r2.Bottom * WindowScaleY: r2.Left = r2.Left * WindowScaleX: r2.Right = r2.Right * WindowScaleX
        
        frmMain.picMiniMap.Width = picMiniMapWidth * WindowScaleX
        frmMain.picMiniMap.Height = picMiniMapHeight * WindowScaleY
        frmMain.picMiniMap.Left = picMiniMapLeft * WindowScaleX
        frmMain.picMiniMap.Top = picMiniMapTop * WindowScaleY
        
        sfcInventory(0).Surface.BltToDC frmMain.hdc, R1, r2
        frmMain.picMiniMap.Visible = True
        

        CreateMiniMap
        ReDrawMiniMap
        DrawMiniMapSection
    End If
    frmMain.Refresh
End Sub

Public Function GetTab() As Long
    GetTab = CurrentTab
End Function

Public Sub CreatePlayerClickWindow(Index As Long)

Dim A As Long

    If player(Index).map = CMap Then
        With ClickWindow
            'Default Icons
            .NumIcons = 3
            .Width = 58
            .Portrait = 0
            ReDim .Icons(1 To 3)
            .Icons(1).IconType = CW_PARTY
            .Icons(1).Picture = 0
            .Icons(1).ToolTip = "Invite to party"
            
            .Icons(2).IconType = CW_TRADE
            .Icons(2).Picture = 1
            If TradeData.Tradestate(0) = TRADE_STATE_INVITED And Index = TradeData.player Then
                .Icons(2).ToolTip = "Accept Trade"
            Else
                .Icons(2).ToolTip = "Invite Trade"
            End If
            
            .Icons(3).IconType = CW_INFO
            .Icons(3).Picture = 2
            .Icons(3).ToolTip = "Info"
            
            If Character.Guild > 0 Then
                .NumIcons = .NumIcons + 1
                ReDim Preserve .Icons(1 To .NumIcons)
                .Icons(.NumIcons).IconType = CW_GUILD
                .Icons(.NumIcons).IconData = 1 'Invite to Guild
                .Icons(.NumIcons).ToolTip = "Invite to Guild"
            End If
            
            For A = 1 To MAX_SKILLS
                If CanUseSkill(A, 0, 1) = 0 Then
                    If Skills(A).TargetType And TT_PLAYER Then
                        If Skills(A).Icon > 0 Then
                            .NumIcons = .NumIcons + 1
                            ReDim Preserve .Icons(1 To .NumIcons)
                            .Icons(.NumIcons).IconData = A
                            .Icons(.NumIcons).IconType = CW_SKILL
                            .Icons(.NumIcons).ToolTip = Skills(A).Name
                            .Icons(.NumIcons).Picture = Skills(A).Icon + 7
                        End If
                    End If
                End If
            Next A
            
            .Height = (((.NumIcons \ 3) + 1) * 18) + 4
            .Target.TargetType = TT_PLAYER
            .Target.Target = Index
            .x = player(Index).XO + 40
            .y = player(Index).YO
            
            If .y < 0 Then .y = 0
            If .y + .Height > 384 Then .y = 384 - .Height
            If .x < 0 Then .x = 0
            If .x + .Width > 384 Then .x = 384 - .Width
            
            .Loaded = True
        End With
        CurrentWindow = WINDOW_PLAYERCLICK
    End If
End Sub

Public Sub CreateCharacterClickWindow()

Dim A As Long

    With ClickWindow
        'Default Icons
        .NumIcons = 1
        .Width = 58
        .Portrait = 0
        ReDim .Icons(1 To 3)

        If Character.Party > 0 Then
            .Icons(1).IconType = CW_PARTY
            .Icons(1).Picture = 0
            .Icons(1).ToolTip = "Leave Party"
        End If
        
        If Character.Guild > 0 Then
            .NumIcons = .NumIcons + 1
            ReDim Preserve .Icons(1 To .NumIcons)
            .Icons(.NumIcons).IconType = CW_GUILD
            .Icons(.NumIcons).IconData = 1 'Leave Guild
            .Icons(.NumIcons).ToolTip = "Leave Guild"
        End If
        
        For A = 1 To MAX_SKILLS
            If CanUseSkill(A, 0, 1) = 0 Then
                If (Skills(A).TargetType And TT_CHARACTER) Or (Skills(A).TargetType And TT_NO_TARGET) Then
                    If Skills(A).Icon > 0 Then
                        .NumIcons = .NumIcons + 1
                        ReDim Preserve .Icons(1 To .NumIcons)
                        .Icons(.NumIcons).IconData = A
                        .Icons(.NumIcons).IconType = CW_SKILL
                        .Icons(.NumIcons).ToolTip = Skills(A).Name
                        .Icons(.NumIcons).Picture = Skills(A).Icon + 7
                    End If
                End If
            End If
        Next A
        
        .Height = (((.NumIcons \ 3) + 1) * 18) + 4
        .Target.TargetType = TT_CHARACTER
        .Target.Target = Character.Index
        .x = Cxo + 40
        .y = CYO
        
        If .y < 0 Then .y = 0
        If .y + .Height > 384 Then .y = 384 - .Height
        If .x < 0 Then .x = 0
        If .x + .Width > 384 Then .x = 384 - .Width
        .Loaded = True

        CurrentWindow = WINDOW_CHARACTERCLICK
    End With
End Sub

Public Sub CreateNPCClickWindow(ByVal NPCNum As Long, ByVal x As Long, ByVal y As Long)

    With ClickWindow
        If NPC(NPCNum).Portrait Then
            .Portrait = NPC(NPCNum).Portrait
        Else
            .Portrait = 0
        End If
        .Width = 58
        'Default Icons
        .NumIcons = 1
        ReDim .Icons(1 To 1)
        .Icons(1).IconType = 1 'NPC Talk
        .Icons(1).Picture = 3
        .Icons(1).ToolTip = "Talk"
        
        If NPC(NPCNum).Flags And NPC_SHOP Then
            .NumIcons = .NumIcons + 1
            ReDim Preserve .Icons(1 To .NumIcons)
            .Icons(.NumIcons).IconType = 2 'NPC SHOP
            .Icons(.NumIcons).Picture = 4
            .Icons(.NumIcons).ToolTip = "Shop"
        End If
        
        If NPC(NPCNum).Flags And NPC_REPAIR Then
            .NumIcons = .NumIcons + 1
            ReDim Preserve .Icons(1 To .NumIcons)
            .Icons(.NumIcons).IconType = 3 'NPC SHOP
            .Icons(.NumIcons).Picture = 5
            .Icons(.NumIcons).ToolTip = "Repair selected item"
        End If
        
        If NPC(NPCNum).Flags And NPC_BANK Then
            .NumIcons = .NumIcons + 1
            ReDim Preserve .Icons(1 To .NumIcons)
            .Icons(.NumIcons).IconType = 4 'NPC BANK
            .Icons(.NumIcons).Picture = 6
            .Icons(.NumIcons).ToolTip = "Bank"
        End If
        
        
        .Height = (((.NumIcons \ 3) + 1) * 18) + 4
        .Target.TargetType = TT_NPC
        .Target.Target = NPCNum
        .Target.x = x
        .Target.y = y
        .x = (x * 32) + 16 - (.Width \ 2)
        .y = (y * 32) - (.Height \ 2) + 16
        If .Portrait Then .y = .y + 34
        
        If .y < 0 Then .y = 0
        If .y + .Height > 384 Then .y = 384 - .Height
        If .x < 0 Then .x = 0
        If .x + .Width > 384 Then .x = 384 - .Width
        
        .Loaded = True

        CurrentWindow = WINDOW_NPCCLICK
    End With
End Sub

Public Sub DrawClickWindow()
    Dim tx As Long, ty As Long
    Dim A As Long, CurIcon As Long, LastY As Long
    tx = CurX * 32 + CurSubX
    ty = CurY * 32 + CurSubY
    
    tx = tx / WindowScaleX
    ty = ty / WindowScaleY
    
    With ClickWindow
       ' Select Case .Target.TargetType
       '     Case TT_PLAYER
       '         'DrawBmpString3D Create3DString(Player(.Target.Target).Name), .X + (58 \ 2), .Y - 16, False, &HFFFFFFFF
       '     Case TT_CHARACTER
       '         'DrawBmpString3D Create3DString(Character.Name), .X + (58 \ 2), .Y - 16, False, &HFFFFFFFF
       '     Case TT_NPC
       '         If .Portrait Then
       '             'DrawBmpString3D Create3DString(NPC(.Target.Target).Name), .X + .Width \ 2, .Y - 16 - 66, False, &HFFFFFFFF
       '         Else
       '             'DrawBmpString3D Create3DString(NPC(.Target.Target).Name), .X + .Width \ 2, .Y - 16, False, &HFFFFFFFF
       '         End If
       ' End Select
        
        CurIcon = 0
        
        If .Portrait > 0 Then
            Draw3D .x - 5, .y - 66, 68, 68, 58, 0, TexControl1, &HBFFFFFFF
            Draw3D .x - 3, .y - 64, 64, 64, 0, 0, TexPortrait0 + .Portrait
        End If


        tx = tx - .x - 2
        ty = ty - .y - 2
        
        If tx >= 0 And tx <= .Width - 2 Then
            If ty >= 0 And ty <= .Height - 2 Then
                CurIcon = (((ty \ (18)) * 3) + (tx \ (18))) + 1
                If CurIcon > .NumIcons Then CurIcon = 0
            End If
        End If
        
        Draw3D .x, .y, 58, 20, 0, 50, TexControl1, &HBFFFFFFF
        LastY = .y + 20
        For A = 1 To .NumIcons
            If .Icons(A).IconType > 0 Then
                If A = CurIcon Then
                    Draw3D .x + 3 + (((A - 1) Mod 3) * 18), .y + 3 + (((A - 1) \ 3) * 18), 16, 16, 128 + ((.Icons(A).Picture Mod 8) * 16), (.Icons(A).Picture \ 8) * 32 + (CurFrame * 16), TexControl1
                Else
                    Draw3D .x + 3 + (((A - 1) Mod 3) * 18), .y + 3 + (((A - 1) \ 3) * 18), 16, 16, 128 + ((.Icons(A).Picture Mod 8) * 16), (.Icons(A).Picture \ 8) * 32, TexControl1
                End If
                If A > 1 And ((A - 1) Mod 3) = 0 Then
                    Draw3D .x, .y + 2 + (((A - 1) \ 3) * 18), 58, 18, 0, 52, TexControl1, &HBFFFFFFF
                    LastY = LastY + 18
                End If
            End If
        Next A
        Draw3D .x, LastY, 58, 2, 0, 70, TexControl1, &HBFFFFFFF
        
        If CurIcon > 0 And CurIcon <= .NumIcons Then
            DrawBmpString3D Create3DString(.Icons(CurIcon).ToolTip), .x + (58 \ 2), LastY, &HFFFFFFFF
        End If
    End With
End Sub

Public Sub DestroyClickWindow()
    ReDim ClickWindow.Icons(0)
    ClickWindow.Loaded = False
    If CurrentWindow = WINDOW_PLAYERCLICK Then CurrentWindow = WINDOW_INVALID
    If CurrentWindow = WINDOW_CHARACTERCLICK Then CurrentWindow = WINDOW_INVALID
    If CurrentWindow = WINDOW_NPCCLICK Then CurrentWindow = WINDOW_INVALID
End Sub

Public Sub SetGUIWindow(WindowType As Long, Optional WindowFlags As Long = 0)
    
    'First check which window is currently open as later I may need to
    'clean up certain things (such as textures)
    Select Case CurrentWindow
        Case WINDOW_TRADE
            Set texTrade.Texture = Nothing
            Character.Trading = False
        Case WINDOW_REPAIR
            Set texRepair.Texture = Nothing
        Case WINDOW_PLAYERCLICK
            DestroyClickWindow
        Case WINDOW_SHOP
            Set texShop.Texture = Nothing
        Case WINDOW_STORAGE
            Set texStorage.Texture = Nothing
            StorageOpen = False
    End Select
    
    CurrentWindow = WindowType
    CurrentWindowFlags = WindowFlags
    
    Select Case CurrentWindow
        Case WINDOW_TRADE
            'Load trade window texture
            Set texTrade.Texture = D3DX.CreateTextureFromFileEx(D3DDevice, AppPath & "Data/Graphics/Interface/Trade.rsc", 512, 256, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texTrade.TexInfo, ByVal 0)
        Case WINDOW_REPAIR
            Set texRepair.Texture = D3DX.CreateTextureFromFileEx(D3DDevice, AppPath & "Data/Graphics/Interface/Repair.rsc", 256, 256, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texRepair.TexInfo, ByVal 0)
        Case WINDOW_SHOP
            Set texShop.Texture = D3DX.CreateTextureFromFileEx(D3DDevice, AppPath & "Data/Graphics/Interface/Shop.rsc", 256, 256, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texShop.TexInfo, ByVal 0)
        Case WINDOW_STORAGE
            Set texStorage.Texture = D3DX.CreateTextureFromFileEx(D3DDevice, AppPath & "Data/Graphics/Interface/Bank.rsc", 256, 256, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, &HFF000000, texStorage.TexInfo, ByVal 0)
            StorageOpen = True
    End Select
End Sub

Public Sub DrawGUIWindow()
    
    Dim A As Long, b As Long, C As Long, D As Long, x1 As Long, y1 As Long, St As String
    Dim Af As Single
    
    If (CurrentWindowFlags And WINDOW_FLAG_INVISIBLE) = 0 Then
        Select Case CurrentWindow
            Case WINDOW_TRADE
                If Character.Trading Then
                    x1 = 192 - (TRADEWINDOWWIDTH \ 2)
                    y1 = 192 - (TRADEWINDOWHEIGHT \ 2)
                    D3DDevice.SetTexture 0, texTrade.Texture
                    Draw3D x1, y1, TRADEWINDOWWIDTH, TRADEWINDOWHEIGHT, 0, 0, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                    If TradeData.Tradestate(1) = TRADE_STATE_ACCEPTED Then
                        Draw3D x1 + 254, y1 + 103, 64, 28, 0, 144, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                    End If
                    If TradeData.Tradestate(0) = TRADE_STATE_ACCEPTED Then
                        Draw3D x1 + 52, y1 + 103, 64, 28, 64, 144, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                    End If
                    'Draw Objects
                    For A = 1 To 10
                        If TradeData.YourObjects(A).Object > 0 Then
                            C = 0
                            If TradeData.YourObjects(A).Prefix > 0 Or TradeData.YourObjects(A).Suffix > 0 Then
                                Select Case ((TradeData.YourObjects(A).PrefixValue \ 64) And 3)
                                    Case 2
                                        C = 4
                                    Case 1
                                        C = 3
                                    Case Else
                                        C = 2
                                End Select
                            End If
                            If (Object(TradeData.YourObjects(A).Object).Class > 0 And ((2 ^ (Character.Class - 1)) And Object(TradeData.YourObjects(A).Object).Class) = 0) Or Character.Level < Object(TradeData.YourObjects(A).Object).MinLevel Then
                                C = 1
                            End If
                                If TradeData.CurrentObject = A Then
                                    D3DDevice.SetTexture 0, texTrade.Texture
                                    Draw3D x1 + 12 + (((A - 1) Mod 5) * 34), y1 + 19 + (((A - 1) \ 5) * 35), 32, 32, 0, 172, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                                    Select Case C
                                        Case 1
                                            Draw3D x1 + 13 + ((A - 1) Mod 5) * 35, y1 + 20 + ((A - 1) \ 5) * 36, 28, 28, 373, 2, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                                        Case 2
                                            Draw3D x1 + 13 + ((A - 1) Mod 5) * 35, y1 + 20 + ((A - 1) \ 5) * 36, 28, 28, 373, 33, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                                        Case 3
                                            Draw3D x1 + 13 + ((A - 1) Mod 5) * 35, y1 + 20 + ((A - 1) \ 5) * 36, 28, 28, 373, 64, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                                        Case 4
                                            Draw3D x1 + 13 + ((A - 1) Mod 5) * 35, y1 + 20 + ((A - 1) \ 5) * 36, 28, 28, 373, 95, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                                    End Select
                                Else
                                    D3DDevice.SetTexture 0, texTrade.Texture
                                    Select Case C
                                        Case 1
                                            Draw3D x1 + 12 + ((A - 1) Mod 5) * 35, y1 + 19 + ((A - 1) \ 5) * 36, 30, 30, 372, 1, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                                        Case 2
                                            Draw3D x1 + 12 + ((A - 1) Mod 5) * 35, y1 + 19 + ((A - 1) \ 5) * 36, 30, 30, 372, 32, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                                        Case 3
                                            Draw3D x1 + 12 + ((A - 1) Mod 5) * 35, y1 + 19 + ((A - 1) \ 5) * 36, 30, 30, 372, 63, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                                        Case 4
                                            Draw3D x1 + 12 + ((A - 1) Mod 5) * 35, y1 + 19 + ((A - 1) \ 5) * 36, 30, 30, 372, 94, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                                    End Select
                                End If
                                LastTexture = 0
                            DrawObject x1 + 12 + (((A - 1) Mod 5) * 35), y1 + 19 + (((A - 1) \ 5) * 36), IIf(ExamineBit(Object(TradeData.YourObjects(A).Object).Flags, 6), Object(TradeData.YourObjects(A).Object).Picture + 255, Object(TradeData.YourObjects(A).Object).Picture)
                        End If
                        If TradeData.TheirObjects(A).Object > 0 Then
                            C = 0
                            If TradeData.TheirObjects(A).Prefix > 0 Or TradeData.TheirObjects(A).Suffix > 0 Then
                                Select Case ((TradeData.TheirObjects(A).PrefixValue \ 64) And 3)
                                    Case 2
                                        C = 4
                                    Case 1
                                        C = 3
                                    Case Else
                                        C = 2
                                End Select
                            End If
                            If (Object(TradeData.TheirObjects(A).Object).Class > 0 And ((2 ^ (Character.Class - 1)) And Object(TradeData.TheirObjects(A).Object).Class) = 0) Or Character.Level < Object(TradeData.TheirObjects(A).Object).MinLevel Then
                                C = 1
                            End If
                            If TradeData.CurrentObject = A - 10 Then
                                D3DDevice.SetTexture 0, texTrade.Texture
                                Draw3D x1 + 189 + (((A - 1) Mod 5) * 34), y1 + 19 + (((A - 1) \ 5) * 35), 32, 32, 0, 172, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                                Select Case C
                                    Case 1
                                        Draw3D x1 + 190 + ((A - 1) Mod 5) * 35, y1 + 20 + ((A - 1) \ 5) * 36, 28, 28, 373, 2, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                                    Case 2
                                        Draw3D x1 + 190 + ((A - 1) Mod 5) * 35, y1 + 20 + ((A - 1) \ 5) * 36, 28, 28, 373, 33, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                                    Case 3
                                        Draw3D x1 + 190 + ((A - 1) Mod 5) * 35, y1 + 20 + ((A - 1) \ 5) * 36, 28, 28, 373, 64, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                                    Case 4
                                        Draw3D x1 + 190 + ((A - 1) Mod 5) * 35, y1 + 20 + ((A - 1) \ 5) * 36, 28, 28, 373, 95, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                                End Select
                            Else
                                D3DDevice.SetTexture 0, texTrade.Texture
                                Select Case C
                                    Case 1
                                        Draw3D x1 + 189 + ((A - 1) Mod 5) * 35, y1 + 19 + ((A - 1) \ 5) * 36, 30, 30, 372, 1, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                                    Case 2
                                        Draw3D x1 + 189 + ((A - 1) Mod 5) * 35, y1 + 19 + ((A - 1) \ 5) * 36, 30, 30, 372, 32, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                                    Case 3
                                        Draw3D x1 + 189 + ((A - 1) Mod 5) * 35, y1 + 19 + ((A - 1) \ 5) * 36, 30, 30, 372, 63, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                                    Case 4
                                        Draw3D x1 + 189 + ((A - 1) Mod 5) * 35, y1 + 19 + ((A - 1) \ 5) * 36, 30, 30, 372, 94, 0, -1, texTrade.TexInfo.Width, texTrade.TexInfo.Height
                                End Select
                            End If
                            LastTexture = 0
                            DrawObject x1 + 189 + (((A - 1) Mod 5) * 35), y1 + 19 + (((A - 1) \ 5) * 36), IIf(ExamineBit(Object(TradeData.TheirObjects(A).Object).Flags, 6), Object(TradeData.TheirObjects(A).Object).Picture + 255, Object(TradeData.TheirObjects(A).Object).Picture)
                        End If
                    Next A
                    
                    DrawBmpString3D Create3DString("Your Objects"), (x1 + 88), (y1 + 84), &HFFFFFFFF, False
                    DrawBmpString3D Create3DString(player(TradeData.player).Name), (x1 + 280), (y1 + 84), &HFFFFFFFF, False
                Else
                    SetGUIWindow WINDOW_INVALID
                End If
            Case WINDOW_PLAYERCLICK, WINDOW_CHARACTERCLICK, WINDOW_NPCCLICK
                DrawClickWindow
            Case WINDOW_REPAIR
                x1 = 192 - REPAIRWINDOWWIDTH \ 2
                y1 = 192 - REPAIRWINDOWHEIGHT \ 2
                D3DDevice.SetTexture 0, texRepair.Texture
                Draw3D x1, y1, REPAIRWINDOWWIDTH, REPAIRWINDOWHEIGHT, 0, 0, 0, -1, texRepair.TexInfo.Width, texRepair.TexInfo.Height
                With RepairObj
                    If .Object > 0 Then
                        St = Object(.Object).Name
                        If .Prefix > 0 Then
                            St = Prefix(.Prefix).Name + " " + St
                        End If
                        If .Suffix > 0 Then
                            St = St + " " + Prefix(.Suffix).Name
                        End If
                        'DrawBmpString3D Create3DString(St), X1 + 65, Y1 + 37, False, -1, False
                        DrawMultilineString3D St, x1 + 64, y1 + 37, 165
                        If Object(.Object).Type = 8 Then  'Ring
                            Af = (.Value * 10) / (Object(.Object).ObjData(1))
                        Else
                            Af = (.Value * 10) / (Object(.Object).ObjData(0))
                        End If
                        DrawBmpString3D Create3DString(CStr(RepairCost)), x1 + 62, y1 + 100, &HFFFFFFFF, False, 1
                        DrawBmpString3D Create3DString(CStr(Round(Af, 2))), x1 + 188, y1 + 100, &HFFFFFFFF, False, 1
                        DrawObject x1 + 21, y1 + 38, IIf(ExamineBit(Object(.Object).Flags, 6), Object(.Object).Picture + 255, Object(.Object).Picture)
                    End If
                End With
            Case WINDOW_SHOP
                x1 = 192 - (352 \ 2)
                If CurShopPage = 1 Then
                    If NumShopItems > 5 Then
                        y1 = 192 - (68 + 50 * 5) \ 2
                        b = 5
                    Else
                        y1 = 192 - ((68 + 50 * NumShopItems) \ 2)
                        b = NumShopItems
                    End If
                Else
                    y1 = 192 - ((68 + 50 * (NumShopItems - 5)) \ 2)
                    b = NumShopItems - 5
                End If
                D3DDevice.SetTexture 0, texShop.Texture
                'Draw left/right of top
                Draw3D x1, y1, 256, 27, 0, 0, 0, -1, texShop.TexInfo.Width, texShop.TexInfo.Height
                Draw3D x1 + 256, y1, 96, 27, 0, 121, 0, -1, texShop.TexInfo.Width, texShop.TexInfo.Height
                For A = 0 To b - 1
                    C = y1 + 27 + (A * 50)
                    D3DDevice.SetTexture 0, texShop.Texture
                    'draw item bar left/right
                        Draw3D x1, C, 256, 50, 0, 28, 0, -1, texShop.TexInfo.Width, texShop.TexInfo.Height
                        Draw3D x1 + 256, C, 96, 50, 0, 149, 0, -1, texShop.TexInfo.Width, texShop.TexInfo.Height
                    If 26 + A + ((CurShopPage - 1) * 5) = CurInvObj Then
                        Draw3D x1 + 9, C + 1, 247, 41, 9, 29, 0, D3DColorARGB(255, 150, 150, 150), texShop.TexInfo.Width, texShop.TexInfo.Height
                        Draw3D x1 + 256, C + 1, 87, 41, 0, 150, 0, D3DColorARGB(255, 150, 150, 150), texShop.TexInfo.Width, texShop.TexInfo.Height
                    End If
                    With SaleItem(A + ((CurShopPage - 1) * 5))
                        If .GiveObject > 0 And .TakeObject > 0 Then
                            If (Object(.GiveObject).Class > 0 And ((2 ^ Character.Class - 1) And Object(.GiveObject).Class) = 0) Or Character.Level < Object(.GiveObject).MinLevel Then
                                D = D3DColorARGB(255, 255, 0, 0) 'red
                            Else
                                D = -1 'white
                            End If
                            DrawObject x1 + 19, C + 6, IIf(ExamineBit(Object(.GiveObject).Flags, 6), Object(.GiveObject).Picture + 255, Object(.GiveObject).Picture)
                            If Object(.GiveObject).Type = 6 Or Object(.GiveObject).Type = 11 Then
                                St = .GiveValue & " " & Object(.GiveObject).Name
                            Else
                                St = Object(.GiveObject).Name
                            End If
                            St = St & " in exchange for "
                            DrawBmpString3D Create3DString(St), 192, C + 2, D, True
                            DrawObject x1 + 176 + 124 + 2, C + 6, IIf(ExamineBit(Object(.TakeObject).Flags, 6), Object(.TakeObject).Picture + 255, Object(.TakeObject).Picture)
                            If Object(.TakeObject).Type = 6 Or Object(.TakeObject).Type = 11 Then
                                St = .TakeValue & " " & Object(.TakeObject).Name
                            Else
                                St = Object(.TakeObject).Name
                            End If
                            DrawBmpString3D Create3DString(St), 192, C + 20, D, True
                        End If
                    End With
                Next A
                D3DDevice.SetTexture 0, texShop.Texture
                'draw left/right of bottom
                Draw3D x1, y1 + 27 + (b * 50), 256, 41, 0, 79, 0, -1, texShop.TexInfo.Width, texShop.TexInfo.Height
                
                If NumShopItems < 6 Then
                    Draw3D 69, y1 + 29 + (b * 50), 33, 33, 8, 79, 0, -1, texShop.TexInfo.Width, texShop.TexInfo.Height
                End If
                
                Draw3D x1 + 256, y1 + 27 + (b * 50), 96, 41, 0, 200, 0, -1, texShop.TexInfo.Width, texShop.TexInfo.Height
            Case WINDOW_STORAGE
                x1 = 192 - (STORAGEWINDOWWIDTH \ 2)
                y1 = 192 - (STORAGEWINDOWHEIGHT \ 2)
                D3DDevice.SetTexture 0, texStorage.Texture
                Draw3D x1, y1, STORAGEWINDOWWIDTH, STORAGEWINDOWHEIGHT, 0, 0, -1, -1, 256, 256
                If Character.CurStoragePage > 1 Then
                    Draw3D x1 + 34, y1 + 216, 28, 24, 208, 226, 0, -1, texStorage.TexInfo.Width, texStorage.TexInfo.Height
                End If
                If Character.CurStoragePage < Character.NumStoragePages Then
                    Draw3D x1 + 151, y1 + 216, 28, 24, 208, 202, 0, -1, texStorage.TexInfo.Width, texStorage.TexInfo.Height
                End If
                For A = 1 To 20
                    If CurStorageObj = A Then
                        D3DDevice.SetTexture 0, texStorage.Texture
                        Draw3D x1 + 16 + ((A - 1) Mod 5) * 36, y1 + 61 + ((A - 1) \ 5) * 35.5, 32, 32, 208, 0, 0, -1, 256, 256
                    End If
                    With Character.Storage(Character.CurStoragePage, A)
                        If .Object > 0 Then
                            C = 0
                            If .Prefix > 0 Or .Suffix > 0 Then
                                Select Case ((.PrefixValue \ 64) And 3)
                                    Case 2
                                        C = 4
                                    Case 1
                                        C = 3
                                    Case Else
                                        C = 2
                                End Select
                            End If
                            If (Object(.Object).Class > 0 And ((2 ^ (Character.Class - 1)) And Object(.Object).Class) = 0) Or Character.Level < Object(.Object).MinLevel Then
                                C = 1
                            End If
                            If C > 0 Then
                                D3DDevice.SetTexture 0, texStorage.Texture
                                If CurStorageObj = A Then
                                    Select Case C
                                        Case 1
                                            Draw3D x1 + 19 + ((A - 1) Mod 5) * 36, y1 + 63 + ((A - 1) \ 5) * 35.5, 28, 28, 210, 34, 0, -1, texStorage.TexInfo.Width, texStorage.TexInfo.Height
                                        Case 2
                                            Draw3D x1 + 19 + ((A - 1) Mod 5) * 36, y1 + 63 + ((A - 1) \ 5) * 35.5, 28, 28, 210, 65, 0, -1, texStorage.TexInfo.Width, texStorage.TexInfo.Height
                                        Case 3
                                            Draw3D x1 + 19 + ((A - 1) Mod 5) * 36, y1 + 63 + ((A - 1) \ 5) * 35.5, 28, 28, 210, 97, 0, -1, texStorage.TexInfo.Width, texStorage.TexInfo.Height
                                        Case 4
                                            Draw3D x1 + 19 + ((A - 1) Mod 5) * 36, y1 + 63 + ((A - 1) \ 5) * 35.5, 28, 28, 210, 127, 0, -1, texStorage.TexInfo.Width, texStorage.TexInfo.Height
                                    End Select
                                Else
                                    Select Case C
                                        Case 1
                                            Draw3D x1 + 17 + ((A - 1) Mod 5) * 36, y1 + 61 + ((A - 1) \ 5) * 35.5, 31, 31, 208, 32, 0, -1, texStorage.TexInfo.Width, texStorage.TexInfo.Height
                                        Case 2
                                            Draw3D x1 + 17 + ((A - 1) Mod 5) * 36, y1 + 61 + ((A - 1) \ 5) * 35.5, 31, 31, 208, 63, 0, -1, texStorage.TexInfo.Width, texStorage.TexInfo.Height
                                        Case 3
                                            Draw3D x1 + 17 + ((A - 1) Mod 5) * 36, y1 + 61 + ((A - 1) \ 5) * 35.5, 31, 31, 208, 95, 0, -1, texStorage.TexInfo.Width, texStorage.TexInfo.Height
                                        Case 4
                                            Draw3D x1 + 17 + ((A - 1) Mod 5) * 36, y1 + 61 + ((A - 1) \ 5) * 35.5, 31, 31, 208, 125, 0, -1, texStorage.TexInfo.Width, texStorage.TexInfo.Height
                                    End Select
                                End If
                                LastTexture = 0
                            End If
                            DrawObject x1 + 16 + ((A - 1) Mod 5) * 36, y1 + 61 + ((A - 1) \ 5) * 35.5, IIf(ExamineBit(Object(Character.Storage(Character.CurStoragePage, A).Object).Flags, 6), Object(Character.Storage(Character.CurStoragePage, A).Object).Picture + 255, Object(Character.Storage(Character.CurStoragePage, A).Object).Picture)
                        End If
                    End With
                Next A
                
                DrawBmpString3D Create3DString("Personal Storage [" & Character.CurStoragePage & "/" & Character.NumStoragePages & "]"), x1 + (STORAGEWINDOWWIDTH \ 2), y1 + 7, &HFFFFFFFF, True
        End Select
    End If
End Sub

Public Function GUIWindow_MouseDown(ByVal x As Long, ByVal y As Long, ByVal Button As Long) As Boolean
    Dim A As Long, b As Long, C As Long
    
    If (CurrentWindowFlags And WINDOW_FLAG_INVISIBLE) = 0 Then
        Select Case CurrentWindow
            Case WINDOW_INVALID
                GUIWindow_MouseDown = False
            Case WINDOW_TRADE
                If Character.Trading Then
                    GUIWindow_MouseDown = True
                    'Translate to window coords
                    x = x - (192 * WindowScaleX - (TRADEWINDOWWIDTH * WindowScaleX / 2))
                    y = y - (192 * WindowScaleY - (TRADEWINDOWHEIGHT * WindowScaleY / 2))
                    x = x / WindowScaleX
                    y = y / WindowScaleY
                    
                    If Button = 1 Then
                        If x >= 0 And x <= TRADEWINDOWWIDTH Then
                            If y >= 0 And y <= TRADEWINDOWHEIGHT Then
                                If x >= 52 And x <= 115 Then 'Accept
                                    If y >= 103 And y <= 130 Then
                                        If TradeData.Tradestate(0) = TRADE_STATE_OPEN Then
                                            SendSocket Chr$(73) + Chr$(6)
                                            TradeData.Tradestate(0) = TRADE_STATE_ACCEPTED
                                            TradeData.Tradestate(2) = 0
                                        ElseIf TradeData.Tradestate(0) = TRADE_STATE_ACCEPTED Then
                                            If TradeData.Tradestate(1) = TRADE_STATE_ACCEPTED Then
                                                
                                                If TradeData.Tradestate(2) = 0 Then SendSocket Chr$(73) + Chr$(8)
                                                TradeData.Tradestate(2) = 1
                                            Else
                                                PrintChat "Other player must accept trade before finalizing.", 7, Options.FontSize
                                            End If
                                        End If
                                    End If
                                End If
                                If x >= 145 And x <= 227 Then 'Cancel
                                    If y >= 111 And y <= 138 Then
                                        SetGUIWindow WINDOW_INVALID
                                        SendSocket Chr$(73) + Chr$(3)
                                    End If
                                End If
                                
                            End If
                        End If
                        If x >= 12 And x <= 180 Then  'Clicked item in your list.
                            If y >= 19 And y <= 84 Then
                                A = (x - 12) \ 33
                                b = (y - 19) \ 33
                                A = (b * 5) + A + 1
                                If A > 0 And A <= 10 Then
                                    TradeData.CurrentObject = A
                                    If TradeData.YourObjects(A).Object > 0 Then
                                        CurInvObj = 105 + A
                                        DrawCurInvObj
                                    End If
                                End If
                            End If
                        End If
                        If x >= 189 And x <= 357 Then  'Clicked item in their list.
                            If y >= 19 And y <= 84 Then
                                A = (x - 189) \ 33
                                b = (y - 19) \ 33
                                A = (b * 5) + A + 1
                                If A > 0 And A <= 10 Then
                                    TradeData.CurrentObject = A + 10
                                    If TradeData.TheirObjects(A).Object > 0 Then
                                        CurInvObj = 115 + A
                                        DrawCurInvObj
                                    End If
                                End If
                            End If
                        End If
                    ElseIf Button = 2 Then
                        If x >= 12 And x <= 180 Then  'right clicked item in your list.
                            If y >= 19 And y <= 84 Then
                                A = (x - 12) \ 33
                                b = (y - 19) \ 33
                                A = (b * 5) + A + 1
                                If A > 0 And A <= 10 Then
                                    If TradeData.YourObjects(A).Object > 0 Then
                                        SendSocket Chr$(73) + Chr$(5) + Chr$(A)
                                        If A = TradeData.CurrentObject Then
                                            TradeData.CurrentObject = 0
                                            CurInvObj = 0
                                            DrawCurInvObj
                                        End If
                                    End If
                                End If
                            End If
                        End If
                            If x >= 52 And x <= 115 Then 'Un-Accept
                                If y >= 103 And y <= 130 Then
                                    If TradeData.Tradestate(0) = TRADE_STATE_ACCEPTED Then
                                        SendSocket Chr$(73) + Chr$(6)
                                    End If
                                End If
                            End If
                    End If
                Else
                    SetGUIWindow WINDOW_INVALID
                End If
            Case WINDOW_PLAYERCLICK
                With ClickWindow
                    If .Loaded And .Target.Target > 0 And .Target.Target <= MAXUSERS Then
                        If .Target.TargetType = TT_PLAYER And CMap = player(.Target.Target).map Then
                            x = CurX * 32 + CurSubX
                            y = CurY * 32 + CurSubY
                            
                            x = x / WindowScaleX
                            y = y / WindowScaleY
                            
                            x = x - .x - 2
                            y = y - .y - 2
                            
                       
                           
                            
                            If x >= 0 And x <= .Width - 2 Then
                                If y >= 0 And y <= .Height - 2 Then
                                    A = (((y \ 18) * 3) + (x \ 18)) + 1
                                    If A > 0 And A <= .NumIcons Then
                                        GUIWindow_MouseDown = True
                                        Select Case .Icons(A).IconType
                                            Case CW_PARTY 'Invite to Party
                                                If player(.Target.Target).Party = 0 Then
                                                    If Character.Party = 0 Then
                                                        SendSocket Chr$(69) + Chr$(1) + Chr$(0)
                                                    End If
                                                    SendSocket Chr$(69) + Chr$(3) + Chr$(.Target.Target)
                                                    PrintChat "You have invited " & player(.Target.Target).Name & " to join your party!", 14, Options.FontSize
                                                Else
                                                    If player(.Target.Target).Party <> Character.Party Then
                                                        PrintChat player(.Target.Target).Name + " is already in a party!", 14, Options.FontSize
                                                    Else
                                                        PrintChat player(.Target.Target).Name + " is already a member of your party!", 14, Options.FontSize
                                                    End If
                                                End If
                                            Case CW_TRADE 'Invite to Trade
                                                If Character.Trading = False Then
                                                    If TradeData.Tradestate(0) = TRADE_STATE_INVITED Then
                                                        SendSocket Chr$(73) + Chr$(2)
                                                    Else
                                                        SendSocket Chr$(73) + Chr$(1) + Chr$(.Target.Target)
                                                        PrintChat "You have invited " & player(.Target.Target).Name & " to trade.", 14, Options.FontSize
                                                    End If
                                                End If
                                            Case CW_INFO 'Examine Player
                                            Case CW_GUILD
                                                If .Target.Target > 0 Then
                                                    SendSocket Chr$(34) + Chr$(.Target.Target)
                                                End If
                                            Case CW_SKILL 'Skill
                                                If .Icons(A).IconData > 0 And .Icons(A).IconData <= MAX_SKILLS Then
                                                    CurrentTarget.Target = .Target.Target
                                                    CurrentTarget.TargetType = TT_PLAYER
                                                    If CanUseSkill(.Icons(A).IconData, 1, 0) = 0 Then
                                                        UseSkill .Icons(A).IconData
                                                    End If
                                                End If
                                        End Select
                                        SetGUIWindow WINDOW_INVALID
                                        GUIWindow_MouseDown = True
                                    End If
                                Else
                                    SetGUIWindow WINDOW_INVALID
                                    GUIWindow_MouseDown = False
                                End If
                            Else
                                SetGUIWindow WINDOW_INVALID
                                GUIWindow_MouseDown = False
                            End If
                        Else
                            SetGUIWindow WINDOW_INVALID
                            GUIWindow_MouseDown = False
                        End If
                    Else
                        SetGUIWindow WINDOW_INVALID
                        GUIWindow_MouseDown = False
                    End If
                End With
            Case WINDOW_CHARACTERCLICK
                With ClickWindow
                    x = CurX * 32 + CurSubX
                    y = CurY * 32 + CurSubY
                    
                    x = x / WindowScaleX
                    y = y / WindowScaleY
                    
                    x = x - .x - 2
                    y = y - .y - 2
                    
                 
                    
                    If x >= 0 And x <= .Width - 2 Then
                        If y >= 0 And y <= .Height - 2 Then
                            A = (((y \ 18) * 3) + (x \ 18)) + 1
                            If A > 0 And A <= .NumIcons Then
                                Select Case .Icons(A).IconType
                                    Case CW_PARTY 'Party
                                        If Character.Party = 0 Then
                                            SendSocket Chr$(69) + Chr$(1) + Chr$(1) 'Create Party
                                        Else
                                            SendSocket Chr$(69) + Chr$(2) 'Leave Party
                                        End If
                                    Case CW_GUILD
                                        If Character.Guild > 0 Then
                                            SendSocket Chr$(32) 'Leave Guild
                                        End If
                                    Case CW_SKILL 'Spell
                                        If .Icons(A).IconData > 0 And .Icons(A).IconData <= MAX_SKILLS Then
                                            CurrentTarget.Target = Character.Index
                                            CurrentTarget.TargetType = TT_CHARACTER
                                            If CanUseSkill(.Icons(A).IconData, 1, 0) = 0 Then
                                                UseSkill .Icons(A).IconData
                                            End If
                                        End If
                                End Select
                                SetGUIWindow WINDOW_INVALID
                                GUIWindow_MouseDown = True
                            End If
                        Else
                            SetGUIWindow WINDOW_INVALID
                            GUIWindow_MouseDown = False
                        End If
                    Else
                        SetGUIWindow WINDOW_INVALID
                        GUIWindow_MouseDown = False
                    End If
                End With
            Case WINDOW_NPCCLICK
                With ClickWindow
                    x = CurX * 32 + CurSubX
                    y = CurY * 32 + CurSubY
                    
                    x = x / WindowScaleX
                    y = y / WindowScaleY
                    
                    x = x - .x - 2
                    y = y - .y - 2
                    
                    If x >= 0 And x <= .Width - 2 Then
                        If y >= 0 And y <= .Height - 2 Then
                            A = (((y \ 18) * 3) + (x \ 18)) + 1
                            If A > 0 And A <= .NumIcons Then
                                Select Case .Icons(A).IconType
                                    Case 1 'NPC Talk
                                        SendSocket Chr$(61) + Chr$(.Target.x) + Chr$(.Target.y)
                                    Case 2 'Shop
                                        SendSocket Chr$(52) + Chr$(.Target.x) + Chr$(.Target.y)
                                    Case 3 'Repair
                                        If CurInvObj > 0 And CurInvObj <= 20 Then
                                            With Character.Inv(CurInvObj)
                                                If .Object > 0 Then
                                                    If Object(.Object).Type = 6 Or Object(.Object).Type = 11 Then
                                                        PrintChat "You cannot repair this item!", 14, Options.FontSize
                                                    Else
                                                        SendSocket Chr$(65) + Chr$(1) + Chr$(ClickWindow.Target.x) + Chr$(ClickWindow.Target.y) + Chr$(CurInvObj)
                                                        MemCopy RepairObj, Character.Inv(CurInvObj), LenB(RepairObj)
                                                    End If
                                                Else
                                                    PrintChat "Please select an item in your inventory to be repaired!", 14, Options.FontSize
                                                End If
                                            End With
                                        Else
                                            PrintChat "Please select an item in your inventory to be repaired!", 14, Options.FontSize
                                        End If
                                    Case 4 'Bank
                                         SendSocket Chr$(75) + Chr$(0) + Chr$(.Target.x) + Chr$(.Target.y)
                                End Select
                                SetGUIWindow WINDOW_INVALID
                                GUIWindow_MouseDown = True
                            End If
                        Else
                            SetGUIWindow WINDOW_INVALID
                            GUIWindow_MouseDown = False
                        End If
                    Else
                        SetGUIWindow WINDOW_INVALID
                        GUIWindow_MouseDown = False
                    End If
                End With
            Case WINDOW_REPAIR
                GUIWindow_MouseDown = True
                'Translate to window coords
                x = x - (192 * WindowScaleX - (REPAIRWINDOWWIDTH * WindowScaleX / 2))
                y = y - (192 * WindowScaleY - (REPAIRWINDOWHEIGHT * WindowScaleY / 2))
                x = x / WindowScaleX
                y = y / WindowScaleY
                
                
                If Button = 1 Then
                    If y >= 128 And y <= 155 Then
                        If x >= 41 And x <= 118 Then 'Repair
                            SendSocket Chr$(65) + Chr$(2) + Chr$(LastNPCX) + Chr$(LastNPCY)
                        End If
                        If x >= 123 And x <= 200 Then 'Cancel
                            SetGUIWindow WINDOW_INVALID
                        End If
                    End If
                End If
            Case WINDOW_SHOP
                GUIWindow_MouseDown = True
                x = x - (192 * WindowScaleX - 352 * WindowScaleX \ 2)
                
                x = x / WindowScaleX
                
                
                If CurShopPage = 1 Then
                    If NumShopItems > 5 Then
                        y = y - (192 * WindowScaleY - ((68 * WindowScaleY + 50 * WindowScaleY * 5) \ 2))
                        y = y / WindowScaleY
                        b = 5
                    Else
                        y = y - (192 * WindowScaleY - ((68 * WindowScaleY + 50 * WindowScaleY * NumShopItems) \ 2))
                        y = y / WindowScaleY
                        b = NumShopItems
                    End If
                Else
                    y = y - (192 * WindowScaleY - ((68 * WindowScaleY + 50 * WindowScaleY * (NumShopItems - 5)) \ 2))
                    y = y / WindowScaleY
                    b = NumShopItems - 5
                End If
                If Button = 1 Then
                    If x >= 185 And x <= 268 Then
                        If y >= 27 + (b * 50) + 5 And y <= 27 + (b * 50) + 33 Then
                            If CurInvObj >= 26 And CurInvObj <= 35 Then
                                CurInvObj = 0
                                DrawCurInvObj
                            End If
                            SetGUIWindow WINDOW_INVALID
                        End If
                    End If
                    If x >= 97 And x <= 180 Then
                        If y >= 27 + (b * 50) + 5 And y <= 27 + (b * 50) + 33 Then
                            If CurInvObj >= 26 And CurInvObj <= 35 Then
                                With SaleItem(CurInvObj - 26)
                                    If .GiveObject >= 1 And .TakeObject >= 1 Then
                                        SendSocket Chr$(53) + Chr$(LastNPCX) + Chr$(LastNPCY) + Chr$(CurInvObj - 26)
                                    End If
                                End With
                            End If
                        End If
                    End If
                    If x >= 53 And x <= 85 Then 'Next list
                        If y >= 27 + (b * 50) + 2 And y <= 27 + (b * 50) + 35 Then
                            If NumShopItems > 5 Then
                                If CurShopPage = 1 Then
                                    CurShopPage = 2
                                Else
                                    CurShopPage = 1
                                End If
                            End If
                        End If
                    End If
                    If y >= 27 And y <= 27 + (b * 50) Then 'Selected item in list
                        C = (y - 27) \ 50
                        If C >= 0 And C < 5 Then
                            CurInvObj = 26 + C + ((CurShopPage - 1) * 5)
                            SetTab tsInventory
                            DrawCurInvObj
                        End If
                    End If
                End If
            Case WINDOW_STORAGE
                GUIWindow_MouseDown = True
                x = x - (192 * WindowScaleX - (STORAGEWINDOWWIDTH * WindowScaleX / 2))
                y = y - (192 * WindowScaleY - (STORAGEWINDOWHEIGHT * WindowScaleY / 2))
                x = x / WindowScaleX
                y = y / WindowScaleY
                
                If Button = 1 Then
                    If y >= 213 And y <= 240 Then
                        If x >= 65 And x <= 147 Then 'Close
                            StorageOpen = False
                            If CurInvObj > 35 And CurInvObj <= 55 Then
                                CurInvObj = 0
                                DrawCurInvObj
                            End If
                            SetGUIWindow WINDOW_INVALID
                        End If
                    End If
                    
                    If x >= 16 And x <= 192 Then
                        If y >= 60 And y <= 196 Then
                            A = (x - 16) \ 35
                            b = (y - 60) \ 35
                            A = (b * 5) + A + 1
                            If A > 0 And A <= 20 Then
                                CurStorageObj = A
                                CurInvObj = 35 + A
                                SetTab tsInventory
                                DrawCurInvObj
                            End If
                        End If
                    End If
                    
                    If y >= 216 And y <= 238 Then 'Previous/Next page
                        If x >= 34 And x <= 61 Then
                            If Character.CurStoragePage > 1 Then
                                SendSocket Chr$(75) + Chr$(4) + Chr$(LastNPCX) + Chr$(LastNPCY)
                            End If
                        End If
                        If x >= 151 And x <= 178 Then
                            If Character.CurStoragePage < Character.NumStoragePages Then
                                SendSocket Chr$(75) + Chr$(3) + Chr$(LastNPCX) + Chr$(LastNPCY)
                            End If
                        End If
                    End If
                ElseIf Button = 2 Then
                    If x >= 16 And x <= 192 Then
                        If y >= 60 And y <= 196 Then
                            A = (x - 16) \ 35
                            b = (y - 60) \ 35
                            A = (b * 5) + A + 1
                            If A > 0 And A <= 20 Then
                                CurStorageObj = A
                                With Character.Storage(Character.CurStoragePage, CurStorageObj)
                                    If .Object > 0 Then
                                        If Object(.Object).Type = 6 Or Object(.Object).Type = 11 Then
                                            CurrentWindowFlags = CurrentWindowFlags Or WINDOW_FLAG_INVISIBLE
                                            TempVar2 = .Value
                                            TempVar1 = TempVar2
                                            TempVar3 = CurStorageObj
                                            frmMain.picDrop.Visible = True
                                            frmMain.lblDropTitle = "Withdraw how much?"
                                            frmMain.txtDrop = TempVar1
                                            frmMain.txtDrop.SelStart = 0
                                            frmMain.txtDrop.SelLength = Len(frmMain.txtDrop)
                                            frmMain.txtDrop.SetFocus
                                        Else
                                            SendSocket Chr$(75) + Chr$(2) + Chr$(LastNPCX) + Chr$(LastNPCY) + Chr$(CurStorageObj) + QuadChar(0)
                                        End If
                                    End If
                                End With
                            End If
                        End If
                    End If
                End If
        End Select
    End If
End Function
