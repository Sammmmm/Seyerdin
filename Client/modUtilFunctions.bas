Attribute VB_Name = "ModUtilFunctions"
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


Private Type guid
   Data1 As Long
   data2 As Integer
   data3 As Integer
   data4(7) As Byte
End Type

Private Type PicBmp
   Size As Long
   Type As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public SkillsPerLevel As Byte
Public StatsPerLevel As Byte

Public StatRate1 As Byte
Public StatRate2 As Byte


Public baseEnergyRegen As Byte
Public baseManaRegen As Byte
Public BaseHPRegen As Byte
Public LevelsPerHpRegen As Byte
Public LevelsPerManaRegen As Byte

Public StrengthPerDamage(1 To 3) As Byte

Public AgilityPerCritChance(1 To 3) As Byte
Public AgilityPerDodgeChance(1 To 3) As Byte

Public EndurancePerBlockChance(1 To 3) As Byte
Public EndurancePerEnergy(1 To 3) As Byte
Public EndurancePerEnergyRegen(1 To 3) As Byte

Public ManaPerIntelligence(1 To 3) As Byte
Public IntelligencePerManaRegen(1 To 3) As Byte

Public ConstitutionPerHPRegen(1 To 3) As Byte
Public HPPerConstitution(1 To 3) As Byte

Public PietyPerMagicResist(1 To 3) As Byte
Public PietyPerHPRegen(1 To 3) As Byte
Public PietyPerManaRegen(1 To 3) As Byte
Public PietyPerHP(1 To 3) As Byte
Public PietyPerMana(1 To 3) As Byte
Public PietyPerDodge(1 To 3) As Byte
Public PietyPerCrit(1 To 3) As Byte
Public PietyPerBlock(1 To 3) As Byte

Public GenericStatPerBonus(1 To 3) As Byte
Public GenericPietyPerBonus(1 To 3) As Byte

Dim oSHA As New clsSHAAlgorithm

'''''''''
'Windows API Declarations
'''''''''
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As PicBmp, RefIID As guid, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
'Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Public Const STD_OUTPUT_HANDLE = -11&
Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
'Declare Function FlushFileBuffers Lib "Kernel32" (ByVal hFile As Long) As Long


Public Sub SetWalkSpeed()
    If map.Tile(cX, cY).Att = 26 Then
        If (KeyDown(Options.RunKey)) And Options.AutoRun Then
            CWalkStep = map.Tile(cX, cY).AttData(0)
        ElseIf ((KeyDown(Options.RunKey)) < 0 Or Options.AutoRun) And Character.Energy > 0 Then
                CWalkStep = map.Tile(cX, cY).AttData(1)
        Else
                CWalkStep = map.Tile(cX, cY).AttData(0)
        End If
        If KeyDown(Options.StrafeKey) Then CWalkStep = map.Tile(cX, cY).AttData(0) - 1
    Else
        If (KeyDown(Options.RunKey)) And Options.AutoRun Then
            CWalkStep = 8
        ElseIf ((KeyDown(Options.RunKey)) Or Options.AutoRun) And Character.Energy > 0 Then
                CWalkStep = 16
        Else
                CWalkStep = 8
        End If
        If KeyDown(Options.StrafeKey) Then CWalkStep = 7
    End If
End Sub

Public Function GetWindowScreenshot() As Long
'
' Function to create screeenshot of specified window and store at specified path
'
    On Error GoTo ErrorHandler

    Dim hDCSrc As Long
    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim WidthSrc As Long
    Dim HeightSrc As Long
    Dim Pic As PicBmp
    Dim IPic As IPicture
    Dim IID_IDispatch As guid
    Dim rc As RECT
    Dim pictr As PictureBox
    
    
    'Bring window on top of all windows if specified
    'If BringFront = 1 Then BringWindowToTop frmMain.hwnd
    
    'Get Window Size
    GetWindowRect frmMain.hwnd, rc
    WidthSrc = rc.Right - rc.Left
    HeightSrc = rc.Bottom - rc.Top
    
    'Get Window  device context
    hDCSrc = GetWindowDC(frmMain.hwnd)
    
    'create a memory device context
    hDCMemory = CreateCompatibleDC(hDCSrc)
    
    'create a bitmap compatible with window hdc
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    
    'copy newly created bitmap into memory device context
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    
    
    'copy window window hdc to memory hdc
    Call BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, _
                hDCSrc, 0, 0, vbSrcCopy)
      
    'Get Bmp from memory Dc
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    
    'release the created objects and free memory
    Call DeleteDC(hDCMemory)
    Call ReleaseDC(frmMain.hwnd, hDCSrc)
    
    'fill in OLE IDispatch Interface ID
    With IID_IDispatch
       .Data1 = &H20400
       .data4(0) = &HC0
       .data4(7) = &H46
     End With
    
    'fill Pic with necessary parts
    With Pic
       .Size = Len(Pic)         'Length of structure
       .Type = vbPicTypeBitmap  'Type of Picture (bitmap)
       .hBmp = hBmp             'Handle to bitmap
       .hPal = 0&               'Handle to palette (may be null)
     End With

    'create OLE Picture object
    Call OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
    
    'return the new Picture object
    For hBmp = 1 To 1000000
        If Dir("screenshot" + Str(hBmp) + ".bmp") = "" Then Exit For
    Next hBmp
    SavePicture IPic, "screenshot" + Str(hBmp) + ".bmp"
    PrintChat "screenshot" + Str(hBmp) + ".bmp saved", 15, Options.FontSize
    GetWindowScreenshot = 1
    Exit Function
    
ErrorHandler:
    GetWindowScreenshot = 0
End Function

Function FtoDW(F As Single) As Long
    Dim Buf As D3DXBuffer
    Dim l As Long
    Set Buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData Buf, 0, 4, 1, F
    D3DX.BufferGetData Buf, 0, 4, 1, l
    FtoDW = l
End Function

Sub calculateAttackDamage()
Dim A As Long
With Character
    .statBaseAttack = 0
    If .Equipped(1).Object > 0 Then
        .statBaseAttack = (Object(.Equipped(1).Object).ObjData(2) * 256 + Object(.Equipped(1).Object).ObjData(3) + Object(.Equipped(1).Object).ObjData(1)) / 2
        If ExamineBit(Object(.Equipped(1).Object).Flags, 3) Then
            .statBaseAttack = .statBaseAttack + (.statBaseAttack * .SkillLevels(SKILL_TWOHAND)) / 100
        End If
    End If
    
    If .Equipped(2).Object > 0 Then
        With Object(.Equipped(2).Object)
            If ExamineBit(.Flags, 5) Then
                Character.statBaseAttack = Character.statBaseAttack + (.ObjData(2) * 256 + .ObjData(3) + .ObjData(1)) / 2
            End If
        End With
    End If
    
    If (.Equipped(5).Object > 0) Then
        With Object(.Equipped(5).Object)
            If .ObjData(0) = 0 Then
                Character.statBaseAttack = Character.statBaseAttack + .ObjData(2)
            End If
        End With
    End If
    
    For A = 1 To 5
        If .Equipped(A).Object > 0 Then
            With .Equipped(A)
                If .Prefix > 0 Then
                    If Prefix(.Prefix).ModType = 18 Then
                        If (.PrefixValue And 63) > 0 Then
                            Character.statBaseAttack = Character.statBaseAttack + (.PrefixValue And 63)
                        End If
                    End If
                End If
                If .Suffix > 0 Then
                    If Prefix(.Suffix).ModType = 18 Then
                        If .SuffixValue > 0 Then
                            Character.statBaseAttack = Character.statBaseAttack + (.SuffixValue And 63)
                        End If
                    End If
                End If
                If .Affix > 0 Then
                    If Prefix(.Affix).ModType = 18 Then
                        If .AffixValue > 0 Then
                            Character.statBaseAttack = Character.statBaseAttack + (.AffixValue And 63)
                        End If
                    End If
                End If
            End With
        End If
    Next A
    
    
    
    
    .statBaseAttack = .statBaseAttack + GetStatPerBonus(.strength, StrengthPerDamage)
    
            If GetStatusEffect(.Index, SE_INVULNERABILITY) Then .statBaseAttack = .statBaseAttack / 2
        
        If GetStatusEffect(.Index, SE_BERSERK) Then .statBaseAttack = .statBaseAttack * 1.3
        If GetStatusEffect(.Index, SE_FIERYESSENCE) > 0 Then .statBaseAttack = .statBaseAttack * 1.1
        If .SkillLevels(SKILL_VENGEFUL) Then .statBaseAttack = .statBaseAttack * 1.1
        If .SkillLevels(SKILL_GREATFORCE) Then .statBaseAttack = .statBaseAttack * 1.15
    
    
End With
End Sub

Sub calculateEvasion()
Dim C As Long, A As Long
    With Character
        .statEvasion = 0
        C = GetStatPerBonus(.Agility, AgilityPerDodgeChance) + GetStatPerBonus(.Wisdom, PietyPerDodge)
        C = C + .SkillLevels(SKILL_EVASION) * 1.5
        C = C + .SkillLevels(SKILL_AGILITY) * 17
        If GetStatusEffect(.Index, SE_ETHEREALITY) Then C = C + .SkillLevels(SKILL_ETHEREALITY) * 2
        .statEvasion = C
        
    For A = 1 To 5
        If .Equipped(A).Object > 0 Then
            With .Equipped(A)
                If .Prefix > 0 Then
                    If Prefix(.Prefix).ModType = 25 Then  'Resist Poison
                        If (.PrefixValue And 63) > 0 Then
                            Character.statEvasion = Character.statEvasion + (.PrefixValue And 63)
                        End If
                    End If
                End If
                If .Suffix > 0 Then
                    If Prefix(.Suffix).ModType = 25 Then
                        If .SuffixValue > 0 Then
                            Character.statEvasion = Character.statEvasion + (.SuffixValue And 63)
                        End If
                    End If
                End If
                If .Affix > 0 Then
                    If Prefix(.Affix).ModType = 25 Then
                        If .AffixValue > 0 Then
                            Character.statEvasion = Character.statEvasion + (.AffixValue And 63)
                        End If
                    End If
                End If
            End With
        End If
    Next A
        
        
    End With
End Sub

Sub calculateBlock()
Dim C As Long, A As Long
    With Character
        .statBlock = 0
        
        If .Equipped(2).Object > 0 Then
            If Not ExamineBit(Object(.Equipped(2).Object).Flags, 5) Then
                .statBlock = Object(.Equipped(2).Object).ObjData(1)
                .statBlock = .statBlock + .SkillLevels(SKILL_SHIELDMASTERY)
                .statBlock = .statBlock + .SkillLevels(SKILL_GUARDIAN) * 5
                .statBlock = .statBlock + GetStatPerBonus(.Endurance, EndurancePerBlockChance) + GetStatPerBonus(.Wisdom, PietyPerBlock)
            End If
        End If
        
        
        For A = 1 To 5
            If .Equipped(A).Object > 0 Then
                With .Equipped(A)
                    If .Prefix > 0 Then
                        If Prefix(.Prefix).ModType = 24 Then
                            If (.PrefixValue And 63) > 0 Then
                                Character.statBlock = Character.statBlock + (.PrefixValue And 63)
                            End If
                        End If
                    End If
                    If .Suffix > 0 Then
                        If Prefix(.Suffix).ModType = 24 Then
                            If .SuffixValue > 0 Then
                                Character.statBlock = Character.statBlock + (.SuffixValue And 63)
                            End If
                        End If
                    End If
                    If .Affix > 0 Then
                        If Prefix(.Affix).ModType = 24 Then
                            If .AffixValue > 0 Then
                                Character.statBlock = Character.statBlock + (.AffixValue And 63)
                            End If
                        End If
                    End If
                End With
            End If
        Next A
        
    End With
End Sub

Sub calculatePoisonResist()
Dim A As Long
    With Character
        .statPoisonResist = .ConstitutionMod / 5
        
        For A = 1 To 5
            If .Equipped(A).Object > 0 Then
                With .Equipped(A)
                    If .Prefix > 0 Then
                        If Prefix(.Prefix).ModType = 16 Then 'Resist Poison
                            If (.PrefixValue And 63) > 0 Then
                                Character.statPoisonResist = Character.statPoisonResist + (.PrefixValue And 63)
                            End If
                        End If
                    End If
                    If .Suffix > 0 Then
                        If Prefix(.Suffix).ModType = 16 Then
                            If .SuffixValue > 0 Then
                                Character.statPoisonResist = Character.statPoisonResist + (.SuffixValue And 63)
                            End If
                        End If
                    End If
                    If .Affix > 0 Then
                        If Prefix(.Affix).ModType = 16 Then
                            If .AffixValue > 0 Then
                                Character.statPoisonResist = Character.statPoisonResist + (.AffixValue And 63)
                            End If
                        End If
                    End If
                End With
            End If
        Next A
        
        .statPoisonResist = (100 + .statPoisonResist) / 10
    
    End With
End Sub

Sub calculateHpRegen()
Dim A As Long
    With Character
        .statHpRegenLow = 2 + (.Level \ LevelsPerHpRegen)
        .statHpRegenLow = .statHpRegenLow + .SkillLevels(SKILL_VIGOR) * 2
        .statHpRegenHigh = 0
        If .SkillLevels(SKILL_CONVALESCENCE) Then
            .statHpRegenLow = .statHpRegenLow + 1 + .SkillLevels(SKILL_CONVALESCENCE) \ 2
        End If
        
        For A = 1 To 5
            If .Equipped(A).Object > 0 Then
                With .Equipped(A)
                    If .Prefix > 0 Then
                        If Prefix(.Prefix).ModType = 7 Then 'hp reg
                            If (.PrefixValue And 63) > 0 Then
                                Character.statHpRegenLow = Character.statHpRegenLow + (.PrefixValue And 63)
                            End If
                        End If
                    End If
                    If .Suffix > 0 Then
                        If Prefix(.Suffix).ModType = 7 Then
                            If .SuffixValue > 0 Then
                                Character.statHpRegenLow = Character.statHpRegenLow + (.SuffixValue And 63)
                            End If
                        End If
                    End If
                    If .Affix > 0 Then
                        If Prefix(.Affix).ModType = 7 Then
                            If .AffixValue > 0 Then
                                Character.statHpRegenLow = Character.statHpRegenLow + (.AffixValue And 63)
                            End If
                        End If
                    End If
                End With
            End If
        Next A

        .statHpRegenHigh = .statHpRegenLow + GetStatPerBonusHigh(.Constitution, ConstitutionPerHPRegen) + GetStatPerBonusHigh(.Wisdom, PietyPerHPRegen)
        .statHpRegenLow = .statHpRegenLow + GetStatPerBonus(.Constitution, ConstitutionPerHPRegen) + GetStatPerBonus(.Wisdom, PietyPerHPRegen)
        
        If combatCounter = 0 Then
            .statHpRegenLow = .statHpRegenLow * 3
            .statHpRegenHigh = .statHpRegenHigh * 3
        End If
    
    End With
End Sub

Sub calculateManaRegen()
Dim A As Long
    With Character
        .statManaRegenHigh = 0
        .statManaRegenLow = 1 + (.Level \ LevelsPerManaRegen)

        For A = 1 To 5
            If .Equipped(A).Object > 0 Then
                With .Equipped(A)
                    If .Prefix > 0 Then
                        If Prefix(.Prefix).ModType = 8 Then 'mana reg
                            If (.PrefixValue And 63) > 0 Then
                                Character.statManaRegenLow = Character.statManaRegenLow + (.PrefixValue And 63)
                            End If
                        End If
                    End If
                    If .Suffix > 0 Then
                        If Prefix(.Suffix).ModType = 8 Then
                            If .SuffixValue > 0 Then
                                Character.statManaRegenLow = Character.statManaRegenLow + (.SuffixValue And 63)
                            End If
                        End If
                    End If
                    If .Affix > 0 Then
                        If Prefix(.Affix).ModType = 8 Then
                            If .AffixValue > 0 Then
                                Character.statManaRegenLow = Character.statManaRegenLow + (.AffixValue And 63)
                            End If
                        End If
                    End If
                End With
            End If
        Next A

        .statManaRegenHigh = .statManaRegenLow + GetStatPerBonusHigh(.Intelligence, IntelligencePerManaRegen) + GetStatPerBonusHigh(.Wisdom, PietyPerManaRegen)
        .statManaRegenLow = .statManaRegenLow + GetStatPerBonus(.Intelligence, IntelligencePerManaRegen) + GetStatPerBonus(.Wisdom, PietyPerManaRegen)

        If combatCounter = 0 Then
            .statManaRegenHigh = .statManaRegenHigh * 3
            .statManaRegenLow = .statManaRegenLow * 3
        End If
    
    End With
End Sub

Sub calculateBodyArmor()
    Dim A As Long
    With Character
        .statBodyArmor = 0
        If (.Equipped(5).Object > 0) Then
            With Object(.Equipped(5).Object)
                If .ObjData(0) = 1 Then
                    Character.statBodyArmor = Character.statBodyArmor + .ObjData(2)
                End If
            End With
        End If
        
        If .Equipped(3).Object > 0 Then
            With .Equipped(3)
                Character.statBodyArmor = Character.statBodyArmor + Object(.Object).ObjData(1) * 256 + Object(.Object).ObjData(2)
            End With
        End If
        
        If .buff = BUFF_HOLYARMOR Then
            .statBodyArmor = .statBodyArmor + (.SkillLevels(SKILL_HOLYARMOR) + 3) / 3 + ((.Wisdom + .WisdomMod) / 20)
        End If
        
        If .SkillLevels(SKILL_ARMORMASTERY) Then
            .statBodyArmor = .statBodyArmor + (.statBodyArmor * (.SkillLevels(SKILL_ARMORMASTERY) + 5) \ 100)
        End If
        If GetStatusEffect(.Index, SE_MASTERDEFENSE) Then
            .statBodyArmor = .statBodyArmor * 2
        End If
        
        If GetStatusEffect(.Index, SE_INVULNERABILITY) Then
            .statBodyArmor = 1000
        End If
        
    End With
End Sub
Sub calculateHeadArmor()
    Dim A As Long
    With Character
        .statHeadArmor = 0
        If (.Equipped(5).Object > 0) Then
            With Object(.Equipped(5).Object)
                If .ObjData(0) = 1 Then
                    Character.statHeadArmor = Character.statHeadArmor + .ObjData(2)
                End If
            End With
        End If
        
        If .Equipped(4).Object > 0 Then
            With .Equipped(4)
                Character.statHeadArmor = Character.statHeadArmor + Object(.Object).ObjData(1) * 256 + Object(.Object).ObjData(2)
            End With
        End If
        
        If .buff = BUFF_HOLYARMOR Then
            .statHeadArmor = .statHeadArmor + (.SkillLevels(SKILL_HOLYARMOR) + 3) / 3 + ((.Wisdom + .WisdomMod) / 20)
        End If
        
        If .SkillLevels(SKILL_ARMORMASTERY) Then
            .statHeadArmor = .statHeadArmor + (.statHeadArmor * (.SkillLevels(SKILL_ARMORMASTERY) + 5) \ 100)
        End If

        If GetStatusEffect(.Index, SE_MASTERDEFENSE) Then
            .statHeadArmor = .statHeadArmor * 2
        End If
        
        If GetStatusEffect(.Index, SE_INVULNERABILITY) Then
            .statHeadArmor = 1000
        End If
        

    End With
End Sub

Sub calculateCriticalChance()
Dim A As Long
Dim b As Long
    With Character
        .statCritical = 5
        .statCritical = .statCritical + .SkillLevels(SKILL_PERCEPTION) * 10
        .statCritical = .statCritical + CInt(.SkillLevels(SKILL_POIGNANCY) / 2.5)
        .statCritical = .statCritical + GetStatPerBonus(.Agility, AgilityPerCritChance)
        
        For A = 1 To 5
            If .Equipped(A).Object > 0 Then
                With .Equipped(A)
                    If .Prefix > 0 Then
                        If Prefix(.Prefix).ModType = 15 Then 'Resist Poison
                            If (.PrefixValue And 63) > 0 Then
                                b = b + (.PrefixValue And 63)
                            End If
                        End If
                    End If
                    If .Suffix > 0 Then
                        If Prefix(.Suffix).ModType = 15 Then
                            If .SuffixValue > 0 Then
                                b = b + (.SuffixValue And 63)
                            End If
                        End If
                    End If
                    If .Affix > 0 Then
                        If Prefix(.Affix).ModType = 15 Then
                            If .AffixValue > 0 Then
                                b = b + (.AffixValue And 63)
                            End If
                        End If
                    End If
                End With
            End If
        Next A
        If b > 25 Then b = 25
        .statCritical = .statCritical + b
        
        If GetStatusEffect(.Index, SE_DEADLYCLARITY) Then
            .statCritical = .statCritical + 10 + .SkillLevels(SKILL_DEADLYCLARITY)
        End If
        
    End With
End Sub


Sub calculatePhysicalDefense()
    With Character
        Dim A As Long, b As Long
        For A = 1 To 5
            If .Equipped(A).Object > 0 Then
                With .Equipped(A)
                    If .Prefix > 0 Then
                        If Prefix(.Prefix).ModType = 21 Then
                            If (.PrefixValue And 63) > 0 Then
                                b = b + (.PrefixValue And 63)
                            End If
                        End If
                    End If
                    If .Suffix > 0 Then
                        If Prefix(.Suffix).ModType = 21 Then
                            If .SuffixValue > 0 Then
                                b = b + (.SuffixValue And 63)
                            End If
                        End If
                    End If
                    If .Affix > 0 Then
                        If Prefix(.Affix).ModType = 21 Then
                            If .AffixValue > 0 Then
                                b = b + (.AffixValue And 63)
                            End If
                        End If
                    End If
                End With
            End If
        Next A
        .statPhysicalDefense = b
    End With
End Sub

Sub calculatePhysicalResist()
    With Character
            
        Dim A As Long, b As Long
        For A = 1 To 5
            If .Equipped(A).Object > 0 Then
                With .Equipped(A)
                    If .Prefix > 0 Then
                        If Prefix(.Prefix).ModType = 22 Then
                            If (.PrefixValue And 63) > 0 Then
                                b = b + (.PrefixValue And 63)
                            End If
                        End If
                    End If
                    If .Suffix > 0 Then
                        If Prefix(.Suffix).ModType = 22 Then
                            If .SuffixValue > 0 Then
                                b = b + (.SuffixValue And 63)
                            End If
                        End If
                    End If
                    If .Affix > 0 Then
                        If Prefix(.Affix).ModType = 22 Then
                            If .AffixValue > 0 Then
                                b = b + (.AffixValue And 63)
                            End If
                        End If
                    End If
                End With
            End If
        Next A
        .statPhysicalResist = b
    End With
End Sub

Sub calculateMagicDefense()
    With Character
        Dim A As Long, b As Long
        For A = 1 To 5
            If .Equipped(A).Object > 0 Then
                With .Equipped(A)
                    If .Prefix > 0 Then
                        If Prefix(.Prefix).ModType = 20 Then
                            If (.PrefixValue And 63) > 0 Then
                                b = b + (.PrefixValue And 63)
                            End If
                        End If
                    End If
                    If .Suffix > 0 Then
                        If Prefix(.Suffix).ModType = 0 Then
                            If .SuffixValue > 0 Then
                                b = b + (.SuffixValue And 63)
                            End If
                        End If
                    End If
                    If .Affix > 0 Then
                        If Prefix(.Affix).ModType = 20 Then
                            If .AffixValue > 0 Then
                                b = b + (.AffixValue And 63)
                            End If
                        End If
                    End If
                End With
            End If
        Next A
        .statMagicDefense = b
    End With
End Sub


Sub calculateMagicResist()
Dim A As Long
    With Character
        .statMagicResist = 0
        If GetStatusEffect(.Index, SE_MASTERDEFENSE) Then
            .statMagicResist = 10
        End If
        .statMagicResist = .statMagicResist + 8 * .SkillLevels(SKILL_FANATACISM)
        
        For A = 1 To 5
            If .Equipped(A).Object > 0 Then
                With .Equipped(A)
                    If .Prefix > 0 Then
                        If Prefix(.Prefix).ModType = 9 Then 'Resist Magic
                            If (.PrefixValue And 63) > 0 Then
                                Character.statMagicResist = Character.statMagicResist + (.PrefixValue And 63)
                            End If
                        End If
                    End If
                    If .Suffix > 0 Then
                        If Prefix(.Suffix).ModType = 9 Then
                            If .SuffixValue > 0 Then
                                Character.statMagicResist = Character.statMagicResist + (.SuffixValue And 63)
                            End If
                        End If
                    End If
                    If .Suffix > 0 Then
                        If Prefix(.Affix).ModType = 9 Then
                            If .SuffixValue > 0 Then
                                Character.statMagicResist = Character.statMagicResist + (.AffixValue And 63)
                            End If
                        End If
                    End If
                End With
            End If
        Next A
        .statMagicResist = .statMagicResist + GetStatPerBonus(.Wisdom, PietyPerMagicResist)
        
    End With
End Sub

Sub calculateAttackSpeed()
Dim A As Long, b As Long
    With Character
        .statAttackSpeed = .AttackSpeed + 10
    End With
End Sub

Sub calculateLeachHp()
Dim A As Long
    With Character
    .statLeachHP = 0
     For A = 1 To 5
        If .Equipped(A).Object > 0 Then
                With .Equipped(A)
                    If .Prefix > 0 Then
                        If Prefix(.Prefix).ModType = 13 Then 'leach
                            If (.PrefixValue And 63) > 0 Then
                                Character.statLeachHP = Character.statLeachHP + (.PrefixValue And 63)
                            End If
                        End If
                    End If
                    If .Suffix > 0 Then
                        If Prefix(.Suffix).ModType = 13 Then
                            If .SuffixValue > 0 Then
                                Character.statLeachHP = Character.statLeachHP + (.SuffixValue And 63)
                            End If
                        End If
                    End If
                    If .Affix > 0 Then
                        If Prefix(.Affix).ModType = 13 Then
                            If .AffixValue > 0 Then
                                Character.statLeachHP = Character.statLeachHP + (.AffixValue And 63)
                            End If
                        End If
                    End If
                End With
            End If
        Next A
    End With
End Sub

Sub calculateMagicFind()
Dim A As Long
    Character.statMagicFind = 0
    With Character
     For A = 1 To 5
        If .Equipped(A).Object > 0 Then
                With .Equipped(A)
                    If .Prefix > 0 Then
                        If Prefix(.Prefix).ModType = 12 Then 'mf
                            If (.PrefixValue And 63) > 0 Then
                                Character.statMagicFind = Character.statMagicFind + (.PrefixValue And 63)
                            End If
                        End If
                    End If
                    If .Suffix > 0 Then
                        If Prefix(.Suffix).ModType = 12 Then
                            If .SuffixValue > 0 Then
                                Character.statMagicFind = Character.statMagicFind + (.SuffixValue And 63)
                            End If
                        End If
                    End If
                    If .Affix > 0 Then
                        If Prefix(.Affix).ModType = 12 Then
                            If .AffixValue > 0 Then
                                Character.statMagicFind = Character.statMagicFind + (.AffixValue And 63)
                            End If
                        End If
                    End If
                End With
            End If
        Next A
    End With
End Sub

Sub calculateEnergyRegen()
Dim A As Long
    With Character
        .statEnergyRegenHigh = 0
        .statEnergyRegenLow = 2

        .statEnergyRegenHigh = .statEnergyRegenLow + GetStatPerBonusHigh(.Endurance, EndurancePerEnergyRegen)
        .statEnergyRegenLow = .statEnergyRegenLow + GetStatPerBonus(.Endurance, EndurancePerEnergyRegen)
    
    End With
End Sub

'I put this here for lack of a better place to put it
Public Function GetTickCount() As Long
    Dim curFreq As Currency
    Dim curTime As Currency
    
    QueryPerformanceFrequency curFreq
    QueryPerformanceCounter curTime
    
    GetTickCount = CLng(curTime / (curFreq / 1000)) '- tickCountMod
End Function

Sub NewChatWindow()
Dim A As Long
    For A = 1 To 5
        If chatForms(A) Is Nothing Then
            Set chatForms(A) = New frmChatWindow
            chatForms(A).Top = (A - 1) * chatForms(A).Height / 18
            chatForms(A).Show vbModeless, frmMain
            frmMain.SetFocus
            Exit Sub
        End If
        If Not chatForms(A).Visible Then
            Unload chatForms(A)
            Set chatForms(A) = New frmChatWindow
            chatForms(A).Top = (A - 1) * chatForms(A).Height / 18
            chatForms(A).Show vbModeless, frmMain
            frmMain.SetFocus
            Exit Sub
        End If
    Next A
End Sub

Sub CreateKeyCodeList()
    Dim A As Long
    For A = 0 To MAXKEYCODES
        KeyCodeList(A).NotChatKey = True
        KeyCodeList(A).CapitalKeyCode = 0
    Next A
    KeyCodeList(0).KeyCode = 255
    KeyCodeList(0).Text = "No Binding"
    
    'numbers
    For A = 1 To 10
        KeyCodeList(A).KeyCode = 47 + A
        KeyCodeList(A).Text = Chr(47 + A)
        KeyCodeList(A).NotChatKey = False
    Next A
    For A = 1 To 10
        If A = 1 Then KeyCodeList(A).CapitalKeyCode = 33
        If A = 2 Then KeyCodeList(A).CapitalKeyCode = 64
        If A = 3 Then KeyCodeList(A).CapitalKeyCode = 35
        If A = 4 Then KeyCodeList(A).CapitalKeyCode = 36
        If A = 5 Then KeyCodeList(A).CapitalKeyCode = 37
        If A = 6 Then KeyCodeList(A).CapitalKeyCode = 94
        If A = 7 Then KeyCodeList(A).CapitalKeyCode = 38
        If A = 8 Then KeyCodeList(A).CapitalKeyCode = 42
        If A = 9 Then KeyCodeList(A).CapitalKeyCode = 40
        If A = 10 Then KeyCodeList(A).CapitalKeyCode = 41
    Next A
    
    
    'letters
    For A = 11 To 36
        KeyCodeList(A).CapitalKeyCode = 86 + A
        KeyCodeList(A).KeyCode = 54 + A
        KeyCodeList(A).Text = Chr(54 + A)
        KeyCodeList(A).NotChatKey = False
    Next A
    
    KeyCodeList(37).Text = "Caps Lock"
    KeyCodeList(37).KeyCode = vbKeyCapital
    
    KeyCodeList(38).Text = "+"
    KeyCodeList(38).KeyCode = 43
    KeyCodeList(38).CapitalKeyCode = 61
    KeyCodeList(38).NotChatKey = False
    
    KeyCodeList(39).Text = "-"
    KeyCodeList(39).KeyCode = 45
    KeyCodeList(39).CapitalKeyCode = 95
    KeyCodeList(39).NotChatKey = False
    
    KeyCodeList(40).KeyCode = vbKeySpace
    KeyCodeList(40).Text = "Spacebar"

    KeyCodeList(41).KeyCode = &H10
    KeyCodeList(41).Text = "Shift"
    KeyCodeList(42).KeyCode = &H11
    KeyCodeList(42).Text = "Control"
    KeyCodeList(43).KeyCode = &H12
    KeyCodeList(43).Text = "Alt"
    
    For A = 44 To 53
        KeyCodeList(A).KeyCode = 112 + A - 44
        KeyCodeList(A).Text = Replace("F" & Str(A - 43), " ", "")
    Next A
    
    KeyCodeList(54).KeyCode = vbKeyUp
    KeyCodeList(54).Text = "Up"
    KeyCodeList(55).KeyCode = vbKeyDown
    KeyCodeList(55).Text = "Down"
    KeyCodeList(56).KeyCode = vbKeyLeft
    KeyCodeList(56).Text = "Left"
    KeyCodeList(57).KeyCode = vbKeyRight
    KeyCodeList(57).Text = "Right"
    
    
    'KeyCodeList(58).CapitalKeyCode = 13
    KeyCodeList(58).KeyCode = vbKeyReturn
    KeyCodeList(58).Text = "Enter"

   ' KeyCodeList(74).KeyCode = vbKeySeparator
   ' KeyCodeList(74).Text = "Numpad Enter"
    
    KeyCodeList(59).KeyCode = vbKeyTab
    KeyCodeList(59).Text = "Tab"
    
    KeyCodeList(60).KeyCode = 91
    KeyCodeList(60).CapitalKeyCode = 123
    KeyCodeList(60).Text = "["
    KeyCodeList(60).NotChatKey = False
    
    KeyCodeList(61).CapitalKeyCode = 125
    KeyCodeList(61).KeyCode = 93
    KeyCodeList(61).Text = "]"
    KeyCodeList(61).NotChatKey = False
    
    KeyCodeList(62).CapitalKeyCode = 124
    KeyCodeList(62).KeyCode = 92
    KeyCodeList(62).Text = "\"
    KeyCodeList(62).NotChatKey = False
    
    KeyCodeList(63).CapitalKeyCode = 58
    KeyCodeList(63).KeyCode = 59
    KeyCodeList(63).Text = ";"
    KeyCodeList(63).NotChatKey = False
    
    KeyCodeList(64).CapitalKeyCode = 39
    KeyCodeList(64).KeyCode = 34
    KeyCodeList(64).Text = "'"
    KeyCodeList(64).NotChatKey = False
    
    KeyCodeList(65).CapitalKeyCode = 60
    KeyCodeList(65).KeyCode = 44
    KeyCodeList(65).Text = ","
    KeyCodeList(65).NotChatKey = False
    
    KeyCodeList(66).CapitalKeyCode = 62
    KeyCodeList(66).KeyCode = 46
    KeyCodeList(66).Text = "."
    KeyCodeList(66).NotChatKey = False
    
    KeyCodeList(67).CapitalKeyCode = 96
    KeyCodeList(67).KeyCode = 126
    KeyCodeList(67).Text = "~"
    KeyCodeList(67).NotChatKey = False
    
    KeyCodeList(68).KeyCode = vbKeyLButton
    KeyCodeList(68).Text = "Mouse1 (Left)"
    
    KeyCodeList(69).KeyCode = vbKeyRButton
    KeyCodeList(69).Text = "Mouse2 (Right)"
    
    KeyCodeList(70).KeyCode = vbKeyMButton
    KeyCodeList(70).Text = "Mouse3 (Middle)"
    
    
    
    'KeyCodeList(84).KeyCode = v
    'KeyCodeList(84).Text = ""
    'KeyCodeList(85).KeyCode = v
    'KeyCodeList(85).Text = ""
    'KeyCodeList(86).KeyCode = v
    ''KeyCodeList(86).Text = ""
    'KeyCodeList(87).KeyCode = v
    'KeyCodeList(87).Text = ""
    'KeyCodeList(88).KeyCode = v
    'KeyCodeList(88).Text = ""
    'KeyCodeList(89).KeyCode = v
    'KeyCodeList(89).Text = ""
    'KeyCodeList(90).KeyCode = v
    'KeyCodeList(90).Text = ""
    'KeyCodeList(91).KeyCode = v
    'KeyCodeList(91).Text = ""
    'KeyCodeList(92).KeyCode = v
    'KeyCodeList(92).Text = ""
    
    
End Sub

Public Function KeyDown(Key As Byte) As Boolean
    If (Chat.Enabled = False Or KeyCodeList(Key).NotChatKey) And KeyCodeList(Key).KeyCode <> 255 Then
        If (GetKeyState(KeyCodeList(Key).KeyCode) < 0) Then 'Or (KeyCodeList(key).CapitalKeyCode <> 0 And GetKeyState(KeyCodeList(key).CapitalKeyCode) > 0)
            KeyDown = True
        Else
            If Key = 68 Or Key = 69 Or Key = 70 Then
                If GetKeyState(KeyCodeList(Key).KeyCode) And &H8000 Then
                    KeyDown = True
                Else
                    KeyDown = False
                End If
            Else
                KeyDown = False
            End If
        End If
    End If
    
    
    
    
End Function

Function FindKeyCode(KeyCode As Byte) As Long
    Dim A As Long
    For A = 0 To MAXKEYCODES
        If KeyCodeList(A).KeyCode = KeyCode Then
            FindKeyCode = A
            Exit Function
        End If
    Next A
End Function

Sub UseMacro(macronum As Long)
    With Macro(macronum)
        Character.MacroSkill = 0
        If .Skill > 0 Then
            If Skills(.Skill).TargetType And TT_TILE Then
                CastingSpell = True
                Character.MacroSkill = .Skill
            Else
                UseSkill .Skill
            End If
        End If
    End With
End Sub

Sub ResizeGameWindow()
Dim Width As Long, Height As Long, currentRes As Long
    Select Case Options.ResolutionIndex
        Case 1:
            currentRes = 200
            Width = currentRes * 4
            Height = currentRes * 3
        Case 2:
            currentRes = 240
            Width = currentRes * 4
            Height = currentRes * 3
        Case 3:
            currentRes = 256
            Width = currentRes * 4
            Height = currentRes * 3
        Case 4:
            currentRes = 300
            Width = currentRes * 4
            Height = currentRes * 3
        Case 5:
            currentRes = 320
            Width = currentRes * 4
            Height = currentRes * 3
        Case 6:
            currentRes = 350
            Width = currentRes * 4
            Height = currentRes * 3
        Case 7:
            currentRes = 360
            Width = currentRes * 4
            Height = currentRes * 3
        Case 8:
            currentRes = 400
            Width = currentRes * 4
            Height = currentRes * 3
        Case 9:
            Width = 1152
            Height = 864
        Case 10:
            Width = 1280
            Height = 720
        Case 11:
            Width = 1280
            Height = 768
        Case 12:
            Width = 1280
            Height = 800
        Case 13:
            Width = 1280
            Height = 1024
        Case 14:
            Width = 1360
            Height = 768
        Case 15:
            Width = 1440
            Height = 900
        Case 16:
            Width = 1600
            Height = 900
        Case 17:
            Width = 1600
            Height = 1024
        Case 18:
            Width = 1680
            Height = 1050
        Case 19:
            Width = 1768
            Height = 992
        Case 20:
            Width = 1920
            Height = 1080
        Case Else
            Width = 800
            Height = 600
    End Select


    If (Width >= 800 And Height >= 600) Then
       ' frmMain.picViewport.width = 384 + width - 800
       ' frmMain.picViewport.height = 384 + height - 600
        
        frmMain.Width = Width * 12000 / 800
        frmMain.Height = Height * 9000 / 600

        WindowScaleX = Width / 800
        WindowScaleY = Height / 600
        TileSizeX = TileSize * WindowScaleX
        TileSizeY = TileSize * WindowScaleY
        
        frmMain.picViewport.Width = roundUp(ViewportWidth * WindowScaleX)
        frmMain.picViewport.Height = roundUp(ViewportHeight * WindowScaleY)
        frmMain.picViewport.Left = Round(ViewportLeft * WindowScaleX)
        frmMain.picViewport.Top = ViewportTop * WindowScaleY
        
        frmMain.lstSkills.Width = lstSkillsWidth * WindowScaleX
        frmMain.lstSkills.Height = lstSkillsHeight * WindowScaleY
        frmMain.lstSkills.Left = lstSkillsLeft * WindowScaleX
        frmMain.lstSkills.Top = lstSkillsTop * WindowScaleY
        
        frmMain.picChat.Width = picChatWidth * WindowScaleX
        frmMain.picChat.Height = picChatHeight * WindowScaleY
        frmMain.picChat.Left = picChatLeft * WindowScaleX
        frmMain.picChat.Top = picChatTop * WindowScaleY
        
        frmMain.picMiniMap.Width = picMiniMapWidth * WindowScaleX
        frmMain.picMiniMap.Height = picMiniMapHeight * WindowScaleY
        frmMain.picMiniMap.Left = picMiniMapLeft * WindowScaleX
        frmMain.picMiniMap.Top = picMiniMapTop * WindowScaleY
        
        
        
        'INVSrcX = cINVSrcX ' * WindowScaleX
        'InvSrcY = cInvSrcY ' * WindowScaleY
        INVDestX = cINVDestX * WindowScaleX
        INVDestY = cINVDestY * WindowScaleY
        INVWIDTH = cINVWIDTH * WindowScaleX
        INVHEIGHT = cINVHEIGHT * WindowScaleY
        
        'HP
        'HPSrcX = cHPSrcX '* WindowScaleX
        'HPSrcY = cHPSrcY '* WindowScaleY
        HPDestX = cHPDestX * WindowScaleX
        HPDestY = cHPDestY * WindowScaleY
        HPWIDTH = cHPWIDTH * WindowScaleX
        HPHEIGHT = cHPHEIGHT * WindowScaleY
        
        'Energy
        'ENERGYSrcX = cENERGYSrcX '* WindowScaleX
        'ENERGYSrcY = cENERGYSrcY '* WindowScaleY
        ENERGYDestX = cENERGYDestX * WindowScaleX
        ENERGYDestY = cENERGYDestY * WindowScaleY
        ENERGYWIDTH = cENERGYWIDTH * WindowScaleX
        ENERGYHEIGHT = cENERGYHEIGHT * WindowScaleY
        
        'Mana
        'MANASrcX = cMANASrcX '* WindowScaleX
        'MANASrcY = cMANASrcY '* WindowScaleY
        MANADestX = cMANADestX * WindowScaleX
        MANADestY = cMANADestY * WindowScaleY
        MANAWIDTH = cMANAWIDTH * WindowScaleX
        MANAHEIGHT = cMANAHEIGHT * WindowScaleY
        
        'EXP Bar
        'EXPSrcX = cEXPSrcX '* WindowScaleX
        'EXPSrcY = cEXPSrcY '* WindowScaleY
        EXPDestX = cEXPDestX * WindowScaleX
        EXPDestY = cEXPDestY * WindowScaleY
        EXPWIDTH = cEXPWIDTH * WindowScaleX
        EXPHEIGHT = cEXPHEIGHT * WindowScaleY
        
        
        
       ' If sfcInventory2.desc.lWidth <> 0 Then InitializeSurfaces
        
        frmMain.Top = 0
        frmMain.Left = 0
        
        
        frmMain.Cls
        frmMain.PaintPicture frmMain.Picture, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight
        
        If sfcInventory2.desc.lWidth <> 0 Then 'dx is init
            DrawHP
            DrawEnergy
            DrawMana
            DrawExperience
            DrawInv
            DrawChat
            DrawMapTitle map.Name
            SetTab CurrentTab
        End If
        
        
    End If
End Sub

Sub FullscreenGameWindow()

 Dim scry As Long, scrx As Long
            Dim twipsPerpixelX As Long
            Dim twipsPerpixelY As Long
            Dim resX As Long
            Dim resY As Long
        
            scrx = Screen.Width
            scry = Screen.Height
        
            twipsPerpixelX = Screen.twipsPerpixelX
            twipsPerpixelY = Screen.twipsPerpixelY
        
            resX = scrx / twipsPerpixelX
            resY = scry / twipsPerpixelY


Dim Width As Long, Height As Long
Width = resX
Height = resY


    If (Width >= 800 And Height >= 600) Then
       ' frmMain.picViewport.width = 384 + width - 800
       ' frmMain.picViewport.height = 384 + height - 600
        
        frmMain.Width = Width * 12000 / 800
        frmMain.Height = Height * 9000 / 600

        WindowScaleX = Width / 800
        WindowScaleY = Height / 600
        TileSizeX = TileSize * WindowScaleX
        TileSizeY = TileSize * WindowScaleY
        
        frmMain.picViewport.Width = roundUp(ViewportWidth * WindowScaleX)
        frmMain.picViewport.Height = roundUp(ViewportHeight * WindowScaleY)
        frmMain.picViewport.Left = ViewportLeft * WindowScaleX
        frmMain.picViewport.Top = ViewportTop * WindowScaleY
        
        frmMain.lstSkills.Width = lstSkillsWidth * WindowScaleX
        frmMain.lstSkills.Height = lstSkillsHeight * WindowScaleY
        frmMain.lstSkills.Left = lstSkillsLeft * WindowScaleX
        frmMain.lstSkills.Top = lstSkillsTop * WindowScaleY
        
        frmMain.picChat.Width = picChatWidth * WindowScaleX
        frmMain.picChat.Height = picChatHeight * WindowScaleY
        frmMain.picChat.Left = picChatLeft * WindowScaleX
        frmMain.picChat.Top = picChatTop * WindowScaleY
        
        frmMain.picMiniMap.Width = picMiniMapWidth * WindowScaleX
        frmMain.picMiniMap.Height = picMiniMapHeight * WindowScaleY
        frmMain.picMiniMap.Left = picMiniMapLeft * WindowScaleX
        frmMain.picMiniMap.Top = picMiniMapTop * WindowScaleY
        
        
        
        'INVSrcX = cINVSrcX ' * WindowScaleX
        'InvSrcY = cInvSrcY ' * WindowScaleY
        INVDestX = cINVDestX * WindowScaleX
        INVDestY = cINVDestY * WindowScaleY
        INVWIDTH = cINVWIDTH * WindowScaleX
        INVHEIGHT = cINVHEIGHT * WindowScaleY
        
        'HP
        'HPSrcX = cHPSrcX '* WindowScaleX
        'HPSrcY = cHPSrcY '* WindowScaleY
        HPDestX = cHPDestX * WindowScaleX
        HPDestY = cHPDestY * WindowScaleY
        HPWIDTH = cHPWIDTH * WindowScaleX
        HPHEIGHT = cHPHEIGHT * WindowScaleY
        
        'Energy
        'ENERGYSrcX = cENERGYSrcX '* WindowScaleX
        'ENERGYSrcY = cENERGYSrcY '* WindowScaleY
        ENERGYDestX = cENERGYDestX * WindowScaleX
        ENERGYDestY = cENERGYDestY * WindowScaleY
        ENERGYWIDTH = cENERGYWIDTH * WindowScaleX
        ENERGYHEIGHT = cENERGYHEIGHT * WindowScaleY
        
        'Mana
        'MANASrcX = cMANASrcX '* WindowScaleX
        'MANASrcY = cMANASrcY '* WindowScaleY
        MANADestX = cMANADestX * WindowScaleX
        MANADestY = cMANADestY * WindowScaleY
        MANAWIDTH = cMANAWIDTH * WindowScaleX
        MANAHEIGHT = cMANAHEIGHT * WindowScaleY
        
        'EXP Bar
        'EXPSrcX = cEXPSrcX '* WindowScaleX
        'EXPSrcY = cEXPSrcY '* WindowScaleY
        EXPDestX = cEXPDestX * WindowScaleX
        EXPDestY = cEXPDestY * WindowScaleY
        EXPWIDTH = cEXPWIDTH * WindowScaleX
        EXPHEIGHT = cEXPHEIGHT * WindowScaleY
        
        
        
       ' If sfcInventory2.desc.lWidth <> 0 Then InitializeSurfaces
        
        frmMain.Top = 0
        frmMain.Left = 0
        
        
        frmMain.Cls
        frmMain.PaintPicture frmMain.Picture, 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight
        
        If sfcInventory2.desc.lWidth <> 0 Then 'dx is init
            DrawHP
            DrawEnergy
            DrawMana
            DrawExperience
            DrawInv
            DrawChat
            DrawMapTitle map.Name
            SetTab CurrentTab
        End If
        
        
    End If
End Sub


Public Sub SetServerData()
    Dim Numbers As String
    If GetTickCount < 0 Then
        MsgBox "There has been an error (gettickcount)", vbOKOnly, TitleString
        End
    End If
    Numbers = "1234567890."
    
    ServerPort = Chr$(51) + Chr$(48) + Chr$(48) + Chr$(56)
    ServerPort = ServerPort + 9
    
    ServerIP = "samuelw.co"
        
    Dim Parts() As String
    Parts = Split(Trim$(Command$), " ")
    
    Dim i As Integer
    For i = 0 To UBound(Parts)
        If LCase(Parts(i)) = "-ip" Then
            ServerIP = Parts(i + 1)
        End If
    
        If LCase(Parts(i)) = "-serverid" Then
            ServerId = Parts(i + 1)
        End If
        
        If LCase(Parts(i)) = "-port" Then
            ServerPort = Parts(i + 1)
        End If
        
        If LCase(Parts(i)) = "-cclasses" Then
            ServerHasCustomClasses = True
        End If
        
        If LCase(Parts(i)) = "-cskilldata" Then
            ServerHasCustomSkilldata = True
        End If
        
        If LCase(Parts(i)) = "-ver" Then
             Dim result As Long
             WriteFile GetStdHandle(STD_OUTPUT_HANDLE), Str(ClientVer), Len(Str(ClientVer)), result, ByVal 0&
             End
        End If
    Next
End Sub


Public Function roundUp(dblValue As Double) As Double
On Error GoTo PROC_ERR
Dim myDec As Long

myDec = InStr(1, CStr(dblValue), ".", vbTextCompare)
If myDec > 0 Then
    roundUp = CDbl(Left(CStr(dblValue), myDec)) + 1
Else
    roundUp = dblValue
End If

PROC_EXIT:
    Exit Function
PROC_ERR:
    MsgBox Err.Description, vbInformation, "Round Up"
End Function

Public Sub InitConstants()
    Dim A As Long
    Dim ServerPart As String
    
    If (ServerHasCustomClasses) Then ServerPart = ServerId
    
    SkillsPerLevel = ReadInt("Settings", "SkillsPerLevel", "classes", ServerHasCustomClasses)
    StatsPerLevel = ReadInt("Settings", "StatsPerLevel", "classes", ServerHasCustomClasses)
    StatRate1 = ReadInt("Settings", "StatRate1", "classes", ServerHasCustomClasses)
    StatRate2 = ReadInt("Settings", "StatRate2", "classes", ServerHasCustomClasses)
    
    baseEnergyRegen = ReadInt("Settings", "BaseEnergyRegen", "classes", ServerHasCustomClasses)
    BaseHPRegen = ReadInt("Settings", "baseHpRegen", "classes", ServerHasCustomClasses)
    baseManaRegen = ReadInt("Settings", "baseManaRegen", "classes", ServerHasCustomClasses)
    
    LevelsPerHpRegen = ReadInt("Settings", "LevelsPerHpRegen", "classes", ServerHasCustomClasses)
    LevelsPerManaRegen = ReadInt("Settings", "LevelsPerManaRegen", "classes", ServerHasCustomClasses)
    
    
    For A = 1 To 3
        StrengthPerDamage(A) = ReadInt("Settings", "StrengthPerDamage" & A, "classes", ServerHasCustomClasses)
        
        AgilityPerCritChance(A) = ReadInt("Settings", "AgilityPerCritChance" & A, "classes", ServerHasCustomClasses)
        AgilityPerDodgeChance(A) = ReadInt("Settings", "AgilityPerDodgeChance" & A, "classes", ServerHasCustomClasses)
        
        EndurancePerBlockChance(A) = ReadInt("Settings", "EndurancePerBlockChance" & A, "classes", ServerHasCustomClasses)
        EndurancePerEnergy(A) = ReadInt("Settings", "EndurancePerEnergy" & A, "classes", ServerHasCustomClasses)
        EndurancePerEnergyRegen(A) = ReadInt("Settings", "EndurancePerEnergyRegen" & A, "classes", ServerHasCustomClasses)
        
        ManaPerIntelligence(A) = ReadInt("Settings", "ManaPerIntelligence" & A, "classes", ServerHasCustomClasses)
        IntelligencePerManaRegen(A) = ReadInt("Settings", "IntelligencePerManaRegen" & A, "classes", ServerHasCustomClasses)
        
        ConstitutionPerHPRegen(A) = ReadInt("Settings", "ConstitutionPerHPRegen" & A, "classes", ServerHasCustomClasses)
        HPPerConstitution(A) = ReadInt("Settings", "HPPerConstitution" & A, "classes", ServerHasCustomClasses)
        
        PietyPerMagicResist(A) = ReadInt("Settings", "PietyPerMagicResist" & A, "classes", ServerHasCustomClasses)
        PietyPerHPRegen(A) = ReadInt("Settings", "PietyPerHPRegen" & A, "classes", ServerHasCustomClasses)
        PietyPerManaRegen(A) = ReadInt("Settings", "PietyPerManaRegen" & A, "classes", ServerHasCustomClasses)
        PietyPerHP(A) = ReadInt("Settings", "PietyPerHP" & A, "classes", ServerHasCustomClasses)
        PietyPerMana(A) = ReadInt("Settings", "PietyPerMana" & A, "classes", ServerHasCustomClasses)
        PietyPerDodge(A) = ReadInt("Settings", "PietyPerDodge" & A, "classes", ServerHasCustomClasses)
        PietyPerCrit(A) = ReadInt("Settings", "PietyPerCrit" & A, "classes", ServerHasCustomClasses)
        PietyPerBlock(A) = ReadInt("Settings", "PietyPerBlock" & A, "classes", ServerHasCustomClasses)
        
        GenericStatPerBonus(A) = ReadInt("Settings", "GenericStatPerBonus" & A, "classes", ServerHasCustomClasses)
        GenericPietyPerBonus(A) = ReadInt("Settings", "GenericPietyPerBonus" & A, "classes", ServerHasCustomClasses)
    Next A
    
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

Function GetGenericStatBonus(ByVal statValue As Long, Optional ByVal pietyValue As Long = 0) As Long
    GetGenericStatBonus = GetStatPerBonus(statValue, GenericStatPerBonus) + GetStatPerBonus(statValue, GenericPietyPerBonus)
End Function

Function SHA256(Value As String)
    SHA256 = oSHA.SHA256FromString(Value)


End Function



