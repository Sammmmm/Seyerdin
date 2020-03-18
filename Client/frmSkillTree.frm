VERSION 5.00
Begin VB.Form frmSkillTree 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   567
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picSkillTree 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   840
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   0
      Top             =   1080
      Width           =   3735
   End
End
Attribute VB_Name = "frmSkillTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Private texSkills As DirectDrawSurface7
Private texDescs As DDSURFACEDESC2
Private texSkillBG As DirectDrawSurface7
Private texBGDesc As DDSURFACEDESC2
Const MAX_TREES = 20
Private treeGroups(1 To MAX_TREES) As Long
Private skillTreeInfo(0 To MAX_SKILLS) As SkillTreeData
Private skillTraining(0 To MAX_SKILLS) As Byte
Private skillCount As Byte
Private skillPoints As Byte
Private hoverSkill As Long
Private mouseDown As Boolean
Private mX As Long, mY As Long
Private mdX As Long, mdY As Long
Const BorderSize = 1000
Const TreeXSpacing = 100 '80
Const TreeYSpacing = 75
Const TreeYOffset = 16
Const TreeXOffset = 34 '64


Private buffer As Long




Private Sub Form_Activate()
Set texSkillBG = DD.CreateSurfaceFromFile("Data/Graphics/Interface/SkillTree.rsc", texBGDesc)
Set texSkills = DD.CreateSurfaceFromFile("Data/Graphics/skillthumbs.rsc", texDescs)
Dim A As Long
For A = 0 To MAX_SKILLS
    skillTraining(A) = 0
Next A
picSkillTree.ForeColor = vbWhite
redoSkillTreeInfo
drawSkills
mouseDown = False
End Sub

Private Sub picskilltree_KeyDown(KeyCode As Integer, Shift As Integer)
    If hoverSkill > 0 Then
   ' If Character.SkillLevels(skillTreeInfo(hoverSkill).Skill) > 0 Then
    Dim b As Long, A As Long
    For b = 0 To 9
        If KeyDown(Options.SpellKey(b)) Then
            If skillTreeInfo(hoverSkill).Skill > 0 Then
                If skillTreeInfo(hoverSkill).Skill <= MAX_SKILLS Then
                    For A = 0 To 9
                        If Macro(A).Skill = skillTreeInfo(hoverSkill).Skill Then
                            Macro(A).Skill = 0
                        End If
                    Next A
                    Macro(b).Skill = skillTreeInfo(hoverSkill).Skill
                    frmMain.DrawLstBox
                    SaveSkillMacros
                    drawSkills
                End If
                Exit For
            End If
        End If
    Next b
    End If
  '  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim A As Long
    For A = 0 To MAX_SKILLS
        skillTraining(A) = 0
        skillTreeInfo(A).Skill = 0
    Next A
    For A = 1 To MAX_TREES
        treeGroups(A) = 0
    Next A
    skillCount = 0
    Set texSkills = Nothing
    Set texSkillBG = Nothing
    Unload Me
End Sub


Sub Form_Resize()
    picSkillTree.Left = 0
    picSkillTree.Top = 0
    picSkillTree.Width = Me.ScaleWidth
    picSkillTree.Height = Me.ScaleHeight
End Sub

Sub redoSkillTreeInfo()
Dim A As Long
Dim b As Long
skillCount = 1

    skillPoints = Character.skillPoints
    For b = 1 To MAX_TREES
        treeGroups(b) = 0
    Next b
    For A = 0 To MAX_SKILLS
        skillTreeInfo(A).Skill = 0
        skillTreeInfo(A).Group = 0
        skillTreeInfo(A).meetsReqs = False
        skillTreeInfo(A).Req1 = 0
        skillTreeInfo(A).Req2 = 0
        skillTreeInfo(A).LevelReq = 0
        skillTreeInfo(A).Req1Level = 0
        skillTreeInfo(A).Req2Level = 0
        skillTreeInfo(A).used = False
        skillTreeInfo(A).drawX = 0
        skillTreeInfo(A).drawY = 0
       ' skillTraining(A) = 0
    Next A

    For A = 0 To MAX_SKILLS
            
            If (2 ^ (Character.Class - 1) And Skills(A).Class) And Skills(A).Name <> "" Then 'this is a class skill
                skillTreeInfo(skillCount).Skill = A
                skillTreeInfo(skillCount).meetsReqs = True
                If Skills(A).Level(Character.Class) > Character.Level Then
                    skillTreeInfo(skillCount).meetsReqs = False
                    skillTreeInfo(skillCount).LevelReq = Skills(A).Level(Character.Class)
                Else
                    skillTreeInfo(skillCount).LevelReq = 0
                End If
                
                For b = 1 To MAX_TREES
                    skillTreeInfo(skillCount).Group = b
                    If treeGroups(b) = Skills(A).Color Then Exit For
                    If treeGroups(b) = 0 Then
                        treeGroups(b) = Skills(A).Color
                        Exit For
                    End If
                Next b
                With Skills(A).Requirements(0)
                    If .Skill > SKILL_INVALID And .Skill < MAX_SKILLS Then
                        skillTreeInfo(skillCount).Req1 = .Skill
                        skillTreeInfo(skillCount).Req1Level = .Level
                        If Character.SkillLevels(.Skill) + skillTraining(.Skill) < .Level Then
                            skillTreeInfo(skillCount).meetsReqs = False
                        End If
                    Else
                        skillTreeInfo(skillCount).Req1 = -1
                    End If
                End With
                With Skills(A).Requirements(1)
                    If .Skill > SKILL_INVALID And .Skill < MAX_SKILLS Then
                        skillTreeInfo(skillCount).Req2 = .Skill
                        skillTreeInfo(skillCount).Req2Level = .Level
                        If Character.SkillLevels(.Skill) + skillTraining(.Skill) < .Level Then
                            skillTreeInfo(skillCount).meetsReqs = False
                        End If
                    Else
                        skillTreeInfo(skillCount).Req2 = -1
                    End If
                End With
                If Not skillTreeInfo(skillCount).meetsReqs And skillTraining(A) > 0 Then
                    skillTraining(A) = 0
                    redoSkillTreeInfo
                    Exit Sub
                End If
                skillPoints = skillPoints - skillTraining(A)
                skillCount = skillCount + 1
            End If
            

    Next A

    
    For A = 1 To MAX_TREES
        If treeGroups(A) = 0 Then
            Exit For
        End If
    Next A
    A = A - 1

    If A < 5 Then buffer = 5 - A

    If A > 5 And A < 7 Then
        Me.Width = A * 1150 + BorderSize * 2 + 400
    ElseIf A > 7 Then
        Me.Width = A * 1150 + BorderSize * 2 + 1300 * 2
    Else
        Me.Width = A * 1150 + BorderSize * 2 + 1150 * buffer
    End If
    Me.Height = findMaxHeight * 1150
    
End Sub

Function findMaxHeight() As Long
    Dim y As Long, max As Long, A As Long, b As Long
   

    For A = 1 To MAX_TREES
        If max < y Then max = y
        y = 0
        For b = 0 To MAX_SKILLS
            If skillTreeInfo(b).Group = A Then y = y + 1
        Next b
    Next A
    If max < y Then max = y
    findMaxHeight = max
End Function

Sub drawSkills()
    Dim A As Long, b As Long, cur As Long, y As Long, C As Long
    drawBG
    
    For A = 0 To MAX_SKILLS
        skillTreeInfo(A).used = False
    Next A
    
    For A = 1 To MAX_TREES
        y = 0
        If treeGroups(A) <> 0 Then
nextB:
            cur = -1
            For b = 0 To MAX_SKILLS
                If skillTreeInfo(b).Group = A And Not skillTreeInfo(b).used Then
                    If cur = -1 Then
                        cur = b
                    ElseIf skillTreeInfo(cur).Req1 = skillTreeInfo(b).Skill Or skillTreeInfo(cur).Req2 = skillTreeInfo(b).Skill Then
                        cur = b
                    ElseIf skillTreeInfo(cur).Req1 = -1 And skillTreeInfo(cur).Req2 = -1 And skillTreeInfo(b).Req1 = -1 And skillTreeInfo(b).Req2 = -1 And Skills(skillTreeInfo(cur).Skill).Level(Character.Class) > Skills(skillTreeInfo(b).Skill).Level(Character.Class) Then
                        cur = b
                    ElseIf skillTreeInfo(cur).meetsReqs = False And skillTreeInfo(b).meetsReqs = True Then
                        cur = b
                    End If
                    For C = 0 To MAX_SKILLS
                        If skillTreeInfo(C).Group = A And skillTreeInfo(C).used Then
                            If skillTreeInfo(b).Req1 = skillTreeInfo(C).Skill And skillTreeInfo(b).Req2 = -1 Then
                                cur = b
                                Exit For
                            End If
                        End If
                    Next C
                End If
            Next b
                If cur <> -1 Then
                    skillTreeInfo(cur).used = True
                    drawSkill (A - 1) * TreeXSpacing + TreeXOffset + (buffer * TreeXSpacing) / 2, TreeYOffset + y * TreeYSpacing, cur, skillTreeInfo(cur).meetsReqs
                    
                    
                    If y > 0 Then
                        picSkillTree.Line ((A - 1) * TreeXSpacing + TreeXOffset + 15 + (buffer * TreeXSpacing) / 2, TreeYOffset + 32 + (y - 1) * TreeYSpacing)-((A - 1) * TreeXSpacing + TreeXOffset + 16 + (buffer * TreeXSpacing) / 2, 16 + (y) * TreeYSpacing - 1), Skills(skillTreeInfo(cur).Skill).Color, BF
                    End If
                    
                    For C = 0 To 9
                        If Macro(C).Skill = skillTreeInfo(cur).Skill Then
                            TextOut picSkillTree.hdc, (A - 1) * TreeXSpacing + TreeXOffset + 34, TreeYOffset + y * TreeYSpacing + 17, "[" + KeyCodeList(Options.SpellKey(C)).Text + "]", Len("[" + KeyCodeList(Options.SpellKey(C)).Text + "]")
                        End If
                    Next C
                    
                    y = y + 1
                    GoTo nextB
                End If
        End If
    Next A
            picSkillTree.FontBold = True
    picSkillTree.FontSize = 12
    picSkillTree.ForeColor = vbWhite
    TextOut picSkillTree.hdc, 5, Me.ScaleHeight - 24, "Free Skill Points: " + Str(skillPoints), Len("Free Skill Points: " + Str(skillPoints))
    
    If hoverSkill > 0 Then
        Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long, textY As Long
        With (skillTreeInfo(hoverSkill))
            If .Group <= (5 - buffer) - 2 Then
                x1 = .drawX + 32
                x2 = x1 + 210
            Else
                x2 = .drawX - 1
                x1 = x2 - 210
            End If
            If .drawY > Me.ScaleHeight / 2 Then
                y2 = .drawY + 31
                y1 = y2 - 180
            Else
                y1 = .drawY
                y2 = y1 + 180
            End If
            
            'x1 = Me.Width / Me.ScaleWidth
                        
                        
            picSkillTree.Line (x1, y1)-(x2, y2), vbBlack, BF
            picSkillTree.Line (x1, y1)-(x2, y2), Skills(.Skill).Color, B
            
            picSkillTree.ForeColor = Skills(.Skill).Color
            picSkillTree.Line (x1 + 5, y1 + 22)-(x2 - 5, y1 + 23), Skills(.Skill).Color, BF
            picSkillTree.FontSize = 12
            picSkillTree.FontBold = True
            TextOut picSkillTree.hdc, x1 + 5, y1 + 5, Skills(.Skill).Name, Len(Skills(.Skill).Name)

            If .meetsReqs Then
                picSkillTree.FontSize = 10
                picSkillTree.FontBold = False
                If skillTraining(.Skill) > 0 Then
                    TextOut picSkillTree.hdc, x1 + 7, y1 + 25, "Skill Level: " + Str(Character.SkillLevels(.Skill)) + " + " + Str(skillTraining(.Skill)) + "/" + Str(Skills(.Skill).MaxLevel), Len("Skill Level: " + Str(Character.SkillLevels(.Skill)) + " + " + Str(skillTraining(.Skill)) + "/" + Str(Skills(.Skill).MaxLevel))
                Else
                    TextOut picSkillTree.hdc, x1 + 7, y1 + 25, "Skill Level: " + Str(Character.SkillLevels(.Skill)) + "/" + Str(Skills(.Skill).MaxLevel), Len("Skill Level: " + Str(Character.SkillLevels(.Skill)) + "/" + Str(Skills(.Skill).MaxLevel))
                End If
                textY = y1 + 25 + 16
                picSkillTree.Line (x1 + 6, textY)-(x2 - 6, textY), Skills(.Skill).Color, BF
                textY = textY + 4
            Else
                picSkillTree.FontSize = 8
                picSkillTree.FontBold = False
                textY = y1 + 25
                If .LevelReq > 0 Then
                    TextOut picSkillTree.hdc, x1 + 7, textY, "Requires Character Level " + Str(.LevelReq), Len("Requires Character Level " + Str(.LevelReq))
                    textY = textY + 12
                End If
                
                If .Req1 >= 0 Then
                    TextOut picSkillTree.hdc, x1 + 7, textY, "Requires " + Str(.Req1Level) + " skill points in " + Skills(.Req1).Name, Len("Requires " + Str(.Req1Level) + " skill points in " + Skills(.Req1).Name)
                    textY = textY + 12
                End If
                If .Req2 >= 0 Then
                    TextOut picSkillTree.hdc, x1 + 7, textY, "Requires " + Str(.Req2Level) + " skill points in " + Skills(.Req2).Name, Len("Requires " + Str(.Req2Level) + " skill points in " + Skills(.Req2).Name)
                    textY = textY + 12
                End If
                textY = textY + 2
                picSkillTree.Line (x1 + 6, textY)-(x2 - 6, textY), Skills(.Skill).Color, BF
                textY = textY + 4
                picSkillTree.FontSize = 10
            End If
            Dim St As String
            Dim R1 As RECT
                St = Skills(.Skill).Description
                R1.Left = x1 + 7
                R1.Right = x2 - 3
                R1.Top = textY
                R1.Bottom = y2
                DrawText picSkillTree.hdc, St, Len(St), R1, DT_WORDBREAK
            
            
            
                'st = "Mana Cost: "
                'If Skills(.Skill).CostConstant > 0 Then st = st + " + " + Str(Skills(.Skill).CostConstant)
                'If Skills(.Skill).CostPerLevel > 0 Then st = st + " + " + Str(Skills(.Skill).CostPerLevel) + " * cLvl"
                'If Skills(.Skill).CostPerSLevel > 0 Then st = st + " + " + Str(Skills(.Skill).CostPerSLevel) + " * sLvl"
                'TextOut picSkillTree.hdc, x1 + 7, textY, st, Len(st)

            
        End With
    End If
    picSkillTree.FontBold = True
    picSkillTree.FontSize = 12
    picSkillTree.ForeColor = vbWhite
    'TextOut picSkillTree.hdc, 5, Me.ScaleHeight - 24, "Free Skill Points: " + Str(skillPoints), Len("Free Skill Points: " + Str(skillPoints))
    
    
    picSkillTree.Refresh
End Sub


Sub drawSkill(x As Long, y As Long, skillC As Long, meetsReqs As Boolean)
    Dim hdc As Long
    With skillTreeInfo(skillC)
    .drawX = x
    .drawY = y
    
    hdc = texSkillBG.GetDC
        BitBlt picSkillTree.hdc, x, y, 32, 32, hdc, 32, Abs(Not meetsReqs) * 32, SRCCOPY
    texSkillBG.ReleaseDC hdc
    
    hdc = texSkills.GetDC
        TransparentBlt picSkillTree.hdc, x + 0, y + 0, 32, 32, hdc, 0, 32 * (.Skill - 1), SRCCOPY
    texSkills.ReleaseDC hdc
        
    picSkillTree.FontSize = 10
    picSkillTree.FontBold = True
    picSkillTree.ForeColor = Skills(.Skill).Color
    If .meetsReqs Then
        If skillTraining(.Skill) > 0 Then
            TextOut picSkillTree.hdc, x + 30, y + 1, Str(Character.SkillLevels(.Skill)) + "+" + Str(skillTraining(.Skill)) + "/" + Str(Skills(.Skill).MaxLevel), Len(Str(Character.SkillLevels(.Skill)) + "+" + Str(skillTraining(.Skill)) + "/" + Str(Skills(.Skill).MaxLevel))
        Else
            TextOut picSkillTree.hdc, x + 30, y + 1, Str(Character.SkillLevels(.Skill)) + "/" + Str(Skills(.Skill).MaxLevel), Len(Str(Character.SkillLevels(.Skill)) + "/" + Str(Skills(.Skill).MaxLevel))
        End If
    End If
    End With
    
    
End Sub

Sub drawBG()
Dim x As Long, y As Long, hdc As Long
   hdc = texSkillBG.GetDC
   
For x = 0 To Me.ScaleWidth + 32 Step 32
    For y = 0 To Me.ScaleHeight + 20 Step 20
        BitBlt picSkillTree.hdc, x, y + 8, 32, 32, hdc, 9, 73, SRCCOPY
    Next y
Next x

For x = 0 To Me.ScaleWidth + 32 Step 32
        BitBlt picSkillTree.hdc, x, 0, 32, 8, hdc, 9, 64, SRCCOPY
Next x

For x = 0 To Me.ScaleWidth + 32 Step 32
        BitBlt picSkillTree.hdc, x, Me.ScaleHeight - 8, 32, 8, hdc, 9, 106, SRCCOPY
Next x

For y = 0 To Me.ScaleHeight + 32 Step 32
        BitBlt picSkillTree.hdc, 0, y, 8, 32, hdc, 0, 73, SRCCOPY
Next y
For y = 0 To Me.ScaleHeight + 32 Step 32
        BitBlt picSkillTree.hdc, Me.ScaleWidth - 8, y, 8, 32, hdc, 42, 73, SRCCOPY
Next y

BitBlt picSkillTree.hdc, 0, 0, 8, 8, hdc, 0, 64, SRCCOPY
BitBlt picSkillTree.hdc, 0, Me.ScaleHeight - 8, 8, 8, hdc, 0, 106, SRCCOPY
BitBlt picSkillTree.hdc, Me.ScaleWidth - 8, 0, 8, 8, hdc, 42, 64, SRCCOPY
BitBlt picSkillTree.hdc, Me.ScaleWidth - 8, Me.ScaleHeight - 8, 8, 8, hdc, 42, 106, SRCCOPY


    BitBlt picSkillTree.hdc, Me.ScaleWidth - 58 - 6, Me.ScaleHeight - 22 - 6, 58, 22, hdc, 70, 0, SRCCOPY
    If (Character.skillPoints > skillPoints) Then
        BitBlt picSkillTree.hdc, Me.ScaleWidth - 58 - 6 - 58 - 6, Me.ScaleHeight - 22 - 6, 58, 22, hdc, 70, 22, SRCCOPY
    End If
    texSkillBG.ReleaseDC hdc
End Sub

Private Sub picSkillTree_mousedown(Button As Integer, Shift As Integer, x As Single, y As Single)

    mdX = x
    mdY = y
    If x >= Me.ScaleWidth - 58 - 6 And x <= Me.ScaleWidth - 6 And y >= Me.ScaleHeight - 22 - 6 And y <= Me.ScaleHeight - 6 Then 'close
        Set texSkills = Nothing
        Set texSkillBG = Nothing
        Unload Me
    End If
        mouseDown = True
    If x >= Me.ScaleWidth - 58 - 58 - 6 - 6 And x <= Me.ScaleWidth - 6 - 58 - 6 And y >= Me.ScaleHeight - 22 - 6 And y <= Me.ScaleHeight - 6 And Character.skillPoints > skillPoints Then 'apply changes
        Dim A As Long
        For A = 0 To MAX_SKILLS
            While (skillTraining(A) > 0)
                Character.skillPoints = Character.skillPoints - 1
                Character.SkillLevels(A) = Character.SkillLevels(A) + 1
                SendSocket Chr$(30) + Chr$(1) + Chr$(A)
                skillTraining(A) = skillTraining(A) - 1
            Wend
        Next A
        UpdateSkills
        redoSkillTreeInfo
        drawSkills
    End If
    If hoverSkill >= 0 Then
        If Button = 1 Then
            If skillPoints > 0 Then
                If skillTreeInfo(hoverSkill).meetsReqs Then
                    If skillTraining(skillTreeInfo(hoverSkill).Skill) + Character.SkillLevels(skillTreeInfo(hoverSkill).Skill) < Skills(skillTreeInfo(hoverSkill).Skill).MaxLevel Then
                        skillTraining(skillTreeInfo(hoverSkill).Skill) = skillTraining(skillTreeInfo(hoverSkill).Skill) + 1
                        redoSkillTreeInfo
                        drawSkills
                    End If
                End If
            End If
        Else
            If skillTreeInfo(hoverSkill).meetsReqs Then
                If skillTraining(skillTreeInfo(hoverSkill).Skill) > 0 Then
                    skillTraining(skillTreeInfo(hoverSkill).Skill) = skillTraining(skillTreeInfo(hoverSkill).Skill) - 1
                    redoSkillTreeInfo
                    drawSkills
                End If
            End If
        End If
    End If
    
End Sub

Private Sub picskilltree_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim A As Long
mX = x
mY = y
If mouseDown Then
Me.Left = Me.Left - (mdX - mX)
mX = mdX
Me.Top = Me.Top - (mdY - mY)
mY = mdY
End If
For A = 0 To MAX_SKILLS
    If skillTreeInfo(A).drawX <= x And skillTreeInfo(A).drawX + 31 >= x And skillTreeInfo(A).drawY <= y And skillTreeInfo(A).drawY + 31 >= y Then
        If hoverSkill <> A Then
            hoverSkill = A
            drawSkills
        End If
        Exit Sub
    End If
Next A
If hoverSkill <> -1 Then
    hoverSkill = -1
    drawSkills
End If


End Sub

Private Sub picSkillTree_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mouseDown = False
End Sub
