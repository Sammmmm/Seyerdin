Attribute VB_Name = "ModSkills"
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

Type SkillRequirementType
    Skill As Byte
    Level As Byte
End Type

'Skill information = 316 bytes total
Type SkillType
    Name As String              'max of 32
    Description As String       'max 256
    Class As Integer
    Level(1 To 10) As Byte
    TargetType As Byte
    Type As Long
    Flags As Long
    Range As Byte
    MaxLevel As Byte
    EXPTable As Byte
    Requirements(1) As SkillRequirementType
    Icon As Byte
    ManaCost1(1 To 4) As Byte
    ManaCost(2 To 3) As Integer
    LocalTick As Long
    GlobalTick As Long
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Public Skills() As SkillType
Public CurrentSkill As Long

Function DoubleChar(ByVal Num As Long) As String
    DoubleChar = Chr$(Int(Num / 256)) + Chr$(Num Mod 256)
End Function

Function GetInt(Chars As String) As Long
    GetInt = Asc(Mid$(Chars, 1, 1)) * 256 + Asc(Mid$(Chars, 2, 1))
End Function

Function QuadChar(ByVal Num As Long) As String
    QuadChar = Chr$(Int(Num / 16777216) Mod 256) + Chr$(Int(Num / 65536) Mod 256) + Chr$(Int(Num / 256) Mod 256) + Chr$(Num Mod 256)
End Function

Function ClipString(St As String) As String
    Dim A As Long
    For A = Len(St) To 1 Step -1
        If Mid$(St, A, 1) <> Chr$(32) Then
            ClipString = Mid$(St, 1, A)
            Exit Function
        End If
    Next A
End Function
