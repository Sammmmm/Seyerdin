VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParticleEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

'local variable to hold collection
Option Explicit
Private mCol As Collection
Public Current As String
Public LastKey As Long

Public Function Add(x As Long, y As Long, SourceType As Long, Red As Long, Green As Long, Blue As Long, Life As Long, VelocityX As Single, VelocityY As Single, NumParticles As Long, ParticleSize As Long, ParticleTexNum As Long, Size1 As Long, TargType As Long, Target As Long) As clsParticleSource
    'create a new object
    Dim objNewMember As clsParticleSource
    Set objNewMember = New clsParticleSource
    Dim b As Single
    
    With objNewMember
        .x = x
        .y = y
        .Red = Red
        .Green = Green
        .Blue = Blue
        .NumParticles = NumParticles
        If SourceType = 9 Then .NumParticles = (((x * 32 - VelocityX * 32) ^ 2 + (y * 32 - VelocityY * 32) ^ 2) ^ (0.5)) / 300
        .SourceType = SourceType
        .Life = Life
        .VelocityX = VelocityX
        .VelocityY = VelocityY
        If ParticleSize > 25 Then ParticleSize = 25
        b = ParticleSize
        .lng0 = FtoDW(0)
        .lng1 = FtoDW(1)
        .ParticleSize = FtoDW(b)
        .ParticleNum = ParticleTexNum
        .Size1 = Size1
        .TargType = TargType
        .Target = Target
        LastKey = LastKey + 1
        If LastKey > 2140000000 Then LastKey = 0
        .Key = "Source" + CStr(LastKey)
        

        
    End With

    mCol.Add objNewMember, objNewMember.Key

    'return the object created
    Set Add = objNewMember
    objNewMember.Init
    Set objNewMember = Nothing
End Function


Public Function Clear()
    Dim TempObj As clsParticleSource
    For Each TempObj In mCol
        mCol.Remove TempObj.Key
    Next
    Current = ""
End Function

Public Property Get item(vntIndexKey As Variant) As clsParticleSource
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Set item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
    If Current = vntIndexKey Then Current = ""
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Call Clear
    Set mCol = Nothing
End Sub

