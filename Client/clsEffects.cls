VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEffects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"DXControl"
Attribute VB_Ext_KEY = "Member0" ,"DXControl"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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

Public Function Add(x As Long, y As Long, Sprite As Long, Speed As Long, TotalFrames As Long, loopcount As Integer, Target As Long, TargType As Long, Optional XO As Byte = 0, Optional YO As Byte = 0, Optional deathCont As Boolean = False) As clsEffect
    'create a new object
    Dim objNewMember As clsEffect
    Set objNewMember = New clsEffect
    'set the properties passed into the method
    With objNewMember
        .Sprite = Sprite
        .TargetType = TargType
        .Target = Target
        .Frame = 0
        .TotalFrames = TotalFrames
        .loopcount = loopcount
        .Speed = Speed
        .x = x
        .y = y
        .XO = XO
        .deathCont = deathCont
        .YO = YO
        LastKey = LastKey + 1
        If LastKey > 2140000000 Then LastKey = 0
        .Key = "Effect" + CStr(LastKey)
    End With
    
    mCol.Add objNewMember, objNewMember.Key

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Function Clear()
    Dim TempObj As clsEffect
    For Each TempObj In mCol
        mCol.Remove TempObj.Key
    Next
    Current = vbNullString
End Function

Public Property Get item(vntIndexKey As Variant) As clsEffect
Attribute item.VB_UserMemId = 0
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
    If Current = vntIndexKey Then Current = vbNullString
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
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


