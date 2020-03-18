VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHashTable"
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

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As _
    Any, source As Any, ByVal bytes As Long)

' default values
Const DEFAULT_HASHSIZE = 1024
Const DEFAULT_LISTSIZE = 2048
Const DEFAULT_CHUNKSIZE = 1024

Private Type SlotType
    Key As String
    Value() As Byte
    nextItem As Long      ' 0 if last item
End Type

' for each hash code this array holds the first element
' in slotTable() with the corresponding hash code
Dim hashTbl() As Long
' the array that holds the data
Dim slotTable() As SlotType

' pointer to first free slot
Dim FreeNdx As Long

' size of hash table
Dim m_HashSize As Long
' size of slot table
Dim m_ListSize As Long
' chunk size
Dim m_ChunkSize As Long
' items in the slot table
Dim m_Count As Long

' member variable for IgnoreCase property
Private m_IgnoreCase As Boolean

' initialize the hash table

Sub SetSize(ByVal HashSize As Long, Optional ByVal ListSize As Long, _
    Optional ByVal ChunkSize As Long)
    m_IgnoreCase = True
    ' provide defaults
    If ListSize <= 0 Then ListSize = m_ListSize
    If ChunkSize <= 0 Then ChunkSize = m_ChunkSize
    ' save size values
    m_HashSize = HashSize
    m_ListSize = ListSize
    m_ChunkSize = ChunkSize
    m_Count = 0
    ' rebuild tables
    FreeNdx = 0
    ReDim hashTbl(0 To HashSize - 1) As Long
    ReDim slotTable(0) As SlotType
    ExpandSlotTable m_ListSize
End Sub

' check whether an item is in the hash table

Function Exists(Key As String) As Boolean
    Exists = GetSlotIndex(Key) <> 0
End Function

' add a new element to the hash table

Sub Add(Key As String, Value() As Byte)
    Dim ndx As Long, Create As Boolean
    
    ' get the index to the slot where the value is
    ' (allocate a new slot if necessary)
    Create = True
    ndx = GetSlotIndex(Key, Create)
    
    If Create Then
        ' the item was actually added
        If IsObject(Value) Then
            'Set slotTable(ndx).Value = Value
        Else
            slotTable(ndx).Value = Value
        End If
    Else
        ' raise error "This key is already associated with an item of this
        ' collection"
        Err.Raise 457
    End If
End Sub

' the value associated to a key
' (empty if not found)

Property Get Item(Key As String) As Byte()
    Dim ndx As Long
    ' get the index to the slot where the value is
    ndx = GetSlotIndex(Key)
    If ndx = 0 Then
        ' return Empty if not found
    'ElseIf IsObject(slotTable(ndx).Value) Then
        'Set Item = slotTable(ndx).Value
    Else
        Item = slotTable(ndx).Value()
    End If
End Property

' remove all items from the hash table

Sub RemoveAll()
    SetSize m_HashSize, m_ListSize, m_ChunkSize
End Sub

' the number of items in the hash table

'Property Get Count() As Long
'    Count = m_Count
'End Property

' the array of all keys

'Property Get Keys() As Variant()
'    Dim i As Long, ndx As Long
'    Dim n As Long
'    ReDim res(0 To m_Count - 1) As Variant
'
'    For i = 0 To m_HashSize - 1
'        ' take the pointer from the hash table
'        ndx = hashTbl(i)
'        ' walk the slottable() array
'        Do While ndx
'            res(n) = slotTable(ndx).Key
'            n = n + 1
'            ndx = slotTable(ndx).nextItem
'        Loop
'    Next
'
'    ' assign to the result
'    Keys = res()
'End Property


'-----------------------------------------
' Private procedures
'-----------------------------------------

Private Sub Class_Initialize()
    ' initialize the tables at default size
    SetSize DEFAULT_HASHSIZE, DEFAULT_LISTSIZE, DEFAULT_CHUNKSIZE
End Sub

' expand the slotTable() array

Private Sub ExpandSlotTable(ByVal numEls As Long)
    Dim newFreeNdx As Long, i As Long
    newFreeNdx = UBound(slotTable) + 1
    
    ReDim Preserve slotTable(0 To UBound(slotTable) + numEls) As SlotType
    ' create the linked list of free items
    For i = newFreeNdx To UBound(slotTable)
        slotTable(i).nextItem = i + 1
    Next
    ' overwrite the last (wrong) value
    slotTable(UBound(slotTable)).nextItem = FreeNdx
    ' we now know where to pick the first free item
    FreeNdx = newFreeNdx
End Sub

' return the hash code of a string

Private Function HashCode(Key As String) As Long
    Dim lastEl As Long, i As Long
    
    ' copy ansi codes into an array of long
    lastEl = (Len(Key) - 1) \ 4
    ReDim codes(lastEl) As Long
    ' this also converts from Unicode to ANSI
    CopyMemory codes(0), ByVal Key, Len(Key)
    
    ' XOR the ANSI codes of all characters
    For i = 0 To lastEl
        HashCode = HashCode Xor codes(i)
    Next
    
End Function

' get the index where an item is stored or 0 if not found
' if Create = True the item is created
'
' on exit Create=True only if a slot has been actually created

Private Function GetSlotIndex(ByVal Key As String, Optional Create As Boolean, _
    Optional HCode As Long, Optional LastNdx As Long) As Long
    Dim ndx As Long
    
    ' raise error if invalid key
    If Len(Key) = 0 Then Err.Raise 1001, , "Invalid key"
    
    ' keep case-unsensitiveness into account
    If m_IgnoreCase Then Key = UCase$(Key)
    ' get the index in the hashTbl() array
    HCode = HashCode(Key) Mod m_HashSize
    ' get the pointer to the slotTable() array
    ndx = hashTbl(HCode)
    
    ' exit if there is no item with that hash code
    Do While ndx
        ' compare key with actual value
        If slotTable(ndx).Key = Key Then Exit Do
        ' remember last pointer
        LastNdx = ndx
        ' check the next item
        ndx = slotTable(ndx).nextItem
    Loop
    
    ' create a new item if not there
    If ndx = 0 And Create Then
        ndx = GetFreeSlot()
        PrepareSlot ndx, Key, HCode, LastNdx
    Else
        ' signal that no item has been created
        Create = False
    End If
    ' this is the return value
    GetSlotIndex = ndx

End Function

' return the first free slot

Private Function GetFreeSlot() As Long
    ' allocate new memory if necessary
    If FreeNdx = 0 Then ExpandSlotTable m_ChunkSize
    ' use the first slot
    GetFreeSlot = FreeNdx
    ' update the pointer to the first slot
    FreeNdx = slotTable(GetFreeSlot).nextItem
    ' signal this as the end of the linked list
    slotTable(GetFreeSlot).nextItem = 0
    ' we have one more item
    m_Count = m_Count + 1
End Function

' assign a key and value to a given slot

Private Sub PrepareSlot(ByVal Index As Long, ByVal Key As String, _
    ByVal HCode As Long, ByVal LastNdx As Long)
    ' assign the key
    ' keep case-sensitiveness into account
    If m_IgnoreCase Then Key = UCase$(Key)
    slotTable(Index).Key = Key
    
    If LastNdx Then
        ' this is the successor of another slot
        slotTable(LastNdx).nextItem = Index
    Else
        ' this is the first slot for a given hash code
        hashTbl(HCode) = Index
    End If
End Sub




