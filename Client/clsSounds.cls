VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSounds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'local variable to hold collection
Option Explicit
Private mCol As Collection
Public Current As String
Public LastKey As Long

Public Function Play(Filename As String) As clsSound
    'create a new object
    Dim objNewMember As clsSound
    Set objNewMember = New clsSound
Dim A As Long
    
    
    With objNewMember
        LastKey = LastKey + 1
        If LastKey > 2140000000 Then LastKey = 0
        .Key = "Sound" + CStr(LastKey)
    End With
    
    mCol.Add objNewMember, objNewMember.Key
    mCol.Item("Sound" & LastKey).Play Filename
    
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Function Clear()
    Dim TempObj As clsWidget
    For Each TempObj In mCol
        mCol.Remove TempObj.Key
    Next
    Current = ""
End Function

Public Property Get Item(vntIndexKey As Variant) As clsWidget
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    Set Item = mCol(vntIndexKey)
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

