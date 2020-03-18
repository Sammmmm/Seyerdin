Attribute VB_Name = "modComEnc"
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


'**************************************************
'         Compression/Encryption Module
'**************************************************

Private m_sBoxRC4(0 To 255) As Integer
Private m_KeyS As String

Public MapKey As String
Public MyKey As String

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)

Public Function DecryptFile(SourceFile As String, Dest() As Byte) As Long

Dim Filenr As Integer
Dim offset As Long
Dim ByteLen As Long

'Make sure the source file do exist
    
    CheckFile SourceFile

    'Open the source file and read the content
    'into a bytearray to decrypt
    Filenr = FreeFile
    Open SourceFile For Binary Access Read As #Filenr
    ReDim Dest(0 To LOF(Filenr) - 1)
    Get #Filenr, , Dest()
    Close #Filenr
    ByteLen = UBound(Dest) + 1
    For offset = 0 To (ByteLen - 1)
        Dest(offset) = Dest(offset) Xor 170
    Next
    DecryptFile = ByteLen
End Function

Public Sub EncryptFile(SourceFile As String, DestFile As String)

Dim Filenr As Long
Dim ByteArray() As Byte
Dim offset As Long
Dim ByteLen As Long

'Make sure the source file do exist

    CheckFile SourceFile

    'Open the source file and read the content
    'into a bytearray to pass onto encryption
    Filenr = FreeFile
    Open SourceFile For Binary Access Read As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr

    'Encrypt the bytearray

    'Get the size of the source array
    ByteLen = UBound(ByteArray) + 1

    'Loop thru the data encrypting it with simply XOR´ing with the key
    For offset = 0 To (ByteLen - 1)
        ByteArray(offset) = ByteArray(offset) Xor 170
    Next

    'If the destination file already exist we need
    'to delete it since opening it for binary use
    'will preserve it if it already exist
    If (Exists(DestFile)) Then Kill DestFile

    'Store the encrypted data in the destination file
    Filenr = FreeFile
    Open DestFile For Binary Access Write As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub

Public Function RC4_EncryptString(Text As String, Optional Key As String) As String

Dim ByteArray() As Byte

'Convert the data into a byte array

    ByteArray() = StrConv(Text, vbFromUnicode)

    'Encrypt the byte array
    Call RC4_EncryptByte(ByteArray(), Key)

    'Convert the byte array back into a string
    RC4_EncryptString = StrConv(ByteArray(), vbUnicode)

End Function

Public Function RC4_DecryptString(Text As String, Optional Key As String) As String

Dim ByteArray() As Byte

'Convert the data into a byte array

    ByteArray() = StrConv(Text, vbFromUnicode)

    'Decrypt the byte array
    Call RC4_DecryptByte(ByteArray(), Key)

    'Convert the byte array back into a string
    RC4_DecryptString = StrConv(ByteArray(), vbUnicode)

End Function

Public Sub RC4_EncryptByte(ByteArray() As Byte, Optional Key As String)

Dim i As Long
Dim j As Long
Dim Temp As Byte
Dim offset As Long
Dim OrigLen As Long
Dim sBox(0 To 255) As Integer

    'Set the new key (optional)
    If (Len(Key) > 0) Then RC4_SetKey Key

    'Create a local copy of the sboxes, this
    'is much more elegant than recreating
    'before encrypting/decrypting anything
    Call CopyMem(sBox(0), m_sBoxRC4(0), 512)

    'Get the size of the source array
    OrigLen = UBound(ByteArray) + 1

    'Encrypt the data
    For offset = 0 To (OrigLen - 1)
        i = (i + 1) Mod 256
        j = (j + sBox(i)) Mod 256
        Temp = sBox(i)
        sBox(i) = sBox(j)
        sBox(j) = Temp
        ByteArray(offset) = ByteArray(offset) Xor (sBox((sBox(i) + sBox(j)) Mod 256))
    Next

End Sub

Public Sub RC4_DecryptByte(ByteArray() As Byte, Optional Key As String)

'The same routine is used for encryption as well
'decryption so why not reuse some code and make
'this class smaller (that is it it wasn't for all
'those damn comments ;))

    Call RC4_EncryptByte(ByteArray(), Key)

End Sub


Public Sub RC4_SetKey(New_Value As String)

Dim A As Long
Dim b As Long
Dim Temp As Byte
Dim Key() As Byte
Dim KeyLen As Long

'Do nothing if the key is buffered

    If (m_KeyS = New_Value) Then Exit Sub

    'Set the new key
    m_KeyS = New_Value

    'Save the password in a byte array
    Key() = StrConv(m_KeyS, vbFromUnicode)
    KeyLen = Len(m_KeyS)

    'Initialize s-boxes
    For A = 0 To 255
        m_sBoxRC4(A) = A
    Next A
    For A = 0 To 255
        b = (b + m_sBoxRC4(A) + Key(A Mod KeyLen)) Mod 256
        Temp = m_sBoxRC4(A)
        m_sBoxRC4(A) = m_sBoxRC4(b)
        m_sBoxRC4(b) = Temp
    Next

End Sub
