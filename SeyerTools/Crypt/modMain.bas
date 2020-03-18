Attribute VB_Name = "modMain"
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

Sub Main()
ChDir (App.Path)
Dim St1 As String
Dim St() As String
St() = Split(Command$, """ """)
For a = 0 To UBound(St())
St1 = Replace$(St(a), """", "")
DecryptFile St1, St1
Next a
End Sub

Public Sub DecryptFile(SourceFile As String, DestFile As String)

Dim Filenr As Long
Dim ByteArray() As Byte
Dim Offset As Long
Dim ByteLen As Long

'Make sure the source file do exist
    
    'CheckFile SourceFile

    'Open the source file and read the content
    'into a bytearray to decrypt
    Filenr = FreeFile
    Open SourceFile For Binary Access Read As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr

    'Encrypt the bytearray

    'Get the size of the source array
    ByteLen = UBound(ByteArray) + 1

    'Loop thru the data encrypting it with simply XOR´ing with the key
    For Offset = 0 To (ByteLen - 1)
        ByteArray(Offset) = ByteArray(Offset) Xor 170
    Next

    'If the destination file already exist we need
    'to delete it since opening it for binary use
    'will preserve it if it already exist
    If Exists(DestFile) Then Kill DestFile
    
    'Store the decrypted data in the destination file
    Filenr = FreeFile
    Open DestFile For Binary Access Write As #Filenr
    Put #Filenr, , ByteArray()
    Close #Filenr

End Sub

Public Function Exists(FileName As String) As Boolean
On Error Resume Next
Open FileName For Input As #1
Close #1
If Err.Number <> 0 Then
   Exists = False
Else
   Exists = True
End If
End Function

