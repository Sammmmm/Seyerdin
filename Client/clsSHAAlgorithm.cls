VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSHAAlgorithm"
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

Private mSHA256 As clsSHA256Strategy

Private Sub Class_Initialize()
    InitializeSHAConstants
End Sub


Public Function SHA256(Message() As Byte) As String

    If mSHA256 Is Nothing Then
        Set mSHA256 = New clsSHA256Strategy
    End If

    SHA256 = Run(Message, mSHA256)

End Function

Public Function SHA256FromString(Message As String) As String

    SHA256FromString = SHA256(StrConv(Message, vbFromUnicode))

End Function


Public Function Run(Message() As Byte, strategy As ISHAStrategy) As String
    Dim RoundHash() As Long
    Dim Hash()      As Long
    Dim Msg()       As Long
    Dim Words()     As Long
    Dim lChunk      As Long
    Dim lRound      As Long
    Dim i           As Long
    Dim NumRounds   As Long
    Dim HashSize    As Long

    Hash = strategy.InitializeHash()
    HashSize = UBound(Hash)
    ReDim RoundHash(HashSize)

    Msg = ConvertToWordArray(Message)

    NumRounds = strategy.NumRounds
    ReDim Words(NumRounds - 1)

    For lChunk = 0 To UBound(Msg) Step 16
        For i = 0 To HashSize
            RoundHash(i) = Hash(i)
        Next

        For lRound = 0 To NumRounds - 1
            If lRound < 16 Then
                Words(lRound) = Msg(lRound + lChunk)
            Else
                Words(lRound) = strategy.Expand(Words, lRound)
            End If

            strategy.Round lRound, RoundHash, Words(lRound)
        Next

        For i = 0 To HashSize
            Hash(i) = Add32(Hash(i), RoundHash(i))
        Next
    Next

    Run = strategy.Output(Hash)

End Function

Public Function ConvertToWordArray(sMessage() As Byte) As Long()

    Dim lMessageLength  As Long
    Dim lNumberOfWords  As Long
    Dim lWordArray()    As Long
    Dim lBytePosition   As Long
    Dim lByteCount      As Long
    Dim lWordCount      As Long
    Dim lByte           As Long

    Const MODULUS_BITS      As Long = 512
    Const CONGRUENT_BITS    As Long = 448

    Const DIFF_BYTES        As Long = (MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE
    Const MODULUS_BYTES     As Long = MODULUS_BITS \ BITS_TO_A_BYTE
    Const MODULUS_WORDS     As Long = MODULUS_BITS \ BITS_TO_A_WORD

    lMessageLength = UBound(sMessage) + 1

    ' Get padded number of words. Message needs to be congruent to 448 bits,
    ' modulo 512 bits. If it is exactly congruent to 448 bits, modulo 512 bits
    ' it must still have another 512 bits added. 512 bits = 64 bytes
    ' (or 16 * 4 byte words), 448 bits = 56 bytes. This means lNumberOfWords must
    ' be a multiple of 16 (i.e. 16 * 4 (bytes) * 8 (bits))

    lNumberOfWords = (((lMessageLength + DIFF_BYTES) \ MODULUS_BYTES) + 1) * MODULUS_WORDS
    ReDim lWordArray(lNumberOfWords - 1)

    ' Combine each block of 4 bytes (ascii code of character) into one long
    ' value and store in the message. The high-order (most significant) bit of
    ' each byte is listed first. However, unlike MD5 we put the high-order
    ' (most significant) byte first in each word.
    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        ' Each word is 4 bytes
        lWordCount = lByteCount \ BYTES_TO_A_WORD

        lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE

        lByte = sMessage(lByteCount)

        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(lByte, lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    ' Terminate according to SHA-256 rules with a 1 bit, zeros and the length in
    ' bits stored in the last two words
    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (3 - (lByteCount Mod BYTES_TO_A_WORD)) * BITS_TO_A_BYTE

    ' Add a terminating 1 bit, all the rest of the bits to the end of the
    ' word array will default to zero
    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

    ' We put the length of the message in bits into the last two words, to get
    ' the length in bits we need to multiply by 8 (or left shift 3). This left
    ' shifted value is put in the last word. Any bits shifted off the left edge
    ' need to be put in the penultimate word, we can work out which bits by shifting
    ' right the length by 29 bits.
    lWordArray(lNumberOfWords - 1) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 2) = RShift(lMessageLength, 29)

    ConvertToWordArray = lWordArray
End Function


