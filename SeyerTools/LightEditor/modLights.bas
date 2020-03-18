Attribute VB_Name = "modLights"
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

Public Type LightTemplate
    Name As String * 30
    Red As Byte
    Green As Byte
    Blue As Byte
    Filler(100) As Byte
End Type


Public Lights(50) As LightTemplate
Public CurrentLight As Long

Public Sub SaveLights()
    Dim St As String, A As Long
    For A = 1 To 50
        With Lights(A)
            St = St & .Name & Chr$(.Red) & Chr$(.Green) & Chr$(.Blue) & String$(100, 0)
        End With
    Next A
    
    Dim B As Long
    For A = 1 To Len(St)
        B = B + Asc(Mid$(St, A, 1))
    Next A
    
    B = B Mod 256
    St = St & Chr$(B)
    ChDir App.Path
    
    Open "Light.rsc" For Binary As #1
        Put #1, , St
    Close #1
End Sub

Public Sub LoadLights()
    Dim St As String * 6651, A As Long, Offset As Long
    ChDir App.Path
    
    If Exists("Light.rsc") Then
        Open "Light.rsc" For Binary As #1
            Get #1, , St
        Close #1
    End If
    
    For A = 1 To 50
        Offset = (A - 1) * 133
        With Lights(A)
            .Name = ClipString$(Mid$(St, Offset + 1, 30))
            .Red = Asc(Mid$(St, Offset + 31, 1))
            .Green = Asc(Mid$(St, Offset + 32, 1))
            .Blue = Asc(Mid$(St, Offset + 33, 1))
            
            frmMain.lstLights.AddItem "[" & A & "] " & .Name
        End With
    Next A

End Sub

Public Function Exists(Filename As String) As Boolean
On Error Resume Next
Open Filename For Input As #1
Close #1
If Err.Number <> 0 Then
   Exists = False
Else
   Exists = True
End If
End Function

Function ClipString(St As String) As String
    Dim A As Long
    For A = Len(St) To 1 Step -1
        If Mid$(St, A, 1) <> Chr$(32) And Mid$(St, A, 1) <> Chr$(0) Then
            ClipString = Mid$(St, 1, A)
            Exit Function
        End If
    Next A
End Function
