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

Attribute VB_Name = "Module1"
Option Compare Database

Sub ArmorData()
Dim wb As Workbook
Dim db As Database, rec As Recordset
Dim x As String, y As String, z As String
Dim objnum As String, minlevel As String, mindam As String, maxdam As String
Dim A As String, C As String, D As String

 Set db = CurrentDb
Set rec = db.OpenRecordset("Objects")
Set wb = Workbooks.Open("C:\Users\Adi\Games\Seyerdin\itembalance.xls", True, True)
    ' open the source workbook, read only

    
    'Dim i As Integer
   For i = 1 To 150
            A = "A" & i
            C = "C" & i
            D = "D" & i
            
            objnum = wb.Worksheets("Sheet2").Range(A).Value
            minlevel = wb.Worksheets("Sheet2").Range(C).Value
            defence = wb.Worksheets("Sheet2").Range(D).Value
            
If IsNumeric(objnum) Then

            sqlString = "UPDATE Objects SET Data3=" & defence & ", Minlevel=" & minlevel & " WHERE Number=" & objnum & ";"
            db.Execute (sqlString)
End If
            'i = i + 1
newline:
   Next
   
    

   
    wb.Close False ' close the source workbook without saving any changes
Set wb = Nothing ' free memory
rec.Close
db.Close
End Sub


Sub WeaponData()
Dim wb As Workbook
Dim db As Database, rec As Recordset
Dim x As String, y As String, z As String
Dim objnum As String, minlevel As String, mindam As String, maxdam As String
Dim A As String, C As String, D As String, E As String


 Set db = CurrentDb
Set rec = db.OpenRecordset("Objects")
Set wb = Workbooks.Open("C:\Users\Adi\Games\Seyerdin\itembalance.xls", True, True)
    ' open the source workbook, read only

    
    'Dim i As Integer
   For i = 1 To 100
            A = "A" & i
            C = "C" & i
            D = "D" & i
            E = "E" & i
            
            objnum = wb.Worksheets("Sheet1").Range(A).Value
            minlevel = wb.Worksheets("Sheet1").Range(C).Value
            mindam = wb.Worksheets("Sheet1").Range(D).Value
            maxdam = wb.Worksheets("Sheet1").Range(E).Value
            

If IsNumeric(objnum) Then

            sqlString = "UPDATE Objects SET Data2=" & mindam & ", Data4= " & maxdam & ", Minlevel=" & minlevel & " WHERE Number=" & objnum & ";"
            db.Execute (sqlString)
End If
            'i = i + 1
newline:
   Next
   
    

   
    wb.Close False ' close the source workbook without saving any changes
Set wb = Nothing ' free memory
rec.Close
db.Close
End Sub

