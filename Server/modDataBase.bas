Attribute VB_Name = "modDataBase"
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

'Database Objects
Option Explicit
Public WS As Workspace
Public db As Database

Public UserRS As Recordset
Public NPCRS As Recordset
Public MonsterRS As Recordset
Public ObjectRS As Recordset
Public DataRS As Recordset
Public MapRS As Recordset
Public GuildRS As Recordset
Public PrefixRS As Recordset
Public BanRS As Recordset
Public HallRS As Recordset
Public ScriptRS As Recordset
Public LightsRS As Recordset

'Public AdoCon As ADODB.Connection



Sub CreateAccountsTable(db As Database)
    Dim A As Long
    Dim TD As TableDef
    Dim NewField As Field
    Dim NewIndex As Index
    
    'Create Accounts Table
    Set TD = db.CreateTableDef("Accounts")

    'Create Fields
    'Account Data
    Set NewField = TD.CreateField("User", dbText, 15)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Email", dbMemo)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Access", dbByte)
    TD.Fields.Append NewField
    
    'Character Data    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Password", dbText, 64)
    TD.Fields.Append NewField

    Set NewField = TD.CreateField("CharNum", dbLong)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Name", dbText, 15)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Class", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Gender", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Sprite", dbByte)
    TD.Fields.Append NewField
    'Set NewField = TD.CreateField("Desc", dbMemo)
    'NewField.AllowZeroLength = True
    'TD.Fields.Append NewField
    Set NewField = TD.CreateField("Squelched", dbByte)
    TD.Fields.Append NewField
    
    'Position Data
    Set NewField = TD.CreateField("Map", dbInteger)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("X", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Y", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("D", dbByte)
    TD.Fields.Append NewField
    
    'Vital Stat Data
    Set NewField = TD.CreateField("MaxHP", dbInteger)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("MaxEnergy", dbInteger)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("MaxMana", dbInteger)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("HP", dbInteger)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Energy", dbInteger)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Mana", dbInteger)
    TD.Fields.Append NewField
    
    Set NewField = TD.CreateField("Level", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Experience", dbLong)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("StatPoints", dbInteger)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("StatusEffect", dbLong)
    TD.Fields.Append NewField
    For A = 1 To 64
        Set NewField = TD.CreateField("StatusData" + CStr(A), dbText, 8)
        TD.Fields.Append NewField
    Next A
    Set NewField = TD.CreateField("SkillLevels", dbText, 255)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("SkillEXP", dbMemo)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("SkillPoints", dbInteger)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Renown", dbLong)
    TD.Fields.Append NewField
    
    'Misc Data
    Set NewField = TD.CreateField("Bank", dbLong)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Status", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("LastPlayed", dbLong)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Flags", dbMemo)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    
    'Statisical Data
    Set NewField = TD.CreateField("Strength", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Agility", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Endurance", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Constitution", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Wisdom", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Intelligence", dbByte)
    TD.Fields.Append NewField
    
    'Inventory Data
    For A = 1 To 20
        Set NewField = TD.CreateField("InvObject" + CStr(A), dbText, 29)
        TD.Fields.Append NewField
    Next A
    
    Set NewField = TD.CreateField("NumStoragePages", dbByte)
    TD.Fields.Append NewField
    For A = 1 To STORAGEPAGES
        Set NewField = TD.CreateField("StorageObjects" + CStr(A), dbMemo)
        TD.Fields.Append NewField
    Next A
    
    For A = 1 To 5
        Set NewField = TD.CreateField("Equipped" + CStr(A), dbText, 29)
        TD.Fields.Append NewField
    Next A
    
    'Create Indexes
    Set NewIndex = TD.CreateIndex("User")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("User")
    NewIndex.Fields.Append NewField
    TD.Indexes.Append NewIndex
    
    Set NewIndex = TD.CreateIndex("Name")
    NewIndex.Primary = False
    NewIndex.Unique = False
    Set NewField = NewIndex.CreateField("Name")
    NewIndex.Fields.Append NewField
    TD.Indexes.Append NewIndex
    
    Set NewIndex = TD.CreateIndex("CharNum")
    NewIndex.Primary = False
    NewIndex.Unique = False
    Set NewField = NewIndex.CreateField("CharNum")
    NewIndex.Fields.Append NewField
    TD.Indexes.Append NewIndex
       
    'Append Accounts Table
    db.TableDefs.Append TD
End Sub

Sub CreateDataTable()
    Dim A As Long
    Dim TD As TableDef
    Dim NewField As Field
    
    'Create Accounts Table
    Set TD = db.CreateTableDef("Data")

    'Create Fields
    Set NewField = TD.CreateField("User", dbText, 15)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Password", dbText, 64)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("MOTD", dbMemo)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("MapResetTime", dbLong)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("TimeAllowance", dbLong)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("BackupInterval", dbLong)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("MoneyObj", dbInteger)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("MsgObj", dbByte)
    TD.Fields.Append NewField
    For A = 0 To 4
        Set NewField = TD.CreateField("StartLocationMap" + CStr(A), dbInteger)
        TD.Fields.Append NewField
        Set NewField = TD.CreateField("StartLocationX" + CStr(A), dbByte)
        TD.Fields.Append NewField
        Set NewField = TD.CreateField("StartLocationY" + CStr(A), dbByte)
        TD.Fields.Append NewField
        Set NewField = TD.CreateField("StartLocationMessage" + CStr(A), dbText, 255)
        NewField.AllowZeroLength = True
        TD.Fields.Append NewField
    Next A
    For A = 1 To 8
        Set NewField = TD.CreateField("StartingObj" + CStr(A), dbInteger)
        TD.Fields.Append NewField
        Set NewField = TD.CreateField("StartingObjVal" + CStr(A), dbLong)
        TD.Fields.Append NewField
    Next A
    Set NewField = TD.CreateField("CharCounter", dbLong)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("MsgCounter", dbLong)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Day", dbLong)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Hour", dbLong)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("LastUpdate", dbLong)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("ObjectData", dbMemo)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Flags", dbMemo)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    For A = 0 To 127
        Set NewField = TD.CreateField("StringFlag" + CStr(A), dbMemo)
        NewField.AllowZeroLength = True
        TD.Fields.Append NewField
    Next A
    'Append Data Table
    db.TableDefs.Append TD
End Sub

Sub CreateMapsTable()
    Dim TD As TableDef
    Dim NewField As Field
    Dim NewIndex As Index
    
    'Create Accounts Table
    Set TD = db.CreateTableDef("Maps")

    'Create Fields
    Set NewField = TD.CreateField("Number", dbInteger)
    TD.Fields.Append NewField
    
    Set NewField = TD.CreateField("Data", dbMemo)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    
    'Create Indexes
    Set NewIndex = TD.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    TD.Indexes.Append NewIndex
    
    'Append Maps Table
    db.TableDefs.Append TD
End Sub
Sub CreateGuildsTable(db As Database)
    Dim A As Long
    Dim TD As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Guilds Table
    Set TD = db.CreateTableDef("Guilds")
    Set NewField = TD.CreateField("Number", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Name", dbText, 20)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Hall", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Sprite", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Bank", dbLong)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("DueDate", dbLong)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("GuildKills", dbInteger)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("GuildDeaths", dbInteger)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Symbol", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Founded", dbDate)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("MOTD", dbMemo)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Info", dbMemo)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField

    
    For A = 0 To 19
        Set NewField = TD.CreateField("MemberName" + CStr(A), dbText, 15)
        NewField.AllowZeroLength = True
        TD.Fields.Append NewField
        Set NewField = TD.CreateField("MemberRank" + CStr(A), dbByte)
        TD.Fields.Append NewField
        Set NewField = TD.CreateField("MemberKills" + CStr(A), dbInteger)
        TD.Fields.Append NewField
        Set NewField = TD.CreateField("MemberDeaths" + CStr(A), dbInteger)
        TD.Fields.Append NewField
        Set NewField = TD.CreateField("MemberRenown" + CStr(A), dbLong)
        TD.Fields.Append NewField
        Set NewField = TD.CreateField("Joined" + CStr(A), dbDate)
        TD.Fields.Append NewField
    Next A
    
    For A = 0 To 9
        Set NewField = TD.CreateField("DeclarationGuild" + CStr(A), dbByte)
        TD.Fields.Append NewField
        Set NewField = TD.CreateField("DeclarationType" + CStr(A), dbByte)
        TD.Fields.Append NewField
    Next A
    
    'Create Indexes
    Set NewIndex = TD.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = False
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    TD.Indexes.Append NewIndex
    
    'Append Guilds Table
    db.TableDefs.Append TD
End Sub
Sub CreateBansTable()
    Dim TD As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Bans Table
    Set TD = db.CreateTableDef("Bans")
    Set NewField = TD.CreateField("Number", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Name", dbText, 15)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Banner", dbText, 15)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("User", dbText, 30)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("IP", dbText, 30)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("UID", dbText, 50)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("iniUID", dbText, 50)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Reason", dbText, 255)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("UnbanDate", dbLong)
    TD.Fields.Append NewField
    
    
    'Create Indexes
    Set NewIndex = TD.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    TD.Indexes.Append NewIndex
    
    'Append Bans Table
    db.TableDefs.Append TD
End Sub
Sub CreateHallsTable(db As Database)
    Dim TD As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Halls Table
    Set TD = db.CreateTableDef("Halls")
    Set NewField = TD.CreateField("Number", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Name", dbText, 15)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Price", dbLong)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Upkeep", dbLong)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("StartLocationMap", dbInteger)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("StartLocationX", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("StartLocationY", dbByte)
    TD.Fields.Append NewField
    
    'Create Indexes
    Set NewIndex = TD.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    TD.Indexes.Append NewIndex
    
    'Append Bans Table
    db.TableDefs.Append TD
End Sub
Sub CreatePreFixTable()
    Dim TD As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Prefix Table
    Set TD = db.CreateTableDef("Prefix")
    Set NewField = TD.CreateField("Number", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Name", dbText, 15)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Data", dbText, 12)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Version", dbLong)
    TD.Fields.Append NewField
    
    'Create Indexes
    Set NewIndex = TD.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = False
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    TD.Indexes.Append NewIndex
    
    'Append Messages Table
    db.TableDefs.Append TD
End Sub

Sub CreateLightsTable()
    Dim TD As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Prefix Table
    Set TD = db.CreateTableDef("Lights")
    Set NewField = TD.CreateField("Number", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Name", dbText, 15)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Data", dbText, 50)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Version", dbLong)
    TD.Fields.Append NewField
    
    'Create Indexes
    Set NewIndex = TD.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = False
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    TD.Indexes.Append NewIndex
    
    'Append Messages Table
    db.TableDefs.Append TD
End Sub

Sub CreateScriptsTable()
    Dim TD As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Messages Table
    Set TD = db.CreateTableDef("Scripts")
    Set NewField = TD.CreateField("Name", dbText, 15)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Source", dbMemo)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Data", dbMemo)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
        
    'Create Indexes
    Set NewIndex = TD.CreateIndex("Name")
    NewIndex.Primary = True
    NewIndex.Unique = False
    Set NewField = NewIndex.CreateField("Name")
    NewIndex.Fields.Append NewField
    TD.Indexes.Append NewIndex
    
    'Append Messages Table
    db.TableDefs.Append TD
End Sub

Sub CreateNPCsTable()
    Dim A As Long
    Dim TD As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create NPC Table
    Set TD = db.CreateTableDef("NPCs")

    'Create Fields
    Set NewField = TD.CreateField("Number", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Name", dbText, 15)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("JoinText", dbText, 255)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("LeaveText", dbText, 255)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    For A = 0 To 4
        Set NewField = TD.CreateField("SayText" + CStr(A), dbText, 255)
        NewField.AllowZeroLength = True
        TD.Fields.Append NewField
    Next A
    Set NewField = TD.CreateField("Flags", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Portrait", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Sprite", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Direction", dbByte)
    TD.Fields.Append NewField
    
    For A = 0 To 9
        Set NewField = TD.CreateField("GiveObject" + CStr(A), dbInteger)
        TD.Fields.Append NewField
        Set NewField = TD.CreateField("GiveValue" + CStr(A), dbLong)
        TD.Fields.Append NewField
        Set NewField = TD.CreateField("TakeObject" + CStr(A), dbInteger)
        TD.Fields.Append NewField
        Set NewField = TD.CreateField("TakeValue" + CStr(A), dbLong)
        TD.Fields.Append NewField
    Next A
            
    'Create Indexes
    Set NewIndex = TD.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    TD.Indexes.Append NewIndex
    
    'Append NPC Table
    db.TableDefs.Append TD
End Sub

Public Sub UpdateDataBase()
On Error Resume Next
'I'm leaving this here so I can use modify it if I need it again
'I never mess with the database portion
Dim NewField As Field
Dim A As Long, x As Long, y As Long

'Set NewField = db.TableDefs("Monsters").CreateField("Alpha", dbByte)
'db.TableDefs("Monsters").Fields.Append NewField
'Set NewField = db.TableDefs("Monsters").CreateField("Red", dbByte)
'db.TableDefs("Monsters").Fields.Append NewField
'Set NewField = db.TableDefs("Monsters").CreateField("Green", dbByte)
'db.TableDefs("Monsters").Fields.Append NewField
'Set NewField = db.TableDefs("Monsters").CreateField("Blue", dbByte)
'db.TableDefs("Monsters").Fields.Append NewField
'Set NewField = db.TableDefs("Monsters").CreateField("Light", dbByte)
'db.TableDefs("Monsters").Fields.Append NewField'

'Set NewField = db.TableDefs("Monsters").CreateField("Flags2", dbByte)
'db.TableDefs("Monsters").Fields.Append NewField


'Set MonsterRS = db.TableDefs("Monsters").OpenRecordset(dbOpenTable)

'    If MonsterRS.BOF = False Then
'        MonsterRS.MoveFirst
'        While MonsterRS.EOF = False
'            a = MonsterRS!Number
'            If a > 0 Then
'                MonsterRS.Edit
'                MonsterRS.Fields("Alpha").Value = 255
'                MonsterRS.Fields("Red").Value = 255
'                MonsterRS.Fields("Green").Value = 255
'                MonsterRS.Fields("Blue").Value = 255
'                MonsterRS.Fields("Light").Value = 0
''                MonsterRS.Update
'            End If
'            MonsterRS.MoveNext
'        Wend
'    End If
    
'    MonsterRS.Close

'Exit Sub

'Set MapRS = db.OpenRecordset("Maps")
'    Dim St As String, st1 As String
'    If MapRS.BOF = False Then
'        MapRS.MoveFirst
'        While MapRS.EOF = False
'            a = MapRS!Number
'            If a > 0 Then
'                St = MapRS!Data
'                st1 = Mid$(St, 1, 30)
'                st1 = st1 + QuadChar(0)
'                st1 = st1 + Mid$(St, 35, 36)
'                For y = 0 To 11
'                    For x = 0 To 11
'                        a = 71 + y * 192 + x * 16
'                        st1 = st1 & Mid$(St, a, 6)
'                        st1 = st1 & Chr2(0) & Chr2(0)
'                        st1 = st1 & Mid$(St, a + 8, 8)
'                    Next x
'                Next y
'                MapRS.Edit
'                MapRS!Data = St
'                MapRS.Update
'            End If
'            MapRS.MoveNext
'        Wend
'    End If
'Exit Sub
'    db.TableDefs.Delete ("Accounts")
'
'    For a = 0 To 7
'        db.TableDefs("Objects").Fields.Delete ("Class" & a)
'    Next a
'
'        Set NewField = db.TableDefs("Objects").CreateField("Class", dbInteger)
'        NewField.DefaultValue = 0
'        db.TableDefs("Objects").Fields.Append NewField

        Set NewField = db.TableDefs("Objects").CreateField("Class", dbInteger)
        NewField.DefaultValue = 0
        db.TableDefs("Objects").Fields.Append NewField


Exit Sub


'DB.TableDefs("Objects").Fields.Delete ("Description")
'Set NewField = DB.TableDefs("Objects").CreateField("Description", dbMemo)
'DB.TableDefs("Objects").Fields.Append NewField
Exit Sub



For A = 0 To 2
    'DB.TableDefs("NPCs").Fields.Delete ("GiveObject" & CStr(A))
    db.TableDefs("Monsters").Fields.Delete ("Object" & CStr(A))
    Set NewField = db.TableDefs("Monsters").CreateField("Object" + CStr(A), dbInteger)
    db.TableDefs("Monsters").Fields.Append NewField
    'Set NewField = DB.TableDefs("NPCs").CreateField("TakeObject" + CStr(A), dbInteger)
    'DB.TableDefs("NPCs").Fields.Append NewField
Next A
    


For A = 0 To 2
    db.TableDefs("Monsters").Fields.Delete ("Object2" & CStr(A))
    'DB.TableDefs("NPCs").Fields.Delete ("TakeObject2" & CStr(A))
Next A
    
End Sub

Sub CreateMonstersTable()
    Dim A As Long
    Dim TD As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create NPC Table
    Set TD = db.CreateTableDef("Monsters")

    'Create Fields
    Set NewField = TD.CreateField("Number", dbInteger)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Name", dbText, 15)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Description", dbText, 255)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Sprite", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("HP", dbInteger)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Min", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Max", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Armor", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Speed", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Sight", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Agility", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Flags", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Flags2", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Experience", dbInteger)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Level", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("MagicResist", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("cStatusEffect", dbLong)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("MonsterType", dbLong)
    TD.Fields.Append NewField
    For A = 0 To 2
        Set NewField = TD.CreateField("Object" + CStr(A), dbInteger)
        TD.Fields.Append NewField
        Set NewField = TD.CreateField("Value" + CStr(A), dbLong)
        TD.Fields.Append NewField
        Set NewField = TD.CreateField("Chance" + CStr(A), dbByte)
        TD.Fields.Append NewField
    Next A
    Set NewField = TD.CreateField("DeathSound", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("AttackSound", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("MoveSpeed", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("AttackSpeed", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Wander", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Alpha", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Red", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Green", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Blue", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Light", dbByte)
    TD.Fields.Append NewField

    'Create Indexes
    Set NewIndex = TD.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    TD.Indexes.Append NewIndex
    
    'Append Monster Table
    db.TableDefs.Append TD
End Sub
Sub CreateObjectsTable(db As Database)
    Dim TD As TableDef
    Dim NewField As Field
    Dim NewIndex As Index

    'Create Objects Table
    Set TD = db.CreateTableDef("Objects")

    'Create Fields
    Set NewField = TD.CreateField("Number", dbInteger)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Name", dbText, 30)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Description", dbMemo)
    NewField.AllowZeroLength = True
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Picture", dbInteger)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Type", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Data1", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Data2", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Data3", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Data4", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Data5", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Data6", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Data7", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Data8", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Data9", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Data10", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Class", dbInteger)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("MinLevel", dbByte)
    TD.Fields.Append NewField
    Set NewField = TD.CreateField("Level", dbByte)
    TD.Fields.Append NewField
    'Flags
    Set NewField = TD.CreateField("Flags", dbByte)
    TD.Fields.Append NewField
    
    Set NewField = TD.CreateField("EquipmentPicture", dbByte)
    TD.Fields.Append NewField
    
    
    'Create Indexes
    Set NewIndex = TD.CreateIndex("Number")
    NewIndex.Primary = True
    NewIndex.Unique = True
    Set NewField = NewIndex.CreateField("Number")
    NewIndex.Fields.Append NewField
    TD.Indexes.Append NewIndex
    
    'Append Object Table
    db.TableDefs.Append TD
End Sub

Public Sub InitDB()
Dim St As String, bDataRSError As Boolean, CurDate As Long
Dim A As Long, B As Long
On Error Resume Next
CurDate = CLng(Date)
    Err.Clear
    
    Set DataRS = db.TableDefs("Data").OpenRecordset(dbOpenTable)
    If Err.Number > 0 Then
        CreateDataTable
        Set DataRS = db.TableDefs("Data").OpenRecordset(dbOpenTable)
    End If
    
    
    
    Err.Clear
    Set UserRS = db.TableDefs("Accounts").OpenRecordset(dbOpenTable)
    If Err.Number > 0 Then
        CreateAccountsTable db
        Set UserRS = db.TableDefs("Accounts").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set NPCRS = db.TableDefs("NPCs").OpenRecordset(dbOpenTable)
    If Err.Number > 0 Then
        CreateNPCsTable
        Set NPCRS = db.TableDefs("NPCs").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set MonsterRS = db.TableDefs("Monsters").OpenRecordset(dbOpenTable)
    If Err.Number > 0 Then
        CreateMonstersTable
        Set MonsterRS = db.TableDefs("Monsters").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set ObjectRS = db.TableDefs("Objects").OpenRecordset(dbOpenTable)
    If Err.Number > 0 Then
        CreateObjectsTable db
        Set ObjectRS = db.TableDefs("Objects").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set MapRS = db.TableDefs("Maps").OpenRecordset(dbOpenTable)
    If Err.Number > 0 Then
        CreateMapsTable
        Set MapRS = db.TableDefs("Maps").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set BanRS = db.TableDefs("Bans").OpenRecordset(dbOpenTable)
    If Err.Number > 0 Then
        CreateBansTable
        Set BanRS = db.TableDefs("Bans").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set GuildRS = db.TableDefs("Guilds").OpenRecordset(dbOpenTable)
    If Err.Number > 0 Then
        CreateGuildsTable db
        Set GuildRS = db.TableDefs("Guilds").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set HallRS = db.TableDefs("Halls").OpenRecordset(dbOpenTable)
    If Err.Number > 0 Then
        CreateHallsTable db
        Set HallRS = db.TableDefs("Halls").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set ScriptRS = db.TableDefs("Scripts").OpenRecordset(dbOpenTable)
    If Err.Number > 0 Then
        CreateScriptsTable
        Set ScriptRS = db.TableDefs("Scripts").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set PrefixRS = db.TableDefs("Prefix").OpenRecordset(dbOpenTable)
    If Err.Number > 0 Then
        CreatePreFixTable
        Set PrefixRS = db.TableDefs("Prefix").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
    Set LightsRS = db.TableDefs("Lights").OpenRecordset(dbOpenTable)
    If Err.Number > 0 Then
        CreateLightsTable
        Set LightsRS = db.TableDefs("Lights").OpenRecordset(dbOpenTable)
    End If
    
    Err.Clear
        
    UserRS.Index = "User"
    ObjectRS.Index = "Number"
    NPCRS.Index = "Number"
    MonsterRS.Index = "Number"
    MapRS.Index = "Number"
    PrefixRS.Index = "Number"
    LightsRS.Index = "Number"
    BanRS.Index = "Number"
    GuildRS.Index = "Number"
    HallRS.Index = "Number"
    ScriptRS.Index = "Name"
    
    CreateModString
    CreateClassData
    
    frmLoading.lblStatus = "Loading World Data.."
    frmLoading.lblStatus.Refresh
    
ReloadData:
    
    Set DataRS = db.TableDefs("Data").OpenRecordset(dbOpenTable)

    'Check if World Data exists
    If DataRS.RecordCount = 0 Then
        'Create default data
       With DataRS
            .AddNew
            !user = ""
            !password = ""
            !BackupInterval = 2
            !MapResetTime = 120000
            !TimeAllowance = 0
            !MoneyObj = 1
            !MsgObj = 2
            !CharCounter = 0
            !MsgCounter = 0
            !MOTD = ""
            !Hour = 12
            !Day = 1
            !LastUpdate = CLng(Date)
            !ObjectData = ""
            !Flags = String$(1024, 0)
            For A = 0 To 127
                .Fields("StringFlag" + CStr(A)) = ""
            Next A
            For A = 0 To 4
                .Fields("StartLocationX" + CStr(A)) = 5
                .Fields("StartLocationY" + CStr(A)) = 5
                .Fields("StartLocationMap" + CStr(A)) = 1
                .Fields("StartLocationMessage" + CStr(A)) = ""
            Next A
            For A = 1 To 8
                .Fields("StartingObj" + CStr(A)) = 0
                .Fields("StartingObjVal" + CStr(A)) = 0
            Next A
            .Update
            .MoveFirst
        End With
    End If
    
    On Error GoTo DataRSError
    
    'Load World Data
    LoadObjectData DataRS!ObjectData
    
    With World
        .BackupInterval = DataRS!BackupInterval
        .MapResetTime = DataRS!MapResetTime
        .TimeAllowance = DataRS!TimeAllowance
        .MoneyObj = DataRS!MoneyObj
        .MsgCounter = DataRS!MsgCounter
        .CharCounter = DataRS!CharCounter
        .Hour = DataRS!Hour
        .LastUpdate = DataRS!LastUpdate
        .MOTD = DataRS!MOTD
        .MapResetTime = 120000
        
        St = DataRS!Flags
        For A = 0 To 255
            .Flag(A) = Asc(Mid$(St, A * 4 + 1, 1)) * 16777216 + Asc(Mid$(St, A * 4 + 2, 1)) * 65536 + Asc(Mid$(St, A * 4 + 3, 1)) * 256& + Asc(Mid$(St, A * 4 + 4, 1))
        Next A
        For A = 0 To 127
            .StringFlag(A) = DataRS("StringFlag" + CStr(A))
        Next A
        For A = 0 To 4
            .StartLocation(A).x = DataRS("StartLocationX" + CStr(A))
            .StartLocation(A).y = DataRS("StartLocationY" + CStr(A))
            .StartLocation(A).map = DataRS("StartLocationMap" + CStr(A))
            .StartLocation(A).Message = DataRS("StartLocationMessage" + CStr(A))
        Next A
        
        For A = 1 To 8
            .StartObjects(A) = DataRS("StartingObj" + CStr(A))
            .StartObjValues(A) = DataRS("StartingObjVal" + CStr(A))
        Next A
    End With
    
    On Error GoTo 0
    
    If bDataRSError Then
        If MsgBox("There was an error loading the server options.  The database will be rebuilt, but some data may be lost.", vbYesNo, TitleString) = vbOK Then
            DataRS.Close
            db.TableDefs.Delete "Data"
            CreateDataTable
            bDataRSError = False
            GoTo ReloadData
        End If
    End If
    
    If World.LastUpdate > CurDate Or Abs(World.LastUpdate - CurDate) >= 30 Then
        CurDate = CLng(Date)
    End If
    
    frmLoading.lblStatus = "Checking Accounts.."
    frmLoading.lblStatus.Refresh
    If UserRS.BOF = False Then
        UserRS.MoveFirst
        While UserRS.EOF = False
            If CurDate - UserRS!lastplayed >= 30 Then
                'DeleteAccount
            End If
            UserRS.MoveNext
        Wend
    End If
    
    frmLoading.lblStatus = "Loading Guilds.."
    frmLoading.lblStatus.Refresh
    If GuildRS.BOF = False Then
        GuildRS.MoveFirst
        While GuildRS.EOF = False
            A = GuildRS!Number
            If A > 0 Then
                With Guild(A)
                    .Name = GuildRS!Name
                    .Bank = GuildRS!Bank
                    .DueDate = GuildRS!DueDate
                    .Hall = GuildRS!Hall
                    .sprite = GuildRS!sprite
                    .MOTD = GuildRS!MOTD
                    .Info = GuildRS!Info
                    .Symbol3 = GuildRS!GuildDeaths
                    .Symbol2 = GuildRS!GuildKills
                    .founded = GuildRS!founded
                    .Symbol1 = GuildRS!Symbol
                    For B = 0 To 19
                        .Member(B).Name = GuildRS("MemberName" + CStr(B))
                        .Member(B).Rank = GuildRS("MemberRank" + CStr(B))
                        .Member(B).Kills = GuildRS("MemberKills" + CStr(B))
                        .Member(B).Deaths = GuildRS("MemberDeaths" + CStr(B))
                        .Member(B).Renown = GuildRS("MemberRenown" + CStr(B))
                        .Member(B).Joined = GuildRS("Joined" + CStr(B))
                    Next B
                    For B = 0 To 9
                        .Declaration(B).Guild = GuildRS("DeclarationGuild" + CStr(B))
                        .Declaration(B).Type = GuildRS("DeclarationType" + CStr(B))
                    Next B
                    .Bookmark = GuildRS.Bookmark
                End With
            End If
            GuildRS.MoveNext
        Wend
    End If
    
    frmLoading.lblStatus = "Loading Halls.."
    frmLoading.lblStatus.Refresh
    If HallRS.BOF = False Then
        HallRS.MoveFirst
        While HallRS.EOF = False
            A = HallRS!Number
            If A > 0 Then
                With Hall(A)
                    .Name = HallRS!Name
                    .Price = HallRS!Price
                    .Upkeep = HallRS!Upkeep
                    .StartLocation.map = HallRS!StartLocationMap
                    .StartLocation.x = HallRS!StartLocationX
                    .StartLocation.y = HallRS!StartLocationY
                End With
            End If
            HallRS.MoveNext
        Wend
    End If
    
    frmLoading.lblStatus = "Loading Objects.."
    frmLoading.lblStatus.Refresh
    If ObjectRS.BOF = False Then
        ObjectRS.MoveFirst
        While ObjectRS.EOF = False
            A = ObjectRS!Number
            If A > 0 Then
                With Object(A)
                    .Name = ObjectRS!Name
                    .Description = CStr(IIf(IsNull(ObjectRS!Description), "", ObjectRS!Description))
                    .Picture = ObjectRS!Picture
                    .Type = ObjectRS!Type
                    .Flags = ObjectRS!Flags
                    .Data(0) = ObjectRS!Data1
                    .Data(1) = ObjectRS!Data2
                    .Data(2) = ObjectRS!Data3
                    .Data(3) = ObjectRS!Data4
                    .Data(4) = ObjectRS!Data5
                    If Not IsNull(ObjectRS!Data6) Then .Data(5) = ObjectRS!Data6
                    If Not IsNull(ObjectRS!Data7) Then .Data(6) = ObjectRS!Data7
                    If Not IsNull(ObjectRS!Data8) Then .Data(7) = ObjectRS!Data8
                    If Not IsNull(ObjectRS!Data9) Then .Data(8) = ObjectRS!Data9
                    If Not IsNull(ObjectRS!Data10) Then .Data(9) = ObjectRS!Data10

                    
                    .Class = IIf(IsNull(ObjectRS!Class), 0, ObjectRS!Class)
                    .MinLevel = ObjectRS!MinLevel
                    .Level = ObjectRS!Level
                    .EquipmentPicture = ObjectRS!EquipmentPicture
                End With
            End If
            ObjectRS.MoveNext
        Wend
    End If
    
    frmLoading.lblStatus = "Loading NPCs.."
    frmLoading.lblStatus.Refresh
    
    If NPCRS.BOF = False Then
        NPCRS.MoveFirst
        While NPCRS.EOF = False
            A = NPCRS!Number
            If A > 0 Then
                With NPC(A)
                    .Name = NPCRS!Name
                    .JoinText = NPCRS!JoinText
                    .LeaveText = NPCRS!LeaveText
                    For B = 0 To 4
                        .SayText(B) = NPCRS("SayText" + CStr(B))
                    Next B
                    For B = 0 To 9
                        With .SaleItem(B)
                            .GiveObject = NPCRS("GiveObject" + CStr(B))
                            .GiveValue = NPCRS("GiveValue" + CStr(B))
                            .TakeObject = NPCRS("TakeObject" + CStr(B))
                            .TakeValue = NPCRS("TakeValue" + CStr(B))
                        End With
                    Next B
                    .Flags = NPCRS!Flags
                    .Portrait = NPCRS!Portrait
                    .sprite = NPCRS!sprite
                    .direction = NPCRS!direction
                End With
            End If
            NPCRS.MoveNext
        Wend
    End If
    
    frmLoading.lblStatus = "Loading Monsters.."
    frmLoading.lblStatus.Refresh
    
    'Set Monster Defaults
    For A = 1 To MAXITEMS
        With monster(A)
            .Wander = 30
            .MoveSpeed = 2
            .AttackSound = 4
            .alpha = 255
            .red = 255
            .green = 255
            .blue = 255
        End With
    Next A
    
    If MonsterRS.BOF = False Then
        MonsterRS.MoveFirst
        While MonsterRS.EOF = False
            A = MonsterRS!Number
            If A > 0 Then
                With monster(A)
                    .Name = MonsterRS!Name
                    .sprite = MonsterRS!sprite
                    .HP = MonsterRS!HP
                    .Min = MonsterRS!Min
                    .Max = MonsterRS!Max
                    .Armor = MonsterRS!Armor
                    .Sight = MonsterRS!Sight
                    .Agility = MonsterRS!Agility
                    .Flags = MonsterRS!Flags
                    .Flags2 = MonsterRS!Flags2
                    .Object(0) = MonsterRS!Object0
                    .Value(0) = MonsterRS!Value0
                    .Chance(0) = MonsterRS!Chance0
                    .Object(1) = MonsterRS!Object1
                    .Value(1) = MonsterRS!Value1
                    .Chance(1) = MonsterRS!Chance1
                    .Object(2) = MonsterRS!Object2
                    .Value(2) = MonsterRS!Value2
                    .Chance(2) = MonsterRS!Chance2
                    .Experience = MonsterRS!Experience
                    .Level = MonsterRS!Level
                    .MagicResist = MonsterRS!MagicResist
                    .cStatusEffect = MonsterRS!cStatusEffect
                    .MonsterType = MonsterRS!MonsterType
                    .DeathSound = MonsterRS!DeathSound
                    .AttackSound = MonsterRS!AttackSound
                    .MoveSpeed = MonsterRS!MoveSpeed
                    .AttackSpeed = MonsterRS!AttackSpeed
                    .Wander = MonsterRS!Wander
                    .alpha = MonsterRS!alpha
                    .red = MonsterRS!red
                    .green = MonsterRS!green
                    .blue = MonsterRS!blue
                    .Light = MonsterRS!Light
                End With
            End If
            MonsterRS.MoveNext
        Wend
    End If
    
    On Error GoTo BanError
bans:
    frmLoading.lblStatus = "Loading Ban List.."
    frmLoading.lblStatus.Refresh
    If BanRS.BOF = False Then
        BanRS.MoveFirst
        While BanRS.EOF = False
            A = BanRS!Number
            If A > 0 Then
                With Ban(A)
                    .Name = BanRS!Name
                    .user = BanRS!user
                    .ip = IIf(IsNull(BanRS!ip), "0.0.0.0", BanRS!ip)
                    .reason = BanRS!reason
                    .UnbanDate = BanRS!UnbanDate
                    .uId = BanRS!uId
                    .iniUID = BanRS!iniUID
                    .Banner = IIf(IsNull(BanRS!Banner), "Script", BanRS!Banner)
                    .InUse = True
                End With
            End If
            BanRS.MoveNext
        Wend
    End If
    
    On Error GoTo 0
    
    frmLoading.lblStatus = "Loading Prefixs.."
    frmLoading.lblStatus.Refresh
    If PrefixRS.BOF = False Then
        PrefixRS.MoveFirst
        While PrefixRS.EOF = False
            A = PrefixRS!Number
            If A > 0 Then
                With prefix(A)
                    .Name = PrefixRS!Name
                    St = PrefixRS!Data
                    If Len(St) >= 10 Then
                        .ModType = Asc(Mid(St, 1, 1))
                        .Min = Asc(Mid(St, 2, 1))
                        .Max = Asc(Mid(St, 3, 1))
                        .Flags = Asc(Mid(St, 4, 1))
                        .Strength1 = Asc(Mid(St, 5, 1))
                        .Strength2 = Asc(Mid(St, 6, 1))
                        .Weakness1 = Asc(Mid(St, 7, 1))
                        .Weakness2 = Asc(Mid(St, 8, 1))
                        .Light.Intensity = Asc(Mid(St, 9, 1))
                        .Light.Radius = Asc(Mid(St, 10, 1))
                        If Len(St) = 10 Then
                            B = 1
                        ElseIf Len(St) >= 11 Then
                            B = Asc(Mid$(St, 11, 1))
                        End If
                        If B < 1 Then B = 1
                        If B > 10 Then B = 10
                        .Rarity = B
                    End If
                End With
            End If
            PrefixRS.MoveNext
        Wend
    End If
    
    GeneratePrefixList
    GenerateSuffixList
    
    frmLoading.lblStatus = "Loading Lights.."
    frmLoading.lblStatus.Refresh
    If LightsRS.BOF = False Then
        LightsRS.MoveFirst
        While LightsRS.EOF = False
            A = LightsRS!Number
            If A > 0 Then
                With Lights(A)
                    .Name = LightsRS!Name
                    St = LightsRS!Data
                    If Len(St) >= 7 Then
                        .red = Asc(Mid$(St, 1, 1))
                        .green = Asc(Mid$(St, 2, 1))
                        .blue = Asc(Mid$(St, 3, 1))
                        .Intensity = Asc(Mid$(St, 4, 1))
                        .Radius = Asc(Mid$(St, 5, 1))
                        .MaxFlicker = Asc(Mid$(St, 6, 1))
                        .FlickerRate = Asc(Mid$(St, 7, 1))
                    End If
                End With
            End If
            LightsRS.MoveNext
        Wend
    End If

    
    
    
    frmLoading.lblStatus = "Loading Maps.."
    frmLoading.lblStatus.Refresh
    If MapRS.BOF = False Then
        MapRS.MoveFirst
        While MapRS.EOF = False
            A = MapRS!Number
            If A > 0 Then
                LoadMap A, MapRS!Data
                ResetMap A
            End If
            MapRS.MoveNext
        Wend
    End If
    
    frmLoading.lblStatus = "Loading Scripts.."
    frmLoading.lblStatus.Refresh
    InitScriptTable
    
Exit Sub

DataRSError:
    bDataRSError = True
    Resume Next
    Exit Sub
BanError:
    BanRS.Close
    Set BanRS = Nothing
    db.TableDefs.Delete "Bans"
    CreateBansTable
    Set BanRS = db.TableDefs("Bans").OpenRecordset(dbOpenTable)
    BanRS.Index = "Number"
    MsgBox "Ban table reset"
    GoTo bans
    
End Sub
