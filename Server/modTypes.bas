Attribute VB_Name = "modTypes"
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

Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

'Type NOTIFYICONDATA
'cbSize As Long
'hwnd As Long
'uId As Long
'uFlags As Long
'ucallbackMessage As Long
'hIcon As Long
'szTip As String * 64
'End Type

'Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)

Public Const TargTypePlayer = 1
Public Const TargTypeMonster = 2
Public Const TargTypeCharacter = 3

Public Const critsize = 15

'Global nid As NOTIFYICONDATA

    
'World constants - No need to store in database.
Public GLOBAL_OBJECT_RESET_RATE As Long
Public GLOBAL_DEATH_DROP_RESET_RATE As Long
Public GLOBAL_MAGIC_DROP_RATE  As Byte
Public Const MAXITEMS = 1000
Public Const MAXSTATUS = 37
Public Const STORAGEPAGES = 10
Public MaxLevel As Byte
Public DeathDropItemsLevel As Byte
Public MaxUsers As Byte
Public SkillsPerLevel As Byte
Public StatsPerLevel As Byte

Public StatRate1 As Byte
Public StatRate2 As Byte

Public baseEnergyRegen As Byte
Public baseManaRegen As Byte
Public BaseHPRegen As Byte
Public LevelsPerHpRegen As Byte
Public LevelsPerManaRegen As Byte

Public StrengthPerDamage(1 To 3) As Byte

Public AgilityPerCritChance(1 To 3) As Byte
Public AgilityPerDodgeChance(1 To 3) As Byte

Public EndurancePerBlockChance(1 To 3) As Byte
Public EndurancePerEnergy(1 To 3) As Byte
Public EndurancePerEnergyRegen(1 To 3) As Byte

Public ManaPerIntelligence(1 To 3) As Byte
Public IntelligencePerManaRegen(1 To 3) As Byte

Public ConstitutionPerHPRegen(1 To 3) As Byte
Public HPPerConstitution(1 To 3) As Byte

Public PietyPerMagicResist(1 To 3) As Byte
Public PietyPerHPRegen(1 To 3) As Byte
Public PietyPerManaRegen(1 To 3) As Byte
Public PietyPerHP(1 To 3) As Byte
Public PietyPerMana(1 To 3) As Byte
Public PietyPerDodge(1 To 3) As Byte
Public PietyPerCrit(1 To 3) As Byte
Public PietyPerBlock(1 To 3) As Byte

Public GenericStatPerBonus(1 To 3) As Byte
Public GenericPietyPerBonus(1 To 3) As Byte


Public Const MaxTraps = 9

Public CloseSocketQueue() As SocketQueueData

Type LightSource
    Radius As Byte
    Intensity As Byte
End Type

Type ScriptData
    Name As String
    source As String
    MCODE() As Byte
End Type

Type MapStartLocationData
    map As Integer
    x As Byte
    y As Byte
End Type

Type HallData
    Name As String
    Price As Long
    Upkeep As Long
    StartLocation As MapStartLocationData
End Type

Type GuildDeclarationData
    Guild As Byte
    Type As Byte
End Type

Type GuildMemberdata
    Name As String
    Rank As Byte
    Kills As Integer
    Deaths As Integer
    Renown As Long
    Joined As Date
End Type

Type GuildData
    Name As String
    Member(0 To 19) As GuildMemberdata
    Declaration(0 To 9) As GuildDeclarationData
    Hall As Byte
    Bank As Long
    sprite As Byte
    DueDate As Long
    Symbol1 As Byte
    Symbol2 As Byte
    MOTD As String
    Info As String
    Symbol3 As Byte
    Bookmark As Variant
    founded As Date
    AverageRenown As Long
    MemberCount As Byte
End Type

Type BanData
    Name As String
    reason As String
    user As String
    UnbanDate As Long
    Banner As String
    InUse As Boolean
    ip As String
    uId As String
    iniUID As String
End Type

Type FlagData
    Value As Long
    ResetCounter As Long
End Type

'Trade Constants
    Public Const TRADE_STATE_INVITER = 1
    Public Const TRADE_STATE_INVITED = 2
    Public Const TRADE_STATE_OPEN = 3
    Public Const TRADE_STATE_ACCEPTED = 4
    Public Const TRADE_STATE_FINISHED = 5
    
Type TradeData
    Trader As Long
    State As Byte
    Slot(1 To 10) As Byte
    Item(1 To 10) As InvObject
End Type

Type StatusData
    timer As Long
    Data(3) As Byte
End Type

Type BuffData
    Type As Byte
    timer As Long
    Data(1) As Long
End Type

Type Widget
    Key As String           'The name of this Widget
    Type As Long            'Type of Widget
    lngData As String       'Any integer value associated with this widget, returned from client
    strData As String       'Any String value associated with this widget, returned from client
    Data(2) As Long         'Any options set by server, MaxLen, Etc
    Flags As Long           'Style Flags set by server, such as BIG, SMALL, for size of buttons .. etc
End Type

Type WidgetData
    NumWidgets As Long      'Number of widgets currently displayed
    Widgets() As Widget     'Array that holds individual widgets, resized at run time
    
    WidgetString As String  'Serialized widget data, built by script commands until ready to send
    WidgetScript As String  'Script called by widget when any button is pressed
    MenuVisible As Boolean  'If a menu is currently being displayed
End Type

'Player Constants/Types
'--------------------------------------------------------------
'               Constants
'--------------------------------------------------------------
Public Const MaxProjectiles = 2 '1-MaxProjectiles
'--------------------------------------------------------------
'               Types
'--------------------------------------------------------------

Type PlayerData
    'Socket Data
    CalculateStats As Boolean
    DeferSends As Boolean
    ScriptUpdateMap As Boolean
    ShieldBlock As Boolean
    Socket As Long
    SocketData As String
    St As String
    ip As String
    ClientVer As String
    InUse As Boolean
    Mode As Byte
    LastMsg As Double
    Leaving As Boolean
    TickCount As Double
    PacketsSent As Byte
    uId As String
    iniUID As String
    'Account Data
    user As String
    Access As Byte
    speedhack As Byte
    moveQueue(0 To 4) As String
    
    walkStamp As Long
    deathStamp As Long
    'Trade Data
    Trade As TradeData
    Trading As Boolean
    'Character Data
    CharNum As Long
    Name As String
    Class As Byte
    Gender As Byte
    sprite As Byte
    'desc As String
    Squelched As Byte
    
    Buff As BuffData
    StatusEffect As Long
    StatusData(1 To MAXSTATUS) As StatusData
    AttackSkill As Byte
    
    Killer As Byte
    KillerPK As Boolean
    
    ProjectileType As Long
    ProjectileDamage As Long
    ProjectileX As Long
    ProjectileY As Long
    
    trapID As Long
    
    combatTimer As Long
    'Position Data
    map As Integer
    x As Byte
    y As Byte
    D As Byte
    WalkCode As Byte
    Light As LightSource
    RadiusMod As Integer
    IntensityMod As Integer
    
    'Vital Stat Data
    MaxHP As Long
    OldHP As Long
    MaxEnergy As Integer
    OldEnergy As Integer
    MaxMana As Integer
    OldMana As Integer
    HP As Integer
    HPRegen As Integer
    Energy As Integer
    Mana As Integer
    ManaRegen As Integer
    
    'Physical Stat Data
    strength As Integer
    OldStrength As Byte
    StrMod(1) As Byte
    Agility As Integer
    OldAgility As Byte
    AgiMod(1) As Byte
    Endurance As Integer
    OldEndurance As Byte
    EndMod(1) As Byte
    Wisdom As Integer
    OldWisdom As Byte
    WisMod(1) As Byte
    Constitution As Integer
    OldConstitution As Byte
    ConMod(1) As Byte
    Intelligence As Integer
    OldIntelligence As Byte
    IntMod(1) As Byte
    Level As Byte
    Experience As Long
    StatPoints As Integer
    Renown As Long
    
    'Misc. Data
    alpha As Byte
    red As Byte
    green As Byte
    blue As Byte
    
    SkillLevel(1 To 255) As Byte
    SkillEXP(1 To 255) As Long
    SkillPoints As Integer
    LocalSpellTick(MAX_SKILLS) As Long
    GlobalSpellTick As Long
    Status As Long
    MagicResist As Integer
    MagicBonus As Byte 'magic find, terrible name
    Leech As Integer
    CriticalBonus As Byte
    ResistPoison As Byte
    
    AttackDamageBonus As Byte
    MagicDamageBonus As Byte
    MagicDefenseBonus As Byte
    PhysicalDefenseBonus As Byte
    PhysicalResistanceBonus As Byte
    ShieldBlockBonus As Byte
    DodgeBonus As Byte
    
    Bank As Long
    
    ScriptTimer(1 To MaxPlayerTimers) As Double
    Script(1 To MaxPlayerTimers) As String
    ScriptCallback As String
    
    Party As Integer
    IParty As Integer
    
    'Guild Data
    Guild As Byte
    GuildRank As Byte
    
    JoinRequest As Byte
    
    'Inventory Data
    Inv(1 To 20) As InvObject
    NumStoragePages As Byte
    CurStoragePage As Byte
    Storage(1 To STORAGEPAGES, 1 To 20) As InvObject
    Equipped(1 To 5) As InvObject
    
    'Flag Data
    Flag(0 To 255) As FlagData
    
    FloodTimer As Long
    SquelchTimer As Long
    LoginStamp As Long
    
    
    AttackSpeed As Integer
    AttackCount As Double
    Frozen As Byte
    
    
    'Target Data
    CurrentRepairTar As Integer
    
    Widgets As WidgetData           'Holds everything to do with widgets
    
    'Database Data
    Bookmark As Variant
End Type

Type ClassData
    Name As String
    StartHP As Integer
    StartEnergy As Integer
    StartMana As Integer
    StartStrength As Byte
    StartAgility As Byte
    StartEndurance As Byte
    StartWisdom As Byte
    StartConstitution As Byte
    StartIntelligence As Byte
    HPIncrement As Byte
    ManaIncrement As Byte
    Enabled As Byte
End Type

Type ObjectData
    Name As String
    Description As String
    Picture As Integer
    Type As Byte
    Data(0 To 9) As Byte
    Flags As Byte
    Class As Integer
    MinLevel As Byte
    Level As Byte
    EquipmentPicture As Byte
End Type

Type MonsterData
    Name As String
    Description As String
    sprite As Byte
    Flags As Byte
    Flags2 As Byte
    HP As Integer
    Min As Integer
    Max As Integer
    Armor As Byte
    Sight As Byte
    Agility As Byte
    Object(0 To 2) As Integer
    Value(0 To 2) As Long
    Chance(0 To 2) As Byte
    Experience As Integer
    Level As Byte
    MagicResist As Integer
    cStatusEffect As Long
    Effect As Long
    MonsterType As Long
    DeathSound As Byte
    AttackSound As Byte
    MoveSpeed As Byte
    AttackSpeed As Byte
    Wander As Byte

    'Color
    alpha As Byte
    red As Byte
    green As Byte
    blue As Byte
    Light As Byte
End Type

Type NPCSaleItemData
    GiveObject As Integer
    GiveValue As Long
    TakeObject As Integer
    TakeValue As Long
End Type

Type NPCData
    Name As String
    JoinText As String
    LeaveText As String
    LastSay As Byte
    SayText(0 To 4) As String
    SaleItem(0 To 9) As NPCSaleItemData
    Flags As Byte
    Portrait As Byte
    sprite As Byte
    direction As Byte
End Type

Type TileData
    Att As Byte
    AttData(0 To 3) As Byte
    TimeStamp As Long
    WallTile As Byte
    Ground As Long
    
    Ground2 As Integer
    BGTile1 As Integer
    Anim(1 To 2) As Byte
    FGTile As Integer
End Type

Type MapDoorData
    Att As Byte
    x As Byte
    y As Byte
    t As Double
    Wall As Byte
    Used As Boolean
End Type

Type MapObjectData
    Object As Integer
    Value As Long
    TimeStamp As Double
    x As Byte
    y As Byte
    prefix As Byte
    prefixVal As Byte
    suffix As Byte
    SuffixVal As Byte
    Affix As Byte
    AffixVal As Byte
    ObjectColor As Byte
    Flags(0 To 3) As Long
    
    deathObj As Boolean
End Type

Type MapMonsterSpawnData
    monster As Integer
    Rate As Byte
End Type

    

'Monster Queue
Public Const QUEUE_EMPTY = 0
Public Const QUEUE_MOVE = 1
Public Const QUEUE_SCRIPT = 2
Public Const QUEUE_PAUSE = 3
Public Const QUEUE_SHIFT = 4
Public Const QUEUE_TURN = 5
Public Type MonsterQueue
    Action As Long
    lngData As Long
    lngData1 As Long
    strData As String
End Type

Type MapMonsterData
    monster As Integer
    x As Byte
    y As Byte
    D As Byte
    Target As Byte
    TargType As Byte
    Distance As Byte
    HP As Integer
    MoveSpeed As Byte
    AttackSpeed As Byte
    MoveCounter As Byte
    AttackCounter As Byte
    R As Byte
    G As Byte
    B As Byte
    A As Byte
    Frozen As Boolean
    Flags(0 To 5) As Long
    Poison As Byte
    PoisonLength As Long
    'Monster Queues
    CurrentQueue As Long
    MonsterQueue(0 To 15) As MonsterQueue
End Type

Type RainData
    Raining As Integer
    R As Byte
    G As Byte
    B As Byte
    FlashChance As Byte
    FlashLength As Byte
End Type

Type MapTrapData
    x As Byte
    y As Byte
    Type As Byte
    strength As Long
    ActiveCounter As Long
    CreatedTime As Long
    player As Long
    trapID As Long
End Type

Type MapProjData
    sprite As Byte
    speed As Byte
    startTime As Long '20 ms delay
    startX As Byte
    startY As Byte
    direction As Byte
    damage As Long
    magical As Byte
    damageString As String
End Type

Type FishTileData
    x As Byte
    y As Byte
    TimeStamp As Long
End Type

Type MapData
    Name As String
    ExitUp As Integer
    ExitDown As Integer
    ExitLeft As Integer
    ExitRight As Integer
    Tile(0 To 11, 0 To 11) As TileData
    Object(0 To 49) As MapObjectData
    monster(0 To 9) As MapMonsterData
    MonsterSpawn(0 To 4) As MapMonsterSpawnData
    Door(0 To 9) As MapDoorData
    trap(0 To MaxTraps) As MapTrapData
    Projectile(1 To 30) As MapProjData
    BootLocation As MapStartLocationData
    Flags(1) As Byte
    NumPlayers As Long
    ResetTimer As Double
    Hall As Byte
    NPC As Byte
    Keep As Boolean
    Version As Long
    CheckSum As Long
    Raining As RainData
    Snowing As Integer
    Fog As Byte
    FlickerDark As Byte
    FlickerLength As Byte
    Zone As Byte
    FishCounter As Byte
    Fish(1 To 10) As FishTileData
End Type

Type StartLocationData
    x As Byte
    y As Byte
    map As Integer
    Message As String
End Type

Type WorldData
    LastUpdate As Long
    MapResetTime As Double
    TimeAllowance As Long
    BackupInterval As Long
    CharCounter As Long
    MsgCounter As Long
    MoneyObj As Integer
    StartLocation(0 To 4) As StartLocationData
    MOTD As String
    Flag(0 To 255) As Long
    StringFlag(0 To 255) As String
    PlayerFlagCounter(0 To 255) As Long
    StartObjects(1 To 8) As Integer
    StartObjValues(1 To 8) As Long
    
    Hour As Long
    HourCounter As Long
End Type

Public NumPrefix(1 To 10) As Long
Public SortedPrefixList(1 To 10, 1 To 50)
Public NumSuffix(1 To 10) As Long
Public SortedSuffixList(1 To 10, 1 To 50)

Type prefix
    Name As String
    ModType As Byte
    Min As Byte
    Max As Byte
    Flags As Byte
    Strength1 As Byte
    Strength2 As Byte
    Weakness1 As Byte
    Weakness2 As Byte
    Light As LightSource
    Rarity As Byte
End Type

Type LightsData
    Name As String
    red As Byte
    green As Byte
    blue As Byte
    Intensity As Byte
    Radius As Byte
    MaxFlicker As Byte
    FlickerRate As Byte
End Type

Type Point
    x As Long
    y As Long
End Type

'Timers
Type uTimerData
    Script As String
    Parm(3) As Long
    InUse As Boolean
    timer As Long
End Type

Public uTimers(1 To 100) As uTimerData

Public Const MAX_CLASS = 10
Public Const MAX_ITEM_GRAPHICS = 350

Public World As WorldData
Public map(1 To 5000) As MapData
Public Guild(1 To 255) As GuildData
Public Hall(1 To 255) As HallData
Public Object(1 To MAXITEMS) As ObjectData
Public monster(1 To MAXITEMS) As MonsterData
Public NPC(1 To 255) As NPCData
Public player() As PlayerData
Public Lights(1 To 255) As LightsData

Public CurrentMaps() As Long


Public Class(1 To MAX_CLASS) As ClassData
Public Ban(1 To 50) As BanData
Public prefix(1 To 255) As prefix

Public Const MaxMod As Byte = 1
Public ModString(0 To MaxMod) As String



'MAP#_X_Y CONSTANTS
Public Const MC_CLICK = 0
Public Const MC_ATTACK = 1
Public Const MC_SKILL = 2
Public Const MC_WALK = 3

'FLOATING TEXT CONSTANTS
Public Const FT_MISS = 1
Public Const FT_INEFFECTIVE = 2
Public Const FT_ENDED = 3
Public Const FT_POISON = 4
Public Const FT_EXHAUST = 5
Public Const FT_VEX = 6
Public Const FT_BLOCK = 7

'MONSTER TYPE CONSTANTS
Public Const MT_FIRE = 1
Public Const MT_ELECTRICITY = 2
Public Const MT_UNDEAD = 4

Public SkillEXPTable(0, 1 To 50) As Long

'Monster Flag Constants
Public Const MONSTER_GUARD = 1
Public Const MONSTER_DAY = 2
Public Const MONSTER_NIGHT = 4
Public Const MONSTER_FRIENDLY = 8
Public Const MONSTER_ATTACK_MONSTER = 16
Public Const MONSTER_NO_SHADOW = 32
Public Const MONSTER_SEE_INVISIBLE = 64
Public Const MONSTER_TICK = 128

'Monster Flag2 Constants
Public Const MONSTER_LARGE = 1
'Public Const MONSTER_DAY = 2
'Public Const MONSTER_NIGHT = 4
'Public Const MONSTER_FRIENDLY = 8
'Public Const MONSTER_ATTACK_MONSTER = 16
'Public Const MONSTER_NO_SHADOW = 32
'Public Const MONSTER_SEE_INVISIBLE = 64
'Public Const MONSTER_TICK = 128


'NPC Flag Constants
Public Const NPC_BANK = 1
Public Const NPC_REPAIR = 2
Public Const NPC_SHOP = 4

'EXP Level Table
Public EXPLevel(255) As Long


Type IpSquelchData
    ip As String
    Time As Byte
End Type
Public ipSquelches(255) As IpSquelchData

Type EditTileData
    Ground As Integer
    Ground2 As Integer
    BGTile1 As Integer
    Anim(1 To 2) As Byte
    FGTile As Integer
    Att As Byte
    AttData(0 To 3) As Byte
    WallTile As Integer
End Type

Type EditMapData
    Name As String
    ExitUp As Integer
    ExitDown As Integer
    ExitLeft As Integer
    ExitRight As Integer
    Tile(0 To 11, 0 To 11) As EditTileData
    MonsterSpawn(0 To 4) As MapMonsterSpawnData
    BootLocation As MapStartLocationData
    Intensity As Byte
    NPC As Byte
    MIDI As Byte
    Flags(1) As Byte
    Version As Long
    Num As Long
    Raining As Integer
    Snowing As Integer
    Fog As Byte
    Zone As Byte
    RainColor As Byte
    SnowColor As Byte
End Type

Public CurEditMap As EditMapData

'Widget Types
    Public Const WIDGET_BUTTON = 1
    Public Const WIDGET_LABEL = 2
    Public Const WIDGET_FRAME = 3
    Public Const WIDGET_TEXTBOX = 4
    Public Const WIDGET_IMAGE = 5
'Widget Style Flags
    Public Const STYLE_NOSTYLE = 0
    Public Const STYLE_SMALL = 1
    Public Const STYLE_MEDIUM = 2
    Public Const STYLE_LARGE = 4
    Public Const STYLE_DYNAMIC = 8
    Public Const STYLE_CENTERED = 16
    Public Const STYLE_SOLID = 32
    Public Const STYLE_MULTILINE = 64

'Item Constants
    Public Const EQ_WEAPON = 1
    Public Const EQ_AMMO = 2
    Public Const EQ_DUALWIELD = 2
    Public Const EQ_SHIELD = 2
    Public Const EQ_ARMOR = 3
    Public Const EQ_HELMET = 4
    Public Const EQ_RING = 5
    
'Projectile COnstants
    Public Const PT_NONE = 0
    Public Const PT_PHYSICAL = 1
    Public Const PT_MAGIC = 2
    
'Attack Type Constants
    Public Const AT_MELEE = 1
    Public Const AT_MAGIC = 2
    Public Const AT_PROJECTILE_PHYSICAL = 3
    Public Const AT_PROJECTILE_MAGIC = 4
    Public Const AT_TRAP = 5
    
    Public startTime As Double
    Public EndTime As Double

