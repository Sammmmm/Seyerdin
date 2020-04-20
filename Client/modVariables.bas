Attribute VB_Name = "modVariables"
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
'INI File Related


Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long


Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'Public Const VK_SHIFT = &H10
'Public Const VK_CONTROL = &H11
'Public Const VK_ALT = &H12
Public Const VK_PRINT = &H2A


Public Const MAXKEYCODES = 70
Public KeyCodeList(0 To MAXKEYCODES) As KeyCode

Public AmbientAlpha As Long
Public AmbientRed As Long
Public AmbientGreen As Long
Public AmbientBlue As Long
Public spamdelay As Byte
Public BanString As String

Public Const LT_NONE = 0
Public Const LT_TILE = 1
Public Const LT_CHARACTER = 2
Public Const LT_PLAYER = 3
Public Const LT_PROJECTILE = 4
Public Const LT_OBJECT = 5
Public Const LT_MONSTER = 6

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function ReleaseCapture Lib "user32" () As Long

'Hook
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Default Paths
Public Const PORTRAIT_DIR = "Data/Graphics/Portraits/"


Public Const MAXITEMS = 1000
Public Const MAXUSERS = 80
Public Const MaxTraps = 9


Public mapChangedBg(0 To 11, 0 To 11) As Boolean
Public mapChanged As Boolean
Public mapFGChanged As Boolean
Public LastBroadcast As String
Public chatForms(1 To 5) As frmChatWindow

Public scripts() As String

Public Type LightData
    Name As String
    Red As Byte
    Green As Byte
    Blue As Byte
    Intensity As Byte
    Radius As Byte
    MaxFlicker As Byte
    FlickerRate As Byte
End Type

Public Type SkillTreeData
    Skill As Byte
    Group As Long
    meetsReqs As Boolean
    Req1 As Long
    Req1Level As Long
    Req2 As Long
    Req2Level As Long
    LevelReq As Long
    drawX As Long
    drawY As Long
    used As Boolean
End Type
Type LightSource
    x As Long
    y As Long
    Radius As Long
    Intensity As Long
    Flicker As Byte
    MaxFlicker As Byte
    FlickerRate As Byte
    FlickerCount As Byte
    Type As Long
    Red As Byte
    Green As Byte
    Blue As Byte
    D As Byte
End Type

Public Type NPCSaleItemData
    GiveObject As Integer
    GiveValue As Long
    TakeObject As Integer
    TakeValue As Long
End Type

Public Type PrefixType
    Name As String
    ModType As Byte
    Flags As Byte
    Light As LightSource
End Type

Type PlayerData
    Name As String
    map As Long
    Sprite As Byte
    Status As Byte
    x As Long
    y As Long
    XO As Long
    YO As Long
    D As Byte
    A As Byte
    W As Long
    WalkStep As Long
    WalkStart As Long
    Guild As Byte
    Color As Long
    Ignore As Boolean
    Party As Integer
    HP As Integer
    Mana As Integer
    Level As Byte
    StatusEffect As Long
    buff As Byte
    Frame As Byte
    Light As LightSource
    LightSourceNumber As Byte
    Red As Byte
    Green As Byte
    Blue As Byte
    alpha As Byte
    equippedPicture(1 To 4) As Integer
End Type

Type MacroData
    Text As String
    LineFeed As Boolean
    Skill As Long
End Type

Type ObjectData
    Name As String
    Description As String
    Type As Byte
    Picture As Integer
    ObjData(0 To 9) As Long
    Flags As Byte
    MinLevel As Byte
    Class As Integer
    EquipmentPicture As Byte
End Type

Type MonsterData
    Name As String
    Sprite As Byte
    MaxHP As Integer
    Flags As Byte
    Flags2 As Byte
    DeathSound As Byte
    AttackSound As Byte
    alpha As Byte
    Red As Byte
    Green As Byte
    Blue As Byte
    Light As Byte
End Type

Type MapStartLocationData
    map As Integer
    x As Byte
    y As Byte
End Type

Type MapDoorData
    Att As Byte
    AttData(3) As Byte
    BGTile1 As Integer
    WallTile As Byte
    x As Long
    y As Long
End Type

Type TileData
    Ground As Integer
    Ground2 As Integer
    BGTile1 As Integer
    Anim(1 To 2) As Byte
    FGTile As Integer
    Att As Byte
    AttData(0 To 3) As Byte
    WallTile As Integer
End Type

Type TileAnimData
    Frame As Byte
    Frame2 As Byte
    AnimDelay As Byte
End Type

Type MapMonsterData
    Monster As Integer
    x As Long
    y As Long
    OX As Long
    OY As Long
    XO As Long
    YO As Long
    D As Byte
    A As Byte
    W As Byte
    HP As Integer
    WalkStep As Byte
    StartTick As Long 'Tick that movement should start at
    EndTick As Long 'Tick that movement should stop at
    R As Byte
    G As Byte
    b As Byte
    alpha As Byte
    LightSourceNumber As Byte
End Type

Type MapDeadBodyData
    MonNum As Integer
    Sprite As Byte
    BodyType As Byte
    Name As String
    Frame As Byte
    Event As Byte
    Counter As Long
    x As Byte
    y As Byte
End Type

Type MapObjectData
    Object As Integer
    x As Byte
    y As Byte
    Value As Long
    Prefix As Byte
    PrefixVal As Byte
    Suffix As Byte
    SuffixVal As Byte
    Affix As Byte
    AffixVal As Byte
    XOffset As Integer
    YOffset As Integer
    Light As LightSource
    TimeStamp As Long
    DeathObj As Boolean
    ObjectColor As Byte
End Type

Type MapMonsterSpawnData
    Monster As Integer
    Rate As Byte
End Type

Type MapTrapData
    x As Byte
    y As Byte
    Created As Long
    Counter As Byte
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
    Monster(0 To 9) As MapMonsterData
    DeadBody(0 To 9) As MapDeadBodyData
    MonsterSpawn(0 To 4) As MapMonsterSpawnData
    Door(0 To 9) As MapDoorData
    Trap(0 To MaxTraps) As MapTrapData
    BootLocation As MapStartLocationData
    Intensity As Byte
    MIDI As Byte
    Flags(1) As Byte
    Version As Long
    Raining As Integer
    Snowing As Integer
    Fog As Byte
    Zone As Byte
    Fish(1 To 10) As FishTileData
    SnowColor As Byte
    RainCOlor As Byte
End Type



Type InvObjData
    Object As Integer
    Value As Long
    Prefix As Byte
    PrefixValue As Byte
    Suffix As Byte
    SuffixValue As Byte
    Affix As Byte
    AffixValue As Byte
    ObjectColor As Byte
End Type

Type GuildData
    Name As String
    Hall As String
    AverageRenown As Long
    Kills As Long
    Deaths As Long
    members As Byte
    MembersOnline As Byte
    Symbol1 As Byte
    Symbol2 As Byte
    Symbol3 As Byte
    hallNum As Byte
End Type

Type HallData
    Name As String
End Type

Type NPCData
    Name As String
    Flags As Byte
    Status As Byte
    Sprite As Byte
    Portrait As Byte
    Direction As Byte
End Type

Type GuildDeclarationData
    Guild As Byte
    Type As Byte
End Type

Type TradeData
    player As Long
    Tradestate(0 To 2) As Byte
    Slot(1 To 10) As Byte
    YourObjects(1 To 10) As InvObjData
    TheirObjects(1 To 10) As InvObjData
    CurrentObject As Byte
End Type

Type CProjectile 'Players current attacking projectile
    Key As String
End Type

Type CharacterData
    Name As String
    Class As Byte
    Level As Byte
    Gender As Byte
    Sprite As Byte
    MacroSkill As Byte
    Counter As Byte
    Light As LightSource
    RadiusMod As Integer
    IntensityMod As Integer
    EnchantRadius As Byte
    EnchantIntensity As Byte
    
    AttackSpeed As Integer
    
    MaxHP As Integer
    MaxEnergy As Integer
    MaxMana As Integer
    HP As Integer
    Energy As Integer
    Mana As Integer
    
    strength As Byte
    StrengthMod As Integer
    Agility As Byte
    AgilityMod As Integer
    Endurance As Byte
    EnduranceMod As Integer
    Wisdom As Byte
    WisdomMod As Integer
    Constitution As Byte
    ConstitutionMod As Integer
    Intelligence As Byte
    IntelligenceMod As Integer
    Experience As Long
    StatPoints As Integer
    skillPoints As Integer
    SkillLevels(1 To 255) As Byte
    SkillEXP(1 To 255) As Long
    LocalSpellTick(1 To MAX_SKILLS) As Long
    GlobalSpellTick As Long
    LastCast As Long
    LastMapSwitch As Long
    
    Status As Byte
    Access As Byte
    Index As Byte
    Guild As Byte
    GuildRank As Byte
    Squelched As Byte
    StatusEffect As Long
    buff As Byte
    Frame As Byte
    Target As Byte
    TargType As Byte
    CurStoragePage As Byte
    NumStoragePages As Byte
    Red As Byte
    Green As Byte
    Blue As Byte
    alpha As Byte
    
    GuildDeclaration(0 To 9) As GuildDeclarationData
    
    desc As String
    Inv(1 To 20) As InvObjData
    Equipped(1 To 5) As InvObjData
    Storage(1 To 10, 1 To 20) As InvObjData
    
    CurProjectile As CProjectile
    Party As Integer
    Trading As Boolean
    
    Frozen As Boolean
    
    statEvasion As Byte
    statBaseAttack As Long
    statBlock As Long
    statPoisonResist As Byte
    statHpRegenLow As Byte
    statHpRegenHigh As Byte
    statManaRegenLow As Byte
    statManaRegenHigh As Byte
    statEnergyRegenLow As Byte
    statEnergyRegenHigh As Byte
    statCritical As Byte
    statMagicResist As Byte
    statMagicDefense As Byte
    statPhysicalDefense As Byte
    statPhysicalResist As Byte
    statAttackSpeed As Long
    statLeachHP As Byte
    statMagicFind As Byte
    statBodyArmor As Integer
    statHeadArmor As Integer
End Type

Type OptionsData
    HideBarsWhenFull As Boolean
    DisableManaBar As Boolean
    StatusDisplayMode As Byte
    
    MIDI As Boolean
    ShowHP As Boolean
    Wav As Boolean
    Broadcasts As Boolean
    Says As Boolean
    Tells As Boolean
    Yells As Boolean
    Emotes As Boolean
    Away As Boolean
    AwayMsg As String
    ForwardUser As String
    FontSize As Byte
    MName As Boolean
    UseFilter As Boolean
    LightQuality As Byte
    MusicVolume As Byte
    SoundVolume As Byte
    WalkSound As Boolean
    AutoRun As Boolean
    pausetime As Long
    highpriority As Boolean
    hightask As Boolean
    ShowFog As Boolean
    fullredraws As Boolean
    dontDisplayHelms As Boolean
    
    ResolutionIndex As Long
    TargetFps As Long
    DisableMultiSampling As Boolean
    VsyncEnabled As Boolean
    
    
    
    'AltKeysEnabled As Boolean
    UpKey As Byte
    DownKey As Byte
    LeftKey As Byte
    RightKey As Byte
    PickupKey As Byte
    AttackKey As Byte
    RunKey As Byte
    StrafeKey As Byte
    CycleKey As Byte
    
    BroadcastKey As Byte
    SayKey As Byte
    TellKey As Byte
    GuildKey As Byte
    PartyKey As Byte
    ChatKey As Byte
    
    SpellKey(0 To 9) As Byte
    
    
    
End Type

Public Const maxmod As Byte = 29
Public ModString(0 To maxmod) As String


Type ListBoxData
    Data(1 To 255) As Byte
    Caption(1 To 255) As String
    YOffset As Long
    Selected As Byte
    MouseState As Boolean
End Type

Public SkillListBox As ListBoxData

Type ItemInfo
    objNum As Integer
    ObjValue As Long
    Prefix As Byte
    PrefixVal As Byte
    PrefixType As Byte
    Suffix As Byte
    SuffixVal As Byte
    SuffixType As Byte
End Type

Type WorldData
    Hour As Byte
    Minute As Byte
    Day As Byte
    Rain As Integer
    Snow As Integer
    Fog As Byte
    FlickerDark As Byte
    FlickerLength As Byte
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
    Description As String
    MaleSprite As Byte
    FemaleSprite As Byte
    Enabled As Byte
End Type

Public Const MAX_CLASS = 10
Public Const MAX_ITEM_GRAPHICS = 350

Public Class(1 To MAX_CLASS) As ClassData

Public RecentTile(6) As Integer
Public StorageOpen As Boolean
Public fTime As Long
Public PlayerScanItem(1 To 30) As ItemInfo

Public Lights(255) As LightData
Public LightSource(0 To 40) As LightSource

Global TargetMonster As Byte
Global RedrawSkills As Boolean

Public World As WorldData

Public SkillEXPTable(0, 1 To 50) As Long

'Chat Variables
Type ChatData
    Text As String
    Color As Long
    Size As Long
    used As Boolean
    Channel As Byte
End Type

Type KeyCode
    Text As String
    KeyCode As Byte
    CapitalKeyCode As Byte
    NotChatKey As Boolean
End Type


Type ChatTypeData
    'AllChat(0 To 255, 0 To 0) As ChatData
    AllChat(0 To 1000) As ChatData
    ChatIndex As Integer
    Enabled As Boolean
    'AllChatPos As Long

    'QuestChat(0 To 255) As ChatData
    'QuestChatPos As Long
End Type

Public Chat As ChatTypeData

'Font Data
Type FontCharData
    srcX As Long
    srcY As Long
    Width As Long
End Type

'Big Fonts
Type UnzTextData
    x As Long
    y As Long
    Lifetime As Long
    Fade As Long
    Text As String
End Type

Type String3D
    Text As String
    Width As Long
    Height As Long
End Type

Public Const FontHeight As Long = 50
Public FontChar(1 To 2, 32 To 128) As FontCharData
Public NumUnzText As Long
Public UnzText() As UnzTextData

'Map Flag Constants
Public Const MAP_FRIENDLY = 1 'bit 1
Public Const MAP_INDOORS = 2 'bit 2
Public Const MAP_RAINING = 1 'bit 1 on flag 2
Public Const MAP_SNOWING = 2 'bit 2 on flag 2
Public Const MAP_TILEFLAGS = 8 'bit 4 on flag 2

'Monster Flag Constants
Public Const MONSTER_GUARD = 1
Public Const MONSTER_DAY = 2
Public Const MONSTER_NIGHT = 4
Public Const MONSTER_FRIENDLY = 8
Public Const MONSTER_ATTACK_MONSTER = 16
Public Const MONSTER_NO_SHADOW = 32
Public Const MONSTER_SEE_INVISIBLE = 64

'Mflag2
Public Const MONSTER_LARGE = 1
Public Const MONSTER_SPRITE255 = 2
Public Const MONSTER_MEDIUM = 4
'NPC Flag Constants
Public Const NPC_BANK = 1
Public Const NPC_REPAIR = 2
Public Const NPC_SHOP = 4

'NPC Status Colors
Public NPCStatusColors(1 To 5) As Long

'Data about last NPC interacted with
Public LastNPCX As Long
Public LastNPCY As Long

Type Point
    x As Long
    y As Long
End Type

'Mapper Undo Variables
Public Type MapAction
    x As Long
    y As Long
    Layer As Long
    OldTile As Long
    AttData(0 To 3) As Byte
End Type

Public MapUndo(250) As MapAction
Public CurrentMapAction As Long
'Public tickCountMod As Long

'Inventory Drag and Drop Variables
Public StartInvX As Long
Public StartInvY As Long
Public OldInvX As Long
Public OldInvY As Long
Public DraggingObj As Boolean
Public ReplaceCursor As Boolean

'Trade Variables
Public TradeData As TradeData
Public Const TRADE_STATE_INVITED As Long = 1
Public Const TRADE_STATE_OPEN As Long = 2
Public Const TRADE_STATE_ACCEPTED As Long = 3

Public EXPLevel(255) As Long

Public LastPlayerTellName As String
Public LastPlayerTellNum As Long

Public Function ToRect(x As Long, y As Long, Width As Long, Height As Long) As RECT
    Dim R1 As RECT
    R1.Left = x
    R1.Right = x + Width
    R1.Top = y
    R1.Bottom = y + Height
    ToRect = R1
End Function


Public Function Create3DString(St As String) As String3D
    Dim A As Long, CurWidth As Long, St1 As String
    Create3DString.Text = St
    Create3DString.Height = 16
    For A = 1 To Len(St)
        St1 = Mid$(St, A, 1)
        If Asc(St1) >= 32 And Asc(St1) <= 127 Then
            CurWidth = CurWidth + FontChar(2, Asc(St1)).Width
        End If
    Next A
    Create3DString.Width = CurWidth
End Function


Public Function Get3DFontWidth(St As String) As Long
    Dim A As Long, b As Long
    If Len(St) > 0 Then
        For A = 1 To Len(St)
            b = b + FontChar(2, Asc(Mid$(St, A, 1))).Width
        Next A
    End If
    Get3DFontWidth = b
End Function
