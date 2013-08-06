Attribute VB_Name = "modTypes"
Option Explicit
Public Const MAX_QUESTS = 500
Public Const MAX_PARTIES = 500
' ---NOTE to future developers!----------
' When loading, types ARE order-sensitive!
' This means do not change the order of variables in between
' versions, and add new variables to the end. This way, we can
' just load the old files! I learned that the hard way :D
' -Pickle

Public Const MAX_PARTY_MEMBERS = 4
Public Const MAX_PARTY_INV_SLOTS = 10
Public Party(1 To MAX_PARTIES) As PartyRec

Type PlayerInvRec
    num As Integer
    Value As Long
    Dur As Integer
End Type

Type BankRec
    num As Integer
    Value As Long
    Dur As Integer
End Type

Type ElementRec
    Name As String * NAME_LENGTH
    Strong As Integer
    Weak As Integer
End Type

Type PetRec
    SPRITE As Long
    Alive As Byte
    Map As Long
    x As Long
    y As Long
    Dir As Byte
    HP As Long
    SP As Long
    MP As Long
    FP As Long
    MAXHP As Long
    MAXSP As Long
    MAXMP As Long
    MaxFP As Long
    MapToGo As Long
    XToGo As Long
    YToGo As Long
    Target As Long
    TargetType As Byte
    AttackTimer As Long
    Level As Long
    SpriteSet As Long
    STR As Long
    DEF As Long
    Speed As Long
    Magi As Long
    POINTS As Long
    EXP As Long
    Name As String
End Type

Type V000PlayerRec
    ' General
    Name As String * NAME_LENGTH
    Guild As String
    GuildAccess As Byte
    Sex As Byte
    Class As Integer
    SPRITE As Long
    Level As Integer
    EXP As Long
    Access As Byte
    PK As Byte

    ' Vitals
    HP As Long
    MP As Long
    SP As Long

    ' Stats
    STR As Long
    DEF As Long
    Speed As Long
    Magi As Long
    POINTS As Long

    ' Worn equipment
    ArmorSlot As Integer
    WeaponSlot As Integer
    HelmetSlot As Integer
    ShieldSlot As Integer
    LegsSlot As Integer
    RingSlot As Integer
    NecklaceSlot As Integer

    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Integer
    Bank(1 To MAX_BANK) As BankRec

    ' Position and movement
    Map As Integer
    x As Byte
    y As Byte
    Dir As Byte

    TargetNPC As Integer

    Head As Integer
    Body As Integer
    Leg As Integer

    paperdoll As Byte

    MAXHP As Long
    MAXMP As Long
    MAXSP As Long
    
    NpcKillType As String
    NpcKillamount As Long
    NpcKillQuestFlag As Long
    QuestFlags(1 To MAX_QUESTS) As Long
    
    PetSprite As Long
    PetAlive As Byte
    PetMap As Long
    PetX As Long
    PetY As Long
    PetDIR As Byte
    PetHP As Long
    PetSP As Long
    PetMP As Long
    PetFP As Long
    PetMaxHP As Long
    PetMaxSP As Long
    PetMaxMP As Long
    PetMaxFP As Long
    PetMapToGo As Long
    PetLevel As Long
    PetSpriteSet As Long
    PetSTR As Long
    PetDEF As Long
    PetSPEED As Long
    PetMAGI As Long
    PetPOINTS As Long
    PetEXP As Long
    PetNAME As String
    PetTNL As Long
    
    InParty As Byte
    LookingForParty As Byte
    PartyInvitedTo As Byte
    PartyInvitedToBy As String
    Party As Byte
End Type

'Nuevo Type para arreglar desincronización
Public Type PlayerRec
    ' General
'090829 Scorpious2k
    Vflag As Byte       ' version flag - always > 127
    Ver As Byte
    SubVer As Byte
    Rel As Byte
'090829 End
    Name As String * NAME_LENGTH
    Guild As String
    GuildAccess As Byte
    Sex As Byte
    Class As Integer
    SPRITE As Long
    Level As Integer
    EXP As Long
    Access As Byte
    PK As Byte

    ' Vitals
    HP As Long
    MP As Long
    SP As Long

    ' Stats
    STR As Long
    DEF As Long
    Speed As Long
    Magi As Long
    POINTS As Long

    ' Worn equipment
    ArmorSlot As Integer
    WeaponSlot As Integer
    HelmetSlot As Integer
    ShieldSlot As Integer
    LegsSlot As Integer
    RingSlot As Integer
    NecklaceSlot As Integer


    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Integer
    Bank(1 To MAX_BANK) As BankRec

    ' Position and movement
    Map As Integer
'090829    X As Byte
'090829    Y As Byte
    x As Integer
    y As Integer
    Dir As Byte

    TargetNPC As Integer

    Head As Integer
    Body As Integer
    Leg As Integer

    paperdoll As Byte

    MAXHP As Long
    MAXMP As Long
    MAXSP As Long
    
    NpcKillType As String
    NpcKillamount As Long
    NpcKillQuestFlag As Long
    QuestFlags(1 To MAX_QUESTS) As Long
    
    PetSprite As Long
    PetAlive As Byte
    PetMap As Long
    PetX As Long
    PetY As Long
    PetDIR As Byte
    PetHP As Long
    PetSP As Long
    PetMP As Long
    PetFP As Long
    PetMaxHP As Long
    PetMaxSP As Long
    PetMaxMP As Long
    PetMaxFP As Long
    PetMapToGo As Long
    PetLevel As Long
    PetSpriteSet As Long
    PetSTR As Long
    PetDEF As Long
    PetSPEED As Long
    PetMAGI As Long
    PetPOINTS As Long
    PetEXP As Long
    PetNAME As String
    PetTNL As Long
    
    InParty As Byte
    LookingForParty As Byte
    PartyInvitedTo As Byte
    PartyInvitedToBy As String
    Party As Byte
    color(1 To 3) As Byte
End Type

Type PlayerTradeRec
    InvNum As Long
    InvName As String
End Type

Public Type AccountRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
    Email As String

    ' Some error here that needs to be fixed. [Mellowz]
    Char(0 To MAX_CHARS) As PlayerRec

    ' None saved local vars
    Buffer As String
    IncBuffer As String
    CharNum As Byte
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    
    PartyID As Long
    InParty As Byte
    PartyPlayer As Long
    PartyStarter As Byte
    Invited As Long
    TargetType As Byte

    Target As Byte
    CastedSpell As Byte

    SpellTime As Long
    SpellVar As Long
    SpellDone As Long
    spellnum As Long

    GettingMap As Byte
    InvitedBy As Byte

    Emoticon As Long

    InTrade As Boolean
    TradePlayer As Long
    TradeOk As Byte
    TradeItemMax As Byte
    TradeItemMax2 As Byte
    Trading(1 To MAX_PLAYER_TRADES) As PlayerTradeRec

    InChat As Byte
    ChatPlayer As Long

    Mute As Boolean
    Locked As Boolean
    LockedSpells As Boolean
    LockedItems As Boolean
    LockedAttack As Boolean
    TargetNPC As Long

    HookShotX As Byte
    HookShotY As Byte

    ' MENUS
    CustomMsg As String
    CustomTitle As String
    Pet As PetRec
End Type

Type TileRec
    Ground As Long
    Mask As Long
    Anim As Long
    Mask2 As Long
    M2Anim As Long
    Fringe As Long
    FAnim As Long
    Fringe2 As Long
    F2Anim As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    String1 As String
    String2 As String
    String3 As String
    Light As Long
    GroundSet As Byte
    MaskSet As Byte
    AnimSet As Byte
    Mask2Set As Byte
    M2AnimSet As Byte
    FringeSet As Byte
    FAnimSet As Byte
    Fringe2Set As Byte
    F2AnimSet As Byte
End Type

Type PartyRec
Leader As Byte
Member(1 To MAX_PARTY_MEMBERS) As Byte
Created As Boolean
TimeCreated As Double
End Type

Type MapRec
    Name As String * 20
    Revision As Integer
    Moral As Byte
    Up As Integer
    Down As Integer
    Left As Integer
    Right As Integer
    music As String
    BootMap As Integer
    BootX As Byte
    BootY As Byte
    Shop As Integer
    Indoors As Byte
    Tile() As TileRec
    NPC(1 To 25) As Integer
    SpawnX(1 To 25) As Byte
    SpawnY(1 To 25) As Byte
    Owner As String
    Scrolling As Byte
    Weather As Integer
End Type

Type ClassRec
    Name As String * NAME_LENGTH

    AdvanceFrom As Long
    LevelReq As Long
    Type As Long
    Locked As Long

    MaleSprite As Long
    FemaleSprite As Long

    STR As Long
    DEF As Long
    Speed As Long
    Magi As Long

    Map As Long
    x As Byte
    y As Byte

    ' Description
    Desc As String
    
    'Genre
    Gender As Byte
    Gender1 As String
    Gender2 As String
End Type

Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 150

    Pic As Long
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    StrReq As Long
    DefReq As Long
    SpeedReq As Long
    MagicReq As Long
    ClassReq As Long
    AccessReq As Byte

    addHP As Long
    addMP As Long
    addSP As Long
    AddStr As Long
    AddDef As Long
    AddMagi As Long
    AddSpeed As Long
    AddEXP As Long
    AttackSpeed As Long
    Price As Long
    Stackable As Byte
    Bound As Byte

    ' Moved back to bottom... I suck :P -Pickle
    TwoHanded As Long
    
    PetSprite As Long
    
End Type
Type MapItemRec
    num As Long
    Value As Long
    Dur As Long

    x As Byte
    y As Byte
End Type

Type NPCEditorRec
    ItemNum As Long
    ItemValue As Long
    chance As Long
End Type

Type NpcRec
    Name As String * NAME_LENGTH
    AttackSay As String * 100

    SPRITE As Long
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte

    STR  As Long
    DEF As Long
    Speed As Long
    Magi As Long
    Big As Long
    MAXHP As Long
    EXP As Long
    SpawnTime As Long

    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec

    Element As Long

    SPRITESIZE As Byte
    Quest As Long
    standstill As Boolean
End Type

Type MapNpcRec
    num As Long

    TargetType As Byte
    Target As Long

    HP As Long
    MP As Long
    SP As Long

    x As Byte
    y As Byte
    Dir As Integer

    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    LastAttack As Long
End Type

Type TradeItemRec
    GiveItem As Long
    GiveValue As Long
    GetItem As Long
    GetValue As Long
End Type

Type ShopItemRec
    ItemNum As Integer
    Price As Integer
    Amount As Integer
End Type

Type ShopRec
    Name As String * NAME_LENGTH
    FixesItems As Byte
    BuysItems As Byte
    ShowInfo As Byte
    ShopItem(1 To MAX_SHOP_ITEMS) As ShopItemRec
    CurrencyItem As Integer
End Type

Type SpellRec
    Name As String * NAME_LENGTH
    ClassReq As Long
    LevelReq As Long
    MPCost As Long
    Sound As Long
    Type As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Range As Byte

    SpellAnim As Long
    SpellTime As Long
    SpellDone As Long

    AE As Long
    Big As Long

    Element As Long
    
    TimeToCast As Long
    CastTimer As Long
End Type

Type TempTileRec
    DoorOpen()  As Byte
    DoorTimer As Long
End Type

Type GuildRec
    Name As String * NAME_LENGTH
    Founder As String * NAME_LENGTH
    Member() As String * NAME_LENGTH
End Type

Type EmoRec
    Pic As Long
    Command As String
End Type

Type ArrowRec
    Name As String
    Pic As Long
    Range As Byte
    Amount As Integer
End Type

Type StatRec
    Level As Long
    STR As Long
    DEF As Long
    Magi As Long
    Speed As Long
End Type
                                
Type QuestRec
Name As String '
LevelIsReq As Byte '
ClassIsReq As Byte '
StartOn As Byte '
LevelReq As Integer '
ClassReq As Integer '

StartItem As Long '
Startval As Long '
ItemReq As Long '
ItemVal As Long '
RewardNum As Long '
RewardVal As Long '
Start As String '
End As String '
During As String '
NotHasItem As String '
Before As String '
After As String '
FishExp As Long
MineExp As Long
LJackingExp As Long
ForagingExp As Long
UnArmedExp As Long
MageWeaponsExp As Long
CombatExp As Long
SmeltingExp As Long
IronForgingExp As Long
LeaderShipExp As Long
GovernmentExp As Long
CriticalHitExp As Long
DodgeExp As Long
RepExp As Long
PKillExp As Long
ThiefExp As Long
LargeBladesExp As Long
SmallBladesExp As Long
BluntWeaponsExp As Long
PolesExp As Long
AxesExp As Long
ThrownExp As Long
XbowsExp As Long
BowsExp As Long
CarpentryExp As Long
MillingExp As Long
SpinningExp As Long
WeavingExp As Long
SewingExp As Long
PlantingExp As Long
HarvestingExp As Long
LeatherWorkingExp As Long
SkinningExp As Long
TanningExp As Long
BodyExp As Long
MindExp As Long
SoulExp As Long
NatureExp As Long
AlchemyExp As Long
QuestingExp As Long
FirstAidExp As Long
QuestExpReward As Long
End Type



Public Quest(1 To MAX_QUESTS) As QuestRec

Function FileExist(ByVal filename As String) As Boolean

    If Dir$(App.Path & "\" & filename) = "" Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function
