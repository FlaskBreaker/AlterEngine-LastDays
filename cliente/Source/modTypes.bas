Attribute VB_Name = "modTypes"
Option Explicit
Public QuestIndex As Integer
Public NpcKillAmount As Long
Public NpcKillName As String
Public NpcKillFinal As String
Public NpcKillFinal2 As String
Public CurrentQuestNum As Long
Public CurrentQuestNpcNum As Long
Public Const MAX_QUESTS As Integer = 500
Public Const MAX_PARTY_MEMBERS As Byte = 4
Public Const MAX_PARTY_INV_SLOTS As Byte = 10

' Party Visual
Type PartyRec
    MemberIndex(1 To MAX_PARTY_MEMBERS) As Long
    MemberNames(1 To MAX_PARTY_MEMBERS) As String
    MemberSprite(1 To MAX_PARTY_MEMBERS) As Long
    Level(1 To MAX_PARTY_MEMBERS) As Long
    Leader As String
End Type

' Types para el editor de cofres
Public ChestItemNum As Long
Public ChestItemAmount As Long

Type QuestRec
name As String
LevelIsReq As Byte
ClassIsReq As Byte
StartOn As Byte
LevelReq As Integer
ClassReq As Integer
StartItem As Long
Startval As Long
ItemReq As Long
ItemVal As Long
RewardNum As Long
RewardVal As Long
Start As String
End As String
During As String
NotHasItem As String
Before As String
After As String
FishEXP As Long
MineEXP As Long
LJackingEXP As Long
ForagingEXP As Long
UnArmedEXP As Long
MageWeaponsEXP As Long
CombatExp As Long
SmeltingEXP As Long
IronForgingEXP As Long
LeaderShipEXP As Long
GovernmentEXP As Long
CriticalHitEXP As Long
DodgeEXP As Long
RepEXP As Long
PKillEXP As Long
ThiefEXP As Long
LargeBladesEXP As Long
SmallBladesEXP As Long
BluntWeaponsEXP As Long
PolesEXP As Long
AXESEXP As Long
THROWNEXP As Long
XbowsEXP As Long
BowsEXP As Long
CarpentryExp As Long
MillingEXP As Long
SpinningEXP As Long
WeavingEXP As Long
SewingEXP As Long
PlantingEXP As Long
HarvestingEXP As Long
LeatherWorkingEXP As Long
SkinningEXP As Long
TanningEXP As Long
BODYEXP As Long
MindEXP As Long
SoulEXP As Long
NatureEXP As Long
AlchemyEXP As Long
QuestingEXP As Long
FirstAidEXP As Long
QuestExpReward As Long
End Type

Type ChatBubble
    Text As String
    Created As Long
End Type

Type ScriptBubble
    Text As String
    Created As Long
    Map As Long
    X As Long
    Y As Long
    Colour As Long
End Type

Type BankRec
    Num As Long
    value As Long
    Dur As Long
End Type

Type PlayerInvRec
    Num As Long
    value As Long
    Dur As Long
End Type

Type ElementRec
    name As String * NAME_LENGTH
    Strong As Long
    Weak As Long
End Type

Type SpellAnimRec
    CastedSpell As Byte

    SpellTime As Long
    SpellVar As Long
    SpellDone As Long

    Target As Long
    TargetType As Long
End Type

Type ScriptSpellAnimRec
    CastedSpell As Byte

    SpellTime As Long
    SpellVar As Long
    SpellDone As Long

    Spellnum As Long
    X As Long
    Y As Long
End Type

Type PlayerArrowRec
    Arrow As Byte
    ArrowNum As Long
    ArrowAnim As Long
    ArrowTime As Long
    ArrowVarX As Long
    ArrowVarY As Long
    ArrowX As Long
    ArrowY As Long
    ArrowPosition As Byte
    ArrowAmount As Long
End Type

Type PlayerRec
    ' General
    name As String * NAME_LENGTH
    Guild As String
    Guildaccess As Byte
    Class As Long
    Sprite As Long
    Level As Long
    Exp As Long
    Access As Byte
    PK As Byte
    input As Byte
    iso As Byte
    Step As Byte

    ' Vitals
    HP As Long
    MP As Long
    SP As Long

    ' Stats
    STR As Long
    DEF As Long
    SPEED As Long
    MAGI As Long
    POINTS As Long

    ' Worn equipment
    ArmorSlot As Long
    WeaponSlot As Long
    HelmetSlot As Long
    ShieldSlot As Long
    LegsSlot As Long
    RingSlot As Long
    NecklaceSlot As Long

    ' Inventory
    Inv(1 To MAX_INV) As PlayerInvRec
    Spell(1 To MAX_PLAYER_SPELLS) As Long
    Bank(1 To MAX_BANK) As BankRec

    ' Position
    Map As Long
    X As Integer
    Y As Integer
    Dir As Byte

    ' Client use only
    MaxHP As Long
    MaxMP As Long
    MaxSP As Long
    XOffset As Integer
    YOffset As Integer
    MovingH As Integer
    MovingV As Integer
    Moving As Byte
    Attacking As Byte
    AttackTimer As Long
    MapGetTimer As Long
    CastedSpell As Byte

    Spellnum As Long
    SpellAnim() As SpellAnimRec
    SpellcdTimer(1 To MAX_PLAYER_SPELLS) As Single
    Spellpos(1 To MAX_PLAYER_SPELLS) As Byte
    spellcdb(1 To MAX_PLAYER_SPELLS) As Boolean
    

    EmoticonNum As Long
    EmoticonTime As Long
    EmoticonVar As Long

    LevelUp As Long
    LevelUpT As Long

    Arrow(1 To MAX_PLAYER_ARROWS) As PlayerArrowRec

    SkilLvl() As Long
    SkilExp() As Long

    Armor As Long
    Helmet As Long
    Shield As Long
    Weapon As Long
    legs As Long
    Ring As Long
    Necklace As Long
    color(1 To 3) As Byte

    head As Long
    body As Long
    leg As Long

    HookShotX As Long
    HookShotY As Long
    HookShotSucces As Long
    HookShotAnim As Long
    HookShotTime As Long
    HookShotToX As Long
    HookShotToY As Long
    HookShotDir As Long

    paperdoll As Byte
    NpcKillType As String
    NpcKillAmount As Long
    NpcKillQuestFlag As Long
    QuestFlags(1 To MAX_QUESTS) As Long
    Pet As PetRec
    InParty As Boolean
    Party As PartyRec
End Type

Type TileRec
    Ground As Long
    mask As Long
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
Type TileRec2
    Ground As String
    mask As String
    Anim As String
    Mask2 As String
    M2Anim As String
    Fringe As String
    FAnim As String
    Fringe2 As String
    F2Anim As String
    Type As String
    Data1 As String
    Data2 As String
    Data3 As String
    String1 As String
    String2 As String
    String3 As String
    Light As String
    GroundSet As String
    MaskSet As String
    AnimSet As String
    Mask2Set As String
    M2AnimSet As String
    FringeSet As String
    FAnimSet As String
    Fringe2Set As String
    F2AnimSet As String
End Type

Type MapRec
    name As String * 20
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
    Npc(1 To 25) As Integer
    SpawnX(1 To 25) As Byte
    SpawnY(1 To 25) As Byte
    owner As String
    scrolling As Byte
    Weather As Integer
End Type

Type MapRec2
    name As String
    Revision As String
    Moral As String
    Up As String
    Down As String
    Left As String
    Right As String
    music As String
    BootMap As String
    BootX As String
    BootY As String
    Shop As String
    Indoors As String
    Tile() As TileRec2
    Npc(1 To 25) As String
    SpawnX(1 To 25) As String
    SpawnY(1 To 25) As String
    owner As String
    scrolling As String
    Weather As String
End Type

Type ClassRec
    name As String * NAME_LENGTH
    MaleSprite As Long
    FemaleSprite As Long
    
    Locked As Long
    
    STR As Long
    DEF As Long
    SPEED As Long
    MAGI As Long
    
    ' For client use
    HP As Long
    MP As Long
    SP As Long
    
    ' Description
    desc As String
    
    'Gender
    gender As Long
    gender1 As String
    gender2 As String
End Type

Type ItemRec
    name As String * NAME_LENGTH
    desc As String * 150
    
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
    
    AddHP As Long
    AddMP As Long
    AddSP As Long
    AddStr As Long
    AddDef As Long
    AddMagi As Long
    AddSpeed As Long
    AddEXP As Long
    AttackSpeed As Long
    Price As Long
    
    Stackable As Long
    Bound As Long
    
    PetSprite As Long
    
End Type
    
Type MapItemRec
    Num As Long
    value As Long
    Dur As Long
    
    X As Byte
    Y As Byte
End Type

Type NPCEditorRec
    ItemNum As Long
    ItemValue As Long
    chance As Long
End Type

Type NpcRec
    name As String * NAME_LENGTH
    AttackSay As String * 100
    
    Sprite As Long
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    SpriteSize As Long
    
    STR  As Long
    DEF As Long
    SPEED As Long
    MAGI As Long
    Big As Long
    MaxHP As Long
    Exp As Long
    SpawnTime As Long
    Spell As Long
    
    ItemNPC(1 To MAX_NPC_DROPS) As NPCEditorRec
    
    Element As Long
    Quest As Integer
    standstill As Boolean
End Type

Type MapNpcRec
    Num As Long
    
    Target As Long
    
    HP As Long
    MaxHP As Long
    MP As Long
    SP As Long
    
    Map As Long
    X As Byte
    Y As Byte
    Dir As Byte
    Big As Byte
    
    ' Client use only
    XOffset As Integer
    YOffset As Integer
    Moving As Byte
    Attacking As Byte
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
    name As String * NAME_LENGTH
    FixesItems As Byte
    BuysItems As Byte
    ShowInfo As Byte
    ShopItem(1 To MAX_SHOP_ITEMS) As ShopItemRec
    currencyItem As Integer
End Type

Type SpellRec
    name As String * NAME_LENGTH
    ClassReq As Long
    LevelReq As Long
    Sound As Long
    MPCost As Long
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
    reload As Long
    
    TimeToCast As Long
    CastTimer As Long

    
End Type

Type TempTileRec
    DoorOpen As Byte
    Ground As Long
    mask As Long
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

Type PlayerTradeRec
    InvNum As Long
    InvName As String
End Type

Type EmoRec
    Pic As Long
    Command As String
End Type

Type DropRainRec
    X As Long
    Y As Long
    Randomized As Boolean
    SPEED As Byte
End Type

Type ItemTradeRec
    ItemGetNum As Long
    ItemGiveNum As Long
    ItemGetVal As Long
    ItemGiveVal As Long
End Type

Type TradeRec
    Items(1 To MAX_TRADES) As ItemTradeRec
    Selected As Long
    SelectedItem As Long
End Type

Type ArrowRec
    name As String
    Pic As Long
    Range As Byte
    Amount As Long
End Type

Type BattleMsgRec
    Msg As String
    index As Byte
    color As Byte
    time As Long
    Done As Byte
    Y As Long
End Type

Type ItemDurRec
    Item As Long
    Dur As Long
    Done As Byte
End Type


Public Quest(0 To MAX_QUESTS) As QuestRec

Sub QuestPrompt(ByVal Hate As Byte)
If Hate = 1 Then
    frmQuest.cmdYes.Visible = False
    frmQuest.cmdNo.Visible = False
    frmQuest.lblChoice.Visible = True
    frmQuest.Show
    frmQuest.SetFocus
    End If
    
    If Hate = 2 Then
    frmQuest.Show
    frmQuest.cmdYes.Visible = True
    frmQuest.cmdNo.Visible = True
    End If
    
    If Hate = 3 Then
    frmQuest.cmdYes.Visible = False
    frmQuest.cmdNo.Visible = False
    frmQuest.lblChoice.Visible = True
    frmQuest.cmdQuit.Visible = True
    frmQuest.Show
    frmQuest.SetFocus
End If
End Sub

Function HasItem(ItemNum As Long, ItemValue As Long) As Boolean
Dim PlayerHas As Boolean
Dim I

For I = 1 To MAX_INV
If GetPlayerInvItemNum(MyIndex, I) = ItemNum Then
If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
If GetPlayerInvItemValue(MyIndex, I) >= ItemValue Then
PlayerHas = True
Exit For
Else
PlayerHas = False
Exit For
End If
Else
PlayerHas = True
Exit For
End If
End If
Next I

If PlayerHas = True Then
HasItem = True
ElseIf PlayerHas = False Then
HasItem = False
End If

End Function

Public Sub QuestEditorCancel()
    InQuestEditor = False
    Unload frmQuestEditor
End Sub

Public Sub QuestEditorOk()
    Quest(EditorIndex).name = frmQuestEditor.txtName.Text
    Quest(EditorIndex).After = frmQuestEditor.txtafter.Text
    Quest(EditorIndex).Before = frmQuestEditor.txtbefore.Text
    Quest(EditorIndex).ClassIsReq = frmQuestEditor.chkcls.value
    Quest(EditorIndex).ClassReq = frmQuestEditor.lstclass.ListIndex
    Quest(EditorIndex).During = frmQuestEditor.txtduring.Text
    Quest(EditorIndex).End = frmQuestEditor.txtend.Text
    Quest(EditorIndex).ItemReq = frmQuestEditor.scrlquestitem.value
    If Item(frmQuestEditor.scrlquestitem.value).Type = 16 Or Item(frmQuestEditor.scrlquestitem.value).Stackable > 0 Then
        Quest(EditorIndex).ItemVal = frmQuestEditor.scrlquestvalue.value
    Else
        Quest(EditorIndex).ItemVal = 1
    End If
    Quest(EditorIndex).LevelIsReq = frmQuestEditor.chklvl
    Quest(EditorIndex).LevelReq = frmQuestEditor.scrllvl.value
    Quest(EditorIndex).NotHasItem = frmQuestEditor.txtnotitem.Text
    Quest(EditorIndex).RewardNum = frmQuestEditor.scrlrewitem.value
    Quest(EditorIndex).RewardVal = frmQuestEditor.scrlrewval.value
    Quest(EditorIndex).Start = frmQuestEditor.txtstart.Text
    Quest(EditorIndex).StartItem = frmQuestEditor.scrlstartnum.value
    Quest(EditorIndex).StartOn = frmQuestEditor.chkstart.value
    Quest(EditorIndex).Startval = frmQuestEditor.scrlstartval
    Quest(EditorIndex).QuestExpReward = frmQuestEditor.scrlQuestExpReward
    
    Call SendSaveQuest(EditorIndex)
    Call QuestEditorCancel
    
End Sub

Public Sub QuestEditorInit()
On Error Resume Next
Dim I As Long

    frmQuestEditor.txtName.Text = Trim(Quest(EditorIndex).name)
    frmQuestEditor.chkcls.value = Quest(EditorIndex).ClassIsReq
    frmQuestEditor.chklvl.value = Quest(EditorIndex).LevelIsReq
    frmQuestEditor.scrlQuestExpReward = Quest(EditorIndex).QuestExpReward
    
    If Quest(EditorIndex).LevelIsReq = 1 Then
        frmQuestEditor.scrllvl.value = Quest(EditorIndex).LevelReq
        frmQuestEditor.lblLevel.Caption = frmQuestEditor.scrllvl.value
        frmQuestEditor.chklvl.value = 1
        frmQuestEditor.scrllvl.Enabled = True
    Else
        frmQuestEditor.scrllvl.value = 1
        frmQuestEditor.lblLevel.Caption = "1"
        frmQuestEditor.scrllvl.Enabled = False
        frmQuestEditor.chklvl.value = 0
    End If

    
    frmQuestEditor.txtafter.Text = Quest(EditorIndex).After
    frmQuestEditor.txtbefore.Text = Quest(EditorIndex).Before
    frmQuestEditor.txtduring.Text = Quest(EditorIndex).During
    frmQuestEditor.txtend.Text = Quest(EditorIndex).End
    frmQuestEditor.txtstart.Text = Quest(EditorIndex).Start
    frmQuestEditor.txtnotitem.Text = Quest(EditorIndex).NotHasItem
    
    For I = 0 To Max_Classes
    frmQuestEditor.lstclass.addItem (I & ":" & Class(I).name)
    Next I
    
    If Quest(EditorIndex).ClassIsReq = 1 Then
        frmQuestEditor.lstclass.ListIndex = Quest(EditorIndex).ClassReq
        frmQuestEditor.chkcls.value = 1
        frmQuestEditor.lstclass.Enabled = True
    Else
        frmQuestEditor.chkcls.value = 0
        frmQuestEditor.lstclass.Enabled = False
        frmQuestEditor.lstclass.ListIndex = 0
    End If
    
    If Quest(EditorIndex).StartOn = 1 Then
        frmQuestEditor.scrlstartnum.value = Quest(EditorIndex).StartItem
        frmQuestEditor.scrlstartval.value = Quest(EditorIndex).Startval
        frmQuestEditor.lblstartitem = Quest(EditorIndex).StartItem & ":" & Item(Quest(EditorIndex).StartItem).name
        frmQuestEditor.lblstartval = Quest(EditorIndex).Startval
        frmQuestEditor.chkstart.value = 1
        If Item(frmQuestEditor.scrlstartnum.value).Type = 16 Or Item(frmQuestEditor.scrlstartnum.value).Stackable > 0 Then
            frmQuestEditor.scrlstartval.Enabled = True
            frmQuestEditor.lblstartval.Caption = frmQuestEditor.scrlstartval.value
        Else
            frmQuestEditor.scrlstartval.Enabled = False
            frmQuestEditor.lblstartval.Caption = "1"
        End If
    Else
        frmQuestEditor.scrlstartnum.value = 1
        frmQuestEditor.scrlstartval.value = 1
        frmQuestEditor.scrlstartnum.Enabled = False
        frmQuestEditor.scrlstartval.Enabled = False
        frmQuestEditor.lblstartitem = "Disabled"
        frmQuestEditor.lblstartval = "Disabled"
        frmQuestEditor.chkstart.value = 0
    End If

        frmQuestEditor.scrlquestitem.value = Quest(EditorIndex).ItemReq
        frmQuestEditor.scrlquestvalue.value = Quest(EditorIndex).ItemVal
        frmQuestEditor.lblquestitem.Caption = Quest(EditorIndex).ItemReq & ":" & Item(Quest(EditorIndex).ItemReq).name
        frmQuestEditor.lblquestval.Caption = Quest(EditorIndex).ItemVal

        If Item(frmQuestEditor.scrlstartnum.value).Type = 16 Or Item(frmQuestEditor.scrlstartnum.value).Stackable > 0 Then
            frmQuestEditor.scrlquestvalue.Enabled = True
            frmQuestEditor.lblquestval.Caption = frmQuestEditor.scrlstartval.value
        Else
            frmQuestEditor.scrlquestvalue.Enabled = False
            frmQuestEditor.lblquestval.Caption = "1"
        End If
    
        frmQuestEditor.scrlrewitem.value = Quest(EditorIndex).RewardNum
        frmQuestEditor.scrlrewval.value = Quest(EditorIndex).RewardVal
        frmQuestEditor.lblrewitem.Caption = Quest(EditorIndex).RewardNum & ":" & Item(Quest(EditorIndex).RewardNum).name
        frmQuestEditor.lblrewval.Caption = Quest(EditorIndex).RewardVal

        If Item(frmQuestEditor.scrlstartnum.value).Type = 16 Or Item(frmQuestEditor.scrlstartnum.value).Stackable > 0 Then
            frmQuestEditor.scrlrewval.Enabled = True
            frmQuestEditor.lblrewval.Caption = frmQuestEditor.scrlrewval.value
        Else
            frmQuestEditor.scrlrewval.Enabled = False
            frmQuestEditor.lblrewval.Caption = "1"
        End If
    
    frmQuestEditor.Show
End Sub

Sub InitQuestEditor()
Dim I As Long

InQuestEditor = True

frmIndex.Show
frmIndex.lstIndex.Clear

' Add the names
For I = 1 To MAX_QUESTS
frmIndex.lstIndex.addItem I & ": " & Trim(Quest(I).name)
Next I

frmIndex.lstIndex.ListIndex = 0
End Sub


Sub SendRequestEditQuest()
Dim packet As String

packet = "REQUESTEDITQUEST" & SEP_CHAR & END_CHAR
Call SendData(packet)
End Sub

Sub SendSaveQuest(ByVal QuestNum As Long)
Dim packet As String

packet = "SAVEQUEST" & SEP_CHAR & QuestNum & SEP_CHAR & Trim(Quest(QuestNum).name) & SEP_CHAR & Trim(Quest(QuestNum).After) & SEP_CHAR & Trim(Quest(QuestNum).Before) & SEP_CHAR & Quest(QuestNum).ClassIsReq & SEP_CHAR & Quest(QuestNum).ClassReq & SEP_CHAR & Trim(Quest(QuestNum).During) & SEP_CHAR & Trim(Quest(QuestNum).End) & SEP_CHAR & Quest(QuestNum).ItemReq & SEP_CHAR & Quest(QuestNum).ItemVal & SEP_CHAR & Quest(QuestNum).LevelIsReq & SEP_CHAR & Quest(QuestNum).LevelReq & SEP_CHAR & Trim(Quest(QuestNum).NotHasItem) & SEP_CHAR & Quest(QuestNum).RewardNum & SEP_CHAR & Quest(QuestNum).RewardVal & SEP_CHAR & Trim(Quest(QuestNum).Start) & SEP_CHAR & Quest(QuestNum).StartItem & SEP_CHAR & Quest(QuestNum).StartOn & SEP_CHAR & Quest(QuestNum).Startval & SEP_CHAR & Quest(QuestNum).FishEXP & SEP_CHAR & Quest(QuestNum).MineEXP & SEP_CHAR & Quest(QuestNum).LJackingEXP & SEP_CHAR & Quest(QuestNum).ForagingEXP & SEP_CHAR & Quest(QuestNum).UnArmedEXP & SEP_CHAR & Quest(QuestNum).MageWeaponsEXP & SEP_CHAR
packet = packet & Quest(QuestNum).CombatExp & SEP_CHAR & Quest(QuestNum).SmeltingEXP & SEP_CHAR & Quest(QuestNum).IronForgingEXP & SEP_CHAR & Quest(QuestNum).LeaderShipEXP & SEP_CHAR & Quest(QuestNum).GovernmentEXP & SEP_CHAR & Quest(QuestNum).CriticalHitEXP & SEP_CHAR & Quest(QuestNum).DodgeEXP & SEP_CHAR & Quest(QuestNum).RepEXP & SEP_CHAR & Quest(QuestNum).PKillEXP & SEP_CHAR & Quest(QuestNum).ThiefEXP & SEP_CHAR & Quest(QuestNum).LargeBladesEXP & SEP_CHAR & Quest(QuestNum).SmallBladesEXP & SEP_CHAR & Quest(QuestNum).BluntWeaponsEXP & SEP_CHAR & Quest(QuestNum).PolesEXP & SEP_CHAR & Quest(QuestNum).AXESEXP & SEP_CHAR & Quest(QuestNum).THROWNEXP & SEP_CHAR & Quest(QuestNum).XbowsEXP & SEP_CHAR & Quest(QuestNum).BowsEXP & SEP_CHAR & Quest(QuestNum).CarpentryExp & SEP_CHAR & Quest(QuestNum).MillingEXP & SEP_CHAR & Quest(QuestNum).SpinningEXP & SEP_CHAR & Quest(QuestNum).WeavingEXP & SEP_CHAR & Quest(QuestNum).SewingEXP & SEP_CHAR & Quest(QuestNum).PlantingEXP & SEP_CHAR
packet = packet & Quest(QuestNum).HarvestingEXP & SEP_CHAR & Quest(QuestNum).LeatherWorkingEXP & SEP_CHAR & Quest(QuestNum).SkinningEXP & SEP_CHAR & Quest(QuestNum).TanningEXP & SEP_CHAR & Quest(QuestNum).BODYEXP & SEP_CHAR & Quest(QuestNum).MindEXP & SEP_CHAR & Quest(QuestNum).SoulEXP & SEP_CHAR & Quest(QuestNum).NatureEXP & SEP_CHAR & Quest(QuestNum).AlchemyEXP & SEP_CHAR & Quest(QuestNum).QuestingEXP & SEP_CHAR & Quest(QuestNum).FirstAidEXP & SEP_CHAR & Quest(QuestNum).QuestExpReward & SEP_CHAR & END_CHAR
Call SendData(packet)
End Sub
