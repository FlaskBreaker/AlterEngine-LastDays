VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found

    szReturn = vbNullString

    sSpaces = Space(5000)

    File = App.Path & "\" & File

    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)

    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    On Error GoTo PutVar_Error
    Dim fpath As String
    fpath = App.Path & "\" & File
    Call WritePrivateProfileString(Header, Var, Value, fpath)

    On Error GoTo 0
    Exit Sub

PutVar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PutVar of Module modDatabase"
End Sub

Public Sub AdminMsg(ByVal Msg As String, ByVal Color As Byte)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerAccess(I) > 0 Then
                Call SendDataTo(I, "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
            End If
        End If
    Next I
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Byte)
    Call SendDataToAll("GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte)
    Call SendDataTo(Index, "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Sub MapMsg(ByVal mapnum As Long, ByVal Msg As String, ByVal Color As Byte)
    Call SendDataToMap(mapnum, "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
    Call SendDataTo(Index, "ALERTMSG" & SEP_CHAR & Msg & END_CHAR)
    Call CloseSocket(Index)
End Sub

Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function

Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).Char(Player(Index).CharNum).Name)
End Function

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).Char(Player(Index).CharNum).Guild)
End Function

Function GetPlayerGuildAccess(ByVal Index As Long) As Long
    GetPlayerGuildAccess = Player(Index).Char(Player(Index).CharNum).GuildAccess
End Function

Sub SetPlayerGuildAccess(ByVal Index As Long, ByVal GuildAccess As Long)
    Player(Index).Char(Player(Index).CharNum).GuildAccess = GuildAccess
End Sub

Sub SetPlayerGuild(ByVal Index As Long, ByVal Guild As String)
    Player(Index).Char(Player(Index).CharNum).Guild = Guild
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Char(Player(Index).CharNum).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Char(Player(Index).CharNum).Class = ClassNum
End Sub

Function GetPlayerClassName(ByVal Index As Long) As String
    GetPlayerClassName = Trim(ClassData(GetPlayerClass(Index)).Name)
End Function

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Char(Player(Index).CharNum).SPRITE
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal SPRITE As Long)
    Player(Index).Char(Player(Index).CharNum).SPRITE = SPRITE
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Char(Player(Index).CharNum).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Char(Player(Index).CharNum).Level = Level
End Sub

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    If GetPlayerLevel(Index) <= MAX_LEVEL Then
        GetPlayerNextLevel = Experience(GetPlayerLevel(Index))
    End If
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Char(Player(Index).CharNum).EXP
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)
    Player(Index).Char(Player(Index).CharNum).EXP = EXP
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Char(Player(Index).CharNum).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Char(Player(Index).CharNum).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).Char(Player(Index).CharNum).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).Char(Player(Index).CharNum).PK = PK
End Sub

Function GetPlayerHP(ByVal Index As Long) As Long
    GetPlayerHP = Player(Index).Char(Player(Index).CharNum).HP
End Function

Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)
    Player(Index).Char(Player(Index).CharNum).HP = HP

    If GetPlayerHP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).HP = 0
    End If

    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
        Player(Index).Char(Player(Index).CharNum).HP = GetPlayerMaxHP(Index)
    End If

    Call SendHP(Index)
End Sub

Function GetPlayerMP(ByVal Index As Long) As Long
    GetPlayerMP = Player(Index).Char(Player(Index).CharNum).MP
End Function

Sub SetPlayerMP(ByVal Index As Long, ByVal MP As Long)
    Player(Index).Char(Player(Index).CharNum).MP = MP

    If GetPlayerMP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).MP = 0
    End If

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then
        Player(Index).Char(Player(Index).CharNum).MP = GetPlayerMaxMP(Index)
    End If

    Call SendMP(Index)
End Sub

Function GetPlayerSP(ByVal Index As Long) As Long
    GetPlayerSP = Player(Index).Char(Player(Index).CharNum).SP
End Function

Sub SetPlayerSP(ByVal Index As Long, ByVal SP As Long)
    Player(Index).Char(Player(Index).CharNum).SP = SP

    If GetPlayerSP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).SP = 0
    End If

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then
        Player(Index).Char(Player(Index).CharNum).SP = GetPlayerMaxSP(Index)
    End If

    Call SendSP(Index)
End Sub

Function GetPlayerMaxHP(ByVal Index As Long) As Long
    Dim CharNum As Long
    Dim Add As Long
    Add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).addHP
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).addHP
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).addHP
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).addHP
    End If
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).addHP
    End If
    If GetPlayerRingSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))).addHP
    End If
    If GetPlayerNecklaceSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))).addHP
    End If

    CharNum = Player(Index).CharNum
    ' GetPlayerMaxHP = ((Player(index).Char(CharNum).Level + Int(GetPlayerSTR(index) / 2) + ClassData(Player(index).Char(CharNum).Class).STR) * 2) + add
    GetPlayerMaxHP = (GetPlayerLevel(Index) * addHP.Level) + (GetPlayerSTR(Index) * addHP.STR) + (GetPlayerDEF(Index) * addHP.DEF) + (GetPlayerMAGI(Index) * addHP.Magi) + (GetPlayerSPEED(Index) * addHP.Speed) + Add
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
    Dim CharNum As Long
    Dim Add As Long

    Add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).addMP
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).addMP
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).addMP
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).addMP
    End If
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).addMP
    End If
    If GetPlayerRingSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))).addMP
    End If
    If GetPlayerNecklaceSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))).addMP
    End If

    CharNum = Player(Index).CharNum
    ' GetPlayerMaxMP = ((Player(index).Char(CharNum).Level + Int(GetPlayerMAGI(index) / 2) + ClassData(Player(index).Char(CharNum).Class).MAGI) * 2) + add
    GetPlayerMaxMP = (GetPlayerLevel(Index) * addMP.Level) + (GetPlayerSTR(Index) * addMP.STR) + (GetPlayerDEF(Index) * addMP.DEF) + (GetPlayerMAGI(Index) * addMP.Magi) + (GetPlayerSPEED(Index) * addMP.Speed) + Add
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
    GetPlayerMaxSP = (Player(Index).Char(Player(Index).CharNum).Level + Int(GetPlayerSPEED(Index) / 2) + ClassData(Player(Index).Char(Player(Index).CharNum).Class).Speed) * 2
End Function

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
    GetClassMaxHP = (1 + Int(ClassData(ClassNum).STR / 2) + ClassData(ClassNum).STR) * 2
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
    GetClassMaxMP = (1 + Int(ClassData(ClassNum).Magi / 2) + ClassData(ClassNum).Magi) * 2
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
    GetClassMaxSP = (1 + Int(ClassData(ClassNum).Speed / 2) + ClassData(ClassNum).Speed) * 2
End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long
    GetClassSTR = ClassData(ClassNum).STR
End Function

Function GetClassDEF(ByVal ClassNum As Long) As Long
    GetClassDEF = ClassData(ClassNum).DEF
End Function

Function GetClassSPEED(ByVal ClassNum As Long) As Long
    GetClassSPEED = ClassData(ClassNum).Speed
End Function

Function GetClassMAGI(ByVal ClassNum As Long) As Long
    GetClassMAGI = ClassData(ClassNum).Magi
End Function

Function GetPlayerSTR(ByVal Index As Long) As Long
    GetPlayerSTR = Player(Index).Char(Player(Index).CharNum).STR
End Function

Sub SetPlayerSTR(ByVal Index As Long, ByVal STR As Long)
    Player(Index).Char(Player(Index).CharNum).STR = STR
End Sub

Function GetPlayerDEF(ByVal Index As Long) As Long
    GetPlayerDEF = Player(Index).Char(Player(Index).CharNum).DEF
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal DEF As Long)
    Player(Index).Char(Player(Index).CharNum).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal Index As Long) As Long
    GetPlayerSPEED = Player(Index).Char(Player(Index).CharNum).Speed
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal Speed As Long)
    Player(Index).Char(Player(Index).CharNum).Speed = Speed
End Sub

Function GetPlayerMAGI(ByVal Index As Long) As Long
    GetPlayerMAGI = Player(Index).Char(Player(Index).CharNum).Magi
End Function

Sub SetPlayerMAGI(ByVal Index As Long, ByVal Magi As Long)
    Player(Index).Char(Player(Index).CharNum).Magi = Magi
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).Char(Player(Index).CharNum).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).Char(Player(Index).CharNum).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = Player(Index).Char(Player(Index).CharNum).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal mapnum As Long)
    If mapnum > 0 And mapnum <= MAX_MAPS Then
        Player(Index).Char(Player(Index).CharNum).Map = mapnum
    End If
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).Char(Player(Index).CharNum).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    Player(Index).Char(Player(Index).CharNum).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Char(Player(Index).CharNum).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Char(Player(Index).CharNum).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Char(Player(Index).CharNum).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Char(Player(Index).CharNum).Dir = Dir
End Sub

Function GetPlayerIP(ByVal Index As Long) As String
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot)
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long, ByVal spellnum As Long)
    Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot) = spellnum
End Sub

Function GetPlayerArmorSlot(ByVal Index As Long) As Long
    GetPlayerArmorSlot = Player(Index).Char(Player(Index).CharNum).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal Index As Long) As Long
    GetPlayerWeaponSlot = Player(Index).Char(Player(Index).CharNum).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal Index As Long) As Long
    GetPlayerHelmetSlot = Player(Index).Char(Player(Index).CharNum).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal Index As Long) As Long
    GetPlayerShieldSlot = Player(Index).Char(Player(Index).CharNum).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).ShieldSlot = InvNum
End Sub

Function GetPlayerNecklaceSlot(ByVal Index As Long) As Long
    GetPlayerNecklaceSlot = Player(Index).Char(Player(Index).CharNum).NecklaceSlot
End Function

Sub SetPlayerNecklaceSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).NecklaceSlot = InvNum
End Sub

Function GetPlayerRingSlot(ByVal Index As Long) As Long
    GetPlayerRingSlot = Player(Index).Char(Player(Index).CharNum).RingSlot
End Function

Sub SetPlayerRingSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).RingSlot = InvNum
End Sub

Function GetPlayerLegsSlot(ByVal Index As Long) As Long
    GetPlayerLegsSlot = Player(Index).Char(Player(Index).CharNum).LegsSlot
End Function

Sub SetPlayerLegsSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).LegsSlot = InvNum
End Sub

Function GetMapBootMap(ByVal Index As Long) As Long
    GetMapBootMap = Map(GetPlayerMap(Index)).BootMap
End Function

Function GetMapBootX(ByVal Index As Long) As Long
    GetMapBootX = Map(GetPlayerMap(Index)).BootX
End Function

Function GetMapBootY(ByVal Index As Long) As Long
    GetMapBootY = Map(GetPlayerMap(Index)).BootY
End Function

Sub MutePlayer(ByVal Index As Integer)
    Player(Index).Mute = True
End Sub

Sub UnMutePlayer(ByVal Index As Integer)
    Player(Index).Mute = False
End Sub

Public Sub MakeDay()
    GameTime = TIME_DAY

    Hours = 7
    Minutes = 0
    Seconds = 0

    Call SendTimeToAll
End Sub

Public Sub MakeNight()
    GameTime = TIME_NIGHT

    Hours = 21
    Minutes = 0
    Seconds = 0

    Call SendTimeToAll
End Sub

Function IsScrolling() As Boolean
    If IS_SCROLLING = 0 Then
        IsScrolling = False
    Else
        IsScrolling = True
    End If
End Function

Function GetMaxPlayers() As Long
    GetMaxPlayers = MAX_PLAYERS
End Function

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long)
    Dim Packet As String
    Dim I As Long

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    I = MapItemSlot

    If I <> 0 Then
        If ItemNum >= 0 Then
            If ItemNum <= MAX_ITEMS Then
                MapItem(mapnum, I).num = ItemNum
                MapItem(mapnum, I).Value = ItemVal
        
                If ItemNum <> 0 Then
                    If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_NECKLACE) Then
                        MapItem(mapnum, I).Dur = ItemDur
                    Else
                        MapItem(mapnum, I).Dur = 0
                    End If
                Else
                    MapItem(mapnum, I).Dur = 0
                End If
        
                MapItem(mapnum, I).X = X
                MapItem(mapnum, I).Y = Y
        
                Packet = "SPAWNITEM" & SEP_CHAR & I & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(mapnum, I).Dur & SEP_CHAR & X & SEP_CHAR & Y & END_CHAR
                Call SendDataToMap(mapnum, Packet)
            End If
        End If
    End If
    
    
End Sub

Function IsConnected(ByVal Index As Long) As Boolean
    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
    Exit Function
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If IsConnected(Index) Then
        If Player(Index).InGame Then
            IsPlaying = True
        Else
            IsPlaying = False
        End If
    Else
        IsPlaying = False
    End If
End Function

Sub SendInventory(ByVal Index As Long)
    Dim Packet As String
    Dim I As Long

    Packet = "PLAYERINV" & SEP_CHAR & Index & SEP_CHAR
    For I = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(Index, I) & SEP_CHAR & GetPlayerInvItemValue(Index, I) & SEP_CHAR & GetPlayerInvItemDur(Index, I) & SEP_CHAR
    Next I
    Packet = Packet & END_CHAR

    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Long)
    Call SendDataToMap(GetPlayerMap(Index), "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & Index & SEP_CHAR & GetPlayerInvItemNum(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(Index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(Index, InvSlot) & SEP_CHAR & Index & END_CHAR)
End Sub

Sub SendIndexInventoryFromMap(ByVal Index As Long)
    Dim Packet As String
    Dim N As Long
    Dim I As Long
    
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = GetPlayerMap(Index) Then
                Packet = "PLAYERINV" & SEP_CHAR & I & SEP_CHAR
                For N = 1 To MAX_INV
                    Packet = Packet & GetPlayerInvItemNum(I, N) & SEP_CHAR & GetPlayerInvItemValue(I, N) & SEP_CHAR & GetPlayerInvItemDur(I, N) & SEP_CHAR
                Next N
                Packet = Packet & END_CHAR

                Call SendDataTo(Index, Packet)
            End If
        End If
    Next I
End Sub

Sub SendWornEquipment(ByVal Index As Long)
    If IsPlaying(Index) Then
        Call SendDataToMap(GetPlayerMap(Index), "PLAYERWORNEQ" & SEP_CHAR & Index & SEP_CHAR & GetPlayerArmorSlot(Index) & SEP_CHAR & GetPlayerWeaponSlot(Index) & SEP_CHAR & GetPlayerHelmetSlot(Index) & SEP_CHAR & GetPlayerShieldSlot(Index) & SEP_CHAR & GetPlayerLegsSlot(Index) & SEP_CHAR & GetPlayerRingSlot(Index) & SEP_CHAR & GetPlayerNecklaceSlot(Index) & END_CHAR)
    End If
End Sub

Sub SendHP(ByVal Index As Long)
    Call SendDataToMap(GetPlayerMap(Index), "PLAYERHP" & SEP_CHAR & Index & SEP_CHAR & GetPlayerMaxHP(Index) & SEP_CHAR & GetPlayerHP(Index) & END_CHAR)
End Sub

Sub SendMP(ByVal Index As Long)
    Call SendDataTo(Index, "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(Index) & SEP_CHAR & GetPlayerMP(Index) & END_CHAR)
End Sub

Sub SendSP(ByVal Index As Long)
    Call SendDataTo(Index, "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(Index) & SEP_CHAR & GetPlayerSP(Index) & END_CHAR)
End Sub

Sub SendPTS(ByVal Index As Long)
    Call SendDataTo(Index, "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(Index) & END_CHAR)
End Sub

Sub SendStats(ByVal Index As Long)
    Call SendDataTo(Index, "PLAYERSTATSPACKET" & SEP_CHAR & GetPlayerSTR(Index) & SEP_CHAR & GetPlayerDEF(Index) & SEP_CHAR & GetPlayerSPEED(Index) & SEP_CHAR & GetPlayerMAGI(Index) & SEP_CHAR & GetPlayerNextLevel(Index) & SEP_CHAR & GetPlayerExp(Index) & SEP_CHAR & GetPlayerLevel(Index) & END_CHAR)
End Sub

Sub Flash(ByVal Index As Long, ByVal flashfile As String)
    Call SendDataTo(Index, "flashevent" & SEP_CHAR & flashfile & END_CHAR)
End Sub

Sub Prompt(ByVal Index As Long, ByVal question As String, ByVal Value As Long)
    Call SendDataTo(Index, "prompt" & SEP_CHAR & question & SEP_CHAR & Value & END_CHAR)
End Sub

Sub PlaySound(ByVal Index As Long, ByVal Sound As String)
    Call SendDataTo(Index, "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & END_CHAR)
End Sub

Sub PlaySoundToMap(ByVal Index As Long, ByVal Sound As String)
    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & END_CHAR)
End Sub

Sub SendPlayerData(ByVal Index As Long)
    Dim Packet As String
    Dim j As Long

    ' Send index's player data to everyone including himself on th emap
    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & Index & SEP_CHAR
    Packet = Packet & GetPlayerName(Index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(Index) & SEP_CHAR
    Packet = Packet & GetPlayerMap(Index) & SEP_CHAR
    Packet = Packet & GetPlayerX(Index) & SEP_CHAR
    Packet = Packet & GetPlayerY(Index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(Index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(Index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(Index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(Index) & SEP_CHAR
    Packet = Packet & GetPlayerHead(Index) & SEP_CHAR
    Packet = Packet & GetPlayerBody(Index) & SEP_CHAR
    Packet = Packet & GetPlayerleg(Index) & SEP_CHAR
    Packet = Packet & GetPlayerPaperdoll(Index) & SEP_CHAR
    Packet = Packet & GetPlayerLevel(Index) & SEP_CHAR
    Packet = Packet & Player(Index).Char(Player(Index).CharNum).Color(1) & SEP_CHAR
    Packet = Packet & Player(Index).Char(Player(Index).CharNum).Color(2) & SEP_CHAR
    Packet = Packet & Player(Index).Char(Player(Index).CharNum).Color(3) & SEP_CHAR
    
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
    
    If Player(Index).Pet.Alive = YES Then
        Packet = "PETDATA" & SEP_CHAR
        Packet = Packet & Index & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Map & SEP_CHAR
        Packet = Packet & Player(Index).Pet.X & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Y & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Dir & SEP_CHAR
        Packet = Packet & Player(Index).Pet.SPRITE & SEP_CHAR
        Packet = Packet & Player(Index).Pet.HP & SEP_CHAR
        Packet = Packet & Player(Index).Pet.STR * 5 & SEP_CHAR
        Packet = Packet & Player(Index).Pet.STR & SEP_CHAR
        Packet = Packet & Player(Index).Pet.DEF & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Speed & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Magi & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Level & SEP_CHAR
        Packet = Packet & Player(Index).Pet.POINTS & SEP_CHAR
        Packet = Packet & GetPetSP(Index) & SEP_CHAR
        Packet = Packet & GetPetMAXSP(Index) & SEP_CHAR
        Packet = Packet & GetPetMP(Index) & SEP_CHAR
        Packet = Packet & GetPetMAXMP(Index) & SEP_CHAR
        Packet = Packet & GetPetFP(Index) & SEP_CHAR
        Packet = Packet & GetPetMAXFP(Index) & SEP_CHAR
        Packet = Packet & GetPetExp(Index) & SEP_CHAR
        Packet = Packet & GetPetNextLevel(Index) & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Name & SEP_CHAR
        Packet = Packet & GetPlayerLevel(Index) & SEP_CHAR
        Packet = Packet & END_CHAR
        Call SendDataToMap(GetPlayerMap(Index), Packet)
    End If
    
    
End Sub

Sub SendDataTo(ByVal Index As Long, ByVal Data As String)
    If IsConnected(Index) Then
        frmServer.Socket(Index).SendData Data
        DoEvents
    End If
End Sub

Sub SendDataToAll(ByVal Data As String)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            Call SendDataTo(I, Data)
        End If
    Next I
End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByVal Data As String)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If I <> Index Then
                Call SendDataTo(I, Data)
            End If
        End If
    Next I
End Sub

Sub SendDataToMap(ByVal mapnum As Long, ByVal Data As String)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = mapnum Then
                Call SendDataTo(I, Data)
            End If
        End If
    Next I
End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal mapnum As Long, ByVal Data As String)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = mapnum And I <> Index Then
                Call SendDataTo(I, Data)
            End If
        End If
    Next I
End Sub

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Char(Player(Index).CharNum).Name = Name
    Call SendPlayerData(Index)
End Sub

Function GetPlayerCharNum(ByVal Index As Long) As Long
    GetPlayerCharNum = Player(Index).CharNum
End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim I As Long

    Name = LCase$(Name)

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If Len(GetPlayerName(I)) >= Len(Name) Then
                If LCase$(GetPlayerName(I)) = Name Then
                    FindPlayer = I
                    Exit Function
                End If
            End If
        End If
    Next I
End Function

Public Sub PlayerWarp(ByVal Index As Long, ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long)
    Dim OldMap As Long

    On Error GoTo WarpErr

    If Not IsPlaying(Index) Then
        Exit Sub
    End If

    If mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)

    If Not OldMap = mapnum Then
        Call SendLeaveMap(Index, OldMap)
    End If

    Call SetPlayerMap(Index, mapnum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)
    
    If Player(Index).Pet.Alive = YES Then
        Player(Index).Pet.MapToGo = mapnum
        Player(Index).Pet.Map = mapnum
        Player(Index).Pet.X = GetPlayerX(Index)
        Player(Index).Pet.Y = GetPlayerY(Index)
        'Call UpdateGrid(Player(Index).Pet.Map, Player(Index).Pet.x, Player(Index).Pet.y, Player(Index).Pet.Map, x, y)
    End If

    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    PlayersOnMap(mapnum) = YES

    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "warp" & END_CHAR)

    Player(Index).GettingMap = YES

    Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & mapnum & SEP_CHAR & Map(mapnum).Revision & END_CHAR)

    Call SendInventory(Index)
    Call SendIndexInventoryFromMap(Index)
    Call SendIndexWornEquipmentFromMap(Index)

    If scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "OnMapLoad " & Index & "," & OldMap & "," & mapnum
    End If
    
    Exit Sub

WarpErr:
    Call AddLog("PlayerWarp error for player index " & Index & " on map " & GetPlayerMap(Index) & ".", "logs\ErrorLog.txt")
End Sub

Sub JoinWarp(ByVal Index As Long, ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long)
    Dim OldMap As Long

    If Not IsPlaying(Index) Then
        Exit Sub
    End If

    If mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(Index)

    Call SendLeaveMap(Index, OldMap)

    Call SetPlayerMap(Index, mapnum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)

    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    PlayersOnMap(mapnum) = YES

    Player(Index).GettingMap = YES

    Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & mapnum & SEP_CHAR & Map(mapnum).Revision & END_CHAR)

    Call SendInventory(Index)
    Call SendIndexWornEquipmentFromMap(Index)
End Sub

Public Sub AddLog(ByVal Text As String, ByVal FN As String)
    Dim filename As String
    Dim FileID As Long

    If ServerLog Then
        filename = App.Path & "\" & FN

        If FileExists(FN) Then
            FileID = FreeFile
            Open filename For Output As #FileID
            Print #FileID, Time & ": " & Text
            Close #FileID
        End If
    End If
End Sub

Sub HackingAttempt(ByVal Index As Long, ByVal Reason As String)
    If Index > 0 Then
        If IsPlaying(Index) Then
            Call AdminMsg(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " ha sido desconectado por (" & Reason & ")", WHITE)
            Call AddLog(GetPlayerName(Index) & " ha sido kickeado por '" & Reason & "'.", "Logs\HackAttempt.txt")
        End If

        Call AlertMsg(Index, "Has perdido la conexión con " & GAME_NAME & ".")
    End If
End Sub

Sub BattleMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte, ByVal Side As Long)
    Call SendDataTo(Index, "damagedisplay" & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function
Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).num
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).num = ItemNum
    Call SendBankUpdate(Index, BankSlot)
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Value = ItemValue
    Call SendBankUpdate(Index, BankSlot)
End Sub

Function GetPlayerBankItemDur(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemDur = Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Dur
End Function

Sub SetPlayerBankItemDur(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemDur As Long)
    Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Dur = ItemDur
End Sub

Function GetPlayerTarget(ByVal Index As Long)
    If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
        GetPlayerTarget = Player(Index).Target
    Else
        GetPlayerTarget = -1
    End If
End Function
Sub SetTimer(ByVal Name As String, ByVal Interval As Long)
    Call AddNewTimer(Name, Interval)
End Sub

Function GetTimer(ByVal Name As String) As Long
    GetTimer = GetTimeLeft(Name)
End Function

Sub RemoveTimer(ByVal Name As String)
    Call GetRidOfTimer(Name)
End Sub

Sub SetTile(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal xset As Long, ByVal yset As Long, ByVal tileset As Long, ByVal layer As Long)
    Call ScriptSetTile(mapper, X, Y, xset, yset, tileset, layer)
End Sub
Function GetTileX(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal layer As Long)
    Select Case layer
        Case 0
            GetTileX = Map(mapper).Tile(X, Y).Ground - Int(Map(mapper).Tile(X, Y).Ground / 14) * 14
        Case 1
            GetTileX = Map(mapper).Tile(X, Y).Mask - Int(Map(mapper).Tile(X, Y).Mask / 14) * 14
        Case 2
            GetTileX = Map(mapper).Tile(X, Y).Anim - Int(Map(mapper).Tile(X, Y).Anim / 14) * 14
        Case 3
            GetTileX = Map(mapper).Tile(X, Y).Mask2 - Int(Map(mapper).Tile(X, Y).Mask2 / 14) * 14
        Case 4
            GetTileX = Map(mapper).Tile(X, Y).M2Anim - Int(Map(mapper).Tile(X, Y).M2Anim / 14) * 14
        Case 5
            GetTileX = Map(mapper).Tile(X, Y).Fringe - Int(Map(mapper).Tile(X, Y).Fringe / 14) * 14
        Case 6
            GetTileX = Map(mapper).Tile(X, Y).FAnim - Int(Map(mapper).Tile(X, Y).FAnim / 14) * 14
        Case 7
            GetTileX = Map(mapper).Tile(X, Y).Fringe2 - Int(Map(mapper).Tile(X, Y).Fringe2 / 14) * 14
        Case 8
            GetTileX = Map(mapper).Tile(X, Y).F2Anim - Int(Map(mapper).Tile(X, Y).F2Anim / 14) * 14
    End Select
End Function
Function GetTileY(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal layer As Long)
    Select Case layer
        Case 0
            GetTileY = Int(Map(mapper).Tile(X, Y).Ground / 14)
        Case 1
            GetTileY = Int(Map(mapper).Tile(X, Y).Mask / 14)
        Case 2
            GetTileY = Int(Map(mapper).Tile(X, Y).Anim / 14)
        Case 3
            GetTileY = Int(Map(mapper).Tile(X, Y).Mask2 / 14)
        Case 4
            GetTileY = Int(Map(mapper).Tile(X, Y).M2Anim / 14)
        Case 5
            GetTileY = Int(Map(mapper).Tile(X, Y).Fringe / 14)
        Case 6
            GetTileY = Int(Map(mapper).Tile(X, Y).FAnim / 14)
        Case 7
            GetTileY = Int(Map(mapper).Tile(X, Y).Fringe2 / 14)
        Case 8
            GetTileY = Int(Map(mapper).Tile(X, Y).F2Anim / 14)
    End Select
End Function
Function GetTileSet(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal layer As Long)
    Select Case layer
        Case 0
            GetTileSet = Map(mapper).Tile(X, Y).GroundSet
        Case 1
            GetTileSet = Map(mapper).Tile(X, Y).MaskSet
        Case 2
            GetTileSet = Map(mapper).Tile(X, Y).AnimSet
        Case 3
            GetTileSet = Map(mapper).Tile(X, Y).Mask2Set
        Case 4
            GetTileSet = Map(mapper).Tile(X, Y).M2AnimSet
        Case 5
            GetTileSet = Map(mapper).Tile(X, Y).FringeSet
        Case 6
            GetTileSet = Map(mapper).Tile(X, Y).FAnimSet
        Case 7
            GetTileSet = Map(mapper).Tile(X, Y).Fringe2Set
        Case 8
            GetTileSet = Map(mapper).Tile(X, Y).F2AnimSet
    End Select
End Function

Sub SendMap(ByVal mapper As Long)
    Call SendDataToMap(mapper, "CHECKFORMAP" & SEP_CHAR & mapper & SEP_CHAR & (Map(mapper).Revision + 1) & END_CHAR)
End Sub

Sub SpellAnim(ByVal spellnum As Long, ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    Call SendDataToMap(mapper, "scriptspellanim" & SEP_CHAR & spellnum & SEP_CHAR & Spell(spellnum).SpellAnim & SEP_CHAR & Spell(spellnum).SpellTime & SEP_CHAR & Spell(spellnum).SpellDone & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Spell(spellnum).Big & END_CHAR)
End Sub
Sub SetAttribute(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal Attrib As Long, ByVal Data1 As Long, ByVal Data2 As Long, ByVal Data3 As Long, ByVal String1 As String, ByVal String2 As String, ByVal String3 As String)
    Call ScriptSetAttribute(mapper, X, Y, Attrib, Data1, Data2, Data3, String1, String2, String3)
End Sub
Function GetAttribute(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetAttribute = Map(mapper).Tile(X, Y).Type
End Function
Function GetTileData1(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetTileData1 = Map(mapper).Tile(X, Y).Data1
End Function
Function GetTileData2(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetTileData2 = Map(mapper).Tile(X, Y).Data2
End Function
Function GetTileData3(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetTileData3 = Map(mapper).Tile(X, Y).Data3
End Function
Function GetTileString1(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetTileString1 = Map(mapper).Tile(X, Y).String1
End Function
Function GetTileString2(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetTileString2 = Map(mapper).Tile(X, Y).String2
End Function
Function GetTileString3(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long)
    GetTileString3 = Map(mapper).Tile(X, Y).String3
End Function
Sub PlayerQueryBox(ByVal Index As Long, ByVal Message As String, ByVal Script As Long)
    Call SendDataTo(Index, "querybox" & SEP_CHAR & Message & SEP_CHAR & Script & END_CHAR)
End Sub
Sub UpdatePaperDoll(ByVal Index As Long)
    Call SendDataToMap(GetPlayerMap(Index), "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
End Sub
Sub UpdateSprite(ByVal Index As Long)
    Call SendDataToMap(GetPlayerMap(Index), "updatesprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index))
End Sub

Sub SetMapNpcNumber(ByVal mapnum As Long, ByVal Index As Long, ByVal Number As Long)
    If Number >= 1 Then
        MapNPC(mapnum, Index).num = Number
    End If
End Sub

Sub SetMapNpcTarget(ByVal mapnum As Long, ByVal Index As Long, ByVal Target As Long)
    MapNPC(mapnum, Index).Target = Target
End Sub

Public Sub SetMapName(ByVal mapnum As Long, ByVal MapName As String)
    Dim I As Long

    Map(mapnum).Name = MapName
    Map(mapnum).Revision = Map(mapnum).Revision + 1

    If GetTotalMapPlayers(mapnum) > 0 Then
        For I = 1 To MAX_PLAYERS
            If IsPlaying(I) Then
                If GetPlayerMap(I) = mapnum Then
                    Call SendMap(mapnum)
                End If
            End If
        Next I
    End If
End Sub

Sub SetMapNpcDir(ByVal mapnum As Long, ByVal Index As Long, ByVal Direction As Long)
    MapNPC(mapnum, Index).Dir = Direction
End Sub

Sub SetMapNpcY(ByVal mapnum As Long, ByVal Index As Long, ByVal NPC_Y As Long)
    If NPC_Y <= MAX_MAPY And NPC_Y >= 0 Then
        MapNPC(mapnum, Index).Y = NPC_Y
    End If
End Sub

Sub SetMapNpcX(ByVal mapnum As Long, ByVal Index As Long, ByVal NPC_X As Long)
    If NPC_X <= MAX_MAPX And NPC_X >= 0 Then
        MapNPC(mapnum, Index).X = NPC_X
    End If
End Sub

Sub SetMapNpcHP(ByVal mapnum As Long, ByVal Index As Long, ByVal HitPoints As Long)
    MapNPC(mapnum, Index).HP = HitPoints
End Sub

Sub SendNPC(ByVal mapnum As Long, ByVal MapNpcNum As Long)
    Call SendDataToMap(mapnum, "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(mapnum, MapNpcNum).num & SEP_CHAR & MapNPC(mapnum, MapNpcNum).X & SEP_CHAR & MapNPC(mapnum, MapNpcNum).Y & SEP_CHAR & MapNPC(mapnum, MapNpcNum).Dir & SEP_CHAR & NPC(MapNPC(mapnum, MapNpcNum).num).Big & END_CHAR)
End Sub

Function GetNpcMaxHP(ByVal Index As Long) As Long
    GetNpcMaxHP = NPC(Index).MAXHP
End Function

Function GetMapNpcNumber(ByVal mapnum As Long, ByVal Index As Long) As Long
    GetMapNpcNumber = MapNPC(mapnum, Index).num
End Function

Function GetMapNpcHP(ByVal mapnum As Long, ByVal Index As Long) As Long
    GetMapNpcHP = MapNPC(mapnum, Index).HP
End Function

Function GetNpcName(ByVal Number As Long) As String
    GetNpcName = Trim$(NPC(Number).Name)
End Function

Function GetNpcBehavior(ByVal Number As Long) As Long
    GetNpcBehavior = NPC(Number).Behavior
End Function

Function GetNpcExp(ByVal Number As Long) As Long
    GetNpcExp = NPC(Number).EXP
End Function

Function GetNpcDefense(ByVal Number As Long) As Long
    GetNpcDefense = NPC(Number).DEF
End Function

Function GetNpcStrength(ByVal Number As Long) As Long
    GetNpcStrength = NPC(Number).STR
End Function

Sub SendIndexWornEquipment(ByVal Index As Long)
    Dim Packet As String
    Dim Armor As Long
    Dim Helmet As Long
    Dim Shield As Long
    Dim Weapon As Long
    Dim Legs As Long
    Dim Ring As Long
    Dim Necklace As Long

    Armor = 0
    Helmet = 0
    Shield = 0
    Weapon = 0
    Legs = 0
    Ring = 0
    Necklace = 0

    If GetPlayerArmorSlot(Index) > 0 Then
        Armor = GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        Helmet = GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        Shield = GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))
    End If
    If GetPlayerWeaponSlot(Index) > 0 Then
        Weapon = GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))
    End If
    If GetPlayerLegsSlot(Index) > 0 Then
        Legs = GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))
    End If
    If GetPlayerRingSlot(Index) > 0 Then
        Ring = GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))
    End If
    If GetPlayerNecklaceSlot(Index) > 0 Then
        Necklace = GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))
    End If

    Packet = "itemworn" & SEP_CHAR & Index & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub SendIndexWornEquipmentFromMap(ByVal Index As Long)
    Dim Packet As String
    Dim I As Long
    Dim Armor As Long
    Dim Helmet As Long
    Dim Shield As Long
    Dim Weapon As Long
    Dim Legs As Long
    Dim Ring As Long
    Dim Necklace As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) = True Then
            If GetPlayerMap(Index) = GetPlayerMap(I) Then
                Armor = 0
                Helmet = 0
                Shield = 0
                Weapon = 0
                Legs = 0
                Ring = 0
                Necklace = 0

                If GetPlayerArmorSlot(I) > 0 Then
                    Armor = GetPlayerInvItemNum(I, GetPlayerArmorSlot(I))
                End If
                If GetPlayerHelmetSlot(I) > 0 Then
                    Helmet = GetPlayerInvItemNum(I, GetPlayerHelmetSlot(I))
                End If
                If GetPlayerShieldSlot(I) > 0 Then
                    Shield = GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))
                End If
                If GetPlayerWeaponSlot(I) > 0 Then
                    Weapon = GetPlayerInvItemNum(I, GetPlayerWeaponSlot(I))
                End If
                If GetPlayerLegsSlot(I) > 0 Then
                    Legs = GetPlayerInvItemNum(I, GetPlayerLegsSlot(I))
                End If
                If GetPlayerRingSlot(I) > 0 Then
                    Ring = GetPlayerInvItemNum(I, GetPlayerRingSlot(I))
                End If
                If GetPlayerNecklaceSlot(I) > 0 Then
                    Necklace = GetPlayerInvItemNum(I, GetPlayerNecklaceSlot(I))
                End If

                Packet = "itemworn" & SEP_CHAR & I & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & END_CHAR
                Call SendDataTo(Index, Packet)
            End If
        End If
    Next I
End Sub

Function GetPlayersOnMap(ByVal mapnum As Long) As Long
    Dim I As Long
    Dim Count As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = mapnum Then
                Count = Count + 1
            End If
        End If
    Next I

    GetPlayersOnMap = Count
End Function

Sub ScriptSpawnNpc(ByVal MapNpcNum As Long, ByVal mapnum As Long, ByVal spawn_X As Long, ByVal spawn_Y As Long, ByVal npcnum As Long)
    ' NPC_index               map_number          X spawn          Y spawn            NPC_number
    Dim Packet As String
    Dim I As Long

    ' Check for subscript out of range
    If MapNpcNum < 0 Or MapNpcNum > MAX_MAP_NPCS Or mapnum <= 0 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    If npcnum = 0 Then
        Map(mapnum).Revision = Map(mapnum).Revision + 1
        MapNPC(mapnum, MapNpcNum).num = 0
        Map(mapnum).NPC(MapNpcNum) = 0
        MapNPC(mapnum, MapNpcNum).Target = 0
        MapNPC(mapnum, MapNpcNum).HP = 0
        MapNPC(mapnum, MapNpcNum).MP = 0
        MapNPC(mapnum, MapNpcNum).SP = 0
        MapNPC(mapnum, MapNpcNum).Dir = 0
        MapNPC(mapnum, MapNpcNum).X = 0
        MapNPC(mapnum, MapNpcNum).Y = 0

        ' Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(mapnum, MapNpcNum).num & SEP_CHAR & MapNpc(mapnum, MapNpcNum).X & SEP_CHAR & MapNpc(mapnum, MapNpcNum).Y & SEP_CHAR & MapNpc(mapnum, MapNpcNum).Dir & SEP_CHAR & Npc(MapNpc(mapnum, MapNpcNum).num).Big & END_CHAR
        ' Call SendDataToMap(mapnum, Packet)
        Call SaveMap(mapnum)
    End If

' MapNpc(mapnum, MapNpcNum).num = 0
' MapNpc(mapnum, MapNpcNum).SpawnWait = GetTickCount
' MapNpc(mapnum, MapNpcNum).HP = 0
' Call SendDataToMap(mapnum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)


    Map(mapnum).Revision = Map(mapnum).Revision + 1

    MapNPC(mapnum, MapNpcNum).num = npcnum
    Map(mapnum).NPC(MapNpcNum) = npcnum

    MapNPC(mapnum, MapNpcNum).Target = 0

    MapNPC(mapnum, MapNpcNum).HP = GetNpcMaxHP(npcnum)
    MapNPC(mapnum, MapNpcNum).MP = GetNpcMaxMP(npcnum)
    MapNPC(mapnum, MapNpcNum).SP = GetNpcMaxSP(npcnum)

    MapNPC(mapnum, MapNpcNum).Dir = Int(Rnd * 4)

    MapNPC(mapnum, MapNpcNum).X = spawn_X
    MapNPC(mapnum, MapNpcNum).Y = spawn_Y

    Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(mapnum, MapNpcNum).num & SEP_CHAR & MapNPC(mapnum, MapNpcNum).X & SEP_CHAR & MapNPC(mapnum, MapNpcNum).Y & SEP_CHAR & MapNPC(mapnum, MapNpcNum).Dir & SEP_CHAR & NPC(MapNPC(mapnum, MapNpcNum).num).Big & END_CHAR
    Call SendDataToMap(mapnum, Packet)

    Call SaveMap(mapnum)

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And GetPlayerMap(I) = mapnum Then
            Call SendDataTo(I, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(I) & SEP_CHAR & Map(GetPlayerMap(I)).Revision & END_CHAR)
        End If
    Next I

End Sub

Sub SpawnItemOnMap(ByVal Map As Long, ByVal X As Long, ByVal Y As Long, ByVal ItemNum As Long, ByVal Amount As Long, ByVal durability As Long)
    Dim I As Long

    If ItemNum > MAX_ITEMS Or ItemNum = 0 Then
        Exit Sub
    End If

    If Map < 0 Or Map > MAX_MAPS Then
        Exit Sub
    End If

    If Amount < 0 Then
        Exit Sub
    End If

    I = FindOpenMapItemSlot(Map)

    If I <> 0 Then

        Call SpawnItemSlot(I, ItemNum, Amount, durability, Map, X, Y)
    Else
        I = MAX_MAP_ITEMS
        Call SpawnItemSlot(I, ItemNum, Amount, durability, Map, X, Y)
    End If
End Sub

Function GetItemName(ByVal Number As Long)
    GetItemName = Item(Number).Name
End Function

Sub ClearItemSlot(ByVal Map As Long, ByVal ItemIndex As Long)
    Call SpawnItemSlot(ItemIndex, 0, 0, 0, Map, MapItem(Map, ItemIndex).X, MapItem(Map, ItemIndex).Y)
End Sub

Sub SendWhosOnline(ByVal Index As Integer)
    Dim S As String
    Dim N As Long, I As Long

    S = vbNullString
    N = 0
    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And I <> Index Then
            S = S & GetPlayerName(I) & ", "
            N = N + 1
        End If
    Next I

    If N = 0 Then
        S = "No hay otros jugadores conectados."
    Else
        S = Mid$(S, 1, Len(S) - 2)
        S = "Hay otros " & N & " jugadores conectados: " & S & "."
    End If

    Call PlayerMsg(Index, S, WhoColor)
End Sub

Sub GoShopping(ByVal Index As Long, ByVal ShopNum As Long)
    Call SendTrade(Index, ShopNum)
End Sub

Public Sub LockPlayer(ByVal Index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(Index).Locked = False
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(Index).Locked = True
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If
End Sub

Function IsPlayerLocked(ByVal Index As Long) As Boolean
    IsPlayerLocked = Player(Index).Locked
End Function

Public Sub SetServerTime(ByVal Time_Hour As Long, ByVal Time_Minute As Long, ByVal Time_Second As Long)
    If Time_Hour <> -1 Then
        If Time_Hour > -1 And Time_Hour < 24 Then
            Hours = Time_Hour
        End If
    End If

    If Time_Minute <> -1 Then
        If Time_Minute > -1 And Time_Minute < 60 Then
            Minutes = Time_Minute
        End If
    End If

    If Time_Second <> -1 Then
        If Time_Second > -1 And Time_Second < 60 Then
            Seconds = Time_Second
        End If
    End If
End Sub

Function GetServerTime(ByVal TimeType As Long) As Long
    Select Case TimeType
        Case 0
            GetServerTime = Hours

        Case 1
            GetServerTime = Minutes

        Case 2
            GetServerTime = Seconds
    End Select
End Function

Function GetInvItemslot(Index, F_item_num, F_Min_value, F_Min_dur)
    Dim F_slot

    F_slot = 1
    GetInvItemslot = 0

    Do While F_slot < 25
        If GetPlayerInvItemNum(Index, F_slot) = F_item_num Then
            If GetPlayerInvItemValue(Index, F_slot) >= F_Min_value Then
                If GetPlayerInvItemDur(Index, F_slot) >= F_Min_dur Then

                    GetInvItemslot = F_slot
                    F_slot = 25

                End If
            End If
        End If
        F_slot = F_slot + 1
    Loop

End Function

Sub giveplayeritemtoslot(Index, F_slot, F_num, F_Val, F_dur)
    Dim Packet

    If F_num < 0 Or F_num > MAX_ITEMS Or F_Val < 0 Or F_slot < 1 Or F_slot > 24 Then
        Exit Sub
    End If

    Player(Index).Char(Player(Index).CharNum).Inv(F_slot).num = F_num
    Player(Index).Char(Player(Index).CharNum).Inv(F_slot).Value = F_Val
    Player(Index).Char(Player(Index).CharNum).Inv(F_slot).Dur = F_dur

    Packet = "PLAYERINVUPDATE" & SEP_CHAR & F_slot & SEP_CHAR & Index & SEP_CHAR & GetPlayerInvItemNum(Index, F_slot) & SEP_CHAR & GetPlayerInvItemValue(Index, F_slot) & SEP_CHAR & GetPlayerInvItemDur(Index, F_slot) & SEP_CHAR & Index & END_CHAR
    Call SendDataToMap(GetPlayerMap(Index), Packet)

End Sub

Function giveplayeritem(Index, F_num, F_Val, F_dur)
    Dim N

    If F_num < 0 Or F_num > MAX_ITEMS Or F_Val < 0 Then
        Exit Function
    End If

    giveplayeritem = 0
    N = FindOpenInvSlot(Index, F_num)

    If N <> 0 Then
        giveplayeritem = 1
        ' Set item in players inventor
        Call SetPlayerInvItemNum(Index, N, F_num)
        If Item(GetPlayerInvItemNum(Index, N)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, N)).Stackable = 1 Then
            Call SetPlayerInvItemValue(Index, N, GetPlayerInvItemValue(Index, N) + F_Val)
        Else
            Call SetPlayerInvItemValue(Index, N, 0)
        End If
        Call SetPlayerInvItemDur(Index, N, F_dur)
        Call SendInventoryUpdate(Index, N)
        Exit Function
    Else
        giveplayeritem = 0
        Exit Function
    End If
End Function

Sub textbubble(player_index, bubble_index, F_text, F_map, F_X, F_Y, F_colour)
    Dim Packet

    If player_index < 1 Or player_index > MAX_PLAYERS Or bubble_index < 1 Or bubble_index > 20 Or F_map < 1 Or F_map > MAX_MAPS Or F_colour < 0 Or F_colour > 15 Then
        Exit Sub
    End If

    Packet = "scriptbubble" & SEP_CHAR & bubble_index & SEP_CHAR & F_text & SEP_CHAR & F_map & SEP_CHAR & F_X & SEP_CHAR & F_Y & SEP_CHAR & F_colour & END_CHAR
    Call SendDataTo(player_index, Packet)

End Sub

Sub maptextbubble(bubble_index, F_text, F_map, F_X, F_Y, F_colour)
    Dim Packet
    Dim I

    If bubble_index < 1 Or bubble_index > 20 Or F_map < 1 Or F_map > MAX_MAPS Or F_colour < 0 Or F_colour > 15 Then
        Exit Sub
    End If

    Packet = "scriptbubble" & SEP_CHAR & bubble_index & SEP_CHAR & F_text & SEP_CHAR & F_map & SEP_CHAR & F_X & SEP_CHAR & F_Y & SEP_CHAR & F_colour & END_CHAR

    For I = 1 To MAX_PLAYERS
        If GetPlayerMap(I) = F_map Then
            Call SendDataTo(I, Packet)
        End If
    Next I

End Sub

Function GetMapName(F_map)
    GetMapName = Trim(Map(F_map).Name)
End Function

Function GetMapUp(F_map)
    GetMapUp = Map(F_map).Up
End Function

Function GetMapDown(F_map)
    GetMapDown = Map(F_map).Down
End Function

Function GetMapLeft(F_map)
    GetMapLeft = Map(F_map).Left
End Function

Function GetMapRight(F_map)
    GetMapRight = Map(F_map).Right
End Function

Sub CustomMenuShow(ByVal Index As Long, ByVal Title As String, ByVal filename As String, ByVal closable As Long)
    Dim Packet As String

    If Index <= 0 Or Index >= MAX_PLAYERS Or closable < 0 Or closable > 1 Then
        Exit Sub
    End If

    Packet = "showcustommenu" & SEP_CHAR & Title & SEP_CHAR & filename & SEP_CHAR & closable & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub CustomMenuClose(ByVal Index As Long)
    Dim Packet As String

    Packet = "closecustommenu" & END_CHAR
    Call SendDataTo(Index, Packet)

End Sub

Sub CustomMenuPicture(ByVal player_index As Long, ByVal picture_index As Long, ByVal filename As String, ByVal Left As Long, ByVal Top As Long)
    Dim Packet As String

    If picture_index < 0 Or player_index <= 0 Or player_index > MAX_PLAYERS Then
        Exit Sub
    End If

    Packet = "loadpiccustommenu" & SEP_CHAR & picture_index & SEP_CHAR & filename & SEP_CHAR & Left & SEP_CHAR & Top & END_CHAR
    Call SendDataTo(player_index, Packet)

End Sub

Sub CustomMenuLabel(ByVal player_index As Long, ByVal picture_index As Long, ByVal Caption As String, ByVal Left As Long, ByVal Top As Long, ByVal customsize As Long, ByVal customcolour As Long, ByVal Alignment As Long, ByVal Width As Long, ByVal Height As Long)
    Dim Packet As String

    If Alignment < 0 Or Alignment > 2 Or Width < 0 Or Height < 0 Or customcolour > 15 Or customcolour < 0 Or picture_index < 0 Or player_index <= 0 Or player_index > MAX_PLAYERS Then
        Exit Sub
    End If

    Packet = "loadlabelcustommenu" & SEP_CHAR & picture_index & SEP_CHAR & Caption & SEP_CHAR & Left & SEP_CHAR & Top & SEP_CHAR & customsize & SEP_CHAR & customcolour & SEP_CHAR & Alignment & SEP_CHAR & Width & SEP_CHAR & Height & END_CHAR
    Call SendDataTo(player_index, Packet)

End Sub

Sub CustomMenuTextBox(ByVal player_index As Long, ByVal customIndex As Long, ByVal Width As Long, ByVal Left As Long, ByVal Top As Long, ByVal Text As String)
    Dim Packet As String

    If customIndex < 0 Or player_index <= 0 Or player_index > MAX_PLAYERS Then
        Exit Sub
    End If

    Packet = "loadtextboxcustommenu" & SEP_CHAR & customIndex & SEP_CHAR & Left & SEP_CHAR & Width & SEP_CHAR & Top & SEP_CHAR & Text & END_CHAR
    Call SendDataTo(player_index, Packet)

End Sub

Function getplayermenuclicktitle(ByVal player_index As Long)
    Dim Msg As String

    If player_index <= 0 Or player_index > MAX_PLAYERS Then
        Exit Function
    End If

    Msg = Player(player_index).CustomTitle

    getplayermenuclicktitle = Trim$(Msg)

End Function

Function getplayermenuclickmsg(ByVal player_index As Long)
    Dim Msg As String

    If player_index <= 0 Or player_index > MAX_PLAYERS Then
        Exit Function
    End If

    Msg = Player(player_index).CustomMsg

    getplayermenuclickmsg = Trim$(Msg)
End Function

Sub Loadinternet(ByVal Index As Long, ByVal address As String)
    Dim Packet As String

    If Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Sub
    End If

    Packet = "loadinternetwindow" & SEP_CHAR & address & END_CHAR
    Call SendDataTo(Index, Packet)

End Sub

Sub updateplayermenutext(ByVal player_index As Long, ByVal box_index As Long)
    Dim Packet As String

    Packet = "returncustomboxmsg" & SEP_CHAR & box_index & END_CHAR
    Call SendDataTo(player_index, Packet)

End Sub

Function getplayermenutext(ByVal player_index As Long)

    getplayermenutext = Player(player_index).CustomMsg

End Function

Sub npcmoving(ByVal mapnum As Long, ByVal MapNpcNum As Long, ByVal Direction As Long, ByVal Speed As Long)

    If CanNpcMove(mapnum, MapNpcNum, Direction) Then
        Call NpcMove(mapnum, MapNpcNum, Direction, Speed)
    End If

End Sub


Function GetPlayerDirX(Index)

    If GetPlayerDir(Index) = 1 Then
        GetPlayerDirX = GetPlayerX(Index)
    ElseIf GetPlayerDir(Index) = 2 Then
        GetPlayerDirX = GetPlayerX(Index) - 1
    ElseIf GetPlayerDir(Index) = 3 Then
        GetPlayerDirX = GetPlayerX(Index) + 1
    ElseIf GetPlayerDir(Index) = 0 Then
        GetPlayerDirX = GetPlayerX(Index)
    End If

End Function


Function GetPlayerDirY(Index)

    If GetPlayerDir(Index) = 1 Then
        GetPlayerDirY = GetPlayerY(Index) + 1
    ElseIf GetPlayerDir(Index) = 2 Then
        GetPlayerDirY = GetPlayerY(Index)
    ElseIf GetPlayerDir(Index) = 3 Then
        GetPlayerDirY = GetPlayerY(Index)
    ElseIf GetPlayerDir(Index) = 0 Then
        GetPlayerDirY = GetPlayerY(Index) - 1
    End If

End Function

Sub SetSpeed(ByVal Index As Long, ByVal Movement As String, ByVal Speed As Long)
    If Speed < 1 Then
        Call PlayerMsg(Index, "Automated message: script error, value " & Speed & " at SetSpeed is too low - must be above 1.", Red)
    ElseIf Speed > 32 Then
        Call PlayerMsg(Index, "Automated message: script error, value " & Speed & " at SetSpeed is to high - must be below 32.", Red)
    Else
        If Int(Log(Speed) / Log(2)) <> Log(Speed) / Log(2) Then
            Exit Sub
        End If

        Call SendDataTo(Index, "setspeed" & SEP_CHAR & Movement & SEP_CHAR & Speed & END_CHAR)
    End If
End Sub


Function GetNpcX(ByVal mapnum As Long, ByVal MapNpcNum As Long)

    If MapNpcNum < 1 Or MapNpcNum > 25 Then
    Else
        GetNpcX = MapNPC(mapnum, MapNpcNum).X
    End If

End Function


Function GetPlayerTargetNpc(ByVal Index As Long)

    If Index > 0 Then
        If Player(Index).TargetType = TARGET_TYPE_NPC Then
            GetPlayerTargetNpc = Player(Index).TargetNPC
        Else
            GetPlayerTargetNpc = -1
        End If
    End If

End Function


Function GetNpcY(ByVal mapnum As Long, ByVal MapNpcNum As Long)

    If MapNpcNum > 0 Then
        GetNpcY = MapNPC(mapnum, MapNpcNum).Y
    End If

End Function

Sub SetWeather(ByVal mapnum As Long, ByVal Weather As Long, ByVal Interval As Long)
    Call SendDataToMap(mapnum, "mapweather" & SEP_CHAR & mapnum & SEP_CHAR & Weather & SEP_CHAR & Interval & END_CHAR)
End Sub

Function GetWeather(ByVal mapnum As Long)
    GetWeather = Map(mapnum).Weather
End Function

Sub Image(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal Tile As Long, ByVal Top As Long, ByVal Height As Long, ByVal Left As Long, ByVal Width As Long)
    Dim Packet As String
    Packet = "fog" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & Tile & SEP_CHAR & Top & SEP_CHAR & Height & SEP_CHAR & Left & SEP_CHAR & Width & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub
Sub SetPlayerNameColor(ByVal Index As Long, ByVal Red As String, ByVal Green As String, ByVal Blue As String)
If CByte(Red) < 0 Then Exit Sub
If CByte(Red) > 255 Then Exit Sub
If CByte(Green) < 0 Then Exit Sub
If CByte(Green) > 255 Then Exit Sub
If CByte(Blue) < 0 Then Exit Sub
If CByte(Blue) > 255 Then Exit Sub

    Player(Index).Char(Player(Index).CharNum).Color(1) = Red
    Player(Index).Char(Player(Index).CharNum).Color(2) = Green
    Player(Index).Char(Player(Index).CharNum).Color(3) = Blue
    Call SendDataToMap(GetPlayerMap(Index), "namecolor" & SEP_CHAR & Index & SEP_CHAR & Red & SEP_CHAR & Green & SEP_CHAR & Blue & END_CHAR)
End Sub
Function GetPlayerNameColor(ByVal Index As Long) As String
GetPlayerNameColor = ""

If Player(Index).Char(Player(Index).CharNum).Color(1) < 0 Or Player(Index).Char(Player(Index).CharNum).Color(1) > 255 Then
    GetPlayerNameColor = ""
Else
    GetPlayerNameColor = Player(Index).Char(Player(Index).CharNum).Color(1)
    If Player(Index).Char(Player(Index).CharNum).Color(2) < 0 Or Player(Index).Char(Player(Index).CharNum).Color(2) > 255 Then
        GetPlayerNameColor = ""
    Else
        GetPlayerNameColor = GetPlayerNameColor & "-" & Player(Index).Char(Player(Index).CharNum).Color(2)
        If Player(Index).Char(Player(Index).CharNum).Color(3) < 0 Or Player(Index).Char(Player(Index).CharNum).Color(3) > 255 Then
            GetPlayerNameColor = ""
        Else
            GetPlayerNameColor = GetPlayerNameColor & "-" & Player(Index).Char(Player(Index).CharNum).Color(3)
        End If
    End If
End If
End Function

Sub ResetPlayerColor(ByVal Index As Long)
    Select Case GetPlayerAccess(Index)
        Case 0
            Call SetPlayerColor(Index, UserCr(1), UserCr(2), UserCr(3))
            Exit Sub
                
        Case 1
            Call SetPlayerColor(Index, ModCr(1), ModCr(2), ModCr(3))
            Exit Sub
                
        Case 2
            Call SetPlayerColor(Index, MapperCr(1), MapperCr(2), MapperCr(3))
            Exit Sub

        Case 3
            Call SetPlayerColor(Index, DeveloperCr(1), DeveloperCr(2), DeveloperCr(3))
            Exit Sub
                
        Case 4
            Call SetPlayerColor(Index, AdminCr(1), AdminCr(2), AdminCr(3))
            Exit Sub
                
        Case 5
            Call SetPlayerColor(Index, OwnerCr(1), OwnerCr(2), OwnerCr(3))
            Exit Sub
    End Select
    Call SendPlayerData(Index)
End Sub
Sub LockSpells(ByVal Index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(Index).LockedSpells = False
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(Index).LockedSpells = True
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If
End Sub

Sub LockItems(ByVal Index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(Index).LockedItems = False
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(Index).LockedItems = True
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If
End Sub

Sub LockAttack(ByVal Index As Long, ByVal Locked As Long)
    If Locked = 0 Then
        Player(Index).LockedAttack = False
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If

    If Locked = 1 Then
        Player(Index).LockedAttack = True
        Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(Index) & SEP_CHAR & Map(GetPlayerMap(Index)).Revision & END_CHAR)
    End If
End Sub

Function GetIndexPlayerOnMap(ByVal Map As Long, ByVal X As Long, ByVal Y As Long)
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) And GetPlayerMap(I) = Map Then
            If GetPlayerX(I) = X And GetPlayerY(I) = Y Then
                GetIndexPlayerOnMap = I
                Exit Function
            End If
        End If
    Next I
End Function

Function GetPlayerHead(ByVal Index As Long)
    GetPlayerHead = Player(Index).Char(Player(Index).CharNum).Head
End Function

Public Sub SetPlayerHead(ByVal Index As Long, ByVal Head As Long)
    Player(Index).Char(Player(Index).CharNum).Head = Head
End Sub

Function GetPlayerBody(ByVal Index As Long)
    GetPlayerBody = Player(Index).Char(Player(Index).CharNum).Body
End Function

Public Sub SetPlayerBody(ByVal Index As Long, ByVal Body As Long)
    Player(Index).Char(Player(Index).CharNum).Body = Body
End Sub

Function GetPlayerleg(ByVal Index As Long)
    GetPlayerleg = Player(Index).Char(Player(Index).CharNum).Leg
End Function

Public Sub SetPlayerLeg(ByVal Index As Long, ByVal Leg As Long)
    Player(Index).Char(Player(Index).CharNum).Leg = Leg
End Sub

Public Sub MovePlayer(ByVal Index As Long, ByVal Direction As Long, ByVal Movement As Long)
    Dim X As Integer
    Dim Y As Integer
   
    Y = GetPlayerY(Index)
    X = GetPlayerX(Index)
   
    Select Case Direction
        Case DIR_UP
            Y = Y - 1
        Case DIR_DOWN
            Y = Y + 1
        Case DIR_LEFT
            X = X - 1
        Case DIR_RIGHT
            X = X + 1
   End Select
   
    Call PlayerMove(Index, Direction, Movement, X, Y)
    Call SendPlayerNewXY(Index)
End Sub

Public Sub SavePlayer(ByVal Index As Long)
    Call SavePlayer(Index)
End Sub

Function GetPlayerGender(ByVal Index As Long) As Long
    GetPlayerGender = Player(Index).Char(Player(Index).CharNum).Sex
End Function

Function GetPlayerPaperdoll(ByVal Index As Long) As Byte
    If Player(Index).InGame Then
        GetPlayerPaperdoll = Player(Index).Char(GetPlayerCharNum(Index)).paperdoll
    End If
End Function

Public Sub SetPlayerPaperdoll(ByVal Index As Long, ByVal Mode As Byte)
    If Mode = 0 Or Mode = 1 Then
        If Player(Index).InGame Then
            Player(Index).Char(GetPlayerCharNum(Index)).paperdoll = Mode
        End If
    End If
End Sub

Public Sub ChangeMapName(ByVal Index As Long, ByVal Name As String)
    Map(Index).Name = Name
    Call SaveMap(Index)
    Call SendMap(Index)
End Sub

' The subs below are added by DFA
Public Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    Call modGameLogic.PlayerMapDropItem(Index, InvNum, Amount)
End Sub

Public Function ScriptGetTickCount() As Long
    ScriptGetTickCount = GetTickCount
End Function


' Eclipse 2.7 Scripting Commands

Public Function GetMaxMapX() As Byte
    GetMaxMapX = MAX_MAPX
End Function

Public Function GetMaxMapY() As Byte
    GetMaxMapY = MAX_MAPY
End Function

Public Sub SetMapMoral(ByVal mapnum As Long, ByVal Moral As Long)
    Map(mapnum).Moral = Moral
End Sub

Public Function GetMapMoral(ByVal mapnum As Long) As Byte
    GetMapMoral = Map(mapnum).Moral
End Function

Public Sub ShopReload(ByVal ShopNum As Long)
    Dim FileID As Integer
    Dim filename As String
    
    filename = App.Path & "\Tiendas\Shop" & ShopNum & ".dat"

    FileID = FreeFile

    Open filename For Binary As #FileID
        Get #FileID, , Shop(ShopNum)
    Close #FileID
End Sub

Sub DamageNPC(ByVal Index As Long, ByVal npcnum As Long, ByVal Damage As Long)
    Call AttackNpc(Index, npcnum, Damage)
    Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & npcnum & END_CHAR)
End Sub

Sub DamagePlayer(ByVal Index As Long, ByVal PIndex As Long, ByVal Damage As Long)
    Call AttackPlayer(Index, PIndex, Damage)
End Sub

Sub NPCAttack(ByVal X As Long, ByVal Target As Long, ByVal Damage As Long)
    Call NpcAttackPlayer(X, Target, Damage)
End Sub

Sub FireArrow(ByVal Index As Long, ByVal arrow As Long, ByVal Dir As Long)
    Call SendDataToMap(GetPlayerMap(Index), "checkarrows" & SEP_CHAR & Index & SEP_CHAR & arrow & SEP_CHAR & Dir & END_CHAR)
End Sub

Function GetTargetSTR(ByVal Index As Long) As Long
Dim Target

Target = GetPlayerTarget(Index)

If Target > 0 Then
GetTargetSTR = Player(Target).Char(Player(Target).CharNum).STR
Else
GetTargetSTR = 0
End If
End Function

Function GetTargetDEF(ByVal Index As Long) As Long
Dim Target

Target = GetPlayerTarget(Index)

If Target > 0 Then
GetTargetDEF = Player(Target).Char(Player(Target).CharNum).DEF
Else
GetTargetDEF = 0
End If
End Function

Function GetTargetMAGI(ByVal Index As Long) As Long
Dim Target

Target = GetPlayerTarget(Index)

If Target > 0 Then
GetTargetMAGI = Player(Target).Char(Player(Target).CharNum).Magi
Else
GetTargetMAGI = 0
End If
End Function

Function GetTargetSPEED(ByVal Index As Long) As Long
Dim Target

Target = GetPlayerTarget(Index)

If Target > 0 Then
GetTargetSPEED = Player(Target).Char(Player(Target).CharNum).Speed
Else
GetTargetSPEED = 0
End If
End Function

Function GetTargetJob(ByVal Index As Long) As String
Dim Target

Target = GetPlayerTarget(Index)

If Target > 0 Then
GetTargetJob = Trim$(ClassData(GetPlayerClass(Target)).Name)
Else
GetTargetJob = 0
End If
End Function
           

Function GetTargetName(ByVal Index As Long) As String
Dim Target

Target = GetPlayerTarget(Index)

If Target > 0 Then
GetTargetName = Trim$(Player(Target).Char(Player(Target).CharNum).Name)
Else
GetTargetName = 0
End If
End Function

Function GetTargetGuild(ByVal Index As Long) As String
Dim Target

Target = GetPlayerTarget(Index)

If Target > 0 Then
GetTargetGuild = Trim$(Player(Target).Char(Player(Target).CharNum).Guild)
Else
GetTargetGuild = 0
End If
End Function

Public Function GetPetHp(ByVal Index As Long) As Long
    GetPetHp = Player(Index).Pet.HP
End Function

Public Sub SetPetHp(ByVal Index As Long, ByVal HP As Long)
    Player(Index).Pet.HP = HP
     
End Sub

Public Function GetPetLevel(ByVal Index As Long) As Long
    GetPetLevel = Player(Index).Pet.Level
End Function

Public Sub SetPetLevel(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Pet.Level = Level
     
End Sub

Public Function GetPetSTR(ByVal Index As Long) As Long
    GetPetSTR = Player(Index).Pet.STR

End Function

Public Sub SetPetSTR(ByVal Index As Long, ByVal STR As Long)
    Player(Index).Pet.STR = STR
     
End Sub

Public Function GetPetDEF(ByVal Index As Long) As Long
    GetPetDEF = Player(Index).Pet.DEF

End Function

Public Function GetPetSPEED(ByVal Index As Long) As Long
    GetPetSPEED = Player(Index).Pet.Speed

End Function

Public Sub SetPetSPEED(ByVal Index As Long, ByVal Speed As Long)
    Player(Index).Pet.Speed = Speed
     
End Sub

Public Function GetPetMAGI(ByVal Index As Long) As Long
    GetPetMAGI = Player(Index).Pet.Magi
     
End Function

Public Sub SetPetMAGI(ByVal Index As Long, ByVal Magi As Long)
    Player(Index).Pet.Magi = Magi
     
End Sub

Public Sub SetPetDEF(ByVal Index As Long, ByVal DEF As Long)
    Player(Index).Pet.DEF = DEF
     
End Sub

Public Function GetPetX(ByVal Index As Long) As Long
GetPetX = Player(Index).Pet.X
End Function

Public Function GetPetY(ByVal Index As Long) As Long
GetPetY = Player(Index).Pet.Y
End Function

Public Function SetPetToGo(ByVal Index As Long, ByVal X As Long, ByVal Y As Long)
Player(Index).Pet.XToGo = X
Player(Index).Pet.YToGo = Y
 
End Function

Public Function IsPetAliveOnLogin(ByVal Index As Long) As Long
IsPetAliveOnLogin = Player(Index).Char(Player(Index).CharNum).PetAlive
End Function


Public Sub SetPlayerPetAlive(ByVal Index As Long, ByVal PetAlive As Long)
    Player(Index).Char(Player(Index).CharNum).PetAlive = PetAlive
     
End Sub

Public Function GetPlayerPetAlive(ByVal Index As Long) As Long
    GetPlayerPetAlive = Player(Index).Char(Player(Index).CharNum).PetAlive
End Function


' Old Functions
Public Function GetPetsprite(ByVal Index As Long) As Long
    GetPetsprite = Player(Index).Pet.SPRITE
End Function
Public Function setPetsprite(ByVal Index As Long, ByVal PetSprite As Long) As Long
    Player(Index).Pet.SPRITE = PetSprite
End Function

Public Function GetPetPOINTS(ByVal Index As Long) As Long
    GetPetPOINTS = Player(Index).Pet.POINTS
End Function

Public Sub SetPetPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).Pet.POINTS = POINTS
     
End Sub

Public Function GetPetMAXHP(ByVal Index As Long) As Long
    GetPetMAXHP = Player(Index).Pet.STR * 5
End Function

Public Function GetPetMAXSP(ByVal Index As Long) As Long
    GetPetMAXSP = Player(Index).Pet.Speed * 5
End Function

Public Function GetPetMAXMP(ByVal Index As Long) As Long
    GetPetMAXMP = Player(Index).Pet.Magi * 5
End Function

Public Function GetPetMAXFP(ByVal Index As Long) As Long
    GetPetMAXFP = 100
End Function

Public Function GetPetSP(ByVal Index As Long) As Long
    GetPetSP = Player(Index).Pet.SP
End Function

Public Sub SetPetSP(ByVal Index As Long, ByVal SP As Long)
    Player(Index).Pet.SP = SP
     
End Sub

Public Function GetPetMP(ByVal Index As Long) As Long
    GetPetMP = Player(Index).Pet.MP
End Function

Public Sub SetPetMP(ByVal Index As Long, ByVal MP As Long)
    Player(Index).Pet.MP = MP
     
End Sub

Public Function GetPetFP(ByVal Index As Long) As Long
    GetPetFP = Player(Index).Pet.FP
End Function

Public Sub SetPetFP(ByVal Index As Long, ByVal FP As Long)
    Player(Index).Pet.FP = FP
     
End Sub

Public Function GetPetExp(ByVal Index As Long) As Long
    GetPetExp = Player(Index).Pet.EXP
End Function

Public Sub SetPetExp(ByVal Index As Long, ByVal EXP As Long)
    Player(Index).Pet.EXP = EXP
     
End Sub

Public Function GetPetNextLevel(ByVal Index As Long) As Long
If Player(Index).Pet.Alive = NO Then Exit Function
    GetPetNextLevel = Experience(GetPetLevel(Index))
End Function

Function GetName(ByVal Index, CharNum) As String
GetName = Trim$(Player(Index).Char(Player(Index).CharNum).Name)
End Function

Sub SendPetData(ByVal Index)
Dim Packet As String

If Index <= 0 Or Index > MAX_PLAYERS Then
    Exit Sub
End If

If Player(Index).Pet.Alive = NO Then
    Exit Sub
End If

        Packet = "PETDATA" & SEP_CHAR
        Packet = Packet & Index & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Map & SEP_CHAR
        Packet = Packet & Player(Index).Pet.X & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Y & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Dir & SEP_CHAR
        Packet = Packet & Player(Index).Pet.SPRITE & SEP_CHAR
        Packet = Packet & Player(Index).Pet.HP & SEP_CHAR
        Packet = Packet & Player(Index).Pet.STR * 5 & SEP_CHAR
        Packet = Packet & Player(Index).Pet.STR & SEP_CHAR
        Packet = Packet & Player(Index).Pet.DEF & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Speed & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Magi & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Level & SEP_CHAR
        Packet = Packet & Player(Index).Pet.POINTS & SEP_CHAR
        Packet = Packet & GetPetSP(Index) & SEP_CHAR
        Packet = Packet & GetPetMAXSP(Index) & SEP_CHAR
        Packet = Packet & GetPetMP(Index) & SEP_CHAR
        Packet = Packet & GetPetMAXMP(Index) & SEP_CHAR
        Packet = Packet & GetPetFP(Index) & SEP_CHAR
        Packet = Packet & GetPetMAXFP(Index) & SEP_CHAR
        Packet = Packet & GetPetExp(Index) & SEP_CHAR
        Packet = Packet & GetPetNextLevel(Index) & SEP_CHAR
        Packet = Packet & Player(Index).Pet.Name & SEP_CHAR
        Packet = Packet & END_CHAR
            Call SendDataToMap(GetPlayerMap(Index), Packet)
End Sub

Sub savepet(ByVal Index As Long)
Call SetPlayerPetAlive(Index, 1)
Call SetPlayerPetSprite(Index, Player(Index).Pet.SPRITE)
Call SetPlayerPetHP(Index, Player(Index).Pet.HP)
Call SetPlayerPetSP(Index, Player(Index).Pet.SP)
Call SetPlayerPetMP(Index, Player(Index).Pet.MP)
Call SetPlayerPetFP(Index, Player(Index).Pet.FP)
Call SetPlayerPetSTR(Index, Player(Index).Pet.STR)
Call SetPlayerPetDEF(Index, Player(Index).Pet.DEF)
Call SetPlayerPetMAGI(Index, Player(Index).Pet.Magi)
Call SetPlayerPetSPEED(Index, Player(Index).Pet.Speed)
Call SetPlayerPetLevel(Index, Player(Index).Pet.Level)
Call SetPlayerPetEXP(Index, Player(Index).Pet.EXP)
Call SetPlayerPetPOINTS(Index, Player(Index).Pet.POINTS)
Call SetPlayerPetNAME(Index, Player(Index).Pet.Name)
Call SetPlayerPetTNL(Index, GetPetNextLevel(Index))
Call PlayerMsg(Index, "[Mascota] Tu mascota se guardo!", BRIGHTGREEN)
End Sub
