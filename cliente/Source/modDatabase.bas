Attribute VB_Name = "modDatabase"
Option Explicit

Sub SaveLocalMap(ByVal MapNum As Long)
    Dim filename As String
    Dim F As Long
    'Call SendData("requestspass" & END_CHAR)
    filename = App.Path & "\Mapas\map" & MapNum & ".dat"
    'If GetSetting(App.EXEName, "Clave", "Clave", "") = SPassWord Then
    'Call EncryptarMap(MapNum)
    'Else
    'Call SendData("needmap" & SEP_CHAR & "YES" & END_CHAR)
    'Call SaveSetting(App.EXEName, "Clave", "Clave", SPassWord)
    'Exit Sub
    'End If
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Map(MapNum)
    Close #F
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean
    If Not RAW Then
        If LenB(Dir(App.Path & "\" & filename)) > 0 Then
            FileExist = True
        End If
    Else
        If LenB(Dir(filename)) > 0 Then
            FileExist = True
        End If
    End If
End Function

Sub LoadMap(ByVal MapNum As Long)
    Dim filename As String
    Dim F As Long

    filename = App.Path & "\Mapas\map" & MapNum & ".dat"
    'Call SendData("requestspass" & END_CHAR) Tuve que sacar este pedido de PassWord por que me tildaba la coneccion :S,pero funciona perfecto
    If FileExists("mapas\map" & MapNum & ".dat") = False Then
        Exit Sub
    End If
    'If GetSetting(App.EXEName, "Clave", "Clave") = SPassWord Then
    F = FreeFile
    Open filename For Binary As #F
    Get #F, , Map(MapNum)
    Close #F
    'Call DesencryptarMap(MapNum)
    'Else
    'Call SendData("needmap" & SEP_CHAR & "YES" & END_CHAR)
    'Call SaveSetting(App.EXEName, "Clave", "Clave", SPassWord)
    'End If
End Sub

Sub LeerLibro(ByVal index As Long, ByVal QueLeer As String)
frmLibro.Text1.Text = QueLeer
frmLibro.Show
End Sub

Function GetMapRevision(ByVal MapNum As Long) As Long
    GetMapRevision = Map(MapNum).Revision
End Function

Sub ClearTempTile()
    Dim X As Long, Y As Long

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            TempTile(X, Y).DoorOpen = NO
        Next X
    Next Y
End Sub

Sub ClearPlayer(ByVal index As Long)
    Dim I As Long
    Dim n As Long

    Player(index).name = vbNullString
    Player(index).Guild = vbNullString
    Player(index).Guildaccess = 0
    Player(index).Class = 0
    Player(index).Level = 0
    Player(index).Sprite = 0
    Player(index).Exp = 0
    Player(index).Access = 0
    Player(index).PK = NO

    Player(index).HP = 0
    Player(index).MP = 0
    Player(index).SP = 0

    Player(index).STR = 0
    Player(index).DEF = 0
    Player(index).SPEED = 0
    Player(index).MAGI = 0

    For n = 1 To MAX_INV
        Player(index).Inv(n).Num = 0
        Player(index).Inv(n).value = 0
        Player(index).Inv(n).Dur = 0
    Next n

    For n = 1 To MAX_BANK
        Player(index).Bank(n).Num = 0
        Player(index).Bank(n).value = 0
        Player(index).Bank(n).Dur = 0
    Next n

    Player(index).ArmorSlot = 0
    Player(index).WeaponSlot = 0
    Player(index).HelmetSlot = 0
    Player(index).ShieldSlot = 0
    Player(index).LegsSlot = 0
    Player(index).RingSlot = 0
    Player(index).ArmorSlot = 0

    Player(index).Map = 0
    Player(index).X = 0
    Player(index).Y = 0
    Player(index).Dir = 0

    ' Client use only
    Player(index).MaxHP = 0
    Player(index).MaxMP = 0
    Player(index).MaxSP = 0
    Player(index).XOffset = 0
    Player(index).YOffset = 0
    Player(index).MovingH = 0
    Player(index).MovingV = 0
    Player(index).Moving = 0
    Player(index).Attacking = 0
    Player(index).AttackTimer = 0
    Player(index).MapGetTimer = 0
    Player(index).CastedSpell = NO
    Player(index).EmoticonNum = -1
    Player(index).EmoticonTime = 0
    Player(index).EmoticonVar = 0

    For I = 1 To MAX_SPELL_ANIM
        Player(index).SpellAnim(I).CastedSpell = NO
        Player(index).SpellAnim(I).SpellTime = 0
        Player(index).SpellAnim(I).SpellVar = 0
        Player(index).SpellAnim(I).SpellDone = 0

        Player(index).SpellAnim(I).Target = 0
        Player(index).SpellAnim(I).TargetType = 0
    Next I

    Player(index).Spellnum = 0

    For I = 1 To MAX_BLT_LINE
        BattlePMsg(I).index = 1
        BattlePMsg(I).time = I
        BattleMMsg(I).index = 1
        BattleMMsg(I).time = I
    Next I

    Inventory = 1
End Sub

Sub ClearItem(ByVal index As Long)
    Item(index).name = vbNullString
    Item(index).desc = vbNullString

    Item(index).Type = 0
    Item(index).Data1 = 0
    Item(index).Data2 = 0
    Item(index).Data3 = 0
    Item(index).StrReq = 0
    Item(index).DefReq = 0
    Item(index).SpeedReq = 0
    Item(index).MagicReq = 0
    Item(index).ClassReq = -1
    Item(index).AccessReq = 0

    Item(index).AddHP = 0
    Item(index).AddMP = 0
    Item(index).AddSP = 0
    Item(index).AddStr = 0
    Item(index).AddDef = 0
    Item(index).AddMagi = 0
    Item(index).AddSpeed = 0
    Item(index).AddEXP = 0
    Item(index).AttackSpeed = 1000
    Item(index).Stackable = 0
End Sub

Sub ClearItems()
    Dim I As Long

    For I = 1 To MAX_ITEMS
        Call ClearItem(I)
    Next I
End Sub

Sub ClearMapItem(ByVal index As Long)
    MapItem(index).Num = 0
    MapItem(index).value = 0
    MapItem(index).Dur = 0
    MapItem(index).X = 0
    MapItem(index).Y = 0
End Sub

Sub ClearMap()
    Dim I As Long
    Dim X As Long
    Dim Y As Long

    For I = 1 To MAX_MAPS
        Map(I).name = vbNullString
        Map(I).Revision = 0
        Map(I).Moral = 0
        Map(I).Up = 0
        Map(I).Down = 0
        Map(I).Left = 0
        Map(I).Right = 0
        Map(I).Indoors = 0

        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Map(I).Tile(X, Y).Ground = 0
                Map(I).Tile(X, Y).mask = 0
                Map(I).Tile(X, Y).Anim = 0
                Map(I).Tile(X, Y).Mask2 = 0
                Map(I).Tile(X, Y).M2Anim = 0
                Map(I).Tile(X, Y).Fringe = 0
                Map(I).Tile(X, Y).FAnim = 0
                Map(I).Tile(X, Y).Fringe2 = 0
                Map(I).Tile(X, Y).F2Anim = 0
                Map(I).Tile(X, Y).Type = 0
                Map(I).Tile(X, Y).Data1 = 0
                Map(I).Tile(X, Y).Data2 = 0
                Map(I).Tile(X, Y).Data3 = 0
                Map(I).Tile(X, Y).String1 = vbNullString
                Map(I).Tile(X, Y).String2 = vbNullString
                Map(I).Tile(X, Y).String3 = vbNullString
                Map(I).Tile(X, Y).Light = 0
                Map(I).Tile(X, Y).GroundSet = 0
                Map(I).Tile(X, Y).MaskSet = 0
                Map(I).Tile(X, Y).AnimSet = 0
                Map(I).Tile(X, Y).Mask2Set = 0
                Map(I).Tile(X, Y).M2AnimSet = 0
                Map(I).Tile(X, Y).FringeSet = 0
                Map(I).Tile(X, Y).FAnimSet = 0
                Map(I).Tile(X, Y).Fringe2Set = 0
                Map(I).Tile(X, Y).F2AnimSet = 0
            Next X
        Next Y
    Next I
End Sub

Sub ClearMapItems()
    Dim X As Long

    For X = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(X)
    Next X
End Sub

Sub ClearMapNpc(ByVal index As Long)
    MapNpc(index).Num = 0
    MapNpc(index).Target = 0
    MapNpc(index).HP = 0
    MapNpc(index).MP = 0
    MapNpc(index).SP = 0
    MapNpc(index).Map = 0
    MapNpc(index).X = 0
    MapNpc(index).Y = 0
    MapNpc(index).Dir = 0

    ' Client use only
    MapNpc(index).XOffset = 0
    MapNpc(index).YOffset = 0
    MapNpc(index).Moving = 0
    MapNpc(index).Attacking = 0
    MapNpc(index).AttackTimer = 0
End Sub

Sub ClearMapNpcs()
    Dim I As Long

    For I = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(I)
    Next I
End Sub

Function GetPlayerName(ByVal index As Long) As String
    If index < 1 Or index > MAX_PLAYERS Then
        Exit Function
    End If
    GetPlayerName = Trim$(Player(index).name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal name As String)
    Player(index).name = name
End Sub

Function GetPlayerGuild(ByVal index As Long) As String
    GetPlayerGuild = Trim$(Player(index).Guild)
End Function

Sub SetPlayerGuild(ByVal index As Long, ByVal Guild As String)
    Player(index).Guild = Guild
End Sub

Function GetPlayerGuildAccess(ByVal index As Long) As Long
    GetPlayerGuildAccess = Player(index).Guildaccess
End Function

Sub SetPlayerGuildAccess(ByVal index As Long, ByVal Guildaccess As Long)
    Player(index).Guildaccess = Guildaccess
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    GetPlayerClass = Player(index).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    Player(index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long
    GetPlayerSprite = Player(index).Sprite
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    Player(index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long
    GetPlayerLevel = Player(index).Level
End Function

Sub SetPlayerLevel(ByVal index As Long, ByVal Level As Long)
    Player(index).Level = Level
End Sub

Function GetPlayerExp(ByVal index As Long) As Long
    GetPlayerExp = Player(index).Exp
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal Exp As Long)
    Player(index).Exp = Exp
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long
    GetPlayerAccess = Player(index).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    Player(index).Access = Access
End Sub

Function GetPlayerPK(ByVal index As Long) As Long
    GetPlayerPK = Player(index).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    Player(index).PK = PK
End Sub

Function GetPlayerHP(ByVal index As Long) As Long
    GetPlayerHP = Player(index).HP
End Function

Sub SetPlayerHP(ByVal index As Long, ByVal HP As Long)
    Player(index).HP = HP

    If GetPlayerHP(index) > GetPlayerMaxHP(index) Then
        Player(index).HP = GetPlayerMaxHP(index)
    End If
End Sub

Function GetPlayerMP(ByVal index As Long) As Long
    GetPlayerMP = Player(index).MP
End Function

Sub SetPlayerMP(ByVal index As Long, ByVal MP As Long)
    Player(index).MP = MP

    If GetPlayerMP(index) > GetPlayerMaxMP(index) Then
        Player(index).MP = GetPlayerMaxMP(index)
    End If
End Sub

Function GetPlayerSP(ByVal index As Long) As Long
    GetPlayerSP = Player(index).SP
End Function

Sub SetPlayerSP(ByVal index As Long, ByVal SP As Long)
    Player(index).SP = SP

    If GetPlayerSP(index) > GetPlayerMaxSP(index) Then
        Player(index).SP = GetPlayerMaxSP(index)
    End If
End Sub

Function GetPlayerMaxHP(ByVal index As Long) As Long
    GetPlayerMaxHP = Player(index).MaxHP
End Function

Function GetPlayerMaxMP(ByVal index As Long) As Long
    GetPlayerMaxMP = Player(index).MaxMP
End Function

Function GetPlayerMaxSP(ByVal index As Long) As Long
    GetPlayerMaxSP = Player(index).MaxSP
End Function

Function GetPlayerSTR(ByVal index As Long) As Long
    GetPlayerSTR = Player(index).STR
End Function

Sub SetPlayerSTR(ByVal index As Long, ByVal STR As Long)
    Player(index).STR = STR
End Sub

Function GetPlayerDEF(ByVal index As Long) As Long
    GetPlayerDEF = Player(index).DEF
End Function

Sub SetPlayerDEF(ByVal index As Long, ByVal DEF As Long)
    Player(index).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal index As Long) As Long
    GetPlayerSPEED = Player(index).SPEED
End Function

Sub SetPlayerSPEED(ByVal index As Long, ByVal SPEED As Long)
    Player(index).SPEED = SPEED
End Sub

Function GetPlayerMAGI(ByVal index As Long) As Long
    GetPlayerMAGI = Player(index).MAGI
End Function

Sub SetPlayerMAGI(ByVal index As Long, ByVal MAGI As Long)
    Player(index).MAGI = MAGI
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long
    GetPlayerPOINTS = Player(index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    Player(index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal index As Long) As Long
    If index <= 0 Then
        Exit Function
    End If
    GetPlayerMap = Player(index).Map
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal MapNum As Long)
    Player(index).Map = MapNum
End Sub

Function GetPlayerX(ByVal index As Long) As Long
    GetPlayerX = Player(index).X
End Function

Sub SetPlayerX(ByVal index As Long, ByVal X As Long)
Player(index).X = X
End Sub

Function GetPlayerY(ByVal index As Long) As Long
    GetPlayerY = Player(index).Y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal Y As Long)
    Player(index).Y = Y
End Sub
Sub SetPlayerLoc(ByVal index As Long, ByVal X As Long, ByVal Y As Long)
    Player(index).X = X
    Player(index).Y = Y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long
    GetPlayerDir = Player(index).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    Player(index).Dir = Dir
End Sub

Function GetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(index).Inv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(index).Inv(InvSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(index).Inv(InvSlot).value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(index).Inv(InvSlot).value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(index).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(index).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerArmorSlot(ByVal index As Long) As Long
    GetPlayerArmorSlot = Player(index).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal index As Long, InvNum As Long)
    Player(index).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal index As Long) As Long
    GetPlayerWeaponSlot = Player(index).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal index As Long, InvNum As Long)
    Player(index).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal index As Long) As Long
    GetPlayerHelmetSlot = Player(index).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal index As Long, InvNum As Long)
    Player(index).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal index As Long) As Long
    GetPlayerShieldSlot = Player(index).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal index As Long, InvNum As Long)
    Player(index).ShieldSlot = InvNum
End Sub
Function GetPlayerLegsSlot(ByVal index As Long) As Long
    GetPlayerLegsSlot = Player(index).LegsSlot
End Function

Sub SetPlayerLegsSlot(ByVal index As Long, InvNum As Long)
    Player(index).LegsSlot = InvNum
End Sub
Function GetPlayerRingSlot(ByVal index As Long) As Long
    GetPlayerRingSlot = Player(index).RingSlot
End Function

Sub SetPlayerRingSlot(ByVal index As Long, InvNum As Long)
    Player(index).RingSlot = InvNum
End Sub
Function GetPlayerNecklaceSlot(ByVal index As Long) As Long
    GetPlayerNecklaceSlot = Player(index).NecklaceSlot
End Function

Sub SetPlayerNecklaceSlot(ByVal index As Long, InvNum As Long)
    Player(index).NecklaceSlot = InvNum
End Sub

Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long
    If BankSlot > MAX_BANK Then
        Exit Function
    End If
    GetPlayerBankItemNum = Player(index).Bank(BankSlot).Num
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
    Player(index).Bank(BankSlot).Num = ItemNum
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Player(index).Bank(BankSlot).value
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Player(index).Bank(BankSlot).value = ItemValue
End Sub

Function GetPlayerBankItemDur(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemDur = Player(index).Bank(BankSlot).Dur
End Function

Sub SetPlayerBankItemDur(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemDur As Long)
    Player(index).Bank(BankSlot).Dur = ItemDur
End Sub

Function GetPlayerHead(ByVal index As Long) As Long
    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerHead = Player(index).head
    End If
End Function

Sub SetPlayerHead(ByVal index As Long, ByVal head As Long)
    If index > 0 And index < MAX_PLAYERS Then
        Player(index).head = head
    End If
End Sub

Function GetPlayerBody(ByVal index As Long) As Long
    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerBody = Player(index).body
    End If
End Function

Sub SetPlayerBody(ByVal index As Long, ByVal body As Long)
    If index > 0 And index < MAX_PLAYERS Then
        Player(index).body = body
    End If
End Sub

Function GetPlayerLeg(ByVal index As Long) As Long
    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerLeg = Player(index).leg
    End If
End Function

Sub SetPlayerLeg(ByVal index As Long, ByVal leg As Long)
    If index > 0 And index < MAX_PLAYERS Then
        Player(index).leg = leg
    End If
End Sub

Function GetPlayerSkillLvl(ByVal index As Long, ByVal skill As Long) As Long
    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerSkillLvl = Player(index).SkilLvl(skill)
    End If
End Function

Sub SetPlayerSkillLvl(ByVal index As Long, ByVal skill As Long, ByVal lvl As Long)
    If index > 0 And index < MAX_PLAYERS Then
        Player(index).SkilLvl(skill) = lvl
    End If
End Sub

Function GetPlayerSkillExp(ByVal index As Long, ByVal skill As Long) As Long
    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerSkillExp = Player(index).SkilExp(skill)
    End If
End Function

Sub SetPlayerSkillExp(ByVal index As Long, ByVal skill As Long, ByVal lvl As Long)
    If index > 0 And index < MAX_PLAYERS Then
        Player(index).SkilExp(skill) = lvl
    End If
End Sub

Function GetPlayerPaperdoll(ByVal index As Long) As Byte
    If index < MAX_PLAYERS And index > 0 Then
        If IsPlaying(index) Then
            GetPlayerPaperdoll = Player(index).paperdoll
        End If
    End If
End Function

Sub SetPlayerPaperdoll(ByVal index As Long, ByVal paperdoll As Byte)
    If index < MAX_PLAYERS And index > 0 Then
        If IsPlaying(index) Then
            Player(index).paperdoll = paperdoll
        End If
    End If
End Sub
