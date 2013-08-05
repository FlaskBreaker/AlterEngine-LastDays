Attribute VB_Name = "modClientTCP"
Option Explicit

Sub TcpInit()
    SEP_CHAR = Chr$(0)
    END_CHAR = Chr$(237)

    PlayerBuffer = vbNullString

    frmMirage.Socket.RemoteHost = ReadINI("IPCONFIG", "IP", App.Path & "\Config.ini")
    frmMirage.Socket.RemotePort = ReadINI("IPCONFIG", "PORT", App.Path & "\Config.ini")
End Sub

Sub LoadSCT(ByVal index As Long, ByVal SpellMemorized As Long)
 Dim packet As String
 packet = "loadsct" & SEP_CHAR & SpellMemorized & END_CHAR
 
 Call SendData(packet)
End Sub

Sub TcpDestroy()
    frmMirage.Socket.close
End Sub

Sub IncomingData(ByVal DataLength As Long)
    Dim buffer As String
    Dim packet As String
    Dim Start As Long

    frmMirage.Socket.GetData buffer, vbString, DataLength

    PlayerBuffer = PlayerBuffer & buffer

    Start = InStr(PlayerBuffer, END_CHAR)
    Do While Start > 0
        packet = Mid$(PlayerBuffer, 1, Start - 1)
        PlayerBuffer = Mid$(PlayerBuffer, Start + 1, Len(PlayerBuffer))
        Start = InStr(PlayerBuffer, END_CHAR)
        If LenB(packet) > 0 Then
            Call HandleData(packet)
        End If
    Loop
End Sub

Function ConnectToServer() As Boolean
    If IsConnected Then
        ConnectToServer = True
        Exit Function
    End If

    Call TcpDestroy
    frmMirage.Socket.Connect

    If IsConnected Then
        ConnectToServer = True
    Else
        ConnectToServer = False
    End If
End Function

Function IsConnected() As Boolean
    If frmMirage.Socket.State = sckConnected Then
        IsConnected = True
    End If
End Function

Function IsPlaying(ByVal index As Long) As Boolean
    If GetPlayerName(index) <> vbNullString Then
        IsPlaying = True
    End If
End Function

Function IsAlphaNumeric(TestString As String) As Boolean
    Dim LoopID As Integer
    Dim sChar As String

    If LenB(TestString) > 0 Then
        For LoopID = 1 To Len(TestString)
            sChar = Mid(TestString, LoopID, 1)
            If Not sChar Like "[0-9A-Za-z]" Then
                Exit Function
            End If
        Next

        IsAlphaNumeric = True
    End If
End Function

Sub SendData(ByVal data As String)
    If IsConnected Then
        frmMirage.Socket.SendData data
        DoEvents
    End If
End Sub

Sub SendNewAccount(ByVal name As String, ByVal Password As String, ByVal Email As String)
    Call SendData("newaccount" & SEP_CHAR & Trim$(name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & Trim$(Email) & END_CHAR)
End Sub

Sub SendDelAccount(ByVal name As String, ByVal Password As String)
    Call SendData("delaccount" & SEP_CHAR & Trim$(name) & SEP_CHAR & Trim$(Password) & END_CHAR)
End Sub

Sub SendLogin(ByVal name As String, ByVal Password As String)
    Call SendData("acclogin" & SEP_CHAR & Trim$(name) & SEP_CHAR & Trim$(Password) & SEP_CHAR & App.Major & SEP_CHAR & App.Minor & SEP_CHAR & App.Revision & SEP_CHAR & SEC_CODE & SEP_CHAR & END_CHAR)
End Sub

Sub SendAddChar(ByVal name As String, ByVal Sex As Long, ByVal ClassNum As Long, ByVal slot As Long, ByVal HeadC As Long, ByVal BodyC As Long, ByVal LegC As Long)
    Call SendData("addchar" & SEP_CHAR & Trim$(name) & SEP_CHAR & Sex & SEP_CHAR & ClassNum & SEP_CHAR & slot & SEP_CHAR & HeadC & SEP_CHAR & BodyC & SEP_CHAR & LegC & END_CHAR)
End Sub

Sub SendDelChar(ByVal slot As Long)
    Call SendData("delchar" & SEP_CHAR & slot & END_CHAR)
End Sub

Sub SendGetClasses()
    Call SendData("getclasses" & END_CHAR)
End Sub

Sub SendUseChar(ByVal CharSlot As Long)
    Call SendData("usechar" & SEP_CHAR & CharSlot & END_CHAR)
End Sub

Sub SayMsg(ByVal Text As String)
    Call SendData("saymsg" & SEP_CHAR & Text & END_CHAR)
End Sub

Sub GlobalMsg(ByVal Text As String)
    Call SendData("globalmsg" & SEP_CHAR & Text & END_CHAR)
End Sub

Sub BroadcastMsg(ByVal Text As String)
    Call SendData("broadcastmsg" & SEP_CHAR & Text & END_CHAR)
End Sub

Sub EmoteMsg(ByVal Text As String)
    Call SendData("emotemsg" & SEP_CHAR & Text & END_CHAR)
End Sub

Sub MapMsg(ByVal Text As String)
    Call SendData("mapmsg" & SEP_CHAR & Text & END_CHAR)
End Sub

Sub PlayerMsg(ByVal Text As String, ByVal MsgTo As String)
    Call SendData("playermsg" & SEP_CHAR & MsgTo & SEP_CHAR & Text & END_CHAR)
End Sub

Sub AdminMsg(ByVal Text As String)
    Call SendData("adminmsg" & SEP_CHAR & Text & END_CHAR)
End Sub

Sub SendPlayerMove()
Call SendData("playermove" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & Player(MyIndex).Moving & SEP_CHAR & GetPlayerX(MyIndex) & SEP_CHAR & GetPlayerY(MyIndex) & END_CHAR)
End Sub

Sub SendPlayerDir()
    Call SendData("playerdir" & SEP_CHAR & GetPlayerDir(MyIndex) & END_CHAR)
End Sub

Sub SendPlayerRequestNewMap(ByVal Cancel As Long)
    Call SendData("requestnewmap" & SEP_CHAR & GetPlayerDir(MyIndex) & SEP_CHAR & Cancel & END_CHAR)
End Sub

Sub SendMap()
    Dim packet As String
    Dim X As Byte
    Dim Y As Byte

    packet = "mapdata" & SEP_CHAR & GetPlayerMap(MyIndex) & SEP_CHAR & Trim$(Map(GetPlayerMap(MyIndex)).name) & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Revision & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Moral & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Up & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Down & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Left & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Right & SEP_CHAR & Map(GetPlayerMap(MyIndex)).music & SEP_CHAR & Map(GetPlayerMap(MyIndex)).BootMap & SEP_CHAR & Map(GetPlayerMap(MyIndex)).BootX & SEP_CHAR & Map(GetPlayerMap(MyIndex)).BootY & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Indoors & SEP_CHAR & Map(GetPlayerMap(MyIndex)).Weather & SEP_CHAR

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            With Map(GetPlayerMap(MyIndex)).Tile(X, Y)
                packet = packet & (.Ground & SEP_CHAR & .mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR & .Light & SEP_CHAR)
                packet = packet & (.GroundSet & SEP_CHAR & .MaskSet & SEP_CHAR & .AnimSet & SEP_CHAR & .Mask2Set & SEP_CHAR & .M2AnimSet & SEP_CHAR & .FringeSet & SEP_CHAR & .FAnimSet & SEP_CHAR & .Fringe2Set & SEP_CHAR & .F2AnimSet & SEP_CHAR)
            End With
        Next X
    Next Y

    With Map(GetPlayerMap(MyIndex))
        For X = 1 To MAX_MAP_NPCS
            packet = packet & (.Npc(X) & SEP_CHAR & .SpawnX(X) & SEP_CHAR & .SpawnY(X) & SEP_CHAR)
        Next X
    End With

    packet = packet & Map(GetPlayerMap(MyIndex)).owner & END_CHAR

    Call SendData(packet)
End Sub

Sub WarpMeTo(ByVal name As String)
    Call SendData("warpmeto" & SEP_CHAR & name & END_CHAR)
End Sub

Sub WarpToMe(ByVal name As String)
    Call SendData("warptome" & SEP_CHAR & name & END_CHAR)
End Sub

Sub WarpTo(ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long)
    Call SendData("warpto" & SEP_CHAR & MapNum & SEP_CHAR & X & SEP_CHAR & Y & END_CHAR)
End Sub

Sub LocalWarp(ByVal X As Long, ByVal Y As Long)
    Call SendData("localwarp" & SEP_CHAR & X & SEP_CHAR & Y & END_CHAR)
End Sub

Sub SendSetAccess(ByVal name As String, ByVal Access As Byte)
    Call SendData("setaccess" & SEP_CHAR & name & SEP_CHAR & Access & END_CHAR)
End Sub

Sub SendKick(ByVal name As String)
    Call SendData("kickplayer" & SEP_CHAR & name & END_CHAR)
End Sub

Sub SendBan(ByVal name As String)
    Call SendData("banplayer" & SEP_CHAR & name & END_CHAR)
End Sub

Sub SendBanList()
    Call SendData("banlist" & END_CHAR)
End Sub

Sub SendRequestEditItem()
    Call SendData("requestedititem" & END_CHAR)
End Sub

Sub SendSaveItem(ByVal ItemNum As Long)
    Call SendData("saveitem" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR & Item(ItemNum).AddHP & SEP_CHAR & Item(ItemNum).AddMP & SEP_CHAR & Item(ItemNum).AddSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound & SEP_CHAR & Item(ItemNum).PetSprite & END_CHAR)
End Sub

Sub SendRequestEditEmoticon()
    Call SendData("requesteditemoticon" & END_CHAR)
End Sub

Sub SendRequestEditElement()
    Call SendData("requesteditelement" & END_CHAR)
End Sub

Sub SendRequestEditSkill()
    Call SendData("requesteditskill" & END_CHAR)
End Sub

Sub SendSaveEmoticon(ByVal EmoNum As Long)
    Call SendData("saveemoticon" & SEP_CHAR & EmoNum & SEP_CHAR & Trim$(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & END_CHAR)
End Sub

Sub SendSaveElement(ByVal ElementNum As Long)
    Call SendData("saveelement" & SEP_CHAR & ElementNum & SEP_CHAR & Trim$(Element(ElementNum).name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & END_CHAR)
End Sub

Sub SendRequestEditArrow()
    Call SendData("requesteditarrow" & END_CHAR)
End Sub

Sub SendSaveArrow(ByVal ArrowNum As Long)
    Call SendData("savearrow" & SEP_CHAR & ArrowNum & SEP_CHAR & Trim$(Arrows(ArrowNum).name) & SEP_CHAR & Arrows(ArrowNum).Pic & SEP_CHAR & Arrows(ArrowNum).Range & SEP_CHAR & Arrows(ArrowNum).Amount & END_CHAR)
End Sub

Sub SendRequestEditNPC()
    Call SendData("requesteditnpc" & END_CHAR)
End Sub

Sub SendSaveNPC(ByVal NpcNum As Long)
    Dim packet As String
    Dim I As Long

    packet = "savenpc" & SEP_CHAR & NpcNum & SEP_CHAR & Trim$(Npc(NpcNum).name) & SEP_CHAR & Trim$(Npc(NpcNum).AttackSay) & SEP_CHAR & Npc(NpcNum).Sprite & SEP_CHAR & Npc(NpcNum).SpawnSecs & SEP_CHAR & Npc(NpcNum).Behavior & SEP_CHAR & Npc(NpcNum).Range & SEP_CHAR & Npc(NpcNum).STR & SEP_CHAR & Npc(NpcNum).DEF & SEP_CHAR & Npc(NpcNum).SPEED & SEP_CHAR & Npc(NpcNum).MAGI & SEP_CHAR & Npc(NpcNum).Big & SEP_CHAR & Npc(NpcNum).MaxHP & SEP_CHAR & Npc(NpcNum).Exp & SEP_CHAR & Npc(NpcNum).SpawnTime & SEP_CHAR & Npc(NpcNum).Element & SEP_CHAR & Npc(NpcNum).SpriteSize & SEP_CHAR & Npc(NpcNum).Quest

    For I = 1 To MAX_NPC_DROPS
        packet = packet & (SEP_CHAR & Npc(NpcNum).ItemNPC(I).chance & SEP_CHAR & Npc(NpcNum).ItemNPC(I).ItemNum & SEP_CHAR & Npc(NpcNum).ItemNPC(I).ItemValue)
    Next I

    packet = packet & SEP_CHAR & Npc(NpcNum).standstill & END_CHAR

    Call SendData(packet)
End Sub

Sub SendMapRespawn()
    Call SendData("maprespawn" & END_CHAR)
End Sub

Sub SendUseItem(ByVal InvNum As Long)
    Call SendData("useitem" & SEP_CHAR & InvNum & END_CHAR)
End Sub

Sub SendScript(ByVal Num As Long)
    Call SendData("scriptedaction" & SEP_CHAR & Num & END_CHAR)
End Sub

Sub SendDropItem(ByVal InvNum As Long, ByVal Amount As Long)
    Call SendData("mapdropitem" & SEP_CHAR & InvNum & SEP_CHAR & Amount & END_CHAR)
End Sub

Sub SendWhosOnline()
    Call SendData("whosonline" & END_CHAR)
End Sub

Sub SendOnlineList()
    Call SendData("onlinelist" & END_CHAR)
End Sub

Sub SendMOTDChange(ByVal MOTD As String)
    Call SendData("setmotd" & SEP_CHAR & MOTD & END_CHAR)
End Sub

Sub SendRequestEditShop()
    Call SendData("requesteditshop" & END_CHAR)
End Sub

Sub SendSaveShop(ByVal shopNum As Long)
    Dim packet As String
    Dim I As Integer

    packet = "saveshop" & SEP_CHAR & shopNum & SEP_CHAR & Trim$(Shop(shopNum).name) & SEP_CHAR & Shop(shopNum).FixesItems & SEP_CHAR & Shop(shopNum).BuysItems & SEP_CHAR & Shop(shopNum).ShowInfo & SEP_CHAR & Shop(shopNum).currencyItem & SEP_CHAR

    For I = 1 To MAX_SHOP_ITEMS
        packet = packet & (Shop(shopNum).ShopItem(I).ItemNum & SEP_CHAR & Shop(shopNum).ShopItem(I).Amount & SEP_CHAR & Shop(shopNum).ShopItem(I).Price & SEP_CHAR)
    Next I

    packet = packet & END_CHAR

    Call SendData(packet)
End Sub

Sub SendRequestEditSpell()
    Call SendData("requesteditspell" & END_CHAR)
End Sub

Sub SendReloadScripts()
    Call SendData("reloadscripts" & END_CHAR)
End Sub

Sub SendSaveSpell(ByVal Spellnum As Long)
    Call SendData("savespell" & SEP_CHAR & Spellnum & SEP_CHAR & Trim$(Spell(Spellnum).name) & SEP_CHAR & Spell(Spellnum).ClassReq & SEP_CHAR & Spell(Spellnum).LevelReq & SEP_CHAR & Spell(Spellnum).Type & SEP_CHAR & Spell(Spellnum).Data1 & SEP_CHAR & Spell(Spellnum).Data2 & SEP_CHAR & Spell(Spellnum).Data3 & SEP_CHAR & Spell(Spellnum).MPCost & SEP_CHAR & Trim$(Spell(Spellnum).Sound) & SEP_CHAR & Spell(Spellnum).Range & SEP_CHAR & Spell(Spellnum).SpellAnim & SEP_CHAR & Spell(Spellnum).SpellTime & SEP_CHAR & Spell(Spellnum).SpellDone & SEP_CHAR & Spell(Spellnum).AE & SEP_CHAR & Spell(Spellnum).Big & SEP_CHAR & Spell(Spellnum).Element & SEP_CHAR & Spell(Spellnum).TimeToCast & SEP_CHAR & Spell(Spellnum).CastTimer & END_CHAR)
End Sub

Sub SendRequestEditMap()
    Call SendData("requesteditmap" & END_CHAR)
End Sub

Sub SendRequestEditHouse()
    Call SendData("requestedithouse" & END_CHAR)
End Sub

Sub SendTradeRequest(ByVal name As String)
    Call SendData("pptrade" & SEP_CHAR & name & END_CHAR)
End Sub

Sub SendAcceptTrade()
    Call SendData("atrade" & END_CHAR)
End Sub

Sub SendDeclineTrade()
    Call SendData("dtrade" & END_CHAR)
End Sub

Sub PartyMsg(ByVal Text As String)
Call SendData("d" & SEP_CHAR & Text & SEP_CHAR & END_CHAR)
End Sub

Sub GuildMsg(ByVal Text As String)
Call SendData("guildmsg" & SEP_CHAR & Text & SEP_CHAR & END_CHAR)
End Sub

Sub SendPartyRequest()
Call SendData("i2" & SEP_CHAR & END_CHAR)
End Sub
Sub InvitePlayer()
Call SendData("57" & SEP_CHAR & END_CHAR)
End Sub
Sub RemoveMember()
Call SendData("c7" & SEP_CHAR & END_CHAR)
End Sub
Sub SendPartyInvite(ByVal name As String)
Call SendData("58" & SEP_CHAR & name & SEP_CHAR & END_CHAR)
End Sub
Sub SendRoll(ByVal ItemSlot As Byte)
Call SendData("i" & SEP_CHAR & ItemSlot & SEP_CHAR & END_CHAR)
End Sub
Sub SendNoRoll(ByVal ItemSlot As Byte)
Call SendData("82" & SEP_CHAR & ItemSlot & SEP_CHAR & END_CHAR)
End Sub
Sub SendFunRoll()
Call SendData("d1" & SEP_CHAR & END_CHAR)
End Sub
Sub SendReturn()
Call SendData("d0" & SEP_CHAR & END_CHAR)
End Sub
Sub SendJoinParty()
Call SendData("62" & SEP_CHAR & END_CHAR)
End Sub
Sub SendLeaveParty()
Player(MyIndex).InParty = False
Call SendData("64" & SEP_CHAR & END_CHAR)
End Sub
Sub SendTarget(ByVal PartySlot As Byte)
Call SendData("f6" & SEP_CHAR & PartySlot & SEP_CHAR & END_CHAR)
End Sub
Sub SendBanDestroy()
    Call SendData("bandestroy" & END_CHAR)
End Sub

'Sub SendSetPlayerSprite(ByVal Name As String, ByVal SpriteNum As Byte)
'   Call SendData("setplayersprite" & SEP_CHAR & Name & SEP_CHAR & SpriteNum & END_CHAR)
'End Sub
' Fix para el Runtime Error 6

Public Sub SendSetPlayerSprite(ByVal name As String, ByVal SpriteNum As Long)
    Call SendData("setplayersprite" & SEP_CHAR & name & SEP_CHAR & CStr(SpriteNum) & END_CHAR)
End Sub

Sub SendHotScript(ByVal value As Byte)
    Call SendData("hotscript" & SEP_CHAR & value & END_CHAR)
End Sub

Sub SendScriptTile(ByVal Text As String)
    Call SendData("scripttile" & SEP_CHAR & Text & END_CHAR)
End Sub

Sub SendPlayerMoveMouse()
    Call SendData("playermovemouse" & SEP_CHAR & GetPlayerDir(MyIndex) & END_CHAR)
End Sub

Sub SendChangeDir()
    Call SendData("warp" & SEP_CHAR & GetPlayerDir(MyIndex) & END_CHAR)
End Sub

Sub SendUseStatPoint(ByVal value As Byte)
    Call SendData("usestatpoint" & SEP_CHAR & value & END_CHAR)
End Sub

Sub SendGuildLeave()
    Call SendData("GUILDLEAVE" & END_CHAR)
End Sub

Sub SendGuildMember(ByVal name As String)
    Call SendData("GUILDMEMBER" & SEP_CHAR & name & END_CHAR)
End Sub

Sub SendRequestSpells()
    Call SendData("spells" & END_CHAR)
End Sub

Sub SendForgetSpell(ByVal SpellID As Long)
    If Player(MyIndex).Spell(SpellID) > 0 Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            If MsgBox("Estas seguro de querer borrar el hechizo?", vbYesNo, "Olvidar Hechizo") = vbYes Then
                Call SendData("forgetspell" & SEP_CHAR & SpellID & END_CHAR)
                frmMirage.picPlayerSpells.Visible = False
                frmMirage.picPlayerSpells.Visible = True
            End If
        End If
    Else
        Call AddText("No hay ningun hechizo aqui.", BRIGHTRED)
    End If
End Sub

Sub SendRequestMyStats()
    Call SendData("getstats" & SEP_CHAR & GetPlayerName(MyIndex) & END_CHAR)
End Sub

Sub SendSetTrainee(ByVal name As String)
    Call SendData("GUILDTRAINEE" & SEP_CHAR & name & END_CHAR)
End Sub

Sub SendGuildDisown(ByVal name As String)
    Call SendData("GUILDDISOWN" & SEP_CHAR & name & END_CHAR)
End Sub

Sub SendChangeGuildAccess(ByVal name As String, ByVal AccessLvl As Long)
    Call SendData("GUILDCHANGEACCESS" & SEP_CHAR & name & SEP_CHAR & AccessLvl & END_CHAR)
End Sub

Sub SendPlayerChat(ByVal name As String)
    Call SendData("playerchat" & SEP_CHAR & name & END_CHAR)
End Sub

' Perfiles de jugadores
Sub SendProfile(ByVal name As String)
Call SendData("profile" & SEP_CHAR & name & END_CHAR)
End Sub

Sub SendGiveItem(ByVal name As String, ByVal Item As Integer)
Dim packet As String

    packet = "GIVEITEM" & SEP_CHAR & name & SEP_CHAR & Item & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendTakeItem(ByVal name As String, ByVal Item As Integer)
Dim packet As String

    packet = "TAKEITEM" & SEP_CHAR & name & SEP_CHAR & Item & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub

Sub SendRequestEditMain()
Dim packet As String

    packet = "REQUESTEDITMAIN" & SEP_CHAR & END_CHAR
    Call SendData(packet)
End Sub
