Attribute VB_Name = "modHandleData"
Option Explicit

Sub HandleData(ByVal index As Long, ByVal Data As String)
    Dim Parse() As String
    Dim N As Long
    Dim PointType As Integer
    Dim FileData As String

    On Error Resume Next

    Parse = Split(Data, SEP_CHAR)

    Select Case LCase$(Parse(0))
        Case "getclasses"
            Call Packet_GetClasses(index)
            Exit Sub
        Case "loadsct"
            Call Packet_loadsct(index, Parse(1))
            Exit Sub
            
        Case "usepetstatpoint"
            PointType = Val(Parse(1))
            'frozen
            If GetPetPOINTS(index) > 0 Then
                If GetPetLevel(index) >= 1 Then
                    Call UsingPetStatPoints(index, PointType)
                End If
            Else
                Call BattleMsg(index, "Tu mascota no tiene puntos para entrenarla!", BRIGHTRED, 0)
            End If
            
            Call SendPlayerData(index)
            Exit Sub
            
        Case "choosepet"
        Call ChoosePet(index, Val(Parse(1)), Parse(2))
        Exit Sub
        
        Case "killpet"
            Call KillPet(index)
            Exit Sub
            
        Case "petmoveselect"
            Call DoPetMoveSelect(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub
    
        Case "newaccount"
            Call Packet_NewAccount(index, Parse(1), Parse(2), Parse(3))
            Exit Sub
            
        Case "delaccount"
            Call Packet_DeleteAccount(index, Parse(1), Parse(2))
            Exit Sub
    
        Case "acclogin"
            Call Packet_AccountLogin(index, Parse(1), Parse(2), Val(Parse(3)), Val(Parse(4)), Val(Parse(5)), Parse(6))
            Exit Sub
    
        Case "givemethemax"
            Call Packet_GiveMeTheMax(index)
            Exit Sub
    
        Case "addchar"
            Call Packet_AddCharacter(index, Parse(1), Val(Parse(2)), Val(Parse(3)), Val(Parse(4)), Val(Parse(5)), Val(Parse(6)), Val(Parse(7)))
            Exit Sub
        
        Case "allchars"
            If Parse(1) > 0 And Parse(1) < MAX_PLAYERS Then
                Call SendChars(Parse(1))
            End If
            Exit Sub
    
        Case "delchar"
            Call Packet_DeleteCharacter(index, Val(Parse(1)))
            Exit Sub
    
        Case "usechar"
            Call Packet_UseCharacter(index, Val(Parse(1)))
            Exit Sub

        Case "guildchangeaccess"
            Call Packet_GuildChangeAccess(index, Parse(1), Val(Parse(2)))
            Exit Sub

        Case "guilddisown"
            Call Packet_GuildDisown(index, Parse(1))
            Exit Sub

        Case "guildleave"
            Call Packet_GuildLeave(index)
            Exit Sub

        Case "guildmake"
            Call Packet_GuildMake(index, Parse(1), Parse(2))
            Exit Sub

        Case "guildmember"
            Call Packet_GuildMember(index, Parse(1))
            Exit Sub

        Case "guildtrainee"
            Call Packet_GuildTrainee(index, Parse(1))
            Exit Sub

        Case "saymsg"
            Call Packet_SayMessage(index, Parse(1))
            Exit Sub

        Case "emotemsg"
            Call Packet_EmoteMessage(index, Parse(1))
            Exit Sub

        Case "broadcastmsg"
            Call Packet_BroadcastMessage(index, Parse(1))
            Exit Sub

        Case "globalmsg"
            Call Packet_GlobalMessage(index, Parse(1))
            Exit Sub

        Case "adminmsg"
            Call Packet_AdminMessage(index, Parse(1))
            Exit Sub

        Case "playermsg"
            Call Packet_PlayerMessage(index, Parse(1), Parse(2))
            Exit Sub

        Case "playermove"
            Call Packet_PlayerMove(index, Val(Parse(1)), Val(Parse(2)), Val(Parse(3)), Val(Parse(4)))
            Exit Sub

        Case "playerdir"
            Call Packet_PlayerDirection(index, Val(Parse(1)))
            Exit Sub

        Case "useitem"
            Call Packet_UseItem(index, Val(Parse(1)))
            Exit Sub

        Case "playermovemouse"
            Call Packet_PlayerMoveMouse(index, Val(Parse(1)))
            Exit Sub

        Case "warp"
            Call Packet_Warp(index, Val(Parse(1)))
            Exit Sub

        Case "endshot"
            Call Packet_EndShot(index, Val(Parse(1)))
            Exit Sub

        Case "attack"
            Call Packet_Attack(index)
            Exit Sub

        Case "usestatpoint"
            Call Packet_UseStatPoint(index, Val(Parse(1)))
            Exit Sub

        Case "setplayersprite"
            Call Packet_SetPlayerSprite(index, Parse(1), Val(Parse(2)))
            Exit Sub

        Case "getstats"
            Call Packet_GetStats(index, Parse(1))
            Exit Sub

        Case "requestnewmap"
            Call Packet_RequestNewMap(index, Val(Parse(1)))
            Exit Sub

        Case "warpmeto"
            Call Packet_WarpMeTo(index, Parse(1))
            Exit Sub

        Case "warptome"
            Call Packet_WarpToMe(index, Parse(1))
            Exit Sub

        Case "mapdata"
            Call Packet_MapData(index, Parse)
            Exit Sub

        Case "needmap"
            Call Packet_NeedMap(index, Parse(1))
            Exit Sub
            
        Case "requestspass"
            Call Packet_SPass(index)
            Exit Sub

        Case "mapgetitem"
            Call Packet_MapGetItem(index)
            Exit Sub
            
        Case "mapdropitem"
            Call Packet_MapDropItem(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub

        Case "maprespawn"
            Call Packet_MapRespawn(index)
            Exit Sub

        Case "kickplayer"
            Call Packet_KickPlayer(index, Parse(1))
            Exit Sub

        Case "banlist"
            Call Packet_BanList(index)
            Exit Sub

        Case "bandestroy"
            Call Packet_BanListDestroy(index)
            Exit Sub

        Case "banplayer"
            Call Packet_BanPlayer(index, Parse(1))
            Exit Sub

        Case "requesteditmap"
            Call Packet_RequestEditMap(index)
            Exit Sub

        Case "requestedititem"
            Call Packet_RequestEditItem(index)
            Exit Sub

        Case "edititem"
            Call Packet_EditItem(index, Val(Parse(1)))
            Exit Sub

        Case "saveitem"
            Call Packet_SaveItem(index, Parse)
            Exit Sub

        Case "enabledaynight"
            Call Packet_EnableDayNight(index)
            Exit Sub

        Case "daynight"
            Call Packet_DayNight(index)
            Exit Sub

        Case "requesteditnpc"
            Call Packet_RequestEditNPC(index)
            Exit Sub

        Case "editnpc"
            Call Packet_EditNPC(index, Val(Parse(1)))
            Exit Sub

        Case "savenpc"
            Call Packet_SaveNPC(index, Parse)
            Exit Sub

        Case "requesteditshop"
            Call Packet_RequestEditShop(index)
            Exit Sub

        Case "editshop"
            Call Packet_EditShop(index, Val(Parse(1)))
            Exit Sub

        Case "saveshop"
            Call Packet_SaveShop(index, Parse)
            Exit Sub

        Case "requesteditspell"
            Call Packet_RequestEditSpell(index)
            Exit Sub

        Case "editspell"
            Call Packet_EditSpell(index, Val(Parse(1)))
            Exit Sub

        Case "savespell"
            Call Packet_SaveSpell(index, Parse)
            Exit Sub

        Case "forgetspell"
            Call Packet_ForgetSpell(index, Val(Parse(1)))
            Exit Sub

        Case "setaccess"
            Call Packet_SetAccess(index, Parse(1), Val(Parse(2)))
            Exit Sub

        Case "whosonline"
            Call Packet_WhoIsOnline(index)
            Exit Sub

        Case "onlinelist"
            Call Packet_OnlineList(index)
            Exit Sub

        Case "setmotd"
            Call Packet_SetMOTD(index, Parse(1))
            Exit Sub

        Case "buy"
            Call Packet_BuyItem(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub

        Case "sellitem"
            Call Packet_SellItem(index, Val(Parse(1)), Val(Parse(2)), Val(Parse(3)), Val(Parse(4)))
            Exit Sub

        Case "fixitem"
            Call Packet_FixItem(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub

        Case "search"
            Call Packet_Search(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub

        Case "playerchat"
            Call Packet_PlayerChat(index, Parse(1))
            Exit Sub

        Case "achat"
            Call Packet_AcceptChat(index)
            Exit Sub

        Case "dchat"
            Call Packet_DenyChat(index)
            Exit Sub

        Case "qchat"
            Call Packet_QuitChat(index)
            Exit Sub

        Case "sendchat"
            Call Packet_SendChat(index, Parse(1))
            Exit Sub

        Case "pptrade"
            Call Packet_PrepareTrade(index, Parse(1))
            Exit Sub

        Case "atrade"
            Call Packet_AcceptTrade(index)
            Exit Sub

        Case "qtrade"
            Call Packet_QuitTrade(index)
            Exit Sub

        Case "dtrade"
            Call Packet_DenyTrade(index)
            Exit Sub
            
        Case "requesteditmain"
            Call Packet_RequestEditMain(index)
            Exit Sub
            
        Case "newmain"
            Call Packet_NewMain(index, Parse(1))
            Exit Sub

        Case "updatetradeinv"
            Call Packet_UpdateTradeInventory(index, Val(Parse(1)), Val(Parse(2)), Parse(3))
            Exit Sub

        Case "swapitems"
            Call Packet_SwapItems(index)
            Exit Sub

        Case "party"
        N = FindPlayer(Parse(1))
        
        ' Prevent partying with self
        If N = index Then
            Exit Sub
        End If
                
        ' Check for a previous party and if so drop it
        If Player(index).InParty = YES Then
            Call PlayerMsg(index, "Ya te encuentras en este grupo!", Red)
            Exit Sub
        End If
        
        If N > 0 Then
            ' Check if its an admin
            If GetPlayerAccess(index) > ADMIN_MONITER Then
                Call PlayerMsg(index, "No puedes entrar a un grupo, eres un admin!", BRIGHTBLUE)
                Exit Sub
            End If
        
            If GetPlayerAccess(N) > ADMIN_MONITER Then
                Call PlayerMsg(index, "Los administradores no pueden entrar en grupos!", BRIGHTBLUE)
                Exit Sub
            End If
            
            ' Make sure they are in right level range
            If GetPlayerLevel(index) + 5 < GetPlayerLevel(N) Or GetPlayerLevel(index) - 5 > GetPlayerLevel(N) Then
                Call PlayerMsg(index, "Este jugador es mayor que tu por 5 niveles, el grupo fallo.", PINK)
                Exit Sub
            End If
            
            ' Check to see if player is already in a party
            If Player(N).InParty = NO Then
                Call PlayerMsg(index, "Una petición de grupo ha sido enviada a " & GetPlayerName(N) & ".", Green)
                Call PlayerMsg(N, GetPlayerName(index) & " quiere que te unas a su grupo.  Escribe /entrar para entrar o /salir para rechazar.", Green)
            
                Player(index).PartyStarter = YES
                Player(N).Char(Player(N).CharNum).PartyInvitedTo = Player(index).PartyID
                Player(N).PartyPlayer = index
            Else
                Call PlayerMsg(index, "Este jugador ya se encuentra en tu grupo!", Red)
            End If
        Else
            Call PlayerMsg(index, "El jugador no esta conectado.", WHITE)
        End If
        Exit Sub

    ' :::::::::::::::::::::::
    ' :: Join party packet ::
    ' :::::::::::::::::::::::
    Case "joinparty"
        N = Player(index).PartyPlayer
        
        If N > 0 Then
            ' Check to make sure they aren't the starter
            If Player(index).PartyStarter = NO Then
                ' Check to make sure that each of there party players match
                If Player(N).PartyPlayer = index Then
                    Call PlayerMsg(index, "Has entrado en el grupo de " & GetPlayerName(N) & "!", Green)
                    Call PlayerMsg(N, GetPlayerName(index) & " ha entrado en tu grupo!", Green)
                    
                    Player(index).InParty = YES
                    Player(N).InParty = YES
                Else
                    Call PlayerMsg(index, "El grupo fallo.", Red)
                End If
            Else
                Call PlayerMsg(index, "No has sido invitado por ningún grupo!", Red)
            End If
        Else
            Call PlayerMsg(index, "No ha sido invitado a unirte a ningún grupo!", Red)
        End If
        Exit Sub
   
    Case "leaveparty"
        N = Player(index).PartyPlayer
        
        If N > 0 Then
            If Player(index).InParty = YES Then
                Call PlayerMsg(index, "Has salido del grupo.", Green)
                Call PlayerMsg(N, GetPlayerName(index) & " ha salido del grupo.", Red)
                
                Player(index).PartyPlayer = 0
                Player(index).PartyStarter = NO
                Player(index).InParty = NO
                Player(N).PartyPlayer = 0
                Player(N).PartyStarter = NO
                Player(N).InParty = NO
            Else
                Call PlayerMsg(index, "Petición de grupo rechazada.", Green)
                Call PlayerMsg(N, GetPlayerName(index) & " ha rechazado tu petición.", Red)
                
                Player(index).PartyPlayer = 0
                Player(index).PartyStarter = NO
                Player(index).InParty = NO
                Player(N).PartyPlayer = 0
                Player(N).PartyStarter = NO
                Player(N).InParty = NO
            End If
        Else
            Call PlayerMsg(index, "No estas en un grupo!", Red)
        End If
        Exit Sub

        Case "partychat"
        Dim i As Long
            If Player(index).PartyID > 0 Then
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(Player(index).PartyID).Member(i) <> 0 Then Call PlayerMsg(Party(Player(index).PartyID).Member(i), GetPlayerName(index) & "-" & Parse(1), Green)
                Next
            Else
                Call PlayerMsg(index, "No estas en un grupo!", Red)
            End If
            Exit Sub


        Case "spells"
            Call Packet_Spells(index)
            Exit Sub

        Case "hotscript"
            Call Packet_HotScript(index, Val(Parse(1)))
            Exit Sub

        Case "scripttile"
            Call Packet_ScriptTile(index, Val(Parse(1)))
            Exit Sub

        Case "cast"
            Call Packet_Cast(index, Val(Parse(1)))
            Exit Sub

        Case "refresh"
            Call Packet_Refresh(index)
            Exit Sub

        Case "buysprite"
            Call Packet_BuySprite(index)
            Exit Sub

        Case "clearowner"
            Call Packet_ClearOwner(index)
            Exit Sub

        Case "requestedithouse"
            Call Packet_RequestEditHouse(index)
            Exit Sub

        Case "buyhouse"
            Call Packet_BuyHouse(index)
            Exit Sub

        Case "checkcommands"
            Call Packet_CheckCommands(index, Parse(1))
            Exit Sub

        Case "prompt"
            Call Packet_Prompt(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub

        Case "querybox"
            Call Packet_QueryBox(index, Parse(1), Val(Parse(2)))
            Exit Sub

        Case "requesteditarrow"
            Call Packet_RequestEditArrow(index)
            Exit Sub

        Case "editarrow"
            Call Packet_EditArrow(index, Val(Parse(1)))
            Exit Sub

        Case "savearrow"
            Call Packet_SaveArrow(index, Val(Parse(1)), Parse(2), Val(Parse(3)), Val(Parse(4)), Val(Parse(5)))
            Exit Sub

        Case "checkarrows"
            Call Packet_CheckArrows(index, Val(Parse(1)))
            Exit Sub

        Case "requesteditemoticon"
            Call Packet_RequestEditEmoticon(index)
            Exit Sub

        Case "requesteditelement"
            Call Packet_RequestEditElement(index)
            Exit Sub

        Case "editemoticon"
            Call Packet_EditEmoticon(index, Val(Parse(1)))
            Exit Sub

        Case "editelement"
            Call Packet_EditElement(index, Val(Parse(1)))
            Exit Sub

        Case "saveemoticon"
            Call Packet_SaveEmoticon(index, Val(Parse(1)), Parse(2), Val(Parse(3)))
            Exit Sub

        Case "saveelement"
            Call Packet_SaveElement(index, Val(Parse(1)), Parse(2), Val(Parse(3)), Val(Parse(4)))
            Exit Sub

        Case "checkemoticons"
            Call Packet_CheckEmoticon(index, Val(Parse(1)))
            Exit Sub

        Case "mapreport"
            Call Packet_MapReport(index)
            Exit Sub

        Case "gmtime"
            Call Packet_GMTime(index, Val(Parse(1)))
            Exit Sub

        Case "weather"
            Call Packet_Weather(index, Val(Parse(1)))
            Exit Sub

        Case "warpto"
            Call Packet_WarpTo(index, Val(Parse(1)), Val(Parse(2)), Val(Parse(3)))
            Exit Sub

        Case "localwarp"
            Call Packet_LocalWarp(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub

        Case "arrowhit"
            Call Packet_ArrowHit(index, Val(Parse(1)), Val(Parse(2)), Val(Parse(3)), Val(Parse(4)))
            Exit Sub

        Case "bankdeposit"
            Call Packet_BankDeposit(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub

        Case "bankwithdraw"
            Call Packet_BankWithdraw(index, Val(Parse(1)), Val(Parse(2)))
            Exit Sub

        Case "reloadscripts"
            Call Packet_ReloadScripts(index)
            Exit Sub

        Case "custommenuclick"
            Call Packet_CustomMenuClick(index, Val(Parse(1)), Val(Parse(2)), Parse(3), Val(Parse(4)), Parse(5))
            Exit Sub

        Case "returningcustomboxmsg"
            Call Packet_CustomBoxReturnMsg(index, Val(Parse(1)))
            Exit Sub
            
        Case "requesteditquest"
           Call callrequstedEditQuest(index)
           Exit Sub
           
           Case "giveitem"

            If GetPlayerAccess(index) < ADMIN_CREATOR Then
                Call HackingAttempt(index, "Trying to use powers not available")
                Exit Sub
            End If

            ' The index
            N = FindPlayer(Parse(1))

            i = Val(Parse(2))

            Call GMGiveItem(N, i)
           
            Exit Sub
           
        Case "takeitem"

            If GetPlayerAccess(index) < ADMIN_CREATOR Then
                Call HackingAttempt(index, "Trying to use powers not available")
                Exit Sub
            End If

            ' The index
            N = FindPlayer(Parse(1))

            i = Val(Parse(2))

            Call GMTakeItem(N, i)
           
            Exit Sub

    Case "editquest"
    
        N = Val(Parse(1))
        If N < 0 Or N > MAX_QUESTS Then
            Call HackingAttempt(index, "Invalid Quest Index")
            Exit Sub
        End If
        Call AddLog(GetPlayerName(index) & " editando quest #" & N & ".", ADMIN_LOG)
        Call SendEditQuestTo(index, N)
        Exit Sub

    Case "savequest"
        N = Val(Parse(1))
        If N < 0 Or N > MAX_QUESTS Then
            Call HackingAttempt(index, "Invalid Quests Index")
            Exit Sub
        End If
        Debug.Print Parse(5) & Parse(6)
        Quest(N).Name = Parse(2)
        Quest(N).After = Parse(3)
        Quest(N).Before = Parse(4)
        Quest(N).ClassIsReq = Val(Parse(5))
        Quest(N).ClassReq = Val(Parse(6))
        Quest(N).During = Parse(7)
        Quest(N).End = Parse(8)
        Quest(N).ItemReq = Val(Parse(9))
        Quest(N).ItemVal = Val(Parse(10))
        Quest(N).LevelIsReq = Val(Parse(11))
        Quest(N).LevelReq = Val(Parse(12))
        Quest(N).NotHasItem = Parse(13)
        Quest(N).RewardNum = Val(Parse(14))
        Quest(N).RewardVal = Val(Parse(15))
        Quest(N).Start = Parse(16)
        Quest(N).StartItem = Val(Parse(17))
        Quest(N).StartOn = Val(Parse(18))
        Quest(N).Startval = Val(Parse(19))
        Quest(N).FishExp = Val(Parse(20))
        Quest(N).MineExp = Val(Parse(21))
        Quest(N).LJackingExp = Val(Parse(22))
        Quest(N).ForagingExp = Val(Parse(23))
        Quest(N).UnArmedExp = Val(Parse(24))
        Quest(N).MageWeaponsExp = Val(Parse(25))
        Quest(N).CombatExp = Val(Parse(26))
        Quest(N).SmeltingExp = Val(Parse(27))
        Quest(N).IronForgingExp = Val(Parse(28))
        Quest(N).LeaderShipExp = Val(Parse(29))
        Quest(N).GovernmentExp = Val(Parse(30))
        Quest(N).CriticalHitExp = Val(Parse(31))
        Quest(N).DodgeExp = Val(Parse(32))
        Quest(N).RepExp = Val(Parse(33))
        Quest(N).PKillExp = Val(Parse(34))
        Quest(N).ThiefExp = Val(Parse(35))
        Quest(N).LargeBladesExp = Val(Parse(36))
        Quest(N).SmallBladesExp = Val(Parse(37))
        Quest(N).BluntWeaponsExp = Val(Parse(38))
        Quest(N).PolesExp = Val(Parse(39))
        Quest(N).AxesExp = Val(Parse(40))
        Quest(N).ThrownExp = Val(Parse(41))
        Quest(N).XbowsExp = Val(Parse(42))
        Quest(N).BowsExp = Val(Parse(43))
        Quest(N).CarpentryExp = Val(Parse(44))
        Quest(N).MillingExp = Val(Parse(45))
        Quest(N).SpinningExp = Val(Parse(46))
        Quest(N).WeavingExp = Val(Parse(47))
        Quest(N).SewingExp = Val(Parse(48))
        Quest(N).PlantingExp = Val(Parse(49))
        Quest(N).HarvestingExp = Val(Parse(50))
        Quest(N).LeatherWorkingExp = Val(Parse(51))
        Quest(N).SkinningExp = Val(Parse(52))
        Quest(N).TanningExp = Val(Parse(53))
        Quest(N).BodyExp = Val(Parse(54))
        Quest(N).MindExp = Val(Parse(55))
        Quest(N).SoulExp = Val(Parse(56))
        Quest(N).NatureExp = Val(Parse(57))
        Quest(N).AlchemyExp = Val(Parse(58))
        Quest(N).QuestingExp = Val(Parse(59))
        Quest(N).FirstAidExp = Val(Parse(60))
        Quest(N).QuestExpReward = Val(Parse(61))
        Call SendUpdateQuestToAll(N)
        Call SaveQuest(N)
        Call AddLog(GetPlayerName(index) & " guarda quest #" & N & ".", ADMIN_LOG)
        Exit Sub
        
        Case "acceptquest"
        Call ActuallyStartQuest(Val(Parse(1)), index, Val(Parse(2)))
        Exit Sub
        
        Case "acceptkillquest"
        Call ActuallyStartKillQuest(index, Val(Parse(1)), Parse(2), Parse(3))
        Exit Sub
        
        Case "stopkillquest"
        Call StopKillQuest(index)
        Exit Sub
        
        Case "questdone"
        Call GiveRewardItem(index, Quest(Val(Parse(1))).RewardNum, Quest(Val(Parse(1))).RewardVal, Val(Parse(3)))
        Exit Sub
        
        Case "profile"
        Call Packet_Profile(index, Parse(1))
        Exit Sub
        
        Case "openchest"
            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_CHEST Then
                If Val(ReadINI(GetPlayerName(index), "Cofre" & GetPlayerMap(index) & "," & GetPlayerX(index) & "," & (GetPlayerY(index) - 1), App.Path & "\Cofres.ini", 0)) = 0 Then
                    Call PlayerMsg(index, "Este cofre contiene " & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data2 & " " & Trim$(Item(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).Name), 1)
                    
                    '6dragon6 chest fix
                    If Item(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1).Stackable = True Then
                        Call GiveItem(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data2)
                    Else
                        For i = 1 To Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data2
                            Call GiveItem(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data1, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Data2)
                        Next
                        
                    End If
                        Call WriteINI(GetPlayerName(index), "Cofre" & GetPlayerMap(index) & "," & GetPlayerX(index) & "," & (GetPlayerY(index) - 1), 1, App.Path & "\Cofres.ini")
                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "cofre" & END_CHAR)
                Else
                
                    Call PlayerMsg(index, "Ya abriste este cofre antes!", Red)
                End If
            Else
                Call HackingAttempt(index, "No intentes hacer trucos con esto!!")
            End If
            Exit Sub
            
            
            Case "57"
        If GetPlayerParty(index) = 0 Then
        Call PlayerMsg(index, "Hacking attempt.", YELLOW)
        Exit Sub
        End If
        If Player(index).TargetType = TARGET_TYPE_NPC Then
        Call PlayerMsg(index, "Trata de invitar a nuevos jugadores a tu grupo para una mejor experiencia de juego.", YELLOW)
        Exit Sub
        End If
        If IsPlaying(Player(index).Target) = False Then
        Call PlayerMsg(index, "Este jugador no está jugando en este momento.", YELLOW)
        Exit Sub
        End If
        If index = Player(index).Target Then
        Call PlayerMsg(index, "No puedes invitarte a ti mismo a tu propio grupo.", YELLOW)
        Exit Sub
        End If
        If Not GetPlayerParty(Player(index).Target) = 0 Then
        Call PlayerMsg(index, "Este jugador se encuentra en otro grupo actualmente.", YELLOW)
        Exit Sub
        End If
        If GetPlayerLevel(index) > (GetPlayerLevel(Player(index).Target) + 3) Or GetPlayerLevel(index) < (GetPlayerLevel(Player(index).Target) - 3) Then
        Call PlayerMsg(index, "No puedes invitar a jugadores mayor o menor de 3 niveles con respecto a ti.", YELLOW)
        Exit Sub
        End If
        If Not Party(GetPlayerParty(index)).Leader = index Then
        Call PlayerMsg(index, "Solo el lider del grupo puede invitar a mas gente al grupo.", YELLOW)
        Else
         Call InvitePlayerToParty(index, Player(index).Target)
       End If
        Exit Sub

' :::::::::::::::::::
' :: Invite packet ::
' :::::::::::::::::::
   Case "58"
   N = FindPlayer(Parse(1))
   If GetPlayerMap(N) <> GetPlayerMap(index) Then
   Call PlayerMsg(index, "El jugador al que invitas debe estar almenos en el mismo mapa que tu estas.", YELLOW)
   Exit Sub
   End If
   Call InvitePlayerToParty(index, N)
   Exit Sub
   
   Case "f6"
   If GetPlayerParty(index) = 0 Then
   Call PlayerMsg(index, "No estas en un grupo!", Red)
   Exit Sub
   End If
   If Party(GetPlayerParty(index)).Member(Val(Parse(1))) = 0 Then
   Call PlayerMsg(index, "Este miembro no existe", Red)
   Exit Sub
   End If
   
   If GetPlayerMap(Party(GetPlayerParty(index)).Member(Val(Parse(1)))) = GetPlayerMap(index) Then
   Player(index).Target = Party(GetPlayerParty(index)).Member(Val(Parse(1)))
   Player(index).TargetType = TARGET_TYPE_PLAYER
   Call PlayerMsg(index, "Tu objetivo es ahora " & GetPlayerName(Party(GetPlayerParty(index)).Member(Val(Parse(1)))) & ".", Green)
   End If
   Exit Sub
   
   
   ' ::::::::::::::::::::::::
    ' :: Leave party packet ::
    ' ::::::::::::::::::::::::
    Case "64"
    ' Check if the player was in the party, if so, remove them.
    If GetPlayerParty(index) > 0 Then
    Call PartyRemoval(index, GetPlayerParty(index), Trim$(GetPlayerName(index)))
    Exit Sub
    Else
    Call PlayerMsg(index, "No estas en un grupo en este momento.", Red)
    Exit Sub
    End If
    
    Case "62"
    ' A player has accepted the invitation.
    
    If GetPlayerInvited(index) = 0 Then
    Call PlayerMsg(index, "No tienes ningun grupo ya que no has sido invitado.", Red)
    Exit Sub
    End If
    
    If Party(GetPlayerInvited(index)).Created = False And GetPlayerInvited(index) > 0 Then
    Call PlayerMsg(index, "No puedes entrar en este grupo ya que ya no existe.", Red)
    Player(index).Char(Player(index).CharNum).PartyInvitedTo = 0
    Player(index).Char(Player(index).CharNum).PartyInvitedToBy = ""
    Exit Sub
    End If
    
    If Not Trim$(Player(index).Char(Player(index).CharNum).PartyInvitedToBy) = Trim$(GetPlayerName(Party(GetPlayerInvited(index)).Leader)) Then
    Call PlayerMsg(index, "La persona que te ha invitado ahora mismo ya no es el lider, por lo que esta invitación no es valida.", Red)
    Player(index).Char(Player(index).CharNum).PartyInvitedTo = 0
    Player(index).Char(Player(index).CharNum).PartyInvitedToBy = ""
    Exit Sub
    End If
    
    If Not GetPlayerMap(index) = GetPlayerMap(Party(GetPlayerInvited(index)).Leader) Then
    Call PlayerMsg(index, "Necesitas estar en el mismo mapa que el lider del grupo para aceptar la invitación.", Red)
    Exit Sub
    End If
    
'6dragon6 group members fix
    If Party(GetPlayerInvited(index)).Created = True Then
       Call SetPlayerParty(index, GetPlayerInvited(index))
             
        Exit Sub
    End If
     
        Case "c7"
        If GetPlayerParty(index) = 0 Then
        Call PlayerMsg(index, "Hacking attempt.", Red)
        Exit Sub
        End If
        If Player(index).TargetType = TARGET_TYPE_NPC Then
        Call PlayerMsg(index, "Porque quieres invitar a un NPC a tu grupo?", YELLOW)
        Exit Sub
        End If
        If IsPlaying(Player(index).Target) = False Then
        Call PlayerMsg(index, "Este jugador no esta jugando en este momento.", YELLOW)
        Exit Sub
        End If
        If Not GetPlayerParty(Player(index).Target) = GetPlayerParty(index) Then
        Call PlayerMsg(index, "Este jugador no esta en tu grupo.", YELLOW)
        Exit Sub
        End If
        If index = Player(index).Target Then
        Call PlayerMsg(index, "No puedes eliminarte a ti mismo con esta opción. En el menu de grupo click en salir, o escribe /salir.", YELLOW)
        Exit Sub
        End If
        If Not Party(GetPlayerParty(index)).Leader = index Then
        Call PlayerMsg(index, "Solo el lider del grupo puede eliminar jugadores de ella.", YELLOW)
        Else
       ' Call PartyMsg(GetPlayerParty(Index), "The leader of the party, " & GetPlayerName(Index) & ", has removed " & GetPlayerName(Player(Index).Target) & " from the party.", Yellow)
        Call PartyRemoval(Player(index).Target, GetPlayerParty(index), GetPlayerName(Player(index).Target))
       End If
        Exit Sub
    
    
        Case "d"
    Dim z As Long
    
    z = GetPlayerParty(index)
    If z <> 0 Then
        Call PartyMsg(z, Parse(1))
    Else
    Call PlayerMsg(index, "No estás en un Party", PINK)
    End If
    Exit Sub
    Case "i2"
        Call CreateParty(index)
        Exit Sub
    
    Case "guildmsg"
    Dim y As String
    
    y = GetPlayerGuild(index)
    
    If y <> vbNullString Then
        Call guildMsg(y, Parse(1))
    Else
    Call PlayerMsg(index, "No estás en un Guild", BLACK)
    End If
    Exit Sub

            
    End Select

    Call HackingAttempt(index, "Recibido paquete invalido: " & Parse(0))
End Sub

Public Sub Packet_GetClasses(ByVal index As Long)
    Call SendNewCharClasses(index)
End Sub

Public Sub Packet_NewAccount(ByVal index As Long, ByVal Username As String, ByVal Password As String, ByVal Email As String)
    If Not IsLoggedIn(index) Then
        If LenB(Username) < 6 Then
            Call PlainMsg(index, "Tu usuario debe ser mayor de 3 caracteres.", 1)
            Exit Sub
        End If

        If LenB(Password) < 6 Then
            Call PlainMsg(index, "Tu contraseña debe ser mayor de 3 caracteres.", 1)
            Exit Sub
        End If

        If EMAIL_AUTH = 1 Then
            If LenB(Email) = 0 Then
                Call PlainMsg(index, "La dirección de e-mail no puedes dejarla en blanco.", 1)
                Exit Sub
            End If
        End If

        If Not IsAlphaNumeric(Username) Then
            Call PlainMsg(index, "Tu usuario debe contener caracteres alfa-numericos.!", 1)
            Exit Sub
        End If

        If Not IsAlphaNumeric(Password) Then
            Call PlainMsg(index, "Tu contraseña debe contener caracteres alfa-numericos.!", 1)
            Exit Sub
        End If

        If Not AccountExists(Username) Then
            Call AddAccount(index, Username, Password, Email)
            Call PlainMsg(index, "Tu cuenta ha sido creada con exito.", 0)
        Else
            Call PlainMsg(index, "Lo sentimos, el usuario ya existe.", 1)
        End If
    End If
End Sub

Public Sub Packet_DeleteAccount(ByVal index As Long, ByVal Username As String, ByVal Password As String)
    Dim i As Long
    
    If Not IsLoggedIn(index) Then
        If Not IsAlphaNumeric(Username) Then
            Call PlainMsg(index, "Tu usuario debe contener caracteres alfa-numericos.", 2)
            Exit Sub
        End If

        If Not IsAlphaNumeric(Password) Then
            Call PlainMsg(index, "Tu contraseña debe contener caracteres alfa-numericos.", 2)
            Exit Sub
        End If

        If Not AccountExists(Username) Then
            Call PlainMsg(index, "El nombre de usuario no existe.", 2)
            Exit Sub
        End If

        If Not PasswordOK(Username, Password) Then
            Call PlainMsg(index, "Has introducido una contraseña incorrecta.", 2)
            Exit Sub
        End If
    
        Call LoadPlayer(index, Username)
        For i = 1 To MAX_CHARS
            If LenB(Trim$(Player(index).Char(i).Name)) <> 0 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnEraseChar " & index & "," & i
                Call DeleteName(Player(index).Char(i).Name)
            End If
        Next i
        Call ClearPlayer(index)

        ' Remove the users main player profile.
        Kill App.Path & "\Cuentas\" & Username & "_Info.ini"
        Kill App.Path & "\Cuentas\" & Username & "\*.*"

        ' Delete the users account directory.
        RmDir App.Path & "\Cuentas\" & Username & "\"
    
        Call PlainMsg(index, "Tu cuenta ha sido borrada.", 0)
    End If
End Sub

Public Sub Packet_AccountLogin(ByVal index As Long, ByVal Username As String, ByVal Password As String, ByVal Major As Long, ByVal Minor As Long, ByVal Revision As Long, ByVal Code As String)
    If IsLoggedIn(index) <> True Then
        ' I'll re-add this when I change it to the new DAT method. [Mellowz]
        'If ACC_VERIFY = 1 Then
        '    If Val(ReadINI("GENERAL", "verified", App.Path & "\Cuentas\" & Trim$(Player(Index).Login) & ".ini")) = 0 Then
        '        Call PlainMsg(Index, "Your account hasn't been verified yet!", 3)
        '        Exit Sub
        '    End If
        'End If

        If Major < CLIENT_MAJOR Or Minor < CLIENT_MINOR Or Revision < CLIENT_REVISION Then
            Call PlainMsg(index, "Versión desactualizada, por favor visita " & Trim$(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "WebSite")), 3)
            Exit Sub
        End If

        If Not IsAlphaNumeric(Username) Then
            Call PlainMsg(index, "Tu usuario debe contener caracteres alfa-numericos.", 3)
            Exit Sub
        End If

        If Not IsAlphaNumeric(Password) Then
            Call PlainMsg(index, "Tu contraseña debe contener caracteres alfa-numericos.", 3)
            Exit Sub
        End If

        If Not AccountExists(Username) Then
            Call PlainMsg(index, "El nombre de usuario no existe", 3)
            Exit Sub
        End If
    
        If Not PasswordOK(Username, Password) Then
            Call PlainMsg(index, "Has introducido una contraseña incorrecta.", 3)
            Exit Sub
        End If
    
        If IsMultiAccounts(Username) Then
            Call PlainMsg(index, "El uso de cuentas multiples no esta permitido.", 3)
            Exit Sub
        End If
    
        If frmServer.Closed.Value = Checked Then
            Call PlainMsg(index, "El servidor esta cerrado para los usuarios.", 3)
            Exit Sub
        End If
    
        If Code <> SEC_CODE Then
            Call AlertMsg(index, "La contraseña del cliente no coincide con la del servidor.")
            Exit Sub
        End If
    
        Call LoadPlayer(index, Username)
        Call SendChars(index)
    
        Call TextAdd(frmServer.txtText(0), GetPlayerLogin(index) & " ha entrado desde " & GetPlayerIP(index) & ".", True)
    Else
        Call AlertMsg(index, "Ya estás logeado, tu cuenta será desconectada!")
    End If
End Sub

Public Sub Packet_Profile(ByVal index As Long, ByVal Name As String)
Dim Player

Player = FindPlayer(Name)

    If Player = 0 Then
        Call PlayerMsg(index, Player & " esta actualmente desconectado.", WHITE)
        Exit Sub
    End If
   
    If scripting = 1 Then
    MyScript.ExecuteStatement "Scripts\Main.txt", "PlayersProfile " & index & "," & Player
    End If
End Sub

Public Sub Packet_GiveMeTheMax(ByVal index As Long)
    Dim Packet As String
    Dim Pass As String
    
    Pass = GetSetting(App.EXEName, "Clave", "Clave")

    Packet = "MAXINFO" & SEP_CHAR
    Packet = Packet & GAME_NAME & SEP_CHAR
    Packet = Packet & MAX_PLAYERS & SEP_CHAR
    Packet = Packet & MAX_ITEMS & SEP_CHAR
    Packet = Packet & MAX_NPCS & SEP_CHAR
    Packet = Packet & MAX_SHOPS & SEP_CHAR
    Packet = Packet & MAX_SPELLS & SEP_CHAR
    Packet = Packet & MAX_MAPS & SEP_CHAR
    Packet = Packet & MAX_MAP_ITEMS & SEP_CHAR
    Packet = Packet & MAX_MAPX & SEP_CHAR
    Packet = Packet & MAX_MAPY & SEP_CHAR
    Packet = Packet & MAX_EMOTICONS & SEP_CHAR
    Packet = Packet & MAX_ELEMENTS & SEP_CHAR
    Packet = Packet & paperdoll & SEP_CHAR
    Packet = Packet & SPRITESIZE & SEP_CHAR
    Packet = Packet & MAX_SCRIPTSPELLS & SEP_CHAR
    Packet = Packet & CUSTOM_SPRITE & SEP_CHAR
    Packet = Packet & Level & SEP_CHAR
    Packet = Packet & MAX_PARTY_MEMBERS & SEP_CHAR
    Packet = Packet & stat1 & SEP_CHAR
    Packet = Packet & stat2 & SEP_CHAR
    Packet = Packet & stat3 & SEP_CHAR
    Packet = Packet & stat4 & SEP_CHAR
    Packet = Packet & MAX_HEAD & SEP_CHAR
    Packet = Packet & MAX_BODY & SEP_CHAR
    Packet = Packet & MAX_LEGS & SEP_CHAR
    Packet = Packet & Pass & SEP_CHAR
    Packet = Packet & PKCr(1) & SEP_CHAR
    Packet = Packet & PKCr(2) & SEP_CHAR
    Packet = Packet & PKCr(3) & SEP_CHAR
    Packet = Packet & END_CHAR

    Call SendDataTo(index, Packet)
    Call SendNewsTo(index)
End Sub

Public Sub Packet_AddCharacter(ByVal index As Long, ByVal Name As String, ByVal Sex As Long, ByVal Class As Long, ByVal CharNum As Long, ByVal Head As Long, ByVal Body As Long, ByVal Leg As Long)
    If CharNum < 1 Or CharNum > MAX_CHARS Then
        Call HackingAttempt(index, "Invalid CharNum")
        Exit Sub
    End If
    
    If LenB(Name) < 6 Then
        Call HackingAttempt(index, "Invalid Name Length")
        Exit Sub
    End If
    
    If Sex <> SEX_MALE And Sex <> SEX_FEMALE Then
        Call HackingAttempt(index, "Invalid Sex")
        Exit Sub
    End If
    
    If Class < 0 Or Class > MAX_CLASSES Then
        Call HackingAttempt(index, "Invalid Class")
        Exit Sub
    End If

   If Not IsAlphaNumeric(Name) Then
        Call PlainMsg(index, "El nombre debe contener caracteres alfa-numericos.", 4)
        Exit Sub
    End If

    If CharExist(index, CharNum) Then
        Call PlainMsg(index, "El personaje ya existe.", 4)
        Exit Sub
    End If
    
    If FindChar(Name) Then
        Call PlainMsg(index, "Lo sentimos, pero el nombre ya esta en uso.", 4)
        Exit Sub
    End If

    Call AddChar(index, Name, Sex, Class, CharNum, Head, Body, Leg)
    Call SendChars(index)

    Call PlainMsg(index, "El personaje ha sido creado con exito.", 5)

    If scripting = 1 Then
        Call MyScript.ExecuteStatement("Scripts\Main.txt", "OnNewChar " & index & "," & CharNum)
    End If
End Sub

Public Sub Packet_DeleteCharacter(ByVal index As Long, ByVal CharNum As Long)
    If CharNum < 1 Or CharNum > MAX_CHARS Then
        Call HackingAttempt(index, "Invalid CharNum")
        Exit Sub
    End If
    
    If CharExist(index, CharNum) Then
        Call DelChar(index, CharNum)
        Call SendChars(index)
    
        Call PlainMsg(index, "El personaje ha sido eliminado.", 5)
    Else
        Call PlainMsg(index, "El personaje no existe.", 5)
    End If
End Sub

Public Sub Packet_UseCharacter(ByVal index As Long, ByVal CharNum As Long)
    Dim FileID As Integer

    If CharNum < 1 Or CharNum > MAX_CHARS Then
        Call HackingAttempt(index, "Invalid CharNum")
        Exit Sub
    End If
    
    If CharExist(index, CharNum) Then
        Player(index).CharNum = CharNum
    
        If frmServer.GMOnly.Value = Checked Then
            If GetPlayerAccess(index) = 0 Then
                Call PlainMsg(index, "El servidor solo esta disponible para el personal.", 5)
                Exit Sub
            End If
        End If
    
        Call JoinGame(index)

        Call TextAdd(frmServer.txtText(0), GetPlayerLogin(index) & "/" & GetPlayerName(index) & " ha entrado en " & GAME_NAME & ".", True)
        Call UpdateTOP
    
        If Not FindChar(GetPlayerName(index)) Then
            FileID = FreeFile
            Open App.Path & "\Cuentas\CharList.txt" For Append As #FileID
                Print #FileID, GetPlayerName(index)
            Close #FileID
        End If
    Else
        Call PlainMsg(index, "El personaje no existe!", 5)
    End If
End Sub

Public Sub Packet_GuildChangeAccess(ByVal index As Long, ByVal Name As String, ByVal Rank As Long)
    Dim NameIndex As Long
    
    If LenB(Name) = 0 Then
        Call PlayerMsg(index, "Necesitas insertar un nombre de usuario para proceder.", WHITE)
        Exit Sub
    End If

    If Rank < 0 Or Rank > 4 Then
        Call PlayerMsg(index, "Necesitas insertar un rango valido para proceder.", Red)
        Exit Sub
    End If

    NameIndex = FindPlayer(Name)

    If NameIndex = 0 Then
        Call PlayerMsg(index, Name & " esta actualmente en uso.", WHITE)
        Exit Sub
    End If

    If GetPlayerGuild(NameIndex) <> GetPlayerGuild(index) Then
        Call PlayerMsg(index, Name & " no esta en tu clan.", Red)
        Exit Sub
    End If

    If GetPlayerGuildAccess(index) < 4 Then
        Call PlayerMsg(index, "No eres el dueño de este clan.", Red)
        Exit Sub
    End If

    Call SetPlayerGuildAccess(NameIndex, Rank)
    Call SendPlayerData(NameIndex)
End Sub

Public Sub Packet_GuildDisown(ByVal index As Long, ByVal Name As String)
    Dim NameIndex As Long

    NameIndex = FindPlayer(Name)

    If NameIndex = 0 Then
        Call PlayerMsg(index, Name & " está actualmente desconectado.", WHITE)
        Exit Sub
    End If

    If GetPlayerGuild(NameIndex) <> GetPlayerGuild(index) Then
        Call PlayerMsg(index, Name & " no está en tu clan.", Red)
        Exit Sub
    End If

    If GetPlayerGuildAccess(NameIndex) > GetPlayerGuildAccess(index) Then
        Call PlayerMsg(index, Name & " tiene mas rango que tu en el clan.", Red)
        Exit Sub
    End If

    Call SetPlayerGuild(NameIndex, vbNullString)
    Call SetPlayerGuildAccess(NameIndex, 0)
    Call SendPlayerData(NameIndex)
End Sub

Public Sub Packet_GuildLeave(ByVal index As Long)
    If LenB(GetPlayerGuild(index)) = 0 Then
        Call PlayerMsg(index, "No estas en un clan.", Red)
        Exit Sub
    End If

    Call SetPlayerGuild(index, vbNullString)
    Call SetPlayerGuildAccess(index, 0)
    Call SendPlayerData(index)
End Sub

Public Sub Packet_GuildMake(ByVal index As Long, ByVal Name As String, ByVal Guild As String)
    Dim NameIndex As Long

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    NameIndex = FindPlayer(Name)

    If NameIndex = 0 Then
        Call PlayerMsg(index, Name & " está actualmente desconectado.", WHITE)
        Exit Sub
    End If

    If LenB(GetPlayerGuild(NameIndex)) <> 0 Then
        Call PlayerMsg(index, Name & " ya esta en tu clan.", Red)
        Exit Sub
    End If

    If LenB(Guild) = 0 Then
        Call PlayerMsg(index, "Por favor, introduce un nombre de clan valido.", Red)
        Exit Sub
    End If

    Call SetPlayerGuild(NameIndex, Guild)
    Call SetPlayerGuildAccess(NameIndex, 4)
    Call SendPlayerData(NameIndex)
End Sub

Public Sub Packet_GuildMember(ByVal index As Long, ByVal Name As String)
    Dim NameIndex As Long

    NameIndex = FindPlayer(Name)

    If NameIndex = 0 Then
        Call PlayerMsg(index, Name & " no esta conectado.", WHITE)
        Exit Sub
    End If

    If GetPlayerGuild(NameIndex) <> GetPlayerGuild(index) Then
        Call PlayerMsg(index, Name & " no esta en tu clan.", Red)
        Exit Sub
    End If

    If GetPlayerGuildAccess(NameIndex) > 1 Then
        Call PlayerMsg(index, Name & " ya ha sido admitido.", WHITE)
        Exit Sub
    End If

    Call SetPlayerGuild(NameIndex, GetPlayerGuild(index))
    Call SetPlayerGuildAccess(NameIndex, 1)
    Call SendPlayerData(NameIndex)
End Sub

Public Sub Packet_GuildTrainee(ByVal index As Long, ByVal Name As String)
    Dim NameIndex As Long

    NameIndex = FindPlayer(Name)

    If NameIndex = 0 Then
        Call PlayerMsg(index, Name & " no esta conectado.", WHITE)
        Exit Sub
    End If

    If LenB(GetPlayerGuild(NameIndex)) <> 0 Then
        Call PlayerMsg(index, Name & " ya esta en tu clan.", Red)
        Exit Sub
    End If

    Call SetPlayerGuild(NameIndex, GetPlayerGuild(index))
    Call SetPlayerGuildAccess(NameIndex, 0)
    Call SendPlayerData(NameIndex)
End Sub

Public Sub Packet_SayMessage(ByVal index As Long, ByVal Message As String)
    If frmServer.chkLogMap.Value = Unchecked Then
        If GetPlayerAccess(index) = 0 Then
            Call PlayerMsg(index, "Los mensajes en el mapa han sido desactivados!", BRIGHTRED)
            Exit Sub
        End If
    End If

    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & ": " & Message, SayColor)
    Call MapMsg2(GetPlayerMap(index), Message, index)

    Call TextAdd(frmServer.txtText(3), GetPlayerName(index) & " en el mapa " & GetPlayerMap(index) & ": " & Message, True)
    Call AddLog("Mapa #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " : " & Message, PLAYER_LOG)
End Sub

Public Sub Packet_EmoteMessage(ByVal index As Long, ByVal Message As String)
    If frmServer.chkLogEmote.Value = Unchecked Then
        If GetPlayerAccess(index) = 0 Then
            Call PlayerMsg(index, "Los emoticonos han sido desactivados en el mapa!", BRIGHTRED)
            Exit Sub
        End If
    End If

    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & ": " & Message, EmoteColor)

    Call TextAdd(frmServer.txtText(6), GetPlayerName(index) & " " & Message, True)
    Call AddLog("Mapa #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " " & Message, PLAYER_LOG)
End Sub

Public Sub Packet_BroadcastMessage(ByVal index As Long, ByVal Message As String)
    If frmServer.chkLogBC.Value = Unchecked Then
        If GetPlayerAccess(index) = 0 Then
            Call PlayerMsg(index, "Los mensajes de difusión han sido desactivados en el mapa!", BRIGHTRED)
            Exit Sub
        End If
    End If

    If Player(index).Mute Then
        Call PlayerMsg(index, "Estas silenciado, no puedes difundir mensajes.", BRIGHTRED)
        Exit Sub
    End If

    Call GlobalMsg(GetPlayerName(index) & ": " & Message, BroadcastColor)

    Call TextAdd(frmServer.txtText(0), GetPlayerName(index) & ": " & Message, True)
    Call TextAdd(frmServer.txtText(1), GetPlayerName(index) & ": " & Message, True)
    Call AddLog(GetPlayerName(index) & ": " & Message, PLAYER_LOG)
End Sub

Public Sub Packet_GlobalMessage(ByVal index As Long, ByVal Message As String)
    If frmServer.chkLogGlobal.Value = Unchecked Then
        If GetPlayerAccess(index) = 0 Then
            Call PlayerMsg(index, "Los mensajes globales han sido desactivados!", BRIGHTRED)
            Exit Sub
        End If
    End If

    If Player(index).Mute Then
        Call PlayerMsg(index, "Estas silenciado, no puedes difundir mensajes.", BRIGHTRED)
        Exit Sub
    End If

    If GetPlayerAccess(index) > 0 Then
        Call GlobalMsg("(Global) " & GetPlayerName(index) & ": " & Message, GlobalColor)

        Call TextAdd(frmServer.txtText(0), "(Global) " & GetPlayerName(index) & ": " & Message, True)
        Call TextAdd(frmServer.txtText(2), GetPlayerName(index) & ": " & Message, True)
        Call AddLog("(Global) " & GetPlayerName(index) & ": " & Message, ADMIN_LOG)
    End If
End Sub

Public Sub Packet_AdminMessage(ByVal index As Long, ByVal Message As String)
    If frmServer.chkLogAdmin.Value = Unchecked Then
        Call PlayerMsg(index, "Los mensajes de administradores han sido desactivados!", BRIGHTRED)
        Exit Sub
    End If

    If GetPlayerAccess(index) > 0 Then
        Call AdminMsg("(Admin " & GetPlayerName(index) & ") " & Message, AdminColor)

        Call TextAdd(frmServer.txtText(5), GetPlayerName(index) & ": " & Message, True)
        Call AddLog("(Admin " & GetPlayerName(index) & ") " & Message, ADMIN_LOG)
    End If
End Sub

Public Sub Packet_PlayerMessage(ByVal index As Long, ByVal Name As String, ByVal Message As String)
    Dim MsgTo As Long
    
    If frmServer.chkLogPM.Value = Unchecked Then
        If GetPlayerAccess(index) = 0 Then
            Call PlayerMsg(index, "Los mensajes privados han sido desactivados!", BRIGHTRED)
            Exit Sub
        End If
    End If

    If LenB(Name) = 0 Then
        Call PlayerMsg(index, "Necesitas seleccionar un nombre de jugador para enviar el MP.", BRIGHTRED)
        Exit Sub
    End If

    If LenB(Message) = 0 Then
        Call PlayerMsg(index, "Necesitas enviar el mensaje privado a otro jugador.", BRIGHTRED)
        Exit Sub
    End If

    MsgTo = FindPlayer(Name)

    If MsgTo = 0 Then
        Call PlayerMsg(index, Name & " no esta conectado.", WHITE)
        Exit Sub
    End If

    Call PlayerMsg(index, "Le dices " & GetPlayerName(MsgTo) & ", '" & Message & "'", TellColor)
    Call PlayerMsg(MsgTo, GetPlayerName(index) & " te dice, '" & Message & "'", TellColor)

    Call TextAdd(frmServer.txtText(4), "A " & GetPlayerName(MsgTo) & " de " & GetPlayerName(index) & ": " & Message, True)
    Call AddLog(GetPlayerName(index) & " dice " & GetPlayerName(MsgTo) & ", " & Message & "'", PLAYER_LOG)
End Sub

Public Sub Packet_PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal Movement As Long, Xpos As Integer, Ypos As Integer)
    If Player(index).GettingMap = YES Then
        Exit Sub
    End If

    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(index, "Invalid Direction")
        Exit Sub
    End If

    If Movement <> 1 And Movement <> 2 Then
        Call HackingAttempt(index, "Invalid Movement")
        Exit Sub
    End If

    If Player(index).CastedSpell = YES Then
        If GetTickCount > Player(index).AttackTimer + 1000 Then
            Player(index).CastedSpell = NO
        Else
            Call SendPlayerXY(index)
            Exit Sub
        End If
    End If

    If Player(index).Locked = True Then
        Call SendPlayerXY(index)
        Exit Sub
    End If

    Call PlayerMove(index, Dir, Movement, Xpos, Ypos)
End Sub

Public Sub Packet_PlayerDirection(ByVal index As Long, ByVal Dir As Long)
    If Player(index).GettingMap = YES Then
        Exit Sub
    End If

    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(index, "Invalid Direction")
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)

    Call SendDataToMapBut(index, GetPlayerMap(index), "PLAYERDIR" & SEP_CHAR & index & SEP_CHAR & GetPlayerDir(index) & END_CHAR)
End Sub

Public Sub Packet_UseItem(ByVal index As Long, ByVal InvNum As Long)
    Dim CharNum As Long
    Dim SpellID As Long
    Dim MinLvl As Long
    Dim x As Long
    Dim y As Long

    If InvNum < 1 Or InvNum > MAX_ITEMS Then
        Call HackingAttempt(index, "Invalid InvNum")
        Exit Sub
    End If

    If Player(index).LockedItems Then
        Call PlayerMsg(index, "Actualmente no puedes usar ningun objeto.", BRIGHTRED)
        Exit Sub
    End If

    CharNum = Player(index).CharNum

    Dim N As Long

    ' Find out what kind of item it is
    Select Case Item(GetPlayerInvItemNum(index, InvNum)).Type
        Case ITEM_TYPE_ARMOR
            If InvNum <> GetPlayerArmorSlot(index) Then
                If ItemIsUsable(index, InvNum) Then
                    Call SetPlayerArmorSlot(index, InvNum)
                End If
            Else
                Call SetPlayerArmorSlot(index, 0)
            End If
            Call SendWornEquipment(index)
    
        Case ITEM_TYPE_WEAPON
            If InvNum <> GetPlayerWeaponSlot(index) Then
                If ItemIsUsable(index, InvNum) Then
                    Call SetPlayerWeaponSlot(index, InvNum)
                End If
            Else
                Call SetPlayerWeaponSlot(index, 0)
            End If
            Call SendWornEquipment(index)
    
        Case ITEM_TYPE_TWO_HAND
            If InvNum <> GetPlayerWeaponSlot(index) Then
                If ItemIsUsable(index, InvNum) Then
                    If GetPlayerShieldSlot(index) <> 0 Then
                        Call SetPlayerShieldSlot(index, 0)
                    End If

                    Call SetPlayerWeaponSlot(index, InvNum)
                End If
            Else
                Call SetPlayerWeaponSlot(index, 0)
            End If
            Call SendWornEquipment(index)
    
        Case ITEM_TYPE_HELMET
            If InvNum <> GetPlayerHelmetSlot(index) Then
                If ItemIsUsable(index, InvNum) Then
                    Call SetPlayerHelmetSlot(index, InvNum)
                End If
            Else
                Call SetPlayerHelmetSlot(index, 0)
            End If
            Call SendWornEquipment(index)
    
        Case ITEM_TYPE_SHIELD
            If InvNum <> GetPlayerShieldSlot(index) Then
                If ItemIsUsable(index, InvNum) Then
                    If GetPlayerWeaponSlot(index) <> 0 Then
                        If Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Type = ITEM_TYPE_TWO_HAND Then
                            Call SetPlayerWeaponSlot(index, 0)
                        End If
                    End If

                    Call SetPlayerShieldSlot(index, InvNum)
                End If
            Else
                Call SetPlayerShieldSlot(index, 0)
            End If
            Call SendWornEquipment(index)
    
        Case ITEM_TYPE_LEGS
            If InvNum <> GetPlayerLegsSlot(index) Then
                If ItemIsUsable(index, InvNum) Then
                    Call SetPlayerLegsSlot(index, InvNum)
                End If
            Else
                Call SetPlayerLegsSlot(index, 0)
            End If
            Call SendWornEquipment(index)
    
        Case ITEM_TYPE_RING
            If InvNum <> GetPlayerRingSlot(index) Then
                If ItemIsUsable(index, InvNum) Then
                    Call SetPlayerRingSlot(index, InvNum)
                End If
            Else
                Call SetPlayerRingSlot(index, 0)
            End If
            Call SendWornEquipment(index)
    
        Case ITEM_TYPE_NECKLACE
            If InvNum <> GetPlayerNecklaceSlot(index) Then
                If ItemIsUsable(index, InvNum) Then
                    Call SetPlayerNecklaceSlot(index, InvNum)
                End If
            Else
                Call SetPlayerNecklaceSlot(index, 0)
            End If
            Call SendWornEquipment(index)

        Case ITEM_TYPE_POTIONADDHP
            Call SetPlayerHP(index, GetPlayerHP(index) + Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1)
            If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 1)
            Else
                Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
            End If
            Call SendHP(index)
    
        Case ITEM_TYPE_POTIONADDMP
            Call SetPlayerMP(index, GetPlayerMP(index) + Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1)
            If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 1)
            Else
                Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
            End If
            Call SendMP(index)
    
        Case ITEM_TYPE_POTIONADDSP
            Call SetPlayerSP(index, GetPlayerSP(index) + Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1)
            If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 1)
            Else
                Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
            End If
            Call SendSP(index)
    
        Case ITEM_TYPE_POTIONSUBHP
            Call SetPlayerHP(index, GetPlayerHP(index) - Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1)
            If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 1)
            Else
                Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
            End If
            Call SendHP(index)
    
        Case ITEM_TYPE_POTIONSUBMP
            Call SetPlayerMP(index, GetPlayerMP(index) - Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1)
            If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 1)
            Else
                Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
            End If
            Call SendMP(index)
    
        Case ITEM_TYPE_POTIONSUBSP
            Call SetPlayerSP(index, GetPlayerSP(index) - Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1)
            If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 1)
            Else
                Call TakeItem(index, Player(index).Char(CharNum).Inv(InvNum).num, 0)
            End If
            Call SendSP(index)
    
        Case ITEM_TYPE_KEY
            Select Case GetPlayerDir(index)
                Case DIR_UP
                    If GetPlayerY(index) > 0 Then
                        x = GetPlayerX(index)
                        y = GetPlayerY(index) - 1
                    Else
                        Exit Sub
                    End If
    
                Case DIR_DOWN
                    If GetPlayerY(index) < MAX_MAPY Then
                        x = GetPlayerX(index)
                        y = GetPlayerY(index) + 1
                    Else
                        Exit Sub
                    End If
    
                Case DIR_LEFT
                    If GetPlayerX(index) > 0 Then
                        x = GetPlayerX(index) - 1
                        y = GetPlayerY(index)
                    Else
                        Exit Sub
                    End If
    
                Case DIR_RIGHT
                    If GetPlayerX(index) < MAX_MAPX Then
                        x = GetPlayerX(index) + 1
                        y = GetPlayerY(index)
                    Else
                        Exit Sub
                    End If
            End Select
    
            ' Check if a key exists.
            If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_KEY Then
                ' Check if the key they are using matches the map key.
                If GetPlayerInvItemNum(index, InvNum) = Map(GetPlayerMap(index)).Tile(x, y).Data1 Then
                    TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                    TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
    
                    Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)

                    If Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1) = vbNullString Then
                        Call MapMsg(GetPlayerMap(index), "La puerta se ha desbloqueado!", WHITE)
                    Else
                        Call MapMsg(GetPlayerMap(index), Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1), WHITE)
                    End If

                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "key" & END_CHAR)
    
                    ' Check if we are supposed to take away the item.
                    If Map(GetPlayerMap(index)).Tile(x, y).Data2 = 1 Then
                        Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                        Call PlayerMsg(index, "La llave se disolvio.", YELLOW)
                    End If
                End If
            End If
    
            If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_DOOR Then
                TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount
    
                Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "key" & END_CHAR)
            End If
    
        Case ITEM_TYPE_SPELL
            SpellID = Item(GetPlayerInvItemNum(index, InvNum)).Data1
    
            If SpellID > 0 Then
                If Spell(SpellID).ClassReq - 1 = GetPlayerClass(index) Or Spell(SpellID).ClassReq = 0 Then
                    If Spell(SpellID).LevelReq = 0 And Player(index).Char(Player(index).CharNum).Access < 1 Then
                        Call PlayerMsg(index, "Este hechizo solo puede ser usado por administradores!", BRIGHTRED)
                        Exit Sub
                    End If

                    MinLvl = GetSpellReqLevel(SpellID)

                    If MinLvl <= GetPlayerLevel(index) Then
                        MinLvl = FindOpenSpellSlot(index)
    
                        If MinLvl > 0 Then
                            If Not HasSpell(index, SpellID) Then
                                Call SetPlayerSpell(index, MinLvl, SpellID)
                                Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                                Call PlayerMsg(index, "Has aprendido un nuevo hechizo.", WHITE)
                            Else
                                Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                                Call PlayerMsg(index, "Ya habias aprendido este hechizo. *El hechizo se deshace*.", BRIGHTRED)
                            End If
                        Else
                            Call PlayerMsg(index, "Ya has aprendido todo lo que debias aprender.", BRIGHTRED)
                        End If
                    Else
                        Call PlayerMsg(index, "Necesitas alcanzar el nivel " & MinLvl & " para aprender este hechizo.", WHITE)
                    End If
                Else
                    Call PlayerMsg(index, "Este hechizo solo puede ser aprendido por " & GetClassName(Spell(SpellID).ClassReq - 1) & ".", WHITE)
                End If
            End If
    
        Case ITEM_TYPE_SCRIPTED
            If scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedItem " & index & "," & Item(Player(index).Char(CharNum).Inv(InvNum).num).Data1
            End If
        
        Case ITEM_TYPE_PET
            Dim PetSprite As Long
                    PetSprite = Item(GetPlayerInvItemNum(index, InvNum)).PetSprite
                   
                    Call PetSub(index, PetSprite)
                    Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
        
        Case ITEM_TYPE_PETREZ
                    Call RezPet(index)
                    Call SendPetData(index)
                    Call TakeItem(index, GetPlayerInvItemNum(index, InvNum), 0)
    End Select
    
    Call SendStats(index)
    Call SendHP(index)
    Call SendMP(index)
    Call SendSP(index)

    Call SendIndexWornEquipment(index)
End Sub

' This packet seems to me like it's incomplete. [Mellowz]
Public Sub Packet_PlayerMoveMouse(ByVal index As Long, ByVal Dir As Long)
    If Player(index).GettingMap = YES Then
        Exit Sub
    End If

    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(index, "Invalid Direction")
        Exit Sub
    End If

    If Player(index).Locked = True Then
        Call SendPlayerXY(index)
        Exit Sub
    End If

    If Player(index).CastedSpell = YES Then
        If GetTickCount > Player(index).AttackTimer + 1000 Then
            Player(index).CastedSpell = NO
        Else
            Call SendPlayerXY(index)
            Exit Sub
        End If
    End If

    If Val(ReadINI("CONFIG", "mouse", App.Path & "\Configuracion.ini", "0")) = 1 Then
        Call SendDataTo(index, "mouse" & END_CHAR)
    End If
End Sub

Public Sub Packet_Warp(ByVal index As Long, ByVal Dir As Long)
    Select Case Dir
        Case DIR_UP
            If Map(GetPlayerMap(index)).Up > 0 Then
                Call PlayerWarp(index, Map(GetPlayerMap(index)).Up, GetPlayerX(index), MAX_MAPY)
                Exit Sub
            End If

        Case DIR_DOWN
            If Map(GetPlayerMap(index)).Down > 0 Then
                Call PlayerWarp(index, Map(GetPlayerMap(index)).Down, GetPlayerX(index), 0)
                Exit Sub
            End If

        Case DIR_LEFT
            If Map(GetPlayerMap(index)).Left > 0 Then
                Call PlayerWarp(index, Map(GetPlayerMap(index)).Left, MAX_MAPX, GetPlayerY(index))
                Exit Sub
            End If

        Case DIR_RIGHT
            If Map(GetPlayerMap(index)).Right > 0 Then
                Call PlayerWarp(index, Map(GetPlayerMap(index)).Right, 0, GetPlayerY(index))
                Exit Sub
            End If
    End Select
End Sub

Public Sub Packet_EndShot(ByVal index As Long, ByVal Unknown As Long)
    If Unknown = 0 Then
        Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & END_CHAR)
        Player(index).Locked = False
        Player(index).HookShotX = 0
        Player(index).HookShotY = 0
        Exit Sub
    End If

    Call PlayerMsg(index, "Ten cuidado al cruzar el cable.", 1)

    Player(index).Locked = False

    Call SetPlayerX(index, Player(index).HookShotX)
    Call SetPlayerY(index, Player(index).HookShotY)

    Player(index).HookShotX = 0
    Player(index).HookShotY = 0

    Call SendPlayerXY(index)
End Sub

Public Sub Packet_Attack(ByVal index As Long)
    Dim i As Long
    Dim Damage As Long

    If Player(index).LockedAttack Then
        Exit Sub
    End If

    If GetPlayerWeaponSlot(index) > 0 Then
        If Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 > 0 Then
            If Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Stackable = 0 Then
                Call SendDataToMap(GetPlayerMap(index), "checkarrows" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & END_CHAR)
            Else
                Call GrapleHook(index)
            End If

            Exit Sub
        End If
    End If

    ' Try to attack another player.
    For i = 1 To MAX_PLAYERS
        If i <> index Then
            If CanAttackPlayer(index, i) Then
            
                Player(index).Target = i
                Player(index).TargetType = TARGET_TYPE_PLAYER
            
                If Not CanPlayerBlockHit(i) Then
                    If Not CanPlayerCriticalHit(index) Then
                        Damage = GetPlayerDamage(index) - GetPlayerProtection(i)
                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & END_CHAR)
                    Else
                        Damage = GetPlayerDamage(index) + Int(Rnd * Int(GetPlayerDamage(index) / 2)) + 1 - GetPlayerProtection(i)

                        Call BattleMsg(index, "Sientes como una increible fuerza se apodera de ti!", BRIGHTCYAN, 0)
                        Call BattleMsg(i, GetPlayerName(index) & " desprende un increible poder!", BRIGHTCYAN, 1)

                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & END_CHAR)
                    End If

                    If Damage > 0 Then
                    If scripting = 1 Then
                        MyScript.ExecuteStatement "Scripts\Main.txt", "OnAttack " & index & "," & Damage
                    Else
                        Call AttackPlayer(index, i, Damage)
                    End If
                    Else
                        If scripting = 1 Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "OnAttack " & index & "," & Damage
                        End If
                        Call PlayerMsg(index, "Tu ataque no hace nada.", BRIGHTRED)
                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & END_CHAR)
                    End If
                Else
                    If scripting = 1 Then
                        MyScript.ExecuteStatement "Scripts\Main.txt", "OnAttack " & index & "," & 0
                    End If

                    Call BattleMsg(index, GetPlayerName(i) & " bloquea tu golpe!", BRIGHTCYAN, 0)
                    Call BattleMsg(i, "Bloqueas el golpe de " & GetPlayerName(index) & ".", BRIGHTCYAN, 1)

                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & END_CHAR)
                End If

                Exit Sub
            End If
        End If
    Next i

    ' Try to attack an NPC.
    For i = 1 To MAX_MAP_NPCS
        If CanAttackNpc(index, i) Then
            ' Get the damage we can do
            Player(index).TargetNPC = i
            Player(index).TargetType = TARGET_TYPE_NPC
            If Not CanPlayerCriticalHit(index) Then
                Damage = GetPlayerDamage(index) - Int(NPC(MapNPC(GetPlayerMap(index), i).num).DEF / 2)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & END_CHAR)
            Else
                Damage = GetPlayerDamage(index) + Int(Rnd * Int(GetPlayerDamage(index) / 2)) + 1 - Int(NPC(MapNPC(GetPlayerMap(index), i).num).DEF / 2)
                Call BattleMsg(index, "Sientes como una increible fuerza se apodera de ti!", BRIGHTCYAN, 0)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & END_CHAR)
            End If
            
            

            If Damage > 0 Then
                If scripting = 1 Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "OnAttack " & index & "," & Damage
                Else
                    Call AttackNpc(index, i, Damage)
                    Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & END_CHAR)
                End If
            Else
                If scripting = 1 Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "OnAttack " & index & "," & Damage
                End If
                
                Call BattleMsg(index, "Tu ataque no hace nada.", BRIGHTRED, 0)

                Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & END_CHAR)
            End If

            Exit Sub
        End If
    Next i
End Sub

Public Sub Packet_UseStatPoint(ByVal index As Long, ByVal PointType As Long)
    If PointType < 0 Or PointType > 3 Then
        Call HackingAttempt(index, "Invalid Point Type")
        Exit Sub
    End If

    If GetPlayerPOINTS(index) > 0 Then
        If scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "UsingStatPoints " & index & "," & PointType
        Else
            Select Case PointType
                Case 0
                    Call SetPlayerSTR(index, GetPlayerSTR(index) + 1)
                    Call BattleMsg(index, "Tu fuerza ha incrementado!", 15, 0)

                Case 1
                    Call SetPlayerDEF(index, GetPlayerDEF(index) + 1)
                    Call BattleMsg(index, "Tu defensa ha incrementado!", 15, 0)

                Case 2
                    Call SetPlayerMAGI(index, GetPlayerMAGI(index) + 1)
                    Call BattleMsg(index, "Tu magia ha incrementado!", 15, 0)

                Case 3
                    Call SetPlayerSPEED(index, GetPlayerSPEED(index) + 1)
                    Call BattleMsg(index, "Tu velocidad ha incrementado!", 15, 0)
            End Select

            Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) - 1)
        End If
    Else
        Call BattleMsg(index, "No tienes puntos de estado para entrenar!", BRIGHTRED, 0)
    End If

    Call SendHP(index)
    Call SendMP(index)
    Call SendSP(index)

    Player(index).Char(Player(index).CharNum).MAXHP = GetPlayerMaxHP(index)
    Player(index).Char(Player(index).CharNum).MAXMP = GetPlayerMaxMP(index)
    Player(index).Char(Player(index).CharNum).MAXSP = GetPlayerMaxSP(index)

    Call SendStats(index)

    Call SendDataTo(index, "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(index) & END_CHAR)
End Sub

Public Sub Packet_GetStats(ByVal index As Long, ByVal Name As String)
    Dim PlayerID As Long
    Dim BlockChance As Long
    Dim CritChance As Long

    PlayerID = FindPlayer(Name)

    If PlayerID > 0 Then
        Call PlayerMsg(index, "Cuenta: " & Trim$(Player(PlayerID).Login) & "; Nombre: " & GetPlayerName(PlayerID), BRIGHTGREEN)

        If GetPlayerAccess(index) > ADMIN_MONITER Then
            Call PlayerMsg(index, "Estado de " & GetPlayerName(PlayerID) & ":", BRIGHTGREEN)
            Call PlayerMsg(index, "Nivel: " & GetPlayerLevel(PlayerID) & "; EXP: " & GetPlayerExp(PlayerID) & "/" & GetPlayerNextLevel(PlayerID), BRIGHTGREEN)
            Call PlayerMsg(index, "PV: " & GetPlayerHP(PlayerID) & "/" & GetPlayerMaxHP(PlayerID) & "; PM: " & GetPlayerMP(PlayerID) & "/" & GetPlayerMaxMP(PlayerID) & "; PS: " & GetPlayerSP(PlayerID) & "/" & GetPlayerMaxSP(PlayerID), BRIGHTGREEN)
            Call PlayerMsg(index, "FRZ: " & GetPlayerSTR(PlayerID) & "; DEF: " & GetPlayerDEF(PlayerID) & "; MA: " & GetPlayerMAGI(PlayerID) & "; Vel: " & GetPlayerSPEED(PlayerID), BRIGHTGREEN)
            
            CritChance = Int(GetPlayerSTR(PlayerID) / 2) + Int(GetPlayerLevel(PlayerID) / 2)
            If CritChance < 0 Then
                CritChance = 0
            End If
            If CritChance > 100 Then
                CritChance = 100
            End If

            BlockChance = Int(GetPlayerDEF(PlayerID) / 2) + Int(GetPlayerLevel(PlayerID) / 2)
            If BlockChance < 0 Then
                BlockChance = 0
            End If
            If BlockChance > 100 Then
                BlockChance = 100
            End If

            Call PlayerMsg(index, "Oportunidad de critico: " & CritChance & "%; Oportunidad de bloqueo: " & BlockChance & "%", BRIGHTGREEN)
        End If
    Else
        Call PlayerMsg(index, Name & " no está conectado.", WHITE)
    End If
End Sub

Public Sub Packet_SetPlayerSprite(ByVal index As Long, ByVal Name As String, ByVal SpriteID As Long)
    Dim PlayerID As Long

    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    PlayerID = FindPlayer(Name)

    If PlayerID > 0 Then
        Call SetPlayerSprite(PlayerID, SpriteID)
        Call SendPlayerData(PlayerID)
    Else
        Call PlayerMsg(index, Name & " no está conectado.", WHITE)
    End If
End Sub

Public Sub Packet_RequestNewMap(ByVal index As Long, ByVal Dir As Long)
    Dim x As Integer
    Dim y As Integer
    
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Call HackingAttempt(index, "Invalid Direction")
        Exit Sub
    End If
    
    y = GetPlayerY(index)
    x = GetPlayerX(index)
   
    Select Case Dir
        Case DIR_UP
            y = y - 1
        Case DIR_DOWN
            y = y + 1
        Case DIR_LEFT
            x = x - 1
        Case DIR_RIGHT
            x = x + 1
   End Select

    Call PlayerMove(index, Dir, 1, x, y)
    Call SendPlayerNewXY(index)
End Sub

Public Sub Packet_WarpMeTo(ByVal index As Long, ByVal Name As String)
    Dim PlayerID As Long

    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If
    
    PlayerID = FindPlayer(Name)

    If PlayerID > 0 Then
        Call PlayerWarp(index, GetPlayerMap(PlayerID), GetPlayerX(PlayerID), GetPlayerY(PlayerID))
    Else
        Call PlayerMsg(index, Name & " no está conectado.", WHITE)
    End If
End Sub

Public Sub Packet_WarpToMe(ByVal index As Long, ByVal Name As String)
    Dim PlayerID As Long

    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If
    
    PlayerID = FindPlayer(Name)

    If PlayerID > 0 Then
        Call PlayerWarp(PlayerID, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    Else
        Call PlayerMsg(index, Name & " no está conectado.", WHITE)
    End If
End Sub


Public Sub Packet_MapData(ByVal index As Long, ByRef MapData() As String)
    Dim MapIndex As Long
    Dim mapnum As Long
    Dim MapRevision As Long
    Dim x As Long
    Dim y As Long
    Dim i As Long
    
    ' Check to see if the user is at least a mapper.
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If
            
    mapnum = GetPlayerMap(index)
            
    ' Get revision number before it clears
    MapRevision = Map(mapnum).Revision + 1
            
    MapIndex = 1

    Call ClearMap(mapnum)

    mapnum = Val(MapData(MapIndex))
    Map(mapnum).Name = MapData(MapIndex + 1)
    Map(mapnum).Revision = MapRevision
    Map(mapnum).Moral = Val(MapData(MapIndex + 3))
    Map(mapnum).Up = Val(MapData(MapIndex + 4))
    Map(mapnum).Down = Val(MapData(MapIndex + 5))
    Map(mapnum).Left = Val(MapData(MapIndex + 6))
    Map(mapnum).Right = Val(MapData(MapIndex + 7))
    Map(mapnum).music = MapData(MapIndex + 8)
    Map(mapnum).BootMap = Val(MapData(MapIndex + 9))
    Map(mapnum).BootX = Val(MapData(MapIndex + 10))
    Map(mapnum).BootY = Val(MapData(MapIndex + 11))
    Map(mapnum).Indoors = Val(MapData(MapIndex + 12))
    Map(mapnum).Weather = Val(MapData(MapIndex + 13))

    MapIndex = MapIndex + 14

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            Map(mapnum).Tile(x, y).Ground = Val(MapData(MapIndex))
            Map(mapnum).Tile(x, y).Mask = Val(MapData(MapIndex + 1))
            Map(mapnum).Tile(x, y).Anim = Val(MapData(MapIndex + 2))
            Map(mapnum).Tile(x, y).Mask2 = Val(MapData(MapIndex + 3))
            Map(mapnum).Tile(x, y).M2Anim = Val(MapData(MapIndex + 4))
            Map(mapnum).Tile(x, y).Fringe = Val(MapData(MapIndex + 5))
            Map(mapnum).Tile(x, y).FAnim = Val(MapData(MapIndex + 6))
            Map(mapnum).Tile(x, y).Fringe2 = Val(MapData(MapIndex + 7))
            Map(mapnum).Tile(x, y).F2Anim = Val(MapData(MapIndex + 8))
            Map(mapnum).Tile(x, y).Type = Val(MapData(MapIndex + 9))
            Map(mapnum).Tile(x, y).Data1 = Val(MapData(MapIndex + 10))
            Map(mapnum).Tile(x, y).Data2 = Val(MapData(MapIndex + 11))
            Map(mapnum).Tile(x, y).Data3 = Val(MapData(MapIndex + 12))
            Map(mapnum).Tile(x, y).String1 = MapData(MapIndex + 13)
            Map(mapnum).Tile(x, y).String2 = MapData(MapIndex + 14)
            Map(mapnum).Tile(x, y).String3 = MapData(MapIndex + 15)
            Map(mapnum).Tile(x, y).Light = Val(MapData(MapIndex + 16))
            Map(mapnum).Tile(x, y).GroundSet = Val(MapData(MapIndex + 17))
            Map(mapnum).Tile(x, y).MaskSet = Val(MapData(MapIndex + 18))
            Map(mapnum).Tile(x, y).AnimSet = Val(MapData(MapIndex + 19))
            Map(mapnum).Tile(x, y).Mask2Set = Val(MapData(MapIndex + 20))
            Map(mapnum).Tile(x, y).M2AnimSet = Val(MapData(MapIndex + 21))
            Map(mapnum).Tile(x, y).FringeSet = Val(MapData(MapIndex + 22))
            Map(mapnum).Tile(x, y).FAnimSet = Val(MapData(MapIndex + 23))
            Map(mapnum).Tile(x, y).Fringe2Set = Val(MapData(MapIndex + 24))
            Map(mapnum).Tile(x, y).F2AnimSet = Val(MapData(MapIndex + 25))

            MapIndex = MapIndex + 26
        Next x
    Next y

    For x = 1 To MAX_MAP_NPCS
        Map(mapnum).NPC(x) = Val(MapData(MapIndex))
        Map(mapnum).SpawnX(x) = Val(MapData(MapIndex + 1))
        Map(mapnum).SpawnY(x) = Val(MapData(MapIndex + 2))
        MapIndex = MapIndex + 3
        Call ClearMapNpc(x, mapnum)
    Next x

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).x, MapItem(GetPlayerMap(index), i).y)
        Call ClearMapItem(i, GetPlayerMap(index))
    Next i

    ' Save the map
    Call SaveMap(mapnum)
            
    ' Mapper is on the map
    PlayersOnMap(mapnum) = YES

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(index))

    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(index))
    Next i

    ' Refresh map for everyone online
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapnum Then
                Call SendDataTo(i, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & Map(GetPlayerMap(i)).Revision & END_CHAR)
            End If
        End If
    Next i
End Sub

Public Sub Packet_NeedMap(ByVal index As Long, ByVal NeedMap As String)
    Dim i As Long

    NeedMap = UCase$(NeedMap)

    If NeedMap = "YES" Then
        Call SendMap(index, GetPlayerMap(index))
    End If

    Call SendMapItemsTo(index, GetPlayerMap(index))
    Call SendMapNpcsTo(index, GetPlayerMap(index))
    Call SendJoinMap(index)
    Call SendDataTo(index, "MAPDONE" & END_CHAR)

    Player(index).GettingMap = NO

    Call SendPlayerData(index)

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendHP(i)
            Call SendIndexWornEquipment(i)
            Call SendWornEquipment(i)
            Call SendPlayerColor(i)
        End If
    Next i
End Sub

Public Sub Packet_MapGetItem(ByVal index As Long)
    Call PlayerMapGetItem(index)
End Sub

Public Sub Packet_MapDropItem(ByVal index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    If InvNum < 1 Or InvNum > MAX_INV Then
        Call HackingAttempt(index, "Invalid InvNum")
        Exit Sub
    End If

    ' Prevent hacking
    If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
        If Amount <= 0 Then
            Call PlayerMsg(index, "Debes soltar almenos la cantidad de 1 del objeto!", BRIGHTRED)
            Exit Sub
        End If

        If Amount > GetPlayerInvItemValue(index, InvNum) Then
            Call PlayerMsg(index, "No tienes esa cantidad para soltar!", BRIGHTRED)
            Exit Sub
        End If
    End If

    ' Prevent hacking
    If Item(GetPlayerInvItemNum(index, InvNum)).Type <> ITEM_TYPE_CURRENCY Then
        If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
            If Amount > GetPlayerInvItemValue(index, InvNum) Then
                Call HackingAttempt(index, "Item amount modification")
                Exit Sub
            End If
        End If
    End If

    Call PlayerMapDropItem(index, InvNum, Amount)

    Call SendStats(index)
    Call SendHP(index)
    Call SendMP(index)
    Call SendSP(index)
End Sub

Public Sub Packet_MapRespawn(ByVal index As Long)
    Dim i As Long

    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    ' Clear out all of the floor items.
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).x, MapItem(GetPlayerMap(index), i).y)
        Call ClearMapItem(i, GetPlayerMap(index))
    Next i

    ' Respawn all of the floor items.
    Call SpawnMapItems(GetPlayerMap(index))

    ' Respawn NPCS
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(index))
    Next i

    Call PlayerMsg(index, "Mapa respawneado.", Blue)
End Sub

Public Sub Packet_KickPlayer(ByVal index As Long, ByVal Name As String)
    Dim PlayerIndex As Long

    If GetPlayerAccess(index) < 1 Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex > 0 Then
        If PlayerIndex <> index Then
            If GetPlayerAccess(PlayerIndex) <= GetPlayerAccess(index) Then
                Call GlobalMsg(GetPlayerName(PlayerIndex) & " ha sido expulsado de " & GAME_NAME & " por " & GetPlayerName(index) & "!", WHITE)
                Call AddLog(GetPlayerName(index) & " ha expulsado a " & GetPlayerName(PlayerIndex) & ".", ADMIN_LOG)
                Call AlertMsg(PlayerIndex, "Has sido expulsado por " & GetPlayerName(index) & "!")
            Else
                Call PlayerMsg(index, "No puedes expulsar un administrador!", WHITE)
            End If
        Else
            Call PlayerMsg(index, "No puedes expulsarte a ti mismo!", WHITE)
        End If
    Else
        Call PlayerMsg(index, "El jugador no está conectado.", WHITE)
    End If
End Sub

Public Sub Packet_BanList(ByVal index As Long)
    Dim FileID As Integer
    Dim PlayerName As String

    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If
            
    If Not FileExists("BanList.txt") Then
        Call PlayerMsg(index, "La lista de bans no ha sido encontrada!", BRIGHTRED)
        Exit Sub
    End If

    FileID = FreeFile

    Open App.Path & "\BanList.txt" For Input As #FileID
    Do While Not EOF(FileID)
        Line Input #FileID, PlayerName
        Call PlayerMsg(index, PlayerName, WHITE)
    Loop
    Close #FileID
End Sub

Public Sub Packet_BanListDestroy(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If FileExists("BanList.txt") Then
        Call Kill(App.Path & "\BanList.txt")
    End If

    Call PlayerMsg(index, "Lista de bans destruida.", WHITE)
End Sub

Public Sub Packet_BanPlayer(ByVal index As Long, ByVal Name As String)
    Dim PlayerIndex As Long

    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex > 0 Then
        If PlayerIndex <> index Then
            If GetPlayerAccess(PlayerIndex) <= GetPlayerAccess(index) Then
                Call BanIndex(PlayerIndex, index)
            Else
                Call PlayerMsg(index, "No puedes banear un administrador!", WHITE)
            End If
        Else
            Call PlayerMsg(index, "No puedes banearte a ti mismo!", WHITE)
        End If
    Else
        Call PlayerMsg(index, "El jugador no esta conectado.", WHITE)
    End If
End Sub

Public Sub Packet_RequestEditMap(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(index, "EDITMAP" & END_CHAR)
End Sub

Public Sub Packet_RequestEditItem(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(index, "ITEMEDITOR" & END_CHAR)
End Sub

Public Sub Packet_EditItem(ByVal index As Long, ByVal ItemNum As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Call HackingAttempt(index, "Invalid Item Index")
        Exit Sub
    End If

    Call SendEditItemTo(index, ItemNum)

    Call AddLog(GetPlayerName(index) & " edita el objeto #" & ItemNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveItem(ByVal index As Long, ByRef ItemData() As String)
    Dim ItemNum As Long

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    ItemNum = Val(ItemData(1))

    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Call HackingAttempt(index, "Invalid Item Index")
        Exit Sub
    End If

    Item(ItemNum).Name = ItemData(2)
    Item(ItemNum).Pic = Val(ItemData(3))
    Item(ItemNum).Type = Val(ItemData(4))
    Item(ItemNum).Data1 = Val(ItemData(5))
    Item(ItemNum).Data2 = Val(ItemData(6))
    Item(ItemNum).Data3 = Val(ItemData(7))
    Item(ItemNum).StrReq = Val(ItemData(8))
    Item(ItemNum).DefReq = Val(ItemData(9))
    Item(ItemNum).SpeedReq = Val(ItemData(10))
    Item(ItemNum).MagicReq = Val(ItemData(11))
    Item(ItemNum).ClassReq = Val(ItemData(12))
    Item(ItemNum).AccessReq = Val(ItemData(13))

    Item(ItemNum).addHP = Val(ItemData(14))
    Item(ItemNum).addMP = Val(ItemData(15))
    Item(ItemNum).addSP = Val(ItemData(16))
    Item(ItemNum).AddStr = Val(ItemData(17))
    Item(ItemNum).AddDef = Val(ItemData(18))
    Item(ItemNum).AddMagi = Val(ItemData(19))
    Item(ItemNum).AddSpeed = Val(ItemData(20))
    Item(ItemNum).AddEXP = Val(ItemData(21))
    Item(ItemNum).Desc = ItemData(22)
    Item(ItemNum).AttackSpeed = Val(ItemData(23))
    Item(ItemNum).Price = Val(ItemData(24))
    Item(ItemNum).Stackable = Val(ItemData(25))
    Item(ItemNum).Bound = Val(ItemData(26))
    
    Item(ItemNum).PetSprite = Val(ItemData(27))

    Call SendUpdateItemToAll(ItemNum)
    Call SaveItem(ItemNum)

    Call AddLog(GetPlayerName(index) & " guarda el objeto #" & ItemNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_EnableDayNight(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If Not TimeDisable Then
        Gamespeed = 0
        frmServer.GameTimeSpeed.Text = 0
        TimeDisable = True
        frmServer.Timer1.Enabled = False
        frmServer.Command69.Caption = "Activar Tiempo"
    Else
        Gamespeed = 1
        frmServer.GameTimeSpeed.Text = 1
        TimeDisable = False
        frmServer.Timer1.Enabled = True
        frmServer.Command69.Caption = "Desactivar Tiempo"
    End If
End Sub

Public Sub Packet_DayNight(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If Hours > 12 Then
        Hours = Hours - 12
    Else
        Hours = Hours + 12
    End If
End Sub

Public Sub Packet_RequestEditNPC(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(index, "NPCEDITOR" & END_CHAR)
End Sub

Public Sub Packet_EditNPC(ByVal index As Long, ByVal npcnum As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If npcnum < 0 Or npcnum > MAX_NPCS Then
        Call HackingAttempt(index, "Invalid NPC Index")
        Exit Sub
    End If

    Call SendEditNpcTo(index, npcnum)

    Call AddLog(GetPlayerName(index) & " edita el npc #" & npcnum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveNPC(ByVal index As Long, ByRef NPCData() As String)
    Dim npcnum As Long
    Dim NPCIndex As Long
    Dim i As Long

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    npcnum = Val(NPCData(1))

    If npcnum < 0 Or npcnum > MAX_NPCS Then
        Call HackingAttempt(index, "Invalid NPC Index")
        Exit Sub
    End If

    NPC(npcnum).Name = NPCData(2)
    NPC(npcnum).AttackSay = NPCData(3)
    NPC(npcnum).SPRITE = Val(NPCData(4))
    NPC(npcnum).SpawnSecs = Val(NPCData(5))
    NPC(npcnum).Behavior = Val(NPCData(6))
    NPC(npcnum).Range = Val(NPCData(7))
    NPC(npcnum).STR = Val(NPCData(8))
    NPC(npcnum).DEF = Val(NPCData(9))
    NPC(npcnum).Speed = Val(NPCData(10))
    NPC(npcnum).Magi = Val(NPCData(11))
    NPC(npcnum).Big = Val(NPCData(12))
    NPC(npcnum).MAXHP = Val(NPCData(13))
    NPC(npcnum).EXP = Val(NPCData(14))
    NPC(npcnum).SpawnTime = Val(NPCData(15))
    NPC(npcnum).Element = Val(NPCData(16))
    NPC(npcnum).SPRITESIZE = Val(NPCData(17))
    NPC(npcnum).Quest = Val(NPCData(18))

    NPCIndex = 19

    For i = 1 To MAX_NPC_DROPS
        NPC(npcnum).ItemNPC(i).chance = Val(NPCData(NPCIndex))
        NPC(npcnum).ItemNPC(i).ItemNum = Val(NPCData(NPCIndex + 1))
        NPC(npcnum).ItemNPC(i).ItemValue = Val(NPCData(NPCIndex + 2))
        NPCIndex = NPCIndex + 3
    Next i
    
    NPC(npcnum).standstill = CBool(NPCData(49))

    Call SendUpdateNpcToAll(npcnum)
    Call SaveNpc(npcnum)

    Call AddLog(GetPlayerName(index) & " guarda el npc #" & npcnum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_RequestEditShop(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(index, "SHOPEDITOR" & END_CHAR)
End Sub

Public Sub Packet_EditShop(ByVal index As Long, ByVal ShopNum As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
        Call HackingAttempt(index, "Invalid Shop Index")
        Exit Sub
    End If

    Call SendEditShopTo(index, ShopNum)

    Call AddLog(GetPlayerName(index) & " edita la tienda #" & ShopNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveShop(ByVal index As Long, ByRef ShopData() As String)
    Dim ShopNum As Long
    Dim ShopIndex As Long
    Dim i As Long

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    ShopNum = Val(ShopData(1))

    If ShopNum < 1 Or ShopNum > MAX_SHOPS Then
        Call HackingAttempt(index, "Invalid Shop Index")
        Exit Sub
    End If

    Shop(ShopNum).Name = ShopData(2)
    Shop(ShopNum).FixesItems = Val(ShopData(3))
    Shop(ShopNum).BuysItems = Val(ShopData(4))
    Shop(ShopNum).ShowInfo = Val(ShopData(5))
    Shop(ShopNum).CurrencyItem = Val(ShopData(6))

    ShopIndex = 7

    For i = 1 To MAX_SHOP_ITEMS
        Shop(ShopNum).ShopItem(i).ItemNum = Val(ShopData(ShopIndex))
        Shop(ShopNum).ShopItem(i).Amount = Val(ShopData(ShopIndex + 1))
        Shop(ShopNum).ShopItem(i).Price = Val(ShopData(ShopIndex + 2))
        ShopIndex = ShopIndex + 3
    Next i

    Call SendUpdateShopToAll(ShopNum)
    Call SaveShop(ShopNum)

    Call AddLog(GetPlayerName(index) & " guarda la tienda #" & ShopNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_RequestEditSpell(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(index, "SPELLEDITOR" & END_CHAR)
End Sub

Public Sub Packet_EditSpell(ByVal index As Long, ByVal spellnum As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If spellnum < 0 Or spellnum > MAX_SPELLS Then
        Call HackingAttempt(index, "Invalid Spell Index")
        Exit Sub
    End If

    Call SendEditSpellTo(index, spellnum)

    Call AddLog(GetPlayerName(index) & " edita el hechizo #" & spellnum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveSpell(ByVal index As Long, ByRef SpellData() As String)
    Dim spellnum As Long
    
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    spellnum = Val(SpellData(1))

    If spellnum < 1 Or spellnum > MAX_SPELLS Then
        Call HackingAttempt(index, "Invalid Spell Index")
        Exit Sub
    End If

    Spell(spellnum).Name = SpellData(2)
    Spell(spellnum).ClassReq = Val(SpellData(3))
    Spell(spellnum).LevelReq = Val(SpellData(4))
    Spell(spellnum).Type = Val(SpellData(5))
    Spell(spellnum).Data1 = Val(SpellData(6))
    Spell(spellnum).Data2 = Val(SpellData(7))
    Spell(spellnum).Data3 = Val(SpellData(8))
    Spell(spellnum).MPCost = Val(SpellData(9))
    Spell(spellnum).Sound = Val(SpellData(10))
    Spell(spellnum).Range = Val(SpellData(11))
    Spell(spellnum).SpellAnim = Val(SpellData(12))
    Spell(spellnum).SpellTime = Val(SpellData(13))
    Spell(spellnum).SpellDone = Val(SpellData(14))
    Spell(spellnum).AE = Val(SpellData(15))
    Spell(spellnum).Big = Val(SpellData(16))
    Spell(spellnum).Element = Val(SpellData(17))
    Spell(spellnum).TimeToCast = Val(SpellData(18))
    Spell(spellnum).CastTimer = Val(SpellData(19))

    Call SendUpdateSpellToAll(spellnum)
    Call SaveSpell(spellnum)

    Call AddLog(GetPlayerName(index) & " guarda el hechizo #" & spellnum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_ForgetSpell(ByVal index As Long, ByVal spellnum As Long)
    If spellnum < 1 Or spellnum > MAX_PLAYER_SPELLS Then
        Call HackingAttempt(index, "Invalid Spell Slot")
        Exit Sub
    End If

    With Player(index).Char(Player(index).CharNum)
        If .Spell(spellnum) = 0 Then
            Call PlayerMsg(index, "No tienes seleccionado un hechizo.", Red)
        Else
            Call PlayerMsg(index, "Has olvidado el hechizo " & Trim$(Spell(.Spell(spellnum)).Name) & ".", Green)

            .Spell(spellnum) = 0

            Call SendSpells(index)
        End If
    End With
End Sub

Public Sub Packet_SetAccess(ByVal index As Long, ByVal Name As String, ByVal AccessLvl As Long)
    Dim PlayerIndex As Long
    
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Call HackingAttempt(index, "Invalid Access")
        Exit Sub
    End If
    
    If AccessLvl < 0 Or AccessLvl > 5 Then
        Call PlayerMsg(index, "Has introducido un privilegio que no existe.", BRIGHTRED)
        Exit Sub
    End If

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex > 0 Then
        If GetPlayerName(index) <> GetPlayerName(PlayerIndex) Then
            If GetPlayerAccess(index) > GetPlayerAccess(PlayerIndex) Then
                Call SetPlayerAccess(PlayerIndex, AccessLvl)
                Call SendPlayerData(PlayerIndex)
    
                If GetPlayerAccess(PlayerIndex) = 0 Then
                    Call GlobalMsg(GetPlayerName(PlayerIndex) & " ha sido bendecido con poderes administrativos.", BRIGHTBLUE)
                End If
    
                Call AddLog(GetPlayerName(index) & " ha modificado el acceso de " & GetPlayerName(PlayerIndex) & ".", ADMIN_LOG)
            Else
                Call PlayerMsg(index, "Tu nivel de acceso es menor que el de " & GetPlayerName(PlayerIndex) & ".", Red)
            End If
        Else
            Call PlayerMsg(index, "No puedes cambiar tu acceso", Red)
        End If
    Else
        Call PlayerMsg(index, "El jugador no esta conectado.", WHITE)
    End If
End Sub

Public Sub Packet_WhoIsOnline(ByVal index As Long)
    Call SendWhosOnline(index)
End Sub

Public Sub Packet_OnlineList(ByVal index As Long)
    Call SendOnlineList
End Sub

Public Sub Packet_SetMOTD(ByVal index As Long, ByVal MOTD As String)
    If GetPlayerAccess(index) < ADMIN_MAPPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call PutVar(App.Path & "\MOTD.ini", "MOTD", "Msg", MOTD)
            
    If scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "ChangeMOTD"
    End If
            
    Call GlobalMsg("MOTD cambiado a: " & MOTD, BRIGHTCYAN)

    Call AddLog(GetPlayerName(index) & " cambia el MOTD a: " & MOTD, ADMIN_LOG)
End Sub

Public Sub Packet_BuyItem(ByVal index As Long, ByVal ShopIndex As Long, ByVal ItemIndex As Long)
    Dim InvItem As Long

    If ShopIndex < 1 Or ShopIndex > MAX_SHOPS Then
        Call HackingAttempt(index, "Invalid Shop Index")
        Exit Sub
    End If
    
    If ItemIndex < 1 Or ItemIndex > MAX_SHOP_ITEMS Then
        Call HackingAttempt(index, "Invalid Shop Item")
        Exit Sub
    End If

    ' Check to see if player's inventory is full.
    InvItem = FindOpenInvSlot(index, Shop(ShopIndex).ShopItem(ItemIndex).ItemNum)
    If InvItem = 0 Then
        Call PlayerMsg(index, "Tu inventario esta al maximo de su capacidad!", BRIGHTRED)
        Exit Sub
    End If

    ' Check to see if they have enough currency.
    If HasItem(index, Shop(ShopIndex).CurrencyItem) >= Shop(ShopIndex).ShopItem(ItemIndex).Price Then
        Call TakeItem(index, Shop(ShopIndex).CurrencyItem, Shop(ShopIndex).ShopItem(ItemIndex).Price)
        Call GiveItem(index, Shop(ShopIndex).ShopItem(ItemIndex).ItemNum, Shop(ShopIndex).ShopItem(ItemIndex).Amount)

        Call PlayerMsg(index, "Compras el objeto.", YELLOW)
    Else
        Call PlayerMsg(index, "No puedes permitirte eso!", Red)
    End If
End Sub

Public Sub Packet_SellItem(ByVal index As Long, ByVal ShopNum As Long, ByVal ItemNum As Long, ByVal ItemSlot As Long, ByVal ItemAmt As Long)
    If ItemIsEquipped(index, ItemNum) Then
        Call PlayerMsg(index, "No puedes vender objetos equipados.", Red)
        Exit Sub
    End If

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        Call PlayerMsg(index, "No puedes vender dinero!.", Red)
        Exit Sub
    End If

    If Item(ItemNum).Stackable = YES Then
        If ItemAmt > GetPlayerInvItemValue(index, ItemSlot) Then
            Call PlayerMsg(index, "No tienes suficientes objetos para vender por esta cantidad!", Red)
            Exit Sub
        End If
    End If

    If Item(ItemNum).Price > 0 Then
        Call TakeItem(index, ItemNum, ItemAmt)
        Call GiveItem(index, Shop(ShopNum).CurrencyItem, Item(ItemNum).Price * ItemAmt)
        Call PlayerMsg(index, "El vendedor te da " & Item(ItemNum).Price * ItemAmt & " " & Trim$(Item(Shop(ShopNum).CurrencyItem).Name) & ".", YELLOW)
    Else
        Call PlayerMsg(index, "Este objeto no puede ser vendido.", Red)
    End If
End Sub

Public Sub Packet_FixItem(ByVal index As Long, ByVal ShopNum As Long, ByVal InvNum As Long)
    Dim ItemNum As Long
    Dim DurNeeded As Long
    Dim GoldNeeded As Long
    Dim i As Long

    If Item(GetPlayerInvItemNum(index, InvNum)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(index, InvNum)).Type > ITEM_TYPE_NECKLACE Then
        Call PlayerMsg(index, "Este objeto no necesita ser reparado.", BRIGHTRED)
        Exit Sub
    End If

    If FindOpenInvSlot(index, GetPlayerInvItemNum(index, InvNum)) = 0 Then
        Call PlayerMsg(index, "No tienes espacio en tu inventario!", BRIGHTRED)
        Exit Sub
    End If

    ItemNum = GetPlayerInvItemNum(index, InvNum)

    i = Int(Item(GetPlayerInvItemNum(index, InvNum)).Data2 / 5)
    If i <= 0 Then
        i = 1
    End If

    DurNeeded = Item(ItemNum).Data1 - GetPlayerInvItemDur(index, InvNum)

    GoldNeeded = Int(DurNeeded * i / 2)
    If GoldNeeded <= 0 Then
        GoldNeeded = 1
    End If

    If DurNeeded = 0 Then
        Call PlayerMsg(index, "Este objeto esta en perfecta condición!", WHITE)
        Exit Sub
    End If

    If HasItem(index, Shop(ShopNum).CurrencyItem) >= i Then
        If HasItem(index, Shop(ShopNum).CurrencyItem) >= GoldNeeded Then
            Call TakeItem(index, Shop(ShopNum).CurrencyItem, GoldNeeded)
            Call SetPlayerInvItemDur(index, InvNum, Item(ItemNum).Data1)

            Call PlayerMsg(index, "EL objeto ha sido totalmente reparado por " & GoldNeeded & " de dinero!", BRIGHTBLUE)
        Else
            DurNeeded = (HasItem(index, Shop(ShopNum).CurrencyItem) / i)
            GoldNeeded = Int(DurNeeded * i / 2)

            If GoldNeeded <= 0 Then
                GoldNeeded = 1
            End If

            Call TakeItem(index, Shop(ShopNum).CurrencyItem, GoldNeeded)
            Call SetPlayerInvItemDur(index, InvNum, GetPlayerInvItemDur(index, InvNum) + DurNeeded)

            Call PlayerMsg(index, "El objeto ha sido parcialmente reparado por " & GoldNeeded & " de dinero!", BRIGHTBLUE)
        End If
    Else
        Call PlayerMsg(index, "No tienes suficiente dinero para reparar este objeto!", BRIGHTRED)
    End If
End Sub

Public Sub Packet_Search(ByVal index As Long, ByVal x As Long, ByVal y As Long)
    Dim i As Long

    If x < 0 Or x > MAX_MAPX Then
        Exit Sub
    End If

    If y < 0 Or y > MAX_MAPY Then
        Exit Sub
    End If

    ' Check for a player
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(index) = GetPlayerMap(i) Then
                If GetPlayerX(i) = x Then
                    If GetPlayerY(i) = y Then
                        If GetPlayerLevel(i) >= GetPlayerLevel(index) + 5 Then
                            Call PlayerMsg(index, "No tienes nada que hacer contra el.", BRIGHTRED)
                        Else
                            If GetPlayerLevel(i) > GetPlayerLevel(index) Then
                                Call PlayerMsg(index, "El tiene cierta ventaja sobre ti.", YELLOW)
                            Else
                                If GetPlayerLevel(i) = GetPlayerLevel(index) Then
                                    Call PlayerMsg(index, "El es duro de pelar, pero tienes posibilidad.", WHITE)
                                Else
                                    If GetPlayerLevel(index) >= GetPlayerLevel(i) + 5 Then
                                        Call PlayerMsg(index, "Puedes vencer a este jugador.", BRIGHTBLUE)
                                    Else
                                        If GetPlayerLevel(index) > GetPlayerLevel(i) Then
                                            Call PlayerMsg(index, "Tienes ventaja absoluta sobre este jugador.", YELLOW)
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        ' Change the target.
                        Player(index).Target = i
                        Player(index).TargetType = TARGET_TYPE_PLAYER
                        
                        If scripting = 1 Then
                        MyScript.ExecuteStatement "Scripts\Main.txt", "OnClickPlayer " & index
                        End If

                        Call PlayerMsg(index, "Tu objetivo es ahora " & GetPlayerName(i) & ".", YELLOW)

                        Exit Sub
                    End If
                End If
            End If

        End If
    Next i

    ' Check for an NPC.
    For i = 1 To MAX_MAP_NPCS
        If MapNPC(GetPlayerMap(index), i).num > 0 Then
            If MapNPC(GetPlayerMap(index), i).x = x Then
                If MapNPC(GetPlayerMap(index), i).y = y Then
                    Player(index).TargetNPC = i
                    Player(index).TargetType = TARGET_TYPE_NPC

                    Call PlayerMsg(index, "Tu objetivo es " & Trim$(NPC(MapNPC(GetPlayerMap(index), i).num).Name) & ".", YELLOW)

                    Exit Sub
                End If
            End If
        End If
    Next i

    ' Check for an item on the ground.
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(GetPlayerMap(index), i).num > 0 Then
            If MapItem(GetPlayerMap(index), i).x = x Then
                If MapItem(GetPlayerMap(index), i).y = y Then
                    Call PlayerMsg(index, "Ves " & Trim$(Item(MapItem(GetPlayerMap(index), i).num).Name) & ".", YELLOW)
                    Exit Sub
                End If
            End If
        End If
    Next i

    ' Check for an OnClick tile.
    If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_ONCLICK Then
        If scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnClick " & index & "," & Map(GetPlayerMap(index)).Tile(x, y).Data1
        End If
    End If
End Sub

Public Sub Packet_PlayerChat(ByVal index As Long, ByVal Name As String)
    Dim PlayerIndex As Long

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, Name & " no está conectado.", WHITE)
        Exit Sub
    End If

    If PlayerIndex = index Then
        Call PlayerMsg(index, "No puedes chatear contigo mismo.", PINK)
        Exit Sub
    End If

    If Player(index).InChat = 1 Then
        Call PlayerMsg(index, "Ya estas en un chat con otro jugador!", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).InChat = 1 Then
        Call PlayerMsg(index, Name & " esta chateando con otro jugador!", PINK)
        Exit Sub
    End If

    Call PlayerMsg(index, "La petición de chat ha sido enviada a " & GetPlayerName(PlayerIndex) & ".", PINK)
    ' ComandoAE
    Call PlayerMsg(PlayerIndex, GetPlayerName(index) & " quiero chatear contigo en privado. Escribe /chat para aceptar, o /rechazarchat para cancelar.", PINK)

    Player(index).ChatPlayer = PlayerIndex
    Player(PlayerIndex).ChatPlayer = index
End Sub

Public Sub Packet_AcceptChat(ByVal index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(index).ChatPlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, "No tienes ninguna petición para chatear con nadie", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).ChatPlayer <> index Then
        Call PlayerMsg(index, "El chat fallo.", PINK)
        Exit Sub
    End If

    Call SendDataTo(index, "PPCHATTING" & SEP_CHAR & PlayerIndex & END_CHAR)
    Call SendDataTo(PlayerIndex, "PPCHATTING" & SEP_CHAR & index & END_CHAR)
End Sub

Public Sub Packet_DenyChat(ByVal index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(index).ChatPlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, "No tienes ninguna petición para chatear con nadie.", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).ChatPlayer <> index Then
        Call PlayerMsg(index, "El chat fallo.", PINK)
        Exit Sub
    End If

    Call PlayerMsg(index, "Petición de chat rechazada.", PINK)
    Call PlayerMsg(PlayerIndex, GetPlayerName(index) & " rechazo tu petición de chat.", PINK)

    Player(index).ChatPlayer = 0
    Player(index).InChat = 0

    Player(PlayerIndex).ChatPlayer = 0
    Player(PlayerIndex).InChat = 0
End Sub

Public Sub Packet_QuitChat(ByVal index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(index).ChatPlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, "No tienes ninguna petición para chatear con nadie.", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).ChatPlayer <> index Then
        Call PlayerMsg(index, "El chat fallo.", PINK)
        Exit Sub
    End If

    Call SendDataTo(index, "qchat" & END_CHAR)
    Call SendDataTo(PlayerIndex, "qchat" & END_CHAR)

    Player(index).ChatPlayer = 0
    Player(index).InChat = 0

    Player(PlayerIndex).ChatPlayer = 0
    Player(PlayerIndex).InChat = 0
End Sub

Public Sub Packet_SendChat(ByVal index As Long, ByVal Message As String)
    Dim PlayerIndex As Long

    PlayerIndex = Player(index).ChatPlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, "No tienes ninguna petición para chatear con nadie.", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).ChatPlayer <> index Then
        Call PlayerMsg(index, "El chat fallo.", PINK)
        Exit Sub
    End If

    Call SendDataTo(PlayerIndex, "sendchat" & SEP_CHAR & Message & SEP_CHAR & index & END_CHAR)
End Sub

Public Sub Packet_PrepareTrade(ByVal index As Long, ByVal Name As String)
    Dim PlayerIndex As Long

    PlayerIndex = FindPlayer(Name)

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, Name & " no está conectado.", WHITE)
        Exit Sub
    End If

    If PlayerIndex = index Then
        Call PlayerMsg(index, "No puedes comerciar contigo mismo!", PINK)
        Exit Sub
    End If

    If GetPlayerMap(index) <> GetPlayerMap(PlayerIndex) Then
        Call PlayerMsg(index, "Necesitas estar en el mismo mapa para comerciar con " & GetPlayerName(PlayerIndex) & "!", PINK)
        Exit Sub
    End If

    If Player(index).InTrade Then
        Call PlayerMsg(index, "Ya estas comerciando con otro jugador!", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).InTrade Then
        Call PlayerMsg(index, Name & " esta actualmente comerciando con otro jugador!", PINK)
        Exit Sub
    End If

    Call PlayerMsg(index, "La petición de comerciar ha sido enviada a " & GetPlayerName(PlayerIndex) & ".", PINK)
    'ComandoAE
    Call PlayerMsg(PlayerIndex, GetPlayerName(index) & " quiere comerciar contigo. Escribe /aceptar para aceptar, o /rechazar para rechazar.", PINK)

    Player(index).TradePlayer = PlayerIndex
    Player(PlayerIndex).TradePlayer = index
End Sub

Public Sub Packet_AcceptTrade(ByVal index As Long)
    Dim PlayerIndex As Long
    Dim i As Long

    PlayerIndex = Player(index).TradePlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, "No tienes ninguna petición para comerciar con nadie.", PINK)
        Exit Sub
    End If

    If Player(PlayerIndex).TradePlayer <> index Then
        Call PlayerMsg(index, "El comercio fallo.", PINK)
        Exit Sub
    End If

    If GetPlayerMap(index) <> GetPlayerMap(PlayerIndex) Then
        Call PlayerMsg(index, "Necesitas estar en el mismo mapa para comerciar con " & GetPlayerName(PlayerIndex) & "!", PINK)
        Exit Sub
    End If

    Call PlayerMsg(index, "Estas comerciando con " & GetPlayerName(PlayerIndex) & "!", PINK)
    Call PlayerMsg(PlayerIndex, GetPlayerName(index) & " acepto tu petición de comercio!", PINK)

    Call SendDataTo(index, "PPTRADING" & END_CHAR)
    Call SendDataTo(PlayerIndex, "PPTRADING" & END_CHAR)

    For i = 1 To MAX_PLAYER_TRADES
        Player(index).Trading(i).InvNum = 0
        Player(index).Trading(i).InvName = vbNullString

        Player(PlayerIndex).Trading(i).InvNum = 0
        Player(PlayerIndex).Trading(i).InvName = vbNullString
    Next i

    Player(index).InTrade = True
    Player(index).TradeItemMax = 0
    Player(index).TradeItemMax2 = 0

    Player(PlayerIndex).InTrade = True
    Player(PlayerIndex).TradeItemMax = 0
    Player(PlayerIndex).TradeItemMax2 = 0
End Sub

Public Sub Packet_QuitTrade(ByVal index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(index).TradePlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, "No tienes ninguna petición para comerciar con nadie.", PINK)
        Exit Sub
    End If

    Call PlayerMsg(index, "El comercio se cerro.", PINK)
    Call PlayerMsg(PlayerIndex, GetPlayerName(index) & " dejó de comerciar contigo!", PINK)

    Player(index).TradeOk = 0
    Player(index).TradePlayer = 0
    Player(index).InTrade = False

    Player(PlayerIndex).TradeOk = 0
    Player(PlayerIndex).TradePlayer = 0
    Player(PlayerIndex).InTrade = False

    Call SendDataTo(index, "qtrade" & END_CHAR)
    Call SendDataTo(PlayerIndex, "qtrade" & END_CHAR)
End Sub

Public Sub Packet_DenyTrade(ByVal index As Long)
    Dim PlayerIndex As Long

    PlayerIndex = Player(index).TradePlayer

    If PlayerIndex = 0 Then
        Call PlayerMsg(index, "No tienes ninguna petición para comerciar con nadie.", PINK)
        Exit Sub
    End If

    Call PlayerMsg(index, "Petición de comercio rechazada.", PINK)
    Call PlayerMsg(PlayerIndex, GetPlayerName(index) & " rechazo tu petición para comerciar.", PINK)

    Player(index).TradePlayer = 0
    Player(index).InTrade = False

    Player(PlayerIndex).TradePlayer = 0
    Player(PlayerIndex).InTrade = False
End Sub

Public Sub Packet_UpdateTradeInventory(ByVal index As Long, ByVal TradeIndex As Long, ByVal ItemNum As Long, ByVal ItemName As String)
    Player(index).Trading(TradeIndex).InvNum = ItemNum
    Player(index).Trading(TradeIndex).InvName = Trim$(ItemName)

    If Player(index).Trading(TradeIndex).InvNum = 0 Then
        Player(index).TradeItemMax = Player(index).TradeItemMax - 1
        Player(index).TradeOk = 0
        Player(TradeIndex).TradeOk = 0

        Call SendDataTo(index, "trading" & SEP_CHAR & 0 & END_CHAR)
        Call SendDataTo(TradeIndex, "trading" & SEP_CHAR & 0 & END_CHAR)
    Else
        Player(index).TradeItemMax = Player(index).TradeItemMax + 1
    End If

    Call SendDataTo(Player(index).TradePlayer, "updatetradeitem" & SEP_CHAR & TradeIndex & SEP_CHAR & Player(index).Trading(TradeIndex).InvNum & SEP_CHAR & Player(index).Trading(TradeIndex).InvName & END_CHAR)
End Sub

Public Sub Packet_SwapItems(ByVal index As Long)
    Dim TradeIndex As Long
    Dim i As Long
    Dim x As Long

    TradeIndex = Player(index).TradePlayer

    If Player(index).TradeOk = 0 Then
        Player(index).TradeOk = 1
        Call SendDataTo(TradeIndex, "trading" & SEP_CHAR & 1 & END_CHAR)
    ElseIf Player(index).TradeOk = 1 Then
        Player(index).TradeOk = 0
        Call SendDataTo(TradeIndex, "trading" & SEP_CHAR & 0 & END_CHAR)
    End If

    If Player(index).TradeOk = 1 Then
        If Player(TradeIndex).TradeOk = 1 Then
            Player(index).TradeItemMax2 = 0
            Player(TradeIndex).TradeItemMax2 = 0
    
            For i = 1 To MAX_INV
                If Player(index).TradeItemMax = Player(index).TradeItemMax2 Then
                    Exit For
                End If

                If GetPlayerInvItemNum(TradeIndex, i) < 1 Then
                    Player(index).TradeItemMax2 = Player(index).TradeItemMax2 + 1
                End If
            Next i
    
            For i = 1 To MAX_INV
                If Player(TradeIndex).TradeItemMax = Player(TradeIndex).TradeItemMax2 Then
                    Exit For
                End If

                If GetPlayerInvItemNum(index, i) < 1 Then
                    Player(TradeIndex).TradeItemMax2 = Player(TradeIndex).TradeItemMax2 + 1
                End If
            Next i
    
            If Player(index).TradeItemMax2 = Player(index).TradeItemMax And Player(TradeIndex).TradeItemMax2 = Player(TradeIndex).TradeItemMax Then
                For i = 1 To MAX_PLAYER_TRADES
                    For x = 1 To MAX_INV
                        If GetPlayerInvItemNum(TradeIndex, x) < 1 Then
                            If Player(index).Trading(i).InvNum > 0 Then
                                Call GiveItem(TradeIndex, GetPlayerInvItemNum(index, Player(index).Trading(i).InvNum), 1)
                                Call TakeItem(index, GetPlayerInvItemNum(index, Player(index).Trading(i).InvNum), 1)
                                Exit For
                            End If
                        End If
                    Next x
                Next i
    
                For i = 1 To MAX_PLAYER_TRADES
                    For x = 1 To MAX_INV
                        If GetPlayerInvItemNum(index, x) < 1 Then
                            If Player(TradeIndex).Trading(i).InvNum > 0 Then
                                Call GiveItem(index, GetPlayerInvItemNum(TradeIndex, Player(TradeIndex).Trading(i).InvNum), 1)
                                Call TakeItem(TradeIndex, GetPlayerInvItemNum(TradeIndex, Player(TradeIndex).Trading(i).InvNum), 1)
                                Exit For
                            End If
                        End If
                    Next x
                Next i

                Call PlayerMsg(index, "El comercio se completo con exito!", BRIGHTGREEN)
                Call PlayerMsg(TradeIndex, "El comercio se completo con exito!", BRIGHTGREEN)

                Call SendInventory(index)
                Call SendInventory(TradeIndex)
            Else
                If Player(index).TradeItemMax2 < Player(index).TradeItemMax Then
                    Call PlayerMsg(index, "Tu inventario esta lleno!", BRIGHTRED)
                    Call PlayerMsg(TradeIndex, GetPlayerName(index) & " tiene el inventario lleno!", BRIGHTRED)
                End If
                        
                If Player(TradeIndex).TradeItemMax2 < Player(TradeIndex).TradeItemMax Then
                    Call PlayerMsg(TradeIndex, "Tu inventario esta lleno!", BRIGHTRED)
                    Call PlayerMsg(index, GetPlayerName(TradeIndex) & " tiene el inventario lleno!", BRIGHTRED)
                End If
            End If
    
            Player(index).TradePlayer = 0
            Player(index).InTrade = False
            Player(index).TradeOk = 0

            Player(TradeIndex).TradePlayer = 0
            Player(TradeIndex).InTrade = False
            Player(TradeIndex).TradeOk = 0

            Call SendDataTo(index, "qtrade" & END_CHAR)
            Call SendDataTo(TradeIndex, "qtrade" & END_CHAR)
        End If
    End If
End Sub



Public Sub Packet_Spells(ByVal index As Long)
    Call SendPlayerSpells(index)
End Sub

Public Sub Packet_HotScript(ByVal index As Long, ByVal ScriptID As Long)
    If scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "HotScript " & index & "," & ScriptID
    End If
End Sub

Public Sub Packet_ScriptTile(ByVal index As Long, ByVal TileNum As Long)
    Call SendDataTo(index, "SCRIPTTILE" & SEP_CHAR & GetVar(App.Path & "\Tiles.ini", "Names", "Tile" & TileNum) & END_CHAR)
End Sub

Public Sub Packet_Cast(ByVal index As Long, ByVal spellnum As Long)
    Call CastSpell(index, spellnum)
End Sub

Public Sub Packet_Refresh(ByVal index As Long)
    Call SendDataToMap(GetPlayerMap(index), "playerxy" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & END_CHAR)
End Sub

Public Sub Packet_BuySprite(ByVal index As Long)
    Dim i As Long

    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type <> TILE_TYPE_SPRITE_CHANGE Then
        Call PlayerMsg(index, "Necesitas estar en un tile de sprite para comprarlo!", BRIGHTRED)
        Exit Sub
    End If

    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 = 0 Then
        Call SetPlayerSprite(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
        Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & END_CHAR)
        Exit Sub
    End If

    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(index, i) = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 Then
            If Item(GetPlayerInvItemNum(index, i)).Type = ITEM_TYPE_CURRENCY Then
                If GetPlayerInvItemValue(index, i) >= Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3 Then
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3)

                    If GetPlayerInvItemValue(index, i) = 0 Then
                        Call SetPlayerInvItemNum(index, i, 0)
                    End If

                    Call PlayerMsg(index, "Has comprado un nuevo sprite!", BRIGHTGREEN)
                    Call SetPlayerSprite(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
                    Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & END_CHAR)
                    Call SendInventory(index)
                End If
            Else
                If GetPlayerWeaponSlot(index) <> i And GetPlayerArmorSlot(index) <> i And GetPlayerShieldSlot(index) <> i And GetPlayerHelmetSlot(index) <> i And GetPlayerLegsSlot(index) <> i And GetPlayerRingSlot(index) <> i And GetPlayerNecklaceSlot(index) <> i Then
                    Call SetPlayerInvItemNum(index, i, 0)
                    Call PlayerMsg(index, "Has comprado un nuevo sprite!", BRIGHTGREEN)
                    Call SetPlayerSprite(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
                    Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & END_CHAR)
                    Call SendInventory(index)
                End If
            End If

            If GetPlayerWeaponSlot(index) <> i And GetPlayerArmorSlot(index) <> i And GetPlayerShieldSlot(index) <> i And GetPlayerHelmetSlot(index) <> i And GetPlayerLegsSlot(index) <> i And GetPlayerRingSlot(index) <> i And GetPlayerNecklaceSlot(index) <> i Then
                Exit Sub
            End If
        End If
    Next i

    Call PlayerMsg(index, "No tienes suficiente dinero para comprar este sprite!", BRIGHTRED)
End Sub

Public Sub Packet_ClearOwner(ByVal index As Long)
    Dim mapnum As Long

    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    mapnum = GetPlayerMap(index)

    Map(mapnum).Owner = 0
    Map(mapnum).Name = "Casa Abandonada"
    Map(mapnum).Revision = Map(mapnum).Revision + 1

    Call SaveMap(mapnum)

    Call SendDataToMap(mapnum, "CHECKFORMAP" & SEP_CHAR & mapnum & SEP_CHAR & (Map(mapnum).Revision + 1) & END_CHAR)

    Call PlayerMsg(index, "El propietario de la casa fue completamente eliminado.", BRIGHTRED)
End Sub

Public Sub Packet_RequestEditHouse(ByVal index As Long)
    If Map(GetPlayerMap(index)).Moral <> MAP_MORAL_HOUSE Then
        Call PlayerMsg(index, "Esto no es una casa!", BRIGHTRED)
        Exit Sub
    End If

    If Map(GetPlayerMap(index)).Owner <> GetPlayerName(index) Then
        Call PlayerMsg(index, "Esta no es tu casa!", BRIGHTRED)
        Exit Sub
    End If

    Call SendDataTo(index, "EDITHOUSE" & END_CHAR)
End Sub

Public Sub Packet_BuyHouse(ByVal index As Long)
    Dim i As Long

    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type <> TILE_TYPE_HOUSE Then
        Call PlayerMsg(index, "Necesitas estar en una tile de casa para poder comprarla!", BRIGHTRED)
        Exit Sub
    End If

    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 = 0 Then
        Map(GetPlayerMap(index)).Owner = GetPlayerName(index)
        Map(GetPlayerMap(index)).Name = "Casa de " & GetPlayerName(index)
        Map(GetPlayerMap(index)).Revision = Map(GetPlayerMap(index)).Revision + 1

        Call SaveMap(GetPlayerMap(index))
        Call SendDataToMap(GetPlayerMap(index), "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & (Map(GetPlayerMap(index)).Revision + 1) & END_CHAR)

        Exit Sub
    End If

    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(index, i) = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 Then
            If Item(GetPlayerInvItemNum(index, i)).Type = ITEM_TYPE_CURRENCY Then
                If GetPlayerInvItemValue(index, i) >= Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 Then
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2)

                    If GetPlayerInvItemValue(index, i) = 0 Then
                        Call SetPlayerInvItemNum(index, i, 0)
                    End If

                    Map(GetPlayerMap(index)).Owner = GetPlayerName(index)
                    Map(GetPlayerMap(index)).Name = "Casa de " & GetPlayerName(index)
                    Map(GetPlayerMap(index)).Revision = Map(GetPlayerMap(index)).Revision + 1

                    Call SaveMap(GetPlayerMap(index))
                    Call SendDataToMap(GetPlayerMap(index), "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & (Map(GetPlayerMap(index)).Revision + 1) & END_CHAR)
                    Call SendInventory(index)

                    Call PlayerMsg(index, "Has comprado una nueva casa!", BRIGHTGREEN)
                End If
            Else
                If GetPlayerWeaponSlot(index) <> i And GetPlayerArmorSlot(index) <> i And GetPlayerShieldSlot(index) <> i And GetPlayerHelmetSlot(index) <> i And GetPlayerLegsSlot(index) <> i And GetPlayerRingSlot(index) <> i And GetPlayerNecklaceSlot(index) <> i Then
                    Call SetPlayerInvItemNum(index, i, 0)

                    Map(GetPlayerMap(index)).Owner = GetPlayerName(index)
                    Map(GetPlayerMap(index)).Name = "Casa de " & GetPlayerName(index)
                    Map(GetPlayerMap(index)).Revision = Map(GetPlayerMap(index)).Revision + 1

                    Call SaveMap(GetPlayerMap(index))
                    Call SendDataToMap(GetPlayerMap(index), "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & (Map(GetPlayerMap(index)).Revision + 1) & END_CHAR)
                    Call SendInventory(index)

                    Call PlayerMsg(index, "Ahora eres dueño de una nueva casa!", BRIGHTGREEN)
                End If
            End If

            If GetPlayerWeaponSlot(index) <> i And GetPlayerArmorSlot(index) <> i And GetPlayerShieldSlot(index) <> i And GetPlayerHelmetSlot(index) <> i And GetPlayerLegsSlot(index) <> i And GetPlayerRingSlot(index) <> i And GetPlayerNecklaceSlot(index) <> i Then
                Exit Sub
            End If
        End If
    Next i

    Call PlayerMsg(index, "No tienes suficiente dinero para comprar esta casa!", BRIGHTRED)
End Sub

Public Sub Packet_CheckCommands(ByVal index As Long, ByVal Command As String)
    If scripting = 1 Then
        PutVar App.Path & "\Scripts\Comandos.ini", "TEMP", "Text" & index, Trim$(Command)
        MyScript.ExecuteStatement "Scripts\Main.txt", "Commands " & index
    Else
        Call PlayerMsg(index, "Este no es un comando valido!", BRIGHTRED)
    End If
End Sub

Public Sub Packet_Prompt(ByVal index As Long, ByVal PromptNum As Long, ByVal Value As Long)
    If scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerPrompt " & index & "," & PromptNum & "," & Value
    End If
End Sub

Public Sub Packet_QueryBox(ByVal index As Long, ByVal Response As String, ByVal PromptNum As Long)
    If scripting = 1 Then
        Call PutVar(App.Path & "\Responses.ini", "Responses", CStr(index), Response)
        MyScript.ExecuteStatement "Scripts\Main.txt", "QueryBox " & index & "," & PromptNum
    End If
End Sub

Public Sub Packet_RequestEditArrow(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(index, "arroweditor" & END_CHAR)
End Sub

Public Sub Packet_EditArrow(ByVal index As Long, ByVal ArrowNum As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If ArrowNum < 0 Or ArrowNum > MAX_ARROWS Then
        Call HackingAttempt(index, "Invalid Arrow Index")
        Exit Sub
    End If

    Call SendEditArrowTo(index, ArrowNum)

    Call AddLog(GetPlayerName(index) & " edita la flecha #" & ArrowNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveArrow(ByVal index As Long, ByVal ArrowNum As Long, ByVal Name As String, ByVal Pic As Long, ByVal Range As Long, ByVal Amount As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If ArrowNum < 0 Or ArrowNum > MAX_ITEMS Then
        Call HackingAttempt(index, "Invalid Arrow Index")
        Exit Sub
    End If

    Arrows(ArrowNum).Name = Name
    Arrows(ArrowNum).Pic = Pic
    Arrows(ArrowNum).Range = Range
    Arrows(ArrowNum).Amount = Amount

    Call SendUpdateArrowToAll(ArrowNum)
    Call SaveArrow(ArrowNum)

    Call AddLog(GetPlayerName(index) & " guarda la flecha #" & ArrowNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_CheckArrows(ByVal index As Long, ByVal ArrowNum As Long)
    Call SendDataToMap(GetPlayerMap(index), "checkarrows" & SEP_CHAR & index & SEP_CHAR & Arrows(ArrowNum).Pic & END_CHAR)
End Sub


Public Sub Packet_RequestEditEmoticon(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(index, "emoticoneditor" & END_CHAR)
End Sub

Public Sub Packet_RequestEditElement(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(index, "elementeditor" & END_CHAR)
End Sub

Public Sub Packet_RequestEditQuest(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Call SendDataTo(index, "questeditor" & END_CHAR)
End Sub

Public Sub Packet_EditEmoticon(ByVal index As Long, ByVal EmoteNum As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If EmoteNum < 0 Or EmoteNum > MAX_EMOTICONS Then
        Call HackingAttempt(index, "Invalid Emoticon Index")
        Exit Sub
    End If

    Call SendEditEmoticonTo(index, EmoteNum)

    Call AddLog(GetPlayerName(index) & " edita el emoticono #" & EmoteNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_EditElement(ByVal index As Long, ByVal ElementNum As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If ElementNum < 0 Or ElementNum > MAX_ELEMENTS Then
        Call HackingAttempt(index, "Invalid Emoticon Index")
        Exit Sub
    End If

    Call SendEditElementTo(index, ElementNum)

    Call AddLog(GetPlayerName(index) & " edita el elemento #" & ElementNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveEmoticon(ByVal index As Long, ByVal EmoteNum As Long, ByVal Command As String, ByVal Pic As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If EmoteNum < 0 Or EmoteNum > MAX_EMOTICONS Then
        Call HackingAttempt(index, "Invalid Emoticon Index")
        Exit Sub
    End If

    Emoticons(EmoteNum).Command = Command
    Emoticons(EmoteNum).Pic = Pic

    Call SendUpdateEmoticonToAll(EmoteNum)
    Call SaveEmoticon(EmoteNum)

    Call AddLog(GetPlayerName(index) & " guarda el emoticono #" & EmoteNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_SaveElement(ByVal index As Long, ByVal ElementNum As Long, ByVal Name As String, ByVal Strong As Long, ByVal Weak As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    If ElementNum < 0 Or ElementNum > MAX_ELEMENTS Then
        Call HackingAttempt(index, "Invalid Element Index")
        Exit Sub
    End If

    Element(ElementNum).Name = Name
    Element(ElementNum).Strong = Strong
    Element(ElementNum).Weak = Weak

    Call SendUpdateElementToAll(ElementNum)
    Call SaveElement(ElementNum)

    Call AddLog(GetPlayerName(index) & " guarda el elemento #" & ElementNum & ".", ADMIN_LOG)
End Sub

Public Sub Packet_CheckEmoticon(ByVal index As Long, ByVal EmoteNum As Long)
    Call SendDataToMap(GetPlayerMap(index), "checkemoticons" & SEP_CHAR & index & SEP_CHAR & Emoticons(EmoteNum).Pic & END_CHAR)
End Sub

Public Sub Packet_MapReport(ByVal index As Long)
    Dim Packet As String
    Dim i As Long

    If GetPlayerAccess(index) < ADMIN_MONITER Then
        Call HackingAttempt(index, "Packet Modification")
        Exit Sub
    End If

    Packet = "mapreport" & SEP_CHAR

    For i = 1 To MAX_MAPS
        Packet = Packet & Map(i).Name & SEP_CHAR
    Next i

    Packet = Packet & END_CHAR

    Call SendDataTo(index, Packet)
End Sub

Public Sub Packet_GMTime(ByVal index As Long, ByVal Time As Long)
    If GetPlayerAccess(index) < ADMIN_MONITER Then
        Call HackingAttempt(index, "Packet Modification")
        Exit Sub
    End If

    GameTime = Time

    Call SendTimeToAll
End Sub

Public Sub Packet_Weather(ByVal index As Long, ByVal WeatherNum As Long)
    If GetPlayerAccess(index) < ADMIN_MONITER Then
        Call HackingAttempt(index, "Packet Modification")
        Exit Sub
    End If

    WeatherType = WeatherNum

    Call SendWeatherToAll
End Sub

Public Sub Packet_WarpTo(ByVal index As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    If GetPlayerAccess(index) < ADMIN_MONITER Then
        Call HackingAttempt(index, "Packet Modification")
        Exit Sub
    End If
    
    If x < 0 Or x > MAX_MAPX Then
        Call PlayerMsg(index, "Por favor, introduce una coordenada X valida.", BRIGHTRED)
        Exit Sub
    End If

    If y < 0 Or y > MAX_MAPY Then
        Call PlayerMsg(index, "Por favor, introduce una coordenada Y valida.", BRIGHTRED)
        Exit Sub
    End If

    Call PlayerWarp(index, mapnum, x, y)
End Sub

Public Sub Packet_LocalWarp(ByVal index As Long, ByVal x As Long, ByVal y As Long)
    If GetPlayerAccess(index) < ADMIN_MONITER Then
        Call HackingAttempt(index, "Packet Modification")
        Exit Sub
    End If
    
    If x < 0 Or x > MAX_MAPX Then
        Call PlayerMsg(index, "Por favor, introduce una coordenada X valida.", BRIGHTRED)
        Exit Sub
    End If

    If y < 0 Or y > MAX_MAPY Then
        Call PlayerMsg(index, "Por favor, introduce una coordenada Y valida.", BRIGHTRED)
        Exit Sub
    End If

    Player(index).Char(Player(index).CharNum).x = x
    Player(index).Char(Player(index).CharNum).y = y

    Call SendPlayerXY(index)
End Sub

Public Sub Packet_ArrowHit(ByVal index As Long, ByVal TargetType As Long, ByVal PlayerIndex As Long, ByVal x As Long, ByVal y As Long)
    Dim Damage As Long
    
    If TargetType = TARGET_TYPE_PLAYER Then
        If PlayerIndex <> index Then
            If CanAttackPlayerWithArrow(index, PlayerIndex) Then
                Player(index).Target = PlayerIndex
                Player(index).TargetType = TARGET_TYPE_PLAYER
                If Not CanPlayerBlockHit(PlayerIndex) Then
                    If Not CanPlayerCriticalHit(index) Then
                        Damage = GetPlayerDamage(index) - GetPlayerProtection(PlayerIndex)
                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & END_CHAR)
                    Else
                        TargetType = GetPlayerDamage(index)
                        Damage = TargetType + Int(Rnd * Int(TargetType / 2)) + 1 - GetPlayerProtection(PlayerIndex)

                        Call BattleMsg(index, "Sientes como tu punteria se afina como un aguila!", BRIGHTCYAN, 0)
                        Call BattleMsg(PlayerIndex, GetPlayerName(index) & " dispara con una punteria bestial!", BRIGHTCYAN, 1)

                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & END_CHAR)
                    End If

                    If Damage > 0 Then
                        If scripting = 1 Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "OnArrowHit " & index & "," & Damage
                        Else
                            Call AttackPlayer(index, PlayerIndex, Damage)
                        End If
                    Else
                        If scripting = 1 Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "OnArrowHit " & index & "," & 0
                        End If
                        Call BattleMsg(index, "Tu ataque no hace nada.", BRIGHTRED, 0)
                        Call BattleMsg(PlayerIndex, GetPlayerName(index) & " ataco sin hacer nada.", BRIGHTRED, 1)

                        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & END_CHAR)
                    End If
                Else
                
                    If scripting = 1 Then
                        MyScript.ExecuteStatement "Scripts\Main.txt", "OnArrowHit " & index & "," & 0
                    End If
                    Call BattleMsg(index, GetPlayerName(PlayerIndex) & " bloqueo tu golpe!", BRIGHTCYAN, 0)
                    Call BattleMsg(PlayerIndex, "Bloqueas el golpe de " & GetPlayerName(index) & "!", BRIGHTCYAN, 1)

                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & END_CHAR)
                End If

                Exit Sub
            End If
        End If
    ElseIf TargetType = TARGET_TYPE_NPC Then
        If CanAttackNpcWithArrow(index, PlayerIndex) Then
        Player(index).TargetType = TARGET_TYPE_NPC
        Player(index).TargetNPC = PlayerIndex
            If Not CanPlayerCriticalHit(index) Then
                Damage = GetPlayerDamage(index) - Int(NPC(MapNPC(GetPlayerMap(index), PlayerIndex).num).DEF / 2)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "attack" & END_CHAR)
            Else
                TargetType = GetPlayerDamage(index)
                Damage = TargetType + Int(Rnd * Int(TargetType / 2)) + 1 - Int(NPC(MapNPC(GetPlayerMap(index), PlayerIndex).num).DEF / 2)

                Call BattleMsg(index, "Sientes como tu punteria se afina como un aguila!", BRIGHTCYAN, 0)

                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "critical" & END_CHAR)
            End If

            If Damage > 0 Then
                If scripting = 1 Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "OnArrowHit " & index & "," & Damage
                Else
                    Call AttackNpc(index, PlayerIndex, Damage)
                    Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & PlayerIndex & END_CHAR)
                End If
            Else
                If scripting = 1 Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "OnArrowHit " & index & "," & Damage
                End If
                Call BattleMsg(index, "Tu ataque no hace nada.", BRIGHTRED, 0)

                Call SendDataTo(index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & PlayerIndex & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "miss" & END_CHAR)
            End If

            Exit Sub
        End If
    End If
End Sub

Public Sub Packet_BankDeposit(ByVal index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    Dim BankSlot As Long
    Dim ItemNum As Long

    ItemNum = GetPlayerInvItemNum(index, InvNum)

    BankSlot = FindOpenBankSlot(index, ItemNum)
    If BankSlot = 0 Then
        Call SendDataTo(index, "bankmsg" & SEP_CHAR & "Bank full!" & END_CHAR)
        Exit Sub
    End If

    If Amount > GetPlayerInvItemValue(index, InvNum) Then
        Call SendDataTo(index, "bankmsg" & SEP_CHAR & "No puedes depositar mas de lo que tienes!" & END_CHAR)
        Exit Sub
    End If

    If GetPlayerWeaponSlot(index) = ItemNum Or GetPlayerArmorSlot(index) = ItemNum Or GetPlayerShieldSlot(index) = ItemNum Or GetPlayerHelmetSlot(index) = ItemNum Or GetPlayerLegsSlot(index) = ItemNum Or GetPlayerRingSlot(index) = ItemNum Or GetPlayerNecklaceSlot(index) = ItemNum Then
        Call SendDataTo(index, "bankmsg" & SEP_CHAR & "No puedes depositar objetos equipados!" & END_CHAR)
        Exit Sub
    End If

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
        If Amount = 0 Then
            Call SendDataTo(index, "bankmsg" & SEP_CHAR & "Necesitas depositar una cantidad mayor de 0!" & END_CHAR)
            Exit Sub
        End If
    End If

    Call TakeItem(index, ItemNum, Amount)
    Call GiveBankItem(index, ItemNum, Amount, BankSlot)

    Call SendBank(index)
End Sub

Public Sub Packet_BankWithdraw(ByVal index As Long, ByVal BankInvNum As Long, ByVal Amount As Long)
    Dim BankItemNum As Long
    Dim BankInvSlot As Long

    BankItemNum = GetPlayerBankItemNum(index, BankInvNum)

    BankInvSlot = FindOpenInvSlot(index, BankItemNum)
    If BankInvSlot = 0 Then
        Call SendDataTo(index, "bankmsg" & SEP_CHAR & "Inventory full!" & END_CHAR)
        Exit Sub
    End If

    If Amount > GetPlayerBankItemValue(index, BankInvNum) Then
        Call SendDataTo(index, "bankmsg" & SEP_CHAR & "No puedes retirar mas de lo que tu tienes!" & END_CHAR)
        Exit Sub
    End If

    If Item(BankItemNum).Type = ITEM_TYPE_CURRENCY Or Item(BankItemNum).Stackable = 1 Then
        If Amount = 0 Then
            Call SendDataTo(index, "bankmsg" & SEP_CHAR & "Necesitas retirar una cantidad mayor de 0!" & END_CHAR)
            Exit Sub
        End If
    End If

    Call TakeBankItem(index, BankItemNum, Amount)
    Call GiveItem(index, BankItemNum, Amount)

    Call SendBank(index)
End Sub

Public Sub Packet_ReloadScripts(ByVal index As Long)
    If GetPlayerAccess(index) < ADMIN_DEVELOPER Then
        Call HackingAttempt(index, "Packet Modification")
        Exit Sub
    End If

    Set MyScript = Nothing
    Set clsScriptCommands = Nothing

    Set MyScript = New clsSadScript
    Set clsScriptCommands = New clsCommands

    MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl
    MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True

    MyScript.ExecuteStatement "Scripts\Main.txt", "OnScriptReload"

    Call TextAdd(frmServer.txtText(0), "Scripts recargados.", True)
    Call AdminMsg("Scripts recargados por " & GetPlayerName(index) & ".", WHITE)
End Sub

Public Sub Packet_CustomMenuClick(ByVal index As Long, ByVal MenuIndex As Long, ByVal ClickIndex As Long, ByVal CustomTitle As String, ByVal MenuType As Long, ByVal CustomMsg As String)
    Player(index).CustomTitle = CustomTitle
    Player(index).CustomMsg = CustomMsg

    If scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "menuscripts " & MenuIndex & "," & ClickIndex & "," & MenuType
    End If
End Sub

Public Sub Packet_CustomBoxReturnMsg(ByVal index As Long, ByVal CustomMsg As String)
    Player(index).CustomMsg = CustomMsg
End Sub

Public Sub Packet_RequestEditMain(ByVal index As Long)
    ' Prevent hacking
    If GetPlayerAccess(index) < ADMIN_CREATOR Then
        Call HackingAttempt(index, "Admin Cloning")
        Exit Sub
    End If

    Dim F
    F = FreeFile
    Open App.Path & "Scripts\Main.txt" For Input As #F
   
    Dim Text
    Text = Input$(LOF(F), F)
    Close #F
   
    Call SendDataTo(index, "MAINEDITOR" & SEP_CHAR & Text & SEP_CHAR & END_CHAR)
   
End Sub

Public Sub Packet_NewMain(ByVal index As Long, FileContents)
    If GetPlayerAccess(index) >= ADMIN_CREATOR Then
        Dim temp As String
        Dim F

        F = FreeFile
        Open App.Path & "Scripts\Main.txt" For Input As #F
        temp = Input$(LOF(F), F)
        Close #F
        F = FreeFile
        Open App.Path & "Scripts\Backup.txt" For Output As #F
        Print #F, temp
        Close #F
        F = FreeFile
        Open App.Path & "Scripts\Main.txt" For Output As #F
        Print #F, FileContents
        Close #F

        If scripting = 1 Then
            Set MyScript = Nothing
            Set clsScriptCommands = Nothing
            Set MyScript = New clsSadScript
            Set clsScriptCommands = New clsCommands
            MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
        End If
       
        Call Packet_ReloadScripts(index)
        Call AddLog(GetPlayerName(index) & " ha actualizado el script.", ADMIN_LOG)
    End If
End Sub

Public Sub Packet_loadsct(ByVal index As Long, ByVal spellnum As Long)
Call SendDataTo(index, "loadsctt" & SEP_CHAR & spellnum & SEP_CHAR & Spell(spellnum).CastTimer & SEP_CHAR & Spell(spellnum).TimeToCast & SEP_CHAR & Spell(spellnum).MPCost & END_CHAR)
End Sub

Public Sub Packet_SPass(ByVal index As Long)
Dim Pass As String

Pass = GetSetting(App.EXEName, "Clave", "Clave")

Call SendDataTo(index, "SPASS" & SEP_CHAR & Pass & END_CHAR)
End Sub

