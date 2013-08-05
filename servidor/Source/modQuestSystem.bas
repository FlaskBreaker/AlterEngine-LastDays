Attribute VB_Name = "modQuestSystem"
' AlterEngine - www.alterengine.net

Option Explicit

Function GetPlayerNpcKillType(ByVal Index As Long) As String
    GetPlayerNpcKillType = Trim$(Player(Index).Char(Player(Index).CharNum).NpcKillType)
End Function

Sub SetPlayerNpcKillType(ByVal Index As Long, _
   ByVal NpcKillType As String)
    Player(Index).Char(Player(Index).CharNum).NpcKillType = NpcKillType
End Sub

Function GetPlayerNpcKillAmount(ByVal Index As Long) As Long
    GetPlayerNpcKillAmount = Player(Index).Char(Player(Index).CharNum).NpcKillamount
End Function

Sub SetPlayerNpcKillAmount(ByVal Index As Long, _
   ByVal NpcKillamount As Long)
    Player(Index).Char(Player(Index).CharNum).NpcKillamount = NpcKillamount
End Sub

Function GetPlayerNpcKillQuestFlag(ByVal Index As Long) As Long
    GetPlayerNpcKillQuestFlag = Player(Index).Char(Player(Index).CharNum).NpcKillQuestFlag
End Function

Sub SetPlayerNpcKillQuestFlag(ByVal Index As Long, _
   ByVal NpcKillQuestFlag As Long)
    Player(Index).Char(Player(Index).CharNum).NpcKillQuestFlag = NpcKillQuestFlag
End Sub

Function GetPlayerQuestFlag(ByVal Index As Long, ByVal QuestFlagSlot As Long) As Long
    GetPlayerQuestFlag = Player(Index).Char(Player(Index).CharNum).QuestFlags(QuestFlagSlot)
End Function

Sub SetPlayerQuestFlag(ByVal Index As Long, _
   ByVal QuestFlagSlot As Long, _
   ByVal QuestFlagnum As Long)
    Player(Index).Char(Player(Index).CharNum).QuestFlags(QuestFlagSlot) = QuestFlagnum
    Call SendPlayerQuestFlags(Index)
End Sub

Sub SendPlayerQuestFlags(ByVal Index As Long)
Dim Packet As String
Dim I As Long

    Packet = "QUESTFLAGS" & SEP_CHAR
    For I = 1 To MAX_QUESTS
        Packet = Packet & GetPlayerQuestFlag(Index, I) & SEP_CHAR
    Next
    Packet = Packet & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SaveQuest(ByVal questnum As Long)
Dim filename As String
Dim F As Long

    filename = App.Path & "\main\quests\quests" & questnum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Quest(questnum)
    Close #F
End Sub

Sub SaveQuests()
Dim I As Long

    Call SetStatus("Guardando quests... ")
    For I = 1 To MAX_QUESTS

        If Not FileExist("main\quests\quests" & I & ".dat") Then
            Call SetStatus("Guardando quests.. " & Int((I / MAX_QUESTS) * 100) & "%")
            DoEvents

            Call SaveQuest(I)
        End If
    Next
End Sub

Sub CheckQuests()
    Call SaveQuests
End Sub
Sub LoadQuests()
Dim filename As String
Dim I As Long
Dim F As Long

    Call CheckQuests
    For I = 1 To MAX_QUESTS
        Call SetStatus("Cargando quests... " & Int((I / MAX_QUESTS) * 100) & "%")
        filename = App.Path & "\Main\Quests\quests" & I & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Quest(I)
        Close #F
        DoEvents

    Next
End Sub

Sub ClearQuests()
Dim I As Long

For I = 1 To MAX_QUESTS
Quest(I).Name = ""
Quest(I).LevelIsReq = 0
Quest(I).ClassIsReq = 0
Quest(I).StartOn = 0
Quest(I).LevelReq = 0
Quest(I).ClassReq = 0

Quest(I).StartItem = 0
Quest(I).Startval = 0
Quest(I).ItemReq = 0
Quest(I).ItemVal = 0
Quest(I).RewardNum = 0
Quest(I).RewardVal = 0
Quest(I).Start = ""
Quest(I).End = ""
Quest(I).During = ""
Quest(I).NotHasItem = ""
Quest(I).Before = ""
Quest(I).After = ""
Quest(I).FishExp = 0
Quest(I).MineExp = 0
Quest(I).LJackingExp = 0
Quest(I).ForagingExp = 0
Quest(I).UnArmedExp = 0
Quest(I).MageWeaponsExp = 0
Quest(I).CombatExp = 0
Quest(I).SmeltingExp = 0
Quest(I).IronForgingExp = 0
Quest(I).LeaderShipExp = 0
Quest(I).GovernmentExp = 0
Quest(I).CriticalHitExp = 0
Quest(I).DodgeExp = 0
Quest(I).RepExp = 0
Quest(I).PKillExp = 0
Quest(I).ThiefExp = 0
Quest(I).LargeBladesExp = 0
Quest(I).SmallBladesExp = 0
Quest(I).BluntWeaponsExp = 0
Quest(I).PolesExp = 0
Quest(I).AxesExp = 0
Quest(I).ThrownExp = 0
Quest(I).XbowsExp = 0
Quest(I).BowsExp = 0
Quest(I).CarpentryExp = 0
Quest(I).MillingExp = 0
Quest(I).SpinningExp = 0
Quest(I).WeavingExp = 0
Quest(I).SewingExp = 0
Quest(I).PlantingExp = 0
Quest(I).HarvestingExp = 0
Quest(I).LeatherWorkingExp = 0
Quest(I).SkinningExp = 0
Quest(I).TanningExp = 0
Quest(I).BodyExp = 0
Quest(I).MindExp = 0
Quest(I).SoulExp = 0
Quest(I).NatureExp = 0
Quest(I).AlchemyExp = 0
Quest(I).QuestingExp = 0
Quest(I).FirstAidExp = 0
Quest(I).QuestExpReward = 0
Next I
End Sub

Sub SendQuest(ByVal Index As Long)
'Dim Packet As String
Dim I As Long

    For I = 1 To MAX_QUESTS
        If Trim(Quest(I).Name) <> "" Then
            Call SendUpdateQuestTo(Index, I)
        End If
    Next I
End Sub

Sub SendUpdateQuestToAll(ByVal questnum As Long)
Dim Packet As String

    Packet = "UPDATEQUEST" & SEP_CHAR & questnum & SEP_CHAR & Trim(Quest(questnum).Name) & SEP_CHAR & Trim(Quest(questnum).After) & SEP_CHAR & Trim(Quest(questnum).Before) & SEP_CHAR & Quest(questnum).ClassIsReq & SEP_CHAR & Quest(questnum).ClassReq & SEP_CHAR & Trim(Quest(questnum).During) & SEP_CHAR & Trim(Quest(questnum).End) & SEP_CHAR & Quest(questnum).ItemReq & SEP_CHAR & Quest(questnum).ItemVal & SEP_CHAR & Quest(questnum).LevelIsReq & SEP_CHAR & Quest(questnum).LevelReq & SEP_CHAR & Trim(Quest(questnum).NotHasItem) & SEP_CHAR & Quest(questnum).RewardNum & SEP_CHAR & Quest(questnum).RewardVal & SEP_CHAR & Trim(Quest(questnum).Start) & SEP_CHAR & Quest(questnum).StartItem & SEP_CHAR & Quest(questnum).StartOn & SEP_CHAR & Quest(questnum).Startval & SEP_CHAR & Quest(questnum).FishExp & SEP_CHAR & Quest(questnum).MineExp & SEP_CHAR & Quest(questnum).LJackingExp & SEP_CHAR & Quest(questnum).ForagingExp & SEP_CHAR & Quest(questnum).UnArmedExp & SEP_CHAR & Quest(questnum).MageWeaponsExp & SEP_CHAR
Packet = Packet & Quest(questnum).CombatExp & SEP_CHAR & Quest(questnum).SmeltingExp & SEP_CHAR & Quest(questnum).IronForgingExp & SEP_CHAR & Quest(questnum).LeaderShipExp & SEP_CHAR & Quest(questnum).GovernmentExp & SEP_CHAR & Quest(questnum).CriticalHitExp & SEP_CHAR & Quest(questnum).DodgeExp & SEP_CHAR & Quest(questnum).RepExp & SEP_CHAR & Quest(questnum).PKillExp & SEP_CHAR & Quest(questnum).ThiefExp & SEP_CHAR & Quest(questnum).LargeBladesExp & SEP_CHAR & Quest(questnum).SmallBladesExp & SEP_CHAR & Quest(questnum).BluntWeaponsExp & SEP_CHAR & Quest(questnum).PolesExp & SEP_CHAR & Quest(questnum).AxesExp & SEP_CHAR & Quest(questnum).ThrownExp & SEP_CHAR & Quest(questnum).XbowsExp & SEP_CHAR & Quest(questnum).BowsExp & SEP_CHAR & Quest(questnum).CarpentryExp & SEP_CHAR & Quest(questnum).MillingExp & SEP_CHAR & Quest(questnum).SpinningExp & SEP_CHAR & Quest(questnum).WeavingExp & SEP_CHAR & Quest(questnum).SewingExp & SEP_CHAR & Quest(questnum).PlantingExp & SEP_CHAR
Packet = Packet & Quest(questnum).HarvestingExp & SEP_CHAR & Quest(questnum).LeatherWorkingExp & SEP_CHAR & Quest(questnum).SkinningExp & SEP_CHAR & Quest(questnum).TanningExp & SEP_CHAR & Quest(questnum).BodyExp & SEP_CHAR & Quest(questnum).MindExp & SEP_CHAR & Quest(questnum).SoulExp & SEP_CHAR & Quest(questnum).NatureExp & SEP_CHAR & Quest(questnum).AlchemyExp & SEP_CHAR & Quest(questnum).QuestingExp & SEP_CHAR & Quest(questnum).FirstAidExp & SEP_CHAR & Quest(questnum).QuestExpReward & SEP_CHAR & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateQuestTo(ByVal Index As Long, ByVal questnum As Long)
Dim Packet As String

    Packet = "UPDATEQUEST" & SEP_CHAR & questnum & SEP_CHAR & Trim(Quest(questnum).Name) & SEP_CHAR & Trim(Quest(questnum).After) & SEP_CHAR & Trim(Quest(questnum).Before) & SEP_CHAR & Quest(questnum).ClassIsReq & SEP_CHAR & Quest(questnum).ClassReq & SEP_CHAR & Trim(Quest(questnum).During) & SEP_CHAR & Trim(Quest(questnum).End) & SEP_CHAR & Quest(questnum).ItemReq & SEP_CHAR & Quest(questnum).ItemVal & SEP_CHAR & Quest(questnum).LevelIsReq & SEP_CHAR & Quest(questnum).LevelReq & SEP_CHAR & Trim(Quest(questnum).NotHasItem) & SEP_CHAR & Quest(questnum).RewardNum & SEP_CHAR & Quest(questnum).RewardVal & SEP_CHAR & Trim(Quest(questnum).Start) & SEP_CHAR & Quest(questnum).StartItem & SEP_CHAR & Quest(questnum).StartOn & SEP_CHAR & Quest(questnum).Startval & SEP_CHAR & Quest(questnum).FishExp & SEP_CHAR & Quest(questnum).MineExp & SEP_CHAR & Quest(questnum).LJackingExp & SEP_CHAR & Quest(questnum).ForagingExp & SEP_CHAR & Quest(questnum).UnArmedExp & SEP_CHAR & Quest(questnum).MageWeaponsExp & SEP_CHAR
Packet = Packet & Quest(questnum).CombatExp & SEP_CHAR & Quest(questnum).SmeltingExp & SEP_CHAR & Quest(questnum).IronForgingExp & SEP_CHAR & Quest(questnum).LeaderShipExp & SEP_CHAR & Quest(questnum).GovernmentExp & SEP_CHAR & Quest(questnum).CriticalHitExp & SEP_CHAR & Quest(questnum).DodgeExp & SEP_CHAR & Quest(questnum).RepExp & SEP_CHAR & Quest(questnum).PKillExp & SEP_CHAR & Quest(questnum).ThiefExp & SEP_CHAR & Quest(questnum).LargeBladesExp & SEP_CHAR & Quest(questnum).SmallBladesExp & SEP_CHAR & Quest(questnum).BluntWeaponsExp & SEP_CHAR & Quest(questnum).PolesExp & SEP_CHAR & Quest(questnum).AxesExp & SEP_CHAR & Quest(questnum).ThrownExp & SEP_CHAR & Quest(questnum).XbowsExp & SEP_CHAR & Quest(questnum).BowsExp & SEP_CHAR & Quest(questnum).CarpentryExp & SEP_CHAR & Quest(questnum).MillingExp & SEP_CHAR & Quest(questnum).SpinningExp & SEP_CHAR & Quest(questnum).WeavingExp & SEP_CHAR & Quest(questnum).SewingExp & SEP_CHAR & Quest(questnum).PlantingExp & SEP_CHAR
Packet = Packet & Quest(questnum).HarvestingExp & SEP_CHAR & Quest(questnum).LeatherWorkingExp & SEP_CHAR & Quest(questnum).SkinningExp & SEP_CHAR & Quest(questnum).TanningExp & SEP_CHAR & Quest(questnum).BodyExp & SEP_CHAR & Quest(questnum).MindExp & SEP_CHAR & Quest(questnum).SoulExp & SEP_CHAR & Quest(questnum).NatureExp & SEP_CHAR & Quest(questnum).AlchemyExp & SEP_CHAR & Quest(questnum).QuestingExp & SEP_CHAR & Quest(questnum).FirstAidExp & SEP_CHAR & Quest(questnum).QuestExpReward & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub SendEditQuestTo(ByVal Index As Long, ByVal questnum As Long)
Dim Packet As String

    Packet = "EDITQUEST" & SEP_CHAR & questnum & SEP_CHAR & Trim(Quest(questnum).Name) & SEP_CHAR & Trim(Quest(questnum).After) & SEP_CHAR & Trim(Quest(questnum).Before) & SEP_CHAR & Quest(questnum).ClassIsReq & SEP_CHAR & Quest(questnum).ClassReq & SEP_CHAR & Trim(Quest(questnum).During) & SEP_CHAR & Trim(Quest(questnum).End) & SEP_CHAR & Quest(questnum).ItemReq & SEP_CHAR & Quest(questnum).ItemVal & SEP_CHAR & Quest(questnum).LevelIsReq & SEP_CHAR & Quest(questnum).LevelReq & SEP_CHAR & Trim(Quest(questnum).NotHasItem) & SEP_CHAR & Quest(questnum).RewardNum & SEP_CHAR & Quest(questnum).RewardVal & SEP_CHAR & Trim(Quest(questnum).Start) & SEP_CHAR & Quest(questnum).StartItem & SEP_CHAR & Quest(questnum).StartOn & SEP_CHAR & Quest(questnum).Startval & SEP_CHAR & Quest(questnum).FishExp & SEP_CHAR & Quest(questnum).MineExp & SEP_CHAR & Quest(questnum).LJackingExp & SEP_CHAR & Quest(questnum).ForagingExp & SEP_CHAR & Quest(questnum).UnArmedExp & SEP_CHAR & Quest(questnum).MageWeaponsExp & SEP_CHAR
    Packet = Packet & Quest(questnum).CombatExp & SEP_CHAR & Quest(questnum).SmeltingExp & SEP_CHAR & Quest(questnum).IronForgingExp & SEP_CHAR & Quest(questnum).LeaderShipExp & SEP_CHAR & Quest(questnum).GovernmentExp & SEP_CHAR & Quest(questnum).CriticalHitExp & SEP_CHAR & Quest(questnum).DodgeExp & SEP_CHAR & Quest(questnum).RepExp & SEP_CHAR & Quest(questnum).PKillExp & SEP_CHAR & Quest(questnum).ThiefExp & SEP_CHAR & Quest(questnum).LargeBladesExp & SEP_CHAR & Quest(questnum).SmallBladesExp & SEP_CHAR & Quest(questnum).BluntWeaponsExp & SEP_CHAR & Quest(questnum).PolesExp & SEP_CHAR & Quest(questnum).AxesExp & SEP_CHAR & Quest(questnum).ThrownExp & SEP_CHAR & Quest(questnum).XbowsExp & SEP_CHAR & Quest(questnum).BowsExp & SEP_CHAR & Quest(questnum).CarpentryExp & SEP_CHAR & Quest(questnum).MillingExp & SEP_CHAR & Quest(questnum).SpinningExp & SEP_CHAR & Quest(questnum).WeavingExp & SEP_CHAR & Quest(questnum).SewingExp & SEP_CHAR & Quest(questnum).PlantingExp & SEP_CHAR
    Packet = Packet & Quest(questnum).HarvestingExp & SEP_CHAR & Quest(questnum).LeatherWorkingExp & SEP_CHAR & Quest(questnum).SkinningExp & SEP_CHAR & Quest(questnum).TanningExp & SEP_CHAR & Quest(questnum).BodyExp & SEP_CHAR & Quest(questnum).MindExp & SEP_CHAR & Quest(questnum).SoulExp & SEP_CHAR & Quest(questnum).NatureExp & SEP_CHAR & Quest(questnum).AlchemyExp & SEP_CHAR & Quest(questnum).QuestingExp & SEP_CHAR & Quest(questnum).FirstAidExp & SEP_CHAR & Quest(questnum).QuestExpReward & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub QuestMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte, ByVal Hate As Byte)
Dim Packet As String

    Packet = "QUESTMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & Hate & SEP_CHAR & END_CHAR
    Call SendDataTo(Index, Packet)
End Sub

Sub callrequstedEditQuest(ByVal Index As Long)
        Call SendDataTo(Index, "QUESTEDITOR" & SEP_CHAR & END_CHAR)
End Sub

Function ActuallyStartQuest(ByVal questnum As Long, ByVal Index As Long, ByVal ncpnum As Long)
Call SetPlayerQuestFlag(Index, questnum, 1)
Call QuestMsg(Index, "* Misión Recibida *", BLACK, 1)
Call QuestMsg(Index, "", BRIGHTGREEN, 1)
Call QuestMsg(Index, "" & Trim(NPC(ncpnum).Name) & " dice: " & Trim(Quest(NPC(ncpnum).Quest).Start) & "'", WHITE, 1)
If Quest(questnum).StartOn = 1 Then
        Call GiveQuestItem(Index, Quest(questnum).StartItem, Quest(questnum).Startval, ncpnum)
        Quest(questnum).StartOn = 0
End If
Call SendPlayerQuestFlags(Index)
End Function

'Tex
Function DoQuest(ByVal questnum As Long, ByVal Index As Long, ByVal npcnum As Long)
Dim BoB

If GetPlayerQuestFlag(Index, NPC(npcnum).Quest) = 0 Then
    If MeetReq(questnum, Index) Then
        'If Quest(questnum).StartOn = 0 Then
            Call SendDataTo(Index, "questinfo" & SEP_CHAR & questnum & SEP_CHAR & npcnum & SEP_CHAR & END_CHAR)
            Call QuestMsg(Index, "* Información de misión recibida *", BLACK, 2)
            Call QuestMsg(Index, "" & Trim(NPC(npcnum).Name) & " dice: " & Trim(Quest(NPC(npcnum).Quest).Before) & "'", WHITE, 2)
            ' Call AddLog(GetPlayerName(Index) & ": " & " has started " & questnum & " ! ", SUGGESTION_LOG)
        'ElseIf Quest(questnum).StartOn = 1 Then
        '    Call GiveQuestItem(Index, Quest(questnum).StartItem, Quest(questnum).Startval, npcnum)
        '    Quest(questnum).StartOn = 0
        '    Call SendDataTo(Index, "questinfo" & SEP_CHAR & questnum & SEP_CHAR & npcnum & SEP_CHAR & END_CHAR)
        '    Call QuestMsg(Index, "* Información de misión recibida *", BLACK, 2)
        '    Call QuestMsg(Index, "" & Trim(NPC(npcnum).Name) & " dice: " & Trim(Quest(NPC(npcnum).Quest).Before) & "'", WHITE, 2)
            'Call BattleMsg(Index, "TESTSTARTON", 4, 0)
        'End If
    Else
        If GetPlayerQuestFlag(Index, NPC(npcnum).Quest) = 2 Then
            Call QuestMsg(Index, "* Registro de la misión *", BLACK, 1)
            Call QuestMsg(Index, "" & Trim(NPC(npcnum).Name) & " dice: " & Trim(Quest(NPC(npcnum).Quest).After) & "'", WHITE, 1)
            Exit Function
        End If
        Call PlayerMsg(Index, "" & Trim(NPC(npcnum).Name) & " dice: " & Trim(Quest(NPC(npcnum).Quest).Before) & "'", WHITE)
    End If
    Exit Function
End If

If GetPlayerQuestFlag(Index, NPC(npcnum).Quest) = 1 Then
    Call SendDataTo(Index, "questprompt" & SEP_CHAR & questnum & SEP_CHAR & npcnum & SEP_CHAR & END_CHAR)
End If

If GetPlayerQuestFlag(Index, NPC(npcnum).Quest) = 2 Then
    Call QuestMsg(Index, "* Registro de la misión *", BLACK, 1)
    Call QuestMsg(Index, "", BRIGHTGREEN, 1)
    Call QuestMsg(Index, "" & Trim(NPC(npcnum).Name) & " dice: " & Trim(Quest(NPC(npcnum).Quest).After) & "'", WHITE, 1)
    Exit Function
End If

End Function

Sub SaveLine(File As Integer, Header As String, Var As String, Value As String)
    Print #File, Var & "=" & Value
End Sub

Function MeetReq(questnum As Long, Index As Long) As Boolean
If Quest(questnum).ClassIsReq = 0 And Quest(questnum).LevelIsReq = 0 Then
    MeetReq = True
    Exit Function
ElseIf Quest(questnum).ClassIsReq = 1 And Quest(questnum).LevelIsReq = 0 Then
    If Player(Index).Char(Player(Index).CharNum).Class = Quest(questnum).ClassReq Then
        MeetReq = True
        Exit Function
    Else
        MeetReq = False
        Exit Function
    End If
ElseIf Quest(questnum).ClassIsReq = 0 And Quest(questnum).LevelIsReq = 1 Then
    If Player(Index).Char(Player(Index).CharNum).Level >= Quest(questnum).LevelReq Then
        MeetReq = True
        Exit Function
    Else
        MeetReq = False
        Exit Function
    End If
ElseIf Quest(questnum).ClassIsReq = 1 And Quest(questnum).LevelIsReq = 1 Then
    If Player(Index).Char(Player(Index).CharNum).Class = Quest(questnum).ClassReq And Player(Index).Char(Player(Index).CharNum).Level >= Quest(questnum).LevelReq Then
        MeetReq = True
        Exit Function
    Else
        MeetReq = False
        Exit Function
    End If
End If

End Function

Sub GiveQuestItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal npcnum As Long)
Dim I As Long
Dim Curr As Boolean
Dim Has As Boolean
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    If (Item(ItemNum).Stackable > 0) Or (Item(ItemNum).Type = 16) Then
       Curr = True
    Else
        Curr = False
    End If
    
    For I = 1 To MAX_INV
        If Curr = True Then
            If GetPlayerInvItemNum(Index, I) = ItemNum Then
                Call SetPlayerInvItemValue(Index, I, GetPlayerInvItemValue(Index, I) + ItemVal)
                Call SendInventoryUpdate(Index, I)
                Has = True
                Exit For
            End If
        Else
            If GetPlayerInvItemNum(Index, I) = 0 Then
                Call SetPlayerInvItemNum(Index, I, ItemNum)
                Call SetPlayerInvItemValue(Index, I, ItemVal)
                If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Then
                    Call SetPlayerInvItemDur(Index, I, Item(ItemNum).Data1)
                End If
                Call SendInventoryUpdate(Index, I)
                Has = True
                Exit For
            End If
        End If
    Next I
    
    If Has = False Then
        For I = 1 To MAX_INV
            If (GetPlayerInvItemNum(Index, I) = 0) And (Curr = True) Then
                Call SetPlayerInvItemNum(Index, I, ItemNum)
                Call SetPlayerInvItemValue(Index, I, ItemVal)
                Call SendInventoryUpdate(Index, I)
                Has = True
                Exit For
            End If
        Next I
    End If
    
    If Has = False Then
        Call PlayerMsg(Index, "Tu inventario esta lleno, por favor vuelve cuando hagas un hueco.", BRIGHTRED)
        Exit Sub
    Else
        'Call WriteINI(Player(Index).Char(Player(Index).CharNum).Name, "QUEST" & Npc(npcnum).Quest, "1", App.Path + "\Main\Quest_Flags\qflag.ini")
        Call SetPlayerQuestFlag(Index, NPC(npcnum).Quest, 1)
        Call QuestMsg(Index, "" & Trim(NPC(npcnum).Name) & " dice: " & Trim(Quest(NPC(npcnum).Quest).Start) & "'", SayColor, 1)
        Call SendPlayerQuestFlags(Index)
    End If

End Sub
Sub GiveRewardItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal npcnum As Long)
Dim I As Long
Dim Curr As Boolean
Dim Has As Boolean
Dim questnum As Long
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
questnum = NPC(npcnum).Quest
    
    If GetPlayerQuestFlag(Index, NPC(npcnum).Quest) = 2 Then
        Call QuestMsg(Index, "" & Trim(NPC(npcnum).Name) & " dice: " & Trim(Quest(NPC(npcnum).Quest).After) & "'", SayColor, 1)
        Exit Sub
    End If
    
    If (Item(ItemNum).Stackable > 0) Or (Item(ItemNum).Type = 16) Then
       Curr = True
    Else
        Curr = False
    End If
    
    For I = 1 To MAX_INV
        If Curr = True Then
            If GetPlayerInvItemNum(Index, I) = ItemNum Then
                Call SetPlayerInvItemValue(Index, I, GetPlayerInvItemValue(Index, I) + ItemVal)
                Call SendInventoryUpdate(Index, I)
                Has = True
                'Call WriteINI(Player(Index).Char(Player(Index).CharNum).Name, "QUEST" & Npc(npcnum).Quest, "2", App.Path + "\Main\Quest_Flags\qflag.ini")
                Call SetPlayerQuestFlag(Index, NPC(npcnum).Quest, 2)
                Call SendPlayerQuestFlags(Index)
                Exit For
            End If
        Else
            If GetPlayerInvItemNum(Index, I) = 0 Then
                Call SetPlayerInvItemNum(Index, I, ItemNum)
                Call SetPlayerInvItemValue(Index, I, ItemVal)
                If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Then
                    Call SetPlayerInvItemDur(Index, I, Item(ItemNum).Data1)
                End If
                Call SendInventoryUpdate(Index, I)
                Has = True
                Exit For
            End If
        End If
    Next I
    
    If Has = False Then
        If Curr = True Then
            For I = 1 To MAX_INV
                If GetPlayerInvItemNum(Index, I) = 0 Then
                    Call SetPlayerInvItemNum(Index, I, ItemNum)
                    Call SetPlayerInvItemValue(Index, I, ItemVal)
                    Call SendInventoryUpdate(Index, I)
                    Has = True
                    ' Call WriteINI(Player(Index).Char(Player(Index).CharNum).Name, "QUEST" & Npc(npcnum).Quest, "2", App.Path + "\Main\Quest_Flags\qflag.ini")
                    Call SetPlayerQuestFlag(Index, NPC(npcnum).Quest, 2)
                    Call SendPlayerQuestFlags(Index)
                    Exit For
                End If
            Next I
        End If
    End If
    
    If Has = False Then
        Call PlayerMsg(Index, "Tu inventario esta lleno, por favor vuelve cuando hagas un hueco.", BRIGHTRED)
        Exit Sub
    Else
        Call QuestMsg(Index, "* Información de Misión *", SayColor, 1)
        Call QuestMsg(Index, "" & Trim(NPC(npcnum).Name) & " dice: " & Trim(Quest(NPC(npcnum).Quest).End) & "'", SayColor, 1)
        Call PlayerMsg(Index, "[Misión Completada] Has recibido tu recompensa.", YELLOW)
        Call SetPlayerQuestFlag(Index, NPC(npcnum).Quest, 2)
        Call SendPlayerQuestFlags(Index)
        Call DetermineExpType(questnum, Index)
        Call SendPlayerData(Index)
        Call CheckPlayerLevelUp(Index)
        If Item(Quest(NPC(npcnum).Quest).RewardNum).Type = 16 Or Item(Quest(NPC(npcnum).Quest).RewardNum).Stackable > 0 Then
            Call TakeItem(Index, Quest(NPC(npcnum).Quest).ItemReq, Quest(NPC(npcnum).Quest).ItemVal)
        Else
            Call TakeItem(Index, Quest(NPC(npcnum).Quest).ItemReq, Quest(NPC(npcnum).Quest).ItemVal)
        End If
                        Call SendInventoryUpdate(Index, I)
    End If

End Sub

Sub DetermineExpType(ByVal questnum As Long, ByVal Index As Long)
Dim ExpAmount As Long
Dim npcnum As Long
Dim I As Long

If Quest(questnum).QuestExpReward < 1 Then
Call PlayerMsg(Index, "Has ganado 0 de experiencia!", BRIGHTRED)
Exit Sub
End If

If Quest(questnum).QuestExpReward > 0 Then
Call SetPlayerExp(Index, GetPlayerExp(Index) + Quest(questnum).QuestExpReward)
Call CheckPlayerLevelUp(Index)
Call PlayerMsg(Index, "Has ganado " & Quest(questnum).QuestExpReward & " de experiencia.", BRIGHTRED)
Exit Sub
End If

Call SendStats(Index)
End Sub

Function DoQuestNpcKillsQuest(ByVal Index As Long, ByVal npcnum As Long)
Dim NpCSelect As Long
Dim NpCFinal As Long
Dim NpcAmount As Long
Dim NpcName As String
Dim NpcFinal2 As String
Dim EPool As Long



'If GetPlayerName(Index) = "Halo" Then
'Call SetPlayerNpcKillQuestFlag(Index, 0)
'Call SetPlayerNpcKillType(Index, "")
'Call SetPlayerNpcKillAmount(Index, 0)
'Call QuestMsg(Index, "reset just for you", Cyan,1)
'Exit Function
'End If

If GetPlayerNpcKillQuestFlag(Index) > 1 Then
If GetPlayerNpcKillAmount(Index) = 0 Then


EPool = 150000

Call SetPlayerNpcKillType(Index, "")
Call SetPlayerNpcKillQuestFlag(Index, 0)
Call SetPlayerExp(Index, GetPlayerExp(Index) + EPool)
Call QuestMsg(Index, "----Chaos Knight----", BRIGHTCYAN, 1)
Call QuestMsg(Index, "----Quest Completed----", BRIGHTGREEN, 1)
Call QuestMsg(Index, "You have Completed Your Chaos Knight Quest !", YELLOW, 1)
Call QuestMsg(Index, "You have gained " & EPool & " Experience Pool !", BRIGHTGREEN, 1)
Call SendPlayerData(Index)
Exit Function
End If
End If

If GetPlayerNpcKillQuestFlag(Index) > 0 Then
Call QuestMsg(Index, "----Chaos Knight----", BRIGHTGREEN, 3)
Call QuestMsg(Index, "You are already in a Quest " & GetPlayerName(Index) & " !", BRIGHTRED, 3)
If GetPlayerNpcKillType(Index) = "Ghost" Then
NpcName = "Ghost"
End If
If GetPlayerNpcKillType(Index) = "Firefly" Then
NpcName = "Firefly"
End If
If GetPlayerNpcKillType(Index) = "Evil Bat" Then
NpcName = "Evil Bat"
End If
If GetPlayerNpcKillType(Index) = "Demon Mage" Then
NpcName = "Demon Mage"
End If
Call QuestMsg(Index, "Return to Me When you have Killed " & GetPlayerNpcKillAmount(Index) & " " & NpcName & "(s) !", YELLOW, 3)
Exit Function
End If

If GetPlayerNpcKillQuestFlag(Index) < 1 Then
NpCSelect = (Rnd * 4)
NpcAmount = (Rnd * 50)

If NpCSelect = 1 Then
NpcFinal2 = "Ghost"
NpcName = "Ghost"
Call SendDataTo(Index, "killinfo" & SEP_CHAR & NpcAmount & SEP_CHAR & NpcName & SEP_CHAR & NpcFinal2 & SEP_CHAR & END_CHAR)
End If

If NpCSelect = 2 Then
NpcFinal2 = "Firefly"
NpcName = "Firefly"
Call SendDataTo(Index, "killinfo" & SEP_CHAR & NpcAmount & SEP_CHAR & NpcName & SEP_CHAR & NpcFinal2 & SEP_CHAR & END_CHAR)
End If

If NpCSelect = 3 Then
NpcFinal2 = "Evil Bat"
NpcName = "Evil Bat"
Call SendDataTo(Index, "killinfo" & SEP_CHAR & NpcAmount & SEP_CHAR & NpcName & SEP_CHAR & NpcFinal2 & SEP_CHAR & END_CHAR)
End If

If NpCSelect = 4 Then
NpcFinal2 = "Demon Mage"
NpcName = "Demon Mage"
Call SendDataTo(Index, "killinfo" & SEP_CHAR & NpcAmount & SEP_CHAR & NpcName & SEP_CHAR & NpcFinal2 & SEP_CHAR & END_CHAR)
End If

End If

End Function

Function ActuallyStartKillQuest(ByVal Index As Long, ByVal NpcAmount As Long, ByVal NpcName As String, ByVal NpcFinal2 As String)

Call SetPlayerNpcKillType(Index, NpcFinal2)
Call SetPlayerNpcKillAmount(Index, NpcAmount)
Call SetPlayerNpcKillQuestFlag(Index, 1)
Call QuestMsg(Index, "----Chaos Knight----", BRIGHTCYAN, 1)
Call QuestMsg(Index, "----Quest Received----", BRIGHTGREEN, 1)
Call QuestMsg(Index, "You have recieved a quest to kill " & NpcAmount & "  " & NpcName & "(s)- !", YELLOW, 1)

End Function

Function DoQuestNpcKills(ByVal Index As Long, ByVal npcnum As Long)
Dim NpCSelect As Long
Dim NpCFinal As Long
Dim NpcAmount As Long

If GetPlayerNpcKillQuestFlag(Index) > 1 Then Exit Function

If GetPlayerNpcKillQuestFlag(Index) > 0 Then
If Trim(NPC(npcnum).Name) = "Ghost" Then
If GetPlayerNpcKillAmount(Index) <= 1 Then
Call SetPlayerNpcKillQuestFlag(Index, 2)
Call SetPlayerNpcKillAmount(Index, 0)
Call PlayerMsg(Index, "Return to The Chaos Knight To Collect Your Reward !", YELLOW)
Exit Function
End If
End If
End If

If GetPlayerNpcKillQuestFlag(Index) > 0 Then
If Trim(NPC(npcnum).Name) = "Firefly" Then
If GetPlayerNpcKillAmount(Index) <= 1 Then
Call SetPlayerNpcKillQuestFlag(Index, 2)
Call SetPlayerNpcKillAmount(Index, 0)
Call PlayerMsg(Index, "Return to The Chaos Knight To Collect Your Reward !", YELLOW)
Exit Function
End If
End If
End If

If GetPlayerNpcKillQuestFlag(Index) > 0 Then
If Trim(NPC(npcnum).Name) = "Evil Bat" Then
If GetPlayerNpcKillAmount(Index) <= 1 Then
Call SetPlayerNpcKillQuestFlag(Index, 2)
Call SetPlayerNpcKillAmount(Index, 0)
Call PlayerMsg(Index, "Return to The Chaos Knight To Collect Your Reward !", YELLOW)
Exit Function
End If
End If
End If


If GetPlayerNpcKillQuestFlag(Index) > 0 Then
If Trim(NPC(npcnum).Name) = "Demon Mage" Then
If GetPlayerNpcKillAmount(Index) <= 1 Then
Call SetPlayerNpcKillQuestFlag(Index, 2)
Call SetPlayerNpcKillAmount(Index, 0)
Call PlayerMsg(Index, "Return to The Chaos Knight To Collect Your Reward !", YELLOW)
Exit Function
End If
End If
End If

If GetPlayerNpcKillQuestFlag(Index) > 0 Then
If Trim(NPC(npcnum).Name) = "Ghost" Then
If GetPlayerNpcKillAmount(Index) > 0 Then
If GetPlayerNpcKillType(Index) = "Ghost" Then
   Call SetPlayerNpcKillAmount(Index, GetPlayerNpcKillAmount(Index) - 1)
   Call PlayerMsg(Index, "You Have Killed A Monster For The Chaos Knight !", BRIGHTCYAN)
   Call PlayerMsg(Index, GetPlayerNpcKillAmount(Index) & " Monsters Remaining !", YELLOW)
   Exit Function
End If
End If
End If
End If

If GetPlayerNpcKillQuestFlag(Index) > 0 Then
If Trim(NPC(npcnum).Name) = "Firefly" Then
If GetPlayerNpcKillAmount(Index) > 0 Then
If GetPlayerNpcKillType(Index) = "Firefly" Then
Call SetPlayerNpcKillAmount(Index, GetPlayerNpcKillAmount(Index) - 1)
Call PlayerMsg(Index, "You Have Killed A Monster For The Chaos Knight !", BRIGHTCYAN)
Call PlayerMsg(Index, GetPlayerNpcKillAmount(Index) & " Monsters Remaining !", YELLOW)
Exit Function
End If
End If
End If
End If

If GetPlayerNpcKillQuestFlag(Index) > 0 Then
If Trim(NPC(npcnum).Name) = "Evil Bat" Then
If GetPlayerNpcKillAmount(Index) > 0 Then
If GetPlayerNpcKillType(Index) = "Evil Bat" Then
Call SetPlayerNpcKillAmount(Index, GetPlayerNpcKillAmount(Index) - 1)
Call PlayerMsg(Index, "You Have Killed A Monster For The Chaos Knight !", BRIGHTCYAN)
Call PlayerMsg(Index, GetPlayerNpcKillAmount(Index) & " Monsters Remaining !", YELLOW)
Exit Function
End If
End If
End If
End If

If GetPlayerNpcKillQuestFlag(Index) > 0 Then
If Trim(NPC(npcnum).Name) = "Demon Mage" Then
If GetPlayerNpcKillAmount(Index) > 0 Then
If GetPlayerNpcKillType(Index) = "Demon Mage" Then
Call SetPlayerNpcKillAmount(Index, GetPlayerNpcKillAmount(Index) - 1)
Call PlayerMsg(Index, "You Have Killed A Monster For The Chaos Knight !", BRIGHTCYAN)
Call PlayerMsg(Index, GetPlayerNpcKillAmount(Index) & " Monsters Remaining !", YELLOW)
Exit Function
End If
End If
End If
End If

End Function

Function StopKillQuest(ByVal Index As Long)
Call SetPlayerNpcKillType(Index, "")
Call SetPlayerNpcKillAmount(Index, 0)
Call SetPlayerNpcKillQuestFlag(Index, 0)
Call QuestMsg(Index, "----Chaos Knight----", BRIGHTCYAN, 1)
Call QuestMsg(Index, "----Quest Canceled----", BRIGHTGREEN, 1)
Call QuestMsg(Index, "You have Canceled Your Quest !", CYAN, 1)
Call QuestMsg(Index, "Perhaps you will Choose to Serve Again for The Kingdom, Speak with me again Later if Your Wish !", YELLOW, 1)
End Function
