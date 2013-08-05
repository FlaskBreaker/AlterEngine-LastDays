Attribute VB_Name = "modServerTCP"
Option Explicit
Sub SendPlayerNewXY(ByVal index As Long)
    Call SendDataTo(index, "PLAYERNEWXY" & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & END_CHAR)
End Sub

Sub UpdateTitle()
    frmServer.Caption = GAME_NAME & " - [AlterEngine | www.alterengine.net]"
End Sub

Sub UpdateTOP()
    frmServer.TPO.Caption = "Jugadores Conectados: " & TotalOnlinePlayers
End Sub

Function IsConnected(ByVal index As Long) As Boolean
    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
    Exit Function
End Function

Function IsPlaying(ByVal index As Long) As Boolean
    If index < 1 Or index > MAX_PLAYERS Then
        Exit Function
    End If

    If IsConnected(index) Then
        If Player(index).InGame Then
            IsPlaying = True
        End If
    End If
End Function

Function IsLoggedIn(ByVal index As Long) As Boolean
    If index < 1 Or index > MAX_PLAYERS Then
        Exit Function
    End If

    If IsConnected(index) Then
        If Trim$(Player(index).Login) <> vbNullString Then
            IsLoggedIn = True
        End If
    End If
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsConnected(i) Then
            If LCase$(Trim$(Player(i).Login)) = LCase$(Trim$(Login)) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If
    Next i
End Function

Function IsBanned(ByVal IPAddr As String) As Boolean
    Dim filename As String
    Dim FileIP As String
    Dim FileID As Long

    filename = App.Path & "\BanList.txt"

    FileID = FreeFile

    ' Check if file exists
    If Not FileExists("BanList.txt") Then
        Open filename For Output As #FileID
        Close #FileID
    End If

    Open filename For Input As #FileID
        Do While Not EOF(FileID)
            Line Input #FileID, FileIP
    
            If FileIP = IPAddr Then
                IsBanned = True
                Exit Do
            End If
        Loop
    Close #FileID
End Function

Sub SendDataTo(ByVal index As Long, ByVal Data As String)
Dim i As Long, N As Long, startc As Long

    If IsConnected(index) Then
        frmServer.Socket(index).SendData Data
        DoEvents
    End If
End Sub

Sub SendDataToAll(ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If
    Next i
End Sub

Sub SendDataToAllBut(ByVal index As Long, ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If i <> index Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
End Sub

Sub SendDataToMap(ByVal mapnum As Long, ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapnum Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal mapnum As Long, ByVal Data As String)
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapnum Then
                If i <> index Then
                    Call SendDataTo(i, Data)
                End If
            End If
        End If
    Next i
End Sub

Sub GlobalMsg(ByVal Msg As String, ByVal color As Byte)
    Call SendDataToAll("GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & color & END_CHAR)
End Sub

Sub Anuncio(ByVal Msg As String)
Dim Packet As String

    Packet = "anuncio" & SEP_CHAR & Msg & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal color As Byte)
    Call SendDataTo(index, "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & color & END_CHAR)
End Sub

Sub AdminMsg(ByVal Msg As String, ByVal color As Byte)
    Dim Packet As String
    Dim i As Long

    Packet = "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & color & END_CHAR

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerAccess(i) > 0 Then
                Call SendDataTo(i, Packet)
            End If
        End If
    Next i
End Sub

Sub MapMsg(ByVal mapnum As Long, ByVal Msg As String, ByVal color As Byte)
    Call SendDataToMap(mapnum, "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & color & END_CHAR)
End Sub

Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
    Call SendDataTo(index, "ALERTMSG" & SEP_CHAR & Msg & END_CHAR)
    Call CloseSocket(index)
End Sub

Sub PlainMsg(ByVal index As Long, ByVal Msg As String, ByVal num As Long)
    Call SendDataTo(index, "PLAINMSG" & SEP_CHAR & Msg & SEP_CHAR & num & END_CHAR)
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)
    If index < 1 Or index > MAX_PLAYERS Then
        Exit Sub
    End If

    If IsPlaying(index) Then
        Call AdminMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " ha sido hechado por (" & Reason & ")", BRIGHTRED)
        Call AlertMsg(index, "Has perdido la conexión con " & GAME_NAME & ".")
    End If
End Sub

Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
On Error Resume Next
Dim i As Long

    If (index = 0) Then
        i = FindOpenPlayerSlot
        
        If i <> 0 Then
            ' Whoho, we can connect them
            frmServer.Socket(i).Close
            frmServer.Socket(i).Accept SocketId
            Call SocketConnected(i)
        End If
    End If
End Sub

Sub SocketConnected(ByVal index As Long)
    If index < 1 Or index > MAX_PLAYERS Then
        Exit Sub
    End If

    If Not IsBanned(GetPlayerIP(index)) Then
        Call TextAdd(frmServer.txtText(0), "Conexión recibida de " & GetPlayerIP(index) & ".", True)
    Else
        Call AlertMsg(index, "Has sido baneado de " & GAME_NAME & ", y no podras volver a jugar.")
    End If
End Sub

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
Dim Buffer As String
Dim Packet As String
Dim Start As Long

    If index > 0 Then
        frmServer.Socket(index).GetData Buffer, vbString, DataLength
            
        Player(index).Buffer = Player(index).Buffer & Buffer
        
        Start = InStr(Player(index).Buffer, END_CHAR)
        Do While Start > 0
            Packet = Mid(Player(index).Buffer, 1, Start - 1)
            Player(index).Buffer = Mid(Player(index).Buffer, Start + 1, Len(Player(index).Buffer))
            Player(index).DataPackets = Player(index).DataPackets + 1
            Start = InStr(Player(index).Buffer, END_CHAR)
            If Len(Packet) > 0 Then
                Call HandleData(index, Packet)
            End If
        Loop
        
    ' Check if elapsed time has passed
     Player(index).DataBytes = Player(index).DataBytes + Len(Buffer)
    If GetTickCount >= Player(index).DataTimer + 1000 Then
        If Player(index).CharNum <> 0 Then
            Player(index).DataTimer = GetTickCount
            Player(index).DataBytes = 0
            Player(index).DataPackets = 0
        End If
    End If

    ' Check for data flooding
    If Player(index).DataBytes > 1000 Then
        If GetPlayerAccess(index) <= 0 Then
            Call HackingAttempt(index, "Data Flooding")
            Exit Sub
        End If
    End If

     'Check for packet flooding
    If Player(index).DataPackets > 100 Then
        If GetPlayerAccess(index) = 0 Then
            Call HackingAttempt(index, "Packet Flooding")
            Exit Sub
        End If
    End If
        
    End If
End Sub

Sub CloseSocket(ByVal index As Long)
    If index > 0 Then
        Call LeftGame(index)
        Call TextAdd(frmServer.txtText(0), "La conexion con " & GetPlayerIP(index) & " ha sido terminada.", True)

        frmServer.Socket(index).Close
       
        Call UpdateTOP
        Call ClearPlayer(index)
    End If
End Sub

Sub SendWhosOnline(ByVal index As Long)
    Dim PlayerNames As String
    Dim PlayerCount As Long
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If i <> index Then
                PlayerNames = PlayerNames & GetPlayerName(i) & ", "
                PlayerCount = PlayerCount + 1
            End If
        End If
    Next i

    If PlayerCount = 0 Then
        PlayerNames = "No hay otros jugadores conectados."
    Else
        PlayerNames = Mid$(PlayerNames, 1, Len(PlayerNames) - 2)
        PlayerNames = "Hay otros " & PlayerCount & " jugadores conectados: " & PlayerNames & "."
    End If

    Call PlayerMsg(index, PlayerNames, WhoColor)
End Sub

Sub SendOnlineList()
    Dim Packet As String
    Dim PlayerCount As Long
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Packet = Packet & SEP_CHAR & GetPlayerName(i) & SEP_CHAR
            PlayerCount = PlayerCount + 1
        End If
    Next i

    Call SendDataToAll("ONLINELIST" & SEP_CHAR & PlayerCount & Packet & END_CHAR)
End Sub

Sub SendChars(ByVal index As Long)
    Dim Packet As String
    Dim i As Long

    Packet = "ALLCHARS" & SEP_CHAR
    For i = 1 To MAX_CHARS
        Packet = Packet & Trim$(Player(index).Char(i).Name) & SEP_CHAR & Trim$(ClassData(Player(index).Char(i).Class).Name) & SEP_CHAR & Player(index).Char(i).Level & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR

    Call SendDataTo(index, Packet)
End Sub

Sub SendJoinMap(ByVal index As Long)
    Dim Packet As String
    Dim i As Long
    Dim j As Long

    Packet = vbNullString

    ' Send all players on current map to index
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            'If I <> index Then
                If GetPlayerMap(i) = GetPlayerMap(index) Then
                    Packet = "PLAYERDATA" & SEP_CHAR
                    Packet = Packet & i & SEP_CHAR
                    Packet = Packet & GetPlayerName(i) & SEP_CHAR
                    Packet = Packet & GetPlayerSprite(i) & SEP_CHAR
                    Packet = Packet & GetPlayerMap(i) & SEP_CHAR
                    Packet = Packet & GetPlayerX(i) & SEP_CHAR
                    Packet = Packet & GetPlayerY(i) & SEP_CHAR
                    Packet = Packet & GetPlayerDir(i) & SEP_CHAR
                    Packet = Packet & GetPlayerAccess(i) & SEP_CHAR
                    Packet = Packet & GetPlayerPK(i) & SEP_CHAR
                    Packet = Packet & GetPlayerGuild(i) & SEP_CHAR
                    Packet = Packet & GetPlayerGuildAccess(i) & SEP_CHAR
                    Packet = Packet & GetPlayerClass(i) & SEP_CHAR
                    Packet = Packet & GetPlayerHead(i) & SEP_CHAR
                    Packet = Packet & GetPlayerBody(i) & SEP_CHAR
                    Packet = Packet & GetPlayerleg(i) & SEP_CHAR
                    Packet = Packet & GetPlayerPaperdoll(i) & SEP_CHAR
                    Packet = Packet & GetPlayerLevel(i) & SEP_CHAR
                    Packet = Packet & Player(index).Char(Player(index).CharNum).color(1) & SEP_CHAR
                    Packet = Packet & Player(index).Char(Player(index).CharNum).color(2) & SEP_CHAR
                    Packet = Packet & Player(index).Char(Player(index).CharNum).color(3) & SEP_CHAR
                    Packet = Packet & END_CHAR
                    Call SendDataTo(index, Packet)
                    
                    If Player(i).Pet.Alive = YES Then
                Packet = "PETDATA" & SEP_CHAR
                Packet = Packet & Player(i).Pet.Alive & SEP_CHAR
                Packet = Packet & Player(i).Pet.Map & SEP_CHAR
                Packet = Packet & Player(i).Pet.x & SEP_CHAR
                Packet = Packet & Player(i).Pet.y & SEP_CHAR
                Packet = Packet & Player(i).Pet.Dir & SEP_CHAR
                Packet = Packet & Player(i).Pet.SPRITE & SEP_CHAR
                Packet = Packet & Player(i).Pet.HP & SEP_CHAR
                Packet = Packet & Player(i).Pet.STR * 5 & SEP_CHAR
                Packet = Packet & Player(i).Pet.STR & SEP_CHAR
                Packet = Packet & Player(i).Pet.DEF & SEP_CHAR
                Packet = Packet & Player(i).Pet.Speed & SEP_CHAR
                Packet = Packet & Player(i).Pet.Magi & SEP_CHAR
                Packet = Packet & Player(i).Pet.Level & SEP_CHAR
                Packet = Packet & Player(i).Pet.POINTS & SEP_CHAR
                Packet = Packet & GetPetSP(i) & SEP_CHAR
                Packet = Packet & GetPetMAXSP(i) & SEP_CHAR
                Packet = Packet & GetPetMP(i) & SEP_CHAR
                Packet = Packet & GetPetMAXMP(i) & SEP_CHAR
                Packet = Packet & GetPetFP(i) & SEP_CHAR
                Packet = Packet & GetPetMAXFP(i) & SEP_CHAR
                Packet = Packet & GetPetExp(i) & SEP_CHAR
                Packet = Packet & GetPetNextLevel(i) & SEP_CHAR
                Packet = Packet & Player(i).Pet.Name & SEP_CHAR
                Packet = Packet & END_CHAR
                Call SendDataTo(index, Packet)
            End If
                End If
            'End If
        End If
    Next i

    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal mapnum As Long)
    Dim Packet As String

    Packet = "leave" & SEP_CHAR & index & END_CHAR
    Call SendDataToMapBut(index, mapnum, Packet)
    
    If Player(index).Pet.Alive = YES Then
        Packet = "PETDATA" & SEP_CHAR
        Packet = Packet & index & SEP_CHAR
        Packet = Packet & Player(index).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(index).Pet.Map & SEP_CHAR
        Packet = Packet & Player(index).Pet.x & SEP_CHAR
        Packet = Packet & Player(index).Pet.y & SEP_CHAR
        Packet = Packet & Player(index).Pet.Dir & SEP_CHAR
        Packet = Packet & Player(index).Pet.SPRITE & SEP_CHAR
        Packet = Packet & Player(index).Pet.HP & SEP_CHAR
        Packet = Packet & Player(index).Pet.STR * 5 & SEP_CHAR
        Packet = Packet & Player(index).Pet.STR & SEP_CHAR
        Packet = Packet & Player(index).Pet.DEF & SEP_CHAR
        Packet = Packet & Player(index).Pet.Speed & SEP_CHAR
        Packet = Packet & Player(index).Pet.Magi & SEP_CHAR
        Packet = Packet & Player(index).Pet.Level & SEP_CHAR
        Packet = Packet & Player(index).Pet.POINTS & SEP_CHAR
        Packet = Packet & GetPetSP(index) & SEP_CHAR
        Packet = Packet & GetPetMAXSP(index) & SEP_CHAR
        Packet = Packet & GetPetMP(index) & SEP_CHAR
        Packet = Packet & GetPetMAXMP(index) & SEP_CHAR
        Packet = Packet & GetPetFP(index) & SEP_CHAR
        Packet = Packet & GetPetMAXFP(index) & SEP_CHAR
        Packet = Packet & GetPetExp(index) & SEP_CHAR
        Packet = Packet & GetPetNextLevel(index) & SEP_CHAR
        Packet = Packet & Player(index).Pet.Name & SEP_CHAR
        Packet = Packet & END_CHAR
        Call SendDataToMapBut(index, mapnum, Packet)
    End If
End Sub

Sub SendPlayerData(ByVal index As Long)
    Dim Packet As String
    Dim j As Long

    ' Send index's player data to everyone including himself on th emap
    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & index & SEP_CHAR
    Packet = Packet & GetPlayerName(index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(index) & SEP_CHAR
    Packet = Packet & GetPlayerMap(index) & SEP_CHAR
    Packet = Packet & GetPlayerX(index) & SEP_CHAR
    Packet = Packet & GetPlayerY(index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(index) & SEP_CHAR
    Packet = Packet & GetPlayerHead(index) & SEP_CHAR
    Packet = Packet & GetPlayerBody(index) & SEP_CHAR
    Packet = Packet & GetPlayerleg(index) & SEP_CHAR
    Packet = Packet & GetPlayerPaperdoll(index) & SEP_CHAR
    Packet = Packet & GetPlayerLevel(index) & SEP_CHAR
    Packet = Packet & Player(index).Char(Player(index).CharNum).color(1) & SEP_CHAR
    Packet = Packet & Player(index).Char(Player(index).CharNum).color(2) & SEP_CHAR
    Packet = Packet & Player(index).Char(Player(index).CharNum).color(3) & SEP_CHAR

    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
    
    If Player(index).Pet.Alive = YES Then
        Packet = "PETDATA" & SEP_CHAR
        Packet = Packet & index & SEP_CHAR
        Packet = Packet & Player(index).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(index).Pet.Map & SEP_CHAR
        Packet = Packet & Player(index).Pet.x & SEP_CHAR
        Packet = Packet & Player(index).Pet.y & SEP_CHAR
        Packet = Packet & Player(index).Pet.Dir & SEP_CHAR
        Packet = Packet & Player(index).Pet.SPRITE & SEP_CHAR
        Packet = Packet & Player(index).Pet.HP & SEP_CHAR
        Packet = Packet & Player(index).Pet.STR * 5 & SEP_CHAR
        Packet = Packet & Player(index).Pet.STR & SEP_CHAR
        Packet = Packet & Player(index).Pet.DEF & SEP_CHAR
        Packet = Packet & Player(index).Pet.Speed & SEP_CHAR
        Packet = Packet & Player(index).Pet.Magi & SEP_CHAR
        Packet = Packet & Player(index).Pet.Level & SEP_CHAR
        Packet = Packet & Player(index).Pet.POINTS & SEP_CHAR
        Packet = Packet & GetPetSP(index) & SEP_CHAR
        Packet = Packet & GetPetMAXSP(index) & SEP_CHAR
        Packet = Packet & GetPetMP(index) & SEP_CHAR
        Packet = Packet & GetPetMAXMP(index) & SEP_CHAR
        Packet = Packet & GetPetFP(index) & SEP_CHAR
        Packet = Packet & GetPetMAXFP(index) & SEP_CHAR
        Packet = Packet & GetPetExp(index) & SEP_CHAR
        Packet = Packet & GetPetNextLevel(index) & SEP_CHAR
        Packet = Packet & Player(index).Pet.Name & SEP_CHAR
        Packet = Packet & GetPlayerLevel(index) & SEP_CHAR
        Packet = Packet & END_CHAR
        Call SendDataToMap(GetPlayerMap(index), Packet)
    End If
End Sub

Public Sub SendMap(ByVal index As Long, ByVal mapnum As Long)
    If LenB(MapCache(mapnum)) = 0 Then
        Call MapCache_Create(mapnum)
    End If

    Call SendDataTo(index, MapCache(mapnum))
End Sub

Public Sub MapCache_Create(ByVal mapnum As Long)
    Dim MapData As String
    Dim x As Long
    Dim y As Long

    MapData = "MAPDATA" & SEP_CHAR & mapnum & SEP_CHAR & Trim$(Map(mapnum).Name) & SEP_CHAR & Map(mapnum).Revision & SEP_CHAR & Map(mapnum).Moral & SEP_CHAR & Map(mapnum).Up & SEP_CHAR & Map(mapnum).Down & SEP_CHAR & Map(mapnum).Left & SEP_CHAR & Map(mapnum).Right & SEP_CHAR & Map(mapnum).music & SEP_CHAR & Map(mapnum).BootMap & SEP_CHAR & Map(mapnum).BootX & SEP_CHAR & Map(mapnum).BootY & SEP_CHAR & Map(mapnum).Indoors & SEP_CHAR & Map(mapnum).Weather & SEP_CHAR

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            With Map(mapnum).Tile(x, y)
                MapData = MapData & (.Ground & SEP_CHAR & .Mask & SEP_CHAR & .Anim & SEP_CHAR & .Mask2 & SEP_CHAR & .M2Anim & SEP_CHAR & .Fringe & SEP_CHAR & .FAnim & SEP_CHAR & .Fringe2 & SEP_CHAR & .F2Anim & SEP_CHAR & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR & .Light & SEP_CHAR)
                MapData = MapData & (.GroundSet & SEP_CHAR & .MaskSet & SEP_CHAR & .AnimSet & SEP_CHAR & .Mask2Set & SEP_CHAR & .M2AnimSet & SEP_CHAR & .FringeSet & SEP_CHAR & .FAnimSet & SEP_CHAR & .Fringe2Set & SEP_CHAR & .F2AnimSet & SEP_CHAR)
            End With
        Next x
    Next y

    For x = 1 To MAX_MAP_NPCS
        MapData = MapData & (Map(mapnum).NPC(x) & SEP_CHAR & Map(mapnum).SpawnX(x) & SEP_CHAR & Map(mapnum).SpawnY(x) & SEP_CHAR)
    Next x

    MapData = MapData & END_CHAR

    MapCache(mapnum) = MapData
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal mapnum As Long)
    Dim Packet As String
    Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        If mapnum > 0 Then
            Packet = Packet & (MapItem(mapnum, i).num & SEP_CHAR & MapItem(mapnum, i).Value & SEP_CHAR & MapItem(mapnum, i).Dur & SEP_CHAR & MapItem(mapnum, i).x & SEP_CHAR & MapItem(mapnum, i).y & SEP_CHAR)
        End If
    Next i
    Packet = Packet & END_CHAR

    Call SendDataTo(index, Packet)
End Sub

Sub SendMapItemsToAll(ByVal mapnum As Long)
    Dim Packet As String
    Dim i As Long

    Packet = "MAPITEMDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_ITEMS
        Packet = Packet & (MapItem(mapnum, i).num & SEP_CHAR & MapItem(mapnum, i).Value & SEP_CHAR & MapItem(mapnum, i).Dur & SEP_CHAR & MapItem(mapnum, i).x & SEP_CHAR & MapItem(mapnum, i).y & SEP_CHAR)
    Next i
    Packet = Packet & END_CHAR

    Call SendDataToMap(mapnum, Packet)
End Sub

Sub SendMapNpcsTo(ByVal index As Long, ByVal mapnum As Long)
    Dim Packet As String
    Dim i As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS
        If mapnum > 0 Then
            Packet = Packet & (MapNPC(mapnum, i).num & SEP_CHAR & MapNPC(mapnum, i).x & SEP_CHAR & MapNPC(mapnum, i).y & SEP_CHAR & MapNPC(mapnum, i).Dir & SEP_CHAR)
        End If
    Next i
    Packet = Packet & END_CHAR

    Call SendDataTo(index, Packet)
End Sub

Sub SendMapNpcsToMap(ByVal mapnum As Long)
    Dim Packet As String
    Dim i As Long

    Packet = "MAPNPCDATA" & SEP_CHAR
    For i = 1 To MAX_MAP_NPCS
        Packet = Packet & (MapNPC(mapnum, i).num & SEP_CHAR & MapNPC(mapnum, i).x & SEP_CHAR & MapNPC(mapnum, i).y & SEP_CHAR & MapNPC(mapnum, i).Dir & SEP_CHAR)
    Next i
    Packet = Packet & END_CHAR

    Call SendDataToMap(mapnum, Packet)
End Sub

Sub SendItems(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_ITEMS
        If Trim$(Item(i).Name) <> vbNullString Then
            Call SendUpdateItemTo(index, i)
        End If
    Next i
End Sub

Sub SendElements(ByVal index As Long)
    Dim i As Long

    For i = 0 To MAX_ELEMENTS
        If Trim$(Element(i).Name) <> vbNullString Then
            Call SendUpdateElementTo(index, i)
        End If
    Next i
End Sub
Sub SendEmoticons(ByVal index As Long)
    Dim i As Long

    For i = 0 To MAX_EMOTICONS
        If Trim$(Emoticons(i).Command) <> vbNullString Then
            Call SendUpdateEmoticonTo(index, i)
        End If
    Next i
End Sub

Sub SendArrows(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_ARROWS
        If Trim$(Arrows(i).Name) <> vbNullString Then
            Call SendUpdateArrowTo(index, i)
        End If
    Next i
End Sub

Sub SendNpcs(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_NPCS
        If Trim$(NPC(i).Name) <> vbNullString Then
            Call SendUpdateNpcTo(index, i)
        End If
    Next i
End Sub
Sub SendBank(ByVal index As Long)
    Dim Packet As String
    Dim i As Integer

    Packet = "PLAYERBANK" & SEP_CHAR
    For i = 1 To MAX_BANK
        Packet = Packet & (GetPlayerBankItemNum(index, i) & SEP_CHAR & GetPlayerBankItemValue(index, i) & SEP_CHAR & GetPlayerBankItemDur(index, i) & SEP_CHAR)
    Next i
    Packet = Packet & END_CHAR

    Call SendDataTo(index, Packet)
End Sub

Sub SendBankUpdate(ByVal index As Long, ByVal BankSlot As Long)
    Call SendDataTo(index, "PLAYERBANKUPDATE" & SEP_CHAR & BankSlot & SEP_CHAR & GetPlayerBankItemNum(index, BankSlot) & SEP_CHAR & GetPlayerBankItemValue(index, BankSlot) & SEP_CHAR & GetPlayerBankItemDur(index, BankSlot) & END_CHAR)
End Sub
Sub SendInventory(ByVal index As Long)
    Dim Packet As String
    Dim i As Long

    Packet = "PLAYERINV" & SEP_CHAR & index & SEP_CHAR
    For i = 1 To MAX_INV
        Packet = Packet & (GetPlayerInvItemNum(index, i) & SEP_CHAR & GetPlayerInvItemValue(index, i) & SEP_CHAR & GetPlayerInvItemDur(index, i) & SEP_CHAR)
    Next i
    Packet = Packet & END_CHAR

    Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal InvSlot As Long)
    Call SendDataToMap(GetPlayerMap(index), "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & index & SEP_CHAR & GetPlayerInvItemNum(index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(index, InvSlot) & SEP_CHAR & index & END_CHAR)
End Sub

Sub SendIndexInventoryFromMap(ByVal index As Long)
    Dim Packet As String
    Dim N As Long
    Dim i As Long
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(index) Then
                Packet = "PLAYERINV" & SEP_CHAR & i & SEP_CHAR
                For N = 1 To MAX_INV
                    Packet = Packet & (GetPlayerInvItemNum(i, N) & SEP_CHAR & GetPlayerInvItemValue(i, N) & SEP_CHAR & GetPlayerInvItemDur(i, N) & SEP_CHAR)
                Next N
                Packet = Packet & END_CHAR

                Call SendDataTo(index, Packet)
            End If
        End If
    Next i
End Sub

Sub SendWornEquipment(ByVal index As Long)
    Dim Packet As String
    If IsPlaying(index) Then
        Packet = "PLAYERWORNEQ" & SEP_CHAR & index & SEP_CHAR & Player(index).Char(Player(index).CharNum).ArmorSlot & SEP_CHAR & Player(index).Char(Player(index).CharNum).WeaponSlot & SEP_CHAR & Player(index).Char(Player(index).CharNum).HelmetSlot & SEP_CHAR & Player(index).Char(Player(index).CharNum).ShieldSlot & SEP_CHAR & Player(index).Char(Player(index).CharNum).LegsSlot & SEP_CHAR & Player(index).Char(Player(index).CharNum).RingSlot & SEP_CHAR & Player(index).Char(Player(index).CharNum).NecklaceSlot & END_CHAR
        Call SendDataToMap(GetPlayerMap(index), Packet)
    End If
End Sub

Sub GetMapWornEquipment(ByVal index As Long)
    Dim Packet As String
    Dim i As Long
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If Player(i).Char(Player(i).CharNum).Map = Player(index).Char(Player(index).CharNum).Map Then
                Packet = "PLAYERWORNEQ" & SEP_CHAR & i & SEP_CHAR & Player(i).Char(Player(i).CharNum).ArmorSlot & SEP_CHAR & Player(i).Char(Player(i).CharNum).WeaponSlot & SEP_CHAR & Player(i).Char(Player(i).CharNum).HelmetSlot & SEP_CHAR & Player(i).Char(Player(i).CharNum).ShieldSlot & SEP_CHAR & Player(i).Char(Player(i).CharNum).LegsSlot & SEP_CHAR & Player(i).Char(Player(i).CharNum).RingSlot & SEP_CHAR & Player(i).Char(Player(i).CharNum).NecklaceSlot & END_CHAR
                Call SendDataTo(index, Packet)
            End If
        End If
    Next i
End Sub

Sub SendHP(ByVal index As Long)
    Call SendDataToMap(GetPlayerMap(index), "PLAYERHP" & SEP_CHAR & index & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & GetPlayerHP(index) & END_CHAR)
End Sub

Sub SendMP(ByVal index As Long)
    Call SendDataTo(index, "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(index) & SEP_CHAR & GetPlayerMP(index) & END_CHAR)
End Sub

Sub SendSP(ByVal index As Long)
    Call SendDataTo(index, "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(index) & SEP_CHAR & GetPlayerSP(index) & END_CHAR)
End Sub

Sub SendPTS(ByVal index As Long)
    Call SendDataTo(index, "PLAYERPOINTS" & SEP_CHAR & GetPlayerPOINTS(index) & END_CHAR)
End Sub

Sub SendEXP(ByVal index As Long)
    Call SendDataTo(index, "PLAYEREXP" & SEP_CHAR & GetPlayerExp(index) & SEP_CHAR & GetPlayerNextLevel(index) & END_CHAR)
End Sub

Sub SendStats(ByVal index As Long)
    Call SendDataTo(index, "PLAYERSTATSPACKET" & SEP_CHAR & GetPlayerSTR(index) & SEP_CHAR & GetPlayerDEF(index) & SEP_CHAR & GetPlayerSPEED(index) & SEP_CHAR & GetPlayerMAGI(index) & SEP_CHAR & GetPlayerNextLevel(index) & SEP_CHAR & GetPlayerExp(index) & SEP_CHAR & GetPlayerLevel(index) & END_CHAR)
End Sub

Sub SendPlayerLevelToAll(ByVal index As Long)
    Call SendDataToAll("PLAYERLEVEL" & SEP_CHAR & index & SEP_CHAR & GetPlayerLevel(index) & END_CHAR)
End Sub

Sub SendClasses(ByVal index As Long)
    Dim Packet As String
    Dim i As Long

    Packet = "CLASSESDATA" & SEP_CHAR & MAX_CLASSES & SEP_CHAR
    For i = 0 To MAX_CLASSES
        Packet = Packet & (GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & ClassData(i).STR & SEP_CHAR & ClassData(i).DEF & SEP_CHAR & ClassData(i).Speed & SEP_CHAR & ClassData(i).Magi & SEP_CHAR & ClassData(i).Locked & SEP_CHAR & ClassData(i).Desc & SEP_CHAR)
    Next i
    Packet = Packet & END_CHAR

    Call SendDataTo(index, Packet)
End Sub

Sub SendNewCharClasses(ByVal index As Long)
    Dim Packet As String
    Dim i As Long
    Dim Gender As Long
    Dim Gender1 As String
    Dim Gender2 As String
    Dim stat1 As String
    Dim stat2 As String
    Dim stat3 As String
    Dim stat4 As String
    

    Packet = "NEWCHARCLASSES" & SEP_CHAR & MAX_CLASSES & SEP_CHAR & CLASSES & SEP_CHAR
    For i = 0 To MAX_CLASSES
        Gender = ClassData(i).Gender
        Gender1 = ClassData(i).Gender1
        Gender2 = ClassData(i).Gender2
        stat1 = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "stat1")
        stat2 = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "stat2")
        stat3 = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "stat3")
        stat4 = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "stat4")
        Packet = Packet & (GetClassName(i) & SEP_CHAR & GetClassMaxHP(i) & SEP_CHAR & GetClassMaxMP(i) & SEP_CHAR & GetClassMaxSP(i) & SEP_CHAR & ClassData(i).STR & SEP_CHAR & ClassData(i).DEF & SEP_CHAR & ClassData(i).Speed & SEP_CHAR & ClassData(i).Magi & SEP_CHAR & ClassData(i).MaleSprite & SEP_CHAR & ClassData(i).FemaleSprite & SEP_CHAR & ClassData(i).Locked & SEP_CHAR & ClassData(i).Desc & SEP_CHAR & Gender & SEP_CHAR & Gender1 & SEP_CHAR & Gender2 & SEP_CHAR)
    Next i
    Packet = Packet & stat1 & SEP_CHAR & stat2 & SEP_CHAR & stat3 & SEP_CHAR & stat4 & SEP_CHAR
    Packet = Packet & END_CHAR

    Call SendDataTo(index, Packet)
End Sub

Sub SendLeftGame(ByVal index As Long)
 Dim Packet As String
    Call SendDataToAllBut(index, "left" & SEP_CHAR & index & END_CHAR)
    
    If Player(index).Pet.Alive <> NO Then
    Packet = "PETDATA" & SEP_CHAR
    Packet = Packet & index & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & 0 & SEP_CHAR
    Packet = Packet & "" & SEP_CHAR
    Packet = Packet & END_CHAR
    End If
    Call SendDataToAllBut(index, Packet)
End Sub

Sub SendPlayerXY(ByVal index As Long)
    Call SendDataToMap(GetPlayerMap(index), "PLAYERXY" & SEP_CHAR & index & SEP_CHAR & GetPlayerX(index) & SEP_CHAR & GetPlayerY(index) & END_CHAR)
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Long)
    Dim Packet As String

    ' Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & END_CHAR
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).addHP & SEP_CHAR & Item(ItemNum).addMP & SEP_CHAR & Item(ItemNum).addSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound
    Packet = Packet & END_CHAR
    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal ItemNum As Long)
    Dim Packet As String

    ' Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Desc & END_CHAR
    Packet = "UPDATEITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).addHP & SEP_CHAR & Item(ItemNum).addMP & SEP_CHAR & Item(ItemNum).addSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound
    Packet = Packet & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendEditItemTo(ByVal index As Long, ByVal ItemNum As Long)
    Dim Packet As String

    Packet = "EDITITEM" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Item(ItemNum).Name) & SEP_CHAR & Item(ItemNum).Pic & SEP_CHAR & Item(ItemNum).Type & SEP_CHAR & Item(ItemNum).Data1 & SEP_CHAR & Item(ItemNum).Data2 & SEP_CHAR & Item(ItemNum).Data3 & SEP_CHAR & Item(ItemNum).StrReq & SEP_CHAR & Item(ItemNum).DefReq & SEP_CHAR & Item(ItemNum).SpeedReq & SEP_CHAR & Item(ItemNum).MagicReq & SEP_CHAR & Item(ItemNum).ClassReq & SEP_CHAR & Item(ItemNum).AccessReq & SEP_CHAR
    Packet = Packet & Item(ItemNum).addHP & SEP_CHAR & Item(ItemNum).addMP & SEP_CHAR & Item(ItemNum).addSP & SEP_CHAR & Item(ItemNum).AddStr & SEP_CHAR & Item(ItemNum).AddDef & SEP_CHAR & Item(ItemNum).AddMagi & SEP_CHAR & Item(ItemNum).AddSpeed & SEP_CHAR & Item(ItemNum).AddEXP & SEP_CHAR & Item(ItemNum).Desc & SEP_CHAR & Item(ItemNum).AttackSpeed & SEP_CHAR & Item(ItemNum).Price & SEP_CHAR & Item(ItemNum).Stackable & SEP_CHAR & Item(ItemNum).Bound
    Packet = Packet & END_CHAR
    Call SendDataTo(index, Packet)
End Sub

Sub SendUpdateEmoticonToAll(ByVal ItemNum As Long)
    Call SendDataToAll("UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & END_CHAR)
End Sub

Sub SendUpdateEmoticonTo(ByVal index As Long, ByVal ItemNum As Long)
    Call SendDataTo(index, "UPDATEEMOTICON" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Emoticons(ItemNum).Command) & SEP_CHAR & Emoticons(ItemNum).Pic & END_CHAR)
End Sub

Sub SendEditEmoticonTo(ByVal index As Long, ByVal EmoNum As Long)
    Call SendDataTo(index, "EDITEMOTICON" & SEP_CHAR & EmoNum & SEP_CHAR & Trim$(Emoticons(EmoNum).Command) & SEP_CHAR & Emoticons(EmoNum).Pic & END_CHAR)
End Sub

Sub SendUpdateElementToAll(ByVal ElementNum As Long)
    Call SendDataToAll("UPDATEELEMENT" & SEP_CHAR & ElementNum & SEP_CHAR & Trim$(Element(ElementNum).Name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & END_CHAR)
End Sub

Sub SendUpdateElementTo(ByVal index As Long, ByVal ElementNum As Long)
    Call SendDataTo(index, "UPDATEELEMENT" & SEP_CHAR & ElementNum & SEP_CHAR & Trim$(Element(ElementNum).Name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & END_CHAR)
End Sub

Sub SendEditElementTo(ByVal index As Long, ByVal ElementNum As Long)
    Call SendDataTo(index, "EDITELEMENT" & SEP_CHAR & ElementNum & SEP_CHAR & Trim$(Element(ElementNum).Name) & SEP_CHAR & Element(ElementNum).Strong & SEP_CHAR & Element(ElementNum).Weak & END_CHAR)
End Sub

Sub SendUpdateArrowToAll(ByVal ItemNum As Long)
    Call SendDataToAll("UPDATEARROW" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & Arrows(ItemNum).Amount & END_CHAR)
End Sub

Sub SendUpdateArrowTo(ByVal index As Long, ByVal ItemNum As Long)
    Call SendDataTo(index, "UPDATEARROW" & SEP_CHAR & ItemNum & SEP_CHAR & Trim$(Arrows(ItemNum).Name) & SEP_CHAR & Arrows(ItemNum).Pic & SEP_CHAR & Arrows(ItemNum).Range & SEP_CHAR & Arrows(ItemNum).Amount & END_CHAR)
End Sub

Sub SendEditArrowTo(ByVal index As Long, ByVal EmoNum As Long)
    Call SendDataTo(index, "EDITARROW" & SEP_CHAR & EmoNum & SEP_CHAR & Trim$(Arrows(EmoNum).Name) & END_CHAR)
End Sub

Sub SendUpdateNpcToAll(ByVal npcnum As Long)
    Call SendDataToAll("UPDATENPC" & SEP_CHAR & npcnum & SEP_CHAR & Trim$(NPC(npcnum).Name) & SEP_CHAR & NPC(npcnum).SPRITE & SEP_CHAR & NPC(npcnum).SPRITESIZE & SEP_CHAR & NPC(npcnum).Big & SEP_CHAR & NPC(npcnum).MAXHP & SEP_CHAR & NPC(npcnum).Quest & END_CHAR)
End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal npcnum As Long)
    Call SendDataTo(index, "UPDATENPC" & SEP_CHAR & npcnum & SEP_CHAR & Trim$(NPC(npcnum).Name) & SEP_CHAR & NPC(npcnum).SPRITE & SEP_CHAR & NPC(npcnum).SPRITESIZE & SEP_CHAR & NPC(npcnum).Big & SEP_CHAR & NPC(npcnum).MAXHP & SEP_CHAR & NPC(npcnum).Quest & END_CHAR)
End Sub

Sub SendEditNpcTo(ByVal index As Long, ByVal npcnum As Long)
    Dim Packet As String
    Dim i As Long

    Packet = "EDITNPC" & SEP_CHAR & npcnum & SEP_CHAR & Trim$(NPC(npcnum).Name) & SEP_CHAR & Trim$(NPC(npcnum).AttackSay) & SEP_CHAR & NPC(npcnum).SPRITE & SEP_CHAR & NPC(npcnum).SpawnSecs & SEP_CHAR & NPC(npcnum).Behavior & SEP_CHAR & NPC(npcnum).Range & SEP_CHAR & NPC(npcnum).STR & SEP_CHAR & NPC(npcnum).DEF & SEP_CHAR & NPC(npcnum).Speed & SEP_CHAR & NPC(npcnum).Magi & SEP_CHAR & NPC(npcnum).Big & SEP_CHAR & NPC(npcnum).MAXHP & SEP_CHAR & NPC(npcnum).EXP & SEP_CHAR & NPC(npcnum).SpawnTime & SEP_CHAR & NPC(npcnum).Element & SEP_CHAR & NPC(npcnum).SPRITESIZE & SEP_CHAR & NPC(npcnum).Quest & SEP_CHAR
    For i = 1 To MAX_NPC_DROPS
        Packet = Packet & (NPC(npcnum).ItemNPC(i).chance & SEP_CHAR & NPC(npcnum).ItemNPC(i).ItemNum & SEP_CHAR & NPC(npcnum).ItemNPC(i).ItemValue & SEP_CHAR)
    Next i
    If NPC(npcnum).standstill = False Or NPC(npcnum).standstill = True Then
    Packet = Packet & SEP_CHAR & NPC(npcnum).standstill & END_CHAR
    Else
    NPC(npcnum).standstill = False
    Packet = Packet & SEP_CHAR & NPC(npcnum).standstill & END_CHAR
    End If
    
    Call SendDataTo(index, Packet)
End Sub

Sub SendShops(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_SHOPS
        If Trim$(Shop(i).Name) <> vbNullString Then
            Call SendUpdateShopTo(index, i)
        End If
    Next i
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
    Dim Packet As String
    Dim i As Integer

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Shop(ShopNum).BuysItems & SEP_CHAR & Shop(ShopNum).ShowInfo & SEP_CHAR & Shop(ShopNum).CurrencyItem & SEP_CHAR
    For i = 1 To MAX_SHOP_ITEMS
        Packet = Packet & (Shop(ShopNum).ShopItem(i).ItemNum & SEP_CHAR & Shop(ShopNum).ShopItem(i).Amount & SEP_CHAR & Shop(ShopNum).ShopItem(i).Price & SEP_CHAR)
    Next i
    Packet = Packet & END_CHAR

    Call SendDataToAll(Packet)
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal ShopNum)
    Dim Packet As String
    Dim i As Integer

    Packet = "UPDATESHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Shop(ShopNum).BuysItems & SEP_CHAR & Shop(ShopNum).ShowInfo & SEP_CHAR & Shop(ShopNum).CurrencyItem & SEP_CHAR
    For i = 1 To MAX_SHOP_ITEMS
        Packet = Packet & (Shop(ShopNum).ShopItem(i).ItemNum & SEP_CHAR & Shop(ShopNum).ShopItem(i).Amount & SEP_CHAR & Shop(ShopNum).ShopItem(i).Price & SEP_CHAR)
    Next i
    Packet = Packet & END_CHAR

    Call SendDataTo(index, Packet)
End Sub

Sub SendEditShopTo(ByVal index As Long, ByVal ShopNum As Long)
    Dim Packet As String
    Dim z As Integer

    Packet = "EDITSHOP" & SEP_CHAR & ShopNum & SEP_CHAR & Trim$(Shop(ShopNum).Name) & SEP_CHAR & Shop(ShopNum).FixesItems & SEP_CHAR & Shop(ShopNum).BuysItems & SEP_CHAR & Shop(ShopNum).ShowInfo & SEP_CHAR & Shop(ShopNum).CurrencyItem & SEP_CHAR
    For z = 1 To MAX_SHOP_ITEMS
        Packet = Packet & (Shop(ShopNum).ShopItem(z).ItemNum & SEP_CHAR & Shop(ShopNum).ShopItem(z).Amount & SEP_CHAR & Shop(ShopNum).ShopItem(z).Price & SEP_CHAR)
    Next z
    Packet = Packet & END_CHAR

    Call SendDataTo(index, Packet)
End Sub

Sub SendSpells(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_SPELLS
        If Trim$(Spell(i).Name) <> vbNullString Then
            Call SendUpdateSpellTo(index, i)
        End If
    Next i
End Sub

Sub SendUpdateSpellToAll(ByVal spellnum As Long)
    Call SendDataToAll("UPDATESPELL" & SEP_CHAR & spellnum & SEP_CHAR & Trim$(Spell(spellnum).Name) & END_CHAR)
End Sub
Sub SendUpdateSpellTo(ByVal index As Long, ByVal spellnum As Long)
    Call SendDataTo(index, "UPDATESPELL" & SEP_CHAR & spellnum & SEP_CHAR & Trim$(Spell(spellnum).Name) & END_CHAR)
End Sub

Sub SendEditSpellTo(ByVal index As Long, ByVal spellnum As Long)
    Call SendDataTo(index, "EDITSPELL" & SEP_CHAR & spellnum & SEP_CHAR & Trim$(Spell(spellnum).Name) & SEP_CHAR & Spell(spellnum).ClassReq & SEP_CHAR & Spell(spellnum).LevelReq & SEP_CHAR & Spell(spellnum).Type & SEP_CHAR & Spell(spellnum).Data1 & SEP_CHAR & Spell(spellnum).Data2 & SEP_CHAR & Spell(spellnum).Data3 & SEP_CHAR & Spell(spellnum).MPCost & SEP_CHAR & Spell(spellnum).Sound & SEP_CHAR & Spell(spellnum).Range & SEP_CHAR & Spell(spellnum).SpellAnim & SEP_CHAR & Spell(spellnum).SpellTime & SEP_CHAR & Spell(spellnum).SpellDone & SEP_CHAR & Spell(spellnum).AE & SEP_CHAR & Spell(spellnum).Big & SEP_CHAR & Spell(spellnum).Element & SEP_CHAR & Spell(spellnum).TimeToCast & SEP_CHAR & Spell(spellnum).CastTimer & END_CHAR)
End Sub


Sub SendTrade(ByVal index As Long, ByVal ShopNum As Long)
    Call SendDataTo(index, "GOSHOP" & SEP_CHAR & ShopNum & END_CHAR)
End Sub

Sub SendPlayerSpells(ByVal index As Long)
    Dim Packet As String
    Dim i As Long

    Packet = "SPELLS" & SEP_CHAR
    For i = 1 To MAX_PLAYER_SPELLS
        Packet = Packet & (GetPlayerSpell(index, i) & SEP_CHAR)
    Next i
    Packet = Packet & END_CHAR

    Call SendDataTo(index, Packet)
End Sub

Sub SendWeatherTo(ByVal index As Long)
    If WeatherLevel <= 0 Then
        WeatherLevel = 1
    End If

    Call SendDataTo(index, "WEATHER" & SEP_CHAR & WeatherType & SEP_CHAR & WeatherLevel & END_CHAR)
End Sub

Sub SendWeatherToAll()
    Dim i As Long
    Dim Weather As String

    Select Case WeatherType
        Case 0
            Weather = "Nada"
        Case 1
            Weather = "Lluvia"
        Case 2
            Weather = "Nieve"
        Case 3
            Weather = "Tormenta"
    End Select

    frmServer.Label5.Caption = "Tiempo Actual: " & Weather

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendWeatherTo(i)
        End If
    Next i
End Sub

Sub SendGameClockTo(ByVal index As Long)
    Call SendDataTo(index, "GAMECLOCK" & SEP_CHAR & Seconds & SEP_CHAR & Minutes & SEP_CHAR & Hours & SEP_CHAR & Gamespeed & END_CHAR)
End Sub

Sub SendGameClockToAll()
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendGameClockTo(i)
        End If
    Next i
End Sub
Sub SendNewsTo(ByVal index As Long)
    Dim Packet As String
    Dim Red As Integer
    Dim Green As Integer
    Dim Blue As Integer

    On Error GoTo NewsError
    Red = Val(ReadINI("COLOR", "Red", App.Path & "\Noticias.ini", "255"))
    Green = Val(ReadINI("COLOR", "Green", App.Path & "\Noticias.ini", "255"))
    Blue = Val(ReadINI("COLOR", "Blue", App.Path & "\Noticias.ini", "255"))

    Packet = "NEWS" & SEP_CHAR & ReadINI("DATA", "NewsTitle", App.Path & "\Noticias.ini", vbNullString) & SEP_CHAR
    Packet = Packet & Red & SEP_CHAR & Green & SEP_CHAR & Blue & SEP_CHAR & ReadINI("DATA", "NewsBody", App.Path & "\Noticias.ini", vbNullString) & END_CHAR

    Call SendDataTo(index, Packet)
    Exit Sub

NewsError:
    ' Error reading the news, so just send white
    Red = 255
    Green = 255
    Blue = 255

    Packet = "NEWS" & SEP_CHAR & ReadINI("DATA", "NewsTitle", App.Path & "\Noticias.ini", vbNullString) & SEP_CHAR
    Packet = Packet & Red & SEP_CHAR & Green & SEP_CHAR & Blue & SEP_CHAR & ReadINI("DATA", "NewsBody", App.Path & "\Noticias.ini", vbNullString) & END_CHAR

    Call SendDataTo(index, Packet)
End Sub


Sub SendTimeTo(ByVal index As Long)
    Call SendDataTo(index, "TIME" & SEP_CHAR & GameTime & END_CHAR)
End Sub

Sub SendTimeToAll()
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendTimeTo(i)
        End If
    Next i

    Call SpawnAllMapNpcs
End Sub

Sub MapMsg2(ByVal mapnum As Long, ByVal Msg As String, ByVal index As Long)
    Call SendDataToMap(mapnum, "MAPMSG2" & SEP_CHAR & Msg & SEP_CHAR & index & END_CHAR)
End Sub

Sub DisabledTime()
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call DisabledTimeTo(i)
        End If
    Next i
End Sub

Sub DisabledTimeTo(ByVal index As Long)
    Call SendDataTo(index, "DTIME" & SEP_CHAR & TimeDisable & END_CHAR)
End Sub

Sub SendSprite(ByVal index As Long, ByVal indexto As Long)
    Call SendDataTo(indexto, "cussprite" & SEP_CHAR & index & SEP_CHAR & Player(index).Char(Player(index).CharNum).Head & SEP_CHAR & Player(index).Char(Player(index).CharNum).Body & SEP_CHAR & Player(index).Char(Player(index).CharNum).Leg & END_CHAR)
End Sub

Sub GrapleHook(ByVal index As Long)
    Dim x As Long, y As Long, mapnum As Long
    mapnum = GetPlayerMap(index)

    If Player(index).HookShotX <> 0 Or Player(index).HookShotY <> 0 Then
        If Player(index).Locked = True Then
            Call PlayerMsg(index, "Solo puedes disparar un grappleshot al mismo tiempo", 1)
            Exit Sub
        End If
    End If

    Player(index).Locked = True
    Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(index) & SEP_CHAR & Map(GetPlayerMap(index)).Revision & END_CHAR)

    If GetPlayerDir(index) = DIR_DOWN Then
        x = GetPlayerX(index)
        y = GetPlayerY(index) + 1
        Do While y <= MAX_MAPY
            If Map(mapnum).Tile(x, y).Type = TILE_TYPE_HOOKSHOT Then
                Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
                Player(index).HookShotX = x
                Player(index).HookShotY = y
                Exit Sub
            Else
                If Map(mapnum).Tile(x, y).Type = TILE_TYPE_BLOCKED Then
                    Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 0 & END_CHAR)
                    Player(index).HookShotX = x
                    Player(index).HookShotY = y
                    Exit Sub
                End If
            End If
            y = y + 1
        Loop
        Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 0 & END_CHAR)
        Player(index).HookShotX = x
        Player(index).HookShotY = y
        Exit Sub
    End If
    If GetPlayerDir(index) = DIR_UP Then
        x = GetPlayerX(index)
        y = GetPlayerY(index) - 1
        Do While y >= 0
            If Map(mapnum).Tile(x, y).Type = TILE_TYPE_HOOKSHOT Then
                Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
                Player(index).HookShotX = x
                Player(index).HookShotY = y
                Exit Sub
            Else
                If Map(mapnum).Tile(x, y).Type = TILE_TYPE_BLOCKED Then
                    Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 0 & END_CHAR)
                    Player(index).HookShotX = x
                    Player(index).HookShotY = y
                    Exit Sub
                End If
            End If
            y = y - 1
        Loop
        Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 0 & END_CHAR)
        Player(index).HookShotX = x
        Player(index).HookShotY = y
        Exit Sub
    End If

    If GetPlayerDir(index) = DIR_RIGHT Then
        x = GetPlayerX(index) + 1
        y = GetPlayerY(index)
        Do While x <= MAX_MAPX
            If Map(mapnum).Tile(x, y).Type = TILE_TYPE_HOOKSHOT Then
                Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
                Player(index).HookShotX = x
                Player(index).HookShotY = y
                Exit Sub
            Else
                If Map(mapnum).Tile(x, y).Type = TILE_TYPE_BLOCKED Then
                    Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 0 & END_CHAR)
                    Player(index).HookShotX = x
                    Player(index).HookShotY = y
                    Exit Sub
                End If
            End If
            x = x + 1
        Loop
        Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 0 & END_CHAR)
        Player(index).HookShotX = x
        Player(index).HookShotY = y
        Exit Sub
    End If

    If GetPlayerDir(index) = DIR_LEFT Then
        x = GetPlayerX(index) - 1
        y = GetPlayerY(index)
        Do While x >= 0
            If Map(mapnum).Tile(x, y).Type = TILE_TYPE_HOOKSHOT Then
                Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
                Player(index).HookShotX = x
                Player(index).HookShotY = y
                Exit Sub
            Else
                If Map(mapnum).Tile(x, y).Type = TILE_TYPE_BLOCKED Then
                    Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 0 & END_CHAR)
                    Player(index).HookShotX = x
                    Player(index).HookShotY = y
                    Exit Sub
                End If
            End If
            x = x - 1
        Loop
        Call SendDataToMap(GetPlayerMap(index), "hookshot" & SEP_CHAR & index & SEP_CHAR & Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3 & SEP_CHAR & GetPlayerDir(index) & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 0 & END_CHAR)
        Player(index).HookShotX = x
        Player(index).HookShotY = y
        Exit Sub
    End If
End Sub
