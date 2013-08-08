Attribute VB_Name = "modGameEditor"
Option Explicit

Public Sub EditorInit()
    Dim I As Long

    InEditor = True

    'frmMapEditor.Show vbModeless
    Call frmMapEditor.Show(vbModeless, frmMirage)

    EditorSet = 0

    MapEditorSelectedType = 1

    If FileExists("GFX\Tiles0.bmp") = True Then
    frmMapEditor.picBackSelect.Picture = LoadPicture(App.Path & "\GFX\Tiles0.bmp")

    EditorSet = 0
    End If
    frmMapEditor.scrlPicture.max = Int((frmMapEditor.picBackSelect.Height - frmMapEditor.picBack.Height) / PIC_Y)
    frmMapEditor.picBack.Width = 448
End Sub

Public Sub MainMenuInit()
    frmLogin.txtName.Text = Trim$(ReadINI("CONFIG", "Account", App.Path & "\config.ini"))
    frmLogin.txtPassword.Text = Trim$(ReadINI("CONFIG", "Password", App.Path & "\config.ini"))

    If frmLogin.Check1.value = 0 Then
        frmLogin.Check2.value = 0
    End If

    If ConnectToServer = True And AutoLogin = 1 Then
        'frmMainMenu.picAutoLogin.Visible = True
        frmChars.Label1.Visible = True
    Else
        'frmMainMenu.picAutoLogin.Visible = False
        frmChars.Label1.Visible = False
    End If
End Sub

Public Sub ParseNews()
    Dim FileData As String
    Dim FileTitle As String
    Dim FileBody As String
    Dim Red As Integer
    Dim Blue As Integer
    Dim GRN As Integer

    FileData = ReadINI("DATA", "News", App.Path & "\Noticias.ini")
    FileTitle = Replace(FileData, "*", vbNewLine)

    FileData = ReadINI("DATA", "Desc", App.Path & "\Noticias.ini")
    FileBody = Replace(FileData, "*", vbNewLine)

    frmMainMenu.picNews.Caption = FileTitle & vbNewLine & vbNewLine & FileBody

    Red = Val(ReadINI("COLOR", "Red", App.Path & "\Noticias.ini"))
    GRN = Val(ReadINI("COLOR", "Green", App.Path & "\Noticias.ini"))
    Blue = Val(ReadINI("COLOR", "Blue", App.Path & "\Noticias.ini"))

    If Red < 0 Or Red > 255 Or GRN < 0 Or GRN > 255 Or Blue < 0 Or Blue > 255 Then
        frmMainMenu.picNews.ForeColor = RGB(255, 255, 255)
    Else
        frmMainMenu.picNews.ForeColor = RGB(Red, GRN, Blue)
    End If
End Sub

Public Sub EditorMouseDown(Button As Integer, Shift As Integer, X As Long, Y As Long)
    Dim X2 As Long, Y2 As Long, PicX As Long

    If InEditor Then

        If frmMapEditor.MousePointer = 2 Then
            If MapEditorSelectedType = 1 Then
                With Map(GetPlayerMap(MyIndex)).Tile(X, Y)
                    If frmMapEditor.optGround.value = True Then
                        PicX = .Ground
                        EditorSet = .GroundSet
                    End If
                    If frmMapEditor.optMask.value = True Then
                        PicX = .mask
                        EditorSet = .MaskSet
                    End If
                    If frmMapEditor.optAnim.value = True Then
                        PicX = .Anim
                        EditorSet = .AnimSet
                    End If
                    If frmMapEditor.optMask2.value = True Then
                        PicX = .Mask2
                        EditorSet = .Mask2Set
                    End If
                    If frmMapEditor.optM2Anim.value = True Then
                        PicX = .M2Anim
                        EditorSet = .M2AnimSet
                    End If
                    If frmMapEditor.optFringe.value = True Then
                        PicX = .Fringe
                        EditorSet = .FringeSet
                    End If
                    If frmMapEditor.optFAnim.value = True Then
                        PicX = .FAnim
                        EditorSet = .FAnimSet
                    End If
                    If frmMapEditor.optFringe2.value = True Then
                        PicX = .Fringe2
                        EditorSet = .Fringe2Set
                    End If
                    If frmMapEditor.optF2Anim.value = True Then
                        PicX = .F2Anim
                        EditorSet = .F2AnimSet
                    End If

                    EditorTileY = Int(PicX / TilesInSheets)
                    EditorTileX = (PicX - Int(PicX / TilesInSheets) * TilesInSheets)
                    frmMapEditor.shpSelected.top = Int(EditorTileY * PIC_Y)
                    frmMapEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
                    frmMapEditor.shpSelected.Height = PIC_Y
                    frmMapEditor.shpSelected.Width = PIC_X
                End With
                
            ElseIf MapEditorSelectedType = 3 Then
                EditorTileY = Int(Map(GetPlayerMap(MyIndex)).Tile(X, Y).Light / TilesInSheets)
                EditorTileX = (Map(GetPlayerMap(MyIndex)).Tile(X, Y).Light - Int(Map(GetPlayerMap(MyIndex)).Tile(X, Y).Light / TilesInSheets) * TilesInSheets)
                frmMapEditor.shpSelected.top = Int(EditorTileY * PIC_Y)
                frmMapEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
                frmMapEditor.shpSelected.Height = PIC_Y
                frmMapEditor.shpSelected.Width = PIC_X
                
            ElseIf MapEditorSelectedType = 2 Then
                With Map(GetPlayerMap(MyIndex)).Tile(X, Y)
                    If .Type = TILE_TYPE_BLOCKED Then
                        frmMapEditor.optBlocked.value = True
                    End If
                    If .Type = TILE_TYPE_WALKTHRU Then
                        frmMapEditor.optWalkThru.value = True
                    End If
                    If .Type = TILE_TYPE_WARP Then
                        EditorWarpMap = .Data1
                        EditorWarpX = .Data2
                        EditorWarpY = .Data3
                        frmMapEditor.optWarp.value = True
                    End If
                    If .Type = TILE_TYPE_HEAL Then
                        frmMapEditor.optHeal.value = True
                    End If
                    If .Type = TILE_TYPE_ROOFBLOCK Then
                        frmMapEditor.optRoofBlock.value = True
                        RoofId = .String1
                    End If
                    If .Type = TILE_TYPE_ROOF Then
                        frmMapEditor.optRoof.value = True
                        RoofId = .String1
                    End If
                    If .Type = TILE_TYPE_KILL Then
                        frmMapEditor.optKill.value = True
                    End If
                    If .Type = TILE_TYPE_ITEM Then
                        ItemEditorNum = .Data1
                        ItemEditorValue = .Data2
                        frmMapEditor.optItem.value = True
                    End If
                    If .Type = TILE_TYPE_NPCAVOID Then
                        frmMapEditor.optNpcAvoid.value = True
                    End If
                    If .Type = TILE_TYPE_KEY Then
                        KeyEditorNum = .Data1
                        KeyEditorTake = .Data2
                        frmMapEditor.optKey.value = True
                    End If
                    If .Type = TILE_TYPE_KEYOPEN Then
                        KeyOpenEditorX = .Data1
                        KeyOpenEditorY = .Data2
                        KeyOpenEditorMsg = .String1
                        frmMapEditor.optKeyOpen.value = True
                    End If
                    If .Type = TILE_TYPE_SHOP Then
                        EditorShopNum = .Data1
                        frmMapEditor.optShop.value = True
                    End If
                    If .Type = TILE_TYPE_CBLOCK Then
                        EditorItemNum1 = .Data1
                        EditorItemNum2 = .Data2
                        EditorItemNum3 = .Data3
                        frmMapEditor.optCBlock.value = True
                    End If
                    If .Type = TILE_TYPE_ARENA Then
                        Arena1 = .Data1
                        Arena2 = .Data2
                        Arena3 = .Data3
                        frmMapEditor.optArena.value = True
                    End If
                    If .Type = TILE_TYPE_SOUND Then
                        SoundFileName = .String1
                        frmMapEditor.optSound.value = True
                    End If
                    If .Type = TILE_TYPE_SPRITE_CHANGE Then
                        SpritePic = .Data1
                        SpriteItem = .Data2
                        SpritePrice = .Data3
                        frmMapEditor.optSprite.value = True
                    End If
                    If .Type = TILE_TYPE_SIGN Then
                        SignLine1 = .String1
                        SignLine2 = .String2
                        SignLine3 = .String3
                        frmMapEditor.optSign.value = True
                    End If
                    If .Type = TILE_TYPE_DOOR Then
                        frmMapEditor.optDoor.value = True
                    End If
                    If .Type = TILE_TYPE_NOTICE Then
                        NoticeTitle = .String1
                        NoticeText = .String2
                        NoticeSound = .String3
                        frmMapEditor.optNotice.value = True
                    End If
                    If .Type = TILE_TYPE_CHEST Then
                        frmMapEditor.optChest.value = True
                    End If
                    If .Type = TILE_TYPE_CLASS_CHANGE Then
                        ClassChange = .Data1
                        ClassChangeReq = .Data2
                        frmMapEditor.optClassChange.value = True
                    End If
                    If .Type = TILE_TYPE_SCRIPTED Then
                        ScriptNum = .Data1
                        frmMapEditor.optScripted.value = True
                    End If
                    If .Type = TILE_TYPE_HOUSE Then
                        HouseItem = .Data1
                        HousePrice = .Data2
                        frmMapEditor.optHouse.value = True
                    End If
                    If .Type = TILE_TYPE_GUILDBLOCK Then
                        GuildBlock = .Data1
                        frmMapEditor.optGuildBlock.value = True
                    End If
                    If .Type = TILE_TYPE_BANK Then
                        frmMapEditor.optBank.value = True
                    End If
                    If .Type = TILE_TYPE_HOOKSHOT Then
                        frmMapEditor.OptGHook.value = True
                    End If
                    If .Type = TILE_TYPE_ONCLICK Then
                        ClickScript = .Data1
                        frmMapEditor.optClick.value = True
                    End If
                    If .Type = TILE_TYPE_LOWER_STAT Then
                        MinusHp = .Data1
                        MinusMp = .Data2
                        MinusSp = .Data3
                        MessageMinus = .String1
                        frmMapEditor.optMinusStat.value = True
                    End If
                End With
            End If
            frmMapEditor.MousePointer = 1
            frmMirage.MousePointer = 1
        Else
            If (Button = 1) And (X >= 0) And (X <= MAX_MAPX) And (Y >= 0) And (Y <= MAX_MAPY) Then
                If frmMapEditor.shpSelected.Height <= PIC_Y And frmMapEditor.shpSelected.Width <= PIC_X Then
                    If MapEditorSelectedType = 1 Then
                        With Map(GetPlayerMap(MyIndex)).Tile(X, Y)
                            If frmMapEditor.optGround.value = True Then
                                .Ground = EditorTileY * TilesInSheets + EditorTileX
                                .GroundSet = EditorSet
                            End If
                            If frmMapEditor.optMask.value = True Then
                                .mask = EditorTileY * TilesInSheets + EditorTileX
                                .MaskSet = EditorSet
                            End If
                            If frmMapEditor.optAnim.value = True Then
                                .Anim = EditorTileY * TilesInSheets + EditorTileX
                                .AnimSet = EditorSet
                            End If
                            If frmMapEditor.optMask2.value = True Then
                                .Mask2 = EditorTileY * TilesInSheets + EditorTileX
                                .Mask2Set = EditorSet
                            End If
                            If frmMapEditor.optM2Anim.value = True Then
                                .M2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .M2AnimSet = EditorSet
                            End If
                            If frmMapEditor.optFringe.value = True Then
                                .Fringe = EditorTileY * TilesInSheets + EditorTileX
                                .FringeSet = EditorSet
                            End If
                            If frmMapEditor.optFAnim.value = True Then
                                .FAnim = EditorTileY * TilesInSheets + EditorTileX
                                .FAnimSet = EditorSet
                            End If
                            If frmMapEditor.optFringe2.value = True Then
                                .Fringe2 = EditorTileY * TilesInSheets + EditorTileX
                                .Fringe2Set = EditorSet
                            End If
                            If frmMapEditor.optF2Anim.value = True Then
                                .F2Anim = EditorTileY * TilesInSheets + EditorTileX
                                .F2AnimSet = EditorSet
                            End If
                        End With
                    ElseIf MapEditorSelectedType = 3 Then
                        Map(GetPlayerMap(MyIndex)).Tile(X, Y).Light = EditorTileY * TilesInSheets + EditorTileX
                    ElseIf MapEditorSelectedType = 2 Then
                        With Map(GetPlayerMap(MyIndex)).Tile(X, Y)
                            If frmMapEditor.optBlocked.value = True Then
                                .Type = TILE_TYPE_BLOCKED
                            End If
                            If frmMapEditor.optRoofBlock.value = True Then
                                .Type = TILE_TYPE_ROOFBLOCK
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = RoofId
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optRoof.value = True Then
                                .Type = TILE_TYPE_ROOF
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = RoofId
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optWarp.value = True Then
                                .Type = TILE_TYPE_WARP
                                .Data1 = EditorWarpMap
                                .Data2 = EditorWarpX
                                .Data3 = EditorWarpY
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If

                            If frmMapEditor.optHeal.value = True Then
                                .Type = TILE_TYPE_HEAL
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If

                            If frmMapEditor.optKill.value = True Then
                                .Type = TILE_TYPE_KILL
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If GetPlayerAccess(MyIndex) >= 3 Then
                                If frmMapEditor.optItem.value = True Then
                                    .Type = TILE_TYPE_ITEM
                                    .Data1 = ItemEditorNum
                                    .Data2 = ItemEditorValue
                                    .Data3 = 0
                                    .String1 = vbNullString
                                    .String2 = vbNullString
                                    .String3 = vbNullString
                                End If
                            End If
                            If frmMapEditor.optNpcAvoid.value = True Then
                                .Type = TILE_TYPE_NPCAVOID
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optKey.value = True Then
                                .Type = TILE_TYPE_KEY
                                .Data1 = KeyEditorNum
                                .Data2 = KeyEditorTake
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optKeyOpen.value = True Then
                                .Type = TILE_TYPE_KEYOPEN
                                .Data1 = KeyOpenEditorX
                                .Data2 = KeyOpenEditorY
                                .Data3 = 0
                                .String1 = KeyOpenEditorMsg
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optShop.value = True Then
                                .Type = TILE_TYPE_SHOP
                                .Data1 = EditorShopNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optCBlock.value = True Then
                                .Type = TILE_TYPE_CBLOCK
                                .Data1 = EditorItemNum1
                                .Data2 = EditorItemNum2
                                .Data3 = EditorItemNum3
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optArena.value = True Then
                                .Type = TILE_TYPE_ARENA
                                .Data1 = Arena1
                                .Data2 = Arena2
                                .Data3 = Arena3
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optSound.value = True Then
                                .Type = TILE_TYPE_SOUND
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = SoundFileName
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optSprite.value = True Then
                                .Type = TILE_TYPE_SPRITE_CHANGE
                                .Data1 = SpritePic
                                .Data2 = SpriteItem
                                .Data3 = SpritePrice
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optSign.value = True Then
                                .Type = TILE_TYPE_SIGN
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = SignLine1
                                .String2 = SignLine2
                                .String3 = SignLine3
                            End If
                            If frmMapEditor.optDoor.value = True Then
                                .Type = TILE_TYPE_DOOR
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optNotice.value = True Then
                                .Type = TILE_TYPE_NOTICE
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = NoticeTitle
                                .String2 = NoticeText
                                .String3 = NoticeSound
                            End If
                            If frmMapEditor.optChest.value = True Then
                                .Type = TILE_TYPE_CHEST
                                .Data1 = ChestItemNum
                                .Data2 = ChestItemAmount
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If GetPlayerAccess(MyIndex) >= 3 Then
                                If frmMapEditor.optClassChange.value = True Then
                                    .Type = TILE_TYPE_CLASS_CHANGE
                                    .Data1 = ClassChange
                                    .Data2 = ClassChangeReq
                                    .Data3 = 0
                                    .String1 = vbNullString
                                    .String2 = vbNullString
                                    .String3 = vbNullString
                                End If
                            End If
                            If frmMapEditor.optScripted.value = True Then
                                .Type = TILE_TYPE_SCRIPTED
                                .Data1 = ScriptNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optHouse.value = True Then
                                .Type = TILE_TYPE_HOUSE
                                .Data1 = HouseItem
                                .Data2 = HousePrice
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optGuildBlock.value = True Then
                                .Type = TILE_TYPE_GUILDBLOCK
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = GuildBlock
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optBank.value = True Then
                                .Type = TILE_TYPE_BANK
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.OptGHook.value = True Then
                                .Type = TILE_TYPE_HOOKSHOT
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optWalkThru.value = True Then
                                .Type = TILE_TYPE_WALKTHRU
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optClick.value = True Then
                                .Type = TILE_TYPE_ONCLICK
                                .Data1 = ClickScript
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optMinusStat.value = True Then
                                .Type = TILE_TYPE_LOWER_STAT
                                .Data1 = MinusHp
                                .Data2 = MinusMp
                                .Data3 = MinusSp
                                .String1 = MessageMinus
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                        End With
                    End If
                Else
                    For Y2 = 0 To Int(frmMapEditor.shpSelected.Height / PIC_Y) - 1
                        For X2 = 0 To Int(frmMapEditor.shpSelected.Width / PIC_X) - 1
                            If X + X2 <= MAX_MAPX Then
                                If Y + Y2 <= MAX_MAPY Then
                                    If MapEditorSelectedType = 1 Then
                                        With Map(GetPlayerMap(MyIndex)).Tile(X + X2, Y + Y2)
                                            If frmMapEditor.optGround.value = True Then
                                                .Ground = (EditorTileY + Y2) * TilesInSheets + (EditorTileX + X2)
                                                .GroundSet = EditorSet
                                            End If
                                            If frmMapEditor.optMask.value = True Then
                                                .mask = (EditorTileY + Y2) * TilesInSheets + (EditorTileX + X2)
                                                .MaskSet = EditorSet
                                            End If
                                            If frmMapEditor.optAnim.value = True Then
                                                .Anim = (EditorTileY + Y2) * TilesInSheets + (EditorTileX + X2)
                                                .AnimSet = EditorSet
                                            End If
                                            If frmMapEditor.optMask2.value = True Then
                                                .Mask2 = (EditorTileY + Y2) * TilesInSheets + (EditorTileX + X2)
                                                .Mask2Set = EditorSet
                                            End If
                                            If frmMapEditor.optM2Anim.value = True Then
                                                .M2Anim = (EditorTileY + Y2) * TilesInSheets + (EditorTileX + X2)
                                                .M2AnimSet = EditorSet
                                            End If
                                            If frmMapEditor.optFringe.value = True Then
                                                .Fringe = (EditorTileY + Y2) * TilesInSheets + (EditorTileX + X2)
                                                .FringeSet = EditorSet
                                            End If
                                            If frmMapEditor.optFAnim.value = True Then
                                                .FAnim = (EditorTileY + Y2) * TilesInSheets + (EditorTileX + X2)
                                                .FAnimSet = EditorSet
                                            End If
                                            If frmMapEditor.optFringe2.value = True Then
                                                .Fringe2 = (EditorTileY + Y2) * TilesInSheets + (EditorTileX + X2)
                                                .Fringe2Set = EditorSet
                                            End If
                                            If frmMapEditor.optF2Anim.value = True Then
                                                .F2Anim = (EditorTileY + Y2) * TilesInSheets + (EditorTileX + X2)
                                                .F2AnimSet = EditorSet
                                            End If
                                        End With
                                    ElseIf MapEditorSelectedType = 3 Then
                                        Map(GetPlayerMap(MyIndex)).Tile(X + X2, Y + Y2).Light = (EditorTileY + Y2) * TilesInSheets + (EditorTileX + X2)
                                    End If
                                End If
                            End If
                        Next X2
                    Next Y2
                End If
            End If

            If (Button = 2) And (X >= 0) And (X <= MAX_MAPX) And (Y >= 0) And (Y <= MAX_MAPY) Then
                If MapEditorSelectedType = 1 Then
                    With Map(GetPlayerMap(MyIndex)).Tile(X, Y)
                        If frmMapEditor.optGround.value = True Then
                            .Ground = 0
                        End If
                        If frmMapEditor.optMask.value = True Then
                            .mask = 0
                        End If
                        If frmMapEditor.optAnim.value = True Then
                            .Anim = 0
                        End If
                        If frmMapEditor.optMask2.value = True Then
                            .Mask2 = 0
                        End If
                        If frmMapEditor.optM2Anim.value = True Then
                            .M2Anim = 0
                        End If
                        If frmMapEditor.optFringe.value = True Then
                            .Fringe = 0
                        End If
                        If frmMapEditor.optFAnim.value = True Then
                            .FAnim = 0
                        End If
                        If frmMapEditor.optFringe2.value = True Then
                            .Fringe2 = 0
                        End If
                        If frmMapEditor.optF2Anim.value = True Then
                            .F2Anim = 0
                        End If
                    End With
                ElseIf MapEditorSelectedType = 3 Then
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).Light = 0
                ElseIf MapEditorSelectedType = 2 Then
                    With Map(GetPlayerMap(MyIndex)).Tile(X, Y)
                        .Type = 0
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End With
                End If
            End If
        End If
    End If
End Sub

Public Sub EditorChooseTile(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        EditorTileX = Int(X / PIC_X)
        EditorTileY = Int(Y / PIC_Y)
    End If
    frmMapEditor.shpSelected.top = Int(EditorTileY * PIC_Y)
    frmMapEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
End Sub

Public Sub EditorTileScroll()
    frmMapEditor.picBackSelect.top = (frmMapEditor.scrlPicture.value * PIC_Y) * -1
End Sub

Public Sub EditorSend()
    Call SendMap
    Call EditorCancel
End Sub

Public Sub EditorCancel()
    ScreenMode = 0
    NightMode = 0
    GridMode = 0

    ' Set the type back to default.
    MapEditorSelectedType = 1

    ' Set the map controls to default.
    frmMapEditor.fraAttribs.Visible = False
    frmMapEditor.fraLayers.Visible = True
    frmMapEditor.frmtile.Visible = True

    InEditor = False
    frmMapEditor.Visible = False

    frmMirage.Show
    frmMapEditor.MousePointer = 1
    frmMirage.MousePointer = 1

    Call LoadMap(GetPlayerMap(MyIndex))
End Sub

Public Sub EditorClearLayer()
    Dim Choice As Integer
    Dim X As Byte
    Dim Y As Byte

    ' Ground Layer
    If frmMapEditor.optGround.value Then
        Choice = MsgBox("Estas seguro de querer borrar por completo la capa Suelo?", vbYesNo, GAME_NAME)

        If Choice = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).Ground = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).GroundSet = 0
                Next X
            Next Y
        End If
    End If

    ' Mask Layer
    If frmMapEditor.optMask.value Then
        Choice = MsgBox("Estas seguro de querer borrar por completo la capa Mascara?", vbYesNo, GAME_NAME)

        If Choice = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).mask = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).MaskSet = 0
                Next X
            Next Y
        End If
    End If

    ' Mask Animation Layer
    If frmMapEditor.optAnim.value Then
        Choice = MsgBox("Estas seguro de querer borrar por completo la capa Animación?", vbYesNo, GAME_NAME)

        If Choice = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).AnimSet = 0
                Next X
            Next Y
        End If
    End If

    ' Mask 2 Layer
    If frmMapEditor.optMask2.value Then
        Choice = MsgBox("Estas seguro de querer borrar por completo la capa Mascara 2?", vbYesNo, GAME_NAME)

        If Choice = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).Mask2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).Mask2Set = 0
                Next X
            Next Y
        End If
    End If

    ' Mask 2 Animation layer
    If frmMapEditor.optM2Anim.value Then
        Choice = MsgBox("Estas seguro de querer borrar por completo la capa Animación 2?", vbYesNo, GAME_NAME)

        If Choice = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).M2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).M2AnimSet = 0
                Next X
            Next Y
        End If
    End If

    ' Fringe Layer
    If frmMapEditor.optFringe.value Then
        Choice = MsgBox("Estas seguro de querer borrar por completo la capa Superior?", vbYesNo, GAME_NAME)

        If Choice = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).Fringe = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).FringeSet = 0
                Next X
            Next Y
        End If
    End If

    ' Fringe Animation Layer
    If frmMapEditor.optFAnim.value Then
        Choice = MsgBox("Estas seguro de querer borrar por completo la capa Animación 3?", vbYesNo, GAME_NAME)

        If Choice = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).FAnim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).FAnimSet = 0
                Next X
            Next Y
        End If
    End If

    ' Fringe 2 Layer
    If frmMapEditor.optFringe2.value Then
        Choice = MsgBox("Estas seguro de querer borrar por completo la capa Superior 2?", vbYesNo, GAME_NAME)

        If Choice = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).Fringe2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).Fringe2Set = 0
                Next X
            Next Y
        End If
    End If

    ' Fringe 2 Animation Layer
    If frmMapEditor.optF2Anim.value Then
        Choice = MsgBox("Estas seguro de querer borrar por completo la capa Animación 4?", vbYesNo, GAME_NAME)

        If Choice = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).F2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).F2AnimSet = 0
                Next X
            Next Y
        End If
    End If
End Sub

Public Sub EditorClearAttribs()
    Dim Choice As Integer
    Dim X As Byte
    Dim Y As Byte

    Choice = MsgBox("¿Estas seguro de querer eliminar todo los atributos del mapa?", vbYesNo, GAME_NAME)

    If Choice = vbYes Then
        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = 0
            Next X
        Next Y
    End If
End Sub

Public Sub EmoticonEditorInit()
    frmEmoticonEditor.scrlEmoticon.max = MAX_EMOTICONS
    frmEmoticonEditor.scrlEmoticon.value = Emoticons(EditorIndex - 1).Pic
    frmEmoticonEditor.txtCommand.Text = Trim$(Emoticons(EditorIndex - 1).Command)

    frmEmoticonEditor.picEmoticons.Picture = LoadPicture(App.Path & "\GFX\Emoticons.bmp")

    frmEmoticonEditor.Show vbModal
End Sub

Public Sub ElementEditorInit()
    frmElementEditor.txtName.Text = Trim$(Element(EditorIndex - 1).name)
    frmElementEditor.scrlStrong.value = Element(EditorIndex - 1).Strong
    frmElementEditor.scrlWeak.value = Element(EditorIndex - 1).Weak
    frmElementEditor.Show vbModal
End Sub

Public Sub EmoticonEditorOk()
    Emoticons(EditorIndex - 1).Pic = frmEmoticonEditor.scrlEmoticon.value
    If frmEmoticonEditor.txtCommand.Text <> "/" Then
        Emoticons(EditorIndex - 1).Command = frmEmoticonEditor.txtCommand.Text
    Else
        Emoticons(EditorIndex - 1).Command = vbNullString
    End If

    Call SendSaveEmoticon(EditorIndex - 1)
    Call EmoticonEditorCancel
End Sub

Public Sub ElementEditorOk()
    Element(EditorIndex - 1).name = frmElementEditor.txtName.Text
    Element(EditorIndex - 1).Strong = frmElementEditor.scrlStrong.value
    Element(EditorIndex - 1).Weak = frmElementEditor.scrlWeak.value
    Call SendSaveElement(EditorIndex - 1)
    Call ElementEditorCancel
End Sub

Public Sub EmoticonEditorCancel()
    InEmoticonEditor = False
    Unload frmEmoticonEditor
End Sub

Public Sub ElementEditorCancel()
    InElementEditor = False
    Unload frmElementEditor
End Sub

Public Sub ArrowEditorInit()
    frmEditArrows.scrlArrow.max = MAX_ARROWS
    If Arrows(EditorIndex).Pic = 0 Then
        Arrows(EditorIndex).Pic = 1
    End If
    frmEditArrows.scrlArrow.value = Arrows(EditorIndex).Pic
    frmEditArrows.txtName.Text = Arrows(EditorIndex).name
    If Arrows(EditorIndex).Range = 0 Then
        Arrows(EditorIndex).Range = 1
    End If
    frmEditArrows.scrlRange.value = Arrows(EditorIndex).Range
    If Arrows(EditorIndex).Amount = 0 Then
        Arrows(EditorIndex).Amount = 1
    End If
    frmEditArrows.scrlAmount.value = Arrows(EditorIndex).Amount

    frmEditArrows.picArrows.Picture = LoadPicture(App.Path & "\GFX\Arrows.bmp")

    frmEditArrows.Show vbModal
End Sub

Public Sub ArrowEditorOk()
    Arrows(EditorIndex).Pic = frmEditArrows.scrlArrow.value
    Arrows(EditorIndex).Range = frmEditArrows.scrlRange.value
    Arrows(EditorIndex).name = frmEditArrows.txtName.Text
    Arrows(EditorIndex).Amount = frmEditArrows.scrlAmount.value
    Call SendSaveArrow(EditorIndex)
    Call ArrowEditorCancel
End Sub

Public Sub ArrowEditorCancel()
    InArrowEditor = False
    Unload frmEditArrows
End Sub

Public Sub ItemEditorInit()
    Dim I As Long

    EditorItemY = Int(Item(EditorIndex).Pic / 6)
    EditorItemX = (Item(EditorIndex).Pic - Int(Item(EditorIndex).Pic / 6) * 6)

    frmItemEditor.scrlClassReq.max = Max_Classes

    frmItemEditor.picItems.Picture = LoadPicture(App.Path & "\GFX\Items.bmp")

    frmItemEditor.txtName.Text = Trim$(Item(EditorIndex).name)
    frmItemEditor.txtDesc.Text = Trim$(Item(EditorIndex).desc)
    frmItemEditor.cmbType.ListIndex = Item(EditorIndex).Type
    frmItemEditor.txtPrice.Text = Item(EditorIndex).Price
    frmItemEditor.chkStackable.value = Item(EditorIndex).Stackable
    frmItemEditor.chkBound.value = Item(EditorIndex).Bound

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_NECKLACE) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.fraAttributes.Visible = True
        If frmItemEditor.cmbType.ListIndex = ITEM_TYPE_WEAPON Then
            frmItemEditor.fraBow.Visible = True
        End If

        frmItemEditor.scrlDurability.value = Item(EditorIndex).Data1
        frmItemEditor.scrlStrength.value = Item(EditorIndex).Data2
        frmItemEditor.scrlStrReq.value = Item(EditorIndex).StrReq
        frmItemEditor.scrlDefReq.value = Item(EditorIndex).DefReq
        frmItemEditor.scrlSpeedReq.value = Item(EditorIndex).SpeedReq
        frmItemEditor.scrlMagicReq.value = Item(EditorIndex).MagicReq
        frmItemEditor.scrlClassReq.value = Item(EditorIndex).ClassReq
        frmItemEditor.scrlAccessReq.value = Item(EditorIndex).AccessReq
        frmItemEditor.scrlAddHP.value = Item(EditorIndex).AddHP
        frmItemEditor.scrlAddMP.value = Item(EditorIndex).AddMP
        frmItemEditor.scrlAddSP.value = Item(EditorIndex).AddSP
        frmItemEditor.scrlAddStr.value = Item(EditorIndex).AddStr
        frmItemEditor.scrlAddDef.value = Item(EditorIndex).AddDef
        frmItemEditor.scrlAddMagi.value = Item(EditorIndex).AddMagi
        frmItemEditor.scrlAddSpeed.value = Item(EditorIndex).AddSpeed
        ' frmItemEditor.scrlAddEXP.Value = Item(EditorIndex).AddEXP
        frmItemEditor.scrlAttackSpeed.value = Item(EditorIndex).AttackSpeed

        If Item(EditorIndex).Data3 > 0 Then
            If Item(EditorIndex).Stackable = 1 Then
                frmItemEditor.chkBow.value = Checked
                frmItemEditor.chkGrapple.value = Checked
            Else
                frmItemEditor.chkBow.value = Checked
                frmItemEditor.chkGrapple.value = Unchecked
            End If
        Else
            frmItemEditor.chkBow.value = Unchecked
        End If


        frmItemEditor.cmbBow.Clear
        If frmItemEditor.chkBow.value = Checked Then
            For I = 1 To 100
                frmItemEditor.cmbBow.addItem I & ": " & Arrows(I).name
            Next I
            frmItemEditor.cmbBow.ListIndex = Item(EditorIndex).Data3 - 1
            frmItemEditor.picBow.top = (Arrows(Item(EditorIndex).Data3).Pic * 32) * -1
            frmItemEditor.cmbBow.Enabled = True
        Else
            frmItemEditor.cmbBow.addItem "Ninguno"
            frmItemEditor.cmbBow.ListIndex = 0
            frmItemEditor.cmbBow.Enabled = False
        End If
        frmItemEditor.chkStackable.Visible = False
    Else
        frmItemEditor.fraEquipment.Visible = False
    End If

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        frmItemEditor.fraVitals.Visible = True
        frmItemEditor.chkStackable.Visible = True
        frmItemEditor.scrlVitalMod.value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraVitals.Visible = False
    End If

    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        frmItemEditor.fraSpell.Visible = True
        frmItemEditor.scrlSpell.value = Item(EditorIndex).Data1
        frmItemEditor.chkStackable.Visible = False
    Else
        frmItemEditor.fraSpell.Visible = False
    End If

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_SCRIPTED) Then
        frmItemEditor.fraScript.Visible = True
        frmItemEditor.scrlScript.value = Item(EditorIndex).Data1
        frmItemEditor.lblScript.Caption = Item(EditorIndex).Data1
        
        frmItemEditor.chkStackable.Visible = True
    Else
        frmItemEditor.fraScript.Visible = False
    End If
    frmItemEditor.VScroll1.value = EditorItemY
    frmItemEditor.picItems.top = (EditorItemY) * -32
    frmItemEditor.Show vbModal
End Sub

Public Sub ItemEditorOk()
    Item(EditorIndex).name = frmItemEditor.txtName.Text
    Item(EditorIndex).desc = frmItemEditor.txtDesc.Text
    Item(EditorIndex).Pic = EditorItemY * 6 + EditorItemX
    Item(EditorIndex).Type = frmItemEditor.cmbType.ListIndex
    Item(EditorIndex).Price = Val(frmItemEditor.txtPrice.Text)
    Item(EditorIndex).Bound = frmItemEditor.chkBound.value
    Item(EditorIndex).PetSprite = frmItemEditor.petspritetxt.Text
    

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_NECKLACE) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.value
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.value
        If frmItemEditor.chkBow.value = Checked Then
            If frmItemEditor.chkGrapple.value = Checked Then
                Item(EditorIndex).Data3 = frmItemEditor.cmbBow.ListIndex + 1
                Item(EditorIndex).Stackable = 1
            Else
                Item(EditorIndex).Data3 = frmItemEditor.cmbBow.ListIndex + 1
                Item(EditorIndex).Stackable = 0
            End If
        Else
            Item(EditorIndex).Data3 = 0
            Item(EditorIndex).Stackable = 0
        End If
        Item(EditorIndex).StrReq = frmItemEditor.scrlStrReq.value
        Item(EditorIndex).DefReq = frmItemEditor.scrlDefReq.value
        Item(EditorIndex).SpeedReq = frmItemEditor.scrlSpeedReq.value
        Item(EditorIndex).MagicReq = frmItemEditor.scrlMagicReq.value

        Item(EditorIndex).ClassReq = frmItemEditor.scrlClassReq.value
        Item(EditorIndex).AccessReq = frmItemEditor.scrlAccessReq.value

        Item(EditorIndex).AddHP = frmItemEditor.scrlAddHP.value
        Item(EditorIndex).AddMP = frmItemEditor.scrlAddMP.value
        Item(EditorIndex).AddSP = frmItemEditor.scrlAddSP.value
        Item(EditorIndex).AddStr = frmItemEditor.scrlAddStr.value
        Item(EditorIndex).AddDef = frmItemEditor.scrlAddDef.value
        Item(EditorIndex).AddMagi = frmItemEditor.scrlAddMagi.value
        Item(EditorIndex).AddSpeed = frmItemEditor.scrlAddSpeed.value
        Item(EditorIndex).AddEXP = frmItemEditor.scrlAddEXP.value
        Item(EditorIndex).AttackSpeed = frmItemEditor.scrlAttackSpeed.value
    End If

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlVitalMod.value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).MagicReq = 0
        Item(EditorIndex).ClassReq = -1
        Item(EditorIndex).AccessReq = 0

        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddStr = 0
        Item(EditorIndex).AddDef = 0
        Item(EditorIndex).AddMagi = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
        Item(EditorIndex).AttackSpeed = 0
        Item(EditorIndex).Stackable = frmItemEditor.chkStackable.value

    End If

    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_NONE) Then
        Item(EditorIndex).Stackable = frmItemEditor.chkStackable.value
    End If

    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlSpell.value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).MagicReq = 0
        Item(EditorIndex).ClassReq = -1
        Item(EditorIndex).AccessReq = 0

        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddStr = 0
        Item(EditorIndex).AddDef = 0
        Item(EditorIndex).AddMagi = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
        Item(EditorIndex).AttackSpeed = 0
        Item(EditorIndex).Stackable = frmItemEditor.chkStackable.value
    End If

    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SCRIPTED) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlScript.value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).MagicReq = 0
        Item(EditorIndex).ClassReq = -1
        Item(EditorIndex).AccessReq = 0

        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddStr = 0
        Item(EditorIndex).AddDef = 0
        Item(EditorIndex).AddMagi = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
        Item(EditorIndex).AttackSpeed = 0
        Item(EditorIndex).Stackable = frmItemEditor.chkStackable.value
    End If
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_THROW) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlScript.value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).MagicReq = 0
        Item(EditorIndex).ClassReq = -1
        Item(EditorIndex).AccessReq = 0

        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddStr = 0
        Item(EditorIndex).AddDef = 0
        Item(EditorIndex).AddMagi = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
        Item(EditorIndex).AttackSpeed = 0
        Item(EditorIndex).Stackable = 0
    End If
    Call SendSaveItem(EditorIndex)
    InItemsEditor = False
    Unload frmItemEditor
End Sub
Public Sub ItemEditorCancel()
    InItemsEditor = False
    Unload frmItemEditor
End Sub

Public Sub NpcEditorInit()
    On Error Resume Next

    frmNpcEditor.picSprites.Picture = LoadPicture(App.Path & "\GFX\Sprites.bmp")

    frmNpcEditor.txtName.Text = Trim$(Npc(EditorIndex).name)
    frmNpcEditor.txtAttackSay.Text = Trim$(Npc(EditorIndex).AttackSay)
    frmNpcEditor.scrlSprite.value = Npc(EditorIndex).Sprite
    frmNpcEditor.txtSpawnSecs.Text = STR(Npc(EditorIndex).SpawnSecs)
    frmNpcEditor.cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
    frmNpcEditor.scrlRange.value = Npc(EditorIndex).Range
    frmNpcEditor.scrlSTR.value = Npc(EditorIndex).STR
    frmNpcEditor.scrlDEF.value = Npc(EditorIndex).DEF
    frmNpcEditor.scrlSPEED.value = Npc(EditorIndex).SPEED
    frmNpcEditor.scrlMAGI.value = Npc(EditorIndex).MAGI
    frmNpcEditor.BigNpc.value = Npc(EditorIndex).Big
    frmNpcEditor.StartHP.value = Npc(EditorIndex).MaxHP
    frmNpcEditor.ExpGive.value = Npc(EditorIndex).Exp
    frmNpcEditor.scrlChance.value = Npc(EditorIndex).ItemNPC(1).chance
    frmNpcEditor.scrlNum.value = Npc(EditorIndex).ItemNPC(1).ItemNum
    frmNpcEditor.scrlValue.value = Npc(EditorIndex).ItemNPC(1).ItemValue
    If Npc(EditorIndex).Quest > 0 Then
    frmNpcEditor.scrlquest.value = Npc(EditorIndex).Quest
    Else
    frmNpcEditor.Label21.Caption = "Ninguno"
    End If
    If Npc(EditorIndex).Behavior = NPC_BEHAVIOR_SCRIPTED Then
        frmNpcEditor.scrlScript.value = Npc(EditorIndex).SpawnSecs
        frmNpcEditor.scrlElement.value = Npc(EditorIndex).Element
    End If
    If Val(0 + Npc(EditorIndex).SpriteSize) = 0 Then
        frmNpcEditor.Opt32.value = 1
        frmNpcEditor.Opt64.value = 0
    Else
        frmNpcEditor.Opt64.value = 1
        frmNpcEditor.Opt32.value = 0
    End If
    If Npc(EditorIndex).SpawnTime = 0 Then
        frmNpcEditor.chkDay.value = Checked
        frmNpcEditor.chkNight.value = Checked
    ElseIf Npc(EditorIndex).SpawnTime = 1 Then
        frmNpcEditor.chkDay.value = Checked
        frmNpcEditor.chkNight.value = Unchecked
    ElseIf Npc(EditorIndex).SpawnTime = 2 Then
        frmNpcEditor.chkDay.value = Unchecked
        frmNpcEditor.chkNight.value = Checked
    End If
    
    If Npc(EditorIndex).standstill = True Then
        frmNpcEditor.chksstill.value = Checked
    ElseIf Npc(EditorIndex).standstill = False Then
        frmNpcEditor.chksstill.value = Unchecked
    End If
    frmNpcEditor.Show vbModal
End Sub

Public Sub NpcEditorOk()
    Npc(EditorIndex).name = frmNpcEditor.txtName.Text
    Npc(EditorIndex).AttackSay = frmNpcEditor.txtAttackSay.Text
    Npc(EditorIndex).Sprite = frmNpcEditor.scrlSprite.value
    Npc(EditorIndex).Behavior = frmNpcEditor.cmbBehavior.ListIndex
    If Npc(EditorIndex).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
        Npc(EditorIndex).SpawnSecs = Val(frmNpcEditor.txtSpawnSecs.Text)
    Else
        Npc(EditorIndex).SpawnSecs = frmNpcEditor.scrlScript.value
    End If
    Npc(EditorIndex).Range = frmNpcEditor.scrlRange.value
    Npc(EditorIndex).STR = frmNpcEditor.scrlSTR.value
    Npc(EditorIndex).DEF = frmNpcEditor.scrlDEF.value
    Npc(EditorIndex).SPEED = frmNpcEditor.scrlSPEED.value
    Npc(EditorIndex).MAGI = frmNpcEditor.scrlMAGI.value
    Npc(EditorIndex).Big = frmNpcEditor.BigNpc.value
    Npc(EditorIndex).MaxHP = frmNpcEditor.StartHP.value
    Npc(EditorIndex).Exp = frmNpcEditor.ExpGive.value
    Npc(EditorIndex).Quest = frmNpcEditor.scrlquest.value

    If frmNpcEditor.Opt64.value = True Then
        Npc(EditorIndex).SpriteSize = 1
    Else
        Npc(EditorIndex).SpriteSize = 0
    End If

    If frmNpcEditor.chkDay.value = Checked And frmNpcEditor.chkNight.value = Checked Then
        Npc(EditorIndex).SpawnTime = 0
    ElseIf frmNpcEditor.chkDay.value = Checked And frmNpcEditor.chkNight.value = Unchecked Then
        Npc(EditorIndex).SpawnTime = 1
    ElseIf frmNpcEditor.chkDay.value = Unchecked And frmNpcEditor.chkNight.value = Checked Then
        Npc(EditorIndex).SpawnTime = 2
    End If
    
    If frmNpcEditor.chksstill.value = Checked Then
        Npc(EditorIndex).standstill = True
    ElseIf frmNpcEditor.chksstill.value = Unchecked Then
        Npc(EditorIndex).standstill = False
    End If

    Call SendSaveNPC(EditorIndex)
    InNpcEditor = False
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorCancel()
    InNpcEditor = False
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorBltSprite()
    If frmNpcEditor.BigNpc.value = Checked Then
        frmNpcEditor.picSprites.top = frmNpcEditor.scrlSprite.value * 64
        frmNpcEditor.picSprites.Left = 3360
        Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, 64, 64, frmNpcEditor.picSprites.hDC, 3 * 64, frmNpcEditor.scrlSprite.value * 64, SRCCOPY)
    Else
        frmNpcEditor.picSprites.Left = 3600

        If SpriteSize = 1 Then
            frmNpcEditor.picSprites.top = frmNpcEditor.scrlSprite.value * 64
            Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, PIC_X, 64, frmNpcEditor.picSprites.hDC, 3 * PIC_X, frmNpcEditor.scrlSprite.value * 64, SRCCOPY)
        Else
            frmNpcEditor.picSprites.top = frmNpcEditor.scrlSprite.value * PIC_Y
            Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, PIC_X, PIC_Y, frmNpcEditor.picSprites.hDC, 3 * PIC_X, frmNpcEditor.scrlSprite.value * PIC_Y, SRCCOPY)
        End If
    End If
End Sub

' Initializes the shop editor
Public Sub ShopEditorInit()
    Dim I As Integer
    Dim itemN As Integer
    Dim cItemMade As Boolean

    On Error GoTo ShopEditorInit_Error


    frmShopEditor.txtName.Text = Trim$(Shop(EditorIndex).name)
    frmShopEditor.chkFixesItems.value = Shop(EditorIndex).FixesItems
    frmShopEditor.chkShow.value = Shop(EditorIndex).ShowInfo
    frmShopEditor.chkSellsItems.value = Shop(EditorIndex).BuysItems

    cItemMade = False

    frmShopEditor.cmbCurrency.Clear
    frmShopEditor.lstItems.Clear

    ' Add all the currency items to cmbCurrency
    For I = 1 To MAX_ITEMS
        If Item(I).Type = ITEM_TYPE_CURRENCY Then
            ' It's a currency item, so add it to the list
            frmShopEditor.cmbCurrency.addItem (I & " - " & Trim(Item(I).name))
            ' Add it to the item data so that we know the number
            frmShopEditor.cmbCurrency.ItemData(frmShopEditor.cmbCurrency.ListCount - 1) = I
            cItemMade = True 'we have at least 1 currency item
            If Shop(EditorIndex).currencyItem = I Then
                frmShopEditor.cmbCurrency.ListIndex = frmShopEditor.cmbCurrency.ListCount - 1
            End If
        End If
    Next I

    If Not cItemMade Then
        Call MsgBox("Por favor crea primero algún tipo de Moneda primero!")
        Call ShopEditorCancel
        Exit Sub
    End If

    ' Add all the items to the list
    For I = 1 To MAX_SHOP_ITEMS
        itemN = Shop(EditorIndex).ShopItem(I).ItemNum

        ' If the item is not empty
        If itemN > 0 Then
            ' Add the item to the shop list
            Call frmShopEditor.AddShopItem(itemN, Shop(EditorIndex).ShopItem(I).Price, Shop(EditorIndex).currencyItem, Shop(EditorIndex).ShopItem(I).Amount)
        End If
    Next I

    ' Add all items to the 'add item' list
    For I = 1 To MAX_ITEMS
        frmShopEditor.cmbItemList.addItem (I & " - " & Trim(Item(I).name))
    Next I

    frmShopEditor.frmAddEditItem.Visible = False

    ' Init shop editor temp array
    frmShopEditor.LoadShopItemData (EditorIndex)

    frmShopEditor.Show vbModal

    On Error GoTo 0
    Exit Sub

ShopEditorInit_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShopEditorInit of Module modGameLogic"
    ' Close the shop editor
    frmShopEditor.Visible = False
    Call ShopEditorCancel
End Sub


Public Sub ShopEditorOk()
    Dim I As Integer
    Dim currencyItem As Integer

    If frmShopEditor.cmbCurrency.ListIndex < 0 Then
        MsgBox "Por favor selecciona una moneda!", vbExclamation
        Exit Sub
    End If

    currencyItem = frmShopEditor.cmbCurrency.ItemData(frmShopEditor.cmbCurrency.ListIndex)

    Shop(EditorIndex).name = frmShopEditor.txtName.Text
    Shop(EditorIndex).FixesItems = frmShopEditor.chkFixesItems.value
    Shop(EditorIndex).BuysItems = frmShopEditor.chkSellsItems.value
    Shop(EditorIndex).ShowInfo = frmShopEditor.chkShow.value
    Shop(EditorIndex).currencyItem = currencyItem

    For I = 1 To MAX_SHOP_ITEMS
        Shop(EditorIndex).ShopItem(I).Amount = frmShopEditor.GetShopItemAmt(I)
        Shop(EditorIndex).ShopItem(I).ItemNum = frmShopEditor.GetShopItemNum(I)
        Shop(EditorIndex).ShopItem(I).Price = frmShopEditor.GetShopItemPrice(I)
    Next I

    Call SendSaveShop(EditorIndex)
    InShopEditor = False
    Unload frmShopEditor
End Sub

Public Sub ShopEditorCancel()
    InShopEditor = False
    Unload frmShopEditor
End Sub

Public Sub SpellEditorInit()
    Dim I As Long

    frmSpellEditor.iconn.Picture = LoadPicture(App.Path & "\GFX\Icons.bmp")

    frmSpellEditor.cmbClassReq.addItem "Todas Clases"
    For I = 0 To Max_Classes
        frmSpellEditor.cmbClassReq.addItem Trim$(Class(I).name)
    Next I

    frmSpellEditor.txtName.Text = Trim$(Spell(EditorIndex).name)
    frmSpellEditor.cmbClassReq.ListIndex = Spell(EditorIndex).ClassReq
    frmSpellEditor.scrlLevelReq.value = Spell(EditorIndex).LevelReq

    frmSpellEditor.cmbType.ListIndex = Spell(EditorIndex).Type
    frmSpellEditor.scrlVitalMod.value = Spell(EditorIndex).Data1

    frmSpellEditor.scrlCost.value = Spell(EditorIndex).MPCost
    frmSpellEditor.scrlSound.value = Spell(EditorIndex).Sound

    If Spell(EditorIndex).Range = 0 Then
        Spell(EditorIndex).Range = 1
    End If
    frmSpellEditor.scrlRange.value = Spell(EditorIndex).Range

    frmSpellEditor.scrlSpellAnim.value = Spell(EditorIndex).SpellAnim
    frmSpellEditor.scrlSpellTime.value = Spell(EditorIndex).SpellTime
    frmSpellEditor.scrlSpellDone.value = Spell(EditorIndex).SpellDone

    frmSpellEditor.chkArea.value = Spell(EditorIndex).AE
    frmSpellEditor.chkBig.value = Spell(EditorIndex).Big

    frmSpellEditor.scrlElement.value = Spell(EditorIndex).Element
    frmSpellEditor.scrlElement.max = MAX_ELEMENTS
    
    frmSpellEditor.TTCHScroll.value = Spell(EditorIndex).TimeToCast
    frmSpellEditor.Label12.Caption = Spell(EditorIndex).TimeToCast
    If Spell(EditorIndex).CastTimer > 100 Then
        frmSpellEditor.CTHScroll.value = Spell(EditorIndex).CastTimer
    Else
        frmSpellEditor.CTHScroll.value = frmSpellEditor.CTHScroll.min
    End If
    frmSpellEditor.Label14.Caption = Spell(EditorIndex).CastTimer

    frmSpellEditor.Show vbModal
End Sub

Public Sub SpellEditorOk()
    Spell(EditorIndex).name = frmSpellEditor.txtName.Text
    Spell(EditorIndex).ClassReq = frmSpellEditor.cmbClassReq.ListIndex
    Spell(EditorIndex).LevelReq = frmSpellEditor.scrlLevelReq.value
    Spell(EditorIndex).Type = frmSpellEditor.cmbType.ListIndex
    Spell(EditorIndex).Data1 = frmSpellEditor.scrlVitalMod.value
    Spell(EditorIndex).Data3 = 0
    Spell(EditorIndex).MPCost = frmSpellEditor.scrlCost.value
    Spell(EditorIndex).Sound = frmSpellEditor.scrlSound.value
    Spell(EditorIndex).Range = frmSpellEditor.scrlRange.value

    Spell(EditorIndex).SpellAnim = frmSpellEditor.scrlSpellAnim.value
    Spell(EditorIndex).SpellTime = frmSpellEditor.scrlSpellTime.value
    Spell(EditorIndex).SpellDone = frmSpellEditor.scrlSpellDone.value

    Spell(EditorIndex).AE = frmSpellEditor.chkArea.value
    Spell(EditorIndex).Big = frmSpellEditor.chkBig.value

    Spell(EditorIndex).Element = frmSpellEditor.scrlElement.value
    
    Spell(EditorIndex).TimeToCast = frmSpellEditor.TTCHScroll.value
    Spell(EditorIndex).CastTimer = frmSpellEditor.CTHScroll.value

    Call SendSaveSpell(EditorIndex)
    InSpellEditor = False
    Unload frmSpellEditor
End Sub

Public Sub SpellEditorCancel()
    InSpellEditor = False
    Unload frmSpellEditor
End Sub
