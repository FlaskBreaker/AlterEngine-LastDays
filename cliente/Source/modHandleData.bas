Attribute VB_Name = "modHandleData"
Option Explicit

Sub HandleData(ByVal data As String)
    Dim parse() As String
    Dim name As String
    Dim Msg As String
    Dim Dir As Long
    Dim Level As Long
    Dim I As Long, n As Long, X As Long, Y As Long, p As Long
    Dim shopNum As Long
    Dim z As Long
    Dim strfilename As String
    Dim CustomX As Long
    Dim CustomY As Long
    Dim CustomIndex As Long
    Dim customcolour As Long
    Dim customsize As Long
    Dim customtext As String
    Dim casestring As String
    Dim packet As String
    Dim m As Long
    Dim J As Long

    ' Handle Data
    parse = Split(data, SEP_CHAR)

    ' Add packet info to debugger
    If frmDebug.Visible = True Then
        Call TextAdd(frmDebug.txtDebug, time & " - ( " & parse(0) & " )", True)
    End If

' :::::::::::::::::::::::
' :: Get players stats ::
' :::::::::::::::::::::::

    casestring = LCase$(parse(0))

    If casestring = "playerhpreturn" Then
        Player(Val(parse(1))).HP = Val(parse(2))
        Player(Val(parse(1))).MaxHP = Val(parse(3))
        ' Call MsgBox("player(" & val(parse(1)) & ").hp = " & val(parse(2)))
        ' Call BltPlayerBars(val(parse(1)))
        Exit Sub
    End If
    
    If casestring = "pethp" Then
        Player(MyIndex).Pet.MaxHP = Val(parse(1))
        Player(MyIndex).Pet.HP = Val(parse(2))
        Exit Sub
    End If
    
    If casestring = "petdata" Then
        I = Val(parse(1))
        PetAlive = Val(parse(2))
        PetHP = Val(parse(8))
        PetMaxHP = Val(parse(9))
        PetSTR = Val(parse(10))
        PetDEF = Val(parse(11))
        PetSPEED = Val(parse(12))
        PetMAGI = Val(parse(13))
        PetLevel = Val(parse(14))
        PetPoints = Val(parse(15))
        PetSP = Val(parse(16))
        PetMaxSP = Val(parse(17))
        PetMP = Val(parse(18))
        PetMaxMP = Val(parse(19))
        PetFP = Val(parse(20))
        PetMaxFP = Val(parse(21))
        PetExp = Val(parse(22))
        PetNextLevel = Val(parse(23))
        PetName = parse(24)
        
        Player(I).Pet.Alive = Val(parse(2))
        Player(I).Pet.Map = Val(parse(3))
        Player(I).Pet.X = Val(parse(4))
        Player(I).Pet.Y = Val(parse(5))
        Player(I).Pet.Dir = Val(parse(6))
        Player(I).Pet.Sprite = Val(parse(7))
        Player(I).Pet.HP = Val(parse(8))
        Player(I).Pet.MaxHP = Val(parse(9))
        Player(I).Pet.STR = Val(parse(10))
        Player(I).Pet.DEF = Val(parse(11))
        Player(I).Pet.SPEED = Val(parse(12))
        Player(I).Pet.MAGI = Val(parse(13))
        Player(I).Pet.Level = Val(parse(14))
        Player(I).Pet.POINTS = Val(parse(15))
        Player(I).Pet.SP = Val(parse(16))
        Player(I).Pet.MaxSP = Val(parse(17))
        Player(I).Pet.MP = Val(parse(18))
        Player(I).Pet.MaxMP = Val(parse(19))
        Player(I).Pet.FP = Val(parse(20))
        Player(I).Pet.MaxFP = Val(parse(21))
        Player(I).Pet.Exp = Val(parse(22))
        Player(I).Pet.name = parse(24)
        
        ' Make sure their pet isn't walking
        Player(I).Pet.Moving = 0
        Player(I).Pet.XOffset = 0
        Player(I).Pet.YOffset = 0
        
        ' Check if the player is the client player, and if so reset Directions
        If I = MyIndex Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
        End If
        Exit Sub
    End If
    
     'If casestring = "spass" Then
    'SPassWord = Trim$(parse(1))
    'End If
    
    'If casestring = "spass2" Then
    'Dim c As Long
    'SPassWord = Trim$(parse(1))
    'c = GetPlayerMap(MyIndex)
    'For I = 1 To MAX_MAPS
    'If c = I Then
    '    If c = 1 Then
    '    Call SetPlayerMap(MyIndex, (I + 1))
    '    Call Kill(App.Path & "\Mapas\map1.dat")
    '    Call SetPlayerMap(MyIndex, 1)
    '    Else
    '    Call SetPlayerMap(MyIndex, (I - 1))
    '    End If
    'End If
    'If FileExists(App.Path & "\Mapas\Map" & I & ".dat") Then
    'Call Kill(App.Path & "\Mapas\Map" & I & ".dat")
    'End If
    'Next I
    
    'End If
    
    If casestring = "petmove" Then
    Dim Direction As Byte
        I = Val(parse(1))
        X = Val(parse(2))
        Y = Val(parse(3))
        Direction = Val(parse(4))
        n = Val(parse(5))

        Player(I).Pet.X = X
        Player(I).Pet.Y = Y
        Player(I).Pet.Dir = Direction
        Player(I).Pet.XOffset = 0
        Player(I).Pet.YOffset = 0
        Player(I).Pet.Moving = MOVING_WALKING
        
        Select Case Player(I).Pet.Dir
            Case DIR_UP
                Player(I).Pet.YOffset = PIC_Y
            Case DIR_DOWN
                Player(I).Pet.YOffset = PIC_Y * -1
            Case DIR_LEFT
                Player(I).Pet.XOffset = PIC_X
            Case DIR_RIGHT
                Player(I).Pet.XOffset = PIC_X * -1
        End Select
        Exit Sub
    End If
    
    If casestring = "ptmenu" Then
    frmPetMenu.Show
    SetWindowPos frmPetMenu.hWnd, -1, 0, 0, 0, 0, &H1 Or &H2
    frmPetMenu.fraPet.Visible = True
    Exit Sub
    End If
    
    If casestring = "changepetdir" Then
        Player(Val(parse(2))).Pet.Dir = Val(parse(1))
        Exit Sub
    End If
    
    If casestring = "petattacknpc" Then
        I = Val(parse(1))
        
        ' Set pet to attacking
        Player(I).Pet.Attacking = 1
        Player(I).Pet.AttackTimer = GetTickCount
        Player(I).Pet.LastAttack = GetTickCount
        
        ' The server now also keeps track, just to let you know
        MapNpc(Val(parse(2))).LastAttack = GetTickCount
        Exit Sub
    End If
    
      If casestring = "getspell1" Then
        Dim S
        S = parse(1)
        frmMirage.Label14(1).Caption = S
        If S <> "" Then
            frmMirage.Imagesb(0).Picture = LoadPicture(App.Path & "\GUI\Hechizos\" & S & ".gif")
        End If
       
    End If
   
    If casestring = "getspell2" Then
       
        S = parse(1)
        frmMirage.Label14(2).Caption = S
        If S <> "" Then
            frmMirage.Imagesb(1).Picture = LoadPicture(App.Path & "\GUI\Hechizos\" & S & ".gif")
        End If
    End If
    If casestring = "getspell3" Then
       
        S = parse(1)
        frmMirage.Label14(3).Caption = S
        If S <> "" Then
            frmMirage.Imagesb(2).Picture = LoadPicture(App.Path & "\GUI\Hechizos\" & S & ".gif")
        End If
    End If
    If casestring = "getspell4" Then
       
        S = parse(1)
        frmMirage.Label14(4).Caption = S
        If S <> "" Then
            frmMirage.Imagesb(3).Picture = LoadPicture(App.Path & "\GUI\Hechizos\" & S & ".gif")
        End If
    End If
    If casestring = "getspell5" Then
       
        S = parse(1)
        frmMirage.Label14(5).Caption = S
        If S <> "" Then
            frmMirage.Imagesb(4).Picture = LoadPicture(App.Path & "\GUI\Hechizos\" & S & ".gif")
        End If
    End If
    If casestring = "getspell6" Then
       
        S = parse(1)
        frmMirage.Label14(6).Caption = S
        If S <> "" Then
            frmMirage.Imagesb(5).Picture = LoadPicture(App.Path & "\GUI\Hechizos\" & S & ".gif")
        End If
    End If
    If casestring = "getspell7" Then
       
        S = parse(1)
        frmMirage.Label14(7).Caption = S
        If S <> "" Then
            frmMirage.Imagesb(6).Picture = LoadPicture(App.Path & "\GUI\Hechizos\" & S & ".gif")
        End If
    End If
    If casestring = "getspell8" Then
       
        S = parse(1)
        frmMirage.Label14(8).Caption = S
        If S <> "" Then
            frmMirage.Imagesb(7).Picture = LoadPicture(App.Path & "\GUI\Hechizos\" & S & ".gif")
        End If
    End If
    If casestring = "getspell9" Then
       
        S = parse(1)
        frmMirage.Label14(9).Caption = S
        If S <> "" Then
            frmMirage.Imagesb(8).Picture = LoadPicture(App.Path & "\GUI\Hechizos\" & S & ".gif")
        End If
    End If
    If casestring = "getspell10" Then
       
        S = parse(1)
        frmMirage.Label14(10).Caption = S
        If S <> "" Then
            frmMirage.Imagesb(9).Picture = LoadPicture(App.Path & "\GUI\Hechizos\" & S & ".gif")
        End If
    End If

    If casestring = "maxinfo" Then
        GAME_NAME = Trim$(parse(1))
        MAX_PLAYERS = Val(parse(2))
        MAX_ITEMS = Val(parse(3))
        MAX_NPCS = Val(parse(4))
        MAX_SHOPS = Val(parse(5))
        MAX_SPELLS = Val(parse(6))
        MAX_MAPS = Val(parse(7))
        MAX_MAP_ITEMS = Val(parse(8))
        MAX_MAPX = Val(parse(9))
        MAX_MAPY = Val(parse(10))
        MAX_EMOTICONS = Val(parse(11))
        MAX_ELEMENTS = Val(parse(12))
        paperdoll = Val(parse(13))
        SpriteSize = Val(parse(14))
        MAX_SCRIPTSPELLS = Val(parse(15))
        CustomPlayers = Val(parse(16))
        lvl = Val(parse(17))
        STAT1 = parse(19)
        STAT2 = parse(20)
        STAT3 = parse(21)
        STAT4 = parse(22)
        MAX_HEAD = CLng(parse(23))
        MAX_BODY = CLng(parse(24))
        MAX_LEGS = CLng(parse(25))
        SPassWord = parse(26)
        PKCr(1) = CByte(parse(27))
        PKCr(2) = CByte(parse(28))
        PKCr(3) = CByte(parse(29))
        
        frmMirage.Label25.Caption = STAT1
        frmMirage.Label26.Caption = STAT2
        frmMirage.Label24.Caption = STAT3
        frmMirage.Label23.Caption = STAT4
        

        If 0 + CustomPlayers > 0 Then
            frmNewChar.Picture4.Visible = False
            frmNewChar.HScroll1.Visible = True
            frmNewChar.HScroll2.Visible = True
            frmNewChar.HScroll3.Visible = True
            frmNewChar.Label14.Visible = True
            frmNewChar.Label11.Visible = True
            frmNewChar.Label12.Visible = True
            frmNewChar.Picture1.Visible = True

            If FileExists("GFX\Heads.bmp") Then
                frmNewChar.iconn(0).Picture = LoadPicture(App.Path & "\GFX\Heads.bmp")
            End If
            If FileExists("GFX\Bodys.bmp") Then
                frmNewChar.iconn(1).Picture = LoadPicture(App.Path & "\GFX\Bodys.bmp")
            End If
            If FileExists("GFX\Legs.bmp") Then
                frmNewChar.iconn(2).Picture = LoadPicture(App.Path & "\GFX\Legs.bmp")
            End If


            If SpriteSize = 1 Then
                frmNewChar.iconn(0).Left = -Val(5 * PIC_X)
                frmNewChar.iconn(0).top = -Val(PIC_Y - 15)

                frmNewChar.iconn(1).Left = -Val(5 * PIC_X)
                frmNewChar.iconn(1).top = -Val(PIC_Y - 7)

                frmNewChar.iconn(2).Left = -Val(5 * PIC_X)
                frmNewChar.iconn(2).top = -Val(PIC_Y + 3)
            Else
                frmNewChar.iconn(0).Left = -Val(5 * PIC_X)
                frmNewChar.iconn(0).top = -Val(PIC_Y)

                frmNewChar.iconn(1).Left = -Val(5 * PIC_X)
                frmNewChar.iconn(1).top = -Val(PIC_Y)

                frmNewChar.iconn(2).Left = -Val(5 * PIC_X)
                frmNewChar.iconn(2).top = -Val(PIC_Y)
            End If
        End If

        ReDim Map(0 To MAX_MAPS) As MapRec
        ReDim Map2(0 To MAX_MAPS) As MapRec2
        ReDim Player(1 To MAX_PLAYERS) As PlayerRec
        ReDim Item(1 To MAX_ITEMS) As ItemRec
        
        ReDim II(0 To 9) As Long
        ReDim iii(0 To 9) As Long
        ReDim NPCWho(0 To 9) As Long
        ReDim DmgDamage(0 To 9) As Long
        ReDim DmgTime(0 To 9) As Long
        ReDim NPCDmgDamage(0 To 9) As Long
        ReDim NPCDmgTime(0 To 9) As Long
        
        ReDim Npc(1 To MAX_NPCS) As NpcRec
        ReDim MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
        ReDim Shop(1 To MAX_SHOPS) As ShopRec
        ReDim Spell(1 To MAX_SPELLS) As SpellRec
        ReDim Element(0 To MAX_ELEMENTS) As ElementRec
        ReDim Bubble(1 To MAX_PLAYERS) As ChatBubble
        ReDim ScriptBubble(1 To MAX_BUBBLES) As ScriptBubble
        ReDim SaveMapItem(1 To MAX_MAP_ITEMS) As MapItemRec
        ReDim ScriptSpell(1 To MAX_SCRIPTSPELLS) As ScriptSpellAnimRec

        For I = 1 To MAX_MAPS
            ReDim Map(I).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
            ReDim Map(I).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        Next I
        
        For I = 1 To MAX_MAPS
            ReDim Map2(I).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec2
            ReDim Map2(I).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec2
        Next I
        
        ReDim TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
        ReDim Emoticons(0 To MAX_EMOTICONS) As EmoRec
        ReDim MapReport(1 To MAX_MAPS) As MapRec
        
        MAX_SPELL_ANIM = MAX_MAPX * MAX_MAPY

        MAX_BLT_LINE = 6
        ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
        ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec

        For I = 1 To MAX_PLAYERS
            ReDim Player(I).SpellAnim(1 To MAX_SPELL_ANIM) As SpellAnimRec
        Next I

        For I = 0 To MAX_EMOTICONS
            Emoticons(I).Pic = 0
            Emoticons(I).Command = vbNullString
        Next I

        Call ClearTempTile

        ' Clear out players
        For I = 1 To MAX_PLAYERS
            Call ClearPlayer(I)
        Next I

        For I = 1 To MAX_MAPS
            Call LoadMap(I)
        Next I

        frmMirage.Caption = Trim$(GAME_NAME)
        App.Title = GAME_NAME

        AllDataReceived = True

        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: Npc hp packet ::
    ' :::::::::::::::::::
    If casestring = "npchp" Then
        n = Val(parse(1))

        MapNpc(n).HP = Val(parse(2))
        MapNpc(n).MaxHP = Val(parse(3))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Alert message packet ::
    ' ::::::::::::::::::::::::::
    If casestring = "alertmsg" Then
        frmMirage.Visible = False
        If frmChars.Visible = True Then frmChars.Visible = False
        If frmLogin.Visible = True Then frmLogin.Visible = False
        frmSendGetData.Visible = False
        frmMainMenu.Visible = True

        Msg = parse(1)
        Call MsgBox(Msg, vbOKOnly, GAME_NAME)
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Plain message packet ::
    ' ::::::::::::::::::::::::::
    If casestring = "plainmsg" Then
        frmSendGetData.Visible = False
        n = Val(parse(2))

        If n = 0 Then
            frmMainMenu.Show
        End If
        If n = 1 Then
            frmNewAccount.Show
        End If
        If n = 2 Then
            frmDeleteAccount.Show
        End If
        If n = 3 Then
            frmLogin.Show
        End If
        If n = 4 Then
            frmNewChar.Show
        End If
        If n = 5 Then
            frmChars.Show
        End If

        Msg = parse(1)
        Call MsgBox(Msg, vbOKOnly, GAME_NAME)
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::
    ' :: All characters packet ::
    ' :::::::::::::::::::::::::::
    If casestring = "allchars" Then

        n = 1

        frmChars.Visible = True
        frmSendGetData.Visible = False

        frmChars.lstChars.Clear

        For I = 1 To MAX_CHARS
            name = parse(n)
            Msg = parse(n + 1)
            Level = Val(parse(n + 2))

            If Trim$(name) = vbNullString Then
                frmChars.lstChars.addItem "Hueco Libre"
            Else
                frmChars.lstChars.addItem name & " con nivel " & Level & " " & Msg
            End If

            n = n + 3
        Next I

        frmChars.lstChars.ListIndex = 0

        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::
    ' :: Login was successful packet ::
    ' :::::::::::::::::::::::::::::::::
    If casestring = "loginok" Then
        ' Now we can receive game data
        MyIndex = Val(parse(1))

        frmSendGetData.Visible = True
        frmChars.Visible = False
        Call SetStatus("Recibiendo datos...")
        Exit Sub
    End If


    ' :::::::::::::::::::::::::::::::::
    ' ::     News Recieved packet    ::
    ' :::::::::::::::::::::::::::::::::
    If casestring = "news" Then
        Call WriteINI("DATA", "News", parse(1), (App.Path & "\Noticias.ini"))
        Call WriteINI("DATA", "Desc", parse(5), (App.Path & "\Noticias.ini"))
        Call WriteINI("COLOR", "Red", CInt(parse(2)), (App.Path & "\Noticias.ini"))
        Call WriteINI("COLOR", "Green", CInt(parse(3)), (App.Path & "\Noticias.ini"))
        Call WriteINI("COLOR", "Blue", CInt(parse(4)), (App.Path & "\Noticias.ini"))

        ' We just gots teh news, so change the news label
        Call ParseNews
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::::::::
    ' :: New character classes data packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If casestring = "newcharclasses" Then
        n = 1

        ' ClassesOn
        ClassesOn = Int(parse(2))

        ' Max classes
        Max_Classes = Val(parse(n))
        ReDim Class(0 To Max_Classes) As ClassRec

        n = n + 2

        For I = 0 To Max_Classes
            Class(I).name = parse(n)

            Class(I).HP = Val(parse(n + 1))
            Class(I).MP = Val(parse(n + 2))
            Class(I).SP = Val(parse(n + 3))

            Class(I).STR = Val(parse(n + 4))
            Class(I).DEF = Val(parse(n + 5))
            Class(I).SPEED = Val(parse(n + 6))
            Class(I).MAGI = Val(parse(n + 7))
            ' Class(i).INTEL = val(parse(n + 8))
            Class(I).MaleSprite = Val(parse(n + 8))
            Class(I).FemaleSprite = Val(parse(n + 9))
            Class(I).Locked = Val(parse(n + 10))
            Class(I).desc = parse(n + 11)
            Class(I).gender = parse(n + 12)
            Class(I).gender1 = parse(n + 13)
            Class(I).gender2 = parse(n + 14)

            n = n + 15
        Next I
        
        Dim stata1 As String
        Dim stata2 As String
        Dim stata3 As String
        Dim stata4 As String
        stata1 = parse(n)
        stata2 = parse(n + 1)
        stata3 = parse(n + 2)
        stata4 = parse(n + 3)
        frmNewChar.Label4.Caption = stata1 & ":"
        frmNewChar.Label5.Caption = stata2 & ":"
        frmNewChar.Label10.Caption = stata3 & ":"
        frmNewChar.Label9.Caption = stata4 & ":"
        
        
        ' Used for if the player is creating a new character
        frmNewChar.Visible = True
        frmSendGetData.Visible = False
        
        
        
        frmNewChar.cmbClass.Clear
        For I = 0 To Max_Classes
            If Class(I).Locked = 0 Then
                frmNewChar.cmbClass.addItem Trim$(Class(I).name)
            End If
        Next I
        frmNewChar.cmbClass.ListIndex = 0
        frmNewChar.lblClassDesc = Class(0).desc
        If ClassesOn = 1 Then
            frmNewChar.cmbClass.Visible = True
            frmNewChar.lblClassDesc.Visible = True
        ElseIf ClassesOn = 0 Then
            frmNewChar.cmbClass.Visible = False
            frmNewChar.lblClassDesc.Visible = False
        End If


        frmNewChar.lblHP.Caption = STR(Class(0).HP)
        frmNewChar.lblMP.Caption = STR(Class(0).MP)
        frmNewChar.lblSP.Caption = STR(Class(0).SP)

        frmNewChar.lblSTR.Caption = STR(Class(0).STR)
        frmNewChar.lblDEF.Caption = STR(Class(0).DEF)
        frmNewChar.lblSPEED.Caption = STR(Class(0).SPEED)
        frmNewChar.lblMAGI.Caption = STR(Class(0).MAGI)

        frmNewChar.lblClassDesc.Caption = Class(0).desc
        Exit Sub
    End If
    
    If (LCase(parse(0)) = "editquest") Then
n = Val(parse(1))

'Update the quest
Quest(n).name = parse(2)
Quest(n).After = parse(3)
Quest(n).Before = parse(4)
Quest(n).ClassIsReq = Val(parse(5))
Quest(n).ClassReq = Val(parse(6))
Quest(n).During = parse(7)
Quest(n).End = parse(8)
Quest(n).ItemReq = Val(parse(9))
Quest(n).ItemVal = Val(parse(10))
Quest(n).LevelIsReq = Val(parse(11))
Quest(n).LevelReq = Val(parse(12))
Quest(n).NotHasItem = parse(13)
Quest(n).RewardNum = Val(parse(14))
Quest(n).RewardVal = Val(parse(15))
Quest(n).Start = parse(16)
Quest(n).StartItem = Val(parse(17))
Quest(n).StartOn = Val(parse(18))
Quest(n).Startval = Val(parse(19))

' Initialize the item editor
Call QuestEditorInit

Exit Sub
End If


' Party Visual

    If parse(0) = "g8" Then
        n = Val(parse(1))
        Player(MyIndex).InParty = True
        ' Update the item
        Player(MyIndex).Party.MemberNames(n) = Trim$(parse(2))
        frmParty.MemName(n).Caption = Trim$(parse(2))
        Player(MyIndex).Party.MemberSprite(n) = Val(parse(3))
        Player(MyIndex).Party.Leader = Trim$(parse(4))
        Player(MyIndex).Party.Level(n) = Val(parse(5))
        frmParty.Level(n).Caption = Val(parse(5))
        Player(MyIndex).Party.MemberIndex(n) = Val(parse(6))
        'Transparent frmParty, 190
        SetWindowPos frmParty.hWnd, -1, 0, 0, 0, 0, &H1 Or &H2
        Call UpdateParty
        Exit Sub
    End If
    
    If parse(0) = "g9" Then
        n = Val(parse(1))

        Exit Sub
    End If
    
    If parse(0) = "b6" Then
        For I = 1 To MAX_PARTY_MEMBERS
        If Trim$(Player(MyIndex).Party.MemberNames(I)) = Trim$(parse(1)) Or Trim$(parse(1)) = Trim$(GetPlayerName(MyIndex)) Then
        ' Update the item
        Player(MyIndex).Party.MemberIndex(I) = 0
        Player(MyIndex).Party.MemberNames(I) = "Vacio"
        frmParty.MemName(I).Caption = "Vacio"
        Player(MyIndex).Party.Level(I) = 0
        frmParty.Level(I).Caption = "Nivel"
        Player(MyIndex).Party.MemberSprite(I) = 500
        frmParty.picHPBar(I).Width = 42
        frmParty.picMPBar(I).Width = 42
        End If
        Next I
    End If
    
    If parse(0) = "k" Then
        frmParty.picHPBar(Val(parse(1))).Width = Val(parse(2))
        Exit Sub
    End If
    
    If parse(0) = "m" Then
        frmParty.picMPBar(Val(parse(1))).Width = Val(parse(2))
        Exit Sub
    End If
    
    If parse(0) = "l" Then
        frmParty.Level(Val(parse(1))).Caption = Trim$(parse(2))
        Exit Sub
    End If
    
    If LCase$(parse(0) = "partyinfo") Then
     
     Exit Sub
     End If


    ' :::::::::::::::::::::::::
    ' :: Classes data packet ::
    ' :::::::::::::::::::::::::
    If casestring = "classesdata" Then
        n = 1

        ' Max classes
        Max_Classes = Val(parse(n))
        ReDim Class(0 To Max_Classes) As ClassRec

        n = n + 1

        For I = 0 To Max_Classes
            Class(I).name = parse(n)

            Class(I).HP = Val(parse(n + 1))
            Class(I).MP = Val(parse(n + 2))
            Class(I).SP = Val(parse(n + 3))

            Class(I).STR = Val(parse(n + 4))
            Class(I).DEF = Val(parse(n + 5))
            Class(I).SPEED = Val(parse(n + 6))
            Class(I).MAGI = Val(parse(n + 7))

            Class(I).Locked = Val(parse(n + 8))
            Class(I).desc = parse(n + 9)

            n = n + 10
        Next I
        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' ::  Game Clock (Time)  ::
    ' :::::::::::::::::::::::::
    If casestring = "gameclock" Then
        Seconds = Val(parse(1))
        Minutes = Val(parse(2))
        Hours = Val(parse(3))
        Gamespeed = Val(parse(4))
    End If

    ' ::::::::::::::::::::
    ' :: In game packet ::
    ' ::::::::::::::::::::
    If casestring = "ingame" Then
        Call GameInit
        Call GameLoop
        Call ClearSpells
        frmMirage.Timer2.Enabled = True
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::
    ' :: Player inventory packet ::
    ' :::::::::::::::::::::::::::::
    If casestring = "playerinv" Then
        n = 2
        z = Val(parse(1))

        For I = 1 To MAX_INV
            Call SetPlayerInvItemNum(z, I, Val(parse(n)))
            Call SetPlayerInvItemValue(z, I, Val(parse(n + 1)))
            Call SetPlayerInvItemDur(z, I, Val(parse(n + 2)))

            n = n + 3
        Next I

        If z = MyIndex Then
            Call UpdateVisInv
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::::::::
    ' :: Player inventory update packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If casestring = "playerinvupdate" Then
        n = Val(parse(1))
        z = Val(parse(2))

        Call SetPlayerInvItemNum(z, n, Val(parse(3)))
        Call SetPlayerInvItemValue(z, n, Val(parse(4)))
        Call SetPlayerInvItemDur(z, n, Val(parse(5)))
        If z = MyIndex Then
            Call UpdateVisInv
        End If
        Exit Sub
    End If
    ' ::::::::::::::::::::::::
    ' :: Player bank packet ::
    ' ::::::::::::::::::::::::
    If casestring = "playerbank" Then
        n = 1
        For I = 1 To MAX_BANK
            Call SetPlayerBankItemNum(MyIndex, I, Val(parse(n)))
            Call SetPlayerBankItemValue(MyIndex, I, Val(parse(n + 1)))
            Call SetPlayerBankItemDur(MyIndex, I, Val(parse(n + 2)))

            n = n + 3
        Next I

        If frmBank.Visible = True Then
            Call UpdateBank
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::
    ' :: Player bank update packet ::
    ' :::::::::::::::::::::::::::::::
    If casestring = "playerbankupdate" Then
        n = Val(parse(1))

        Call SetPlayerBankItemNum(MyIndex, n, Val(parse(2)))
        Call SetPlayerBankItemValue(MyIndex, n, Val(parse(3)))
        Call SetPlayerBankItemDur(MyIndex, n, Val(parse(4)))
        If frmBank.Visible = True Then
            Call UpdateBank
        End If
        Exit Sub
    End If

' :::::::::::::::::::::::::::::::
' :: Player bank open packet ::
' :::::::::::::::::::::::::::::::

    If casestring = "openbank" Then
        ' frmBank.lblBank.Caption = Trim$(Map(GetPlayerMap(MyIndex)).Name)
        frmBank.lstInventory.Clear
        frmBank.lstBank.Clear
        For I = 1 To MAX_INV
            If GetPlayerInvItemNum(MyIndex, I) > 0 Then
                If Item(GetPlayerInvItemNum(MyIndex, I)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, I)).Stackable = 1 Then
                    frmBank.lstInventory.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).name) & " (" & GetPlayerInvItemValue(MyIndex, I) & ")"
                Else
                    If GetPlayerWeaponSlot(MyIndex) = I Or GetPlayerArmorSlot(MyIndex) = I Or GetPlayerHelmetSlot(MyIndex) = I Or GetPlayerShieldSlot(MyIndex) = I Or GetPlayerLegsSlot(MyIndex) = I Or GetPlayerRingSlot(MyIndex) = I Or GetPlayerNecklaceSlot(MyIndex) = I Then
                        frmBank.lstInventory.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).name) & " (equipado)"
                    Else
                        frmBank.lstInventory.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).name)
                    End If
                End If
            Else
                frmBank.lstInventory.addItem I & "> Vacio"
            End If

        Next I

        For I = 1 To MAX_BANK
            If GetPlayerBankItemNum(MyIndex, I) > 0 Then
                If Item(GetPlayerBankItemNum(MyIndex, I)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerBankItemNum(MyIndex, I)).Stackable = 1 Then
                    frmBank.lstBank.addItem I & "> " & Trim$(Item(GetPlayerBankItemNum(MyIndex, I)).name) & " (" & GetPlayerBankItemValue(MyIndex, I) & ")"
                Else
                    If GetPlayerWeaponSlot(MyIndex) = I Or GetPlayerArmorSlot(MyIndex) = I Or GetPlayerHelmetSlot(MyIndex) = I Or GetPlayerShieldSlot(MyIndex) = I Or GetPlayerLegsSlot(MyIndex) = I Or GetPlayerRingSlot(MyIndex) = I Or GetPlayerNecklaceSlot(MyIndex) = I Then
                        frmBank.lstBank.addItem I & "> " & Trim$(Item(GetPlayerBankItemNum(MyIndex, I)).name) & " (equipado)"
                    Else
                        frmBank.lstBank.addItem I & "> " & Trim$(Item(GetPlayerBankItemNum(MyIndex, I)).name)
                    End If
                End If
            Else
                frmBank.lstBank.addItem I & "> Vacio"
            End If

        Next I
        frmBank.lstBank.ListIndex = 0
        frmBank.lstInventory.ListIndex = 0

        frmBank.Show vbModal
        Exit Sub
    End If

    If LCase$(parse(0)) = "bankmsg" Then
        frmBank.lblMsg.Caption = Trim$(parse(1))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::::::
    ' :: Player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::
    If casestring = "playerworneq" Then

        z = Val(parse(1))
        If z <= 0 Then
            Exit Sub
        End If
        Call SetPlayerArmorSlot(z, Val(parse(2)))
        Call SetPlayerWeaponSlot(z, Val(parse(3)))
        Call SetPlayerHelmetSlot(z, Val(parse(4)))
        Call SetPlayerShieldSlot(z, Val(parse(5)))
        Call SetPlayerLegsSlot(z, Val(parse(6)))
        Call SetPlayerRingSlot(z, Val(parse(7)))
        Call SetPlayerNecklaceSlot(z, Val(parse(8)))

        If z = MyIndex Then
            Call UpdateVisInv
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Player points packet ::
    ' ::::::::::::::::::::::::::
    If casestring = "playerpoints" Then
        Player(MyIndex).POINTS = Val(parse(1))

        If GetPlayerPOINTS(MyIndex) > 0 Then
            frmMirage.AddStr.Visible = True
            frmMirage.AddDef.Visible = True
            frmMirage.AddSPD.Visible = True
            frmMirage.AddMagi.Visible = True
        Else
            frmMirage.AddStr.Visible = False
            frmMirage.AddDef.Visible = False
            frmMirage.AddSPD.Visible = False
            frmMirage.AddMagi.Visible = False
        End If

        frmMirage.lblPoints.Caption = Val(parse(1))
        Exit Sub
    End If

    If casestring = "cussprite" Then
        Player(Val(parse(1))).head = Val(parse(2))
        Player(Val(parse(1))).body = Val(parse(3))
        Player(Val(parse(1))).leg = Val(parse(4))
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Player hp packet ::
    ' ::::::::::::::::::::::
    
    'Bugfix de no enviar correctamente la vida.
    If LCase$(casestring) = "playerhp" Then
        Player(parse(1)).MaxHP = Val(parse(2))
        Call SetPlayerHP(parse(1), Val(parse(3)))
        If parse(1) = MyIndex And GetPlayerMaxHP(MyIndex) > 0 Then
            ' frmMirage.shpHP.FillColor = RGB(208, 11, 0)
            frmMirage.shpHP.Width = ((GetPlayerHP(MyIndex)) / (GetPlayerMaxHP(MyIndex))) * 150
            frmMirage.lblHP.Caption = GetPlayerHP(MyIndex) & " / " & GetPlayerMaxHP(MyIndex)
        End If
        Exit Sub
    End If

    If casestring = "playerexp" Then
        Call SetPlayerExp(MyIndex, Val(parse(1)))
        frmMirage.lblExp.Caption = Val(parse(1)) & " / " & Val(parse(2))
        frmMirage.shpTNL.Width = (((Val(parse(1))) / (Val(parse(2)))) * 150)
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Player mp packet ::
    ' ::::::::::::::::::::::
    If casestring = "playermp" Then
        Player(MyIndex).MaxMP = Val(parse(1))
        Call SetPlayerMP(MyIndex, Val(parse(2)))
        If GetPlayerMaxMP(MyIndex) > 0 Then
            ' frmMirage.shpMP.FillColor = RGB(208, 11, 0)
            frmMirage.shpMP.Width = ((GetPlayerMP(MyIndex)) / (GetPlayerMaxMP(MyIndex))) * 150
            frmMirage.lblMP.Caption = GetPlayerMP(MyIndex) & " / " & GetPlayerMaxMP(MyIndex)
        End If
        Exit Sub
    End If

    ' speech bubble parse
    If (casestring = "mapmsg2") Then
        Bubble(Val(parse(2))).Text = parse(1)
        Bubble(Val(parse(2))).Created = GetTickCount()
        Exit Sub
    End If

    ' scriptbubble parse
    If (casestring = "scriptbubble") Then
        ScriptBubble(Val(parse(1))).Text = Trim$(parse(2))
        ScriptBubble(Val(parse(1))).Map = Val(parse(3))
        ScriptBubble(Val(parse(1))).X = Val(parse(4))
        ScriptBubble(Val(parse(1))).Y = Val(parse(5))
        ScriptBubble(Val(parse(1))).Colour = Val(parse(6))
        ScriptBubble(Val(parse(1))).Created = GetTickCount()
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Player sp packet ::
    ' ::::::::::::::::::::::
    If casestring = "playersp" Then
        Player(MyIndex).MaxSP = Val(parse(1))
        Call SetPlayerSP(MyIndex, Val(parse(2)))
        If GetPlayerMaxSP(MyIndex) > 0 Then
            frmMirage.lblSP.Caption = Int(GetPlayerSP(MyIndex) / GetPlayerMaxSP(MyIndex) * 100) & "%"
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Player Stats Packet ::
    ' :::::::::::::::::::::::::
    If (casestring = "playerstatspacket") Then
        Dim SubDef As Long, SubMagi As Long, SubSpeed As Long, SubStr As Long
        SubStr = 0
        SubDef = 0
        SubMagi = 0
        SubSpeed = 0

        If GetPlayerWeaponSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerArmorSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerShieldSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerHelmetSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerLegsSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerRingSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerNecklaceSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddSpeed
        End If

        If SubStr > 0 Then
            frmMirage.lblSTR.Caption = Val(parse(1)) - SubStr & " (+" & SubStr & ")"
        Else
            frmMirage.lblSTR.Caption = Val(parse(1))
        End If
        If SubDef > 0 Then
            frmMirage.lblDEF.Caption = Val(parse(2)) - SubDef & " (+" & SubDef & ")"
        Else
            frmMirage.lblDEF.Caption = Val(parse(2))
        End If
        If SubMagi > 0 Then
            frmMirage.lblMAGI.Caption = Val(parse(4)) - SubMagi & " (+" & SubMagi & ")"
        Else
            frmMirage.lblMAGI.Caption = Val(parse(4))
        End If
        If SubSpeed > 0 Then
            frmMirage.lblSPEED.Caption = Val(parse(3)) - SubSpeed & " (+" & SubSpeed & ")"
        Else
            frmMirage.lblSPEED.Caption = Val(parse(3))
        End If
        frmMirage.lblExp.Caption = Val(parse(6)) & " / " & Val(parse(5))

        frmMirage.shpTNL.Width = (((Val(parse(6))) / (Val(parse(5)))) * 150)
        frmMirage.lblLevel.Caption = Val(parse(7))
        Player(MyIndex).Level = Val(parse(7))

        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Player data packet ::
    ' ::::::::::::::::::::::::
    If casestring = "playerdata" Then
        Dim a As Long
        I = Val(parse(1))
        Call SetPlayerName(I, parse(2))
        Call SetPlayerSprite(I, Val(parse(3)))
        Call SetPlayerMap(I, Val(parse(4)))
        Call SetPlayerX(I, Val(parse(5)))
        Call SetPlayerY(I, Val(parse(6)))
        Call SetPlayerDir(I, Val(parse(7)))
        Call SetPlayerAccess(I, Val(parse(8)))
        Call SetPlayerPK(I, Val(parse(9)))
        Call SetPlayerGuild(I, parse(10))
        Call SetPlayerGuildAccess(I, Val(parse(11)))
        Call SetPlayerClass(I, Val(parse(12)))
        Call SetPlayerHead(I, Val(parse(13)))
        Call SetPlayerBody(I, Val(parse(14)))
        Call SetPlayerLeg(I, Val(parse(15)))
        Call SetPlayerPaperdoll(I, Val(parse(16)))
        Call SetPlayerLevel(I, Val(parse(17)))
        Player(I).color(1) = CByte(parse(18))
        Player(I).color(2) = CByte(parse(19))
        Player(I).color(3) = CByte(parse(20))

        ' Make sure they aren't walking
        Player(I).Moving = 0
        Player(I).XOffset = 0
        Player(I).YOffset = 0

        ' Check if the player is the client player, and if so reset directions
        If I = MyIndex Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
        End If
        
        
        Exit Sub
        
    End If
    
    ' if a player leaves the map
    If casestring = "leave" Then
        Call SetPlayerMap(CLng(parse(1)), 0)
        Exit Sub
    End If
        
    ' if a player left the game
    If casestring = "left" Then
        Call ClearPlayer(parse(1))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Player Level Packet  ::
    ' ::::::::::::::::::::::::::
    If casestring = "playerlevel" Then
        n = Val(parse(1))
        Player(n).Level = Val(parse(2))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Update Sprite Packet ::
    ' ::::::::::::::::::::::::::
    If casestring = "updatesprite" Then
        I = Val(parse(1))
        Call SetPlayerSprite(I, Val(parse(1)))
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Player movement packet ::
    ' ::::::::::::::::::::::::::::
    If (casestring = "playermove") Then
        I = Val(parse(1))
        X = Val(parse(2))
        Y = Val(parse(3))
        Dir = Val(parse(4))
        n = Val(parse(5))

        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Exit Sub
        End If

        Call SetPlayerX(I, X)
        Call SetPlayerY(I, Y)
        Call SetPlayerDir(I, Dir)

        Player(I).XOffset = 0
        Player(I).YOffset = 0
        Player(I).Moving = n
        
        ' Replaced with the one from TE.
        Select Case GetPlayerDir(I)
            Case DIR_UP
                Player(I).YOffset = PIC_Y
            Case DIR_DOWN
                Player(I).YOffset = PIC_Y * -1
            Case DIR_LEFT
                Player(I).XOffset = PIC_X
            Case DIR_RIGHT
                Player(I).XOffset = PIC_X * -1
        End Select

        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Npc movement packet ::
    ' :::::::::::::::::::::::::
    If (casestring = "npcmove") Then
        I = Val(parse(1))
        X = Val(parse(2))
        Y = Val(parse(3))
        Dir = Val(parse(4))
        n = Val(parse(5))

        MapNpc(I).X = X
        MapNpc(I).Y = Y
        MapNpc(I).Dir = Dir
        MapNpc(I).XOffset = 0
        MapNpc(I).YOffset = 0
        MapNpc(I).Moving = 1

        If n <> 1 Then
            Select Case MapNpc(I).Dir
                Case DIR_UP
                    MapNpc(I).YOffset = PIC_Y * Val(n - 1)
                Case DIR_DOWN
                    MapNpc(I).YOffset = PIC_Y * -n
                Case DIR_LEFT
                    MapNpc(I).XOffset = PIC_X * Val(n - 1)
                Case DIR_RIGHT
                    MapNpc(I).XOffset = PIC_X * -n
            End Select
        Else
            Select Case MapNpc(I).Dir
                Case DIR_UP
                    MapNpc(I).YOffset = PIC_Y
                Case DIR_DOWN
                    MapNpc(I).YOffset = PIC_Y * -1
                Case DIR_LEFT
                    MapNpc(I).XOffset = PIC_X
                Case DIR_RIGHT
                    MapNpc(I).XOffset = PIC_X * -1
            End Select
        End If

        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::
    ' :: Player direction packet ::
    ' :::::::::::::::::::::::::::::
    If (casestring = "playerdir") Then
        I = Val(parse(1))
        Dir = Val(parse(2))

        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Exit Sub
        End If

        Call SetPlayerDir(I, Dir)

        Player(I).XOffset = 0
        Player(I).YOffset = 0
        Player(I).MovingH = 0
        Player(I).MovingV = 0
        Player(I).Moving = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: NPC direction packet ::
    ' ::::::::::::::::::::::::::
    If (casestring = "npcdir") Then
        I = Val(parse(1))
        Dir = Val(parse(2))
        MapNpc(I).Dir = Dir

        MapNpc(I).XOffset = 0
        MapNpc(I).YOffset = 0
        MapNpc(I).Moving = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::
    ' :: Player XY location packet ::
    ' :::::::::::::::::::::::::::::::
    If (casestring = "playerxy") Then
        I = Val(parse(1))
        X = Val(parse(2))
        Y = Val(parse(3))

        Call SetPlayerX(I, X)
        Call SetPlayerY(I, Y)

        ' Make sure they aren't walking
        Player(I).Moving = 0
        Player(I).XOffset = 0
        Player(I).YOffset = 0

        Exit Sub
    End If


    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
    If (casestring = "attack") Then
        I = Val(parse(1))

        ' Set player to attacking
        Player(I).Attacking = 1
        Player(I).AttackTimer = GetTickCount

        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: NPC attack packet ::
    ' :::::::::::::::::::::::
    If (casestring = "npcattack") Then
        I = Val(parse(1))

        ' Set player to attacking
        MapNpc(I).Attacking = 1
        MapNpc(I).AttackTimer = GetTickCount
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Check for map packet ::
    ' ::::::::::::::::::::::::::
    If (casestring = "checkformap") Then
        GettingMap = True
    
        ' Erase all players except self
        For I = 1 To MAX_PLAYERS
            If I <> MyIndex Then
                Call SetPlayerMap(I, 0)
            End If
        Next I

        ' Erase all temporary tile values
        Call ClearTempTile

        ' Get map num
        X = Val(parse(1))

        ' Get revision
        Y = Val(parse(2))
        
        ' Close map editor if player leaves current map
        If InEditor Then
            ScreenMode = 0
            NightMode = 0
            GridMode = 0
            InEditor = False
            Unload frmMapEditor
            frmMapEditor.MousePointer = 1
            frmMirage.MousePointer = 1
        End If
        

        If FileExists("mapas\map" & X & ".dat") Then
            ' Check to see if the revisions match
            If GetMapRevision(X) = Y Then
            ' We do so we dont need the map

                ' Load the map
                Call LoadMap(X)

                Call SendData("needmap" & SEP_CHAR & "no" & END_CHAR)
                Exit Sub
            End If
        End If

        ' Either the revisions didn't match or we dont have the map, so we need it
        Call SendData("needmap" & SEP_CHAR & "yes" & END_CHAR)
        Exit Sub
    End If

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::

    If casestring = "mapdata" Then
        n = 1

        Map(Val(parse(1))).name = parse(n + 1)
        Map(Val(parse(1))).Revision = Val(parse(n + 2))
        Map(Val(parse(1))).Moral = Val(parse(n + 3))
        Map(Val(parse(1))).Up = Val(parse(n + 4))
        Map(Val(parse(1))).Down = Val(parse(n + 5))
        Map(Val(parse(1))).Left = Val(parse(n + 6))
        Map(Val(parse(1))).Right = Val(parse(n + 7))
        Map(Val(parse(1))).music = parse(n + 8)
        Map(Val(parse(1))).BootMap = Val(parse(n + 9))
        Map(Val(parse(1))).BootX = Val(parse(n + 10))
        Map(Val(parse(1))).BootY = Val(parse(n + 11))
        Map(Val(parse(1))).Indoors = Val(parse(n + 12))
        Map(Val(parse(1))).Weather = Val(parse(n + 13))

        n = n + 14

        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Map(Val(parse(1))).Tile(X, Y).Ground = Val(parse(n))
                Map(Val(parse(1))).Tile(X, Y).mask = Val(parse(n + 1))
                Map(Val(parse(1))).Tile(X, Y).Anim = Val(parse(n + 2))
                Map(Val(parse(1))).Tile(X, Y).Mask2 = Val(parse(n + 3))
                Map(Val(parse(1))).Tile(X, Y).M2Anim = Val(parse(n + 4))
                Map(Val(parse(1))).Tile(X, Y).Fringe = Val(parse(n + 5))
                Map(Val(parse(1))).Tile(X, Y).FAnim = Val(parse(n + 6))
                Map(Val(parse(1))).Tile(X, Y).Fringe2 = Val(parse(n + 7))
                Map(Val(parse(1))).Tile(X, Y).F2Anim = Val(parse(n + 8))
                Map(Val(parse(1))).Tile(X, Y).Type = Val(parse(n + 9))
                Map(Val(parse(1))).Tile(X, Y).Data1 = Val(parse(n + 10))
                Map(Val(parse(1))).Tile(X, Y).Data2 = Val(parse(n + 11))
                Map(Val(parse(1))).Tile(X, Y).Data3 = Val(parse(n + 12))
                Map(Val(parse(1))).Tile(X, Y).String1 = parse(n + 13)
                Map(Val(parse(1))).Tile(X, Y).String2 = parse(n + 14)
                Map(Val(parse(1))).Tile(X, Y).String3 = parse(n + 15)
                Map(Val(parse(1))).Tile(X, Y).Light = Val(parse(n + 16))
                Map(Val(parse(1))).Tile(X, Y).GroundSet = Val(parse(n + 17))
                Map(Val(parse(1))).Tile(X, Y).MaskSet = Val(parse(n + 18))
                Map(Val(parse(1))).Tile(X, Y).AnimSet = Val(parse(n + 19))
                Map(Val(parse(1))).Tile(X, Y).Mask2Set = Val(parse(n + 20))
                Map(Val(parse(1))).Tile(X, Y).M2AnimSet = Val(parse(n + 21))
                Map(Val(parse(1))).Tile(X, Y).FringeSet = Val(parse(n + 22))
                Map(Val(parse(1))).Tile(X, Y).FAnimSet = Val(parse(n + 23))
                Map(Val(parse(1))).Tile(X, Y).Fringe2Set = Val(parse(n + 24))
                Map(Val(parse(1))).Tile(X, Y).F2AnimSet = Val(parse(n + 25))
                n = n + 26
            Next X
        Next Y

        For X = 1 To 25
            Map(Val(parse(1))).Npc(X) = Val(parse(n))
            Map(Val(parse(1))).SpawnX(X) = Val(parse(n + 1))
            Map(Val(parse(1))).SpawnY(X) = Val(parse(n + 2))
            n = n + 3
        Next X

        ' Save the map
        Call SaveLocalMap(Val(parse(1)))

        ' Check if we get a map from someone else and if we were editing a map cancel it out
        If InEditor Then
            InEditor = False
            frmMapEditor.Visible = False
            frmMapEditor.Visible = False
            frmMirage.Show
' frmMirage.picMapEditor.Visible = False

            If frmMapWarp.Visible Then
                Unload frmMapWarp
            End If

            If frmMapProperties.Visible Then
                Unload frmMapProperties
            End If
        End If

        Exit Sub
    End If

    If casestring = "tilecheck" Then
        n = 5
        X = Val(parse(2))
        Y = Val(parse(3))

        Select Case Val(parse(4))
            Case 0
                Map(Val(parse(1))).Tile(X, Y).Ground = Val(parse(n))
                Map(Val(parse(1))).Tile(X, Y).GroundSet = Val(parse(n + 1))
            Case 1
                Map(Val(parse(1))).Tile(X, Y).mask = Val(parse(n))
                Map(Val(parse(1))).Tile(X, Y).MaskSet = Val(parse(n + 1))
            Case 2
                Map(Val(parse(1))).Tile(X, Y).Anim = Val(parse(n))
                Map(Val(parse(1))).Tile(X, Y).AnimSet = Val(parse(n + 1))
            Case 3
                Map(Val(parse(1))).Tile(X, Y).Mask2 = Val(parse(n))
                Map(Val(parse(1))).Tile(X, Y).Mask2Set = Val(parse(n + 1))
            Case 4
                Map(Val(parse(1))).Tile(X, Y).M2Anim = Val(parse(n))
                Map(Val(parse(1))).Tile(X, Y).M2AnimSet = Val(parse(n + 1))
            Case 5
                Map(Val(parse(1))).Tile(X, Y).Fringe = Val(parse(n))
                Map(Val(parse(1))).Tile(X, Y).FringeSet = Val(parse(n + 1))
            Case 6
                Map(Val(parse(1))).Tile(X, Y).FAnim = Val(parse(n))
                Map(Val(parse(1))).Tile(X, Y).FAnimSet = Val(parse(n + 1))
            Case 7
                Map(Val(parse(1))).Tile(X, Y).Fringe2 = Val(parse(n))
                Map(Val(parse(1))).Tile(X, Y).Fringe2Set = Val(parse(n + 1))
            Case 8
                Map(Val(parse(1))).Tile(X, Y).F2Anim = Val(parse(n))
                Map(Val(parse(1))).Tile(X, Y).F2AnimSet = Val(parse(n + 1))
        End Select
        Call SaveLocalMap(Val(parse(1)))
    End If

    If casestring = "tilecheckattribute" Then
        n = 5
        X = Val(parse(2))
        Y = Val(parse(3))

        Map(Val(parse(1))).Tile(X, Y).Type = Val(parse(n - 1))
        Map(Val(parse(1))).Tile(X, Y).Data1 = Val(parse(n))
        Map(Val(parse(1))).Tile(X, Y).Data2 = Val(parse(n + 1))
        Map(Val(parse(1))).Tile(X, Y).Data3 = Val(parse(n + 2))
        Map(Val(parse(1))).Tile(X, Y).String1 = parse(n + 3)
        Map(Val(parse(1))).Tile(X, Y).String2 = parse(n + 4)
        Map(Val(parse(1))).Tile(X, Y).String3 = parse(n + 5)
        Call SaveLocalMap(Val(parse(1)))
    End If

    ' :::::::::::::::::::::::::::
    ' :: Map items data packet ::
    ' :::::::::::::::::::::::::::
    If casestring = "mapitemdata" Then
        n = 1

        For I = 1 To MAX_MAP_ITEMS
            SaveMapItem(I).Num = Val(parse(n))
            SaveMapItem(I).value = Val(parse(n + 1))
            SaveMapItem(I).Dur = Val(parse(n + 2))
            SaveMapItem(I).X = Val(parse(n + 3))
            SaveMapItem(I).Y = Val(parse(n + 4))

            n = n + 5
        Next I

        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Map npc data packet ::
    ' :::::::::::::::::::::::::
    If casestring = "mapnpcdata" Then
        n = 1

        For I = 1 To 25
            SaveMapNpc(I).Num = Val(parse(n))
            SaveMapNpc(I).X = Val(parse(n + 1))
            SaveMapNpc(I).Y = Val(parse(n + 2))
            SaveMapNpc(I).Dir = Val(parse(n + 3))

            n = n + 4
        Next I

        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::
    ' :: Map send completed packet ::
    ' :::::::::::::::::::::::::::::::
    If casestring = "mapdone" Then
        ' Map = SaveMap

        For I = 1 To MAX_MAP_ITEMS
            MapItem(I) = SaveMapItem(I)
        Next I

        For I = 1 To MAX_MAP_NPCS
            MapNpc(I) = SaveMapNpc(I)
        Next I

        GettingMap = False

        ' Play music
        If Trim$(Map(GetPlayerMap(MyIndex)).music) <> "Ninguna" Then
            Call MapMusic(Map(GetPlayerMap(MyIndex)).music)
        End If

        If GameWeather = WEATHER_RAINING And Map(GetPlayerMap(MyIndex)).Indoors = 0 Then
            Call PlayBGS("rain.wav")
        End If
        
        If GameWeather = WEATHER_THUNDER And Map(GetPlayerMap(MyIndex)).Indoors = 0 Then
            Call PlayBGS("thunder.wav")
        End If

        Exit Sub
    End If

    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
        If (casestring = "saymsg") Or (casestring = "broadcastmsg") Or (casestring = "globalmsg") Or (casestring = "playermsg") Or (casestring = "mapmsg") Or (casestring = "adminmsg") Or (casestring = "d") Or (casestring = "guildmsg") Then
        Call AddText(parse(1), Val(parse(2)))
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Item spawn packet ::
    ' :::::::::::::::::::::::
    If casestring = "spawnitem" Then
        n = Val(parse(1))

        MapItem(n).Num = Val(parse(2))
        MapItem(n).value = Val(parse(3))
        MapItem(n).Dur = Val(parse(4))
        MapItem(n).X = Val(parse(5))
        MapItem(n).Y = Val(parse(6))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Item editor packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "itemeditor") Then
        InItemsEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_ITEMS
            frmIndex.lstIndex.addItem I & ": " & Trim$(Item(I).name)
        Next I

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Update item packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "updateitem") Then
        n = Val(parse(1))

        ' Update the item
        Item(n).name = parse(2)
        Item(n).Pic = Val(parse(3))
        Item(n).Type = Val(parse(4))
        Item(n).Data1 = Val(parse(5))
        Item(n).Data2 = Val(parse(6))
        Item(n).Data3 = Val(parse(7))
        Item(n).StrReq = Val(parse(8))
        Item(n).DefReq = Val(parse(9))
        Item(n).SpeedReq = Val(parse(10))
        Item(n).MagicReq = Val(parse(11))
        Item(n).ClassReq = Val(parse(12))
        Item(n).AccessReq = Val(parse(13))

        Item(n).AddHP = Val(parse(14))
        Item(n).AddMP = Val(parse(15))
        Item(n).AddSP = Val(parse(16))
        Item(n).AddStr = Val(parse(17))
        Item(n).AddDef = Val(parse(18))
        Item(n).AddMagi = Val(parse(19))
        Item(n).AddSpeed = Val(parse(20))
        Item(n).AddEXP = Val(parse(21))
        Item(n).desc = parse(22)
        Item(n).AttackSpeed = Val(parse(23))
        Item(n).Price = Val(parse(24))
        Item(n).Stackable = Val(parse(25))
        Item(n).Bound = Val(parse(26))
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Edit item packet :: <- Used for item editor admins only
    ' ::::::::::::::::::::::
    If (casestring = "edititem") Then
        n = Val(parse(1))

        ' Update the item
        Item(n).name = parse(2)
        Item(n).Pic = Val(parse(3))
        Item(n).Type = Val(parse(4))
        Item(n).Data1 = Val(parse(5))
        Item(n).Data2 = Val(parse(6))
        Item(n).Data3 = Val(parse(7))
        Item(n).StrReq = Val(parse(8))
        Item(n).DefReq = Val(parse(9))
        Item(n).SpeedReq = Val(parse(10))
        Item(n).MagicReq = Val(parse(11))
        Item(n).ClassReq = Val(parse(12))
        Item(n).AccessReq = Val(parse(13))

        Item(n).AddHP = Val(parse(14))
        Item(n).AddMP = Val(parse(15))
        Item(n).AddSP = Val(parse(16))
        Item(n).AddStr = Val(parse(17))
        Item(n).AddDef = Val(parse(18))
        Item(n).AddMagi = Val(parse(19))
        Item(n).AddSpeed = Val(parse(20))
        Item(n).AddEXP = Val(parse(21))
        Item(n).desc = parse(22)
        Item(n).AttackSpeed = Val(parse(23))
        Item(n).Price = Val(parse(24))
        Item(n).Stackable = Val(parse(25))
        Item(n).Bound = Val(parse(26))

        ' Initialize the item editor
        Call ItemEditorInit

        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: mouse packet  ::
    ' :::::::::::::::::::
    If (casestring = "mouse") Then
        Player(MyIndex).input = 1
        Exit Sub
    End If

    ' ::::::::::::::::::
    ' ::Weather Packet::
    ' ::::::::::::::::::
    If (casestring = "mapweather") Then
        If 0 + Val(parse(1)) <> 0 Then
            Map(Val(parse(1))).Weather = Val(parse(2))
            If Val(parse(1)) = 2 Then
                frmMirage.tmrSnowDrop.Interval = Val(parse(3))
            ElseIf Val(parse(1)) = 1 Then
                frmMirage.tmrRainDrop.Interval = Val(parse(3))
            End If
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
    If casestring = "spawnnpc" Then
        n = Val(parse(1))

        MapNpc(n).Num = Val(parse(2))
        MapNpc(n).X = Val(parse(3))
        MapNpc(n).Y = Val(parse(4))
        MapNpc(n).Dir = Val(parse(5))
        MapNpc(n).Big = Val(parse(6))

        ' Client use only
        MapNpc(n).XOffset = 0
        MapNpc(n).YOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    If casestring = "npcdead" Then
        n = Val(parse(1))

        MapNpc(n).Num = 0
        'MapNpc(n).X = 0
        'MapNpc(n).Y = 0
        MapNpc(n).Dir = 0

        ' Client use only
        MapNpc(n).XOffset = 0
        MapNpc(n).YOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Npc editor packet ::
    ' :::::::::::::::::::::::
    If (casestring = "npceditor") Then
        InNpcEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_NPCS
            frmIndex.lstIndex.addItem I & ": " & Trim$(Npc(I).name)
        Next I

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Update npc packet ::
    ' :::::::::::::::::::::::
    If (casestring = "updatenpc") Then
        n = Val(parse(1))

        ' Update the item
        Npc(n).name = parse(2)
        Npc(n).AttackSay = vbNullString
        Npc(n).Sprite = Val(parse(3))
        Npc(n).SpriteSize = Val(parse(4))
        Npc(n).Big = Val(parse(5))
        Npc(n).MaxHP = Val(parse(6))
        Npc(n).Quest = Val(parse(7))
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit npc packet ::
    ' :::::::::::::::::::::
    If (casestring = "editnpc") Then
        n = Val(parse(1))

        ' Update the npc
        Npc(n).name = parse(2)
        Npc(n).AttackSay = parse(3)
        Npc(n).Sprite = Val(parse(4))
        Npc(n).SpawnSecs = Val(parse(5))
        Npc(n).Behavior = Val(parse(6))
        Npc(n).Range = Val(parse(7))
        Npc(n).STR = Val(parse(8))
        Npc(n).DEF = Val(parse(9))
        Npc(n).SPEED = Val(parse(10))
        Npc(n).MAGI = Val(parse(11))
        Npc(n).Big = Val(parse(12))
        Npc(n).MaxHP = Val(parse(13))
        Npc(n).Exp = Val(parse(14))
        Npc(n).SpawnTime = Val(parse(15))
        Npc(n).Element = Val(parse(16))
        Npc(n).SpriteSize = Val(parse(17))
        Npc(n).Quest = Val(parse(18))

        ' Call GlobalMsg("At editnpc..." & Npc(n).Element)
        z = 19
        For I = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(I).chance = Val(parse(z))
            Npc(n).ItemNPC(I).ItemNum = Val(parse(z + 1))
            Npc(n).ItemNPC(I).ItemValue = Val(parse(z + 2))
            z = z + 3
        Next I
        
        Npc(n).standstill = parse(50)

        ' Initialize the npc editor
        Call NpcEditorInit

        Exit Sub
    End If

    ' ::::::::::::::::::::
    ' :: Map key packet ::
    ' ::::::::::::::::::::
    If (casestring = "mapkey") Then
        X = Val(parse(1))
        Y = Val(parse(2))
        n = Val(parse(3))

        TempTile(X, Y).DoorOpen = n

        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit map packet ::
    ' :::::::::::::::::::::
    If (casestring = "editmap") Then
        Call EditorInit
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Shop editor packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "shopeditor") Then
        InShopEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_SHOPS
            frmIndex.lstIndex.addItem I & ": " & Trim$(Shop(I).name)
        Next I

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Update shop packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "updateshop") Then
        n = Val(parse(1))

        ' Update the shop name
        Shop(n).name = parse(2)
        Shop(n).FixesItems = Val(parse(3))
        Shop(n).BuysItems = Val(parse(4))
        Shop(n).ShowInfo = Val(parse(5))
        Shop(n).currencyItem = Val(parse(6))

        m = 7
        ' Get shop items
        For I = 1 To MAX_SHOP_ITEMS
            Shop(n).ShopItem(I).ItemNum = Val(parse(m))
            Shop(n).ShopItem(I).Amount = Val(parse(m + 1))
            Shop(n).ShopItem(I).Price = Val(parse(m + 2))
            m = m + 3
        Next I

        Exit Sub
    End If


    If casestring = "loadsctt" Then
    n = Val(parse(1))
    Spell(n).CastTimer = parse(2)
    Spell(n).TimeToCast = parse(3)
    Spell(n).MPCost = parse(4)
    End If
    
    ' ::::::::::::::::::::::
    ' :: Edit shop packet :: <- Used for shop editor admins only
    ' ::::::::::::::::::::::
    If (casestring = "editshop") Then

        shopNum = Val(parse(1))

        ' Update the shop
        Shop(shopNum).name = parse(2)
        Shop(shopNum).FixesItems = Val(parse(3))
        Shop(shopNum).BuysItems = Val(parse(4))
        Shop(shopNum).ShowInfo = Val(parse(5))
        Shop(shopNum).currencyItem = Val(parse(6))

        m = 7
        For I = 1 To 25
            Shop(shopNum).ShopItem(I).ItemNum = Val(parse(m))
            Shop(shopNum).ShopItem(I).Amount = Val(parse(m + 1))
            Shop(shopNum).ShopItem(I).Price = Val(parse(m + 2))
            m = m + 3
        Next I

        ' Initialize the shop editor
        Call ShopEditorInit

        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Spell editor packet ::
    ' :::::::::::::::::::::::::
    If (casestring = "spelleditor") Then
        InSpellEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        ' Add the names
        For I = 1 To MAX_SPELLS
            frmIndex.lstIndex.addItem I & ": " & Trim$(Spell(I).name)
        Next I

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Update spell packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "updatespell") Then
        n = Val(parse(1))

        ' Update the spell name
        Spell(n).name = parse(2)
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Edit spell packet :: <- Used for spell editor admins only
    ' :::::::::::::::::::::::
    If (casestring = "editspell") Then
        n = Val(parse(1))

        ' Update the spell
        Spell(n).name = parse(2)
        Spell(n).ClassReq = Val(parse(3))
        Spell(n).LevelReq = Val(parse(4))
        Spell(n).Type = Val(parse(5))
        Spell(n).Data1 = Val(parse(6))
        Spell(n).Data2 = Val(parse(7))
        Spell(n).Data3 = Val(parse(8))
        Spell(n).MPCost = Val(parse(9))
        Spell(n).Sound = Val(parse(10))
        Spell(n).Range = Val(parse(11))
        Spell(n).SpellAnim = Val(parse(12))
        Spell(n).SpellTime = Val(parse(13))
        Spell(n).SpellDone = Val(parse(14))
        Spell(n).AE = Val(parse(15))
        Spell(n).Big = Val(parse(16))
        Spell(n).Element = Val(parse(17))
        Spell(n).TimeToCast = Val(parse(18))
        Spell(n).CastTimer = Val(parse(19))


        ' Initialize the spell editor
        Call SpellEditorInit

        Exit Sub
    End If

    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    If (casestring = "goshop") Then
        shopNum = Val(parse(1))
        ' Show the shop
        Call GoShop(shopNum)
    End If

    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If (casestring = "spells") Then

        frmMirage.picPlayerSpells.Visible = True
        frmMirage.lstSpells.Clear

        ' Put spells known in player record
        For I = 1 To MAX_PLAYER_SPELLS
            Player(MyIndex).Spell(I) = Val(parse(I))
            If Player(MyIndex).Spell(I) <> 0 Then
                frmMirage.lstSpells.addItem I & ": " & Trim$(Spell(Player(MyIndex).Spell(I)).name)
            Else
                frmMirage.lstSpells.addItem "-- Libre --"
            End If
        Next I

        frmMirage.lstSpells.ListIndex = 0
    End If

    ' ::::::::::::::::::::
    ' :: Weather packet ::
    ' ::::::::::::::::::::
    If (casestring = "weather") Then
        If Val(parse(1)) = WEATHER_RAINING And GameWeather <> WEATHER_RAINING And Map(GetPlayerMap(MyIndex)).Indoors = 0 Then
            Call AddText("Ves como la lluvia cae sobre ti!", BRIGHTGREEN)
            Call PlayBGS("rain.mp3")
        End If
        If Val(parse(1)) = WEATHER_THUNDER And GameWeather <> WEATHER_THUNDER And Map(GetPlayerMap(MyIndex)).Indoors = 0 Then
            Call AddText("Ves como el cielo se ilumina con los relampagos!", BRIGHTGREEN)
            Call PlayBGS("thunder.mp3")
        End If
        If Val(parse(1)) = WEATHER_SNOWING And GameWeather <> WEATHER_SNOWING Then
            Call AddText("Ves como la nieve cae sobre ti!", BRIGHTGREEN)
        End If

        If Val(parse(1)) = WEATHER_NONE Then
            If GameWeather = WEATHER_RAINING Then
                Call AddText("La lluvia se fue.", BRIGHTGREEN)
                Call StopSound
            ElseIf GameWeather = WEATHER_SNOWING Then
                Call AddText("La nieve dejo de caer.", BRIGHTGREEN)
            ElseIf GameWeather = WEATHER_THUNDER Then
                Call AddText("Los truenos se disiparon.", BRIGHTGREEN)
                Call StopSound
            End If
        End If
        GameWeather = Val(parse(1))
        RainIntensity = Val(parse(2))
        If MAX_RAINDROPS <> RainIntensity Then
            MAX_RAINDROPS = RainIntensity
            ReDim DropRain(1 To MAX_RAINDROPS) As DropRainRec
            ReDim DropSnow(1 To MAX_RAINDROPS) As DropRainRec
        End If
    End If
    
    ' ::::::::::::::::::::::::::::::::
    ' :: editor de scripts online   ::
    ' ::::::::::::::::::::::::::::::::
    
        If casestring = "maineditor" Then
   
        'If LCase(Dir(App.Path & "Scripts", vbDirectory)) <> "scripts" Then
            'Call MkDir(App.Path & "Scripts")
        'End If
       
        Dim AFileName
        AFileName = "Scripts\Main.txt"
             
        Dim F
        F = FreeFile
        Open App.Path & "" & AFileName For Output As #F
            Print #F, parse(1)
        Close #F
       
        Unload frmEditor
        frmEditor.Show
    End If

    ' ::::::::::::::::::::::::::::::::
    ' :: playername coloring packet ::
    ' ::::::::::::::::::::::::::::::::
    If (casestring = "namecolor") Then
        Player(Val(parse(1))).color(1) = CByte(parse(2))
        Player(Val(parse(1))).color(2) = CByte(parse(3))
        Player(Val(parse(1))).color(3) = CByte(parse(4))
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: image packet      ::
    ' :::::::::::::::::::::::
    If (LCase$(parse(0)) = "fog") Then
        rec.top = Int(Val(parse(4)))
        rec.Bottom = Int(Val(parse(5)))
        rec.Left = Int(Val(parse(6)))
        rec.Right = Int(Val(parse(7)))
        Call DD_BackBuffer.BltFast(Val(parse(1)), Val(parse(2)), DD_TileSurf(Val(parse(3))), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Get Online List ::
    ' ::::::::::::::::::::::::::
    If casestring = "onlinelist" Then
        frmMirage.lstOnline.Clear

        n = 2
        z = Val(parse(1))
        For X = n To (z + 1)
            frmMirage.lstOnline.addItem Trim$(parse(n))
            n = n + 2
        Next X
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Blit Player Damage ::
    ' ::::::::::::::::::::::::
    If casestring = "blitplayerdmg" Then
       
        For X = 9 To 1 Step -1
            iii(X) = iii(X - 1)
            DmgDamage(X) = DmgDamage(X - 1)
            NPCWho(X) = NPCWho(X - 1)
            DmgTime(X) = DmgTime(X - 1)
        Next
       
        DmgDamage(0) = Val(parse(1))
        NPCWho(0) = Val(parse(2))
        DmgTime(0) = GetTickCount
        iii(0) = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Blit NPC Damage ::
    ' :::::::::::::::::::::
    If casestring = "blitnpcdmg" Then
       
        For X = 9 To 1 Step -1
            II(X) = II(X - 1)
            NPCDmgDamage(X) = NPCDmgDamage(X - 1)
            NPCDmgTime(X) = NPCDmgTime(X - 1)
        Next
       
        NPCDmgDamage(0) = Val(parse(1))
        NPCDmgTime(0) = GetTickCount
        II(0) = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::::::
    ' :: Retrieve the player's inventory ::
    ' :::::::::::::::::::::::::::::::::::::
    If casestring = "pptrading" Then
        frmPlayerTrade.Items1.Clear
        frmPlayerTrade.Items2.Clear
        For I = 1 To MAX_PLAYER_TRADES
            Trading(I).InvNum = 0
            Trading(I).InvName = vbNullString
            Trading2(I).InvNum = 0
            Trading2(I).InvName = vbNullString
            frmPlayerTrade.Items1.addItem I & ": <Nada>"
            frmPlayerTrade.Items2.addItem I & ": <Nada>"
        Next I

        frmPlayerTrade.Items1.ListIndex = 0

        Call UpdateTradeInventory
        frmPlayerTrade.Show vbModeless, frmMirage
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If casestring = "qtrade" Then
        For I = 1 To MAX_PLAYER_TRADES
            Trading(I).InvNum = 0
            Trading(I).InvName = vbNullString
            Trading2(I).InvNum = 0
            Trading2(I).InvName = vbNullString
        Next I

        frmPlayerTrade.Command1.ForeColor = &H0&
        frmPlayerTrade.Command2.ForeColor = &H0&

        frmPlayerTrade.Visible = False
        Exit Sub
    End If

    ' ::::::::::::::::::
    ' :: Disable Time ::
    ' ::::::::::::::::::
    If casestring = "dtime" Then
        If parse(1) = "True" Then
            frmMirage.lblGameClock.Caption = vbNullString
            frmMirage.lblGameClock.Visible = False
            frmMirage.tmrGameClock.Enabled = False
        Else
            frmMirage.lblGameClock.Caption = vbNullString
            frmMirage.lblGameClock.Visible = True
            frmMirage.tmrGameClock.Enabled = True
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If casestring = "updatetradeitem" Then
        n = Val(parse(1))

        Trading2(n).InvNum = Val(parse(2))
        Trading2(n).InvName = parse(3)

        If STR(Trading2(n).InvNum) <= 0 Then
            frmPlayerTrade.Items2.List(n - 1) = n & ": <Nada>"
        Else
            frmPlayerTrade.Items2.List(n - 1) = n & ": " & Trim$(Trading2(n).InvName)
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If casestring = "trading" Then
        n = Val(parse(1))
        If n = 0 Then
            frmPlayerTrade.Command2.ForeColor = &H0&
        End If
        If n = 1 Then
            frmPlayerTrade.Command2.ForeColor = &HFF00&
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Chat System Packets ::
    ' :::::::::::::::::::::::::
    If casestring = "ppchatting" Then
        frmPlayerChat.txtChat.Text = vbNullString
        frmPlayerChat.txtSay.Text = vbNullString
        frmPlayerChat.Label1.Caption = "" & Trim$(Player(Val(parse(1))).name)

        frmPlayerChat.Show vbModeless, frmMirage
        Exit Sub
    End If

    If casestring = "qchat" Then
        frmPlayerChat.txtChat.Text = vbNullString
        frmPlayerChat.txtSay.Text = vbNullString
        frmPlayerChat.Visible = False
        frmPlayerTrade.Command2.ForeColor = &H8000000F
        frmPlayerTrade.Command1.ForeColor = &H8000000F
        Exit Sub
    End If

    If casestring = "sendchat" Then

        S = vbNewLine & GetPlayerName(Val(parse(2))) & "> " & Trim$(parse(1))
        frmPlayerChat.txtChat.SelStart = Len(frmPlayerChat.txtChat.Text)
        frmPlayerChat.txtChat.SelColor = QBColor(BROWN)
        frmPlayerChat.txtChat.SelText = S
        frmPlayerChat.txtChat.SelStart = Len(frmPlayerChat.txtChat.Text) - 1
        Exit Sub
    End If
' :::::::::::::::::::::::::::::
' :: END Chat System Packets ::
' :::::::::::::::::::::::::::::

    ' :::::::::::::::::::::::
    ' :: Play Sound Packet ::
    ' :::::::::::::::::::::::
    If casestring = "sound" Then
        S = LCase$(parse(1))
        Select Case Trim$(S)
            Case "attack"
                Call PlaySound("sword.wav")
            Case "critical"
                Call PlaySound("critical.wav")
            Case "miss"
                Call PlaySound("miss.wav")
            Case "key"
                Call PlaySound("key.wav")
            Case "cofre"
                Call PlaySound("cofre.wav")
            Case "magic"
                Call PlaySound("magic" & Val(parse(2)) & ".wav")
            Case "warp"
                If FileExists("SFX\warp.wav") Then
                    Call PlaySound("warp.wav")
                End If
            Case "pain"
                Call PlaySound("pain.wav")
            Case "soundattribute"
                Call PlaySound(Trim$(parse(2)))
        End Select
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::::::::
    ' :: Sprite Change Confirmation Packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If casestring = "spritechange" Then
        If Val(parse(1)) = 1 Then
            I = MsgBox("¿Estas seguro de comprar este sprite?", 4, "Comprando Sprite")
            If I = 6 Then
                Call SendData("buysprite" & END_CHAR)
            End If
        Else
            Call SendData("buysprite" & END_CHAR)
        End If
        Exit Sub
    End If
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: House Buy Confirmation Packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If casestring = "housebuy" Then
        If Val(parse(1)) = 1 Then
            I = MsgBox("¿Estas seguro de comprar esta casa?", 4, "Comprando Casa")
            If I = 6 Then
                Call SendData("buyhouse" & END_CHAR)
            End If
        Else
            Call SendData("buyhouse" & END_CHAR)
        End If
        Exit Sub
    End If
    ' ::::::::::::::::::::::::::::::::::::
    ' :: Change Player Direction Packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If casestring = "changedir" Then
        Player(Val(parse(2))).Dir = Val(parse(1))
        Exit Sub
    End If
    
    ' Paquete de Anuncio IN-GAME por Stream
    ' /////////////////////////////////////
    ' BETA
    
    If casestring = "anuncio" Then
    frmMirage.anunciobox.Visible = True
    frmMirage.anuncio.Visible = True
    frmMirage.anuncio.Caption = parse(1)
    
    frmMirage.anunciotimer.Enabled = True
    frmMirage.anunciotimer.Interval = 1000
    frmMirage.anunciocuenta.Caption = "10"
    End If

    ' ::::::::::::::::::::::::::::::
    ' :: Flash Movie Event Packet ::
    ' ::::::::::::::::::::::::::::::
    If casestring = "flashevent" Then
        If LCase$(Mid$(Trim$(parse(1)), 1, 7)) = "http://" Then
            WriteINI "CONFIG", "Music", 0, App.Path & "\config.ini"
            WriteINI "CONFIG", "Sound", 0, App.Path & "\config.ini"
            Call StopBGM
            Call StopSound
            frmFlash.Flash.LoadMovie 0, Trim$(parse(1))
            frmFlash.Flash.Play
            frmFlash.Check.Enabled = True
            frmFlash.Show vbModeless, frmMirage
        ElseIf FileExists("Flashs\" & Trim$(parse(1))) = True Then
            WriteINI "CONFIG", "Music", 0, App.Path & "\config.ini"
            WriteINI "CONFIG", "Sound", 0, App.Path & "\config.ini"
            Call StopBGM
            Call StopSound
            frmFlash.Flash.LoadMovie 0, App.Path & "\Flashs\" & Trim$(parse(1))
            frmFlash.Flash.Play
            frmFlash.Check.Enabled = True
            frmFlash.Show vbModeless, frmMirage
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: Prompt Packet ::
    ' :::::::::::::::::::
    If casestring = "prompt" Then
        I = MsgBox(Trim$(parse(1)), vbYesNo)
        Call SendData("prompt" & SEP_CHAR & I & SEP_CHAR & Val(parse(2)) & END_CHAR)
        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: Prompt Packet ::
    ' :::::::::::::::::::
    If casestring = "querybox" Then
        frmQuery.Label1.Caption = Trim$(parse(1))
        frmQuery.Label2.Caption = parse(2)
        frmQuery.Show vbModal
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Emoticon editor packet ::
    ' ::::::::::::::::::::::::::::
    If (casestring = "emoticoneditor") Then
        InEmoticonEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        For I = 0 To MAX_EMOTICONS
            frmIndex.lstIndex.addItem I & ": " & Trim$(Emoticons(I).Command)
        Next I

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Element editor packet ::
    ' ::::::::::::::::::::::::::::
    If (casestring = "elementeditor") Then
        InElementEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        For I = 0 To MAX_ELEMENTS
            frmIndex.lstIndex.addItem I & ": " & Trim$(Element(I).name)
        Next I

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    If (casestring = "editelement") Then
        n = Val(parse(1))

        Element(n).name = parse(2)
        Element(n).Strong = Val(parse(3))
        Element(n).Weak = Val(parse(4))

        Call ElementEditorInit
        Exit Sub
    End If

    If (casestring = "updateelement") Then
        n = Val(parse(1))

        Element(n).name = parse(2)
        Element(n).Strong = Val(parse(3))
        Element(n).Weak = Val(parse(4))
        Exit Sub
    End If

    If (casestring = "editemoticon") Then
        n = Val(parse(1))

        Emoticons(n).Command = parse(2)
        Emoticons(n).Pic = Val(parse(3))

        Call EmoticonEditorInit
        Exit Sub
    End If

    If (casestring = "updateemoticon") Then
        n = Val(parse(1))

        Emoticons(n).Command = parse(2)
        Emoticons(n).Pic = Val(parse(3))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Arrow editor packet ::
    ' ::::::::::::::::::::::::::::
    If (casestring = "arroweditor") Then
        InArrowEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        For I = 1 To MAX_ARROWS
            frmIndex.lstIndex.addItem I & ": " & Trim$(Arrows(I).name)
        Next I

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    If (casestring = "updatearrow") Then
        n = Val(parse(1))

        Arrows(n).name = parse(2)
        Arrows(n).Pic = Val(parse(3))
        Arrows(n).Range = Val(parse(4))
        Arrows(n).Amount = Val(parse(5))
        Exit Sub
    End If

    If (casestring = "editarrow") Then
        n = Val(parse(1))

        Arrows(n).name = parse(2)

        Call ArrowEditorInit
        Exit Sub
    End If

    If (casestring = "updatearrow") Then
        n = Val(parse(1))

        Arrows(n).name = parse(2)
        Arrows(n).Pic = Val(parse(3))
        Arrows(n).Range = Val(parse(4))
        Arrows(n).Amount = Val(parse(5))
        Exit Sub
    End If

    If (casestring = "hookshot") Then
        n = Val(parse(1))
        I = Val(parse(3))

        Player(n).HookShotAnim = Arrows(Val(parse(2))).Pic
        Player(n).HookShotTime = GetTickCount
        Player(n).HookShotToX = Val(parse(4))
        Player(n).HookShotToY = Val(parse(5))
        Player(n).HookShotX = GetPlayerX(n)
        Player(n).HookShotY = GetPlayerY(n)
        Player(n).HookShotSucces = Val(parse(6))
        Player(n).HookShotDir = Val(parse(3))

        Call PlaySound("grapple.wav")
        Call PlaySound("grapple-fire.wav")

        If I = DIR_DOWN Then
            Player(n).HookShotX = GetPlayerX(n)
            Player(n).HookShotY = GetPlayerY(n) + 1
            If Player(n).HookShotX - 1 > MAX_MAPY Then
                Player(n).HookShotX = 0
                Player(n).HookShotY = 0
                Exit Sub
            End If
        End If
        If I = DIR_UP Then
            Player(n).HookShotX = GetPlayerX(n)
            Player(n).HookShotY = GetPlayerY(n) - 1
            If Player(n).HookShotY + 1 < 0 Then
                Player(n).HookShotX = 0
                Player(n).HookShotY = 0
                Exit Sub
            End If
        End If
        If I = DIR_RIGHT Then
            Player(n).HookShotX = GetPlayerX(n) + 1
            Player(n).HookShotY = GetPlayerY(n)
            If Player(n).HookShotX - 1 > MAX_MAPX Then
                Player(n).HookShotX = 0
                Player(n).HookShotY = 0
                Exit Sub
            End If
        End If
        If I = DIR_LEFT Then
            Player(n).HookShotX = GetPlayerX(n) - 1
            Player(n).HookShotY = GetPlayerY(n)
            If Player(n).HookShotX + 1 < 0 Then
                Player(n).Arrow(X).Arrow = 0
                Exit Sub
            End If
        End If
        Exit Sub
    End If

    If (casestring = "checkarrows") Then
        n = Val(parse(1))
        z = Val(parse(2))
        I = Val(parse(3))

        For X = 1 To MAX_PLAYER_ARROWS
            If Player(n).Arrow(X).Arrow = 0 Then
                Player(n).Arrow(X).Arrow = 1
                Player(n).Arrow(X).ArrowNum = z
                Player(n).Arrow(X).ArrowAnim = Arrows(z).Pic
                Player(n).Arrow(X).ArrowTime = GetTickCount
                Player(n).Arrow(X).ArrowVarX = 0
                Player(n).Arrow(X).ArrowVarY = 0
                Player(n).Arrow(X).ArrowY = GetPlayerY(n)
                Player(n).Arrow(X).ArrowX = GetPlayerX(n)
                Player(n).Arrow(X).ArrowAmount = p

                If I = DIR_DOWN Then
                    Player(n).Arrow(X).ArrowY = GetPlayerY(n) + 1
                    Player(n).Arrow(X).ArrowPosition = 0
                    If Player(n).Arrow(X).ArrowY - 1 > MAX_MAPY Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                If I = DIR_UP Then
                    Player(n).Arrow(X).ArrowY = GetPlayerY(n) - 1
                    Player(n).Arrow(X).ArrowPosition = 1
                    If Player(n).Arrow(X).ArrowY + 1 < 0 Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                If I = DIR_RIGHT Then
                    Player(n).Arrow(X).ArrowX = GetPlayerX(n) + 1
                    Player(n).Arrow(X).ArrowPosition = 2
                    If Player(n).Arrow(X).ArrowX - 1 > MAX_MAPX Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                If I = DIR_LEFT Then
                    Player(n).Arrow(X).ArrowX = GetPlayerX(n) - 1
                    Player(n).Arrow(X).ArrowPosition = 3
                    If Player(n).Arrow(X).ArrowX + 1 < 0 Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                Exit For
            End If
        Next X
        Exit Sub
    End If

    If (casestring = "checksprite") Then
        n = Val(parse(1))

        Player(n).Sprite = Val(parse(2))
        Exit Sub
    End If

    If (casestring = "mapreport") Then
        n = 1

        frmMapReport.lstIndex.Clear
        For I = 1 To MAX_MAPS
            frmMapReport.lstIndex.addItem I & ": " & Trim$(parse(n))
            n = n + 1
        Next I

        frmMapReport.Show vbModeless, frmMirage
        Exit Sub
    End If

    ' :::::::::::::::::
    ' :: Time packet ::
    ' :::::::::::::::::
    If (casestring = "time") Then
        GameTime = Val(parse(1))
        If GameTime = TIME_DAY Then
            Call AddText("El dia llego a este lugar.", WHITE)
        Else
            Call AddText("La noche cae sobre este lugar.", WHITE)
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Spell anim packet ::
    ' :::::::::::::::::::::::
    If (casestring = "spellanim") Then
        Dim Spellnum As Long
        Spellnum = Val(parse(1))

        Spell(Spellnum).SpellAnim = Val(parse(2))
        Spell(Spellnum).SpellTime = Val(parse(3))
        Spell(Spellnum).SpellDone = Val(parse(4))
        Spell(Spellnum).Big = Val(parse(9))

        Player(Val(parse(5))).Spellnum = Spellnum

        For I = 1 To MAX_SPELL_ANIM
            If Player(Val(parse(5))).SpellAnim(I).CastedSpell = NO Then
                Player(Val(parse(5))).SpellAnim(I).SpellDone = 0
                Player(Val(parse(5))).SpellAnim(I).SpellVar = 0
                Player(Val(parse(5))).SpellAnim(I).SpellTime = GetTickCount
                Player(Val(parse(5))).SpellAnim(I).TargetType = Val(parse(6))
                Player(Val(parse(5))).SpellAnim(I).Target = Val(parse(7))
                Player(Val(parse(5))).SpellAnim(I).CastedSpell = YES
                Exit For
            End If
        Next I
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Script Spell anim packet ::
    ' :::::::::::::::::::::::
    If (casestring = "scriptspellanim") Then
        Spell(Val(parse(1))).SpellAnim = Val(parse(2))
        Spell(Val(parse(1))).SpellTime = Val(parse(3))
        Spell(Val(parse(1))).SpellDone = Val(parse(4))
        Spell(Val(parse(1))).Big = Val(parse(7))


        For I = 1 To MAX_SCRIPTSPELLS
            If ScriptSpell(I).CastedSpell = NO Then
                ScriptSpell(I).Spellnum = Val(parse(1))
                ScriptSpell(I).SpellDone = 0
                ScriptSpell(I).SpellVar = 0
                ScriptSpell(I).SpellTime = GetTickCount
                ScriptSpell(I).X = Val(parse(5))
                ScriptSpell(I).Y = Val(parse(6))
                ScriptSpell(I).CastedSpell = YES
                Exit For
            End If
        Next I
        Exit Sub
    End If

    If (casestring = "checkemoticons") Then
        n = Val(parse(1))

        Player(n).EmoticonNum = Val(parse(2))
        Player(n).EmoticonTime = GetTickCount
        Player(n).EmoticonVar = 0
        Exit Sub
    End If


    If casestring = "levelup" Then
        Player(Val(parse(1))).LevelUpT = GetTickCount
        Player(Val(parse(1))).LevelUp = 1
        Exit Sub
    End If

    If casestring = "damagedisplay" Then
        For I = 1 To MAX_BLT_LINE
            If Val(parse(1)) = 0 Then
                If BattlePMsg(I).index <= 0 Then
                    BattlePMsg(I).index = 1
                    BattlePMsg(I).Msg = parse(2)
                    BattlePMsg(I).color = Val(parse(3))
                    BattlePMsg(I).time = GetTickCount
                    BattlePMsg(I).Done = 1
                    BattlePMsg(I).Y = 0
                    Exit Sub
                Else
                    BattlePMsg(I).Y = BattlePMsg(I).Y - 15
                End If
            Else
                If BattleMMsg(I).index <= 0 Then
                    BattleMMsg(I).index = 1
                    BattleMMsg(I).Msg = parse(2)
                    BattleMMsg(I).color = Val(parse(3))
                    BattleMMsg(I).time = GetTickCount
                    BattleMMsg(I).Done = 1
                    BattleMMsg(I).Y = 0
                    Exit Sub
                Else
                    BattleMMsg(I).Y = BattleMMsg(I).Y - 15
                End If
            End If
        Next I

        z = 1
        If Val(parse(1)) = 0 Then
            For I = 1 To MAX_BLT_LINE
                If I < MAX_BLT_LINE Then
                    If BattlePMsg(I).Y < BattlePMsg(I + 1).Y Then
                        z = I
                    End If
                Else
                    If BattlePMsg(I).Y < BattlePMsg(1).Y Then
                        z = I
                    End If
                End If
            Next I

            BattlePMsg(z).index = 1
            BattlePMsg(z).Msg = parse(2)
            BattlePMsg(z).color = Val(parse(3))
            BattlePMsg(z).time = GetTickCount
            BattlePMsg(z).Done = 1
            BattlePMsg(z).Y = 0
        Else
            For I = 1 To MAX_BLT_LINE
                If I < MAX_BLT_LINE Then
                    If BattleMMsg(I).Y < BattleMMsg(I + 1).Y Then
                        z = I
                    End If
                Else
                    If BattleMMsg(I).Y < BattleMMsg(1).Y Then
                        z = I
                    End If
                End If
            Next I

            BattleMMsg(z).index = 1
            BattleMMsg(z).Msg = parse(2)
            BattleMMsg(z).color = Val(parse(3))
            BattleMMsg(z).time = GetTickCount
            BattleMMsg(z).Done = 1
            BattleMMsg(z).Y = 0
        End If
        Exit Sub
    End If

    If casestring = "itembreak" Then
        ItemDur(Val(parse(1))).Item = Val(parse(2))
        ItemDur(Val(parse(1))).Dur = Val(parse(3))
        ItemDur(Val(parse(1))).Done = 1
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::::::::::::
    ' :: Index player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::::::::
    If casestring = "itemworn" Then
        Player(Val(parse(1))).Armor = Val(parse(2))
        Player(Val(parse(1))).Weapon = Val(parse(3))
        Player(Val(parse(1))).Helmet = Val(parse(4))
        Player(Val(parse(1))).Shield = Val(parse(5))
        Player(Val(parse(1))).legs = Val(parse(6))
        Player(Val(parse(1))).Ring = Val(parse(7))
        Player(Val(parse(1))).Necklace = Val(parse(8))
        Exit Sub
    End If

    If casestring = "scripttile" Then
        frmScript.lblScript.Caption = parse(1)
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Set player speed ::
    ' ::::::::::::::::::::::
    If casestring = "setspeed" Then
        SetSpeed parse(1), Val(parse(2))
        Exit Sub
    End If

    ' ::::::::::::::::::
    ' :: Custom Menu  ::
    ' ::::::::::::::::::
    If (casestring = "showcustommenu") Then
        ' Error handling
        If Not FileExists(parse(2)) Then
            Call MsgBox(parse(2) & " no encontrado. Carga el menu abortada. Por favor, contacta con un GM.", vbExclamation)
            Exit Sub
        End If

        CUSTOM_TITLE = parse(1)
        CUSTOM_IS_CLOSABLE = Val(parse(3))

        frmCustom1.picBackground.top = 0
        frmCustom1.picBackground.Left = 0
        frmCustom1.picBackground = LoadPicture(App.Path & parse(2))
        frmCustom1.Height = PixelsToTwips(24 + frmCustom1.picBackground.Height, 1)
        frmCustom1.Width = PixelsToTwips(6 + frmCustom1.picBackground.Width, 0)
        frmCustom1.Visible = True

        Exit Sub
    End If

    If (casestring = "closecustommenu") Then

        CUSTOM_TITLE = "CLOSED"
        Unload frmCustom1

        Exit Sub
    End If

    If (casestring = "loadpiccustommenu") Then

        CustomIndex = parse(1)
        strfilename = parse(2)
        CustomX = Val(parse(3))
        CustomY = Val(parse(4))
                
        If Not IsInArray(frmCustom1.picCustom, CInt(CustomIndex)) Then
            Load frmCustom1.picCustom(CustomIndex)
        End If

        If strfilename = vbNullString Then
            strfilename = "MEGAUBERBLANKNESSOFUNHOLYPOWER" 'smooth :\    -Pickle
        End If

        If FileExists(strfilename) = True Then
            frmCustom1.picCustom(CustomIndex) = LoadPicture(App.Path & strfilename)
            frmCustom1.picCustom(CustomIndex).top = CustomY
            frmCustom1.picCustom(CustomIndex).Left = CustomX
            frmCustom1.picCustom(CustomIndex).Visible = True
        Else
            frmCustom1.picCustom(CustomIndex).Picture = LoadPicture()
            frmCustom1.picCustom(CustomIndex).Visible = False
        End If

        Exit Sub
    End If

    If (casestring = "loadlabelcustommenu") Then

        CustomIndex = parse(1)
        strfilename = parse(2)
        CustomX = Val(parse(3))
        CustomY = Val(parse(4))
        customsize = Val(parse(5))
        customcolour = Val(parse(6))
        
        If CustomIndex > frmCustom1.BtnCustom.UBound Then
            Load frmCustom1.BtnCustom(CustomIndex)
        End If

        frmCustom1.BtnCustom(CustomIndex).Caption = strfilename
        frmCustom1.BtnCustom(CustomIndex).top = CustomY
        frmCustom1.BtnCustom(CustomIndex).Left = CustomX
        frmCustom1.BtnCustom(CustomIndex).Font.Bold = True
        frmCustom1.BtnCustom(CustomIndex).Font.Size = customsize
        frmCustom1.BtnCustom(CustomIndex).ForeColor = QBColor(customcolour)
        frmCustom1.BtnCustom(CustomIndex).Visible = True
        frmCustom1.BtnCustom(CustomIndex).Alignment = parse(7)

        If parse(8) <= 0 Or parse(9) <= 0 Then
            frmCustom1.BtnCustom(CustomIndex).AutoSize = True
        Else
            frmCustom1.BtnCustom(CustomIndex).AutoSize = False
            frmCustom1.BtnCustom(CustomIndex).Width = parse(8)
            frmCustom1.BtnCustom(CustomIndex).Height = parse(9)
        End If

        Exit Sub
    End If

    If (casestring = "loadtextboxcustommenu") Then

        CustomIndex = parse(1)
        strfilename = parse(2)
        CustomX = Val(parse(3))
        CustomY = Val(parse(4))
        customtext = parse(5)
        
        If CustomIndex > frmCustom1.txtCustom.UBound Then
            Load frmCustom1.txtCustom(CustomIndex)
            Load frmCustom1.txtcustomOK(CustomIndex)
        End If

        frmCustom1.txtCustom(CustomIndex).Text = customtext
        frmCustom1.txtCustom(CustomIndex).top = CustomY
        frmCustom1.txtCustom(CustomIndex).Left = strfilename
        frmCustom1.txtCustom(CustomIndex).Width = CustomX - 32
        frmCustom1.txtcustomOK(CustomIndex).top = CustomY
        frmCustom1.txtcustomOK(CustomIndex).Left = frmCustom1.txtCustom(CustomIndex).Left + frmCustom1.txtCustom(CustomIndex).Width
        frmCustom1.txtcustomOK(CustomIndex).Visible = True
        frmCustom1.txtCustom(CustomIndex).Visible = True

        Exit Sub
    End If

    If (casestring = "loadinternetwindow") Then
        customtext = parse(1)
        ' DEBUG STRING
        ' Call AddText(customtext, 15)
        ShellExecute 1, "open", Trim(customtext), vbNullString, vbNullString, 1
        Exit Sub
    End If

    If (casestring = "returncustomboxmsg") Then
        customsize = parse(1)

        packet = "returningcustomboxmsg" & SEP_CHAR & frmCustom1.txtCustom(customsize).Text & END_CHAR
        Call SendData(packet)

        Exit Sub
    End If

    If (LCase(parse(0)) = "killinfo") Then
    NpcKillAmount = Val(parse(1))
    NpcKillName = parse(2)
    NpcKillFinal2 = parse(3)
    frmQuest.lblNpcAmount = NpcKillAmount
    frmQuest.lblNpcName = NpcKillName
    frmQuest.Show
    frmQuest.picNpcQuests.Visible = True
    Exit Sub
    End If
    
    If (LCase(parse(0)) = "questinfo") Then
    CurrentQuestNum = Val(parse(1))
    CurrentQuestNpcNum = Val(parse(2))
    Exit Sub
    End If
    
    If (LCase(parse(0)) = "questmsg") Then
    Dim Hate As Byte
    Call AddQuestText(parse(1), Val(parse(2)))
    Hate = Val(parse(3))
    Call QuestPrompt(Hate)
    Exit Sub
    End If

If (LCase(parse(0)) = "questflags") Then
       ' frmMirage.lstIndex.Clear
       ' frmMirage.picListQuest.Visible = True
        For I = 1 To MAX_QUESTS
            Player(MyIndex).QuestFlags(I) = Val(parse(I))
            If Player(MyIndex).QuestFlags(I) <= 0 Then
             '   frmMirage.lstIndex.AddItem "" & I & ". Quest (" & Trim(Quest(I).Name) & ")  -Not Started-"
            ElseIf Player(MyIndex).QuestFlags(I) = 1 Then
             '   frmMirage.lstIndex.AddItem "" & I & ". Quest (" & Trim(Quest(I).Name) & ")  -Quest Started-"
            ElseIf Player(MyIndex).QuestFlags(I) >= 2 Then
              '  frmMirage.lstIndex.AddItem "" & I & ". Quest (" & Trim(Quest(I).Name) & ")  -Quest Completed-"
            End If
        Next I
        QuestIndex = 1
    End If

If (LCase(parse(0)) = "updatequest") Then
n = Val(parse(1))

'Update the quest
Quest(n).name = parse(2)
Quest(n).After = parse(3)
Quest(n).Before = parse(4)
Quest(n).ClassIsReq = Val(parse(5))
Quest(n).ClassReq = Val(parse(6))
Quest(n).During = parse(7)
Quest(n).End = parse(8)
Quest(n).ItemReq = Val(parse(9))
Quest(n).ItemVal = Val(parse(10))
Quest(n).LevelIsReq = Val(parse(11))
Quest(n).LevelReq = Val(parse(12))
Quest(n).NotHasItem = parse(13)
Quest(n).RewardNum = Val(parse(14))
Quest(n).RewardVal = Val(parse(15))
Quest(n).Start = parse(16)
Quest(n).StartItem = Val(parse(17))
Quest(n).StartOn = Val(parse(18))
Quest(n).Startval = Val(parse(19))
End If

If (LCase(parse(0)) = "questprompt") Then
Dim Awnser2 As Variant
        
        Awnser2 = MsgBox(Quest(parse(1)).During, vbYesNo, Quest(parse(1)).name)
        
        If Awnser2 = 7 Then
            
        Else
            
            If HasItem(Quest(parse(1)).ItemReq, Quest(parse(1)).ItemVal) Then
                Call SendData("questdone" & SEP_CHAR & parse(1) & SEP_CHAR & MyIndex & SEP_CHAR & parse(2) & SEP_CHAR & END_CHAR)
            Else
                Call MsgBox(Quest(parse(1)).NotHasItem, vbInformation, Quest(parse(1)).name)
            End If
        
        End If
End If

    If (casestring = "playernewxy") Then
        X = Val(parse(1))
        Y = Val(parse(2))

        If Not GetPlayerX(MyIndex) = X Then Call SetPlayerX(MyIndex, X)
        If Not GetPlayerY(MyIndex) = Y Then Call SetPlayerY(MyIndex, Y)

        Exit Sub
    End If
End Sub

Public Sub AddQuestText(ByVal Msg As String, ByVal color As Integer)
Dim S As String
Dim c As Long
Dim filename As String
filename = App.Path & "\Main\Config\config.ini"
  
    S = vbNewLine & Msg
    c = frmQuest.txtChat.SelStart
    frmQuest.txtChat.SelStart = Len(frmQuest.txtChat.Text)
    frmQuest.txtChat.SelColor = QBColor(color)
    frmQuest.txtChat.SelText = S
    frmQuest.txtChat.SelStart = Len(frmQuest.txtChat.Text) - 1
    'If Val(GetVar(filename, "CONFIG", "VideoMemory")) = 0 Then frmQuest.txtChat.SelStart = C
End Sub
