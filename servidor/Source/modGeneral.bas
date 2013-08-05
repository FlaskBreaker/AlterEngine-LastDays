Attribute VB_Name = "modGeneral"
Option Explicit

Sub InitServer()
    Dim index As Integer

    Call SetStatus("Comprobando carpetas...")

    ' Check folders
    If Not FolderExists(App.Path & "\Mapas") Then
        Call MkDir(App.Path & "\Mapas")
    End If

    If Not FolderExists(App.Path & "\Logs") Then
        Call MkDir(App.Path & "\Logs")
    End If

    If Not FolderExists(App.Path & "\Cuentas") Then
        Call MkDir(App.Path & "\Cuentas")
    End If

    If Not FolderExists(App.Path & "\NPCs") Then
        Call MkDir(App.Path & "\NPCs")
    End If

    If Not FolderExists(App.Path & "\Objetos") Then
        Call MkDir(App.Path & "\Objetos")
    End If

    If Not FolderExists(App.Path & "\Hechizos") Then
        Call MkDir(App.Path & "\Hechizos")
    End If

    If Not FolderExists(App.Path & "\Tiendas") Then
        Call MkDir(App.Path & "\Tiendas")
    End If

    If Not FolderExists(App.Path & "\Bancos") Then
        Call MkDir(App.Path & "\Bancos")
    End If
    
    If Not FolderExists(App.Path & "\Clases") Then
        Call MkDir(App.Path & "\Clases")
    End If
    
    Call SetStatus("Comprobando archivos...")

    If Not FileExists("Configuracion.ini") Then
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "GameName", "AlterEngine"
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "WebSite", vbNullString
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "Port", 4001
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "HPRegen", 1
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "HPTimer", 1000
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "MPRegen", 1
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "MPTimer", 1000
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "SPRegen", 1
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "SPTimer", 1000
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "NPCRegen", 1
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "Stat1", "Fuerza"
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "Stat2", "Defensa"
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "Stat3", "Velocidad"
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "Stat4", "Magia"
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "PlayerCard", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "Scrolling", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "ScrollX", 30
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "ScrollY", 30
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "Scripting", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "ScriptErrors", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "PaperDoll", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "SaveTime", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "SpriteSize", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "Custom", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "Level", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "PKMinLvl", 10
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "Email", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "VerifyAcc", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "Classes", 1
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "SPAttack", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "SPRunning", 0
        PutVar App.Path & "\Configuracion.ini", "MAX", "MAX_PLAYERS", 200
        PutVar App.Path & "\Configuracion.ini", "MAX", "MAX_ITEMS", 100
        PutVar App.Path & "\Configuracion.ini", "MAX", "MAX_NPCS", 100
        PutVar App.Path & "\Configuracion.ini", "MAX", "MAX_SHOPS", 100
        PutVar App.Path & "\Configuracion.ini", "MAX", "MAX_SPELLS", 100
        PutVar App.Path & "\Configuracion.ini", "MAX", "MAX_MAPS", 500
        PutVar App.Path & "\Configuracion.ini", "MAX", "MAX_MAP_ITEMS", 20
        PutVar App.Path & "\Configuracion.ini", "MAX", "MAX_GUILDS", 20
        PutVar App.Path & "\Configuracion.ini", "MAX", "MAX_GUILD_MEMBERS", 10
        PutVar App.Path & "\Configuracion.ini", "MAX", "MAX_EMOTICONS", 10
        PutVar App.Path & "\Configuracion.ini", "MAX", "MAX_LEVEL", 500
        PutVar App.Path & "\Configuracion.ini", "MAX", "MAX_PARTY_MEMBERS", 4
        PutVar App.Path & "\Configuracion.ini", "MAX", "MAX_ELEMENTS", 20
        PutVar App.Path & "\Configuracion.ini", "MAX", "MAX_SCRIPTSPELLS", 500
        PutVar App.Path & "\Configuracion.ini", "MAX", "Max_HEAD", 50
        PutVar App.Path & "\Configuracion.ini", "MAX", "MAX_BODY", 50
        PutVar App.Path & "\Configuracion.ini", "MAX", "MAX_LEGS", 50
        
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "UsercrR", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "UsercrG", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "UsercrB", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "ModCrR", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "ModCrG", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "ModCrB", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "MapperCrR", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "MapperCrG", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "MapperCrB", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "DeveloperCrR", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "DeveloperCrG", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "DeveloperCrB", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "AdminCrR", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "AdminCrG", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "AdminCrB", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "OwnerCrR", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "OwnerCrG", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "OwnerCrB", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "PKCrR", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "PKCrG", 0
        PutVar App.Path & "\Configuracion.ini", "CONFIG", "PKCrB", 0
        
        
        
        
    End If

    If Not FileExists("Estados.ini") Then
        PutVar App.Path & "\Estados.ini", "HP", "AddPerLevel", 10
        PutVar App.Path & "\Estados.ini", "HP", "AddPerStr", 10
        PutVar App.Path & "\Estados.ini", "HP", "AddPerDef", 0
        PutVar App.Path & "\Estados.ini", "HP", "AddPerMagi", 0
        PutVar App.Path & "\Estados.ini", "HP", "AddPerSpeed", 0
        PutVar App.Path & "\Estados.ini", "MP", "AddPerLevel", 10
        PutVar App.Path & "\Estados.ini", "MP", "AddPerStr", 0
        PutVar App.Path & "\Estados.ini", "MP", "AddPerDef", 0
        PutVar App.Path & "\Estados.ini", "MP", "AddPerMagi", 10
        PutVar App.Path & "\Estados.ini", "MP", "AddPerSpeed", 0
        PutVar App.Path & "\Estados.ini", "SP", "AddPerLevel", 10
        PutVar App.Path & "\Estados.ini", "SP", "AddPerStr", 0
        PutVar App.Path & "\Estados.ini", "SP", "AddPerDef", 0
        PutVar App.Path & "\Estados.ini", "SP", "AddPerMagi", 0
        PutVar App.Path & "\Estados.ini", "SP", "AddPerSpeed", 20
    End If

    If Not FileExists("Noticias.ini") Then
        PutVar App.Path & "\Noticias.ini", "DATA", "NewsTitle", "Cambia el titulo en Noticias.ini."
        PutVar App.Path & "\Noticias.ini", "DATA", "NewsBody", "Cambia el mensaje en Noticias.ini."
        PutVar App.Path & "\Noticias.ini", "COLOR", "Red", 255
        PutVar App.Path & "\Noticias.ini", "COLOR", "Green", 255
        PutVar App.Path & "\Noticias.ini", "COLOR", "Blue", 255
    End If
    
    If Not FileExists("MOTD.ini") Then
        PutVar App.Path & "\MOTD.ini", "MOTD", "Msg", "Cambia este mensaje en MOTD.ini."
    End If

    If Not FileExists("Tiles.ini") Then
        For index = 0 To 100
            PutVar App.Path & "\Tiles.ini", "Names", "Tile" & index, CStr(index)
        Next index
    End If

    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExists("Cuentas\Charlist.txt") Then
        index = FreeFile
        Open App.Path & "\Cuentas\CharList.txt" For Output As #index
        Close #index
    End If

    Call SetStatus("Cargando configuraciÛn...")

    On Error GoTo LoadErr
    addHP.Level = Val(GetVar(App.Path & "\Estados.ini", "HP", "AddPerLevel"))
    addHP.STR = Val(GetVar(App.Path & "\Estados.ini", "HP", "AddPerStr"))
    addHP.DEF = Val(GetVar(App.Path & "\Estados.ini", "HP", "AddPerDef"))
    addHP.Magi = Val(GetVar(App.Path & "\Estados.ini", "HP", "AddPerMagi"))
    addHP.Speed = Val(GetVar(App.Path & "\Estados.ini", "HP", "AddPerSpeed"))
    addMP.Level = Val(GetVar(App.Path & "\Estados.ini", "MP", "AddPerLevel"))
    addMP.STR = Val(GetVar(App.Path & "\Estados.ini", "MP", "AddPerStr"))
    addMP.DEF = Val(GetVar(App.Path & "\Estados.ini", "MP", "AddPerDef"))
    addMP.Magi = Val(GetVar(App.Path & "\Estados.ini", "MP", "AddPerMagi"))
    addMP.Speed = Val(GetVar(App.Path & "\Estados.ini", "MP", vbNullString))
    addSP.Level = Val(GetVar(App.Path & "\Estados.ini", "SP", "AddPerLevel"))
    addSP.STR = Val(GetVar(App.Path & "\Estados.ini", "SP", "AddPerStr"))
    addSP.DEF = Val(GetVar(App.Path & "\Estados.ini", "SP", "AddPerDef"))
    addSP.Magi = Val(GetVar(App.Path & "\Estados.ini", "SP", "AddPerMagi"))
    addSP.Speed = Val(GetVar(App.Path & "\Estados.ini", "SP", "AddPerSpeed"))

    GAME_NAME = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "GameName")
    GAME_PORT = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Port")
    MAX_PLAYERS = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_PLAYERS")
    MAX_ITEMS = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_ITEMS")
    MAX_NPCS = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_NPCS")
    MAX_SHOPS = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_SHOPS")
    MAX_SPELLS = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_SPELLS")
    MAX_MAPS = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_MAPS")
    MAX_MAP_ITEMS = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_MAP_ITEMS")
    MAX_GUILDS = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_GUILDS")
    MAX_GUILD_MEMBERS = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_GUILD_MEMBERS")
    MAX_EMOTICONS = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_EMOTICONS")
    MAX_LEVEL = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_LEVEL")
    MAX_ELEMENTS = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_ELEMENTS")
    MAX_SCRIPTSPELLS = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_SCRIPTSPELLS")
    scripting = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Scripting")
    paperdoll = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "PaperDoll")
    SPRITESIZE = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "SpriteSize")
    HP_REGEN = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "HPRegen")
    HP_TIMER = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "HPTimer")
    MP_REGEN = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "MPRegen")
    MP_TIMER = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "MPTimer")
    SP_REGEN = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "SPRegen")
    SP_TIMER = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "SPTimer")
    NPC_REGEN = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "NPCRegen")
    stat1 = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Stat1")
    stat2 = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Stat2")
    stat3 = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Stat3")
    stat4 = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Stat4")
    MAX_HEAD = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_HEAD")
    MAX_BODY = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_BODY")
    MAX_LEGS = GetVar(App.Path & "\Configuracion.ini", "MAX", "MAX_LEGS")
    SP_ATTACK = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "SPAttack")
    SP_RUNNING = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "SPRunning")
    CUSTOM_SPRITE = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Custom")
    EMAIL_AUTH = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Email")
    SAVETIME = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "SaveTime")
    Level = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Level")
    PKMINLVL = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "PKMinLvl")
    ACC_VERIFY = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "VerifyAcc")
    CLASSES = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Classes")
    UserCr(1) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "UserCrR"))
    UserCr(2) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "UserCrG"))
    UserCr(3) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "UserCrB"))
    ModCr(1) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "ModCrR"))
    ModCr(2) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "ModCrG"))
    ModCr(3) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "ModCrB"))
    MapperCr(1) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "MapperCrR"))
    MapperCr(2) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "MapperCrG"))
    MapperCr(3) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "MapperCrB"))
    DeveloperCr(1) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "DeveloperCrR"))
    DeveloperCr(2) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "DeveloperCrG"))
    DeveloperCr(3) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "DeveloperCrB"))
    AdminCr(1) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "AdminCrR"))
    AdminCr(2) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "AdminCrG"))
    AdminCr(3) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "AdminCrB"))
    OwnerCr(1) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "OwnerCrR"))
    OwnerCr(2) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "OwnerCrG"))
    OwnerCr(3) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "OwnerCrB"))
    PKCr(1) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "PKCrR"))
    PKCr(2) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "PKCrG"))
    PKCr(3) = CByte(GetVar(App.Path & "\Configuracion.ini", "CONFIG", "PKCrB"))

    If GetVar(App.Path & "\Configuracion.ini", "CONFIG", "Scrolling") = 0 Then
        IS_SCROLLING = 0
        MAX_MAPX = 19
        MAX_MAPY = 14
    Else
        IS_SCROLLING = 1
        MAX_MAPX = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "ScrollX")
        MAX_MAPY = GetVar(App.Path & "\Configuracion.ini", "CONFIG", "ScrollY")
    End If

    ' Weather variables.
    WeatherType = WEATHER_NONE
    WeatherLevel = 25

    SEP_CHAR = Chr$(0)
    END_CHAR = Chr$(237)

    ServerLog = True
    
    GoTo LoadSuccess

LoadErr:
    Call MsgBox("Error reading from Configuracion.ini or Estados.ini.", vbOKOnly)
    End

LoadSuccess:
    ' Restore error handling
    On Error GoTo 0

    ReDim Map(1 To MAX_MAPS) As MapRec
    ReDim MapCache(1 To MAX_MAPS) As String
    ReDim TempTile(1 To MAX_MAPS) As TempTileRec
    ReDim PlayersOnMap(1 To MAX_MAPS) As Long
    ReDim Player(1 To MAX_PLAYERS) As AccountRec
    ReDim Item(0 To MAX_ITEMS) As ItemRec
    ReDim NPC(0 To MAX_NPCS) As NpcRec
    ReDim MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
    ReDim MapNPC(1 To MAX_MAPS, 1 To MAX_MAP_NPCS) As MapNpcRec
    ReDim Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Spell(1 To MAX_SPELLS) As SpellRec
    ReDim Guild(1 To MAX_GUILDS) As GuildRec
    ReDim Emoticons(0 To MAX_EMOTICONS) As EmoRec
    ReDim Element(0 To MAX_ELEMENTS) As ElementRec

    For index = 1 To MAX_GUILDS
        ReDim Guild(index).Member(1 To MAX_GUILD_MEMBERS) As String * NAME_LENGTH
    Next index
    
    For index = 1 To MAX_MAPS
        ReDim Map(index).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        ReDim TempTile(index).DoorOpen(0 To MAX_MAPX, 0 To MAX_MAPY) As Byte
    Next index

    ReDim Experience(1 To MAX_LEVEL) As Long

    START_MAP = 1
    START_X = MAX_MAPX / 2
    START_Y = MAX_MAPY / 2

    Set CTimers = New Collection

    Call IncrementBar

    On Error GoTo ScriptErr

    ' Scripting
    frmServer.lblScriptOn.Caption = "Scripts: OFF"
    
    ' Check for Main.txt
    If Not FileExists("\Scripts\Main.txt") Then
        Call MsgBox("Main.txt no encontrado. Scripts desactivados.", vbExclamation)
        scripting = 0
    End If
    
    ' Continue
    If scripting = 1 Then
        Call SetStatus("Loading scripts...")
        Set MyScript = New clsSadScript
        Set clsScriptCommands = New clsCommands
        MyScript.ReadInCode App.Path & "\Scripts\Main.txt", "Scripts\Main.txt", MyScript.SControl
        MyScript.SControl.AddObject "ScriptHardCode", clsScriptCommands, True
        frmServer.lblScriptOn.Caption = "Scripts: ON"
    End If

    Call IncrementBar

    GoTo ScriptsGood

ScriptErr:
    Call MsgBox("Error cargando el motor de scripting.", vbOKOnly)
    End

ScriptsGood:

    On Error GoTo 0

    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = GAME_PORT

    ' Init all the player sockets
    Call SetStatus("Comenzando conexiones...")
    For index = 1 To MAX_PLAYERS
        Call ClearPlayer(index)

    Load frmServer.Socket(index)
    Next index

    For index = 1 To MAX_PLAYERS
        Call ShowPLR(index)
    Next index

    Call IncrementBar


    Call SetStatus("Limpiando tiles sucias...")
    Call ClearTempTile
    Call SetStatus("Limpiando mapas...")
    Call ClearMaps
    Call SetStatus("Limpiando objetos de mapa...")
    Call ClearMapItems
    Call SetStatus("Limpiando npcs de mapa...")
    Call ClearMapNpcs
    Call SetStatus("Limpiando npcs...")
    Call ClearNpcs
    Call SetStatus("Limpiando objetos...")
    Call ClearItems
    Call SetStatus("Limpiando tiendas...")
    Call ClearShops
    Call SetStatus("Limpiando hechizos...")
    Call ClearSpells
    Call SetStatus("Limpiando exp...")
    Call ClearExperience
    Call SetStatus("Limpiando emoticonos...")
    Call ClearEmoticon
    Call IncrementBar
    Call SetStatus("Cargando emoticonos..")
    Call IncrementBar
    Call LoadEmoticon
    Call SetStatus("Cargando elementos...")
    Call IncrementBar
    Call LoadElements
    Call SetStatus("Limpiando flechas...")
    Call ClearArrows
    Call SetStatus("Cargando flechas...")
    Call IncrementBar
    Call LoadArrows
    Call SetStatus("Cargando exp...")
    Call IncrementBar
    Call LoadExperience
    Call SetStatus("Cargando clases...")
    Call IncrementBar
    Call LoadClasses
    Call SetStatus("Cargando mapas...")
    Call IncrementBar
    Call LoadMaps
    Call SetStatus("Cargando objetos...")
    Call IncrementBar
    Call LoadItems
    Call SetStatus("Cargando npcs...")
    Call IncrementBar
    Call LoadNpcs
    Call SetStatus("Cargando Quests...")
    Call IncrementBar
    Call LoadQuests
    Call SetStatus("Cargando tiendas...")
    Call IncrementBar
    Call LoadShops
    Call SetStatus("Cargando hechizos...")
    Call IncrementBar
    Call LoadSpells
    Call SetStatus("Spawneando objetos en mapa...")
    Call SpawnAllMapsItems
    Call SetStatus("Spawneando NPCS...")
    Call SpawnAllMapNpcs
    Call IncrementBar
    
    ' Funcion por la cual se conecta a AE.net y comprueba la versiÛn mediante un TXT
    ' Para modificar simplemente cambiar el host y subir un archivo con el contenido "vX"
    ' Donde "X" cambiar por la versiÛn ejmplo: v1.0, v1.1. Asi como tambiÈn cambiar en el
    ' Textbox "ae_versionactual", para que los compare entresi.
    
    Call SetStatus("Cargando VersiÛn...")
    frmServer.ae_actualizaciones.Text = frmServer.tolpene.OpenURL("http://www.alterengine.net/internet/version.txt")
    If frmServer.ae_actualizaciones.Text = frmServer.ae_versionactual.Text Then
    frmServer.tieneslaultima.Visible = True
    frmServer.notienes.Visible = False
    Else
    frmServer.tieneslaultima.Visible = False
    frmServer.notienes.Visible = True
    frmServer.versionnueva.Visible = True
    'Call MsgBox("Una nueva versiÛn de AE ha sido lanzada, descargatela de www.alterengine.net.", vbOKOnly)
    End If
    Call IncrementBar

    frmServer.MapList.Clear

    For index = 1 To MAX_MAPS
        frmServer.MapList.AddItem index & ": " & Map(index).Name
    Next index

    frmServer.MapList.Selected(0) = True
    frmServer.tmrPlayerSave.Enabled = True
    frmServer.tmrSpawnMapItems.Enabled = True
    frmServer.Timer1.Enabled = True

    ' Error handling for 'Address in use' error
    Err.Clear
    On Error Resume Next
    
    ' Start listening
    frmServer.Socket(0).Listen

    ' RTE 10048 occured
    If Err.Number = 10048 Then
        Call MsgBox("El puerto actualmente esta siendo usado.", vbOKOnly)
        End
    End If
    
    If scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "OnServerLoad"
    End If

    ' Restore error handling
    On Error GoTo 0

    Call UpdateTitle
    Call UpdateTOP

    frmLoad.Visible = False
    frmServer.Show

    SpawnSeconds = 0
    frmServer.tmrGameAI.Enabled = True
    frmServer.tmrScriptedTimer.Enabled = True
    
End Sub

Sub DestroyServer()
    Dim I As Long
    
    Call SaveAllPlayersOnline

    frmServer.Visible = False
    frmLoad.Visible = True

    For I = 1 To MAX_PLAYERS
        temp = I / MAX_PLAYERS * 100
        Call SetStatus("Liberando sockets... " & temp & "%")
        Unload frmServer.Socket(I)
    Next I

    End
End Sub

Sub SetStatus(ByVal Status As String)
    frmLoad.lblStatus.Caption = Status
    DoEvents
End Sub

Sub IncrementBar()
    On Error Resume Next
    ' Increment prog bar
    frmLoad.loadProgressBar.Value = frmLoad.loadProgressBar.Value + 1
End Sub

Sub ServerLogic()
    Call CheckGiveVitals
    Call GameAI
    Call ScriptedTimer
End Sub

Sub CheckSpawnMapItems()
    Dim X As Long
    Dim Y As Long

    ' Used for map item respawning
    SpawnSeconds = SpawnSeconds + 1

    ' Respawns the map items.
    If SpawnSeconds >= 120 Then
        ' 2 minutes have passed
        For Y = 1 To MAX_MAPS
            ' Make sure no one is on the map when it respawns
            If PlayersOnMap(Y) = NO Then
                ' Clear out unnecessary junk
                For X = 1 To MAX_MAP_ITEMS
                    Call ClearMapItem(X, Y)
                Next X

                ' Spawn the items
                Call SpawnMapItems(Y)
                Call SendMapItemsToAll(Y)
            End If
        Next Y

        SpawnSeconds = 0
    End If
End Sub

Sub GameAI()
Dim I As Long, X As Long, Y As Long, N As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, TickCount As Long
Dim Damage As Long, DistanceX As Long, DistanceY As Long, npcnum As Long, Target As Long
Dim DidWalk As Boolean

    'WeatherSeconds = WeatherSeconds + 1
    'TimeSeconds = TimeSeconds + 1
    ' Lets change the weather if its time to
   ' If WeatherSeconds >= 60 Then
     '   i = Int(Rnd * 3)

     '   If i <> GameWeather Then
     '       GameWeather = i
      '      Call SendWeatherToAll
      '  End If
      '  WeatherSeconds = 0
   ' End If

    ' Check if we need to switch from day to night or night to day
   ' If TimeSeconds >= 60 Then
     '   If GameTime = TIME_DAY Then
     '       GameTime = TIME_NIGHT
     '   Else
     '       GameTime = TIME_DAY
     '   End If
     ''   Call SendTimeToAll
      '  TimeSeconds = 0
    'End If
    For Y = 1 To MAX_MAPS

        If PlayersOnMap(Y) = YES Then
            TickCount = GetTickCount

            ' ////////////////////////////////////
            ' // This is used for closing doors //
            ' ////////////////////////////////////
            If TickCount > TempTile(Y).DoorTimer + 5000 Then
                For Y1 = 0 To MAX_MAPY
                    For X1 = 0 To MAX_MAPX

                        If Map(Y).Tile(X1, Y1).Type = TILE_TYPE_KEY And TempTile(Y).DoorOpen(X1, Y1) = YES Then
                            TempTile(Y).DoorOpen(X1, Y1) = NO
                            Call SendDataToMap(Y, "MAPKEY" & SEP_CHAR & X1 & SEP_CHAR & Y1 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                        End If

                        If Map(Y).Tile(X1, Y1).Type = TILE_TYPE_DOOR And TempTile(Y).DoorOpen(X1, Y1) = YES Then
                            TempTile(Y).DoorOpen(X1, Y1) = NO
                            Call SendDataToMap(Y, "MAPKEY" & SEP_CHAR & X1 & SEP_CHAR & Y1 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
                        End If
                    Next
                Next
            End If
            For X = 1 To MAX_MAP_NPCS
                npcnum = MapNPC(Y, X).num

                ' /////////////////////////////////////////
                ' // This is used for ATTACKING ON SIGHT //
                ' /////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).NPC(X) > 0 And MapNPC(Y, X).num > 0 Then
                
                If NPC(npcnum).Behavior = NPC_BEHAVIOR_GUARD Then
                        For I = 1 To MAX_PLAYERS

                            If IsPlaying(I) Then
                                If GetPlayerMap(I) = Y And MapNPC(Y, X).Target = 0 And GetPlayerAccess(I) <= ADMIN_MONITER Then
                                    N = NPC(npcnum).Range
                                    DistanceX = MapNPC(Y, X).X - GetPlayerX(I)
                                    DistanceY = MapNPC(Y, X).Y - GetPlayerY(I)

                                    ' Make sure we get a positive value
                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                    If DistanceY < 0 Then DistanceY = DistanceY * -1

                                    ' Are they in range?  if so GET'M!
                                    If DistanceX <= N And DistanceY <= N Then
                                        If GetPlayerPK(I) = YES Then
                                            If Trim$(NPC(npcnum).AttackSay) <> "" Then
                                                Call PlayerMsg(I, "A " & Trim$(NPC(npcnum).Name) & " : " & Trim$(NPC(npcnum).AttackSay) & "", SayColor)
                                            End If
                                            MapNPC(Y, X).TargetType = TARGET_TYPE_PLAYER
                                            MapNPC(Y, X).Target = I
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                    
                    Dim spellnum As Long
                    Dim Victim As Long

                    ' If the npc is a attack on sight, search for a player on the map
                    If NPC(npcnum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Then
                        For I = 1 To MAX_PLAYERS

                            If IsPlaying(I) Then
                                If GetPlayerMap(I) = Y And MapNPC(Y, X).Target = 0 And GetPlayerAccess(I) <= ADMIN_MONITER Then
                                    N = NPC(npcnum).Range
                                    DistanceX = MapNPC(Y, X).X - GetPlayerX(I)
                                    DistanceY = MapNPC(Y, X).Y - GetPlayerY(I)

                                    ' Make sure we get a positive value
                                    If DistanceX < 0 Then DistanceX = DistanceX * -1
                                    If DistanceY < 0 Then DistanceY = DistanceY * -1

                                    ' Are they in range?  if so GET'M!
                                    If DistanceX <= N And DistanceY <= N Then
                                        If NPC(npcnum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Then
                                            If Trim$(NPC(npcnum).AttackSay) <> "" Then
                                                Call PlayerMsg(I, "A " & Trim$(NPC(npcnum).Name) & " : " & Trim$(NPC(npcnum).AttackSay) & "", SayColor)
                                            End If
                                            MapNPC(Y, X).TargetType = TARGET_TYPE_PLAYER
                                            MapNPC(Y, X).Target = I
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
                

                ' /////////////////////////////////////////////
                ' // This is used for NPC walking/targetting //
                ' /////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).NPC(X) > 0 And MapNPC(Y, X).num > 0 Then
                    Target = MapNPC(Y, X).Target

                    ' Check to see if its time for the npc to walk
                    If NPC(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(npcnum).standstill = False Then

                        ' Check to see if we are following a player or not
                        If Target > 0 Then
                            If MapNPC(Y, X).TargetType = TARGET_TYPE_PLAYER Then

                                ' Check if the player is even playing, if so follow'm
                                If IsPlaying(Target) And GetPlayerMap(Target) = Y Then
                                    DidWalk = False
                                    I = Int(Rnd * 5)

                                    ' Lets move the npc
                                    Select Case I

                                        Case 0

                                            ' Up
                                            If MapNPC(Y, X).Y > GetPlayerY(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_UP) Then
                                                    Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Down
                                            If MapNPC(Y, X).Y < GetPlayerY(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_DOWN) Then
                                                    Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Left
                                            If MapNPC(Y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_LEFT) Then
                                                    Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Right
                                            If MapNPC(Y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                    Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                        Case 1

                                            ' Right
                                            If MapNPC(Y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                    Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Left
                                            If MapNPC(Y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_LEFT) Then
                                                    Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Down
                                            If MapNPC(Y, X).Y < GetPlayerY(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_DOWN) Then
                                                    Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Up
                                            If MapNPC(Y, X).Y > GetPlayerY(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_UP) Then
                                                    Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                        Case 2

                                            ' Down
                                            If MapNPC(Y, X).Y < GetPlayerY(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_DOWN) Then
                                                    Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Up
                                            If MapNPC(Y, X).Y > GetPlayerY(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_UP) Then
                                                    Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Right
                                            If MapNPC(Y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                    Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Left
                                            If MapNPC(Y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_LEFT) Then
                                                    Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                        Case 3

                                            ' Left
                                            If MapNPC(Y, X).X > GetPlayerX(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_LEFT) Then
                                                    Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Right
                                            If MapNPC(Y, X).X < GetPlayerX(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                    Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Up
                                            If MapNPC(Y, X).Y > GetPlayerY(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_UP) Then
                                                    Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Down
                                            If MapNPC(Y, X).Y < GetPlayerY(Target) And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_DOWN) Then
                                                    Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                    End Select

                                    ' Check if we can't move and if player is behind something and if we can just switch dirs
                                    If Not DidWalk Then
                                        If MapNPC(Y, X).X - 1 = GetPlayerX(Target) And MapNPC(Y, X).Y = GetPlayerY(Target) Then
                                            If MapNPC(Y, X).Dir <> DIR_LEFT Then
                                                Call NpcDir(Y, X, DIR_LEFT)
                                            End If
                                            DidWalk = True
                                        End If

                                        If MapNPC(Y, X).X + 1 = GetPlayerX(Target) And MapNPC(Y, X).Y = GetPlayerY(Target) Then
                                            If MapNPC(Y, X).Dir <> DIR_RIGHT Then
                                                Call NpcDir(Y, X, DIR_RIGHT)
                                            End If
                                            DidWalk = True
                                        End If

                                        If MapNPC(Y, X).X = GetPlayerX(Target) And MapNPC(Y, X).Y - 1 = GetPlayerY(Target) Then
                                            If MapNPC(Y, X).Dir <> DIR_UP Then
                                                Call NpcDir(Y, X, DIR_UP)
                                            End If
                                            DidWalk = True
                                        End If

                                        If MapNPC(Y, X).X = GetPlayerX(Target) And MapNPC(Y, X).Y + 1 = GetPlayerY(Target) Then
                                            If MapNPC(Y, X).Dir <> DIR_DOWN Then
                                                Call NpcDir(Y, X, DIR_DOWN)
                                            End If
                                            DidWalk = True
                                        End If

                                        ' We could not move so player must be behind something, walk randomly.
                                        If Not DidWalk Then
                                            I = Int(Rnd * 2)

                                            If I = 1 Then
                                                I = Int(Rnd * 4)

                                                If CanNpcMove(Y, X, I) Then
                                                    Call NpcMove(Y, X, I, MOVING_WALKING)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    MapNPC(Y, X).Target = 0
                                End If
                            Else

                                ' Check if the pet is even playing, if so follow'm
                                If IsPlaying(Target) And Player(Target).Pet.Map = Y Then
                                    DidWalk = False
                                    I = Int(Rnd * 5)

                                    ' Lets move the npc
                                    Select Case I

                                        Case 0

                                            ' Up
                                            If MapNPC(Y, X).Y > Player(Target).Pet.Y And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_UP) Then
                                                    Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Down
                                            If MapNPC(Y, X).Y < Player(Target).Pet.Y And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_DOWN) Then
                                                    Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Left
                                            If MapNPC(Y, X).X > Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_LEFT) Then
                                                    Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Right
                                            If MapNPC(Y, X).X < Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                    Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                        Case 1

                                            ' Right
                                            If MapNPC(Y, X).X < Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                    Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Left
                                            If MapNPC(Y, X).X > Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_LEFT) Then
                                                    Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Down
                                            If MapNPC(Y, X).Y < Player(Target).Pet.Y And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_DOWN) Then
                                                    Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Up
                                            If MapNPC(Y, X).Y > Player(Target).Pet.Y And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_UP) Then
                                                    Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                        Case 2

                                            ' Down
                                            If MapNPC(Y, X).Y < Player(Target).Pet.Y And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_DOWN) Then
                                                    Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Up
                                            If MapNPC(Y, X).Y > Player(Target).Pet.Y And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_UP) Then
                                                    Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Right
                                            If MapNPC(Y, X).X < Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                    Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Left
                                            If MapNPC(Y, X).X > Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_LEFT) Then
                                                    Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                        Case 3

                                            ' Left
                                            If MapNPC(Y, X).X > Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_LEFT) Then
                                                    Call NpcMove(Y, X, DIR_LEFT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Right
                                            If MapNPC(Y, X).X < Player(Target).Pet.X And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_RIGHT) Then
                                                    Call NpcMove(Y, X, DIR_RIGHT, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Up
                                            If MapNPC(Y, X).Y > Player(Target).Pet.Y And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_UP) Then
                                                    Call NpcMove(Y, X, DIR_UP, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If

                                            ' Down
                                            If MapNPC(Y, X).Y < Player(Target).Pet.Y And DidWalk = False Then
                                                If CanNpcMove(Y, X, DIR_DOWN) Then
                                                    Call NpcMove(Y, X, DIR_DOWN, MOVING_WALKING)
                                                    DidWalk = True
                                                End If
                                            End If
                                    End Select

                                    ' Check if we can't move and if pet is behind something and if we can just switch dirs
                                    If Not DidWalk Then
                                        If MapNPC(Y, X).X - 1 = Player(Target).Pet.X And MapNPC(Y, X).Y = Player(Target).Pet.Y Then
                                            If MapNPC(Y, X).Dir <> DIR_LEFT Then
                                                Call NpcDir(Y, X, DIR_LEFT)
                                            End If
                                            DidWalk = True
                                        End If

                                        If MapNPC(Y, X).X + 1 = Player(Target).Pet.X And MapNPC(Y, X).Y = Player(Target).Pet.Y Then
                                            If MapNPC(Y, X).Dir <> DIR_RIGHT Then
                                                Call NpcDir(Y, X, DIR_RIGHT)
                                            End If
                                            DidWalk = True
                                        End If

                                        If MapNPC(Y, X).X = Player(Target).Pet.X And MapNPC(Y, X).Y - 1 = Player(Target).Pet.Y Then
                                            If MapNPC(Y, X).Dir <> DIR_UP Then
                                                Call NpcDir(Y, X, DIR_UP)
                                            End If
                                            DidWalk = True
                                        End If

                                        If MapNPC(Y, X).X = Player(Target).Pet.X And MapNPC(Y, X).Y + 1 = Player(Target).Pet.Y Then
                                            If MapNPC(Y, X).Dir <> DIR_DOWN Then
                                                Call NpcDir(Y, X, DIR_DOWN)
                                            End If
                                            DidWalk = True
                                        End If

                                        ' We could not move so pet must be behind something, walk randomly.
                                        If Not DidWalk Then
                                            I = Int(Rnd * 2)

                                            If I = 1 Then
                                                I = Int(Rnd * 4)

                                                If CanNpcMove(Y, X, I) Then
                                                    Call NpcMove(Y, X, I, MOVING_WALKING)
                                                End If
                                            End If
                                        End If
                                    End If
                                Else
                                    MapNPC(Y, X).Target = 0
                                End If
                            End If
                        Else
                            I = Int(Rnd * 4)

                            If I = 1 Then
                                I = Int(Rnd * 4)

                                If CanNpcMove(Y, X, I) Then
                                    Call NpcMove(Y, X, I, MOVING_WALKING)
                                End If
                            End If
                        End If
                    End If
                End If

                ' //////////////////////////////////////////////////////
                ' // This is used for npcs to attack players and pets //
                ' //////////////////////////////////////////////////////
                ' Make sure theres a npc with the map
                If Map(Y).NPC(X) > 0 And MapNPC(Y, X).num > 0 Then
                    Target = MapNPC(Y, X).Target

                    If MapNPC(Y, X).TargetType <> TARGET_TYPE_LOCATION And MapNPC(Y, X).TargetType <> TARGET_TYPE_NPC Then

                        ' Check if the npc can attack the targeted player player
                        If Target > 0 Then
                            If MapNPC(Y, X).TargetType = TARGET_TYPE_PLAYER Then

                                ' Is the target playing and on the same map?
                                If IsPlaying(Target) And GetPlayerMap(Target) = Y Then

                                    ' Can the npc attack the player?
                                    If CanNpcAttackPlayer(X, Target) Then
                                        If Not CanPlayerBlockHit(Target) Then
                                            Damage = NPC(npcnum).STR - GetPlayerProtection(Target) + (Rnd * 5) - 2

                                            If Damage > 0 Then
                                                Call NpcAttackPlayer(X, Target, Damage)
                                            Else
                                                Call BattleMsg(Target, "" & Trim$(NPC(npcnum).Name) & " no puede hacerte daÒo!", BRIGHTBLUE, 1)

                                                'Call PlayerMsg(Target, "The " & Trim$(Npc(NpcNum).Name) & "'s hit didn't even phase you!", BrightBlue)
                                            End If
                                        Else
                                            Call BattleMsg(Target, "Bloqueas el golpe de " & Trim$(NPC(npcnum).Name) & "!", BRIGHTCYAN, 1)

                                            'Call PlayerMsg(Target, "Your " & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerShieldSlot(Target))).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                            'Call PlayerMsg(Target, "Your " & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerLegsSlot(Target))).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                            'Call PlayerMsg(Target, "Your " & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerBootsSlot(Target))).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                            'Call PlayerMsg(Target, "Your " & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerGlovesSlot(Target))).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                            'Call PlayerMsg(Target, "Your " & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerRing1Slot(Target))).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                            'Call PlayerMsg(Target, "Your " & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerRing2Slot(Target))).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                            'Call PlayerMsg(Target, "Your " & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerAmuletSlot(Target))).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                            'Call PlayerMsg(Target, "Your " & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerWingsSlot(Target))).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                            'Call PlayerMsg(Target, "Your " & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerBeltSlot(Target))).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                            'Call PlayerMsg(Target, "Your " & Trim$(Item(GetPlayerInvItemNum(Target, GetPlayerCapeSlot(Target))).Name) & " blocks the " & Trim$(Npc(NpcNum).Name) & "'s hit!", BrightCyan)
                                            End If
                                    End If
                                Else

                                    ' Player left map or game, set target to 0
                                    MapNPC(Y, X).Target = 0
                                End If
                            Else

                                ' Is the target playing and on the same map?
                                If IsPlaying(Target) And Player(Target).Pet.Map = Y Then

                                    ' Can the npc attack the pet?
                                    If CanNpcAttackPet(X, Target) Then
                                        Damage = NPC(npcnum).STR - Player(Target).Pet.Level + (Rnd * 5) - 2

                                        If Damage > 0 Then
                                            Call NpcAttackPet(X, Target, Damage)
                                        End If
                                    End If
                                Else

                                    ' Pet left map or game, set target to 0
                                    MapNPC(Y, X).Target = 0
                                End If
                            End If
                        End If
                    End If
                End If

                ' ////////////////////////////////////////////
                ' // This is used for regenerating NPC's HP //
                ' ////////////////////////////////////////////
                ' Check to see if we want to regen some of the npc's hp
                If MapNPC(Y, X).num > 0 And TickCount > GiveNPCHPTimer + 10000 Then
                    If MapNPC(Y, X).HP > 0 Then
                        MapNPC(Y, X).HP = MapNPC(Y, X).HP + GetNpcHPRegen(npcnum)

                        ' Check if they have more then they should and if so just set it to max
                        If MapNPC(Y, X).HP > GetNpcMaxHP(npcnum) Then
                            MapNPC(Y, X).HP = GetNpcMaxHP(npcnum)
                        End If
                        Call SendDataToMap(Y, "NPCHP" & SEP_CHAR & X & SEP_CHAR & MapNPC(Y, X).HP & SEP_CHAR & GetNpcMaxHP(MapNPC(Y, X).num) & SEP_CHAR & END_CHAR)
                    End If
                End If

                ' ////////////////////////////////////////////////////////
                ' // This is used for checking if an NPC is dead or not //
                ' ////////////////////////////////////////////////////////
                ' Check if the npc is dead or not
                'If MapNpc(y, x).Num > 0 Then
                '    If MapNpc(y, x).HP <= 0 And Npc(MapNpc(y, x).Num).str > 0 And Npc(MapNpc(y, x).Num).DEF > 0 Then
                '        MapNpc(y, x).Num = 0
                '        MapNpc(y, x).SpawnWait = TickCount
                '   End If
                'End If
                ' //////////////////////////////////////
                ' // This is used for spawning an NPC //
                ' //////////////////////////////////////
                ' Check if we are supposed to spawn an npc or not
                If MapNPC(Y, X).num = 0 And Map(Y).NPC(X) > 0 Then
                    If TickCount > MapNPC(Y, X).SpawnWait + (NPC(Map(Y).NPC(X)).SpawnSecs * 1000) Then
                        Call SpawnNpc(X, Y)
                    End If
                End If

                If MapNPC(Y, X).num > 0 Then

                    ' If the NPC hasn't been fighting, why send it's HP?
                    If GetTickCount < MapNPC(Y, X).LastAttack + 6000 Then
                        Call SendDataToMap(Y, "NPCHP" & SEP_CHAR & X & SEP_CHAR & MapNPC(Y, X).HP & SEP_CHAR & GetNpcMaxHP(MapNPC(Y, X).num) & SEP_CHAR & END_CHAR)
                    End If
                End If
            Next
        End If
        DoEvents
    Next

    ' Make sure we reset the timer for npc hp regeneration
    If GetTickCount > GiveNPCHPTimer + 10000 Then
        GiveNPCHPTimer = GetTickCount
    End If

    ' Make sure we reset the timer for door closing
    If GetTickCount > KeyTimer + 15000 Then
        KeyTimer = GetTickCount
    End If

    ' //////////////////////////////////////////////////////////
    ' // Used for moving pets (it took a while it get going!) //
    ' //////////////////////////////////////////////////////////
    For X = 1 To MAX_PLAYERS
    
   ' If Player(x).CorpseMap > 0 Then
      '   If GetTickCount > CLng(Player(x).CorpseTimer + CLng((400000))) Then
      '    Call ClearCorpse(x)
       '   Call SendCorpseToAll(x)
       '  End If
       ' End If

        If Player(X).Pet.Alive = YES Then
            X1 = Player(X).Pet.X
            Y1 = Player(X).Pet.Y
            X2 = Player(X).Pet.XToGo
            Y2 = Player(X).Pet.YToGo

            If Player(X).Pet.Target > 0 Then
                If Player(X).Pet.TargetType = TARGET_TYPE_PLAYER Then
If CanPetAttackPlayer(X, Player(X).Pet.Target) Then
Damage = (Player(X).Pet.Level + GetPlayerSTR(X)) - GetPlayerProtection(Player(X).Pet.Target) + (Rnd * 20) - 5

If Damage > 0 Then
Call PetAttackPlayer(X, Player(X).Pet.Target, Damage)
X2 = X1
Y2 = Y1
End If

Else
X2 = GetPlayerX(Player(X).Pet.Target)
Y2 = GetPlayerY(Player(X).Pet.Target)
End If
End If

                If Player(X).Pet.TargetType = TARGET_TYPE_NPC Then
                    If CanPetAttackNpc(X, Player(X).Pet.Target) Then
                        Damage = Player(X).Pet.STR - NPC(Player(X).Pet.Target).STR + (Rnd * 5) - 2

                        If Damage > 0 Then
                            Call PetAttackNpc(X, Player(X).Pet.Target, Damage)
                            X2 = X1
                            Y2 = Y1
                        End If
                    Else
                        X2 = MapNPC(Player(X).Pet.Map, Player(X).Pet.Target).X
                        Y2 = MapNPC(Player(X).Pet.Map, Player(X).Pet.Target).Y
                    End If
                End If
            Else

                If Player(X).Pet.Map = GetPlayerMap(X) Or Player(X).Pet.MapToGo = 0 Then
                    If Player(X).Pet.XToGo = -1 Or Player(X).Pet.YToGo = -1 Then
                        I = Int(Rnd * 4)

                        If I = 1 Then
                            I = Int(Rnd * 4)

                            If I = DIR_UP Then
                                Y2 = Y1 - 1
                                X2 = Player(X).Pet.X
                            End If

                            If I = DIR_DOWN Then
                                Y2 = Y1 + 1
                                X2 = Player(X).Pet.X
                            End If

                            If I = DIR_RIGHT Then
                                X2 = X1 + 1
                                Y2 = Player(X).Pet.Y
                            End If

                            If I = DIR_LEFT Then
                                X2 = X1 - 1
                                Y2 = Player(X).Pet.Y
                            End If

                            If Not IsValid(X2, Y2) Then
                                X2 = X1
                                Y2 = Y1
                            End If
                           ' If Grid(Player(x).Pet.Map).Loc(X2, Y2).Blocked = True Then
                            '    X2 = X1
                            '    Y2 = Y1
                           ' End If
                        Else
                            X2 = X1
                            Y2 = Y1
                        End If
                    End If
                Else

                    If Map(Player(X).Pet.Map).Up = Player(X).Pet.MapToGo Then
                        Y2 = Y1 - 1
                    Else

                        If Map(Player(X).Pet.Map).Down = Player(X).Pet.MapToGo Then
                            Y2 = Y1 + 1
                        Else

                            If Map(Player(X).Pet.Map).Left = Player(X).Pet.MapToGo Then
                                X2 = X1 - 1
                            Else

                                If Map(Player(X).Pet.Map).Right = Player(X).Pet.MapToGo Then
                                    X2 = X1 + 1
                                Else
                                    I = Int(Rnd * 4)

                                    If I = 1 Then
                                        I = Int(Rnd * 4)

                                        If I = DIR_UP Then Y2 = Y1 - 1
                                        If I = DIR_DOWN Then Y2 = Y1 + 1
                                        If I = DIR_RIGHT Then X2 = X1 + 1
                                        If I = DIR_LEFT Then X2 = X1 - 1
                                        If Not IsValid(X2, Y2) Then
                                            X2 = X1
                                            Y2 = Y1
                                        End If
                                        'If Grid(Player(x).Pet.Map).Loc(X2, Y2).Blocked = True Then
                                        '    X2 = X1
                                       '     Y2 = Y1
                                       ' End If
                                    Else
                                        X2 = X1
                                        Y2 = Y1
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            If X1 < X2 Then

                ' RIGHT not left
                If Y1 < Y2 Then

                    ' DOWN not up
                    If X2 - X1 > Y2 - Y1 Then

                        ' RIGHT not down
                        If CanPetMove(X, DIR_RIGHT) Then

                            ' RIGHT works
                            Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                        Else

                            If CanPetMove(X, DIR_DOWN) Then

                                ' DOWN works and right doesn't
                                Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                            Else

                                ' Nothing works, random time
                                I = Int(Rnd * 4)

                                If CanPetMove(X, I) Then
                                    Call PetMove(X, I, MOVING_WALKING)
                                End If
                            End If
                        End If
                    Else

                        If X2 - X1 <> Y2 - Y1 Then

                            ' DOWN not right
                            If CanPetMove(X, DIR_DOWN) Then

                                ' DOWN works
                                Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                            Else

                                If CanPetMove(X, DIR_RIGHT) Then

                                    ' RIGHT works and down doesn't
                                    Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                Else

                                    ' Nothing works, random time
                                    I = Int(Rnd * 4)

                                    If CanPetMove(X, I) Then
                                        Call PetMove(X, I, MOVING_WALKING)
                                    End If
                                End If
                            End If
                        Else

                            ' Both are equal
                            If CanPetMove(X, DIR_RIGHT) Then

                                ' RIGHT works
                                If CanPetMove(X, DIR_DOWN) Then

                                    ' DOWN and RIGHT work
                                    I = (Int(Rnd * 2) * 2) + 1

                                    If CanPetMove(X, I) Then
                                        Call PetMove(X, I, MOVING_WALKING)
                                    End If
                                Else

                                    ' RIGHT works only
                                    Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                End If
                            Else

                                If CanPetMove(X, DIR_DOWN) Then

                                    ' DOWN works only
                                    Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                Else

                                    ' Nothing works, random time
                                    I = Int(Rnd * 4)

                                    If CanPetMove(X, I) Then
                                        Call PetMove(X, I, MOVING_WALKING)
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else

                    If Y1 <> Y2 Then

                        ' UP not down
                        If X2 - X1 > Y1 - Y2 Then

                            ' RIGHT not up
                            If CanPetMove(X, DIR_RIGHT) Then

                                ' RIGHT works
                                Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                            Else

                                If CanPetMove(X, DIR_UP) Then

                                    ' UP works and right doesn't
                                    Call PetMove(X, DIR_UP, MOVING_WALKING)
                                Else

                                    ' Nothing works, random time
                                    I = Int(Rnd * 4)

                                    If CanPetMove(X, I) Then
                                        Call PetMove(X, I, MOVING_WALKING)
                                    End If
                                End If
                            End If
                        Else

                            If X2 - X1 <> Y1 - Y2 Then

                                ' UP not right
                                If CanPetMove(X, DIR_UP) Then

                                    ' UP works
                                    Call PetMove(X, DIR_UP, MOVING_WALKING)
                                Else

                                    If CanPetMove(X, DIR_RIGHT) Then

                                        ' RIGHT works and up doesn't
                                        Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                    Else

                                        ' Nothing works, random time
                                        I = Int(Rnd * 4)

                                        If CanPetMove(X, I) Then
                                            Call PetMove(X, I, MOVING_WALKING)
                                        End If
                                    End If
                                End If
                            Else

                                ' Both are equal
                                If CanPetMove(X, DIR_RIGHT) Then

                                    ' RIGHT works
                                    If CanPetMove(X, DIR_UP) Then

                                        ' UP and RIGHT work
                                        I = Int(Rnd * 2) * 3

                                        If CanPetMove(X, I) Then
                                            Call PetMove(X, I, MOVING_WALKING)
                                        End If
                                    Else

                                        ' RIGHT works only
                                        Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                    End If
                                Else

                                    If CanPetMove(X, DIR_UP) Then

                                        ' UP works only
                                        Call PetMove(X, DIR_UP, MOVING_WALKING)
                                    Else

                                        ' Nothing works, random time
                                        I = Int(Rnd * 4)

                                        If CanPetMove(X, I) Then
                                            Call PetMove(X, I, MOVING_WALKING)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else

                        ' Target is horizontal
                        If CanPetMove(X, DIR_RIGHT) Then

                            ' RIGHT works
                            Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                        Else

                            ' Right doesn't work
                            If CanPetMove(X, DIR_UP) Then
                                If CanPetMove(X, DIR_DOWN) Then

                                    ' UP and DOWN work
                                    I = Int(Rnd * 2)
                                    Call PetMove(X, I, MOVING_WALKING)
                                Else

                                    ' Only UP works
                                    Call PetMove(X, DIR_UP, MOVING_WALKING)
                                End If
                            Else

                                If CanPetMove(X, DIR_DOWN) Then

                                    ' Only DOWN works
                                    Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                Else

                                    ' Nothing works, only left is left (heh)
                                    If CanPetMove(X, DIR_LEFT) Then
                                        Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                    Else

                                        ' Nothing works at all, let it die
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Else

                If X1 <> X2 Then

                    ' LEFT not right
                    If Y1 < Y2 Then

                        ' DOWN not up
                        If X1 - X2 > Y2 - Y1 Then

                            ' LEFT not down
                            If CanPetMove(X, DIR_LEFT) Then

                                ' LEFT works
                                Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                            Else

                                If CanPetMove(X, DIR_DOWN) Then

                                    ' DOWN works and left doesn't
                                    Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                Else

                                    ' Nothing works, random time
                                    I = Int(Rnd * 4)

                                    If CanPetMove(X, I) Then
                                        Call PetMove(X, I, MOVING_WALKING)
                                    End If
                                End If
                            End If
                        Else

                            If X1 - X2 <> Y2 - Y1 Then

                                ' DOWN not left
                                If CanPetMove(X, DIR_DOWN) Then

                                    ' DOWN works
                                    Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                Else

                                    If CanPetMove(X, DIR_LEFT) Then

                                        ' LEFT works and down doesn't
                                        Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                    Else

                                        ' Nothing works, random time
                                        I = Int(Rnd * 4)

                                        If CanPetMove(X, I) Then
                                            Call PetMove(X, I, MOVING_WALKING)
                                        End If
                                    End If
                                End If
                            Else

                                ' Both are equal
                                If CanPetMove(X, DIR_LEFT) Then

                                    ' LEFT works
                                    If CanPetMove(X, DIR_DOWN) Then

                                        ' DOWN and LEFT work
                                        I = Int(Rnd * 2) + 1
                                        Call PetMove(X, I, MOVING_WALKING)
                                    Else

                                        ' LEFT works only
                                        Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                    End If
                                Else

                                    If CanPetMove(X, DIR_DOWN) Then

                                        ' DOWN works only
                                        Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                    Else

                                        ' Nothing works, random time
                                        I = Int(Rnd * 4)

                                        If CanPetMove(X, I) Then
                                            Call PetMove(X, I, MOVING_WALKING)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Else

                        If Y1 <> Y2 Then

                            ' UP not down
                            If X1 - X2 > Y1 - Y2 Then

                                ' LEFT not up
                                If CanPetMove(X, DIR_LEFT) Then

                                    ' LEFT works
                                    Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                Else

                                    If CanPetMove(X, DIR_UP) Then

                                        ' UP works and left doesn't
                                        Call PetMove(X, DIR_UP, MOVING_WALKING)
                                    Else

                                        ' Nothing works, random time
                                        I = Int(Rnd * 4)

                                        If CanPetMove(X, I) Then
                                            Call PetMove(X, I, MOVING_WALKING)
                                        End If
                                    End If
                                End If
                            Else

                                If X1 - X2 <> Y1 - Y2 Then

                                    ' UP not LEFT
                                    If CanPetMove(X, DIR_UP) Then

                                        ' UP works
                                        Call PetMove(X, DIR_UP, MOVING_WALKING)
                                    Else

                                        If CanPetMove(X, DIR_LEFT) Then

                                            ' LEFT works and up doesn't
                                            Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                        Else

                                            ' Nothing works, random time
                                            I = Int(Rnd * 4)

                                            If CanPetMove(X, I) Then
                                                Call PetMove(X, I, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                Else

                                    ' Both are equal
                                    If CanPetMove(X, DIR_LEFT) Then

                                        ' LEFT works
                                        If CanPetMove(X, DIR_UP) Then

                                            ' UP and LEFT work
                                            I = Int(Rnd * 2) * 2
                                            Call PetMove(X, I, MOVING_WALKING)
                                        Else

                                            ' LEFT works only
                                            Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                        End If
                                    Else

                                        If CanPetMove(X, DIR_UP) Then

                                            ' UP works only
                                            Call PetMove(X, DIR_UP, MOVING_WALKING)
                                        Else

                                            ' Nothing works, random time
                                            I = Int(Rnd * 4)

                                            If CanPetMove(X, I) Then
                                                Call PetMove(X, I, MOVING_WALKING)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        Else

                            ' Target is horizontal
                            If CanPetMove(X, DIR_LEFT) Then

                                ' LEFT works
                                Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                            Else

                                ' LEFT doesn't work
                                If CanPetMove(X, DIR_UP) Then
                                    If CanPetMove(X, DIR_DOWN) Then

                                        ' UP and DOWN work
                                        I = Int(Rnd * 2)
                                        Call PetMove(X, I, MOVING_WALKING)
                                    Else

                                        ' Only UP works
                                        Call PetMove(X, DIR_UP, MOVING_WALKING)
                                    End If
                                Else

                                    If CanPetMove(X, DIR_DOWN) Then

                                        ' Only DOWN works
                                        Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                    Else

                                        ' Nothing works, only right is left (heh)
                                        If CanPetMove(X, DIR_RIGHT) Then
                                            Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                        Else

                                            ' Nothing works at all, let it die
                                            Player(X).Pet.MapToGo = Player(X).Pet.Map
                                            Player(X).Pet.XToGo = -1
                                            Player(X).Pet.YToGo = -1
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else

                    ' Target is vertical
                    If Y1 < Y2 Then

                        ' DOWN not up
                        If CanPetMove(X, DIR_DOWN) Then
                            Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                        Else

                            ' Down doesn't work
                            If CanPetMove(X, DIR_RIGHT) Then
                                If CanPetMove(X, DIR_LEFT) Then

                                    ' RIGHT and LEFT work
                                    I = Int((Rnd * 2) + 2)
                                    Call PetMove(X, I, MOVING_WALKING)
                                Else

                                    ' RIGHT works only
                                    Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                End If
                            Else

                                If CanPetMove(X, DIR_LEFT) Then

                                    ' LEFT works only
                                    Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                Else

                                    ' Nothing works, lets try up
                                    If CanPetMove(X, DIR_UP) Then
                                        Call PetMove(X, DIR_UP, MOVING_WALKING)
                                    Else

                                        ' Nothing at all works, let it die
                                        Player(X).Pet.MapToGo = Player(X).Pet.Map
                                        Player(X).Pet.XToGo = -1
                                        Player(X).Pet.YToGo = -1
                                    End If
                                End If
                            End If
                        End If
                    Else

                        If Y1 <> Y2 Then

                            ' UP not down
                            If CanPetMove(X, DIR_UP) Then
                                Call PetMove(X, DIR_UP, MOVING_WALKING)
                            Else

                                ' UP doesn't work
                                If CanPetMove(X, DIR_RIGHT) Then
                                    If CanPetMove(X, DIR_LEFT) Then

                                        ' RIGHT and LEFT work
                                        I = Int((Rnd * 2) + 2)
                                        Call PetMove(X, I, MOVING_WALKING)
                                    Else

                                        ' RIGHT works only
                                        Call PetMove(X, DIR_RIGHT, MOVING_WALKING)
                                    End If
                                Else

                                    If CanPetMove(X, DIR_LEFT) Then

                                        ' LEFT works only
                                        Call PetMove(X, DIR_LEFT, MOVING_WALKING)
                                    Else

                                        ' Nothing works, lets try down
                                        If CanPetMove(X, DIR_DOWN) Then
                                            Call PetMove(X, DIR_DOWN, MOVING_WALKING)
                                        Else

                                            ' Nothing at all works, let it die
                                            Player(X).Pet.MapToGo = Player(X).Pet.Map
                                            Player(X).Pet.XToGo = -1
                                            Player(X).Pet.YToGo = -1
                                        End If
                                    End If
                                End If
                            End If
                        Else
                            Player(X).Pet.MapToGo = Player(X).Pet.Map
                            Player(X).Pet.XToGo = -1
                            Player(X).Pet.YToGo = -1
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub


Sub ScriptedTimer()
    Dim X As Long, N As Long
    Dim CustomTimer As clsCTimers

    N = 0
    X = CTimers.Count
    For Each CustomTimer In CTimers
        N = N + 1
        If GetTickCount > CustomTimer.tmrWait Then
            MyScript.ExecuteStatement "Scripts\Main.txt", CustomTimer.Name ' & " " & Index & "," & PointType
            If CTimers.Count < X Then
                N = N - X - CTimers.Count
                X = CTimers.Count
            End If
            If N > 0 Then
                CTimers.Item(N).tmrWait = GetTickCount + CustomTimer.Interval
            Else
                Exit For
            End If
        End If
    Next CustomTimer
End Sub

Sub CheckGiveVitals()
    Dim I As Long

    If HP_REGEN = 1 Then
        If GetTickCount >= GiveHPTimer + HP_TIMER Then
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) Then
                    If GetPlayerHP(I) < GetPlayerMaxHP(I) Then
                        Call SetPlayerHP(I, GetPlayerHP(I) + GetPlayerHPRegen(I))
                        Call SendHP(I)
                    End If
                End If
            Next I

            GiveHPTimer = GetTickCount
        End If
    End If

    If MP_REGEN = 1 Then
        If GetTickCount >= GiveMPTimer + MP_TIMER Then
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) Then
                    If GetPlayerMP(I) < GetPlayerMaxMP(I) Then
                        Call SetPlayerMP(I, GetPlayerMP(I) + GetPlayerMPRegen(I))
                        Call SendMP(I)
                    End If
                End If
            Next I

            GiveMPTimer = GetTickCount
        End If
    End If

    If SP_REGEN = 1 Then
        If GetTickCount >= GiveSPTimer + SP_TIMER Then
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) Then
                    If GetPlayerSP(I) < GetPlayerMaxSP(I) Then
                        Call SetPlayerSP(I, GetPlayerSP(I) + GetPlayerSPRegen(I))
                        Call SendSP(I)
                    End If
                End If
            Next I

            GiveSPTimer = GetTickCount
        End If
    End If
End Sub

Sub PlayerSaveTimer()
    Dim I As Long

    PLYRSAVE_TIMER = PLYRSAVE_TIMER + 1

    If SAVETIME <> 0 Then
        If PLYRSAVE_TIMER >= SAVETIME Then
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) Then
                    Call SavePlayer(I)
                End If
            Next I
    
            PlayerI = 1

            frmServer.PlayerTimer.Enabled = True
            frmServer.tmrPlayerSave.Enabled = False

            PLYRSAVE_TIMER = 0
        End If
    Else
        PLYRSAVE_TIMER = 0
    End If
End Sub

Function IsAlphaNumeric(TestString As String) As Boolean
    Dim LoopID As Integer
    Dim sChar As String

    IsAlphaNumeric = False

    If LenB(TestString) > 0 Then
        For LoopID = 1 To Len(TestString)
            sChar = Mid(TestString, LoopID, 1)
            If Not sChar Like "[0-9A-Za-zÒ—·¡È…ÌÕÛ”˙⁄ ]" Then
                Exit Function
            End If
        Next

        IsAlphaNumeric = True
    End If
End Function

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function
