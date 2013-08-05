Attribute VB_Name = "ModPetSystem"
Option Explicit

'Written by:Morpheus
Public Function GetPetHp(ByVal index As Long) As Long
    GetPetHp = Player(index).Pet.HP
End Function

Public Sub SetPetHp(ByVal index As Long, ByVal HP As Long)
    Player(index).Pet.HP = HP
     
End Sub

Public Function GetPetLevel(ByVal index As Long) As Long
    GetPetLevel = Player(index).Pet.Level
End Function

Public Sub SetPetLevel(ByVal index As Long, ByVal Level As Long)
    Player(index).Pet.Level = Level
     
End Sub

Public Function GetPetSTR(ByVal index As Long) As Long
    GetPetSTR = Player(index).Pet.STR

End Function

Public Sub SetPetSTR(ByVal index As Long, ByVal STR As Long)
    Player(index).Pet.STR = STR
     
End Sub

Public Function GetPetDEF(ByVal index As Long) As Long
    GetPetDEF = Player(index).Pet.DEF

End Function

Public Function GetPetSPEED(ByVal index As Long) As Long
    GetPetSPEED = Player(index).Pet.Speed

End Function

Public Sub SetPetSPEED(ByVal index As Long, ByVal Speed As Long)
    Player(index).Pet.Speed = Speed
     
End Sub

Public Function GetPetMAGI(ByVal index As Long) As Long
    GetPetMAGI = Player(index).Pet.Magi
     
End Function

Public Sub SetPetMAGI(ByVal index As Long, ByVal Magi As Long)
    Player(index).Pet.Magi = Magi
     
End Sub

Public Sub SetPetDEF(ByVal index As Long, ByVal DEF As Long)
    Player(index).Pet.DEF = DEF
     
End Sub

Public Function GetPetX(ByVal index As Long) As Long
GetPetX = Player(index).Pet.x
End Function

Public Function GetPetY(ByVal index As Long) As Long
GetPetY = Player(index).Pet.y
End Function

Public Function SetPetToGo(ByVal index As Long, ByVal x As Long, ByVal y As Long)
Player(index).Pet.XToGo = x
Player(index).Pet.YToGo = y
 
End Function

Public Function IsPetAliveOnLogin(ByVal index As Long) As Long
IsPetAliveOnLogin = Player(index).Char(Player(index).CharNum).PetAlive
End Function

Public Sub SetPlayerPetSprite(ByVal index As Long, ByVal PetSprite As Long)
    Player(index).Char(Player(index).CharNum).PetSprite = PetSprite
     
End Sub

Public Function GetPlayerPetSprite(ByVal index As Long) As Long
    GetPlayerPetSprite = Player(index).Char(Player(index).CharNum).PetSprite
End Function

Public Sub SetPlayerPetAlive(ByVal index As Long, ByVal PetAlive As Long)
    Player(index).Char(Player(index).CharNum).PetAlive = PetAlive
     
End Sub

Public Function GetPlayerPetAlive(ByVal index As Long) As Long
    GetPlayerPetAlive = Player(index).Char(Player(index).CharNum).PetAlive
End Function

Public Sub SetPlayerPetMap(ByVal index As Long, ByVal PetMap As Long)
    Player(index).Char(Player(index).CharNum).PetMap = PetMap
     
End Sub

Public Function GetPlayerPetMap(ByVal index As Long) As Long
    GetPlayerPetMap = Player(index).Char(Player(index).CharNum).PetMap
End Function

Public Sub SetPlayerPetX(ByVal index As Long, ByVal PetX As Long)
    Player(index).Char(Player(index).CharNum).PetX = PetX
     
End Sub

Public Function GetPlayerPetX(ByVal index As Long) As Long
    GetPlayerPetX = Player(index).Char(Player(index).CharNum).PetX
End Function

Public Sub SetPlayerPetY(ByVal index As Long, ByVal PetY As Long)
    Player(index).Char(Player(index).CharNum).PetY = PetY
     
End Sub

Public Function GetPlayerPetY(ByVal index As Long) As Long
    GetPlayerPetY = Player(index).Char(Player(index).CharNum).PetY
End Function

Public Sub SetPlayerPetDIR(ByVal index As Long, ByVal PetDIR As Long)
    Player(index).Char(Player(index).CharNum).PetDIR = PetDIR
     
End Sub

Public Function GetPlayerPetDIR(ByVal index As Long) As Long
    GetPlayerPetDIR = Player(index).Char(Player(index).CharNum).PetDIR
End Function

Public Sub SetPlayerPetHP(ByVal index As Long, ByVal PetHP As Long)
    Player(index).Char(Player(index).CharNum).PetHP = PetHP
     
End Sub

Public Function GetPlayerPetHP(ByVal index As Long) As Long
    GetPlayerPetHP = Player(index).Char(Player(index).CharNum).PetHP
End Function

Public Sub SetPlayerPetSP(ByVal index As Long, ByVal PetSP As Long)
    Player(index).Char(Player(index).CharNum).PetSP = PetSP
     
End Sub

Public Function GetPlayerPetSP(ByVal index As Long) As Long
    GetPlayerPetSP = Player(index).Char(Player(index).CharNum).PetSP
End Function

Public Sub SetPlayerPetMP(ByVal index As Long, ByVal PetMP As Long)
    Player(index).Char(Player(index).CharNum).PetMP = PetMP
     
End Sub

Public Function GetPlayerPetMP(ByVal index As Long) As Long
    GetPlayerPetMP = Player(index).Char(Player(index).CharNum).PetMP
End Function

Public Sub SetPlayerPetFP(ByVal index As Long, ByVal PetFP As Long)
    Player(index).Char(Player(index).CharNum).PetFP = PetFP
     
End Sub

Public Function GetPlayerPetFP(ByVal index As Long) As Long
    GetPlayerPetFP = Player(index).Char(Player(index).CharNum).PetFP
End Function

Public Sub SetPlayerPetMaxHP(ByVal index As Long, ByVal PetMaxHP As Long)
    Player(index).Char(Player(index).CharNum).PetMaxHP = PetMaxHP
     
End Sub

Public Function GetPlayerPetMaxHP(ByVal index As Long) As Long
    GetPlayerPetMaxHP = Player(index).Char(Player(index).CharNum).PetMaxHP
End Function

Public Sub SetPlayerPetMaxSP(ByVal index As Long, ByVal PetMaxSP As Long)
    Player(index).Char(Player(index).CharNum).PetMaxSP = PetMaxSP
     
End Sub

Public Function GetPlayerPetMaxSP(ByVal index As Long) As Long
    GetPlayerPetMaxSP = Player(index).Char(Player(index).CharNum).PetMaxSP
End Function

Public Sub SetPlayerPetMaxMP(ByVal index As Long, ByVal PetMaxMP As Long)
    Player(index).Char(Player(index).CharNum).PetMaxMP = PetMaxMP
     
End Sub

Public Function GetPlayerPetMaxMP(ByVal index As Long) As Long
    GetPlayerPetMaxMP = Player(index).Char(Player(index).CharNum).PetMaxMP
End Function

Public Sub SetPlayerPetMaxFP(ByVal index As Long, ByVal PetMaxFP As Long)
    Player(index).Char(Player(index).CharNum).PetMaxFP = PetMaxFP
End Sub

Public Function GetPlayerPetMaxFP(ByVal index As Long) As Long
    GetPlayerPetMaxFP = Player(index).Char(Player(index).CharNum).PetMaxFP
End Function

Public Sub SetPlayerPetLevel(ByVal index As Long, ByVal PetLevel As Long)
    Player(index).Char(Player(index).CharNum).PetLevel = PetLevel
End Sub

Public Function GetPlayerPetLevel(ByVal index As Long) As Long
    GetPlayerPetLevel = Player(index).Char(Player(index).CharNum).PetLevel
End Function

Public Sub SetPlayerPetSTR(ByVal index As Long, ByVal PetSTR As Long)
    Player(index).Char(Player(index).CharNum).PetSTR = PetSTR
     
End Sub

Public Function GetPlayerPetSTR(ByVal index As Long) As Long
    GetPlayerPetSTR = Player(index).Char(Player(index).CharNum).PetSTR
End Function

Public Sub SetPlayerPetDEF(ByVal index As Long, ByVal PetDEF As Long)
    Player(index).Char(Player(index).CharNum).PetDEF = PetDEF
     
End Sub

Public Function GetPlayerPetDEF(ByVal index As Long) As Long
    GetPlayerPetDEF = Player(index).Char(Player(index).CharNum).PetDEF
End Function

Public Sub SetPlayerPetMAGI(ByVal index As Long, ByVal PetMAGI As Long)
    Player(index).Char(Player(index).CharNum).PetMAGI = PetMAGI
     
End Sub

Public Function GetPlayerPetMAGI(ByVal index As Long) As Long
    GetPlayerPetMAGI = Player(index).Char(Player(index).CharNum).PetMAGI
End Function

Public Sub SetPlayerPetSPEED(ByVal index As Long, ByVal PetSPEED As Long)
    Player(index).Char(Player(index).CharNum).PetSPEED = PetSPEED
     
End Sub

Public Function GetPlayerPetSPEED(ByVal index As Long) As Long
    GetPlayerPetSPEED = Player(index).Char(Player(index).CharNum).PetSPEED
End Function

Public Sub SetPlayerPetEXP(ByVal index As Long, ByVal PetEXP As Long)
    Player(index).Char(Player(index).CharNum).PetEXP = PetEXP
     
End Sub

Public Function GetPlayerPetEXP(ByVal index As Long) As Long
    GetPlayerPetEXP = Player(index).Char(Player(index).CharNum).PetEXP
End Function

Public Sub SetPlayerPetPOINTS(ByVal index As Long, ByVal PetPOINTS As Long)
    Player(index).Char(Player(index).CharNum).PetPOINTS = PetPOINTS
     
End Sub

Public Function GetPlayerPetPOINTS(ByVal index As Long) As Long
    GetPlayerPetPOINTS = Player(index).Char(Player(index).CharNum).PetPOINTS
End Function

Public Sub SetPlayerPetNAME(ByVal index As Long, ByVal PetNAME As String)
    Player(index).Char(Player(index).CharNum).PetNAME = PetNAME
     
End Sub

Public Function GetPlayerPetNAME(ByVal index As Long) As String
    GetPlayerPetNAME = Player(index).Char(Player(index).CharNum).PetNAME
End Function

Public Sub SetPlayerPetTNL(ByVal index As Long, ByVal PetTNL As Long)
    Player(index).Char(Player(index).CharNum).PetTNL = PetTNL
     
End Sub

Public Function GetPlayerPetTNL(ByVal index As Long) As Long
    GetPlayerPetTNL = Player(index).Char(Player(index).CharNum).PetTNL
End Function

' Old Functions

Public Function GetPetPOINTS(ByVal index As Long) As Long
    GetPetPOINTS = Player(index).Pet.POINTS
End Function

Public Sub SetPetPOINTS(ByVal index As Long, ByVal POINTS As Long)
    Player(index).Pet.POINTS = POINTS
     
End Sub

Public Function GetPetMAXHP(ByVal index As Long) As Long
    GetPetMAXHP = Player(index).Pet.STR * 5
End Function

Public Function GetPetMAXSP(ByVal index As Long) As Long
    GetPetMAXSP = Player(index).Pet.Speed * 5
End Function

Public Function GetPetMAXMP(ByVal index As Long) As Long
    GetPetMAXMP = Player(index).Pet.Magi * 5
End Function

Public Function GetPetMAXFP(ByVal index As Long) As Long
    GetPetMAXFP = 100
End Function

Public Function GetPetSP(ByVal index As Long) As Long
    GetPetSP = Player(index).Pet.SP
End Function

Public Sub SetPetSP(ByVal index As Long, ByVal SP As Long)
    Player(index).Pet.SP = SP
     
End Sub

Public Function GetPetMP(ByVal index As Long) As Long
    GetPetMP = Player(index).Pet.MP
End Function

Public Sub SetPetMP(ByVal index As Long, ByVal MP As Long)
    Player(index).Pet.MP = MP
     
End Sub

Public Function GetPetFP(ByVal index As Long) As Long
    GetPetFP = Player(index).Pet.FP
End Function

Public Sub SetPetFP(ByVal index As Long, ByVal FP As Long)
    Player(index).Pet.FP = FP
    
End Sub

Public Function GetPetExp(ByVal index As Long) As Long
    GetPetExp = Player(index).Pet.EXP
End Function

Public Sub SetPetExp(ByVal index As Long, ByVal EXP As Long)
    Player(index).Pet.EXP = EXP
    
End Sub

Public Function GetPetNextLevel(ByVal index As Long) As Long
If Player(index).Pet.Alive = NO Then Exit Function
    GetPetNextLevel = Experience(GetPetLevel(index))
End Function




'---------------

Function CanPetAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
Dim mapnum As Long
Dim i As Long
Dim x As Long
Dim y As Long
Dim Dir As Long

CanPetAttackPlayer = False

' Check for subscript out of range
If IsPlaying(Attacker) = False Or Victim <= 0 Then
Exit Function
End If

' Make sure we arent trying to kill ourselves
If Attacker = Victim Then
Exit Function
End If

mapnum = Player(Attacker).Pet.Map
i = Victim

' Make sure the player isn't already dead
If Player(i).Char(Player(i).CharNum).HP <= 0 Then
Exit Function
End If

' Make sure they are on the same map
If IsPlaying(Attacker) Then
If i > 0 And GetPlayerMap(Victim) = mapnum And GetTickCount > Player(Attacker).Pet.AttackTimer + 1000 Then
'Cuchu pwns!
If Player(i).Char(Player(i).CharNum).Access < 4 Then

For Dir = 0 To 3

' Check if at same coordinates
x = DirToX(Player(Attacker).Pet.x, Dir)
y = DirToY(Player(Attacker).Pet.y, Dir)

If (Player(i).Char(Player(i).CharNum).y = y) And (Player(i).Char(Player(i).CharNum).x = x) Then
CanPetAttackPlayer = True
End If

Next

End If
End If
End If

End Function

Sub PetAttackPlayer(ByVal Attacker As Long, _
ByVal Victim As Long, _
ByVal Damage As Long)
Dim Name As String
Dim N As Long, i As Long
Dim Dir As Long, x As Long, y As Long
Dim Packet As String
Dim mapnum As Long
Dim Char As Byte

' Check for subscript out of range

End Sub

Sub ChoosePet(ByVal index As Long, ByVal SPRITE As Long, ByVal Name As String)
Dim Packet As String
'Player(index).Pet.SPRITE = SPRITE
Player(index).Pet.Name = Trim(Name)
Call SetPlayerPetSprite(index, SPRITE)
Call SetPlayerPetNAME(index, Trim(Name))
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
Call SendDataToMap(GetPlayerMap(index), Packet)
Call PlayerMsg(index, "[Mascota] Has elegido tu mascota!", WHITE)
Call savepet(index)
End Sub

Sub savepet(ByVal index As Long)
Call SetPlayerPetAlive(index, 1)
Call SetPlayerPetSprite(index, Player(index).Pet.SPRITE)
Call SetPlayerPetHP(index, Player(index).Pet.HP)
Call SetPlayerPetSP(index, Player(index).Pet.SP)
Call SetPlayerPetMP(index, Player(index).Pet.MP)
Call SetPlayerPetFP(index, Player(index).Pet.FP)
Call SetPlayerPetSTR(index, Player(index).Pet.STR)
Call SetPlayerPetDEF(index, Player(index).Pet.DEF)
Call SetPlayerPetMAGI(index, Player(index).Pet.Magi)
Call SetPlayerPetSPEED(index, Player(index).Pet.Speed)
Call SetPlayerPetLevel(index, Player(index).Pet.Level)
Call SetPlayerPetEXP(index, Player(index).Pet.EXP)
Call SetPlayerPetPOINTS(index, Player(index).Pet.POINTS)
Call SetPlayerPetNAME(index, Player(index).Pet.Name)
Call SetPlayerPetTNL(index, GetPetNextLevel(index))
Call PlayerMsg(index, "[Mascota] Tu mascota se guardo!", BRIGHTGREEN)
End Sub

Sub savepetDeath(ByVal index As Long)
Call SetPlayerPetAlive(index, 0)
Call SetPlayerPetSprite(index, Player(index).Pet.SPRITE)
Call SetPlayerPetHP(index, Player(index).Pet.HP)
Call SetPlayerPetSP(index, Player(index).Pet.SP)
Call SetPlayerPetMP(index, Player(index).Pet.MP)
Call SetPlayerPetFP(index, Player(index).Pet.FP)
Call SetPlayerPetSTR(index, Player(index).Pet.STR)
Call SetPlayerPetDEF(index, Player(index).Pet.DEF)
Call SetPlayerPetMAGI(index, Player(index).Pet.Magi)
Call SetPlayerPetSPEED(index, Player(index).Pet.Speed)
Call SetPlayerPetLevel(index, Player(index).Pet.Level)
Call SetPlayerPetEXP(index, Player(index).Pet.EXP)
Call SetPlayerPetPOINTS(index, Player(index).Pet.POINTS)
Call SetPlayerPetNAME(index, Player(index).Pet.Name)
Call SetPlayerPetTNL(index, GetPetNextLevel(index))
Call SavePlayer(index)
End Sub

Sub SpawnPet(ByVal index As Long)
Dim Packet As String
Dim STR As Long
Dim DEF As Long
Dim Magi As Long
Dim Speed As Long
Dim EXP As Long
Dim Level As Long
Dim HP As Long
Dim SP As Long
Dim MP As Long
Dim FP As Long
Dim TNL As Long
Dim POINTS As Long
Dim SPRITE As Long
Dim PetNAME As String
STR = GetPlayerPetSTR(index) 'GetVar(App.Path & "\Main\Pets\" & GetPlayerName(Index) & ".ini", "PET", "STR")
DEF = GetPlayerPetDEF(index) 'GetVar(App.Path & "\Main\Pets\" & GetPlayerName(Index) & ".ini", "PET", "DEF")
Speed = GetPlayerPetSPEED(index) 'GetVar(App.Path & "\Main\Pets\" & GetPlayerName(Index) & ".ini", "PET", "SPEED")
Magi = GetPlayerPetMAGI(index) 'GetVar(App.Path & "\Main\Pets\" & GetPlayerName(Index) & ".ini", "PET", "MAGI")
EXP = GetPlayerPetEXP(index) 'GetVar(App.Path & "\Main\Pets\" & GetPlayerName(Index) & ".ini", "PET", "EXP")
Level = GetPlayerPetLevel(index) 'GetVar(App.Path & "\Main\Pets\" & GetPlayerName(Index) & ".ini", "PET", "LEVEL")
HP = GetPlayerPetHP(index) 'GetVar(App.Path & "\Main\Pets\" & GetPlayerName(Index) & ".ini", "PET", "HP")
SP = GetPlayerPetSP(index) 'GetVar(App.Path & "\Main\Pets\" & GetPlayerName(Index) & ".ini", "PET", "SP")
MP = GetPlayerPetMP(index) 'GetVar(App.Path & "\Main\Pets\" & GetPlayerName(Index) & ".ini", "PET", "MP")
FP = GetPlayerPetFP(index) 'GetVar(App.Path & "\Main\Pets\" & GetPlayerName(Index) & ".ini", "PET", "FP")
TNL = GetPlayerPetTNL(index) 'GetVar(App.Path & "\Main\Pets\" & GetPlayerName(Index) & ".ini", "PET", "TNL")
POINTS = GetPlayerPetPOINTS(index) 'GetVar(App.Path & "\Main\Pets\" & GetPlayerName(Index) & ".ini", "PET", "POINTS")
SPRITE = GetPlayerPetSprite(index) 'GetVar(App.Path & "\Main\Pets\" & GetPlayerName(Index) & ".ini", "PET", "SPRITE")
PetNAME = GetPlayerPetNAME(index) 'GetVar(App.Path & "\Main\Pets\" & GetPlayerName(Index) & ".ini", "PET", "NAME")
Player(index).Pet.Alive = YES
Player(index).Pet.Dir = DIR_DOWN
Player(index).Pet.Map = GetPlayerMap(index)
Player(index).Pet.MapToGo = 0
Player(index).Pet.SPRITE = SPRITE
Player(index).Pet.SpriteSet = 0
Player(index).Pet.Name = PetNAME
Call SetPetSTR(index, STR)
Call SetPetDEF(index, DEF)
Call SetPetSPEED(index, Speed)
Call SetPetMAGI(index, Magi)
Call SetPetPOINTS(index, POINTS)
Call SetPetExp(index, EXP)
Call SetPetLevel(index, Level)
Player(index).Pet.x = GetPlayerX(index) + Int(Rnd * 3 - 1)
If Player(index).Pet.x < 0 Or Player(index).Pet.x > MAX_MAPX Then Player(index).Pet.x = GetPlayerX(index)
Player(index).Pet.XToGo = -1
Player(index).Pet.y = GetPlayerY(index) + Int(Rnd * 3 - 1)
If Player(index).Pet.y < 0 Or Player(index).Pet.y > MAX_MAPY Then Player(index).Pet.y = GetPlayerY(index)
Player(index).Pet.YToGo = -1
Call SetPetHp(index, HP)
Call SetPetSP(index, SP)
Call SetPetMP(index, MP)
Call SetPetFP(index, FP)
'Call AddToGrid(Player(Index).Pet.Map, Player(Index).Pet.x, Player(Index).Pet.y)
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
Call SendDataToMap(GetPlayerMap(index), Packet)
Call PlayerMsg(index, "[Mascota] Tu mascota espera tus ordenes!", WHITE)
End Sub

Sub ClearPet(ByVal index As Long)
Call SetPlayerPetAlive(index, 0)
Call SetPlayerPetSprite(index, 0)
Call SetPlayerPetHP(index, 0)
Call SetPlayerPetMP(index, 0)
Call SetPlayerPetSP(index, 0)
Call SetPlayerPetFP(index, 0)
Call SetPlayerPetSTR(index, 0)
Call SetPlayerPetDEF(index, 0)
Call SetPlayerPetSPEED(index, 0)
Call SetPlayerPetMAGI(index, 0)
Call SetPlayerPetLevel(index, 1)
Call SetPlayerPetEXP(index, 0)
Call SetPlayerPetNAME(index, "")
Call SetPlayerPetTNL(index, 0)
End Sub

Sub RezPet(ByVal index As Long)
Dim chance As Long
If IsPetAliveOnLogin(index) > 0 Then
Call PlayerMsg(index, "Ya tienes una mascota!", BRIGHTRED)
Exit Sub
End If
Call SetPlayerPetAlive(index, 1)
Call SpawnPet(index)
Call PlayerMsg(index, "[Mascota] Has resucitado a tu mascota!", BRIGHTGREEN)
Call SendPlayerData(index)
End Sub

Function CanNpcAttackPet(ByVal MapNpcNum As Long, ByVal index As Long) As Boolean
Dim mapnum As Long, npcnum As Long
Dim x As Long
Dim y As Long

    CanNpcAttackPet = False

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(index) = False Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNPC(GetPlayerMap(index), MapNpcNum).num <= 0 Then
        Exit Function
    End If
    mapnum = Player(index).Pet.Map
    npcnum = MapNPC(mapnum, MapNpcNum).num

    ' Make sure the npc isn't already dead
    If MapNPC(mapnum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Make sure npcs dont attack more then once a second
    If GetTickCount < MapNPC(mapnum, MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If
    MapNPC(mapnum, MapNpcNum).AttackTimer = GetTickCount

    ' Make sure they are on the same map
    If IsPlaying(index) Then
        If npcnum > 0 Then
            x = DirToX(MapNPC(mapnum, MapNpcNum).x, MapNPC(mapnum, MapNpcNum).Dir)
            y = DirToY(MapNPC(mapnum, MapNpcNum).y, MapNPC(mapnum, MapNpcNum).Dir)

            ' Check if at same coordinates
            If (Player(index).Pet.y = y) And (Player(index).Pet.x = x) Then
                CanNpcAttackPet = True
            End If
        End If
    End If
End Function

Function CanPetAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
Dim mapnum As Long, npcnum As Long
Dim x As Long
Dim y As Long
Dim Dir As Long

    CanPetAttackNpc = False

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for subscript out of range
    If MapNPC(Player(Attacker).Pet.Map, MapNpcNum).num <= 0 Then
        Exit Function
    End If
    mapnum = Player(Attacker).Pet.Map
    npcnum = MapNPC(mapnum, MapNpcNum).num

    ' Make sure the npc isn't already dead
    If MapNPC(mapnum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Make sure they are on the same map
    If IsPlaying(Attacker) Then
        If npcnum > 0 And GetTickCount > Player(Attacker).Pet.AttackTimer + 1000 Then
            If NPC(npcnum).Behavior <> Int(NPC_BEHAVIOR_FRIENDLY) And NPC(npcnum).Behavior <> Int(NPC_BEHAVIOR_SHOPKEEPER) Then
                For Dir = 0 To 3

                    ' Check if at same coordinates
                    x = DirToX(Player(Attacker).Pet.x, Dir)
                    y = DirToY(Player(Attacker).Pet.y, Dir)

                    If (MapNPC(mapnum, MapNpcNum).y = y) And (MapNPC(mapnum, MapNpcNum).x = x) Then
                        CanPetAttackNpc = True
                    End If
                Next
            End If
        End If
    End If
End Function

Function CanPetMove(ByVal PetNum As Long, ByVal Dir) As Boolean
Dim x As Long, y As Long
Dim i As Long, Packet As String
Dim mapnum As Long
Dim TileType
Dim PetX
Dim PetY

    CanPetMove = True

    If PetNum <= 0 Or PetNum > MAX_PLAYERS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then Exit Function
    x = DirToX(Player(PetNum).Pet.x, Dir)
    y = DirToY(Player(PetNum).Pet.y, Dir)
    mapnum = Player(PetNum).Pet.Map
    PetX = Player(PetNum).Pet.x
    PetY = Player(PetNum).Pet.y
    'Call GlobalMsg(PetX & "," & PetY, 4)
    If IsValid(x, y) Then
        If Dir = DIR_UP Then
                
        
            If Map(Player(PetNum).Pet.Map).Up > 0 And Map(Player(PetNum).Pet.Map).Up = Player(PetNum).Pet.MapToGo Then
                CanPetMove = True
            End If
            
            
            TileType = Map(mapnum).Tile(PetX, PetY - 1).Type
            If TileType = TILE_TYPE_BLOCKED Then
                CanPetMove = False
            End If
            
        
        End If

        If Dir = DIR_DOWN Then
            If Map(Player(PetNum).Pet.Map).Down > 0 And Map(Player(PetNum).Pet.Map).Down = Player(PetNum).Pet.MapToGo Then
                CanPetMove = True
            End If
            
            
            TileType = Map(mapnum).Tile(PetX, PetY + 1).Type
            If TileType = TILE_TYPE_BLOCKED Then
                CanPetMove = False
            End If
            
        End If

        If Dir = DIR_LEFT Then
            If Map(Player(PetNum).Pet.Map).Left > 0 And Map(Player(PetNum).Pet.Map).Left = Player(PetNum).Pet.MapToGo Then
                CanPetMove = True
            End If
            
            TileType = Map(mapnum).Tile(PetX - 1, PetY).Type
            If TileType = TILE_TYPE_BLOCKED Then
                CanPetMove = False
            End If
            
        End If

        If Dir = DIR_RIGHT Then
        
            If Map(Player(PetNum).Pet.Map).Right > 0 And Map(Player(PetNum).Pet.Map).Right = Player(PetNum).Pet.MapToGo Then

                'i = Player(PetNum).Pet.Map
                'Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Right
                'Packet = "PETDATA" & SEP_CHAR
                'Packet = Packet & PetNum & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.Alive & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.Map & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.x & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.y & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.Dir & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.Sprite & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.HP & SEP_CHAR
                'Packet = Packet & Player(PetNum).Pet.Level * 5 & SEP_CHAR
                'Packet = Packet & END_CHAR
                'Call SendDataToMap(Player(PetNum).Pet.Map, Packet)
                'Call SendDataToMap(i, Packet)
                CanPetMove = True
            End If
            
            TileType = Map(mapnum).Tile(PetX + 1, PetY).Type
            If TileType = TILE_TYPE_BLOCKED Then
                CanPetMove = False
            End If
            
            
        End If
        Exit Function
    End If

    'If Grid(Player(PetNum).Pet.Map).Loc(x, y).Blocked = True Then Exit Function
    'CanPetMove = True
End Function

Sub NpcAttackPet(ByVal MapNpcNum As Long, _
   ByVal Victim As Long, _
   ByVal Damage As Long)
Dim Name As String
Dim mapnum As Long
Dim Packet As String

    ' Check for subscript out of range
    If MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for subscript out of range
    If MapNPC(Player(Victim).Pet.Map, MapNpcNum).num <= 0 Then
        Exit Sub
    End If

    ' Send this packet so they can see the npc attacking
    Call SendDataToMap(Player(Victim).Pet.Map, "NPCATTACKPET" & SEP_CHAR & MapNpcNum & SEP_CHAR & Victim & SEP_CHAR & END_CHAR)
    mapnum = Player(Victim).Pet.Map
    Name = Trim$(NPC(MapNPC(mapnum, MapNpcNum).num).Name)

    If Damage >= Player(Victim).Pet.HP Then
        Call BattleMsg(Victim, "Tu mascota ha muerto!", Red, 1)
        Call SetPlayerPetAlive(Victim, 0)
        Player(Victim).Pet.Alive = NO
        Call savepetDeath(Victim)
      '  Call TakeFromGrid(Player(Victim).Pet.Map, Player(Victim).Pet.x, Player(Victim).Pet.y)
        MapNPC(mapnum, MapNpcNum).Target = 0
        Packet = "PETDATA" & SEP_CHAR
        Packet = Packet & Victim & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Alive & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Map & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.x & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.y & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Dir & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.SPRITE & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.HP & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.STR * 5 & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.STR & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.DEF & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Speed & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Magi & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Level & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.POINTS & SEP_CHAR
        Packet = Packet & GetPetSP(Victim) & SEP_CHAR
        Packet = Packet & GetPetMAXSP(Victim) & SEP_CHAR
        Packet = Packet & GetPetMP(Victim) & SEP_CHAR
        Packet = Packet & GetPetMAXMP(Victim) & SEP_CHAR
        Packet = Packet & GetPetFP(Victim) & SEP_CHAR
        Packet = Packet & GetPetMAXFP(Victim) & SEP_CHAR
        Packet = Packet & GetPetExp(Victim) & SEP_CHAR
        Packet = Packet & GetPetNextLevel(Victim) & SEP_CHAR
        Packet = Packet & Player(Victim).Pet.Name & SEP_CHAR
        Packet = Packet & END_CHAR
        Call SendDataTo(Victim, Packet)
        Call SendDataToMapBut(Victim, Player(Victim).Pet.Map, Packet)
    Else

        ' Pet not dead, just do the damage
        Player(Victim).Pet.HP = Player(Victim).Pet.HP - Damage
        Packet = "PETHP" & SEP_CHAR & Player(Victim).Pet.STR * 5 & SEP_CHAR & Player(Victim).Pet.HP & SEP_CHAR & END_CHAR
        Call SendDataTo(Victim, Packet)
    End If

    'Call SendDataTo(Victim, "BLITNPCDMGPET" & SEP_CHAR & Damage & SEP_CHAR & END_CHAR)
End Sub

Sub PetAttackNpc(ByVal Attacker As Long, _
   ByVal MapNpcNum As Long, _
   ByVal Damage As Long)
Dim Name As String
Dim N As Long, i As Long
Dim mapnum As Long, npcnum As Long
Dim Dir As Long, x As Long, y As Long
Dim Packet As String

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    ' Send this packet so they can see the pet attacking
    Call SendDataToMap(Player(Attacker).Pet.Map, "PETATTACKNPC" & SEP_CHAR & Attacker & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
    Dim NpcExp As Long
    mapnum = Player(Attacker).Pet.Map
    npcnum = MapNPC(mapnum, MapNpcNum).num
    Name = Trim$(NPC(npcnum).Name)
    MapNPC(mapnum, MapNpcNum).LastAttack = GetTickCount
    For Dir = 0 To 3

        'If MapNpc(mapnum, npcnum).x = DirToX(Player(Attacker).Pet.x, Dir) And MapNpc(mapnum, npcnum).y = DirToY(Player(Attacker).Pet.y, Dir) Then
         '   Packet = "CHANGEPETDIR" & SEP_CHAR & Dir & SEP_CHAR & Attacker & SEP_CHAR & END_CHAR
          '  Call SendDataToMap(Player(Attacker).Pet.Map, Packet)
        'End If
    Next

    If Damage >= MapNPC(mapnum, MapNpcNum).HP Then
        For i = 1 To MAX_NPC_DROPS

            ' Drop the goods if they get it
            N = Int(Rnd * NPC(npcnum).ItemNPC(i).chance) + 1

            If N = 1 Then
                Call SpawnItem(NPC(npcnum).ItemNPC(i).itemnum, NPC(npcnum).ItemNPC(i).ItemValue, mapnum, MapNPC(mapnum, MapNpcNum).x, MapNPC(mapnum, MapNpcNum).y)
            End If
        Next
        Call BattleMsg(Attacker, "[Mascota] Tu mascota ha matado a " & Name & ".", Red, 1)
       ' Call GoLeaderShip(Attacker)
        NpcExp = NPC(npcnum).EXP
        Call SetPetExp(Attacker, GetPetExp(Attacker) + NpcExp)
        Call BattleMsg(Attacker, "[Mascota] Tu mascota ha ganado " & NpcExp & " puntos de experiencia.", BRIGHTGREEN, 0)
        Call CheckPetLevelUp(Attacker)
        Call SendPlayerData(Attacker)
        Call SendStats(Attacker)
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNPC(mapnum, MapNpcNum).num = 0
        MapNPC(mapnum, MapNpcNum).SpawnWait = GetTickCount
        MapNPC(mapnum, MapNpcNum).HP = 0
        Call SendDataToMap(mapnum, "NPCDEAD" & SEP_CHAR & MapNpcNum & SEP_CHAR & END_CHAR)
        'Call TakeFromGrid(mapnum, MapNPC(mapnum, MapNpcNum).x, MapNPC(mapnum, MapNpcNum).y)

        ' Check if target is npc that died and if so set target to 0
        If Player(Attacker).Pet.TargetType = TARGET_TYPE_NPC And Player(Attacker).Pet.Target = MapNpcNum Then
            Player(Attacker).Pet.Target = 0
            Player(Attacker).Pet.TargetType = 0
            Player(Attacker).Pet.MapToGo = 0
        End If
    Else

        ' NPC not dead, just do the damage
        MapNPC(mapnum, MapNpcNum).HP = MapNPC(mapnum, MapNpcNum).HP - Damage

        ' Set the NPC target to the pet
        MapNPC(mapnum, MapNpcNum).TargetType = TARGET_TYPE_PET
        MapNPC(mapnum, MapNpcNum).Target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNPC(mapnum, MapNpcNum).num).Behavior = NPC_BEHAVIOR_GUARD Then
            For i = 1 To MAX_MAP_NPCS

                If MapNPC(mapnum, i).num = MapNPC(mapnum, MapNpcNum).num Then
                    MapNPC(mapnum, i).TargetType = TARGET_TYPE_PET
                    MapNPC(mapnum, i).Target = Attacker
                End If
            Next
        End If
    End If

    'Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNpc(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNpc(MapNum, MapNpcNum).num) & SEP_CHAR & END_CHAR)
    ' Reset attack timer
    Player(Attacker).Pet.AttackTimer = GetTickCount
End Sub


Sub PetMove(ByVal PetNum As Long, _
   ByVal Dir As Long, _
   ByVal Movement As Long)
Dim Packet As String
Dim x As Long
Dim y As Long
Dim i As Long

    If GetPlayerMap(PetNum) <= 0 Or GetPlayerMap(PetNum) > MAX_MAPS Or PetNum <= 0 Or PetNum > MAX_PLAYERS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then Exit Sub
    Player(PetNum).Pet.Dir = Dir
    x = DirToX(Player(PetNum).Pet.x, Dir)
    y = DirToY(Player(PetNum).Pet.y, Dir)

    If IsValid(x, y) Then
        'If Grid(Player(PetNum).Pet.Map).Loc(x, y).Blocked = True Then
            Packet = "CHANGEPETDIR" & SEP_CHAR & Dir & SEP_CHAR & PetNum & SEP_CHAR & END_CHAR
            Call SendDataToMap(Player(PetNum).Pet.Map, Packet)
          '  Exit Sub
        'End If
       ' Call UpdateGrid(Player(PetNum).Pet.Map, Player(PetNum).Pet.x, Player(PetNum).Pet.y, Player(PetNum).Pet.Map, x, y)
        Player(PetNum).Pet.y = y
        Player(PetNum).Pet.x = x
        Packet = "PETMOVE" & SEP_CHAR & PetNum & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
        Call SendDataToMap(Player(PetNum).Pet.Map, Packet)
    Else
        i = Player(PetNum).Pet.Map

        If Dir = DIR_UP Then
            Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Up
            Player(PetNum).Pet.y = MAX_MAPY
        End If

        If Dir = DIR_DOWN Then
            Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Down
            Player(PetNum).Pet.y = 0
        End If

        If Dir = DIR_LEFT Then
            Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Left
            Player(PetNum).Pet.x = MAX_MAPX
        End If

        If Dir = DIR_RIGHT Then
            Player(PetNum).Pet.Map = Map(Player(PetNum).Pet.Map).Right
            Player(PetNum).Pet.x = 0
        End If
        Packet = "PETDATA" & SEP_CHAR
        Packet = Packet & PetNum & SEP_CHAR
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
        Call SendDataToMap(Player(PetNum).Pet.Map, Packet)
        Call SendDataToMap(i, Packet)
    End If
End Sub

Sub PetSub(ByVal index As Long, ByVal PetSprite As Long)
Dim Packet As String
Player(index).Pet.Alive = YES
                        Player(index).Pet.Dir = DIR_UP
                        Player(index).Pet.Map = GetPlayerMap(index)
                        Player(index).Pet.MapToGo = GetPlayerMap(index)
                        Player(index).Pet.SPRITE = PetSprite
                        Player(index).Pet.SpriteSet = 0
                        
                        Call SetPetSTR(index, 5)
                        Call SetPetDEF(index, 5)
                        Call SetPetSPEED(index, 5)
                        Call SetPetMAGI(index, 5)
                        Call SetPetPOINTS(index, 7)
                        Call SetPetExp(index, 0)
                        Call SetPetLevel(index, 1)
                        
                        Player(index).Pet.x = GetPlayerX(index) + Int(Rnd * 3 - 1)

                        If Player(index).Pet.x < 0 Or Player(index).Pet.x > MAX_MAPX Then Player(index).Pet.x = GetPlayerX(index)
                        Player(index).Pet.XToGo = -1
                        Player(index).Pet.y = GetPlayerY(index) + Int(Rnd * 3 - 1)

                        If Player(index).Pet.y < 0 Or Player(index).Pet.y > MAX_MAPY Then Player(index).Pet.y = GetPlayerY(index)
                        Player(index).Pet.YToGo = -1

                        Call SetPetHp(index, GetPetMAXHP(index))
                        Call SetPetSP(index, GetPetMAXSP(index))
                        Call SetPetMP(index, GetPetMAXMP(index))
                        Call SetPetFP(index, GetPetMAXFP(index))
                        
                        
                        'Call AddToGrid(Player(Index).Pet.Map, Player(Index).Pet.x, Player(Index).Pet.y)
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
                        Call SendDataToMap(GetPlayerMap(index), Packet)
                        
                        Call PlayerMsg(index, "[Mascota] Tienes una mascota!", WHITE)
                        Call savepet(index)
                        Call SendDataTo(index, "ptmenu" & SEP_CHAR & END_CHAR)
End Sub

Sub CallKillPet(ByVal index As Long)
Dim Packet As String
            If Player(index).Pet.Alive = YES Then
                Player(index).Pet.Alive = NO
                'Player(Index).Pet.SPRITE = 0
                'Player(Index).Pet.STR = 0
                'Player(Index).Pet.DEF = 0
                'Player(Index).Pet.Speed = 0
                'Player(Index).Pet.Magi = 0
                'Player(Index).Pet.Name = ""
              '  Call TakeFromGrid(Player(Index).Pet.Map, Player(Index).Pet.x, Player(Index).Pet.y)
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
                Packet = Packet & Player(index).Pet.Name & SEP_CHAR
                Packet = Packet & END_CHAR
                Call SendDataToMap(GetPlayerMap(index), Packet)
            ElseIf Player(index).Pet.Alive = NO Then
                Call PlayerMsg(index, "No tienes una mascota.", Red)
            End If
            Exit Sub
End Sub
Public Sub SendPetData(ByVal index)
Dim Packet As String

If index <= 0 Or index > MAX_PLAYERS Then
    Exit Sub
End If

If Player(index).Pet.Alive = NO Then
    Exit Sub
End If

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
            Call SendDataToMap(GetPlayerMap(index), Packet)
End Sub

Sub KillPet(ByVal index As Long)
Dim Packet As String


Player(index).Pet.Alive = NO
'Player(Index).Pet.SPRITE = 0
'Call SetPetSTR(Index, 5)
'Call SetPetDEF(Index, 5)
'Call SetPetSPEED(Index, 5)
'Call SetPetMAGI(Index, 5)
'Call SetPetPOINTS(Index, 0)
'Call SetPetSP(Index, 0)
'Call SetPetMP(Index, 0)
'Call SetPetFP(Index, 0)
'Call SetPetHp(Index, 0)
'Call SetPetExp(Index, 0)

            'Call TakeFromGrid(Player(Index).Pet.Map, Player(Index).Pet.x, Player(Index).Pet.y)
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
            Call SendDataToMap(GetPlayerMap(index), Packet)
            Call savepetDeath(index)
        'Call ClearPet(Index)
End Sub

Sub UsingPetStatPoints(ByVal index As Long, ByVal PointType As Long)

Select Case PointType
    Case 0
    'Gives you a set max
        If GetPetSTR(index) > 99 Then
           Call BattleMsg(index, "Tu mascota tiene el maximo de su fuerza!", 12, 0)
           Exit Sub
        End If
        Call SetPetSTR(index, GetPetSTR(index) + 1)
        Call BattleMsg(index, "Tu mascota ha ganado mas fuerza!", 15, 0)
        Call SetPetHp(index, GetPetMAXHP(index))
    Case 1
    'Gives you a set max
        If GetPetDEF(index) > 99 Then
           Call BattleMsg(index, "Tu mascota tiene el maximo de su defensa!", 12, 0)
           Exit Sub
        End If
        Call SetPetDEF(index, GetPetDEF(index) + 1)
        Call BattleMsg(index, "Tu mascota ha ganado mas defensa!", 15, 0)
    Case 2
    'Gives you a set max
        If GetPetMAGI(index) > 99 Then
           Call BattleMsg(index, "Tu mascota tiene el maximo de su magia!", 12, 0)
           Exit Sub
        End If
        Call SetPetMAGI(index, GetPetMAGI(index) + 1)
        Call BattleMsg(index, "Tu mascota ha ganado mas magia!", 15, 0)
        Call SetPetMP(index, GetPetMAXMP(index))
    Case 3
    'Gives you a set max
        If GetPetSPEED(index) > 99 Then
           Call BattleMsg(index, "Tu mascota tiene el maximo de su velocidad!", 12, 0)
           Exit Sub
        End If
        Call SetPetSPEED(index, GetPetSPEED(index) + 1)
        Call BattleMsg(index, "Tu mascota ha ganado mas velocidad!", 15, 0)
        Call SetPetSP(index, GetPetMAXSP(index))
End Select
Call SetPetPOINTS(index, GetPetPOINTS(index) - 1)
End Sub

Sub CheckPetLevelUp(ByVal index As Long)
Dim i As Long
Dim d As Long
Dim c As Long
Dim xT As Long

    xT = 2
    c = 0

    If GetPetExp(index) >= GetPetNextLevel(index) Then
        If GetPetLevel(index) < MAX_LEVEL Then
            
                Do Until GetPetExp(index) < GetPetNextLevel(index)
                    DoEvents

                    If GetPetLevel(index) < MAX_LEVEL Then
                        If GetPetExp(index) >= GetPetNextLevel(index) Then
                            d = GetPetExp(index) - GetPetNextLevel(index)
                            Call SetPetLevel(index, GetPetLevel(index) + 1)
                            Call SendDataTo(index, "sound" & SEP_CHAR & "CongratulationsNewLevelAchieved" & SEP_CHAR & END_CHAR)
                            Call SetPetExp(index, d)
                            Call SetPetPOINTS(index, GetPetPOINTS(index) + xT)
                            Call SetPetHp(index, GetPetMAXHP(index))
                            Call SetPetSP(index, GetPetMAXSP(index))
                            Call SetPetMP(index, GetPetMAXMP(index))
                            c = c + 1
                        End If
                    End If

                Loop

                If c > 1 Then
                    Call GlobalMsg(GetPlayerName(index) & " ve como su mascota ha ganado " & c & " niveles!", 14)
                Else
                    Call GlobalMsg(GetPlayerName(index) & " ve como su mascota ha ganado un nivel!", 14)
                End If

            Call SendDataToMap(GetPlayerMap(index), "levelup" & SEP_CHAR & index & SEP_CHAR & END_CHAR)
        End If

        If GetPetLevel(index) = MAX_LEVEL Then
            Call SetPetExp(index, Experience(MAX_LEVEL))
        End If
    End If
        
    Call SendPlayerData(index)
End Sub

Sub DoPetMoveSelect(ByVal index As Long, ByVal x As Long, ByVal y As Long)
Dim i As Long
Player(index).Pet.MapToGo = GetPlayerMap(index)
            Player(index).Pet.Target = 0
            Player(index).Pet.XToGo = x
            Player(index).Pet.YToGo = y
            Player(index).Pet.AttackTimer = GetTickCount
            For i = 1 To MAX_PLAYERS

                If IsPlaying(i) Then
                    If GetPlayerMap(i) = Player(index).Pet.Map Then
                        If GetPlayerX(i) = x And GetPlayerY(i) = y Then
                            Player(index).Pet.TargetType = TARGET_TYPE_PLAYER
                            Player(index).Pet.Target = i
                            Call PlayerMsg(index, "[Mascota] El objetivo de tu mascota es " & Trim$(GetPlayerName(i)) & ".", YELLOW)
                            Exit Sub
                        End If
                    End If
                End If
            Next
            For i = 1 To MAX_MAP_NPCS
                If MapNPC(Player(index).Pet.Map, i).num > 0 Then
                    If MapNPC(Player(index).Pet.Map, i).x = x And MapNPC(Player(index).Pet.Map, i).y = y Then
                        Player(index).Pet.TargetType = TARGET_TYPE_NPC
                        Player(index).Pet.Target = i
                        Call PlayerMsg(index, "[Mascota] El objetivo de tu mascota es " & Trim$(NPC(MapNPC(Player(index).Pet.Map, i).num).Name) & ".", YELLOW)
                        Exit Sub
                    End If
                End If
            Next
            Call PlayerMsg(index, "[Mascota] Tu mascota se mueve a (" & x & "," & y & ")", YELLOW)
End Sub

Public Function IsValid(ByVal x As Long, _
   ByVal y As Long) As Boolean
    IsValid = True

    If x < 0 Or x > MAX_MAPX Or y < 0 Or y > MAX_MAPY Then IsValid = False
End Function

