Attribute VB_Name = "modDatabase"
Option Explicit

' ---------------------------------------------------------------------------------------
' Procedure : GetVar
' Purpose   :  Reads a variable from an INI file
' ---------------------------------------------------------------------------------------
Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found

    On Error GoTo GetVar_Error

    szReturn = vbNullString

    sSpaces = Space(5000)

    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)

    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

    On Error GoTo 0
    Exit Function

GetVar_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetVar of Module modDatabase"
End Function

' ---------------------------------------------------------------------------------------
' Procedure : PutVar
' Purpose   : Writes a file to an INI file
' ---------------------------------------------------------------------------------------
Sub PutVar(File As String, Header As String, Var As String, Value As String)
    On Error GoTo PutVar_Error

    Call WritePrivateProfileString(Header, Var, Value, File)

    On Error GoTo 0
    Exit Sub

PutVar_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PutVar of Module modDatabase"
End Sub

Function FileExists(filename As String) As Boolean
    On Error GoTo ErrorHandler
    ' get the attributes and ensure that it isn't a directory
    FileExists = (GetAttr(App.Path & "\" & filename) And vbDirectory) = 0
ErrorHandler:
' if an error occurs, this function returns False
End Function

Function FolderExists(inPath As String) As Boolean
    If LenB(Dir(inPath, vbDirectory)) = 0 Then
        FolderExists = False
    Else
        FolderExists = True
    End If
End Function

Sub LoadExperience()
    On Error GoTo ExpErr
    Dim filename As String
    Dim i As Integer

    Call CheckExperience

    filename = App.Path & "\Experiencia.ini"

    For i = 1 To MAX_LEVEL
        temp = i / MAX_LEVEL * 100
        Call SetStatus("Cargando experiencia... " & temp & "%")
        Experience(i) = GetVar(filename, "EXPERIENCE", "Exp" & i)
    Next i
    Exit Sub

ExpErr:
    Call MsgBox("Error cargando la EXP para el nivel " & i & ". Asegurate que el archivo Experiencia.ini tiene correctas las variables ERR: " & Err.Number & ", Desc: " & Err.Description, vbCritical)
    Call DestroyServer
End Sub

Sub CheckExperience()
    If Not FileExists("Experiencia.ini") Then
        Dim i As Integer

        For i = 1 To MAX_LEVEL
            temp = i / MAX_LEVEL * 100
            Call SetStatus("Gaurdando Experiencia... " & temp & "%")
            Call PutVar(App.Path & "\Experiencia.ini", "EXPERIENCE", "Exp" & i, i * 1500)
        Next i
    End If
End Sub

Sub ClearExperience()
    Dim i As Integer

    For i = 1 To MAX_LEVEL
        Experience(i) = 0
    Next i
End Sub

Sub LoadEmoticon()
    Dim filename As String
    Dim i As Integer

    Call CheckEmoticon

    filename = App.Path & "\Emoticonos.ini"

    For i = 0 To MAX_EMOTICONS
        temp = i / MAX_EMOTICONS * 100
        Call SetStatus("Cargando emoticonos... " & temp & "%")
        Emoticons(i).Pic = GetVar(filename, "EMOTICONS", "Emoticon" & i)
        Emoticons(i).Command = GetVar(filename, "EMOTICONS", "EmoticonC" & i)
    Next i
End Sub

Sub SaveEmoticon(ByVal EmoNum As Long)
    Dim filename As String

    filename = App.Path & "\Emoticonos.ini"

    Call PutVar(filename, "EMOTICONS", "EmoticonC" & EmoNum, Trim$(Emoticons(EmoNum).Command))
    Call PutVar(filename, "EMOTICONS", "Emoticon" & EmoNum, Val(Emoticons(EmoNum).Pic))
End Sub

Sub CheckEmoticon()
    If Not FileExists("Emoticonos.ini") Then
        Dim i As Integer

        For i = 0 To MAX_EMOTICONS
            temp = i / MAX_LEVEL * 100
            Call SetStatus("Guardando emoticonos... " & temp & "%")
            Call PutVar(App.Path & "\Emoticonos.ini", "EMOTICONS", "Emoticon" & i, 0)
            Call PutVar(App.Path & "\Emoticonos.ini", "EMOTICONS", "EmoticonC" & i, vbNullString)
        Next i
    End If
End Sub

Sub ClearEmoticon()
    Dim i As Integer

    For i = 0 To MAX_EMOTICONS
        Emoticons(i).Pic = 0
        Emoticons(i).Command = vbNullString
    Next i
End Sub

Sub LoadElements()
    On Error GoTo ElementErr
    Dim filename As String
    Dim i As Integer

    Call CheckElements

    filename = App.Path & "\Elementos.ini"

    For i = 0 To MAX_ELEMENTS
        temp = i / MAX_ELEMENTS * 100
        Call SetStatus("Cargando elementos... " & temp & "%")
        Element(i).Name = GetVar(filename, "ELEMENTS", "ElementName" & i)
        Element(i).Strong = Val(GetVar(filename, "ELEMENTS", "ElementStrong" & i))
        Element(i).Weak = Val(GetVar(filename, "ELEMENTS", "ElementWeak" & i))
    Next i
    Exit Sub

ElementErr:
    Call MsgBox("Error cargando el elemento " & i & ". Comprueba que las variables estan correctas!", vbCritical)
    Call DestroyServer
    End
End Sub

Sub CheckElements()
    If Not FileExists("Elementos.ini") Then
        Dim i As Integer

        For i = 0 To MAX_ELEMENTS
            temp = i / MAX_ELEMENTS * 100
            Call SetStatus("Guardando elementos... " & temp & "%")
            Call PutVar(App.Path & "\Elementos.ini", "ELEMENTS", "ElementName" & i, vbNullString)
            Call PutVar(App.Path & "\Elementos.ini", "ELEMENTS", "ElementStrong" & i, 0)
            Call PutVar(App.Path & "\Elementos.ini", "ELEMENTS", "ElementWeak" & i, 0)
        Next i
    End If
End Sub

Sub SaveElement(ByVal ElementNum As Long)
    Dim filename As String

    filename = App.Path & "\Elementos.ini"

    Call PutVar(filename, "ELEMENTS", "ElementName" & ElementNum, Trim$(Element(ElementNum).Name))
    Call PutVar(filename, "ELEMENTS", "ElementStrong" & ElementNum, Val(Element(ElementNum).Strong))
    Call PutVar(filename, "ELEMENTS", "ElementWeak" & ElementNum, Val(Element(ElementNum).Weak))
End Sub

Sub SavePlayer(ByVal index As Long)
    Dim filename As String
    Dim F As Long 'File
    Dim i As Integer

    On Error Resume Next

    ' Save login information first
    filename = App.Path & "\Cuentas\" & Trim$(Player(index).Login) & "_Info.ini"

    Call PutVar(filename, "ACCESS", "Login", Trim$(Player(index).Login))
    Call PutVar(filename, "ACCESS", "Password", Trim$(Player(index).Password))
    Call PutVar(filename, "ACCESS", "Email", Trim$(Player(index).Email))

    ' Make the directory
    If LCase$(Dir(App.Path & "\Cuentas\" & Trim$(Player(index).Login), vbDirectory)) <> LCase$(Trim$(Player(index).Login)) Then
        Call MkDir(App.Path & "\Cuentas\" & Trim$(Player(index).Login))
    End If

    ' Now save their characters
    For i = 1 To MAX_CHARS
        filename = App.Path & "\Cuentas\" & Trim$(Player(index).Login) & "\Char" & i & ".dat"

        ' Save the character
        F = FreeFile
        Open filename For Binary As #F
        Put #F, , Player(index).Char(i)
        Close #F

    Next i
End Sub
Function ConvertV000(filename As String) As PlayerRec
    Dim OldRec As V000PlayerRec
    Dim NewRec As PlayerRec
    Dim F As Long
    Dim N As Integer
   
    F = FreeFile
    Open filename For Binary As #F
        Get #F, , OldRec
    Close #F

        ' General
    NewRec.Name = OldRec.Name
    NewRec.Guild = OldRec.Guild
    NewRec.GuildAccess = OldRec.GuildAccess
    NewRec.Sex = OldRec.Sex
    NewRec.Class = OldRec.Class
    NewRec.SPRITE = OldRec.SPRITE
    NewRec.Level = OldRec.Level
    NewRec.EXP = OldRec.EXP
    NewRec.Access = OldRec.Access
    NewRec.PK = OldRec.PK

    ' Vitals
    NewRec.HP = OldRec.HP
    NewRec.MP = OldRec.MP
    NewRec.SP = OldRec.SP

    ' Stats
    NewRec.STR = OldRec.STR
    NewRec.DEF = OldRec.DEF
    NewRec.Speed = OldRec.Speed
    NewRec.Magi = OldRec.Magi
    NewRec.POINTS = OldRec.POINTS

    ' Worn equipment
    NewRec.ArmorSlot = OldRec.ArmorSlot
    NewRec.WeaponSlot = OldRec.WeaponSlot
    NewRec.HelmetSlot = OldRec.HelmetSlot
    NewRec.ShieldSlot = OldRec.ShieldSlot
    NewRec.LegsSlot = OldRec.LegsSlot
    NewRec.RingSlot = OldRec.RingSlot
    NewRec.NecklaceSlot = OldRec.NecklaceSlot

    ' Inventory
    For N = 1 To MAX_INV
        NewRec.Inv(N) = OldRec.Inv(N)
    Next N
    For N = 1 To MAX_PLAYER_SPELLS
        NewRec.Spell(N) = OldRec.Spell(N)
    Next N
    For N = 1 To MAX_BANK
        NewRec.Bank(N) = OldRec.Bank(N)
    Next N

    ' Position
    NewRec.Map = OldRec.Map
    NewRec.x = OldRec.x
    NewRec.y = OldRec.y
    NewRec.Dir = OldRec.Dir

    NewRec.TargetNPC = OldRec.TargetNPC

    NewRec.Head = OldRec.Head
    NewRec.Body = OldRec.Body
    NewRec.Leg = OldRec.Leg

    NewRec.paperdoll = OldRec.paperdoll

    NewRec.MAXHP = OldRec.MAXHP
    NewRec.MAXMP = OldRec.MAXMP
    NewRec.MAXSP = OldRec.MAXSP
    
    NewRec.NpcKillType = OldRec.NpcKillType
    NewRec.NpcKillamount = OldRec.NpcKillamount
    NewRec.NpcKillQuestFlag = OldRec.NpcKillQuestFlag
    For N = 1 To MAX_QUESTS
        NewRec.QuestFlags(N) = OldRec.QuestFlags(N)
    Next N

    NewRec.PetSprite = OldRec.PetSprite
    NewRec.PetAlive = OldRec.PetAlive
    NewRec.PetMap = OldRec.PetMap
    NewRec.PetX = OldRec.PetX
    NewRec.PetY = OldRec.PetY
    NewRec.PetDIR = OldRec.PetDIR
    NewRec.PetHP = OldRec.PetHP
    NewRec.PetSP = OldRec.PetSP
    NewRec.PetMP = OldRec.PetMP
    NewRec.PetFP = OldRec.PetFP
    NewRec.PetMaxHP = OldRec.PetMaxHP
    NewRec.PetMaxSP = OldRec.PetMaxSP
    NewRec.PetMaxMP = OldRec.PetMaxMP
    NewRec.PetMaxFP = OldRec.PetMaxFP
    NewRec.PetMapToGo = OldRec.PetMapToGo
    NewRec.PetLevel = OldRec.PetLevel
    NewRec.PetSpriteSet = OldRec.PetSpriteSet
    NewRec.PetSTR = OldRec.PetSTR
    NewRec.PetDEF = OldRec.PetDEF
    NewRec.PetSPEED = OldRec.PetSPEED
    NewRec.PetMAGI = OldRec.PetMAGI
    NewRec.PetPOINTS = OldRec.PetPOINTS
    NewRec.PetEXP = OldRec.PetEXP
    NewRec.PetNAME = OldRec.PetNAME
    NewRec.PetTNL = OldRec.PetTNL
    
    NewRec.InParty = OldRec.InParty
    NewRec.LookingForParty = OldRec.LookingForParty
    NewRec.PartyInvitedTo = OldRec.PartyInvitedTo
    NewRec.PartyInvitedToBy = OldRec.PartyInvitedToBy
    NewRec.Party = OldRec.Party

    ' *** add new fields ***

    ' version info
   
    NewRec.Vflag = 128
    NewRec.Ver = 2
    NewRec.SubVer = 8
    NewRec.Rel = 0

    ConvertV000 = NewRec
End Function

Public Sub LoadPlayer(ByVal index As Long, ByVal Name As String)
    Dim F As Long
    Dim i As Integer
    Dim filename As String

    On Error GoTo PlayerErr

    Call ClearPlayer(index)

    ' Load the account settings
    filename = App.Path & "\Cuentas\" & Trim$(Name) & "_Info.ini"

    Player(index).Login = Name
    Player(index).Password = GetVar(filename, "ACCESS", "Password")
    Player(index).Email = GetVar(filename, "ACCESS", "Email")

    ' Load the .dat
    For i = 1 To MAX_CHARS
        filename = App.Path & "\Cuentas\" & Trim$(Player(index).Login) & "\Char" & i & ".dat"

        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Player(index).Char(i)
        Close #F
        If Player(index).Char(i).Vflag <> 128 Then
        Player(index).Char(i) = ConvertV000(filename)
    End If
    Next i

    Exit Sub

PlayerErr:
    Call MsgBox("Couldn't load index " & index & " for " & Name & "!", vbCritical)
    Call DestroyServer
End Sub

Function AccountExists(ByVal Name As String) As Boolean
    If FileExists("\Cuentas\" & Trim$(Name) & "_Info.ini") Then
        AccountExists = True
    Else
        AccountExists = False
    End If
End Function

Function CharExist(ByVal index As Long, ByVal CharNum As Long) As Boolean
    If Trim$(Player(index).Char(CharNum).Name) <> vbNullString Then
        CharExist = True
    Else
        CharExist = False
    End If
End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim RightPassword As String

    PasswordOK = False

    If AccountExists(Name) Then
        RightPassword = GetVar(App.Path & "\Cuentas\" & Trim$(Name) & "_Info.ini", "ACCESS", "Password")

        If Trim$(Password) = Trim$(RightPassword) Then
            PasswordOK = True
        End If
    End If
End Function

Public Sub AddAccount(ByVal index As Long, ByVal Name As String, ByVal Password As String, ByVal Email As String)
    Dim i As Long

    Player(index).Login = Name
    Player(index).Password = Password
    Player(index).Email = Email

    For i = 1 To MAX_CHARS
        Call ClearChar(index, i)
    Next i

    Call SavePlayer(index)

    If ACC_VERIFY = 1 Then
        Call PutVar(App.Path & "\Cuentas\" & Trim$(Player(index).Login) & "_Info.ini", "ACCESS", "verified", 0)
    End If
    
    Call ClearPlayer(index)
End Sub

Sub AddChar(ByVal index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal CharNum As Long, ByVal headc As Long, ByVal bodyc As Long, ByVal logc As Long)
    Dim F As Long
    Dim i As Long

    If Trim$(Player(index).Char(CharNum).Name) = vbNullString Then
        Player(index).CharNum = CharNum

        Player(index).Char(CharNum).Name = Name
        Player(index).Char(CharNum).Sex = Sex
        Player(index).Char(CharNum).Class = ClassNum

        If Player(index).Char(CharNum).Sex = SEX_MALE Then
            Player(index).Char(CharNum).SPRITE = ClassData(ClassNum).MaleSprite
        Else
            Player(index).Char(CharNum).SPRITE = ClassData(ClassNum).FemaleSprite
        End If

        Player(index).Char(CharNum).Level = 1

        Player(index).Char(CharNum).STR = ClassData(ClassNum).STR
        Player(index).Char(CharNum).DEF = ClassData(ClassNum).DEF
        Player(index).Char(CharNum).Speed = ClassData(ClassNum).Speed
        Player(index).Char(CharNum).Magi = ClassData(ClassNum).Magi

        If ClassData(ClassNum).Map <= 0 Then
            ClassData(ClassNum).Map = 1
        End If
        If ClassData(ClassNum).x < 0 Or ClassData(ClassNum).x > MAX_MAPX Then
            ClassData(ClassNum).x = Int(ClassData(ClassNum).x / 2)
        End If
        If ClassData(ClassNum).y < 0 Or ClassData(ClassNum).y > MAX_MAPY Then
            ClassData(ClassNum).y = Int(ClassData(ClassNum).y / 2)
        End If
        Player(index).Char(CharNum).Map = ClassData(ClassNum).Map
        Player(index).Char(CharNum).x = ClassData(ClassNum).x
        Player(index).Char(CharNum).y = ClassData(ClassNum).y

        Player(index).Char(CharNum).HP = GetPlayerMaxHP(index)
        Player(index).Char(CharNum).MP = GetPlayerMaxMP(index)
        Player(index).Char(CharNum).SP = GetPlayerMaxSP(index)

        Player(index).Char(CharNum).MAXHP = GetPlayerMaxHP(index)
        Player(index).Char(CharNum).MAXMP = GetPlayerMaxMP(index)
        Player(index).Char(CharNum).MAXSP = GetPlayerMaxSP(index)

        Player(index).Char(CharNum).Head = headc
        Player(index).Char(CharNum).Body = bodyc
        Player(index).Char(CharNum).Leg = logc
        
            ' version info
        Player(index).Char(CharNum).Vflag = 128
        Player(index).Char(CharNum).Ver = 2
        Player(index).Char(CharNum).SubVer = 8
        Player(index).Char(CharNum).Rel = 0

        Player(index).Char(CharNum).paperdoll = 1
        
        Player(index).Char(CharNum).NpcKillType = ""
        Player(index).Char(CharNum).NpcKillamount = 0
        Player(index).Char(CharNum).NpcKillQuestFlag = 0
        Player(index).Char(CharNum).color(1) = UserCr(1)
        Player(index).Char(CharNum).color(2) = UserCr(2)
        Player(index).Char(CharNum).color(3) = UserCr(3)
        
        
        For i = 1 To MAX_QUESTS
        Player(index).Char(CharNum).QuestFlags(i) = 0
        Next

        ' Append name to file
        F = FreeFile
        Open App.Path & "\Cuentas\CharList.txt" For Append As #F
        Print #F, Name
        Close #F

        Call SavePlayer(index)

        Exit Sub
    End If
End Sub

Sub DelChar(ByVal index As Long, ByVal CharNum As Long)
    MyScript.ExecuteStatement "Scripts\Main.txt", "OnEraseChar " & index & "," & CharNum
    Call DeleteName(Player(index).Char(CharNum).Name)
    Call ClearChar(index, CharNum)
    Call SavePlayer(index)
End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim F As Long
    Dim S As String

    FindChar = False

    F = FreeFile
    Open App.Path & "\Cuentas\CharList.txt" For Input As #F
    Do While Not EOF(F)
        Input #F, S

        If Trim$(LCase$(S)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #F
            Exit Function
        End If
    Loop
    Close #F
End Function

Sub SaveAllPlayersOnline()
    Dim i As Integer

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SavePlayer(i)
        End If
    Next i
End Sub

Sub LoadClasses()
    Dim filename As String
    Dim i As Long

    ' -1 because it is 0 indexed and it needs to include 0
    MAX_CLASSES = -1

    Do While FileExists("\Clases\Class" & i & ".ini")
        MAX_CLASSES = MAX_CLASSES + 1
        i = i + 1
    Loop
    
    If MAX_CLASSES = -1 Then
        MAX_CLASSES = 0
    End If

    ReDim ClassData(0 To MAX_CLASSES) As ClassRec

    Call ClearClasses

    For i = 0 To MAX_CLASSES

        On Error Resume Next ' used if next line tries to divide by 0

        temp = i / MAX_CLASSES * 100

        On Error GoTo ClassErr

        Call SetStatus("Cargando clases... " & temp & "%")
        filename = App.Path & "\Clases\Class" & i & ".ini"

        ' Check if class exists
        If Not FileExists("\Clases\Class" & i & ".ini") Then
            Call PutVar(filename, "CLASS", "Name", Trim$(ClassData(i).Name))
            Call PutVar(filename, "CLASS", "MaleSprite", CStr(ClassData(i).MaleSprite))
            Call PutVar(filename, "CLASS", "FemaleSprite", CStr(ClassData(i).FemaleSprite))
            Call PutVar(filename, "CLASS", "Description", CStr(ClassData(i).Desc))
            Call PutVar(filename, "CLASS", "STR", CStr(ClassData(i).STR))
            Call PutVar(filename, "CLASS", "DEF", CStr(ClassData(i).DEF))
            Call PutVar(filename, "CLASS", "SPEED", CStr(ClassData(i).Speed))
            Call PutVar(filename, "CLASS", "MAGI", CStr(ClassData(i).Magi))
            Call PutVar(filename, "CLASS", "MAP", CStr(ClassData(i).Map))
            Call PutVar(filename, "CLASS", "X", CStr(ClassData(i).x))
            Call PutVar(filename, "CLASS", "Y", CStr(ClassData(i).y))
            Call PutVar(filename, "CLASS", "Locked", CStr(ClassData(i).Locked))
            Call PutVar(filename, "CLASS", "Gender", "1")
            Call PutVar(filename, "CLASS", "Gender1", "Gender1")
            Call PutVar(filename, "CLASS", "Gender2", "Gender2")
            
            
        End If

        ClassData(i).Name = GetVar(filename, "CLASS", "Name")
        ClassData(i).MaleSprite = GetVar(filename, "CLASS", "MaleSprite")
        ClassData(i).FemaleSprite = GetVar(filename, "CLASS", "FemaleSprite")
        ClassData(i).Desc = GetVar(filename, "CLASS", "Desc")
        ClassData(i).STR = Val(GetVar(filename, "CLASS", "STR"))
        ClassData(i).DEF = Val(GetVar(filename, "CLASS", "DEF"))
        ClassData(i).Speed = Val(GetVar(filename, "CLASS", "SPEED"))
        ClassData(i).Magi = Val(GetVar(filename, "CLASS", "MAGI"))
        ClassData(i).Map = Val(GetVar(filename, "CLASS", "MAP"))
        ClassData(i).x = Val(GetVar(filename, "CLASS", "X"))
        ClassData(i).y = Val(GetVar(filename, "CLASS", "Y"))
        ClassData(i).Locked = Val(GetVar(filename, "CLASS", "Locked"))
        ClassData(i).Gender = Val(GetVar(filename, "CLASS", "Gender"))
        ClassData(i).Gender1 = GetVar(filename, "CLASS", "Gender1")
        ClassData(i).Gender2 = GetVar(filename, "CLASS", "Gender2")
    Next i
    Exit Sub

ClassErr:
    Call MsgBox("Error loading class " & i & ". Check that all the variables in your class files exist!")
    Call DestroyServer
    End
End Sub

Sub SaveClasses()
    Dim filename As String
    Dim i As Long

    For i = 0 To MAX_CLASSES
        On Error Resume Next ' if MAX_CLASSES is 0
    
        temp = i / MAX_CLASSES * 100
        Call SetStatus("Guardando clases... " & temp & "%")
        filename = App.Path & "\Clases\Class" & i & ".ini"
        If Not FileExists("Classes\Class" & i & ".ini") Then
            Call PutVar(filename, "CLASS", "Name", Trim$(ClassData(i).Name))
            Call PutVar(filename, "CLASS", "MaleSprite", CStr(ClassData(i).MaleSprite))
            Call PutVar(filename, "CLASS", "FemaleSprite", CStr(ClassData(i).FemaleSprite))
            Call PutVar(filename, "CLASS", "STR", CStr(ClassData(i).STR))
            Call PutVar(filename, "CLASS", "DEF", CStr(ClassData(i).DEF))
            Call PutVar(filename, "CLASS", "SPEED", CStr(ClassData(i).Speed))
            Call PutVar(filename, "CLASS", "MAGI", CStr(ClassData(i).Magi))
            Call PutVar(filename, "CLASS", "MAP", CStr(ClassData(i).Map))
            Call PutVar(filename, "CLASS", "X", CStr(ClassData(i).x))
            Call PutVar(filename, "CLASS", "Y", CStr(ClassData(i).y))
            Call PutVar(filename, "CLASS", "Locked", CStr(ClassData(i).Locked))
            ClassData(i).Gender = Val(GetVar(filename, "CLASS", "Gender"))
            ClassData(i).Gender1 = GetVar(filename, "CLASS", "Gender1", "Gender1")
            ClassData(i).Gender2 = GetVar(filename, "CLASS", "Gender2", "Gender2")
        End If
    Next i
End Sub

Sub SaveItems()
    Dim i As Long

    Call SetStatus("Guardando Objetos... ")
    For i = 1 To MAX_ITEMS
        If Not FileExists("objetos\item" & i & ".dat") Then
            temp = i / MAX_ITEMS * 100
            Call SetStatus("Guardando Objetos... " & temp & "%")
            Call SaveItem(i)
        End If
    Next i
End Sub

Sub SaveItem(ByVal ItemNum As Long)
    Dim filename As String
    Dim F  As Long
    filename = App.Path & "\Objetos\item" & ItemNum & ".dat"

    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Item(ItemNum)
    Close #F
End Sub

Sub LoadItems()
    Dim filename As String
    Dim i As Long
    Dim F As Long

    Call CheckItems

    For i = 1 To MAX_ITEMS
        temp = i / MAX_ITEMS * 100
        Call SetStatus("Cargando objetos... " & temp & "%")

        filename = App.Path & "\Objetos\Item" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Item(i)
        Close #F

    Next i
End Sub

Sub CheckItems()
    Call SaveItems
End Sub

Sub SaveShops()
    Dim i As Long

    Call SetStatus("Guardando tiendas... ")
    For i = 1 To MAX_SHOPS
        If Not FileExists("tiendas\shop" & i & ".dat") Then
            temp = i / MAX_SHOPS * 100
            Call SetStatus("Guardando tiendas... " & temp & "%")
            Call SaveShop(i)
        End If
    Next i
End Sub

Sub SaveShop(ByVal ShopNum As Long)
    Dim filename As String
    Dim F As Long

    filename = App.Path & "\Tiendas\shop" & ShopNum & ".dat"

    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Shop(ShopNum)
    Close #F
End Sub

Sub LoadShops()
    Dim filename As String
    Dim i As Long, F As Long

    Call CheckShops

    For i = 1 To MAX_SHOPS
        temp = i / MAX_SHOPS * 100
        Call SetStatus("Cargando tiendas... " & temp & "%")
        filename = App.Path & "\Tiendas\shop" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Shop(i)
        Close #F

    Next i
End Sub

Sub CheckShops()
    Call SaveShops
End Sub

Sub SaveSpell(ByVal spellnum As Long)
    Dim filename As String
    Dim F As Long

    filename = App.Path & "\Hechizos\Hechizos" & spellnum & ".dat"

    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Spell(spellnum)
    Close #F
End Sub

Sub SaveSpells()
    Dim i As Long

    Call SetStatus("Guardando hechizos... ")
    For i = 1 To MAX_SPELLS
        If Not FileExists("hechizos\Hechizos" & i & ".dat") Then
            temp = i / MAX_SPELLS * 100
            Call SetStatus("Guardando hechizos... " & temp & "%")
            Call SaveSpell(i)
        End If
    Next i
End Sub

Sub LoadSpells()
    Dim filename As String
    Dim i As Long
    Dim F As Long

    Call CheckSpells

    For i = 1 To MAX_SPELLS
        temp = i / MAX_SPELLS * 100
        Call SetStatus("Cargando hechizos... " & temp & "%")

        filename = App.Path & "\Hechizos\Hechizos" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Spell(i)
        Close #F

    Next i
End Sub

Sub CheckSpells()
    Call SaveSpells
End Sub

Sub SaveNpcs()
    Dim i As Long

    Call SetStatus("Guardando NPCS... ")

    For i = 1 To MAX_NPCS
        If Not FileExists("npcs\npc" & i & ".dat") Then
            temp = i / MAX_NPCS * 100
            Call SetStatus("Guardando npcs... " & temp & "%")
            Call SaveNpc(i)
        End If
    Next i
End Sub

Sub SaveNpc(ByVal npcnum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.Path & "\npcs\npc" & npcnum & ".dat"

    F = FreeFile
    Open filename For Binary As #F
    Put #F, , NPC(npcnum)
    Close #F
End Sub

Sub LoadNpcs()
    Dim filename As String
    Dim i As Integer
    Dim F As Long

    Call CheckNpcs

    For i = 1 To MAX_NPCS
        temp = i / MAX_NPCS * 100
        Call SetStatus("Cargando NPCS... " & temp & "%")
        filename = App.Path & "\npcs\npc" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , NPC(i)
        Close #F

    Next i
End Sub

Sub CheckNpcs()
    Call SaveNpcs
End Sub

Sub SaveMap(ByVal mapnum As Long)
    Dim filename As String
    Dim F As Long

    filename = App.Path & "\Mapas\map" & mapnum & ".dat"

    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Map(mapnum)
    Close #F
End Sub

Sub LoadMaps()
    Dim filename As String
    Dim i As Integer
    Dim F As Integer

    Call CheckMaps

    For i = 1 To MAX_MAPS
        temp = i / MAX_MAPS * 100
        Call SetStatus("Cargando mapas... " & temp & "%")
        filename = App.Path & "\Mapas\map" & i & ".dat"

        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Map(i)
        Close #F

    Next i

End Sub

Sub CheckMaps()
    Dim filename As String
    Dim i As Integer

    Call ClearMaps

    For i = 1 To MAX_MAPS
        filename = "mapas\map" & i & ".dat"

        ' Check to see if map exists, if it doesn't, create it.
        If Not FileExists(filename) Then
            temp = i / MAX_MAPS * 100
            Call SetStatus("Guardando mapas... " & temp & "%")
            Call SaveMap(i)
        End If
    Next i
End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
    Dim filename As String
    Dim FileID As Long

    filename = App.Path & "\BanList.txt"

    FileID = FreeFile
    Open filename For Append As #FileID
        Print #FileID, GetPlayerIP(BanPlayerIndex)
    Close #FileID

    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " ha sido baneado de " & GAME_NAME & " por " & GetPlayerName(BannedByIndex) & "!", WHITE)
    Call AddLog(GetPlayerName(BannedByIndex) & " ha sido baneado " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "Has sido baneado por " & GetPlayerName(BannedByIndex) & "!")
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

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long, f2 As Long
    Dim S As String

    Call FileCopy(App.Path & "\Cuentas\CharList.txt", App.Path & "\Cuentas\chartemp.txt")

    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\Cuentas\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\Cuentas\CharList.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, S
        If Trim$(LCase$(S)) <> Trim$(LCase$(Name)) Then
            Print #f2, S
        End If
    Loop

    Close #f1
    Close #f2

    Call Kill(App.Path & "\Cuentas\chartemp.txt")
End Sub

Sub BanByServer(ByVal index As Long, ByVal Reason As String)
    Dim filename As String
    Dim FileID As Long
    
    If IsPlaying(index) Then
        filename = App.Path & "\BanList.txt"

        FileID = FreeFile
        Open filename For Append As #FileID
            Print #FileID, GetPlayerIP(index)
        Close #FileID

        If LenB(Reason) <> 0 Then
            Call GlobalMsg(GetPlayerName(index) & " ha sido baneado por el servidor. - Razón(" & Reason & ")", WHITE)
            Call AddLog("El servidor ha baneado a " & GetPlayerName(index) & ". Razón(" & Reason & ")", ADMIN_LOG)
            Call AlertMsg(index, "Has sido baneado por el servidor. - Razón:(" & Reason & ")")
        Else
            Call GlobalMsg(GetPlayerName(index) & " ha sido baneado por el servidor!", WHITE)
            Call AddLog("El servidor ha baneado a " & GetPlayerName(index) & ".", ADMIN_LOG)
            Call AlertMsg(index, "Has sido baneado por el servidor!")
        End If
    End If
End Sub

Sub SaveLogs()
    Dim filename As String
    Dim CurDate As String
    Dim CurTime As String
    Dim FileID As Integer

    On Error Resume Next

    If Not FolderExists(App.Path & "\Logs") Then
        Call MkDir(App.Path & "\Logs")
    End If

    'CurDate = Date
    CurDate = Replace(Date, "/", "-")
    CurTime = Replace(Time, ":", "-")

    If Not FolderExists(App.Path & "\Logs\" & CurDate) Then
        Call MkDir(App.Path & "\Logs\" & CurDate)
    End If
   
    Call MkDir(App.Path & "\Logs\" & CurDate & "/" & CurTime)
    frmServer.versionnueva.Caption = "Logs Guardados! " & CurDate & "-" & CurTime
    FileID = FreeFile

    filename = App.Path & "\Logs\" & CurDate & "/" & CurTime & "\Main.txt"
    Open filename For Output As #FileID
        Print #FileID, frmServer.txtText(0).Text
    Close #FileID

    filename = App.Path & "\Logs\" & CurDate & "/" & CurTime & "\Broadcast.txt"
    Open filename For Output As #FileID
        Print #FileID, frmServer.txtText(1).Text
    Close #FileID

    filename = App.Path & "\Logs\" & CurDate & "/" & CurTime & "\Global.txt"
    Open filename For Output As #FileID
        Print #FileID, frmServer.txtText(2).Text
    Close #FileID

    filename = App.Path & "\Logs\" & CurDate & "/" & CurTime & "\Map.txt"
    Open filename For Output As #FileID
        Print #FileID, frmServer.txtText(3).Text
    Close #FileID

    filename = App.Path & "\Logs\" & CurDate & "/" & CurTime & "\Private.txt"
    Open filename For Output As #FileID
        Print #FileID, frmServer.txtText(4).Text
    Close #FileID

    filename = App.Path & "\Logs\" & CurDate & "/" & CurTime & "\Admin.txt"
    Open filename For Output As #FileID
        Print #FileID, frmServer.txtText(5).Text
    Close #FileID

    filename = App.Path & "\Logs\" & CurDate & "/" & CurTime & "\Emote.txt"
    Open filename For Output As #FileID
        Print #FileID, frmServer.txtText(6).Text
    Close #FileID
End Sub

Sub LoadArrows()
    Dim filename As String
    Dim i As Long

    Call CheckArrows

    filename = App.Path & "\Flechas.ini"

    For i = 1 To MAX_ARROWS
        temp = i / MAX_ARROWS * 100
        Call SetStatus("Cargando Flechas... " & temp & "%")
        Arrows(i).Name = GetVar(filename, "Arrow" & i, "ArrowName")
        Arrows(i).Pic = GetVar(filename, "Arrow" & i, "ArrowPic")
        Arrows(i).Range = GetVar(filename, "Arrow" & i, "ArrowRange")
        Arrows(i).Amount = GetVar(filename, "Arrow" & i, "ArrowAmount")

    Next i
End Sub

Sub CheckArrows()
    If Not FileExists("Flechas.ini") Then
        Dim i As Long

        For i = 1 To MAX_ARROWS
            temp = i / MAX_ARROWS * 100
            Call SetStatus("Guardando Flechas... " & temp & "%")

            Call PutVar(App.Path & "\Flechas.ini", "Arrow" & i, "ArrowName", vbNullString)
            Call PutVar(App.Path & "\Flechas.ini", "Arrow" & i, "ArrowPic", 0)
            Call PutVar(App.Path & "\Flechas.ini", "Arrow" & i, "ArrowRange", 0)
            Call PutVar(App.Path & "\Flechas.ini", "Arrow" & i, "ArrowAmount", 0)
        Next i
    End If
End Sub

Sub ClearArrows()
    Dim i As Long

    For i = 1 To MAX_ARROWS
        Arrows(i).Name = vbNullString
        Arrows(i).Pic = 0
        Arrows(i).Range = 0
        Arrows(i).Amount = 0
    Next i
End Sub

Sub SaveArrow(ByVal ArrowNum As Long)
    Dim filename As String

    filename = App.Path & "\Flechas.ini"

    Call PutVar(filename, "Arrow" & ArrowNum, "ArrowName", Trim$(Arrows(ArrowNum).Name))
    Call PutVar(filename, "Arrow" & ArrowNum, "ArrowPic", Val(Arrows(ArrowNum).Pic))
    Call PutVar(filename, "Arrow" & ArrowNum, "ArrowRange", Val(Arrows(ArrowNum).Range))
    Call PutVar(filename, "Arrow" & ArrowNum, "ArrowAmount", Val(Arrows(ArrowNum).Amount))
End Sub

Sub ClearTempTile()
    Dim i As Long, y As Long, x As Long

    For i = 1 To MAX_MAPS
        TempTile(i).DoorTimer = 0

        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                TempTile(i).DoorOpen(x, y) = NO
            Next x
        Next y
    Next i
End Sub

Sub ClearClasses()
    Dim i As Long

    For i = 0 To MAX_CLASSES
        ClassData(i).Name = vbNullString
        ClassData(i).AdvanceFrom = 0
        ClassData(i).LevelReq = 0
        ClassData(i).Type = 1
        ClassData(i).STR = 0
        ClassData(i).DEF = 0
        ClassData(i).Speed = 0
        ClassData(i).Magi = 0
        ClassData(i).FemaleSprite = 0
        ClassData(i).MaleSprite = 0
        ClassData(i).Desc = vbNullString
        ClassData(i).Map = 0
        ClassData(i).x = 0
        ClassData(i).y = 0
    Next i
End Sub

Sub ClearPlayer(ByVal index As Long)
    Dim i As Long
    Dim N As Long

    Player(index).Login = vbNullString
    Player(index).Password = vbNullString
    For i = 1 To MAX_CHARS
        Player(index).Char(i).Name = vbNullString
        Player(index).Char(i).Class = 0
        Player(index).Char(i).Level = 0
        Player(index).Char(i).SPRITE = 0
        Player(index).Char(i).EXP = 0
        Player(index).Char(i).Access = 0
        Player(index).Char(i).PK = NO
        Player(index).Char(i).POINTS = 0
        Player(index).Char(i).Guild = vbNullString

        Player(index).Char(i).HP = 0
        Player(index).Char(i).MP = 0
        Player(index).Char(i).SP = 0

        Player(index).Char(i).MAXHP = 0
        Player(index).Char(i).MAXMP = 0
        Player(index).Char(i).MAXSP = 0

        Player(index).Char(i).STR = 0
        Player(index).Char(i).DEF = 0
        Player(index).Char(i).Speed = 0
        Player(index).Char(i).Magi = 0
        
        Player(index).Char(i).PetSprite = 0
        Player(index).Char(i).PetAlive = 0
        Player(index).Char(i).PetMap = 0
        Player(index).Char(i).PetX = 0
        Player(index).Char(i).PetY = 0
        Player(index).Char(i).PetDIR = 0
        Player(index).Char(i).PetHP = 0
        Player(index).Char(i).PetSP = 0
        Player(index).Char(i).PetMP = 0
        Player(index).Char(i).PetFP = 0
        Player(index).Char(i).PetMaxHP = 0
        Player(index).Char(i).PetMaxSP = 0
        Player(index).Char(i).PetMaxMP = 0
        Player(index).Char(i).PetMaxFP = 0
        Player(index).Char(i).PetLevel = 0
        Player(index).Char(i).PetSpriteSet = 0
        Player(index).Char(i).PetSTR = 0
        Player(index).Char(i).PetDEF = 0
        Player(index).Char(i).PetMAGI = 0
        Player(index).Char(i).PetSPEED = 0
        Player(index).Char(i).PetPOINTS = 0
        Player(index).Char(i).PetEXP = 0
        Player(index).Char(i).PetNAME = ""
        Player(index).Char(i).PetTNL = 0

        For N = 1 To MAX_INV
            Player(index).Char(i).Inv(N).num = 0
            Player(index).Char(i).Inv(N).Value = 0
            Player(index).Char(i).Inv(N).Dur = 0
        Next N
        For N = 1 To MAX_BANK
            Player(index).Char(i).Bank(N).num = 0
            Player(index).Char(i).Bank(N).Value = 0
            Player(index).Char(i).Bank(N).Dur = 0
        Next N
        For N = 1 To MAX_PLAYER_SPELLS
            Player(index).Char(i).Spell(N) = 0
        Next N
        For N = 1 To MAX_QUESTS
            Player(index).Char(i).QuestFlags(N) = 0
        Next

        Player(index).Char(i).ArmorSlot = 0
        Player(index).Char(i).WeaponSlot = 0
        Player(index).Char(i).HelmetSlot = 0
        Player(index).Char(i).ShieldSlot = 0
        Player(index).Char(i).LegsSlot = 0
        Player(index).Char(i).RingSlot = 0
        Player(index).Char(i).NecklaceSlot = 0
        
        Player(index).Char(i).NpcKillType = ""
        Player(index).Char(i).NpcKillamount = 0
        Player(index).Char(i).NpcKillQuestFlag = 0

        Player(index).Char(i).Map = 0
        Player(index).Char(i).x = 0
        Player(index).Char(i).y = 0
        Player(index).Char(i).Dir = 0

        Player(index).Locked = False
        Player(index).LockedSpells = False
        Player(index).LockedItems = False
        Player(index).LockedAttack = False

        ' Temporary vars
        Player(index).Buffer = vbNullString
        Player(index).IncBuffer = vbNullString
        Player(index).CharNum = 0
        Player(index).InGame = False
        Player(index).AttackTimer = 0
        Player(index).DataTimer = 0
        Player(index).DataBytes = 0
        Player(index).DataPackets = 0
        Player(index).PartyID = 0
        Player(index).InParty = 0
        Player(index).Invited = 0
        Player(index).PartyPlayer = 0
        Player(index).Target = 0
        Player(index).TargetType = 0
        Player(index).CastedSpell = NO
        Player(index).PartyStarter = NO
        Player(index).GettingMap = NO
        Player(index).Emoticon = -1
        Player(index).InTrade = False
        Player(index).TradePlayer = 0
        Player(index).TradeOk = 0
        Player(index).TradeItemMax = 0
        Player(index).TradeItemMax2 = 0
        Player(index).PartyPlayer = 0
        Player(index).InParty = 0
        Player(index).PartyStarter = NO
        Player(index).PartyPlayer = 0
        For N = 1 To MAX_PLAYER_TRADES
            Player(index).Trading(N).InvName = vbNullString
            Player(index).Trading(N).InvNum = 0
        Next N
        Player(index).ChatPlayer = 0
    Next i
Player(index).Pet.Alive = NO
End Sub

Sub ClearChar(ByVal index As Long, ByVal CharNum As Long)
    Dim N As Long
        ' version info
    Player(index).Char(CharNum).Vflag = 128
    Player(index).Char(CharNum).Ver = 2
    Player(index).Char(CharNum).SubVer = 8
    Player(index).Char(CharNum).Rel = 0

    Player(index).Char(CharNum).Name = vbNullString
    Player(index).Char(CharNum).Class = 0
    Player(index).Char(CharNum).SPRITE = 0
    Player(index).Char(CharNum).Level = 0
    Player(index).Char(CharNum).EXP = 0
    Player(index).Char(CharNum).Access = 0
    Player(index).Char(CharNum).PK = NO
    Player(index).Char(CharNum).POINTS = 0
    Player(index).Char(CharNum).Guild = vbNullString

    Player(index).Char(CharNum).HP = 0
    Player(index).Char(CharNum).MP = 0
    Player(index).Char(CharNum).SP = 0

    Player(index).Char(CharNum).MAXHP = 0
    Player(index).Char(CharNum).MAXMP = 0
    Player(index).Char(CharNum).MAXSP = 0

    Player(index).Char(CharNum).STR = 0
    Player(index).Char(CharNum).DEF = 0
    Player(index).Char(CharNum).Speed = 0
    Player(index).Char(CharNum).Magi = 0

    For N = 1 To MAX_INV
        Player(index).Char(CharNum).Inv(N).num = 0
        Player(index).Char(CharNum).Inv(N).Value = 0
        Player(index).Char(CharNum).Inv(N).Dur = 0
    Next N
    For N = 1 To MAX_BANK
        Player(index).Char(CharNum).Bank(N).num = 0
        Player(index).Char(CharNum).Bank(N).Value = 0
        Player(index).Char(CharNum).Bank(N).Dur = 0
    Next N
    For N = 1 To MAX_PLAYER_SPELLS
        Player(index).Char(CharNum).Spell(N) = 0
    Next N
    
    For N = 1 To MAX_QUESTS
        Player(index).Char(CharNum).QuestFlags(N) = 0
    Next

    Player(index).Char(CharNum).ArmorSlot = 0
    Player(index).Char(CharNum).WeaponSlot = 0
    Player(index).Char(CharNum).HelmetSlot = 0
    Player(index).Char(CharNum).ShieldSlot = 0
    Player(index).Char(CharNum).LegsSlot = 0
    Player(index).Char(CharNum).RingSlot = 0
    Player(index).Char(CharNum).NecklaceSlot = 0
    
    Player(index).Char(CharNum).NpcKillType = ""
    Player(index).Char(CharNum).NpcKillamount = 0
    Player(index).Char(CharNum).NpcKillQuestFlag = 0

    Player(index).Char(CharNum).Map = 0
    Player(index).Char(CharNum).x = 0
    Player(index).Char(CharNum).y = 0
    Player(index).Char(CharNum).Dir = 0
    Player(index).Char(CharNum).PetSprite = 0
    Player(index).Char(CharNum).PetAlive = 0
    Player(index).Char(CharNum).PetMap = 0
    Player(index).Char(CharNum).PetX = 0
    Player(index).Char(CharNum).PetY = 0
    Player(index).Char(CharNum).PetDIR = 0
    Player(index).Char(CharNum).PetHP = 0
    Player(index).Char(CharNum).PetSP = 0
    Player(index).Char(CharNum).PetMP = 0
    Player(index).Char(CharNum).PetFP = 0
    Player(index).Char(CharNum).PetMaxHP = 0
    Player(index).Char(CharNum).PetMaxSP = 0
    Player(index).Char(CharNum).PetMaxMP = 0
    Player(index).Char(CharNum).PetMaxFP = 0
    Player(index).Char(CharNum).PetLevel = 0
    Player(index).Char(CharNum).PetSpriteSet = 0
    Player(index).Char(CharNum).PetSTR = 0
    Player(index).Char(CharNum).PetDEF = 0
    Player(index).Char(CharNum).PetMAGI = 0
    Player(index).Char(CharNum).PetSPEED = 0
    Player(index).Char(CharNum).PetPOINTS = 0
    Player(index).Char(CharNum).PetEXP = 0
    Player(index).Char(CharNum).PetNAME = ""
    Player(index).Char(CharNum).PetTNL = 0
    Player(index).Char(CharNum).PartyInvitedTo = 0
    Player(index).Char(CharNum).PartyInvitedToBy = 0
    Player(index).Char(CharNum).LookingForParty = 0
    Player(index).Char(CharNum).InParty = 0
    Player(index).Char(CharNum).Party = 0
End Sub

Sub ClearItem(ByVal index As Long)
    Item(index).Name = vbNullString
    Item(index).Desc = vbNullString

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

    Item(index).addHP = 0
    Item(index).addMP = 0
    Item(index).addSP = 0
    Item(index).AddStr = 0
    Item(index).AddDef = 0
    Item(index).AddMagi = 0
    Item(index).AddSpeed = 0
    Item(index).AddEXP = 0
    Item(index).AttackSpeed = 1000
    Item(index).Price = 0
    Item(index).Stackable = 0
    Item(index).Bound = 0
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
End Sub

Sub ClearNpc(ByVal index As Long)
    Dim i As Long
    NPC(index).Name = vbNullString
    NPC(index).AttackSay = vbNullString
    NPC(index).SPRITE = 0
    NPC(index).SpawnSecs = 0
    NPC(index).Behavior = 0
    NPC(index).Range = 0
    NPC(index).STR = 0
    NPC(index).DEF = 0
    NPC(index).Speed = 0
    NPC(index).Magi = 0
    NPC(index).Big = 0
    NPC(index).MAXHP = 0
    NPC(index).EXP = 0
    NPC(index).SpawnTime = 0
    NPC(index).Element = 0
    
    For i = 1 To MAX_NPC_DROPS
        NPC(index).ItemNPC(i).chance = 0
        NPC(index).ItemNPC(i).ItemNum = 0
        NPC(index).ItemNPC(i).ItemValue = 0
    Next i
    NPC(index).Quest = 1
End Sub

Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next i
End Sub

Sub ClearMapItem(ByVal index As Long, ByVal mapnum As Long)
    MapItem(mapnum, index).num = 0
    MapItem(mapnum, index).Value = 0
    MapItem(mapnum, index).Dur = 0
    MapItem(mapnum, index).x = 0
    MapItem(mapnum, index).y = 0
    Call SendDataToMap(mapnum, "spawnitem " & index & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
End Sub

Sub ClearMapItems()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next x
    Next y
End Sub

Sub ClearMapNpc(ByVal index As Long, ByVal mapnum As Long)
    MapNPC(mapnum, index).num = 0
    MapNPC(mapnum, index).Target = 0
    MapNPC(mapnum, index).HP = 0
    MapNPC(mapnum, index).MP = 0
    MapNPC(mapnum, index).SP = 0
    MapNPC(mapnum, index).x = 0
    MapNPC(mapnum, index).y = 0
    MapNPC(mapnum, index).Dir = 0

    ' Server use only
    MapNPC(mapnum, index).SpawnWait = 0
    MapNPC(mapnum, index).AttackTimer = 0
End Sub

Sub ClearMapNpcs()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next x
    Next y
End Sub

Sub ClearMap(ByVal mapnum As Long)
    Dim x As Long
    Dim y As Long

    Map(mapnum).Name = vbNullString
    Map(mapnum).Revision = 0
    Map(mapnum).Moral = 0
    Map(mapnum).Up = 0
    Map(mapnum).Down = 0
    Map(mapnum).Left = 0
    Map(mapnum).Right = 0
    Map(mapnum).Indoors = 0
    Map(mapnum).Weather = 0

    For x = 1 To MAX_MAP_NPCS
        Map(mapnum).NPC(x) = 0
    Next x

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            Map(mapnum).Tile(x, y).Ground = 0
            Map(mapnum).Tile(x, y).Mask = 0
            Map(mapnum).Tile(x, y).Anim = 0
            Map(mapnum).Tile(x, y).Mask2 = 0
            Map(mapnum).Tile(x, y).M2Anim = 0
            Map(mapnum).Tile(x, y).Fringe = 0
            Map(mapnum).Tile(x, y).FAnim = 0
            Map(mapnum).Tile(x, y).Fringe2 = 0
            Map(mapnum).Tile(x, y).F2Anim = 0
            Map(mapnum).Tile(x, y).Type = 0
            Map(mapnum).Tile(x, y).Data1 = 0
            Map(mapnum).Tile(x, y).Data2 = 0
            Map(mapnum).Tile(x, y).Data3 = 0
            Map(mapnum).Tile(x, y).String1 = vbNullString
            Map(mapnum).Tile(x, y).String2 = vbNullString
            Map(mapnum).Tile(x, y).String3 = vbNullString
            Map(mapnum).Tile(x, y).Light = 0
            Map(mapnum).Tile(x, y).GroundSet = 0
            Map(mapnum).Tile(x, y).MaskSet = 0
            Map(mapnum).Tile(x, y).AnimSet = 0
            Map(mapnum).Tile(x, y).Mask2Set = 0
            Map(mapnum).Tile(x, y).M2AnimSet = 0
            Map(mapnum).Tile(x, y).FringeSet = 0
            Map(mapnum).Tile(x, y).FAnimSet = 0
            Map(mapnum).Tile(x, y).Fringe2Set = 0
            Map(mapnum).Tile(x, y).F2AnimSet = 0
        Next x
    Next y

    ' Reset the values for if a player is on the map or not
    PlayersOnMap(mapnum) = NO

    ' Reset the map cache array for this map.
    MapCache(mapnum) = vbNullString
End Sub

Sub ClearMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next i
End Sub

Sub ClearShop(ByVal index As Long)
    Dim i As Long

    Shop(index).Name = vbNullString
    Shop(index).CurrencyItem = 1
    Shop(index).FixesItems = 0
    Shop(index).ShowInfo = 0
    For i = 1 To MAX_SHOP_ITEMS
        Shop(index).ShopItem(i).ItemNum = 0
        Shop(index).ShopItem(i).Amount = 0
        Shop(index).ShopItem(i).Price = 0
    Next i

End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next i
End Sub

Sub ClearSpell(ByVal index As Long)
    Spell(index).Name = vbNullString
    Spell(index).ClassReq = 0
    Spell(index).LevelReq = 0
    Spell(index).Type = 0
    Spell(index).Data1 = 0
    Spell(index).Data2 = 0
    Spell(index).Data3 = 0
    Spell(index).MPCost = 0
    Spell(index).Sound = 0
    Spell(index).Range = 0

    Spell(index).SpellAnim = 0
    Spell(index).SpellTime = 40
    Spell(index).SpellDone = 1

    Spell(index).AE = 0
    Spell(index).Big = 0

    Spell(index).Element = 0
End Sub

Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next i
End Sub

Function GetPlayerLogin(ByVal index As Long) As String
    GetPlayerLogin = Trim$(Player(index).Login)
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)
    Player(index).Login = Login
End Sub

Function GetPlayerPassword(ByVal index As Long) As String
    GetPlayerPassword = Trim$(Player(index).Password)
End Function

Sub SetPlayerPassword(ByVal index As Long, ByVal Password As String)
    Player(index).Password = Password
End Sub

Function GetPlayerName(ByVal index As Long) As String
    GetPlayerName = Trim$(Player(index).Char(Player(index).CharNum).Name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    Player(index).Char(Player(index).CharNum).Name = Name
End Sub

Function GetPlayerGuild(ByVal index As Long) As String
    GetPlayerGuild = Trim$(Player(index).Char(Player(index).CharNum).Guild)
End Function

Sub SetPlayerGuild(ByVal index As Long, ByVal Guild As String)
    Player(index).Char(Player(index).CharNum).Guild = Guild
End Sub

Function GetPlayerGuildAccess(ByVal index As Long) As Long
    GetPlayerGuildAccess = Player(index).Char(Player(index).CharNum).GuildAccess
End Function

Sub SetPlayerGuildAccess(ByVal index As Long, ByVal GuildAccess As Long)
    Player(index).Char(Player(index).CharNum).GuildAccess = GuildAccess
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
    GetPlayerClass = Player(index).Char(Player(index).CharNum).Class
End Function

Sub SetPlayerClassData(ByVal index As Long, ByVal ClassNum As Long)
    Player(index).Char(Player(index).CharNum).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long
    GetPlayerSprite = Player(index).Char(Player(index).CharNum).SPRITE
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal SPRITE As Long)
    If index > 0 And index <= MAX_PLAYERS Then
        Player(index).Char(Player(index).CharNum).SPRITE = SPRITE
    End If
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long
    GetPlayerLevel = Player(index).Char(Player(index).CharNum).Level
End Function

Sub SetPlayerLevel(ByVal index As Long, ByVal Level As Long)
    Player(index).Char(Player(index).CharNum).Level = Level
    
    If GetPlayerParty(index) > 0 Then
    Dim i As Long, N As Long, StatPercent As Long, Packet As String
    For i = 1 To MAX_PARTY_MEMBERS
    If Party(GetPlayerParty(index)).Member(i) = index Then
    StatPercent = GetPlayerLevel(index)
    Packet = "l" & SEP_CHAR & i & SEP_CHAR & StatPercent & SEP_CHAR & END_CHAR
    Call SendDataToParty(GetPlayerParty(index), Packet)
    End If
    Next i
    End If
End Sub

Function GetPlayerNextLevel(ByVal index As Long) As Long
    GetPlayerNextLevel = Experience(GetPlayerLevel(index))
End Function

Function GetPlayerExp(ByVal index As Long) As Long
    GetPlayerExp = Player(index).Char(Player(index).CharNum).EXP
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal EXP As Long)
    Player(index).Char(Player(index).CharNum).EXP = EXP
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long
    GetPlayerAccess = Player(index).Char(Player(index).CharNum).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
    Player(index).Char(Player(index).CharNum).Access = Access
End Sub

Function GetPlayerPK(ByVal index As Long) As Long
    GetPlayerPK = Player(index).Char(Player(index).CharNum).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
    Player(index).Char(Player(index).CharNum).PK = PK
End Sub

Function GetPlayerHP(ByVal index As Long) As Long
    GetPlayerHP = Player(index).Char(Player(index).CharNum).HP
End Function

Sub SetPlayerHP(ByVal index As Long, ByVal HP As Long)
    Player(index).Char(Player(index).CharNum).HP = HP

    If GetPlayerHP(index) > GetPlayerMaxHP(index) Then
        Player(index).Char(Player(index).CharNum).HP = GetPlayerMaxHP(index)
    End If
    If GetPlayerHP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).HP = 0
    End If
    'Call SendStats(Index)
    If GetPlayerParty(index) > 0 Then
    Dim i As Long, N As Long, StatPercent As Byte, Packet As String
    For i = 1 To MAX_PARTY_MEMBERS
    If Party(GetPlayerParty(index)).Member(i) = index Then
    StatPercent = Val(((GetPlayerHP(index) / 100) / (GetPlayerMaxHP(index) / 100)) * 42)
    Packet = "k" & SEP_CHAR & i & SEP_CHAR & StatPercent & SEP_CHAR & END_CHAR
    Call SendDataToParty(GetPlayerParty(index), Packet)
    End If
    Next i
    End If
End Sub

Function GetPlayerMP(ByVal index As Long) As Long
    GetPlayerMP = Player(index).Char(Player(index).CharNum).MP
End Function

Sub SetPlayerMP(ByVal index As Long, ByVal MP As Long)
    Player(index).Char(Player(index).CharNum).MP = MP

    If GetPlayerMP(index) > GetPlayerMaxMP(index) Then
        Player(index).Char(Player(index).CharNum).MP = GetPlayerMaxMP(index)
    End If
    If GetPlayerMP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).MP = 0
    End If
    
        If GetPlayerParty(index) > 0 Then
    Dim i As Long, N As Long, StatPercent As Byte, Packet As String
    For i = 1 To MAX_PARTY_MEMBERS
    If Party(GetPlayerParty(index)).Member(i) = index Then
    StatPercent = Val(((GetPlayerMP(index) / 100) / (GetPlayerMaxMP(index) / 100)) * 42)
    Packet = "m" & SEP_CHAR & i & SEP_CHAR & StatPercent & SEP_CHAR & END_CHAR
    Call SendDataToParty(GetPlayerParty(index), Packet)
    End If
    Next i
    End If
End Sub

Function GetPlayerSP(ByVal index As Long) As Long
    GetPlayerSP = Player(index).Char(Player(index).CharNum).SP
End Function

Sub SetPlayerSP(ByVal index As Long, ByVal SP As Long)
    Player(index).Char(Player(index).CharNum).SP = SP

    If GetPlayerSP(index) > GetPlayerMaxSP(index) Then
        Player(index).Char(Player(index).CharNum).SP = GetPlayerMaxSP(index)
    End If
    If GetPlayerSP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).SP = 0
    End If
End Sub

Function GetPlayerMaxHP(ByVal index As Long) As Long
    Dim CharNum As Long
    Dim Add As Long
    Add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        Add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).addHP
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).addHP
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).addHP
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).addHP
    End If
    If GetPlayerLegsSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).addHP
    End If
    If GetPlayerRingSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).addHP
    End If
    If GetPlayerNecklaceSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).addHP
    End If

    CharNum = Player(index).CharNum
    ' GetPlayerMaxHP = ((Player(index).Char(CharNum).Level + Int(GetPlayerSTR(index) / 2) + ClassData(Player(index).Char(CharNum).Class).STR) * 2) + add
    GetPlayerMaxHP = (GetPlayerLevel(index) * addHP.Level) + (GetPlayerSTR(index) * addHP.STR) + (GetPlayerDEF(index) * addHP.DEF) + (GetPlayerMAGI(index) * addHP.Magi) + (GetPlayerSPEED(index) * addHP.Speed) + Add
End Function

Function GetPlayerMaxMP(ByVal index As Long) As Long
    Dim CharNum As Long
    Dim Add As Long
    Add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        Add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).addMP
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).addMP
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).addMP
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).addMP
    End If
    If GetPlayerLegsSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).addMP
    End If
    If GetPlayerRingSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).addMP
    End If
    If GetPlayerNecklaceSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).addMP
    End If

    CharNum = Player(index).CharNum
    ' GetPlayerMaxMP = ((Player(index).Char(CharNum).Level + Int(GetPlayerMAGI(index) / 2) + ClassData(Player(index).Char(CharNum).Class).MAGI) * 2) + add
    GetPlayerMaxMP = (GetPlayerLevel(index) * addMP.Level) + (GetPlayerSTR(index) * addMP.STR) + (GetPlayerDEF(index) * addMP.DEF) + (GetPlayerMAGI(index) * addMP.Magi) + (GetPlayerSPEED(index) * addMP.Speed) + Add
End Function

Function GetPlayerMaxSP(ByVal index As Long) As Long
    Dim CharNum As Long
    Dim Add As Long
    Add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        Add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).addSP
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).addSP
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).addSP
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).addSP
    End If
    If GetPlayerLegsSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).addSP
    End If
    If GetPlayerRingSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).addSP
    End If
    If GetPlayerNecklaceSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).addSP
    End If

    CharNum = Player(index).CharNum
    ' GetPlayerMaxSP = ((Player(index).Char(CharNum).Level + Int(GetPlayerSPEED(index) / 2) + ClassData(Player(index).Char(CharNum).Class).SPEED) * 2) + add
    GetPlayerMaxSP = (GetPlayerLevel(index) * addSP.Level) + (GetPlayerSTR(index) * addSP.STR) + (GetPlayerDEF(index) * addSP.DEF) + (GetPlayerMAGI(index) * addSP.Magi) + (GetPlayerSPEED(index) * addSP.Speed) + Add
End Function

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(ClassData(ClassNum).Name)
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

Function GetPlayerSTR(ByVal index As Long) As Long
    Dim Add As Long
    Add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        Add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddStr
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddStr
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddStr
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddStr
    End If
    If GetPlayerLegsSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).AddStr
    End If
    If GetPlayerRingSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).AddStr
    End If
    If GetPlayerNecklaceSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).AddStr
    End If
    GetPlayerSTR = Player(index).Char(Player(index).CharNum).STR + Add
End Function

Sub SetPlayerSTR(ByVal index As Long, ByVal STR As Long)
    Player(index).Char(Player(index).CharNum).STR = STR
End Sub

Function GetPlayerDEF(ByVal index As Long) As Long
    Dim Add As Long
    Add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        Add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddDef
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddDef
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddDef
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddDef
    End If
    If GetPlayerLegsSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).AddDef
    End If
    If GetPlayerRingSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).AddDef
    End If
    If GetPlayerNecklaceSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).AddDef
    End If
    GetPlayerDEF = Player(index).Char(Player(index).CharNum).DEF + Add
End Function

Sub SetPlayerDEF(ByVal index As Long, ByVal DEF As Long)
    Player(index).Char(Player(index).CharNum).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal index As Long) As Long
    Dim Add As Long
    Add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        Add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddSpeed
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddSpeed
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddSpeed
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddSpeed
    End If
    If GetPlayerLegsSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).AddSpeed
    End If
    If GetPlayerRingSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).AddSpeed
    End If
    If GetPlayerNecklaceSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).AddSpeed
    End If
    GetPlayerSPEED = Player(index).Char(Player(index).CharNum).Speed + Add
End Function

Sub SetPlayerSPEED(ByVal index As Long, ByVal Speed As Long)
    Player(index).Char(Player(index).CharNum).Speed = Speed
End Sub

Function GetPlayerMAGI(ByVal index As Long) As Long
    Dim Add As Long
    Add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        Add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddMagi
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddMagi
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddMagi
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddMagi
    End If
    If GetPlayerLegsSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))).AddMagi
    End If
    If GetPlayerRingSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerRingSlot(index))).AddMagi
    End If
    If GetPlayerNecklaceSlot(index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))).AddMagi
    End If
    GetPlayerMAGI = Player(index).Char(Player(index).CharNum).Magi + Add
End Function

Sub SetPlayerMAGI(ByVal index As Long, ByVal Magi As Long)
    Player(index).Char(Player(index).CharNum).Magi = Magi
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long
    GetPlayerPOINTS = Player(index).Char(Player(index).CharNum).POINTS
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
    Player(index).Char(Player(index).CharNum).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal index As Long) As Long
    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerMap = Player(index).Char(Player(index).CharNum).Map
    End If
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal mapnum As Long)
    If mapnum > 0 And mapnum <= MAX_MAPS Then
        Player(index).Char(Player(index).CharNum).Map = mapnum
    End If
End Sub

Function GetPlayerX(ByVal index As Long) As Long
    GetPlayerX = Player(index).Char(Player(index).CharNum).x
End Function

Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
    Player(index).Char(Player(index).CharNum).x = x
End Sub

Function GetPlayerY(ByVal index As Long) As Long
    GetPlayerY = Player(index).Char(Player(index).CharNum).y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
        Player(index).Char(Player(index).CharNum).y = y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long
    GetPlayerDir = Player(index).Char(Player(index).CharNum).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
    Player(index).Char(Player(index).CharNum).Dir = Dir
End Sub

Function GetPlayerIP(ByVal index As Long) As String
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long
    If InvSlot > 0 Then
        GetPlayerInvItemNum = Player(index).Char(Player(index).CharNum).Inv(InvSlot).num
    End If
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(index).Char(Player(index).CharNum).Spell(SpellSlot)
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long, ByVal spellnum As Long)
    Player(index).Char(Player(index).CharNum).Spell(SpellSlot) = spellnum
End Sub

Function GetPlayerArmorSlot(ByVal index As Long) As Long
    GetPlayerArmorSlot = Player(index).Char(Player(index).CharNum).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal index As Long) As Long
    GetPlayerWeaponSlot = Player(index).Char(Player(index).CharNum).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal index As Long) As Long
    GetPlayerHelmetSlot = Player(index).Char(Player(index).CharNum).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal index As Long) As Long
    GetPlayerShieldSlot = Player(index).Char(Player(index).CharNum).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).ShieldSlot = InvNum
End Sub
Function GetPlayerLegsSlot(ByVal index As Long) As Long
    GetPlayerLegsSlot = Player(index).Char(Player(index).CharNum).LegsSlot
End Function

Sub SetPlayerLegsSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).LegsSlot = InvNum
End Sub
Function GetPlayerRingSlot(ByVal index As Long) As Long
    GetPlayerRingSlot = Player(index).Char(Player(index).CharNum).RingSlot
End Function

Sub SetPlayerRingSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).RingSlot = InvNum
End Sub
Function GetPlayerNecklaceSlot(ByVal index As Long) As Long
    GetPlayerNecklaceSlot = Player(index).Char(Player(index).CharNum).NecklaceSlot
End Function

Sub SetPlayerNecklaceSlot(ByVal index As Long, InvNum As Long)
    Player(index).Char(Player(index).CharNum).NecklaceSlot = InvNum
End Sub

Sub BattleMsg(ByVal index As Long, ByVal Msg As String, ByVal color As Byte, ByVal Side As Byte)
    Call SendDataTo(index, "damagedisplay" & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & color & END_CHAR)
End Sub

Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Player(index).Char(Player(index).CharNum).Bank(BankSlot).num
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
    Player(index).Char(Player(index).CharNum).Bank(BankSlot).num = ItemNum
    Call SendBankUpdate(index, BankSlot)
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Player(index).Char(Player(index).CharNum).Bank(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Player(index).Char(Player(index).CharNum).Bank(BankSlot).Value = ItemValue
    Call SendBankUpdate(index, BankSlot)
End Sub

Function GetPlayerBankItemDur(ByVal index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemDur = Player(index).Char(Player(index).CharNum).Bank(BankSlot).Dur
End Function

Sub SetPlayerBankItemDur(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemDur As Long)
    Player(index).Char(Player(index).CharNum).Bank(BankSlot).Dur = ItemDur
End Sub

Function GetPlayerHead(ByVal index As Long) As Integer
    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerHead = Player(index).Char(Player(index).CharNum).Head
    End If
End Function

Sub SetPlayerHead(ByVal index As Long, ByVal Head As Long)
    If index > 0 And index < MAX_PLAYERS Then
        Player(index).Char(Player(index).CharNum).Head = Head
    End If
End Sub

Function GetPlayerBody(ByVal index As Long) As Integer
    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerBody = Player(index).Char(Player(index).CharNum).Body
    End If
End Function

Sub SetPlayerBody(ByVal index As Long, ByVal Body As Long)
    If index > 0 And index < MAX_PLAYERS Then
        Player(index).Char(Player(index).CharNum).Body = Body
    End If
End Sub

Function GetPlayerleg(ByVal index As Long) As Integer
    If index > 0 And index < MAX_PLAYERS Then
        GetPlayerleg = Player(index).Char(Player(index).CharNum).Leg
    End If
End Function

Sub SetPlayerLeg(ByVal index As Long, ByVal Leg As Long)
    If index > 0 And index < MAX_PLAYERS Then
        Player(index).Char(Player(index).CharNum).Leg = Leg
    End If
End Sub

Function GetPlayerPaperdoll(ByVal index As Long) As Byte
    If index < MAX_PLAYERS And index > 0 Then
        If Player(index).InGame Then
            GetPlayerPaperdoll = Player(index).Char(Player(index).CharNum).paperdoll
        End If
    End If
End Function

Sub SetPlayerPaperdoll(ByVal index As Long, ByVal Mode As Byte)
    If index < MAX_PLAYERS And index > 0 Then
        If Mode = 0 Or Mode = 1 Then
            If Player(index).InGame Then
                Player(index).Char(Player(index).CharNum).paperdoll = Mode
            End If
        End If
    End If
End Sub

Function GetSpellReqLevel(ByVal spellnum As Long) As Long
    GetSpellReqLevel = Spell(spellnum).LevelReq
End Function

Function GetPlayerTargetNpc(ByVal index As Long) As Long
    GetPlayerTargetNpc = Player(index).TargetNPC
End Function

Public Sub SetPlayerColor(ByVal index As Long, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte)
If R < 0 Then Exit Sub
If R > 255 Then Exit Sub
If G < 0 Then Exit Sub
If G > 255 Then Exit Sub
If B < 0 Then Exit Sub
If B > 255 Then Exit Sub


Player(index).Char(Player(index).CharNum).color(1) = R
Player(index).Char(Player(index).CharNum).color(3) = B
Player(index).Char(Player(index).CharNum).color(2) = G

Call SendDataToMap(GetPlayerMap(index), "namecolor" & SEP_CHAR & index & SEP_CHAR & R & SEP_CHAR & G & SEP_CHAR & B & END_CHAR)

End Sub

Public Sub SendPlayerColor(ByVal index As Long)
    Call SendDataToMap(GetPlayerMap(index), "namecolor" & SEP_CHAR & index & SEP_CHAR & Player(index).Char(Player(index).CharNum).color(1) & SEP_CHAR & Player(index).Char(Player(index).CharNum).color(2) & SEP_CHAR & Player(index).Char(Player(index).CharNum).color(3) & END_CHAR)
End Sub
