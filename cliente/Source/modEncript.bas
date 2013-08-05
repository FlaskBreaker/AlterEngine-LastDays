Attribute VB_Name = "modEncript"
Option Explicit

'---------------------------------------------------------------------
' código Modificado por Harvey T. para encriptar y desencriptar cadenas
'---------------------------------------------------------------------

Function Encriptar( _
    UserKey As String, Text As String, Action As String _
    ) As String
    Dim UserKeyX As String
    Dim Temp     As Integer
    Dim Times    As Integer
    Dim I        As Integer
    Dim J        As Integer
    Dim n        As Integer
    Dim rtn      As String
    
    If Text = vbNullString Or UserKey = vbNullString Then
       Encriptar = vbNullString
       Exit Function
    End If
    '//Get UserKey characters
    n = Len(UserKey)
    ReDim UserKeyASCIIS(1 To n)
    For I = 1 To n
        UserKeyASCIIS(I) = Asc(Mid$(UserKey, I, 1))
    Next
        
    '//Get Text characters
    ReDim TextASCIIS(Len(Text)) As Integer
    For I = 1 To Len(Text)
        TextASCIIS(I) = Asc(Mid$(Text, I, 1))
    Next
    
    '//Encryption/Decryption
    If Action = "ENCRYPT" Then
       For I = 1 To Len(Text)
           J = IIf(J + 1 >= n, 1, J + 1)
           Temp = TextASCIIS(I) + UserKeyASCIIS(J)
           If Temp > 255 Then
              Temp = Temp - 255
           End If
           rtn = rtn + Chr$(Temp)
       Next
    ElseIf Action = "DECRYPT" Then
       For I = 1 To Len(Text)
           J = IIf(J + 1 >= n, 1, J + 1)
           Temp = TextASCIIS(I) - UserKeyASCIIS(J)
           If Temp < 0 Then
              Temp = Temp + 255
           End If
           rtn = rtn + Chr$(Temp)
       Next
    End If
    
    '//Return
    Encriptar = rtn
End Function

Sub EncryptarMap(ByVal MapNum As Long)
Dim MapIndex As Long
Dim x As Long
Dim y As Long
    Map2(MapNum).Name = Map(MapNum).Name
    Map2(MapNum).Name = Encriptar(SPassWord, "" & Map2(MapNum).Name, "ENCRYPT")
    Map2(MapNum).Revision = Map(MapNum).Revision
    Map2(MapNum).Revision = Encriptar(SPassWord, "" & Map2(MapNum).Revision, "ENCRYPT")
    Map2(MapNum).Moral = Map(MapNum).Moral
    Map2(MapNum).Moral = Encriptar(SPassWord, "" & Map2(MapNum).Moral, "ENCRYPT")
    Map2(MapNum).Up = Map(MapNum).Up
    Map2(MapNum).Up = Encriptar(SPassWord, "" & Map2(MapNum).Up, "ENCRYPT")
    Map2(MapNum).Down = Map(MapNum).Down
    Map2(MapNum).Down = Encriptar(SPassWord, "" & Map2(MapNum).Down, "ENCRYPT")
    Map2(MapNum).Left = Map(MapNum).Left
    Map2(MapNum).Left = Encriptar(SPassWord, "" & Map2(MapNum).Left, "ENCRYPT")
    Map2(MapNum).Right = Map(MapNum).Right
    Map2(MapNum).Right = Encriptar(SPassWord, "" & Map2(MapNum).Right, "ENCRYPT")
    Map2(MapNum).music = Map(MapNum).music
    Map2(MapNum).music = Encriptar(SPassWord, "" & Map2(MapNum).music, "ENCRYPT")
    Map2(MapNum).BootMap = Map(MapNum).BootMap
    Map2(MapNum).BootMap = Encriptar(SPassWord, "" & Map2(MapNum).BootMap, "ENCRYPT")
    Map2(MapNum).BootX = Map(MapNum).BootX
    Map2(MapNum).BootX = Encriptar(SPassWord, "" & Map2(MapNum).BootX, "ENCRYPT")
    Map2(MapNum).BootY = Map(MapNum).BootY
    Map2(MapNum).BootY = Encriptar(SPassWord, "" & Map2(MapNum).BootY, "ENCRYPT")
    Map2(MapNum).Indoors = Map(MapNum).Indoors
    Map2(MapNum).Indoors = Encriptar(SPassWord, "" & Map2(MapNum).Indoors, "ENCRYPT")
    Map2(MapNum).Weather = Map(MapNum).Weather
    Map2(MapNum).Weather = Encriptar(SPassWord, "" & Map2(MapNum).Weather, "ENCRYPT")
    MapIndex = MapIndex + 14

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            Map2(MapNum).Tile(x, y).Ground = Map(MapNum).Tile(x, y).Ground
            Map2(MapNum).Tile(x, y).Ground = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Ground, "ENCRYPT")
            Map2(MapNum).Tile(x, y).Mask = Map(MapNum).Tile(x, y).Mask
            Map2(MapNum).Tile(x, y).Mask = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Mask, "ENCRYPT")
            Map2(MapNum).Tile(x, y).Anim = Map(MapNum).Tile(x, y).Anim
            Map2(MapNum).Tile(x, y).Anim = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Anim, "ENCRYPT")
            Map2(MapNum).Tile(x, y).Mask2 = Map(MapNum).Tile(x, y).Mask2
            Map2(MapNum).Tile(x, y).Mask2 = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Mask2, "ENCRYPT")
            Map2(MapNum).Tile(x, y).M2Anim = Map(MapNum).Tile(x, y).M2Anim
            Map2(MapNum).Tile(x, y).M2Anim = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).M2Anim, "ENCRYPT")
            Map2(MapNum).Tile(x, y).Fringe = Map(MapNum).Tile(x, y).Fringe
            Map2(MapNum).Tile(x, y).Fringe = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Fringe, "ENCRYPT")
            Map2(MapNum).Tile(x, y).FAnim = Map(MapNum).Tile(x, y).FAnim
            Map2(MapNum).Tile(x, y).FAnim = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).FAnim, "ENCRYPT")
            Map2(MapNum).Tile(x, y).Fringe2 = Map(MapNum).Tile(x, y).Fringe2
            Map2(MapNum).Tile(x, y).Fringe2 = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Fringe2, "ENCRYPT")
            Map2(MapNum).Tile(x, y).F2Anim = Map(MapNum).Tile(x, y).F2Anim
            Map2(MapNum).Tile(x, y).FAnim = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).F2Anim, "ENCRYPT")
            Map2(MapNum).Tile(x, y).Type = Map(MapNum).Tile(x, y).Type
            Map2(MapNum).Tile(x, y).Type = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Type, "ENCRYPT")
            Map2(MapNum).Tile(x, y).Data1 = Map(MapNum).Tile(x, y).Data1
            Map2(MapNum).Tile(x, y).Data1 = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Data1, "ENCRYPT")
            Map2(MapNum).Tile(x, y).Data2 = Map(MapNum).Tile(x, y).Data2
            Map2(MapNum).Tile(x, y).Data2 = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Data2, "ENCRYPT")
            Map2(MapNum).Tile(x, y).Data3 = Map(MapNum).Tile(x, y).Data3
            Map2(MapNum).Tile(x, y).Data3 = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Data3, "ENCRYPT")
            Map2(MapNum).Tile(x, y).String1 = Map(MapNum).Tile(x, y).String1
            Map2(MapNum).Tile(x, y).String1 = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).String1, "ENCRYPT")
            Map2(MapNum).Tile(x, y).String2 = Map(MapNum).Tile(x, y).String2
            Map2(MapNum).Tile(x, y).String2 = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).String2, "ENCRYPT")
            Map2(MapNum).Tile(x, y).String3 = Map(MapNum).Tile(x, y).String3
            Map2(MapNum).Tile(x, y).String3 = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).String3, "ENCRYPT")
            Map2(MapNum).Tile(x, y).Light = Map(MapNum).Tile(x, y).Light
            Map2(MapNum).Tile(x, y).Light = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Light, "ENCRYPT")
            Map2(MapNum).Tile(x, y).GroundSet = Map(MapNum).Tile(x, y).GroundSet
            Map2(MapNum).Tile(x, y).GroundSet = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).GroundSet, "ENCRYPT")
            Map2(MapNum).Tile(x, y).MaskSet = Map(MapNum).Tile(x, y).MaskSet
            Map2(MapNum).Tile(x, y).MaskSet = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).MaskSet, "ENCRYPT")
            Map2(MapNum).Tile(x, y).AnimSet = Map(MapNum).Tile(x, y).AnimSet
            Map2(MapNum).Tile(x, y).AnimSet = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).AnimSet, "ENCRYPT")
            Map2(MapNum).Tile(x, y).Mask2Set = Map(MapNum).Tile(x, y).Mask2Set
            Map2(MapNum).Tile(x, y).Mask2Set = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Mask2Set, "ENCRYPT")
            Map2(MapNum).Tile(x, y).M2AnimSet = Map(MapNum).Tile(x, y).M2AnimSet
            Map2(MapNum).Tile(x, y).M2AnimSet = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).M2AnimSet, "ENCRYPT")
            Map2(MapNum).Tile(x, y).FringeSet = Map(MapNum).Tile(x, y).FringeSet
            Map2(MapNum).Tile(x, y).FringeSet = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).FringeSet, "ENCRYPT")
            Map2(MapNum).Tile(x, y).FAnimSet = Map(MapNum).Tile(x, y).FAnimSet
            Map2(MapNum).Tile(x, y).FAnimSet = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).FAnimSet, "ENCRYPT")
            Map2(MapNum).Tile(x, y).Fringe2Set = Map(MapNum).Tile(x, y).Fringe2Set
            Map2(MapNum).Tile(x, y).Fringe2Set = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Fringe2Set, "ENCRYPT")
            Map2(MapNum).Tile(x, y).F2AnimSet = Map(MapNum).Tile(x, y).F2AnimSet
            Map2(MapNum).Tile(x, y).F2AnimSet = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).F2AnimSet, "ENCRYPT")

            MapIndex = MapIndex + 26
        Next x
    Next y

    For x = 1 To MAX_MAP_NPCS
        Map2(MapNum).Npc(x) = Map(MapNum).Npc(x)
        Map2(MapNum).Npc(x) = Encriptar(SPassWord, "" & Map2(MapNum).Npc(x), "ENCRYPT")
        Map2(MapNum).SpawnX(x) = Map(MapNum).SpawnX(x)
        Map2(MapNum).SpawnX(x) = Encriptar(SPassWord, "" & Map2(MapNum).SpawnX(x), "ENCRYPT")
        Map2(MapNum).SpawnY(x) = Map(MapNum).SpawnY(x)
        Map2(MapNum).SpawnY(x) = Encriptar(SPassWord, "" & Map2(MapNum).SpawnY(x), "ENCRYPT")
        MapIndex = MapIndex + 3
    Next x
    
    
End Sub

Sub DesencryptarMap(ByVal MapNum As Long)
Dim MapIndex As Long
Dim x As Long
Dim y As Long
Dim I As Variant
    Map2(MapNum).Name = Encriptar(SPassWord, "" & Map2(MapNum).Name, "DECRYPT")
    Map(MapNum).Name = Map2(MapNum).Name
    Map2(MapNum).Revision = Encriptar(SPassWord, "" & Map2(MapNum).Revision, "DECRYPT")
    Map(MapNum).Revision = Map2(MapNum).Revision
    Map2(MapNum).Moral = Encriptar(SPassWord, "" & Map2(MapNum).Moral, "DECRYPT")
    Map(MapNum).Moral = Map2(MapNum).Moral
    Map2(MapNum).Up = Encriptar(SPassWord, "" & Map2(MapNum).Up, "DECRYPT")
    Map(MapNum).Up = Map2(MapNum).Up
    Map2(MapNum).Down = Encriptar(SPassWord, "" & Map2(MapNum).Down, "DECRYPT")
    Map(MapNum).Down = Map2(MapNum).Down
    Map2(MapNum).Left = Encriptar(SPassWord, "" & Map2(MapNum).Left, "DECRYPT")
    Map(MapNum).Left = Map2(MapNum).Left
    Map2(MapNum).Right = Encriptar(SPassWord, "" & Map2(MapNum).Right, "DECRYPT")
    Map(MapNum).Right = Map2(MapNum).Right
    Map2(MapNum).music = Encriptar(SPassWord, "" & Map2(MapNum).music, "DECRYPT")
    Map(MapNum).music = Map2(MapNum).music
    Map2(MapNum).BootMap = Encriptar(SPassWord, "" & Map2(MapNum).BootMap, "DECRYPT")
    Map(MapNum).BootMap = Map2(MapNum).BootMap
    Map2(MapNum).BootX = Encriptar(SPassWord, "" & Map2(MapNum).BootX, "DECRYPT")
    Map(MapNum).BootX = Map2(MapNum).BootX
    Map2(MapNum).BootY = Encriptar(SPassWord, "" & Map2(MapNum).BootY, "DECRYPT")
    Map(MapNum).BootY = Map2(MapNum).BootY
    Map2(MapNum).Indoors = Encriptar(SPassWord, "" & Map2(MapNum).Indoors, "DECRYPT")
    Map(MapNum).Indoors = Map2(MapNum).Indoors
    Map2(MapNum).Weather = Encriptar(SPassWord, "" & Map2(MapNum).Weather, "DECRYPT")
    Map(MapNum).Weather = Map2(MapNum).Weather
    MapIndex = MapIndex + 14

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            Map2(MapNum).Tile(x, y).Ground = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Ground, "DECRYPT")
            Map(MapNum).Tile(x, y).Ground = Map2(MapNum).Tile(x, y).Ground
            Map2(MapNum).Tile(x, y).Mask = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Mask, "DECRYPT")
            Map(MapNum).Tile(x, y).Mask = Map2(MapNum).Tile(x, y).Mask
            Map2(MapNum).Tile(x, y).Anim = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Anim, "DECRYPT")
            Map(MapNum).Tile(x, y).Anim = Map2(MapNum).Tile(x, y).Anim
            Map2(MapNum).Tile(x, y).Mask2 = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Mask2, "DECRYPT")
            Map(MapNum).Tile(x, y).Mask2 = Map2(MapNum).Tile(x, y).Mask2
            Map2(MapNum).Tile(x, y).M2Anim = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).M2Anim, "DECRYPT")
            Map(MapNum).Tile(x, y).M2Anim = Map2(MapNum).Tile(x, y).M2Anim
            Map2(MapNum).Tile(x, y).Fringe = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Fringe, "DECRYPT")
            Map(MapNum).Tile(x, y).Fringe = Map2(MapNum).Tile(x, y).Fringe
            Map2(MapNum).Tile(x, y).FAnim = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).FAnim, "DECRYPT")
            Map(MapNum).Tile(x, y).FAnim = Map2(MapNum).Tile(x, y).FAnim
            Map2(MapNum).Tile(x, y).Fringe2 = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Fringe2, "DECRYPT")
            Map(MapNum).Tile(x, y).Fringe2 = Map2(MapNum).Tile(x, y).Fringe2
            Map2(MapNum).Tile(x, y).FAnim = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).F2Anim, "DECRYPT")
            Map(MapNum).Tile(x, y).F2Anim = Map2(MapNum).Tile(x, y).F2Anim
            Map2(MapNum).Tile(x, y).Type = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Type, "DECRYPT")
            Map(MapNum).Tile(x, y).Type = Map2(MapNum).Tile(x, y).Type
            Map2(MapNum).Tile(x, y).Data1 = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Data1, "DECRYPT")
            Map(MapNum).Tile(x, y).Data1 = Map2(MapNum).Tile(x, y).Data1
            Map2(MapNum).Tile(x, y).Data2 = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Data2, "DECRYPT")
            Map(MapNum).Tile(x, y).Data2 = Map2(MapNum).Tile(x, y).Data2
            Map2(MapNum).Tile(x, y).Data3 = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Data3, "DECRYPT")
            Map(MapNum).Tile(x, y).Data3 = Map2(MapNum).Tile(x, y).Data3
            Map2(MapNum).Tile(x, y).String1 = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).String1, "DECRYPT")
            Map(MapNum).Tile(x, y).String1 = Map2(MapNum).Tile(x, y).String1
            Map2(MapNum).Tile(x, y).String2 = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).String2, "DECRYPT")
            Map(MapNum).Tile(x, y).String2 = Map2(MapNum).Tile(x, y).String2
            Map2(MapNum).Tile(x, y).String3 = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).String3, "DECRYPT")
            Map(MapNum).Tile(x, y).String3 = Map2(MapNum).Tile(x, y).String3
            Map2(MapNum).Tile(x, y).Light = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Light, "DECRYPT")
            Map(MapNum).Tile(x, y).Light = Map2(MapNum).Tile(x, y).Light
            Map2(MapNum).Tile(x, y).GroundSet = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).GroundSet, "DECRYPT")
            Map(MapNum).Tile(x, y).GroundSet = Map2(MapNum).Tile(x, y).GroundSet
            Map2(MapNum).Tile(x, y).MaskSet = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).MaskSet, "DECRYPT")
            Map(MapNum).Tile(x, y).MaskSet = Map2(MapNum).Tile(x, y).MaskSet
            Map2(MapNum).Tile(x, y).AnimSet = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).AnimSet, "DECRYPT")
            Map(MapNum).Tile(x, y).AnimSet = Map2(MapNum).Tile(x, y).AnimSet
            Map2(MapNum).Tile(x, y).Mask2Set = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Mask2Set, "DECRYPT")
            Map(MapNum).Tile(x, y).Mask2Set = Map2(MapNum).Tile(x, y).Mask2Set
            Map2(MapNum).Tile(x, y).M2AnimSet = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).M2AnimSet, "DECRYPT")
            Map(MapNum).Tile(x, y).M2AnimSet = Map2(MapNum).Tile(x, y).M2AnimSet
            Map2(MapNum).Tile(x, y).FringeSet = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).FringeSet, "DECRYPT")
            Map(MapNum).Tile(x, y).FringeSet = Map2(MapNum).Tile(x, y).FringeSet
            Map2(MapNum).Tile(x, y).FAnimSet = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).FAnimSet, "DECRYPT")
            Map(MapNum).Tile(x, y).FAnimSet = Map2(MapNum).Tile(x, y).FAnimSet
            Map2(MapNum).Tile(x, y).Fringe2Set = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).Fringe2Set, "DECRYPT")
            Map(MapNum).Tile(x, y).Fringe2Set = Map2(MapNum).Tile(x, y).Fringe2Set
            Map2(MapNum).Tile(x, y).F2AnimSet = Encriptar(SPassWord, "" & Map2(MapNum).Tile(x, y).F2AnimSet, "DECRYPT")
            Map(MapNum).Tile(x, y).F2AnimSet = Map2(MapNum).Tile(x, y).F2AnimSet

            MapIndex = MapIndex + 26
        Next x
    Next y

    For x = 1 To MAX_MAP_NPCS
        Map2(MapNum).Npc(x) = Encriptar(SPassWord, "" & Map2(MapNum).Npc(x), "DECRYPT")
        Map(MapNum).Npc(x) = Map2(MapNum).Npc(x)
        Map2(MapNum).SpawnX(x) = Encriptar(SPassWord, "" & Map2(MapNum).SpawnX(x), "DECRYPT")
        Map(MapNum).SpawnX(x) = Map2(MapNum).SpawnX(x)
        Map2(MapNum).SpawnY(x) = Encriptar(SPassWord, "" & Map2(MapNum).SpawnY(x), "DECRYPT")
        Map(MapNum).SpawnY(x) = Map2(MapNum).SpawnY(x)
        MapIndex = MapIndex + 3
    Next x
    
    

End Sub


