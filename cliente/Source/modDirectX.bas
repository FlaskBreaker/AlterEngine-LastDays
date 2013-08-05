Attribute VB_Name = "modDirectX"
Option Explicit

Public dx As DirectX7
Public DD As DirectDraw7

Public DD_Clip As DirectDrawClipper

Public DD_PrimarySurf As DirectDrawSurface7
Public DDSD_Primary As DDSURFACEDESC2

Public DD_SpriteSurf As DirectDrawSurface7
Public DDSD_Sprite As DDSURFACEDESC2

Public DD_ItemSurf As DirectDrawSurface7
Public DDSD_Item As DDSURFACEDESC2

Public DD_EmoticonSurf As DirectDrawSurface7
Public DDSD_Emoticon As DDSURFACEDESC2

Public DD_BackBuffer As DirectDrawSurface7
Public DDSD_BackBuffer As DDSURFACEDESC2

Public DD_BigSpriteSurf As DirectDrawSurface7
Public DDSD_BigSprite As DDSURFACEDESC2

Public DD_SpellAnim As DirectDrawSurface7
Public DDSD_SpellAnim As DDSURFACEDESC2

Public DD_BigSpellAnim As DirectDrawSurface7
Public DDSD_BigSpellAnim As DDSURFACEDESC2

Public DD_TileSurf(0 To ExtraSheets) As DirectDrawSurface7
Public DDSD_Tile(0 To ExtraSheets) As DDSURFACEDESC2
Public TileFile(0 To ExtraSheets) As Byte

Public DDSD_ArrowAnim As DDSURFACEDESC2
Public DD_ArrowAnim As DirectDrawSurface7

Public DD_player_head As DirectDrawSurface7
Public DDSD_player_head As DDSURFACEDESC2

Public DD_player_body As DirectDrawSurface7
Public DDSD_player_body As DDSURFACEDESC2

Public DD_player_legs As DirectDrawSurface7
Public DDSD_player_legs As DDSURFACEDESC2

Public DD_PetSpriteSurf As DirectDrawSurface7
Public DDSD_PetSprite As DDSURFACEDESC2

Public rec As RECT
Public rec_pos As RECT

Sub InitDirectX()
    On Error GoTo DXErr

    ' Initialize DirextX
    Set dx = New DirectX7

    ' Initialize DirectDraw
    Set DD = dx.DirectDrawCreate(vbNullString)

    ' Indicate windows mode application
    Call DD.SetCooperativeLevel(frmMirage.hWnd, DDSCL_NORMAL)

    ' Init type and get the primary surface
    DDSD_Primary.lFlags = DDSD_CAPS
    DDSD_Primary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_SYSTEMMEMORY
    Set DD_PrimarySurf = DD.CreateSurface(DDSD_Primary)

    ' Create the clipper
    Set DD_Clip = DD.CreateClipper(0)

    ' Associate the picture hwnd with the clipper
    DD_Clip.SetHWnd frmMirage.picScreen.hWnd

    ' Have the blits to the screen clipped to the picture box
    DD_PrimarySurf.SetClipper DD_Clip

    ' Initialize all surfaces
    Call InitSurfaces
    Exit Sub

    ' Error handling
DXErr:
    Call MsgBox("Error initializing DirectDraw! Make sure you have DirectX 7 or higher installed and a compatible graphics device. Err: " & Err.Number & ", Desc: " & Err.Description, vbCritical)
    Call GameDestroy
    End
End Sub

Sub InitSurfaces()
    Dim Key As DDCOLORKEY
    Dim I As Long
    Dim DC As Long

    ' Check for files existing
    If Not FileExists("\GFX\Sprites.bmp") Or Not FileExists("\GFX\Items.bmp") Or Not FileExists("\GFX\BigSprites.bmp") Or Not FileExists("\GFX\Emoticons.bmp") Or Not FileExists("\GFX\Arrows.bmp") Then
        Call MsgBox("Your missing some graphic files!", vbOKOnly, GAME_NAME)
        Call GameDestroy
    End If

    ' Set the key for masks
    Key.low = 0
    Key.high = 0

    ' Initialize back buffer
    DDSD_BackBuffer.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    DDSD_BackBuffer.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    DDSD_BackBuffer.lWidth = (MAX_MAPX + 1) * PIC_X
    DDSD_BackBuffer.lHeight = (MAX_MAPY + 1) * PIC_Y
    Set DD_BackBuffer = DD.CreateSurface(DDSD_BackBuffer)

    ' Init sprite ddsd type and load the bitmap
    DDSD_Sprite.lFlags = DDSD_CAPS
    DDSD_Sprite.ddsCaps.lCaps = DDSCAPS_SYSTEMMEMORY
    Set DD_SpriteSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\Sprites.bmp", DDSD_Sprite)
    SetMaskColorFromPixel DD_SpriteSurf, 0, 0
    
    ' Init sprite ddsd type and load the bitmap
    DDSD_PetSprite.lFlags = DDSD_CAPS
    DDSD_PetSprite.ddsCaps.lCaps = DDSCAPS_SYSTEMMEMORY
    Set DD_PetSpriteSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\Mascotas.bmp", DDSD_PetSprite)
    SetMaskColorFromPixel DD_PetSpriteSurf, 0, 0
    
    ' Init tiles ddsd type and load the bitmap
    For I = 0 To ExtraSheets
        If Dir$(App.Path & "\GFX\Tiles" & I & ".bmp") <> vbNullString Then
            DDSD_Tile(I).lFlags = DDSD_CAPS
            DDSD_Tile(I).ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
            Set DD_TileSurf(I) = DD.CreateSurfaceFromFile(App.Path & "\GFX\Tiles" & I & ".bmp", DDSD_Tile(I))
            SetMaskColorFromPixel DD_TileSurf(I), 0, 0
            TileFile(I) = 1
        Else
            TileFile(I) = 0
        End If
    Next I

    ' Init items ddsd type and load the bitmap
    DDSD_Item.lFlags = DDSD_CAPS
    DDSD_Item.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_ItemSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\Items.bmp", DDSD_Item)
    SetMaskColorFromPixel DD_ItemSurf, 0, 0
    
    ' Init big sprites ddsd type and load the bitmap
    DDSD_BigSprite.lFlags = DDSD_CAPS
    DDSD_BigSprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_BigSpriteSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\BigSprites.bmp", DDSD_BigSprite)
    SetMaskColorFromPixel DD_BigSpriteSurf, 0, 0

    ' Init emoticons ddsd type and load the bitmap
    DDSD_Emoticon.lFlags = DDSD_CAPS
    DDSD_Emoticon.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_EmoticonSurf = DD.CreateSurfaceFromFile(App.Path & "\GFX\Emoticons.bmp", DDSD_Emoticon)
    SetMaskColorFromPixel DD_EmoticonSurf, 0, 0

    ' Init spells ddsd type and load the bitmap
    DDSD_SpellAnim.lFlags = DDSD_CAPS
    DDSD_SpellAnim.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_SpellAnim = DD.CreateSurfaceFromFile(App.Path & "\GFX\Spells.bmp", DDSD_SpellAnim)
    SetMaskColorFromPixel DD_SpellAnim, 0, 0

    ' Init spells ddsd type and load the bitmap
    DDSD_BigSpellAnim.lFlags = DDSD_CAPS
    DDSD_BigSpellAnim.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_BigSpellAnim = DD.CreateSurfaceFromFile(App.Path & "\GFX\BigSpells.bmp", DDSD_BigSpellAnim)
    SetMaskColorFromPixel DD_BigSpellAnim, 0, 0

    ' Init arrows ddsd type and load the bitmap
    DDSD_ArrowAnim.lFlags = DDSD_CAPS
    DDSD_ArrowAnim.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
    Set DD_ArrowAnim = DD.CreateSurfaceFromFile(App.Path & "\GFX\Arrows.bmp", DDSD_ArrowAnim)
    SetMaskColorFromPixel DD_ArrowAnim, 0, 0

    If CustomPlayers <> 0 Then
        ' Init head ddsd type and load the bitmap
        DDSD_player_head.lFlags = DDSD_CAPS
        DDSD_player_head.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set DD_player_head = DD.CreateSurfaceFromFile(App.Path & "\GFX\heads.bmp", DDSD_player_head)
        SetMaskColorFromPixel DD_player_head, 0, 0

        ' Init body ddsd type and load the bitmap
        DDSD_player_body.lFlags = DDSD_CAPS
        DDSD_player_body.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set DD_player_body = DD.CreateSurfaceFromFile(App.Path & "\GFX\bodys.bmp", DDSD_player_body)
        SetMaskColorFromPixel DD_player_body, 0, 0

        ' Init legs ddsd type and load the bitmap
        DDSD_player_legs.lFlags = DDSD_CAPS
        DDSD_player_legs.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        Set DD_player_legs = DD.CreateSurfaceFromFile(App.Path & "\GFX\legs.bmp", DDSD_player_legs)
        SetMaskColorFromPixel DD_player_legs, 0, 0
    End If
End Sub

Sub DestroyDirectX()
    Dim I As Long

    Set dx = Nothing
    Set DD = Nothing

    Set DD_Clip = Nothing

    Set DD_PrimarySurf = Nothing
    Set DD_BackBuffer = Nothing

    Set DD_SpriteSurf = Nothing
    Set DD_PetSpriteSurf = Nothing

    For I = 0 To ExtraSheets
        If TileFile(I) = 1 Then
            Set DD_TileSurf(I) = Nothing
        End If
    Next I

    Set DD_ItemSurf = Nothing
    Set DD_BigSpriteSurf = Nothing
    Set DD_EmoticonSurf = Nothing
    Set DD_SpellAnim = Nothing
    Set DD_BigSpellAnim = Nothing
    Set DD_ArrowAnim = Nothing

    Set DD_player_head = Nothing
    Set DD_player_body = Nothing
    Set DD_player_legs = Nothing

End Sub

Function NeedToRestoreSurfaces() As Boolean
    Dim TestCoopRes As Long

    TestCoopRes = DD.TestCooperativeLevel

    If (TestCoopRes = DD_OK) Then
        NeedToRestoreSurfaces = False
    Else
        NeedToRestoreSurfaces = True
    End If
End Function

Public Sub SetMaskColorFromPixel(ByRef TheSurface As DirectDrawSurface7, ByVal X As Long, ByVal Y As Long)
    Dim TmpR As RECT
    Dim TmpDDSD As DDSURFACEDESC2
    Dim TmpColorKey As DDCOLORKEY

    With TmpR
        .Left = X
        .top = Y
        .Right = X
        .Bottom = Y
    End With

    TheSurface.Lock TmpR, TmpDDSD, DDLOCK_WAIT Or DDLOCK_READONLY, 0

    With TmpColorKey
        .low = TheSurface.GetLockedPixel(X, Y)
        .high = .low
    End With

    TheSurface.SetColorKey DDCKEY_SRCBLT, TmpColorKey

    TheSurface.Unlock TmpR
End Sub

Sub DisplayFx(ByRef surfDisplay As DirectDrawSurface7, intX As Long, intY As Long, intWidth As Long, intHeight As Long, lngROP As Long, blnFxCap As Boolean, Tile As Long)
    Dim lngSrcDC As Long
    Dim lngDestDC As Long

    lngDestDC = DD_BackBuffer.GetDC
    lngSrcDC = surfDisplay.GetDC
    BitBlt lngDestDC, intX, intY, intWidth, intHeight, lngSrcDC, (Tile - Int(Tile / TilesInSheets) * TilesInSheets) * PIC_X, Int(Tile / TilesInSheets) * PIC_Y, lngROP
    surfDisplay.ReleaseDC lngSrcDC
    DD_BackBuffer.ReleaseDC lngDestDC
End Sub

Public Function GetScreenLeft(ByVal index As Long) As Long
    GetScreenLeft = GetPlayerX(index) - 11
End Function

Public Function GetScreenTop(ByVal index As Long) As Long
    GetScreenTop = GetPlayerY(index) - 8
End Function

Public Function GetScreenRight(ByVal index As Long) As Long
    GetScreenRight = GetPlayerX(index) + 10
End Function

Public Function GetScreenBottom(ByVal index As Long) As Long
    GetScreenBottom = GetPlayerY(index) + 8
End Function

Sub Night()
    Dim X As Long, Y As Long

    If TileFile(10) = 0 Then
        Exit Sub
    End If

    For Y = ScreenY To ScreenY2
        For X = ScreenX To ScreenX2
            If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Light <= 0 Then
                DisplayFx DD_TileSurf(10), (X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, 31
            Else
                DisplayFx DD_TileSurf(10), (X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, 32, 32, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, CLng(Map(GetPlayerMap(MyIndex)).Tile(X, Y).Light)
            End If
        Next X
    Next Y
End Sub

Sub BltTile2(ByVal X As Long, ByVal Y As Long, ByVal Tile As Long)
    If TileFile(10) = 0 Then
        Exit Sub
    End If

    rec.top = Int(Tile / TilesInSheets) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (Tile - Int(Tile / TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) + sx - NewXOffset, Y - (NewPlayerY * PIC_Y) + sx - NewYOffset, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
'    DisplayFx DD_TileSurf(10), (x - NewPlayerX * PIC_X) + sx - NewXOffset, y - (NewPlayerY * PIC_Y) + sx - NewYOffset, 32, 16, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, Tile
End Sub

Sub BltPlayerText(ByVal index As Long)
    Dim TextX As Long
    Dim TextY As Long
    Dim intLoop As Integer
    Dim intLoop2 As Integer

    Dim bytLineCount As Byte
    Dim bytLineLength As Byte
    Dim strLine(0 To MAX_LINES - 1) As String
    Dim strWords() As String
    strWords() = Split(Bubble(index).Text, " ")

    If Len(Bubble(index).Text) < MAX_LINE_LENGTH Then
        DISPLAY_BUBBLE_WIDTH = 2 + ((Len(Bubble(index).Text) * 9) \ PIC_X)

        If DISPLAY_BUBBLE_WIDTH > MAX_BUBBLE_WIDTH Then
            DISPLAY_BUBBLE_WIDTH = MAX_BUBBLE_WIDTH
        End If
    Else
        DISPLAY_BUBBLE_WIDTH = 6
    End If

    TextX = GetPlayerX(index) * PIC_X + Player(index).XOffset + Int(PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2) - 6
    TextY = GetPlayerY(index) * PIC_Y + Player(index).YOffset - Int(PIC_Y) + 75

    Call DD_BackBuffer.ReleaseDC(TexthDC)

    ' Draw the fancy box with tiles.
    Call BltTile2(TextX - 10, TextY - 10, 1)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY - 10, 2)

    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)

        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY - 10, 16)
    Next intLoop

    TexthDC = DD_BackBuffer.GetDC

    ' Loop through all the words.
    For intLoop = 0 To UBound(strWords)
        ' Increment the line length.
        bytLineLength = bytLineLength + Len(strWords(intLoop)) + 1

        ' If we have room on the current line.
        If bytLineLength < MAX_LINE_LENGTH Then
            ' Add the text to the current line.
            strLine(bytLineCount) = strLine(bytLineCount) & strWords(intLoop) & " "
        Else
            bytLineCount = bytLineCount + 1

            If bytLineCount = MAX_LINES Then
                bytLineCount = bytLineCount - 1
                Exit For
            End If

            strLine(bytLineCount) = Trim$(strWords(intLoop)) & " "
            bytLineLength = 0
        End If
    Next intLoop

    Call DD_BackBuffer.ReleaseDC(TexthDC)

    If bytLineCount > 0 Then
        For intLoop = 6 To (bytLineCount - 2) * PIC_Y + 6
            Call BltTile2(TextX - 10, TextY - 10 + intLoop, 19)
            Call BltTile2(TextX - 10 + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X, TextY - 10 + intLoop, 17)

            For intLoop2 = 1 To DISPLAY_BUBBLE_WIDTH - 2
                Call BltTile2(TextX - 10 + (intLoop2 * PIC_X), TextY + intLoop - 10, 5)
            Next intLoop2
        Next intLoop
    End If

    Call BltTile2(TextX - 10, TextY + (bytLineCount * 4) - 4, 3)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY + (bytLineCount * 4) - 4, 4)

    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY + (bytLineCount * 16) - 4, 15)
    Next intLoop

    TexthDC = DD_BackBuffer.GetDC

    For intLoop = 0 To (MAX_LINES - 1)
        If strLine(intLoop) <> vbNullString Then
            Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) + sx - NewXOffset + (((DISPLAY_BUBBLE_WIDTH) * PIC_X) / 2) - ((Len(strLine(intLoop)) * 8) \ 2) - 4, TextY - (NewPlayerY * PIC_Y) + sx - NewYOffset, strLine(intLoop), QBColor(WHITE))
            TextY = TextY + 16
        End If
    Next intLoop
End Sub

Sub Bltscriptbubble(ByVal index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal Colour As Long)
    Dim TextX As Long
    Dim TextY As Long
    Dim intLoop As Integer
    Dim intLoop2 As Integer

    Dim bytLineCount As Byte
    Dim bytLineLength As Byte
    Dim strLine(0 To MAX_LINES - 1) As String
    Dim strWords() As String

    strWords() = Split(ScriptBubble(index).Text, " ")

    If Len(ScriptBubble(index).Text) < MAX_LINE_LENGTH Then
        DISPLAY_BUBBLE_WIDTH = 2 + ((Len(ScriptBubble(index).Text) * 9) \ PIC_X)

        If DISPLAY_BUBBLE_WIDTH > MAX_BUBBLE_WIDTH Then
            DISPLAY_BUBBLE_WIDTH = MAX_BUBBLE_WIDTH
        End If
    Else
        DISPLAY_BUBBLE_WIDTH = 6
    End If

    ' TextX = X * PIC_X + Int(PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2) - 6
    TextX = X * PIC_X - 22
    TextY = Y * PIC_Y - 22

    Call DD_BackBuffer.ReleaseDC(TexthDC)

    ' Draw the fancy box with tiles.
    Call BltTile2(TextX - 10, TextY - 10, 1)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY - 10, 2)

    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY - 10, 16)
    Next intLoop

    TexthDC = DD_BackBuffer.GetDC

    ' Loop through all the words.
    For intLoop = 0 To UBound(strWords)
        ' Increment the line length.
        bytLineLength = bytLineLength + Len(strWords(intLoop)) + 1

        ' If we have room on the current line.
        If bytLineLength < MAX_LINE_LENGTH Then
            ' Add the text to the current line.
            strLine(bytLineCount) = strLine(bytLineCount) & strWords(intLoop) & " "
        Else
            bytLineCount = bytLineCount + 1

            If bytLineCount = MAX_LINES Then
                bytLineCount = bytLineCount - 1
                Exit For
            End If

            strLine(bytLineCount) = Trim$(strWords(intLoop)) & " "
            bytLineLength = 0
        End If
    Next intLoop

    Call DD_BackBuffer.ReleaseDC(TexthDC)

    If bytLineCount > 0 Then
        For intLoop = 6 To (bytLineCount - 2) * PIC_Y + 6
            Call BltTile2(TextX - 10, TextY - 10 + intLoop, 19)
            Call BltTile2(TextX - 10 + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X, TextY - 10 + intLoop, 17)

            For intLoop2 = 1 To DISPLAY_BUBBLE_WIDTH - 2
                Call BltTile2(TextX - 10 + (intLoop2 * PIC_X), TextY + intLoop - 10, 5)
            Next intLoop2
        Next intLoop
    End If

    Call BltTile2(TextX - 10, TextY + (bytLineCount * 16) - 4, 3)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY + (bytLineCount * 16) - 4, 4)

    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY + (bytLineCount * 16) - 4, 15)
    Next intLoop

    TexthDC = DD_BackBuffer.GetDC

    For intLoop = 0 To (MAX_LINES - 1)
        If strLine(intLoop) <> vbNullString Then
            Call DrawText(TexthDC, TextX + (((DISPLAY_BUBBLE_WIDTH) * PIC_X) / 2) - ((Len(strLine(intLoop)) * 8) \ 2) - 7, TextY, strLine(intLoop), QBColor(Colour))
            TextY = TextY + 16
        End If
    Next intLoop
End Sub

Sub BltPlayerBars(ByVal index As Long)
    Dim X As Long, Y As Long

    X = (GetPlayerX(index) * PIC_X + sx + Player(index).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
    Y = (GetPlayerY(index) * PIC_Y + sx + Player(index).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset


    If Player(index).HP = 0 Then
        Exit Sub
    End If
    If SpriteSize = 1 Then
        ' draws the back bars
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(X, Y - 30, X + 32, Y - 34)

        ' draws HP
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        Call DD_BackBuffer.DrawBox(X, Y - 30, X + ((Player(index).HP / 100) / (Player(index).MaxHP / 100) * 32), Y - 34)
    Else
        If SpriteSize = 2 Then
            ' draws the back bars
            Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
            Call DD_BackBuffer.DrawBox(X, Y - 30 - PIC_Y, X + 32, Y - 34 - PIC_Y)

            ' draws HP
            Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
            Call DD_BackBuffer.DrawBox(X, Y - 30 - PIC_Y, X + ((Player(index).HP / 100) / (Player(index).MaxHP / 100) * 32), Y - 34 - PIC_Y)
        Else
            ' draws the back bars
            Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
            Call DD_BackBuffer.DrawBox(X, Y + 2, X + 32, Y - 2)

            ' draws HP
            Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
            Call DD_BackBuffer.DrawBox(X, Y + 2, X + ((Player(index).HP / 100) / (Player(index).MaxHP / 100) * 32), Y - 2)
        End If
    End If
End Sub
Sub BltSpellsBar(ByVal index As Long, ByVal Spellnum As Long)
        Dim value As Variant
        Dim Bar As String
        frmMirage.Timer1.Enabled = True
        frmMirage.Timer1.Interval = 1
End Sub

Sub BltNpcBars(ByVal index As Long)
    Dim X As Long, Y As Long

    On Error GoTo BltNpcBars_Error

    If MapNpc(index).HP = 0 Then
        Exit Sub
    End If
    If MapNpc(index).Num < 1 Then
        Exit Sub
    End If

    If Npc(MapNpc(index).Num).Big = 1 Then
        X = (MapNpc(index).X * PIC_X + sx - 9 + MapNpc(index).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
        Y = (MapNpc(index).Y * PIC_Y + sx + MapNpc(index).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset

        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(X, Y + 32, X + 50, Y + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        If MapNpc(index).MaxHP < 1 Then
            Call DD_BackBuffer.DrawBox(X, Y + 32, X + ((MapNpc(index).HP / 100) / ((MapNpc(index).MaxHP + 1) / 100) * 50), Y + 36)
        Else
            Call DD_BackBuffer.DrawBox(X, Y + 32, X + ((MapNpc(index).HP / 100) / (MapNpc(index).MaxHP / 100) * 50), Y + 36)
        End If
    Else
        X = (MapNpc(index).X * PIC_X + sx + MapNpc(index).XOffset) - (NewPlayerX * PIC_X) - NewXOffset
        Y = (MapNpc(index).Y * PIC_Y + sx + MapNpc(index).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset

        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(X, Y + 32, X + 32, Y + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))

        If MapNpc(index).MaxHP < 1 Then
            Call DD_BackBuffer.DrawBox(X, Y + 32, X + ((MapNpc(index).HP / 100) / ((MapNpc(index).MaxHP + 1) / 100) * 32), Y + 36)
        Else
            Call DD_BackBuffer.DrawBox(X, Y + 32, X + ((MapNpc(index).HP / 100) / (MapNpc(index).MaxHP / 100) * 32), Y + 36)
        End If

    End If


    On Error GoTo 0
    Exit Sub

BltNpcBars_Error:

    If Err.Number = DDERR_CANTCREATEDC Then

    End If

End Sub

Sub BltWeather()
    Dim I As Long

    Call DD_BackBuffer.SetForeColor(RGB(0, 0, 200))

    If GameWeather = WEATHER_RAINING Or GameWeather = WEATHER_THUNDER Then
        For I = 1 To MAX_RAINDROPS
            If DropRain(I).Randomized = False Then
                If frmMirage.tmrRainDrop.Enabled = False Then
                    BLT_RAIN_DROPS = 1
                    frmMirage.tmrRainDrop.Enabled = True
                    If frmMirage.tmrRainDrop.Tag = vbNullString Then
                        frmMirage.tmrRainDrop.Interval = 100
                        frmMirage.tmrRainDrop.Tag = "123"
                    End If
                End If
            End If
        Next I
    ElseIf GameWeather = WEATHER_SNOWING Then
        For I = 1 To MAX_RAINDROPS
            If DropSnow(I).Randomized = False Then
                If frmMirage.tmrSnowDrop.Enabled = False Then
                    BLT_SNOW_DROPS = 1
                    frmMirage.tmrSnowDrop.Enabled = True
                    If frmMirage.tmrSnowDrop.Tag = vbNullString Then
                        frmMirage.tmrSnowDrop.Interval = 200
                        frmMirage.tmrSnowDrop.Tag = "123"
                    End If
                End If
            End If
        Next I
    Else
        If BLT_RAIN_DROPS > 0 And BLT_RAIN_DROPS <= RainIntensity Then
            Call ClearRainDrop(BLT_RAIN_DROPS)
        End If
        frmMirage.tmrRainDrop.Tag = vbNullString
    End If

    For I = 1 To MAX_RAINDROPS
        If Not ((DropRain(I).X = 0) Or (DropRain(I).Y = 0)) Then
            rec.top = 0
            rec.Bottom = rec.top + PIC_Y
            rec.Left = 6 * PIC_X
            rec.Right = rec.Left + PIC_X
            DropRain(I).X = DropRain(I).X + DropRain(I).SPEED
            DropRain(I).Y = DropRain(I).Y + DropRain(I).SPEED
            Call DD_BackBuffer.BltFast(DropRain(I).X + DropRain(I).SPEED, DropRain(I).Y + DropRain(I).SPEED, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            If (DropRain(I).X > (MAX_MAPX + 1) * PIC_X) Or (DropRain(I).Y > (MAX_MAPY + 1) * PIC_Y) Then
                DropRain(I).Randomized = False
            End If
        End If
    Next I
    If TileFile(10) = 1 Then
        rec.top = Int(14 / TilesInSheets) * PIC_Y
        rec.Bottom = rec.top + PIC_Y
        rec.Left = (14 - Int(14 / TilesInSheets) * TilesInSheets) * PIC_X
        rec.Right = rec.Left + PIC_X
        For I = 1 To MAX_RAINDROPS
            If Not ((DropSnow(I).X = 0) Or (DropSnow(I).Y = 0)) Then
                DropSnow(I).X = DropSnow(I).X + DropSnow(I).SPEED
                DropSnow(I).Y = DropSnow(I).Y + DropSnow(I).SPEED
                Call DD_BackBuffer.BltFast(DropSnow(I).X + DropSnow(I).SPEED, DropSnow(I).Y + DropSnow(I).SPEED, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                If (DropSnow(I).X > (MAX_MAPX + 1) * PIC_X) Or (DropSnow(I).Y > (MAX_MAPY + 1) * PIC_Y) Then
                    DropSnow(I).Randomized = False
                End If
            End If
        Next I
    End If

    ' If it's thunder, make the screen randomly flash white
    If GameWeather = WEATHER_THUNDER Then
        If Int((100 - 1 + 1) * Rnd) + 1 = 8 Then
            DD_BackBuffer.SetFillColor RGB(255, 255, 255)

            Call DD_BackBuffer.DrawBox(0, 0, (MAX_MAPX + 1) * PIC_X, (MAX_MAPY + 1) * PIC_Y)
        End If
    End If
End Sub

Sub BltMapWeather()
    Dim I As Long

    Call DD_BackBuffer.SetForeColor(RGB(0, 0, 200))

    If Map(GetPlayerMap(MyIndex)).Weather = 1 Or Map(GetPlayerMap(MyIndex)).Weather = 3 Then
        For I = 1 To MAX_RAINDROPS
            If DropRain(I).Randomized = False Then
                If frmMirage.tmrRainDrop.Enabled = False Then
                    BLT_RAIN_DROPS = 1
                    frmMirage.tmrRainDrop.Enabled = True
                End If
            End If
        Next I
        For I = 1 To MAX_RAINDROPS
            If Not ((DropRain(I).X = 0) Or (DropRain(I).Y = 0)) Then
                rec.top = (14 - Int(14 / TilesInSheets)) * PIC_Y
                rec.Bottom = rec.top + PIC_Y
                rec.Left = 6 * PIC_X
                rec.Right = rec.Left + PIC_X
                DropRain(I).X = DropRain(I).X + DropRain(I).SPEED
                DropRain(I).Y = DropRain(I).Y + DropRain(I).SPEED
                Call DD_BackBuffer.BltFast(DropRain(I).X + DropRain(I).SPEED, DropRain(I).Y + DropRain(I).SPEED, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                If (DropRain(I).X > (MAX_MAPX + 1) * PIC_X) Or (DropRain(I).Y > (MAX_MAPY + 1) * PIC_Y) Then
                    DropRain(I).Randomized = False
                End If
            End If
        Next I

        If Map(GetPlayerMap(MyIndex)).Weather = 3 Then
            If Int((100 - 1 + 1) * Rnd) + 1 < 3 Then
                DD_BackBuffer.SetFillColor RGB(255, 255, 255)

                Call DD_BackBuffer.DrawBox(0, 0, (MAX_MAPX + 1) * PIC_X, (MAX_MAPY + 1) * PIC_Y)
            End If
        End If

    ElseIf Map(GetPlayerMap(MyIndex)).Weather = 2 Then
        For I = 1 To MAX_RAINDROPS
            If DropSnow(I).Randomized = False Then
                If frmMirage.tmrSnowDrop.Enabled = False Then
                    BLT_SNOW_DROPS = 1
                    frmMirage.tmrSnowDrop.Enabled = True
                End If
            End If
        Next I
        If TileFile(10) = 1 Then
            rec.top = Int(14 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.top + PIC_Y
            rec.Left = (14 - Int(14 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            For I = 1 To MAX_RAINDROPS
                If Not ((DropSnow(I).X = 0) Or (DropSnow(I).Y = 0)) Then
                    DropSnow(I).X = DropSnow(I).X + DropSnow(I).SPEED
                    DropSnow(I).Y = DropSnow(I).Y + DropSnow(I).SPEED
                    Call DD_BackBuffer.BltFast(DropSnow(I).X + DropSnow(I).SPEED, DropSnow(I).Y + DropSnow(I).SPEED, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    If (DropSnow(I).X > (MAX_MAPX + 1) * PIC_X) Or (DropSnow(I).Y > (MAX_MAPY + 1) * PIC_Y) Then
                        DropSnow(I).Randomized = False
                    End If
                End If
            Next I
        End If
    Else
        If BLT_RAIN_DROPS > 0 And BLT_RAIN_DROPS <= frmMirage.tmrRainDrop.Interval Then
            Call ClearRainDrop(BLT_RAIN_DROPS)
        End If
        frmMirage.tmrRainDrop.Tag = vbNullString
    End If
End Sub

Sub RNDRainDrop(ByVal RDNumber As Long)
Start:
    DropRain(RDNumber).X = Int((((MAX_MAPX + 1) * PIC_X) * Rnd) + 1)
    DropRain(RDNumber).Y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd) + 1)
    If (DropRain(RDNumber).Y > (MAX_MAPY + 1) * PIC_Y / 4) And (DropRain(RDNumber).X > (MAX_MAPX + 1) * PIC_X / 4) Then
        GoTo Start
    End If
    DropRain(RDNumber).SPEED = Int((10 * Rnd) + 6)
    DropRain(RDNumber).Randomized = True
End Sub

Sub ClearRainDrop(ByVal RDNumber As Long)
    On Error Resume Next
    DropRain(RDNumber).X = 0
    DropRain(RDNumber).Y = 0
    DropRain(RDNumber).SPEED = 0
    DropRain(RDNumber).Randomized = False
End Sub

Sub RNDSnowDrop(ByVal RDNumber As Long)
Start:
    DropSnow(RDNumber).X = Int((((MAX_MAPX + 1) * PIC_X) * Rnd) + 1)
    DropSnow(RDNumber).Y = Int((((MAX_MAPY + 1) * PIC_Y) * Rnd) + 1)
    If (DropSnow(RDNumber).Y > (MAX_MAPY + 1) * PIC_Y / 4) And (DropSnow(RDNumber).X > (MAX_MAPX + 1) * PIC_X / 4) Then
        GoTo Start
    End If
    DropSnow(RDNumber).SPEED = Int((10 * Rnd) + 6)
    DropSnow(RDNumber).Randomized = True
End Sub

Sub ClearSnowDrop(ByVal RDNumber As Long)
    On Error Resume Next
    DropSnow(RDNumber).X = 0
    DropSnow(RDNumber).Y = 0
    DropSnow(RDNumber).SPEED = 0
    DropSnow(RDNumber).Randomized = False
End Sub

Sub BltSpell(ByVal index As Long)
    Dim X As Long, Y As Long, I As Long

    If Player(index).Spellnum <= 0 Or Player(index).Spellnum > MAX_SPELLS Then
        Exit Sub
    End If


    For I = 1 To MAX_SPELL_ANIM
        ' IF SPELL IS NOT BIG
        If Spell(Player(index).Spellnum).Big = 0 Then
            If Player(index).SpellAnim(I).CastedSpell = YES Then
                If Player(index).SpellAnim(I).SpellDone < Spell(Player(index).Spellnum).SpellDone Then

                    rec.top = Spell(Player(index).Spellnum).SpellAnim * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    rec.Left = Player(index).SpellAnim(I).SpellVar * PIC_X
                    rec.Right = rec.Left + PIC_X

                    If Player(index).SpellAnim(I).TargetType = 0 Then

                        ' SMALL: IF TARGET IS A PLAYER
                        If Player(index).SpellAnim(I).Target > 0 Then

                            ' SMALL: IF TARGET IS SELF
                            If Player(index).SpellAnim(I).Target = MyIndex Then
                                X = NewX + sx
                                Y = NewY + sx
                                Call DD_BackBuffer.BltFast(X, Y, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

                            ' SMALL: IF TARGET IS ANOTHER PLAYER
                            Else
                                X = GetPlayerX(Player(index).SpellAnim(I).Target) * PIC_X + sx + Player(Player(index).SpellAnim(I).Target).XOffset
                                Y = GetPlayerY(Player(index).SpellAnim(I).Target) * PIC_Y + sx + Player(Player(index).SpellAnim(I).Target).YOffset
                                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                            End If
                        End If

                    ' SMALL: IF TARGET IS AN NPC
                    Else
                        X = MapNpc(Player(index).SpellAnim(I).Target).X * PIC_X + sx + MapNpc(Player(index).SpellAnim(I).Target).XOffset
                        Y = MapNpc(Player(index).SpellAnim(I).Target).Y * PIC_Y + sx + MapNpc(Player(index).SpellAnim(I).Target).YOffset
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If


' SMALL: ADVANCE SPELL ONE CYCLE

                    If GetTickCount > Player(index).SpellAnim(I).SpellTime + Spell(Player(index).Spellnum).SpellTime Then
                        Player(index).SpellAnim(I).SpellTime = GetTickCount
                        Player(index).SpellAnim(I).SpellVar = Player(index).SpellAnim(I).SpellVar + 1
                    End If

                    If Player(index).SpellAnim(I).SpellVar > 12 Then
                        Player(index).SpellAnim(I).SpellDone = Player(index).SpellAnim(I).SpellDone + 1
                        Player(index).SpellAnim(I).SpellVar = 0
                    End If

                Else
                    Player(index).SpellAnim(I).CastedSpell = NO
                End If
            End If
        Else
            If Player(index).SpellAnim(I).CastedSpell = YES Then
                If Player(index).SpellAnim(I).SpellDone < Spell(Player(index).Spellnum).SpellDone Then

                    rec.top = Spell(Player(index).Spellnum).SpellAnim * (PIC_Y * 3)
                    rec.Bottom = rec.top + PIC_Y + 64
                    rec.Left = Player(index).SpellAnim(I).SpellVar * PIC_X
                    rec.Right = rec.Left + PIC_X + 64

                    If Player(index).SpellAnim(I).TargetType = 0 Then

                        ' BIG: IF TARGET IS A PLAYER
                        If Player(index).SpellAnim(I).Target > 0 Then

                            ' BIG: IF TARGET IS SELF
                            If Player(index).SpellAnim(I).Target = MyIndex Then
                                X = NewX + sx - 32
                                Y = NewY + sx - 32

                                If Y < 0 Then
                                    rec.top = rec.top + (Y * -1)
                                    Y = 0
                                End If

                                If X < 0 Then
                                    rec.Left = rec.Left + (X * -1)
                                    X = 0
                                End If

                                If (X + 64) > (MAX_MAPX * 32) Then
                                    rec.Right = rec.Left + 64
                                End If

                                If (Y + 64) > (MAX_MAPY * 32) Then
                                    rec.Bottom = rec.top + 64
                                End If

                                Call DD_BackBuffer.BltFast(X, Y, DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

                            ' BIG: IF TARGET IS A DIFFERENT PLAYER
                            Else
                                X = GetPlayerX(Player(index).SpellAnim(I).Target) * PIC_X + sx - 32 + Player(Player(index).SpellAnim(I).Target).XOffset
                                Y = GetPlayerY(Player(index).SpellAnim(I).Target) * PIC_Y + sx - 32 + Player(Player(index).SpellAnim(I).Target).YOffset

                                If Y < 0 Then
                                    rec.top = rec.top + (Y * -1)
                                    Y = 0
                                End If

                                If X < 0 Then
                                    rec.Left = rec.Left + (X * -1)
                                    X = 0
                                End If

                                If (X + 64) > (MAX_MAPX * 32) Then
                                    rec.Right = rec.Left + 64
                                End If

                                If (Y + 64) > (MAX_MAPY * 32) Then
                                    rec.Bottom = rec.top + 64
                                End If

                                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                            End If
                        End If

                    ' BIG: IF TARGET IS AN NPC
                    Else
                        X = MapNpc(Player(index).SpellAnim(I).Target).X * PIC_X + sx - 32 + MapNpc(Player(index).SpellAnim(I).Target).XOffset
                        Y = MapNpc(Player(index).SpellAnim(I).Target).Y * PIC_Y + sx - 32 + MapNpc(Player(index).SpellAnim(I).Target).YOffset

                        If Y < 0 Then
                            rec.top = rec.top + (Y * -1)
                            Y = 0
                        End If

                        If X < 0 Then
                            rec.Left = rec.Left + (X * -1)
                            X = 0
                        End If

                        If (X + 64) > (MAX_MAPX * 32) Then
                            rec.Right = rec.Left + 64
                        End If

                        If (Y + 64) > (MAX_MAPY * 32) Then
                            rec.Bottom = rec.top + 64
                        End If

                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' BIG: ADVANCE SPELL ONE CYCLE
                    If GetTickCount > Player(index).SpellAnim(I).SpellTime + Spell(Player(index).Spellnum).SpellTime Then
                        Player(index).SpellAnim(I).SpellTime = GetTickCount
                        Player(index).SpellAnim(I).SpellVar = Player(index).SpellAnim(I).SpellVar + 3
                    End If

                    If Player(index).SpellAnim(I).SpellVar > 36 Then
                        Player(index).SpellAnim(I).SpellDone = Player(index).SpellAnim(I).SpellDone + 1
                        Player(index).SpellAnim(I).SpellVar = 0
                    End If

                Else
                    Player(index).SpellAnim(I).CastedSpell = NO
                End If
            End If
        End If
    Next I
End Sub

' Scripted Spell
Sub BltScriptSpell(ByVal I As Long)
    Dim rec As RECT
    Dim X As Long, Y As Long

    X = ScriptSpell(I).X
    Y = ScriptSpell(I).Y

    If Spell(ScriptSpell(I).Spellnum).Big = 0 Then
        If ScriptSpell(I).SpellDone < Spell(ScriptSpell(I).Spellnum).SpellDone Then
            rec.top = Spell(ScriptSpell(I).Spellnum).SpellAnim * PIC_Y
            rec.Bottom = rec.top + PIC_Y
            rec.Left = ScriptSpell(I).SpellVar * PIC_X
            rec.Right = rec.Left + PIC_X

            X = X * PIC_X + sx
            Y = Y * PIC_Y + sx

            If ScriptSpell(I).SpellVar > 10 Then
                ScriptSpell(I).SpellDone = ScriptSpell(I).SpellDone + 1
                ScriptSpell(I).SpellVar = 0
            End If

            If GetTickCount > ScriptSpell(I).SpellTime + Spell(ScriptSpell(I).Spellnum).SpellTime Then
                ScriptSpell(I).SpellTime = GetTickCount
                ScriptSpell(I).SpellVar = ScriptSpell(I).SpellVar + 1
            End If

            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X), Y - (NewPlayerY * PIC_Y), DD_SpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

        Else ' spell is done
            ScriptSpell(I).CastedSpell = NO
        End If
    Else
        If ScriptSpell(I).SpellDone < Spell(ScriptSpell(I).Spellnum).SpellDone Then
            rec.top = Spell(ScriptSpell(I).Spellnum).SpellAnim * (PIC_Y * 3)
            rec.Bottom = rec.top + PIC_Y + 64
            rec.Left = ScriptSpell(I).SpellVar * PIC_X
            rec.Right = rec.Left + PIC_X + 64

            X = X * PIC_X + sx - 32
            Y = Y * PIC_Y + sx - 32

            If Y < 0 Then
                rec.top = rec.top + (Y * -1)
                Y = 0
            End If

            If X < 0 Then
                rec.Left = rec.Left + (X * -1)
                X = 0
            End If

            If (X + 64) > (MAX_MAPX * 32) Then
                rec.Right = rec.Left + 64
            End If

            If (Y + 64) > (MAX_MAPY * 32) Then
                rec.Bottom = rec.top + 64
            End If

            If ScriptSpell(I).SpellVar > 30 Then
                ScriptSpell(I).SpellDone = ScriptSpell(I).SpellDone + 1
                ScriptSpell(I).SpellVar = 0
            End If

            If GetTickCount > ScriptSpell(I).SpellTime + Spell(ScriptSpell(I).Spellnum).SpellTime Then
                ScriptSpell(I).SpellTime = GetTickCount
                ScriptSpell(I).SpellVar = ScriptSpell(I).SpellVar + 3
            End If

            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X), Y - (NewPlayerY * PIC_Y), DD_BigSpellAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else 'spell is done
            ScriptSpell(I).CastedSpell = NO
        End If
    End If
End Sub

Sub BltEmoticons(ByVal index As Long)
    Dim X2 As Long, Y2 As Long
    Dim ETime As Long
    ETime = 1300

    If Player(index).EmoticonNum < 0 Then
        Exit Sub
    End If

    If Player(index).EmoticonTime + ETime > GetTickCount Then
        If GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 1) Then
            Player(index).EmoticonVar = 0
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 2) Then
            Player(index).EmoticonVar = 1
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 3) Then
            Player(index).EmoticonVar = 2
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 4) Then
            Player(index).EmoticonVar = 3
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 5) Then
            Player(index).EmoticonVar = 4
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 6) Then
            Player(index).EmoticonVar = 5
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 7) Then
            Player(index).EmoticonVar = 6
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 8) Then
            Player(index).EmoticonVar = 7
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 9) Then
            Player(index).EmoticonVar = 8
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 10) Then
            Player(index).EmoticonVar = 9
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 11) Then
            Player(index).EmoticonVar = 10
        ElseIf GetTickCount < Player(index).EmoticonTime + Int((ETime / 12) * 12) Then
            Player(index).EmoticonVar = 11
        End If

        rec.top = Player(index).EmoticonNum * PIC_Y
        rec.Bottom = rec.top + PIC_Y
        rec.Left = Player(index).EmoticonVar * PIC_X
        rec.Right = rec.Left + PIC_X

        If index = MyIndex Then
            X2 = NewX + sx + 16
            Y2 = NewY + sx - 32

            If Y2 < 0 Then
                Exit Sub
            End If

            Call DD_BackBuffer.BltFast(X2, Y2, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else
            X2 = GetPlayerX(index) * PIC_X + sx + Player(index).XOffset + 16
            Y2 = GetPlayerY(index) * PIC_Y + sx + Player(index).YOffset - 32

            If Y2 < 0 Then
                Exit Sub
            End If

            Call DD_BackBuffer.BltFast(X2 - (NewPlayerX * PIC_X) - NewXOffset, Y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub Bltgrapple(ByVal index As Long)
    Dim z As Integer
    Dim BX As Long, BY As Long

    If Player(index).HookShotX > 0 Or Player(index).HookShotY <> 0 Then

        Select Case Player(index).HookShotDir
            Case 0
                z = 1
            Case 1
                z = 0
            Case 2
                z = 3
            Case 3
                z = 2
        End Select

        rec.top = Player(index).HookShotAnim * PIC_Y
        rec.Bottom = rec.top + PIC_Y
        rec.Left = z * PIC_X
        rec.Right = rec.Left + PIC_X

        If GetTickCount > Player(index).HookShotTime + 50 Then
            If Player(index).HookShotSucces = 1 Then
                If index = MyIndex Then
                Call SendData("endshot" & SEP_CHAR & 1 & END_CHAR)
                End If
                Player(index).HookShotX = 0
                Player(index).HookShotY = 0
            Else
                If index = MyIndex Then
                Call SendData("endshot" & SEP_CHAR & 0 & END_CHAR)
                End If
                Player(index).HookShotX = 0
                Player(index).HookShotY = 0
            End If
        End If

        BX = GetPlayerX(index)
        BY = GetPlayerY(index)

        If Player(index).HookShotDir = DIR_DOWN Then
            Do While BY <= Player(index).HookShotToY
                If BY <= MAX_MAPY Then
                    Call DD_BackBuffer.BltFast((BX - NewPlayerX) * PIC_X + sx - NewXOffset, (BY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                BY = BY + 1
            Loop
        End If

        If Player(index).HookShotDir = DIR_UP Then
            Do While BY >= Player(index).HookShotToY
                If BY >= 0 Then
                    Call DD_BackBuffer.BltFast((BX - NewPlayerX) * PIC_X + sx - NewXOffset, (BY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                BY = BY - 1
            Loop
        End If

        If Player(index).HookShotDir = DIR_RIGHT Then
            Do While BX <= Player(index).HookShotToX
                If BX <= MAX_MAPX Then
                    Call DD_BackBuffer.BltFast((BX - NewPlayerX) * PIC_X + sx - NewXOffset, (BY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                BX = BX + 1
            Loop
        End If

        If Player(index).HookShotDir = DIR_LEFT Then
            Do While BX >= Player(index).HookShotToX
                If BX >= 0 Then
                    Call DD_BackBuffer.BltFast((BX - NewPlayerX) * PIC_X + sx - NewXOffset, (BY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                BX = BX - 1
            Loop
        End If
    End If
End Sub

Sub BltArrow(ByVal index As Long)
    Dim X As Long
    Dim Y As Long
    Dim I As Long
    Dim z As Long

    For z = 1 To MAX_PLAYER_ARROWS
        If Player(index).Arrow(z).Arrow > 0 Then
            rec.top = Player(index).Arrow(z).ArrowAnim * PIC_Y
            rec.Bottom = rec.top + PIC_Y
            rec.Left = Player(index).Arrow(z).ArrowPosition * PIC_X
            rec.Right = rec.Left + PIC_X

            If GetTickCount > Player(index).Arrow(z).ArrowTime + 30 Then
                Player(index).Arrow(z).ArrowTime = GetTickCount
                Player(index).Arrow(z).ArrowVarX = Player(index).Arrow(z).ArrowVarX + 10
                Player(index).Arrow(z).ArrowVarY = Player(index).Arrow(z).ArrowVarY + 10
            End If

            If Player(index).Arrow(z).ArrowPosition = 0 Then
                X = Player(index).Arrow(z).ArrowX
                Y = Player(index).Arrow(z).ArrowY + Int(Player(index).Arrow(z).ArrowVarY / 32)

                If Y > Player(index).Arrow(z).ArrowY + Arrows(Player(index).Arrow(z).ArrowNum).Range - 2 Then
                    Player(index).Arrow(z).Arrow = 0
                End If

                If Y <= MAX_MAPY Then
                    Call DD_BackBuffer.BltFast((Player(index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset + Player(index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

            If Player(index).Arrow(z).ArrowPosition = 1 Then
                X = Player(index).Arrow(z).ArrowX
                Y = Player(index).Arrow(z).ArrowY - Int(Player(index).Arrow(z).ArrowVarY / 32)

                If Y < Player(index).Arrow(z).ArrowY - Arrows(Player(index).Arrow(z).ArrowNum).Range + 2 Then
                    Player(index).Arrow(z).Arrow = 0
                End If

                If Y >= 0 Then
                    Call DD_BackBuffer.BltFast((Player(index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset - Player(index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

            If Player(index).Arrow(z).ArrowPosition = 2 Then
                X = Player(index).Arrow(z).ArrowX + Int(Player(index).Arrow(z).ArrowVarX / 32)
                Y = Player(index).Arrow(z).ArrowY

                If X > Player(index).Arrow(z).ArrowX + Arrows(Player(index).Arrow(z).ArrowNum).Range - 2 Then
                    Player(index).Arrow(z).Arrow = 0
                End If

                If X <= MAX_MAPX Then
                    Call DD_BackBuffer.BltFast((Player(index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset + Player(index).Arrow(z).ArrowVarX, (Player(index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

            If Player(index).Arrow(z).ArrowPosition = 3 Then
                X = Player(index).Arrow(z).ArrowX - Int(Player(index).Arrow(z).ArrowVarX / 32)
                Y = Player(index).Arrow(z).ArrowY

                If X < Player(index).Arrow(z).ArrowX - Arrows(Player(index).Arrow(z).ArrowNum).Range + 2 Then
                    Player(index).Arrow(z).Arrow = 0
                End If

                If X >= 0 Then
                    Call DD_BackBuffer.BltFast((Player(index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset - Player(index).Arrow(z).ArrowVarX, (Player(index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

            If X >= 0 And X <= MAX_MAPX Then
                If Y >= 0 And Y <= MAX_MAPY Then
                    If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
                        Player(index).Arrow(z).Arrow = 0
                    End If
                End If
            End If

            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) Then
                    If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                        If GetPlayerX(I) = X Then
                            If GetPlayerY(I) = Y Then
                                If index = MyIndex Then
                                    Call SendData("arrowhit" & SEP_CHAR & 0 & SEP_CHAR & I & SEP_CHAR & X & SEP_CHAR & Y & END_CHAR)
                                End If

                                If index <> I Then
                                    Player(index).Arrow(z).Arrow = 0
                                End If

                                Exit Sub
                            End If
                        End If
                    End If
                End If
            Next I

            For I = 1 To MAX_MAP_NPCS
                If MapNpc(I).Num > 0 Then
                    If MapNpc(I).X = X Then
                        If MapNpc(I).Y = Y Then
                            If index = MyIndex Then
                                Call SendData("arrowhit" & SEP_CHAR & 1 & SEP_CHAR & I & SEP_CHAR & X & SEP_CHAR & Y & END_CHAR)
                            End If

                            Player(index).Arrow(z).Arrow = 0

                            Exit Sub
                        End If
                    End If
                End If
            Next I
        End If
    Next z
End Sub

Sub BltLevelUp(ByVal index As Long)
    Dim rec As RECT
    Dim X As Integer
    Dim Y As Integer

    If Player(index).LevelUpT + 3000 > GetTickCount Then
        If GetPlayerMap(index) = GetPlayerMap(MyIndex) Then
            rec.top = PIC_Y * 2
            rec.Bottom = rec.top + PIC_Y
            rec.Left = PIC_X * 4
            rec.Right = rec.Left + 96

            X = GetPlayerX(index) * PIC_X + Player(index).XOffset + sx
            Y = GetPlayerY(index) * PIC_Y + Player(index).YOffset + sx

            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - PIC_X - NewXOffset, Y - (NewPlayerY * PIC_Y) - PIC_Y - NewYOffset - 8, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

            If Player(index).LevelUp >= 3 Then
                Player(index).LevelUp = Player(index).LevelUp - 1
            ElseIf Player(index).LevelUp >= 1 Then
                Player(index).LevelUp = Player(index).LevelUp + 1
            End If
        Else
            Player(index).LevelUpT = 0
        End If
    End If
End Sub

Sub BltSpriteChange(ByVal X As Long, ByVal Y As Long)
    If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_SPRITE_CHANGE Then
        If SpriteSize = 0 Then
            rec.top = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Data1 * PIC_Y + 16
            rec.Bottom = rec.top + PIC_Y - 16
        Else
            rec.top = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Data1 * 64 + 16
            rec.Bottom = rec.top + 64 - 16
        End If
        
        rec.Left = 96
        rec.Right = rec.Left + PIC_X

        X = X * PIC_X + sx
        Y = Y * PIC_Y + (sx / 2) '- 16

        If Y < 0 Then
            Exit Sub
        End If

        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub
