Attribute VB_Name = "ModPets"
Option Explicit

'Written by: Morpheus

Public PetName As String
Public PetAlive As Byte
Public PetHP As Long
Public PetMaxHP As Long
Public PetSTR As Integer
Public PetDEF As Integer
Public PetSPEED As Integer
Public PetMAGI As Integer
Public PetLevel As Integer
Public PetPoints As Integer
Public PetSP As Long
Public PetMaxSP As Long
Public PetMP As Long
Public PetMaxMP As Long
Public PetFP As Long
Public PetMaxFP As Long
Public PetExp As Long
Public PetNextLevel As Long

Type PetRec
    Sprite As Long
    Alive As Byte
    HP As Long
    MaxHP As Long
    SP As Long
    MaxSP As Long
    MP As Long
    MaxMP As Long
    FP As Long
    MaxFP As Long
    Map As Long
    x As Long
    y As Long
    Dir As Byte
    Moving As Byte
    XOffset As Long
    YOffset As Long
    AttackTimer As Long
    Attacking As Byte
    LastAttack As Long
    Level As Long
    SpriteSet As Long
    STR As Long
    DEF As Long
    SPEED As Long
    MAGI As Long
    POINTS As Long
    Exp As Long
    Name As String
End Type


Sub BltPet(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim AttackSpeed As Long

    ' Only used if ever want to switch to blt rather then bltfast
    ' I suggest you don't use, because custom sizes won't work any longer
    With rec_pos
        .Top = Player(Index).Pet.y * PIC_Y + Player(Index).Pet.YOffset - (32 - PIC_Y)
        .Bottom = .Top + PIC_Y
        .Left = Player(Index).Pet.x * PIC_X + Player(Index).Pet.XOffset + ((32 - PIC_X) / 2)
        .Right = .Left + PIC_X + ((32 - PIC_X) / 2)
    End With
   
    ' Check for animation
    Anim = 0
    If Player(Index).Pet.Attacking = 0 Then
        Select Case Player(Index).Pet.Dir
            Case DIR_UP
                If (Player(Index).Pet.YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (Player(Index).Pet.YOffset > PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (Player(Index).Pet.XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (Player(Index).Pet.XOffset > PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If Player(Index).Pet.AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
   
    ' Check to see if we want to stop making him attack
    If Player(Index).Pet.AttackTimer + 1000 < GetTickCount Then
        Player(Index).Pet.Attacking = 0
        Player(Index).Pet.AttackTimer = 0
    End If
   
    rec.Top = Player(Index).Pet.Sprite * 32 + (32 - PIC_Y)
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Player(Index).Pet.Dir * (3 * (32 / PIC_X)) + (Anim * (32 / PIC_X))) * PIC_X
    rec.Right = rec.Left + 32

    x = Player(Index).Pet.x * PIC_X - (32 - PIC_X) / 2 + sx + Player(Index).Pet.XOffset
    y = Player(Index).Pet.y * PIC_Y - (32 - PIC_Y) + sx + Player(Index).Pet.YOffset + (32 - PIC_Y)
   
    If 32 > PIC_X Then
        If x < 0 Then
            x = Player(Index).Pet.XOffset + sx + ((32 - PIC_X) / 2)
            If Player(Index).Pet.Dir = DIR_RIGHT And Player(Index).Pet.Moving > 0 Then
                rec.Left = rec.Left - Player(Index).Pet.XOffset
            Else
                rec.Left = rec.Left - Player(Index).Pet.XOffset + ((32 - PIC_X) / 2)
            End If
        End If
       
        If x > MAX_MAPX * 32 Then
            x = MAX_MAPX * 32 + sx - ((32 - PIC_X) / 2) + Player(Index).Pet.XOffset
            If Player(Index).Pet.Dir = DIR_LEFT And Player(Index).Pet.Moving > 0 Then
                rec.Right = rec.Right + Player(Index).Pet.XOffset
            Else
                rec.Right = rec.Right + Player(Index).Pet.XOffset - ((32 - PIC_X) / 2)
            End If
        End If
    End If
   
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PetSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPetTop(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long
Dim AttackSpeed As Long

    ' Only used if ever want to switch to blt rather then bltfast
    ' I suggest you don't use, because custom sizes won't work any longer
    With rec_pos
        .Top = Player(Index).Pet.y * PIC_Y + Player(Index).Pet.YOffset - (32 - PIC_Y)
        .Bottom = .Top + PIC_Y
        .Left = Player(Index).Pet.x * PIC_X + Player(Index).Pet.XOffset + ((32 - PIC_X) / 2)
        .Right = .Left + PIC_X + ((32 - PIC_X) / 2)
    End With
   
    ' Check for animation
    Anim = 0
    If Player(Index).Pet.Attacking = 0 Then
        Select Case Player(Index).Pet.Dir
            Case DIR_UP
                If (Player(Index).Pet.YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (Player(Index).Pet.YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (Player(Index).Pet.XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (Player(Index).Pet.XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If Player(Index).Pet.AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
   
    ' Check to see if we want to stop making him attack
    If Player(Index).Pet.AttackTimer + 1000 < GetTickCount Then
        Player(Index).Pet.Attacking = 0
        Player(Index).Pet.AttackTimer = 0
    End If
   
    rec.Top = Player(Index).Pet.Sprite * 32
    rec.Bottom = rec.Top + (32 - PIC_Y)
    rec.Left = (Player(Index).Pet.Dir * (3 * (32 / PIC_X)) + (Anim * (32 / PIC_X))) * PIC_X
    rec.Right = rec.Left + 32

    x = Player(Index).Pet.x * PIC_X - (32 - PIC_X) / 2 + sx + Player(Index).Pet.XOffset
    y = Player(Index).Pet.y * PIC_Y - (32 - PIC_Y) + sx + Player(Index).Pet.YOffset
   
   
    If y < 0 Then
        y = 0
        If Player(Index).Pet.Dir = DIR_DOWN And Player(Index).Pet.Moving > 0 Then
            rec.Top = rec.Top - Player(Index).Pet.YOffset
        Else
            rec.Top = rec.Top - Player(Index).Pet.YOffset + (32 - PIC_Y)
        End If
    End If
   
    If 32 > PIC_X Then
        If x < 0 Then
            x = Player(Index).Pet.XOffset + sx + ((32 - PIC_X) / 2)
            If Player(Index).Pet.Dir = DIR_RIGHT And Player(Index).Pet.Moving > 0 Then
                rec.Left = rec.Left - Player(Index).Pet.XOffset
            Else
                rec.Left = rec.Left - Player(Index).Pet.XOffset + ((32 - PIC_X) / 2)
            End If
        End If
       
        If x > MAX_MAPX * 32 Then
            x = MAX_MAPX * 32 + sx - ((32 - PIC_X) / 2) + Player(Index).Pet.XOffset
            If Player(Index).Pet.Dir = DIR_LEFT And Player(Index).Pet.Moving > 0 Then
                rec.Right = rec.Right + Player(Index).Pet.XOffset
            Else
                rec.Right = rec.Right + Player(Index).Pet.XOffset - ((32 - PIC_X) / 2)
            End If
        End If
    End If
   
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_PetSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub


Sub BltPetName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long
    
    ' Check access level
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerAccess(Index)
            Case 0
                color = QBColor(YELLOW)
            Case 1
                color = QBColor(YELLOW)
            Case 2
                color = QBColor(YELLOW)
            Case 3
                color = QBColor(YELLOW)
            Case 4
                color = QBColor(YELLOW)
        End Select
    Else
        color = QBColor(BRIGHTRED)
    End If
'If Not ChangingMap Then
' Draw name
' If PetName = "" Then
' TextX = Player(Index).Pet.x * PIC_X + sx + Player(Index).Pet.XOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index) & "'s Pet") / 2) * 8)
' TextY = Player(Index).Pet.y * PIC_Y + sx + Player(Index).Pet.YOffset - Int(PIC_Y / 2) - (SIZE_Y - PIC_Y)
' Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index) & "'s Pet", Color)
' Else
TextX = Player(Index).Pet.x * PIC_X + sx + Player(Index).Pet.XOffset + Int(PIC_X / 2) - ((Len(Player(Index).Pet.Name) / 2) * 8)
TextY = Player(Index).Pet.y * PIC_Y + sx + Player(Index).Pet.YOffset - Int(PIC_Y / 6) - (SIZE_Y - PIC_Y)
Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Player(Index).Pet.Name, color)
' End If


'End If
End Sub

Sub ProcessPetMovement(ByVal PetNum As Long)
    ' Check if pet is walking, and if so process moving them over
    If Player(PetNum).Pet.Moving = MOVING_WALKING Then
        Select Case Player(PetNum).Pet.Dir
            Case DIR_UP
                Player(PetNum).Pet.YOffset = Player(PetNum).Pet.YOffset - WALK_SPEED
            Case DIR_DOWN
                Player(PetNum).Pet.YOffset = Player(PetNum).Pet.YOffset + WALK_SPEED
            Case DIR_LEFT
                Player(PetNum).Pet.XOffset = Player(PetNum).Pet.XOffset - WALK_SPEED
            Case DIR_RIGHT
                Player(PetNum).Pet.XOffset = Player(PetNum).Pet.XOffset + WALK_SPEED
        End Select
        
        ' Check if completed walking over to the next tile
        If (Player(PetNum).Pet.XOffset = 0) And (Player(PetNum).Pet.YOffset = 0) Then
            Player(PetNum).Pet.Moving = 0
        End If
    End If
End Sub

Public Sub PetMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim X1, Y1 As Long
    If Player(MyIndex).Pet.Alive = NO Then Exit Sub
    
    X1 = Int(x / PIC_X)
    Y1 = Int(y / PIC_Y)
    
    If (Button = 1) And (X1 >= 0) And (X1 <= MAX_MAPX) And (Y1 >= 0) And (Y1 <= MAX_MAPY) Then
        Call SendData("PETMOVESELECT" & SEP_CHAR & X1 & SEP_CHAR & Y1 & SEP_CHAR & END_CHAR)
    End If
End Sub

Sub BltPetBars(ByVal Index As Long)
Dim x As Long, y As Long

    x = (Player(Index).Pet.x * PIC_X + sx + Player(Index).Pet.XOffset) - (NewPlayerX * PIC_X) - NewXOffset
    y = (Player(Index).Pet.y * PIC_Y + sx + Player(Index).Pet.YOffset) - (NewPlayerY * PIC_Y) - NewYOffset
    
    If Player(Index).HP = 0 Then Exit Sub
    'draws the back bars
   ' Call DD_MiddleBuffer.SetFillColor(RGB(0, 0, 255))
   ' Call DD_MiddleBuffer.DrawBox(X, Y + 32, X + 32, Y + 36)
    
    'draws HP
    If Int((Player(Index).Pet.HP / Player(Index).Pet.MaxHP) * 100) > 50 Then
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
    End If
    If Int((Player(Index).Pet.HP / Player(Index).Pet.MaxHP) * 100) > 20 And Int((Player(Index).Pet.HP / Player(Index).Pet.MaxHP) * 100) <= 50 Then
        Call DD_BackBuffer.SetFillColor(RGB(255, 255, 0))
    End If
    If Int((Player(Index).Pet.HP / Player(Index).Pet.MaxHP) * 100) <= 20 Then
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
    End If
    Call DD_BackBuffer.DrawBox(x, y + PIC_Y, x + ((Player(Index).Pet.HP / 100) / (Player(Index).Pet.MaxHP / 100) * SIZE_X), y + 36)
End Sub
