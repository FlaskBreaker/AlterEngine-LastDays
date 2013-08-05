Attribute VB_Name = "modGameLogic"
Option Explicit

Function GetPlayerDamage(ByVal Index As Long) As Long
    Dim WeaponSlot As Long, RingSlot As Long, NecklaceSlot As Long
    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If

    ' GetPlayerDamage in script - TODO LATER - Can't get it to work. :(
    ' If Scripting = 1 Then
    ' GetPlayerDamage = MyScript.RunCodeReturn("Scripts\Main.txt", "GetPlayerDamage ", index)
    ' Else
    GetPlayerDamage = Int(GetPlayerSTR(Index) / 2)
' End If

    If GetPlayerDamage <= 0 Then
        GetPlayerDamage = 1
    End If

    If GetPlayerWeaponSlot(Index) > 0 Then
        WeaponSlot = GetPlayerWeaponSlot(Index)

        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data2

        If GetPlayerInvItemDur(Index, WeaponSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, WeaponSlot, GetPlayerInvItemDur(Index, WeaponSlot) - 1)

            If GetPlayerInvItemDur(Index, WeaponSlot) = 0 Then
                Call BattleMsg(Index, "Tu " & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " se ha roto.", YELLOW, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, WeaponSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, WeaponSlot) <= 10 Then
                    Call BattleMsg(Index, "Tu " & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Name) & " esta apunto de romperse. Duración: " & GetPlayerInvItemDur(Index, WeaponSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, WeaponSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If

    If GetPlayerRingSlot(Index) > 0 Then
        RingSlot = GetPlayerRingSlot(Index)

        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(Index, RingSlot)).Data2

        If GetPlayerInvItemDur(Index, RingSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, RingSlot, GetPlayerInvItemDur(Index, RingSlot) - 1)

            If GetPlayerInvItemDur(Index, RingSlot) = 0 Then
                Call BattleMsg(Index, "Tu " & Trim$(Item(GetPlayerInvItemNum(Index, RingSlot)).Name) & " se ha roto.", YELLOW, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, RingSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, RingSlot) <= 10 Then
                    Call BattleMsg(Index, "Tu " & Trim$(Item(GetPlayerInvItemNum(Index, RingSlot)).Name) & " esta apunto de romperse. Duración: " & GetPlayerInvItemDur(Index, RingSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, RingSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If

    If GetPlayerNecklaceSlot(Index) > 0 Then
        NecklaceSlot = GetPlayerNecklaceSlot(Index)

        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(Index, NecklaceSlot)).Data2

        If GetPlayerInvItemDur(Index, NecklaceSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, NecklaceSlot, GetPlayerInvItemDur(Index, NecklaceSlot) - 1)

            If GetPlayerInvItemDur(Index, NecklaceSlot) = 0 Then
                Call BattleMsg(Index, "Tu " & Trim$(Item(GetPlayerInvItemNum(Index, NecklaceSlot)).Name) & " se ha roto.", YELLOW, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, NecklaceSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, NecklaceSlot) <= 10 Then
                    Call BattleMsg(Index, "Tu " & Trim$(Item(GetPlayerInvItemNum(Index, NecklaceSlot)).Name) & " esta apunto de romperse. Duración: " & GetPlayerInvItemDur(Index, NecklaceSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, NecklaceSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If



    If GetPlayerDamage < 0 Then
        GetPlayerDamage = 0
    End If
End Function

Function GetPlayerProtection(ByVal Index As Long) As Long
    Dim ArmorSlot As Long, HelmSlot As Long, ShieldSlot As Long, LegsSlot As Long

    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > MAX_PLAYERS Then
        Exit Function
    End If

    ArmorSlot = GetPlayerArmorSlot(Index)
    HelmSlot = GetPlayerHelmetSlot(Index)
    ShieldSlot = GetPlayerShieldSlot(Index)
    LegsSlot = GetPlayerLegsSlot(Index)
    GetPlayerProtection = Int(GetPlayerDEF(Index) / 5)

    If ArmorSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, ArmorSlot)).Data2
        If GetPlayerInvItemDur(Index, ArmorSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, ArmorSlot, GetPlayerInvItemDur(Index, ArmorSlot) - 1)

            If GetPlayerInvItemDur(Index, ArmorSlot) = 0 Then
                Call BattleMsg(Index, "Tu " & Trim$(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " se ha roto.", YELLOW, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, ArmorSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, ArmorSlot) <= 10 Then
                    Call BattleMsg(Index, "Tu " & Trim$(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " esta apunto de romperse. Duración: " & GetPlayerInvItemDur(Index, ArmorSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If

    If HelmSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, HelmSlot)).Data2
        If GetPlayerInvItemDur(Index, HelmSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, HelmSlot, GetPlayerInvItemDur(Index, HelmSlot) - 1)

            If GetPlayerInvItemDur(Index, HelmSlot) <= 0 Then
                Call BattleMsg(Index, "Tu " & Trim$(Item(GetPlayerInvItemNum(Index, HelmSlot)).Name) & " se ha roto.", YELLOW, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, HelmSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, HelmSlot) <= 10 Then
                    Call BattleMsg(Index, "Tu " & Trim$(Item(GetPlayerInvItemNum(Index, HelmSlot)).Name) & " " & Trim$(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " esta apunto de romperse. Duración: " & GetPlayerInvItemDur(Index, HelmSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, HelmSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If

    If ShieldSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, ShieldSlot)).Data2
        If GetPlayerInvItemDur(Index, ShieldSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, ShieldSlot, GetPlayerInvItemDur(Index, ShieldSlot) - 1)

            If GetPlayerInvItemDur(Index, ShieldSlot) <= 0 Then
                Call BattleMsg(Index, "Tu " & Trim$(Item(GetPlayerInvItemNum(Index, ShieldSlot)).Name) & " se ha roto.", YELLOW, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, ShieldSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, ShieldSlot) <= 10 Then
                    Call BattleMsg(Index, "Tu " & Trim$(Item(GetPlayerInvItemNum(Index, ShieldSlot)).Name) & " esta apunto de romperse. Duración: " & GetPlayerInvItemDur(Index, ShieldSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, ShieldSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If

    If LegsSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(Index, LegsSlot)).Data2
        If GetPlayerInvItemDur(Index, LegsSlot) > -1 Then
            Call SetPlayerInvItemDur(Index, LegsSlot, GetPlayerInvItemDur(Index, LegsSlot) - 1)

            If GetPlayerInvItemDur(Index, LegsSlot) <= 0 Then
                Call BattleMsg(Index, "Tu " & Trim$(Item(GetPlayerInvItemNum(Index, LegsSlot)).Name) & " se ha roto.", YELLOW, 0)
                Call TakeItem(Index, GetPlayerInvItemNum(Index, LegsSlot), 0)
            Else
                If GetPlayerInvItemDur(Index, LegsSlot) <= 10 Then
                    Call BattleMsg(Index, "Tu " & Trim$(Item(GetPlayerInvItemNum(Index, LegsSlot)).Name) & " " & Trim$(Item(GetPlayerInvItemNum(Index, ArmorSlot)).Name) & " esta apunto de romperse. Duración: " & GetPlayerInvItemDur(Index, LegsSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(Index, LegsSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If
End Function

Function FindOpenPlayerSlot() As Long
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If Not IsConnected(I) Then
            FindOpenPlayerSlot = I
            Exit Function
        End If
    Next I
End Function

Public Function FindOpenInvSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim I As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Check to see if they already have the item.
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
        For I = 1 To MAX_INV
            If GetPlayerInvItemNum(Index, I) = ItemNum Then
                FindOpenInvSlot = I
                Exit Function
            End If
        Next I
    End If

    ' Try to find an open inventory slot.
    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(Index, I) = 0 Then
            FindOpenInvSlot = I
            Exit Function
        End If
    Next I
End Function

Function FindOpenBankSlot(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim I As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Check to see if they already have the item.
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
        For I = 1 To MAX_BANK
            If GetPlayerBankItemNum(Index, I) = ItemNum Then
                FindOpenBankSlot = I
                Exit Function
            End If
        Next I
    End If

    ' Try to find an open bank slot.
    For I = 1 To MAX_BANK
        If GetPlayerBankItemNum(Index, I) = 0 Then
            FindOpenBankSlot = I
            Exit Function
        End If
    Next I
End Function

Function FindOpenMapItemSlot(ByVal mapnum As Long) As Long
    Dim I As Long

    ' Check for subscript out of range.
    If mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Function
    End If

    For I = 1 To MAX_MAP_ITEMS
        If MapItem(mapnum, I).num = 0 Then
            FindOpenMapItemSlot = I
            Exit Function
        End If
    Next I
End Function

Function FindOpenSpellSlot(ByVal Index As Long) As Long
    Dim I As Long

    For I = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, I) = 0 Then
            FindOpenSpellSlot = I
            Exit Function
        End If
    Next I
End Function

Function HasSpell(ByVal Index As Long, ByVal spellnum As Long) As Boolean
    Dim I As Long

    For I = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(Index, I) = spellnum Then
            HasSpell = True
            Exit Function
        End If
    Next I
End Function

Function TotalOnlinePlayers() As Long
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If
    Next I
End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim I As Long

    Name = LCase$(Name)

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If Len(GetPlayerName(I)) >= Len(Name) Then
                If LCase$(GetPlayerName(I)) = Name Then
                    FindPlayer = I
                    Exit Function
                End If
            End If
        End If
    Next I
End Function

Function HasItem(ByVal Index As Long, ByVal ItemNum As Long) As Long
    Dim I As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Check to see if the player has the item.
    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(Index, I) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                HasItem = GetPlayerInvItemValue(Index, I)
            Else
                HasItem = 1
            End If

            Exit Function
        End If
    Next I
End Function

Sub TakeItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
    Dim I As Long, N As Long
    Dim TakeItem As Boolean

    TakeItem = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    For I = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(Index, I) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(Index, I) Then
                    TakeItem = True
                Else
                    Call SetPlayerInvItemValue(Index, I, GetPlayerInvItemValue(Index, I) - ItemVal)
                    Call SendInventoryUpdate(Index, I)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerInvItemNum(Index, I)).Type
                    Case ITEM_TYPE_WEAPON
                        If GetPlayerWeaponSlot(Index) > 0 Then
                            If I = GetPlayerWeaponSlot(Index) Then
                                Call SetPlayerWeaponSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If

                    Case ITEM_TYPE_ARMOR
                        If GetPlayerArmorSlot(Index) > 0 Then
                            If I = GetPlayerArmorSlot(Index) Then
                                Call SetPlayerArmorSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If

                    Case ITEM_TYPE_HELMET
                        If GetPlayerHelmetSlot(Index) > 0 Then
                            If I = GetPlayerHelmetSlot(Index) Then
                                Call SetPlayerHelmetSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If

                    Case ITEM_TYPE_SHIELD
                        If GetPlayerShieldSlot(Index) > 0 Then
                            If I = GetPlayerShieldSlot(Index) Then
                                Call SetPlayerShieldSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If

                    Case ITEM_TYPE_LEGS
                        If GetPlayerLegsSlot(Index) > 0 Then
                            If I = GetPlayerLegsSlot(Index) Then
                                Call SetPlayerLegsSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If

                    Case ITEM_TYPE_RING
                        If GetPlayerRingSlot(Index) > 0 Then
                            If I = GetPlayerRingSlot(Index) Then
                                Call SetPlayerRingSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If

                    Case ITEM_TYPE_NECKLACE
                        If GetPlayerNecklaceSlot(Index) > 0 Then
                            If I = GetPlayerNecklaceSlot(Index) Then
                                Call SetPlayerNecklaceSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                End Select


                N = Item(GetPlayerInvItemNum(Index, I)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (N <> ITEM_TYPE_WEAPON) And (N <> ITEM_TYPE_ARMOR) And (N <> ITEM_TYPE_HELMET) And (N <> ITEM_TYPE_SHIELD) And (N <> ITEM_TYPE_LEGS) And (N <> ITEM_TYPE_RING) And (N <> ITEM_TYPE_NECKLACE) Then
                    TakeItem = True
                End If
            End If

            If TakeItem = True Then
                Call SetPlayerInvItemNum(Index, I, 0)
                Call SetPlayerInvItemValue(Index, I, 0)
                Call SetPlayerInvItemDur(Index, I, 0)

                ' Send the inventory update
                Call SendInventoryUpdate(Index, I)
                Exit Sub
            End If
        End If
    Next I
End Sub

Sub GiveItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
    Dim I As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Sub
    End If

    I = FindOpenInvSlot(Index, ItemNum)

    ' Check to see if inventory is full
    If I > 0 Then
        Call SetPlayerInvItemNum(Index, I, ItemNum)
        Call SetPlayerInvItemValue(Index, I, GetPlayerInvItemValue(Index, I) + ItemVal)

        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Or (Item(ItemNum).Type = ITEM_TYPE_LEGS) Or (Item(ItemNum).Type = ITEM_TYPE_RING) Or (Item(ItemNum).Type = ITEM_TYPE_NECKLACE) Then
            Call SetPlayerInvItemDur(Index, I, Item(ItemNum).Data1)
        End If

        Call SendInventoryUpdate(Index, I)
    Else
        Call PlayerMsg(Index, "Tu inventario esta lleno!", BRIGHTRED)
    End If
End Sub

Sub TakeBankItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
    Dim I As Long, N As Long
    Dim TakeBankItem As Boolean

    TakeBankItem = False

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    For I = 1 To MAX_BANK
        ' Check to see if the player has the item
        If GetPlayerBankItemNum(Index, I) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                ' Is what we are trying to take away more then what they have? If so just set it to zero
                If ItemVal >= GetPlayerBankItemValue(Index, I) Then
                    TakeBankItem = True
                Else
                    Call SetPlayerBankItemValue(Index, I, GetPlayerBankItemValue(Index, I) - ItemVal)
                    Call SendBankUpdate(Index, I)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerBankItemNum(Index, I)).Type
                    Case ITEM_TYPE_WEAPON
                        If GetPlayerWeaponSlot(Index) > 0 Then
                            If I = GetPlayerWeaponSlot(Index) Then
                                Call SetPlayerWeaponSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerWeaponSlot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If

                    Case ITEM_TYPE_ARMOR
                        If GetPlayerArmorSlot(Index) > 0 Then
                            If I = GetPlayerArmorSlot(Index) Then
                                Call SetPlayerArmorSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerArmorSlot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If

                    Case ITEM_TYPE_HELMET
                        If GetPlayerHelmetSlot(Index) > 0 Then
                            If I = GetPlayerHelmetSlot(Index) Then
                                Call SetPlayerHelmetSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerHelmetSlot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If

                    Case ITEM_TYPE_SHIELD
                        If GetPlayerShieldSlot(Index) > 0 Then
                            If I = GetPlayerShieldSlot(Index) Then
                                Call SetPlayerShieldSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerShieldSlot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If

                    Case ITEM_TYPE_LEGS
                        If GetPlayerLegsSlot(Index) > 0 Then
                            If I = GetPlayerLegsSlot(Index) Then
                                Call SetPlayerLegsSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerLegsSlot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If

                    Case ITEM_TYPE_RING
                        If GetPlayerRingSlot(Index) > 0 Then
                            If I = GetPlayerRingSlot(Index) Then
                                Call SetPlayerRingSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerRingSlot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If

                    Case ITEM_TYPE_NECKLACE
                        If GetPlayerNecklaceSlot(Index) > 0 Then
                            If I = GetPlayerNecklaceSlot(Index) Then
                                Call SetPlayerNecklaceSlot(Index, 0)
                                Call SendWornEquipment(Index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(Index, GetPlayerNecklaceSlot(Index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If
                End Select


                N = Item(GetPlayerBankItemNum(Index, I)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (N <> ITEM_TYPE_WEAPON) And (N <> ITEM_TYPE_ARMOR) And (N <> ITEM_TYPE_HELMET) And (N <> ITEM_TYPE_SHIELD) And (N <> ITEM_TYPE_LEGS) And (N <> ITEM_TYPE_RING) And (N <> ITEM_TYPE_NECKLACE) Then
                    TakeBankItem = True
                End If
            End If

            If TakeBankItem = True Then
                Call SetPlayerBankItemNum(Index, I, 0)
                Call SetPlayerBankItemValue(Index, I, 0)
                Call SetPlayerBankItemDur(Index, I, 0)

                ' Send the Bank update
                Call SendBankUpdate(Index, I)
                Exit Sub
            End If
        End If
    Next I
End Sub

Sub GiveBankItem(ByVal Index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal BankSlot As Long)
    Dim I As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Sub
    End If

    I = BankSlot

    ' Check to see if Bankentory is full
    If I > 0 Then
        Call SetPlayerBankItemNum(Index, I, ItemNum)
        Call SetPlayerBankItemValue(Index, I, GetPlayerBankItemValue(Index, I) + ItemVal)

        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Or (Item(ItemNum).Type = ITEM_TYPE_LEGS) Or (Item(ItemNum).Type = ITEM_TYPE_RING) Or (Item(ItemNum).Type = ITEM_TYPE_NECKLACE) Then
            Call SetPlayerBankItemDur(Index, I, Item(ItemNum).Data1)
        End If
    Else
        Call SendDataTo(Index, "bankmsg" & SEP_CHAR & "Banco lleno!" & END_CHAR)
    End If
End Sub

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long)
    Dim I As Long

    ' Check for subscript out of range.
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot.
    I = FindOpenMapItemSlot(mapnum)

    Call SpawnItemSlot(I, ItemNum, ItemVal, Item(ItemNum).Data1, mapnum, X, Y)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long)
    Dim I As Long

    ' Check for subscript out of range.
    If mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If MapItemSlot < 1 Or MapItemSlot > MAX_MAP_ITEMS Then
        Exit Sub
    End If

    I = MapItemSlot

    If I > 0 Then
        MapItem(mapnum, I).num = ItemNum
        MapItem(mapnum, I).Value = ItemVal

        If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_NECKLACE) Then
            MapItem(mapnum, I).Dur = ItemDur
        Else
            MapItem(mapnum, I).Dur = 0
        End If

        MapItem(mapnum, I).X = X
        MapItem(mapnum, I).Y = Y

        Call SendDataToMap(mapnum, "SPAWNITEM" & SEP_CHAR & I & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(mapnum, I).Dur & SEP_CHAR & X & SEP_CHAR & Y & END_CHAR)
    End If
End Sub

Sub SpawnAllMapsItems()
    Dim I As Long

    For I = 1 To MAX_MAPS
        Call SpawnMapItems(I)
    Next I
End Sub

Sub SpawnMapItems(ByVal mapnum As Long)
    Dim X As Integer
    Dim Y As Integer

    ' Check for subscript out of range.
    If mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn all the mapped items on their specified tile.
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            If Map(mapnum).Tile(X, Y).Type = TILE_TYPE_ITEM Then
                If (Item(Map(mapnum).Tile(X, Y).Data1).Type = ITEM_TYPE_CURRENCY Or Item(Map(mapnum).Tile(X, Y).Data1).Stackable = 1) And Map(mapnum).Tile(X, Y).Data2 <= 0 Then
                    Call SpawnItem(Map(mapnum).Tile(X, Y).Data1, 1, mapnum, X, Y)
                Else
                    Call SpawnItem(Map(mapnum).Tile(X, Y).Data1, Map(mapnum).Tile(X, Y).Data2, mapnum, X, Y)
                End If
            End If
        Next X
    Next Y
End Sub

Sub PlayerMapGetItem(ByVal Index As Long)
    Dim I As Long
    Dim N As Long
    Dim mapnum As Long
    Dim Msg As String

    If IsPlaying(Index) = False Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(Index)

    For I = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(mapnum, I).num > 0) Then
            If (MapItem(mapnum, I).num <= MAX_ITEMS) Then
        
                ' Check if item is at the same location as the player
                If (MapItem(mapnum, I).X = GetPlayerX(Index)) Then
                    If (MapItem(mapnum, I).Y = GetPlayerY(Index)) Then
                    
                        ' Find open slot
                        N = FindOpenInvSlot(Index, MapItem(mapnum, I).num)
        
                        ' Open slot available?
                        If N <> 0 Then
                            ' Set item in players inventory
                            Call SetPlayerInvItemNum(Index, N, MapItem(mapnum, I).num)
                            If Item(GetPlayerInvItemNum(Index, N)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, N)).Stackable = 1 Then
                                Call SetPlayerInvItemValue(Index, N, GetPlayerInvItemValue(Index, N) + MapItem(mapnum, I).Value)
                                Msg = "Tu obtienes " & MapItem(mapnum, I).Value & " " & Trim$(Item(GetPlayerInvItemNum(Index, N)).Name) & "."
                            Else
                                Call SetPlayerInvItemValue(Index, N, 0)
                                Msg = "Tu obtienes " & Trim$(Item(GetPlayerInvItemNum(Index, N)).Name) & "."
                            End If
                            Call SetPlayerInvItemDur(Index, N, MapItem(mapnum, I).Dur)
        
                            ' Borra todos los objetos del mapa.
                            Call ClearMapItem(I, mapnum)
        
                            Call SendInventoryUpdate(Index, N)
                            Call SpawnItemSlot(I, 0, 0, 0, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
                            Call PlayerMsg(Index, Msg, YELLOW)
                            Exit Sub
                        Else
                            Call PlayerMsg(Index, "Tu inventario está lleno!", BRIGHTRED)
                            Exit Sub
                        End If
                    End If
                End If
                
            End If
        End If
    Next I
End Sub

Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    Dim I As Long
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If

    If (GetPlayerInvItemNum(Index, InvNum) > 0) Then
        If (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
            I = FindOpenMapItemSlot(GetPlayerMap(Index))
    
            If I <> 0 Then
                MapItem(GetPlayerMap(Index), I).Dur = 0
    
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type
                    Case ITEM_TYPE_ARMOR
                        If InvNum = GetPlayerArmorSlot(Index) Then
                            Call SetPlayerArmorSlot(Index, 0)
                            Call SendWornEquipment(Index)
                            Call SendIndexWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
    
                    Case ITEM_TYPE_WEAPON
                        If InvNum = GetPlayerWeaponSlot(Index) Then
                            Call SetPlayerWeaponSlot(Index, 0)
                            Call SendWornEquipment(Index)
                            Call SendIndexWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
    
                    Case ITEM_TYPE_HELMET
                        If InvNum = GetPlayerHelmetSlot(Index) Then
                            Call SetPlayerHelmetSlot(Index, 0)
                            Call SendWornEquipment(Index)
                            Call SendIndexWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
    
                    Case ITEM_TYPE_SHIELD
                        If InvNum = GetPlayerShieldSlot(Index) Then
                            Call SetPlayerShieldSlot(Index, 0)
                            Call SendWornEquipment(Index)
                            Call SendIndexWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
                    Case ITEM_TYPE_LEGS
                        If InvNum = GetPlayerLegsSlot(Index) Then
                            Call SetPlayerLegsSlot(Index, 0)
                            Call SendWornEquipment(Index)
                            Call SendIndexWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
                    Case ITEM_TYPE_RING
                        If InvNum = GetPlayerRingSlot(Index) Then
                            Call SetPlayerRingSlot(Index, 0)
                            Call SendWornEquipment(Index)
                            Call SendIndexWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
                    Case ITEM_TYPE_NECKLACE
                        If InvNum = GetPlayerNecklaceSlot(Index) Then
                            Call SetPlayerNecklaceSlot(Index, 0)
                            Call SendWornEquipment(Index)
                            Call SendIndexWornEquipment(Index)
                        End If
                        MapItem(GetPlayerMap(Index), I).Dur = GetPlayerInvItemDur(Index, InvNum)
                End Select
    
                MapItem(GetPlayerMap(Index), I).num = GetPlayerInvItemNum(Index, InvNum)
                MapItem(GetPlayerMap(Index), I).X = GetPlayerX(Index)
                MapItem(GetPlayerMap(Index), I).Y = GetPlayerY(Index)
    
                If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(Index, InvNum)).Stackable = 1 Then
                    ' Check if its more then they have and if so drop it all
                    If Amount >= GetPlayerInvItemValue(Index, InvNum) Then
                        MapItem(GetPlayerMap(Index), I).Value = GetPlayerInvItemValue(Index, InvNum)
                        ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, InvNum) & " " & Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemNum(Index, InvNum, 0)
                        Call SetPlayerInvItemValue(Index, InvNum, 0)
                        Call SetPlayerInvItemDur(Index, InvNum, 0)
                    Else
                        MapItem(GetPlayerMap(Index), I).Value = Amount
                        ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & Amount & " " & Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemValue(Index, InvNum, GetPlayerInvItemValue(Index, InvNum) - Amount)
                    End If
                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(Index), I).Value = 0
    
                    ' Normally messages for item drops would go here but it's scripted now
    
                    Call SetPlayerInvItemNum(Index, InvNum, 0)
                    Call SetPlayerInvItemValue(Index, InvNum, 0)
                    Call SetPlayerInvItemDur(Index, InvNum, 0)
                End If
    
                ' Send inventory update
                Call SendInventoryUpdate(Index, InvNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(I, MapItem(GetPlayerMap(Index), I).num, Amount, MapItem(GetPlayerMap(Index), I).Dur, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))
    
                If scripting = 1 Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "onitemdrop " & Index & "," & GetPlayerMap(Index) & "," & MapItem(GetPlayerMap(Index), I).num & "," & Amount & "," & MapItem(GetPlayerMap(Index), I).Dur & "," & I & "," & InvNum
                End If
    
            Else
                Call PlayerMsg(Index, "Hay demasiado objetos en el suelo.", BRIGHTRED)
            End If
        End If
        
    End If
End Sub

Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal mapnum As Long)
    Dim Packet As String
    Dim npcnum As Long
    Dim I As Long
    Dim X As Long
    Dim Y As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Or mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    npcnum = Map(mapnum).NPC(MapNpcNum)

    If npcnum > 0 Then
        If GameTime = TIME_NIGHT Then
            If NPC(npcnum).SpawnTime = 1 Then
                MapNPC(mapnum, MapNpcNum).num = 0
                MapNPC(mapnum, MapNpcNum).SpawnWait = GetTickCount
                MapNPC(mapnum, MapNpcNum).HP = 0
                Call SendDataToMap(mapnum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)
                Exit Sub
            End If
        Else
            If NPC(npcnum).SpawnTime = 2 Then
                MapNPC(mapnum, MapNpcNum).num = 0
                MapNPC(mapnum, MapNpcNum).SpawnWait = GetTickCount
                MapNPC(mapnum, MapNpcNum).HP = 0
                Call SendDataToMap(mapnum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)
                Exit Sub
            End If
        End If

        MapNPC(mapnum, MapNpcNum).num = npcnum
        MapNPC(mapnum, MapNpcNum).Target = 0

        MapNPC(mapnum, MapNpcNum).HP = GetNpcMaxHP(npcnum)
        MapNPC(mapnum, MapNpcNum).MP = GetNpcMaxMP(npcnum)
        MapNPC(mapnum, MapNpcNum).SP = GetNpcMaxSP(npcnum)

        MapNPC(mapnum, MapNpcNum).Dir = Int(Rnd * 4)

        ' This means the admin wants to do a random spawn. [Mellowz]
        If Map(mapnum).SpawnX(MapNpcNum) = 0 Or Map(mapnum).SpawnY(MapNpcNum) = 0 Then
            For I = 1 To 100
                X = Int(Rnd * MAX_MAPX)
                Y = Int(Rnd * MAX_MAPY)
    
                If Map(mapnum).Tile(X, Y).Type = TILE_TYPE_WALKABLE Then
                    MapNPC(mapnum, MapNpcNum).X = X
                    MapNPC(mapnum, MapNpcNum).Y = Y
                    Spawned = True
                    Exit For
                End If
            Next I

            ' Didn't spawn, so now we'll just try to find a free tile
            If Not Spawned Then
                For Y = 0 To MAX_MAPY
                    For X = 0 To MAX_MAPX
                        If Map(mapnum).Tile(X, Y).Type = TILE_TYPE_WALKABLE Then
                            MapNPC(mapnum, MapNpcNum).X = X
                            MapNPC(mapnum, MapNpcNum).Y = Y
                            Spawned = True
                        End If
                    Next X
                Next Y
            End If
        Else
            ' We subtract one because Rand is ListIndex 0. [Mellowz]
            MapNPC(mapnum, MapNpcNum).X = Map(mapnum).SpawnX(MapNpcNum) - 1
            MapNPC(mapnum, MapNpcNum).Y = Map(mapnum).SpawnY(MapNpcNum) - 1
            Spawned = True
        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(mapnum, MapNpcNum).num & SEP_CHAR & MapNPC(mapnum, MapNpcNum).X & SEP_CHAR & MapNPC(mapnum, MapNpcNum).Y & SEP_CHAR & MapNPC(mapnum, MapNpcNum).Dir & SEP_CHAR & NPC(MapNPC(mapnum, MapNpcNum).num).Big & END_CHAR
            Call SendDataToMap(mapnum, Packet)
        End If
    End If

    ' Enable this to display HP when monsters spawn.
    ' Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNPC(MapNum, MapNpcNum).num) & END_CHAR)
End Sub

Sub SpawnMapNpcs(ByVal mapnum As Long)
    Dim I As Long

    For I = 1 To MAX_MAP_NPCS
        If Map(mapnum).NPC(I) > 0 Then
            Call SpawnNpc(I, mapnum)
        End If
    Next I
End Sub

Sub SpawnAllMapNpcs()
    Dim I As Long

    For I = 1 To MAX_MAPS
        If PlayersOnMap(I) = YES Then
            Call SpawnMapNpcs(I)
        End If
    Next I
End Sub

Function CanAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
    Dim AttackSpeed As Long

    If GetPlayerWeaponSlot(Attacker) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
    Else
        AttackSpeed = 1000
    End If
    CanAttackPlayer = False

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Then
        Exit Function
    End If

    ' Make sure they have more then 0 hp
    If GetPlayerHP(Victim) <= 0 Then
        Exit Function
    End If

    ' Make sure we dont attack the player if they are switching maps
    If Player(Victim).GettingMap = YES Then
        Exit Function
    End If

    ' Make sure they are on the same map
    If (GetPlayerMap(Attacker) = GetPlayerMap(Victim)) And (GetTickCount > Player(Attacker).AttackTimer + AttackSpeed) Then

        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
                If (GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
                        ' Check to make sure the victim isn't an admin
                        If GetPlayerAccess(Victim) > ADMIN_MONITER Then
                            Call PlayerMsg(Attacker, "No puedes atacar a " & GetPlayerName(Victim) & "!", BRIGHTRED)
                        Else
                            ' Check if map is attackable
                            If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                                ' Make sure they are high enough level
                                If GetPlayerLevel(Attacker) < PKMINLVL Then
                                    Call PlayerMsg(Attacker, "Tu nivel es menor de " & PKMINLVL & ", no puedes atacar a otro jugador aun!", BRIGHTRED)
                                Else
                                    If GetPlayerLevel(Victim) < PKMINLVL Then
                                        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " su nivel es menor de " & PKMINLVL & ", no puedes atacarlo aun!", BRIGHTRED)
                                    Else
                                        If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                                            If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                                                CanAttackPlayer = True
                                            Else
                                                Call PlayerMsg(Attacker, "No puedes atacar miembros de tu propio clan!", BRIGHTRED)
                                            End If
                                        Else
                                            CanAttackPlayer = True
                                        End If
                                    End If
                                End If
                            Else
                                Call PlayerMsg(Attacker, "Esto es una zona segura!", BRIGHTRED)
                            End If
                        End If
                    ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If

            Case DIR_DOWN
                If (GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
                        ' Check to make sure that they dont have access
                        ' Check if map is attackable
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                            ' Make sure they are high enough level
                            If GetPlayerLevel(Attacker) < PKMINLVL Then
                                Call PlayerMsg(Attacker, "Tu nivel es menor de " & PKMINLVL & ", no puedes atacar a otro jugador aun.", BRIGHTRED)
                            Else
                                If GetPlayerLevel(Victim) < PKMINLVL Then
                                    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " su nivel es menor de " & PKMINLVL & ", no puedes atacarlo aun!", BRIGHTRED)
                                Else
                                    If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                                        If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                                            CanAttackPlayer = True
                                        Else
                                            Call PlayerMsg(Attacker, "No puedes atacar a un miembro de tu propio clan!", BRIGHTRED)
                                        End If
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            End If
                        Else
                            Call PlayerMsg(Attacker, "Esto es una zona segura!", BRIGHTRED)
                        End If
                    End If
                End If
                If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                    CanAttackPlayer = True
                End If

            Case DIR_LEFT
                If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
                        ' Check if map is attackable
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                            ' Make sure they are high enough level
                            If GetPlayerLevel(Attacker) < PKMINLVL Then
                                Call PlayerMsg(Attacker, "Tu nivel es menor de " & PKMINLVL & ", no puedes atacar a otro jugador aun.", BRIGHTRED)
                            Else
                                If GetPlayerLevel(Victim) < PKMINLVL Then
                                    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " su nivel es menor de " & PKMINLVL & ", no puedes atacarlo aun.", BRIGHTRED)
                                Else
                                    If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                                        If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                                            CanAttackPlayer = True
                                        Else
                                                Call PlayerMsg(Attacker, "No puedes atacar a un miembro de tu propio clan!", BRIGHTRED)
                                        End If
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            End If
                        Else
                            Call PlayerMsg(Attacker, "Esto es una zona segura!", BRIGHTRED)
                        End If
                    ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If

            Case DIR_RIGHT
                If (GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker)) Then
                    If Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA Then
                        ' Check if map is attackable
                        If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
                            ' Make sure they are high enough level
                            If GetPlayerLevel(Attacker) < PKMINLVL Then
                                Call PlayerMsg(Attacker, "Tu nivel es menor de " & PKMINLVL & ", no puedes atacar a otro jugador aun.", BRIGHTRED)
                            Else
                                If GetPlayerLevel(Victim) < PKMINLVL Then
                                    Call PlayerMsg(Attacker, GetPlayerName(Victim) & " su nivel es menor de " & PKMINLVL & ", no puedes atacarlo aun.", BRIGHTRED)
                                Else
                                    If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                                        If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                                            CanAttackPlayer = True
                                        Else
                                            Call PlayerMsg(Attacker, "No puedes atacar a un miembro de tu clan!", BRIGHTRED)
                                        End If
                                    Else
                                        CanAttackPlayer = True
                                    End If
                                End If
                            End If
                        Else
                            Call PlayerMsg(Attacker, "Esto es una zona segura!", BRIGHTRED)
                        End If
                    ElseIf Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA Then
                        CanAttackPlayer = True
                    End If
                End If
        End Select
    End If
End Function

Function CanAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
    Dim mapnum As Long
    Dim npcnum As Long
    Dim AttackSpeed As Long
'MsgBox "CanAttackNpc launched"
    ' Check for sub-script out of range.
    If Not IsPlaying(Attacker) Then
        Exit Function
    End If

    ' Check for sub-script out of range.
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for sub-script out of range.
    If MapNPC(GetPlayerMap(Attacker), MapNpcNum).num = 0 Then
        Exit Function
    End If

    ' Get the players weapon attack speed.
    If GetPlayerWeaponSlot(Attacker) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
    Else
        AttackSpeed = 1000
    End If

    ' Get the players map number.
    mapnum = GetPlayerMap(Attacker)

    ' Get the NPCs map index.
    npcnum = MapNPC(mapnum, MapNpcNum).num

    ' Make sure the npc isn't already dead
    If MapNPC(mapnum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Checks to see if the player can attack.
    If GetTickCount > Player(Attacker).AttackTimer + AttackSpeed Then
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
                If (MapNPC(mapnum, MapNpcNum).Y + 1 = GetPlayerY(Attacker)) And (MapNPC(mapnum, MapNpcNum).X = GetPlayerX(Attacker)) Then
                    If NPC(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(npcnum).Behavior <> NPC_BEHAVIOR_SCRIPTED And NPC(npcnum).Behavior <> NPC_BEHAVIOR_QUEST And NPC(npcnum).Behavior <> NPC_BEHAVIOR_CHAOSKNIGHT Then
                        CanAttackNpc = True
                    Else
                        If NPC(npcnum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & NPC(npcnum).SpawnSecs
                        ElseIf NPC(npcnum).Behavior = NPC_BEHAVIOR_QUEST Then
                            Call DoQuest(NPC(npcnum).Quest, Attacker, npcnum)
                        ElseIf NPC(npcnum).Behavior = NPC_BEHAVIOR_CHAOSKNIGHT Then
                        Call DoQuestNpcKillsQuest(Attacker, npcnum)
                        Else
                            Call PlayerMsg(Attacker, Trim$(NPC(npcnum).Name) & " : " & Trim$(NPC(npcnum).AttackSay), Green)
                        End If
                    End If
                End If

            Case DIR_DOWN
                If (MapNPC(mapnum, MapNpcNum).Y - 1 = GetPlayerY(Attacker)) And (MapNPC(mapnum, MapNpcNum).X = GetPlayerX(Attacker)) Then
                    If NPC(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(npcnum).Behavior <> NPC_BEHAVIOR_SCRIPTED And NPC(npcnum).Behavior <> NPC_BEHAVIOR_QUEST And NPC(npcnum).Behavior <> NPC_BEHAVIOR_CHAOSKNIGHT Then
                        CanAttackNpc = True
                    Else
                        If NPC(npcnum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & NPC(npcnum).SpawnSecs
                        ElseIf NPC(npcnum).Behavior = NPC_BEHAVIOR_QUEST Then
                            Call DoQuest(NPC(npcnum).Quest, Attacker, npcnum)
                        ElseIf NPC(npcnum).Behavior = NPC_BEHAVIOR_CHAOSKNIGHT Then
                            Call DoQuestNpcKillsQuest(Attacker, npcnum)
                        Else
                              Call PlayerMsg(Attacker, Trim$(NPC(npcnum).Name) & " : " & Trim$(NPC(npcnum).AttackSay), Green)
                        End If
                    End If
                End If

            Case DIR_LEFT
                If (MapNPC(mapnum, MapNpcNum).Y = GetPlayerY(Attacker)) And (MapNPC(mapnum, MapNpcNum).X + 1 = GetPlayerX(Attacker)) Then
                    If NPC(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(npcnum).Behavior <> NPC_BEHAVIOR_SCRIPTED And NPC(npcnum).Behavior <> NPC_BEHAVIOR_QUEST And NPC(npcnum).Behavior <> NPC_BEHAVIOR_CHAOSKNIGHT Then
                        CanAttackNpc = True
                    Else
                        If NPC(npcnum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & NPC(npcnum).SpawnSecs
                        ElseIf NPC(npcnum).Behavior = NPC_BEHAVIOR_QUEST Then
                            Call DoQuest(NPC(npcnum).Quest, Attacker, npcnum)
                        ElseIf NPC(npcnum).Behavior = NPC_BEHAVIOR_CHAOSKNIGHT Then
                            Call DoQuestNpcKillsQuest(Attacker, npcnum)
                        Else
                            Call PlayerMsg(Attacker, Trim$(NPC(npcnum).Name) & " : " & Trim$(NPC(npcnum).AttackSay), Green)
                        End If
                    End If
                End If

            Case DIR_RIGHT
                If (MapNPC(mapnum, MapNpcNum).Y = GetPlayerY(Attacker)) And (MapNPC(mapnum, MapNpcNum).X - 1 = GetPlayerX(Attacker)) Then
                    If NPC(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(npcnum).Behavior <> NPC_BEHAVIOR_SCRIPTED And NPC(npcnum).Behavior <> NPC_BEHAVIOR_QUEST And NPC(npcnum).Behavior <> NPC_BEHAVIOR_CHAOSKNIGHT Then
                        CanAttackNpc = True
                    Else
                        If NPC(npcnum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & NPC(npcnum).SpawnSecs
                        ElseIf NPC(npcnum).Behavior = NPC_BEHAVIOR_QUEST Then
                            Call DoQuest(NPC(npcnum).Quest, Attacker, npcnum)
                        ElseIf NPC(npcnum).Behavior = NPC_BEHAVIOR_CHAOSKNIGHT Then
                            Call DoQuestNpcKillsQuest(Attacker, npcnum)
                        Else
                            Call PlayerMsg(Attacker, Trim$(NPC(npcnum).Name) & " : " & Trim$(NPC(npcnum).AttackSay), Green)
                        End If
                    End If
                End If
        End Select
    End If
End Function


Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Index As Long) As Boolean
    Dim mapnum As Long
    Dim npcnum As Long

    If Not IsPlaying(Index) Then
        Exit Function
    End If

    ' Make sure the NPC map number isn't out-of-range.
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Make sure that it's a valid NPC.
    If MapNPC(GetPlayerMap(Index), MapNpcNum).num < 1 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(Index)
    npcnum = MapNPC(mapnum, MapNpcNum).num

    ' Make sure that the NPC isn't already dead.
    If MapNPC(mapnum, MapNpcNum).HP < 1 Then
        Exit Function
    End If

    ' Make sure that NPCs don't attack more then once a second.
    If GetTickCount < MapNPC(mapnum, MapNpcNum).AttackTimer + 1000 Then
        Exit Function
    End If

    ' Make sure we don't attack a player if they are switching maps.
    If Player(Index).GettingMap = YES Then
        Exit Function
    End If

    MapNPC(mapnum, MapNpcNum).AttackTimer = GetTickCount

    If IsPlaying(Index) Then
        If npcnum > 0 Then
            If (GetPlayerY(Index) + 1 = MapNPC(mapnum, MapNpcNum).Y) And (GetPlayerX(Index) = MapNPC(mapnum, MapNpcNum).X) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(Index) - 1 = MapNPC(mapnum, MapNpcNum).Y) And (GetPlayerX(Index) = MapNPC(mapnum, MapNpcNum).X) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(Index) = MapNPC(mapnum, MapNpcNum).Y) And (GetPlayerX(Index) + 1 = MapNPC(mapnum, MapNpcNum).X) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(Index) = MapNPC(mapnum, MapNpcNum).Y) And (GetPlayerX(Index) - 1 = MapNPC(mapnum, MapNpcNum).X) Then
                            CanNpcAttackPlayer = True
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function

Sub AttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim EXP As Long
    Dim N As Long

    ' Make sure the attack is a valid index.
    If Not IsPlaying(Attacker) Then
        Exit Sub
    End If

    ' Make sure the victim is a valid index.
    If Not IsPlaying(Victim) Then
        Exit Sub
    End If

    ' Remove one SP point every time the player attacks.
    If SP_ATTACK = 1 Then
        If GetPlayerSP(Attacker) > 0 Then
            Call SetPlayerSP(Attacker, GetPlayerSP(Attacker) - 1)
            Call SendSP(Attacker)
        Else
            Call PlayerMsg(Attacker, "Te sientes agotado por la lucha.", Blue)
            Exit Sub
        End If
    End If
 
    ' If damage is below one, exit this sub routine.
    If Damage < 1 Then
        Exit Sub
    End If

    ' Check for weapon
    If GetPlayerWeaponSlot(Attacker) > 0 Then
        N = GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))
    End If

    ' Send this packet so they can see the person attacking
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), "ATTACK" & SEP_CHAR & Attacker & END_CHAR)

    If Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type <> TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type <> TILE_TYPE_ARENA Then
        If Damage >= GetPlayerHP(Victim) Then
            Call SetPlayerHP(Victim, 0)

            If scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "OnPVPDeath " & Attacker & "," & Victim
            Else
                Call GlobalMsg(GetPlayerName(Victim) & " ha sido asesinado por " & GetPlayerName(Attacker), BRIGHTRED)
            End If

            If Map(GetPlayerMap(Victim)).Moral <> MAP_MORAL_NO_PENALTY Then
                If scripting = 1 Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "DropItems " & Victim
                Else
                    If GetPlayerWeaponSlot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerWeaponSlot(Victim), 0)
                    End If

                    If GetPlayerArmorSlot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerArmorSlot(Victim), 0)
                    End If

                    If GetPlayerHelmetSlot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerHelmetSlot(Victim), 0)
                    End If

                    If GetPlayerShieldSlot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerShieldSlot(Victim), 0)
                    End If

                    If GetPlayerLegsSlot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerLegsSlot(Victim), 0)
                    End If

                    If GetPlayerRingSlot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerRingSlot(Victim), 0)
                    End If

                    If GetPlayerNecklaceSlot(Victim) > 0 Then
                        Call PlayerMapDropItem(Victim, GetPlayerNecklaceSlot(Victim), 0)
                    End If
                End If

                ' Calculate exp to give attacker
                EXP = Int(GetPlayerExp(Victim) / 10)

                ' Make sure we dont get less then 0
                If EXP < 0 Then
                    EXP = 0
                End If

                If GetPlayerLevel(Victim) = MAX_LEVEL Then
                    Call BattleMsg(Victim, "No puedes perder mas experiencia!", BRIGHTRED, 1)
                    Call BattleMsg(Attacker, GetPlayerName(Victim) & " está al maximo nivel!", BRIGHTBLUE, 0)
                Else
                    If EXP = 0 Then
                        Call SetPlayerExp(Victim, 0)
                        Call BattleMsg(Victim, "No pierdes ninguna experiencia.", BRIGHTRED, 1)
                        Call BattleMsg(Attacker, "No recibes ninguna experiencia.", BRIGHTBLUE, 0)
                    Else
                        Call SetPlayerExp(Victim, GetPlayerExp(Victim) - EXP)
                        Call BattleMsg(Victim, "Pierdes " & EXP & " de experiencia.", BRIGHTRED, 1)
                        Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + EXP)
                        Call BattleMsg(Attacker, "Ganas " & EXP & " de experiencia por matar a " & GetPlayerName(Victim) & ".", BRIGHTBLUE, 0)
                    End If
                    
                    Call SendEXP(Victim)
                    Call SendEXP(Attacker)
                End If
            End If

            ' Warp player away
            If scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & Victim
            Else
                If Map(GetPlayerMap(Victim)).BootMap > 0 Then
                    Call PlayerWarp(Victim, Map(GetPlayerMap(Victim)).BootMap, Map(GetPlayerMap(Victim)).BootX, Map(GetPlayerMap(Victim)).BootY)
                Else
                    Call PlayerWarp(Victim, START_MAP, START_X, START_Y)
                End If
            End If

            ' Restore vitals
            Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
            Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
            Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
            Call SendHP(Victim)
            Call SendMP(Victim)
            Call SendSP(Victim)

            ' Check for a level up
            Call CheckPlayerLevelUp(Attacker)

            ' Check if target is player who died and if so set target to 0
            If Player(Attacker).TargetType = TARGET_TYPE_PLAYER And Player(Attacker).Target = Victim Then
                Player(Attacker).Target = 0
                Player(Attacker).TargetType = 0
            End If

            If GetPlayerPK(Victim) = NO Then
                If GetPlayerPK(Attacker) = NO Then
                    Call SetPlayerPK(Attacker, YES)
                    Call SendPlayerData(Attacker)
                    Call GlobalMsg(GetPlayerName(Attacker) & " es considerado ahora un PK!", BRIGHTRED)
                End If
            Else
                Call SetPlayerPK(Victim, NO)
                Call SendPlayerData(Victim)
                Call GlobalMsg(GetPlayerName(Victim) & " ha pagado el precio por ser un asesino de jugadores!", BRIGHTRED)
            End If
        Else
            Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
            Call SendHP(Victim)
        End If
    ElseIf Map(GetPlayerMap(Attacker)).Tile(GetPlayerX(Attacker), GetPlayerY(Attacker)).Type = TILE_TYPE_ARENA And Map(GetPlayerMap(Victim)).Tile(GetPlayerX(Victim), GetPlayerY(Victim)).Type = TILE_TYPE_ARENA Then
        If Damage >= GetPlayerHP(Victim) Then
            Call SetPlayerHP(Victim, 0)

            ' Check if target is player who died and if so set target to 0
            If Player(Attacker).TargetType = TARGET_TYPE_PLAYER And Player(Attacker).Target = Victim Then
                Player(Attacker).Target = 0
                Player(Attacker).TargetType = 0
            End If

            If scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "OnArenaDeath " & Attacker & "," & Victim
            End If
        Else
            Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
            Call SendHP(Victim)
        End If
    End If

    Player(Attacker).AttackTimer = GetTickCount

    Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "pain" & SEP_CHAR & Player(Victim).Char(Player(Victim).CharNum).Sex & END_CHAR)
End Sub

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim EXP As Long
    Dim mapnum As Long

    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Sub
    End If
    
    If Not IsPlaying(Victim) Then
        Exit Sub
    End If

    If Damage < 1 Then
        Exit Sub
    End If

    If MapNPC(GetPlayerMap(Victim), MapNpcNum).num <= 0 Then
        Exit Sub
    End If

    ' Send this packet so they can see the person attacking
    Call SendDataToMap(GetPlayerMap(Victim), "NPCATTACK" & SEP_CHAR & MapNpcNum & END_CHAR)

    mapnum = GetPlayerMap(Victim)

    Name = Trim$(NPC(MapNPC(mapnum, MapNpcNum).num).Name)

    If Damage >= GetPlayerHP(Victim) Then
        Call GlobalMsg(GetPlayerName(Victim) & " fue matado por " & Name, BRIGHTRED)

        If Map(GetPlayerMap(Victim)).Moral <> MAP_MORAL_NO_PENALTY Then
            If scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "DropItems " & Victim
            Else
                If GetPlayerWeaponSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerWeaponSlot(Victim), 0)
                End If

                If GetPlayerArmorSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerArmorSlot(Victim), 0)
                End If

                If GetPlayerHelmetSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerHelmetSlot(Victim), 0)
                End If

                If GetPlayerShieldSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerShieldSlot(Victim), 0)
                End If

                If GetPlayerShieldSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerShieldSlot(Victim), 0)
                End If

                If GetPlayerLegsSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerLegsSlot(Victim), 0)
                End If

                If GetPlayerRingSlot(Victim) > 0 Then
                    Call PlayerMapDropItem(Victim, GetPlayerRingSlot(Victim), 0)
                End If
            End If

            ' Calculate exp to give attacker
            EXP = Int(GetPlayerExp(Victim) / 3)

            ' Make sure we dont get less then 0
            If EXP < 0 Then
                EXP = 0
            End If

            If EXP = 0 Then
                Call SetPlayerExp(Victim, 0)
                Call BattleMsg(Victim, "No puedes perder experiencia.", BRIGHTRED, 0)
            Else
                Call SetPlayerExp(Victim, GetPlayerExp(Victim) - EXP)
                Call BattleMsg(Victim, "Pierdes " & EXP & " de experiencia.", BRIGHTRED, 0)
            End If

            Call SendEXP(Victim)
        End If

        ' Warp player away
        If scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & Victim
        Else
            If Map(GetPlayerMap(Victim)).BootMap > 0 Then
                Call PlayerWarp(Victim, Map(GetPlayerMap(Victim)).BootMap, Map(GetPlayerMap(Victim)).BootX, Map(GetPlayerMap(Victim)).BootY)
            Else
                Call PlayerWarp(Victim, START_MAP, START_X, START_Y)
            End If
        End If

        ' Restore vitals
        Call SetPlayerHP(Victim, GetPlayerMaxHP(Victim))
        Call SetPlayerMP(Victim, GetPlayerMaxMP(Victim))
        Call SetPlayerSP(Victim, GetPlayerMaxSP(Victim))
        Call SendHP(Victim)
        Call SendMP(Victim)
        Call SendSP(Victim)

        ' Set NPC target to 0
        MapNPC(mapnum, MapNpcNum).Target = 0

        ' If the player the attacker killed was a pk then take it away
        If GetPlayerPK(Victim) = YES Then
            Call SetPlayerPK(Victim, NO)
            Call SendPlayerData(Victim)
        End If
    Else
        Call SetPlayerHP(Victim, GetPlayerHP(Victim) - Damage)
        Call SendHP(Victim)
    End If

    Call SendDataTo(Victim, "BLITNPCDMG" & SEP_CHAR & Damage & END_CHAR)
    Call SendDataToMap(GetPlayerMap(Victim), "sound" & SEP_CHAR & "pain" & SEP_CHAR & Player(Victim).Char(Player(Victim).CharNum).Sex & END_CHAR)
End Sub

Sub AttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long)
    Dim Name As String
    Dim EXP As Long
    Dim N As Long, I As Long, q As Integer, X As Long
    Dim mapnum As Long, npcnum As Long

    ' Removes one SP when you attack.
    If SP_ATTACK = 1 Then
        If GetPlayerSP(Attacker) > 0 Then
            Call SetPlayerSP(Attacker, GetPlayerSP(Attacker) - 1)
            Call SendSP(Attacker)
        Else
            Call PlayerMsg(Attacker, "Te sientes agotado por la lucha.", Blue)
            Exit Sub
        End If
    End If

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Or Damage < 0 Then
        Exit Sub
    End If

    ' Check for weapon
    If GetPlayerWeaponSlot(Attacker) > 0 Then
        N = GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))
    Else
        N = 0
    End If

    ' Send this packet so they can see the person attacking
    Call SendDataToMapBut(Attacker, GetPlayerMap(Attacker), "ATTACK" & SEP_CHAR & Attacker & END_CHAR)

    mapnum = GetPlayerMap(Attacker)
    npcnum = MapNPC(mapnum, MapNpcNum).num
    Name = Trim$(NPC(npcnum).Name)

    If Damage >= MapNPC(mapnum, MapNpcNum).HP Then
        ' Check for a weapon and say damage
        Player(Attacker).TargetNPC = 0
        
        If GetPlayerNpcKillQuestFlag(Attacker) > 0 And GetPlayerNpcKillQuestFlag(Attacker) < 2 Then
            Call DoQuestNpcKills(Attacker, npcnum)
        End If

        ' Call BattleMsg(Attacker, "You killed a " & Name, BrightRed, 0)
        If scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnNPCDeath " & Attacker & "," & mapnum & "," & npcnum & "," & MapNpcNum
        End If

        Dim Add As String

        Add = 0
        If GetPlayerWeaponSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AddEXP
        End If
        If GetPlayerArmorSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerArmorSlot(Attacker))).AddEXP
        End If
        If GetPlayerShieldSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerShieldSlot(Attacker))).AddEXP
        End If
        If GetPlayerLegsSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerLegsSlot(Attacker))).AddEXP
        End If
        If GetPlayerRingSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerRingSlot(Attacker))).AddEXP
        End If
        If GetPlayerNecklaceSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerNecklaceSlot(Attacker))).AddEXP
        End If
        If GetPlayerHelmetSlot(Attacker) > 0 Then
            Add = Add + Item(GetPlayerInvItemNum(Attacker, GetPlayerHelmetSlot(Attacker))).AddEXP
        End If

        If Add > 0 Then
            If Add < 100 Then
                If Add < 10 Then
                    Add = 0 & ".0" & Right$(Add, 2)
                Else
                    Add = 0 & "." & Right$(Add, 2)
                End If
            Else
                Add = Mid$(Add, 1, 1) & "." & Right$(Add, 2)
            End If
        End If

        ' Calculate exp to give attacker
        If Add > 0 Then
            EXP = NPC(npcnum).EXP + (NPC(npcnum).EXP * Val(Add))
        Else
            EXP = NPC(npcnum).EXP
        End If

        ' Make sure we dont get less then 0
        If EXP < 0 Then
            EXP = 1
        End If
        
        If Player(Attacker).InParty = YES Then
            Call CheckPlayerLevelUp(Player(Attacker).PartyPlayer)
        End If

        ' Check if in party, if so divide up the exp
        Dim o As Long
        If Player(Attacker).PartyID = 0 Then
            If GetPlayerLevel(Attacker) = MAX_LEVEL Then
                Call SetPlayerExp(Attacker, Experience(MAX_LEVEL))
                Call BattleMsg(Attacker, "No puedes ganar más experiencia!", BRIGHTBLUE, 0)
            Else
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + EXP)
                Call BattleMsg(Attacker, "Has ganado " & EXP & " puntos de experiencia.", BRIGHTBLUE, 0)
            End If
        Else
            o = 1
            For I = 1 To MAX_PARTY_MEMBERS

                If Party(Player(Attacker).PartyID).Member(I) <> Attacker Then
                    If Party(Player(Attacker).PartyID).Member(I) <> 0 Then
                        If GetPlayerMap(Attacker) = GetPlayerMap(Party(Player(Attacker).PartyID).Member(I)) Then
                            o = o + 1
                        End If
                    End If
                End If
            Next

            If GetPlayerLevel(Attacker) = MAX_LEVEL Then
                Call SetPlayerExp(Attacker, Experience(MAX_LEVEL))
                Call BattleMsg(Attacker, "No puedes ganar más experiencia!", BRIGHTBLUE, 0)
            Else

                If o > 1 Then
                    Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Int(EXP * 0.75))
                    Call BattleMsg(Attacker, "Has ganado " & Int(EXP * 0.75) & " puntos de experiencia y compartes " & Int(EXP * 0.25) & " puntos con tu grupo.", BRIGHTBLUE, 0)
                Else
                    Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + EXP)
                    Call BattleMsg(Attacker, "Has ganado " & EXP & " puntos de experiencia pero no compartes nada con tu grupo.", BRIGHTBLUE, 0)
                End If
            End If
            
                '6dragon6 partyid test - result: didn't work, works now
                'Call BattleMsg(Attacker, Player(Attacker).PartyID, 4, 0)
           '6dragon6 group member list test - result: didn't work, works now
            'For I = 1 To 4
            '    Call BattleMsg(Attacker, Party(Player(Attacker).PartyID).Member(I), 4, 0)
            'Next

            If o > 1 Then
                For I = 1 To o

                    If Party(Player(Attacker).PartyID).Member(I) <> Attacker And Party(Player(Attacker).PartyID).Member(I) <> 0 Then
                        If GetPlayerLevel(Party(Player(Attacker).PartyID).Member(I)) = MAX_LEVEL Then
                            Call SetPlayerExp(Party(Player(Attacker).PartyID).Member(I), Experience(MAX_LEVEL))
                            Call BattleMsg(Party(Player(Attacker).PartyID).Member(I), "No puedes ganar más experiencia!", BRIGHTBLUE, 0)
                        Else
                            Call SetPlayerExp(Party(Player(Attacker).PartyID).Member(I), Player(Party(Player(Attacker).PartyID).Member(I)).Char(Player(Party(Player(Attacker).PartyID).Member(I)).CharNum).EXP + Int(EXP * (0.25 / o)))
                            Call BattleMsg(Party(Player(Attacker).PartyID).Member(I), "Has ganado " & Int(EXP * (0.25 / o)) & " puntos de experiencia de tu grupo.", BRIGHTBLUE, 0)
                            'Call BattleMsg(Party(Player(Attacker).PartyID).Member(I), "PUTA", Red, 0)
                            Call SendStats(Party(Player(Attacker).PartyID).Member(I))
                            Call SendPlayerData(Party(Player(Attacker).PartyID).Member(I))
                        End If
                        
                    End If
                Next
            End If
        End If
        ' Drop the items if they earn it.
        For I = 1 To MAX_NPC_DROPS
            If NPC(npcnum).ItemNPC(I).ItemNum > 0 Then
                N = Int(Rnd * NPC(npcnum).ItemNPC(I).chance) + 1
                If N = 1 Then
                    Call SpawnItem(NPC(npcnum).ItemNPC(I).ItemNum, NPC(npcnum).ItemNPC(I).ItemValue, mapnum, MapNPC(mapnum, MapNpcNum).X, MapNPC(mapnum, MapNpcNum).Y)
                End If
            End If
        Next I

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNPC(mapnum, MapNpcNum).num = 0
        MapNPC(mapnum, MapNpcNum).SpawnWait = GetTickCount
        MapNPC(mapnum, MapNpcNum).HP = 0
        Call SendDataToMap(mapnum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)

        ' Check for level up
        Call CheckPlayerLevelUp(Attacker)

        
                ' Check for level up party member
         If Player(Attacker).InParty = YES Then
            For X = 1 To MAX_PARTY_MEMBERS

                If Party(Player(Attacker).PartyID).Member(X) <> 0 Then
                    Call CheckPlayerLevelUp(Party(Player(Attacker).PartyID).Member(X))
                End If
            Next
        End If

        ' Check for level up party member
        If Player(Attacker).InParty = True Then
            Call CheckPlayerLevelUp(Player(Attacker).PartyPlayer)
        End If

        ' Check if target is npc that died and if so set target to 0
        If Player(Attacker).TargetType = TARGET_TYPE_NPC And Player(Attacker).Target = MapNpcNum Then
            Player(Attacker).Target = 0
            Player(Attacker).TargetType = 0
        End If
    Else
        ' NPC not dead, just do the damage
        MapNPC(mapnum, MapNpcNum).HP = MapNPC(mapnum, MapNpcNum).HP - Damage
        Player(Attacker).TargetNPC = MapNpcNum

' Check for a weapon and say damage
' Call BattleMsg(Attacker, "You hit a " & Name & " for " & Damage & " damage.", White, 0)

        If N = 0 Then
        ' Call PlayerMsg(Attacker, "You hit a " & Name & " for " & Damage & " hit points.", White)
        Else
        ' Call PlayerMsg(Attacker, "You hit a " & Name & " with a " & Trim$(Item(n).Name) & " for " & Damage & " hit points.", White)
        End If

        ' Check if we should send a message
        If MapNPC(mapnum, MapNpcNum).Target = 0 And MapNPC(mapnum, MapNpcNum).Target <> Attacker Then
            If Trim$(NPC(npcnum).AttackSay) <> vbNullString Then
                Call PlayerMsg(Attacker, "A " & Trim$(NPC(npcnum).Name) & " : " & Trim$(NPC(npcnum).AttackSay) & vbNullString, SayColor)
            End If
        End If

        ' Set the NPC target to the player
        MapNPC(mapnum, MapNpcNum).Target = Attacker

        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNPC(mapnum, MapNpcNum).num).Behavior = NPC_BEHAVIOR_GUARD Then
            For I = 1 To MAX_MAP_NPCS
                If MapNPC(mapnum, I).num = MapNPC(mapnum, MapNpcNum).num Then
                    MapNPC(mapnum, I).Target = Attacker
                End If
            Next I
        End If
    End If

    Call SendDataToMap(mapnum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(mapnum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNPC(mapnum, MapNpcNum).num) & END_CHAR)

    ' Reset attack timer
    Player(Attacker).AttackTimer = GetTickCount
End Sub

Sub JoinWarp(ByVal Index As Long, ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long)
    Dim OldMap As Long

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Save current number map the player is on.
    OldMap = GetPlayerMap(Index)

    Call SendLeaveMap(Index, OldMap)

    Call SetPlayerMap(Index, mapnum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)

    ' Check to see if anyone is on the map.
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(mapnum) = YES

    Player(Index).GettingMap = YES

    Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & mapnum & SEP_CHAR & Map(mapnum).Revision & END_CHAR)

    Call SendInventory(Index)
    Call SendIndexWornEquipmentFromMap(Index)
End Sub

Sub PlayerWarp(ByVal Index As Long, ByVal mapnum As Long, ByVal X As Long, ByVal Y As Long)
    Dim OldMap As Long

    On Error GoTo WarpErr

    ' Check for subscript out of range.
    If Not IsPlaying(Index) Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Save current number map the player is on.
    OldMap = GetPlayerMap(Index)

    If Not OldMap = mapnum Then
        Call SendLeaveMap(Index, OldMap)
    End If

    Call SetPlayerMap(Index, mapnum)
    Call SetPlayerX(Index, X)
    Call SetPlayerY(Index, Y)

    ' Check to see if anyone is on the map.
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If
    
    If Player(Index).Pet.Alive = YES Then
        Player(Index).Pet.MapToGo = mapnum
        Player(Index).Pet.Map = mapnum
        Player(Index).Pet.X = X
        Player(Index).Pet.Y = Y
        Call SetPlayerPetX(Index, X)
        Call SetPlayerPetY(Index, Y)
        Call SetPlayerPetMap(Index, mapnum)
    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(mapnum) = YES

    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "warp" & END_CHAR)

    Player(Index).GettingMap = YES

    Call SendDataTo(Index, "CHECKFORMAP" & SEP_CHAR & mapnum & SEP_CHAR & Map(mapnum).Revision & END_CHAR)

    Call SendInventory(Index)
    Call SendIndexInventoryFromMap(Index)
    Call SendIndexWornEquipmentFromMap(Index)

    If scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "OnMapLoad " & Index & "," & OldMap & "," & mapnum
    End If
    
    Exit Sub

WarpErr:
    Call AddLog("PlayerWarp error for player index " & Index & " on map " & GetPlayerMap(Index) & ".", "logs\ErrorLog.txt")
End Sub

Sub PlayerMove(ByVal Index As Long, ByVal Dir As Long, ByVal Movement As Long, Xpos As Integer, Ypos As Integer)
    Dim Packet As String
    Dim mapnum As Long
    Dim X As Long
    Dim Y As Long
    Dim I As Long
    Dim Moved As Byte
    Dim sheet As Long
    Dim a As Long
    '   variables we will need
    Dim pmap As Integer
    Dim Xold As Integer
    Dim Yold As Integer

    ' They tried to hack
    ' If Moved = NO Then
    ' Call HackingAttempt(index, "Position Modification")
    ' Exit Sub
    ' End If

    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If
    
    If Player(Index).GettingMap = True Then
        Exit Sub
    End If

    ' Check for scrolling to prevent RTE 9
    If GetPlayerX(Index) > MAX_MAPX Or GetPlayerY(Index) > MAX_MAPY Then
        Call PlayerWarp(Index, GetPlayerMap(Index), 0, 0)
        Exit Sub
    End If

    Call SetPlayerDir(Index, Dir)
    
    If Player(Index).Pet.Alive = YES Then

    
        If Player(Index).Pet.Map = GetPlayerMap(Index) And Player(Index).Pet.X = X And Player(Index).Pet.Y = Y Then
           ' If Grid(GetPlayerMap(Index)).Loc(DirToX(x, Dir), DirToY(y, Dir)).Blocked = False Then
             '   Call UpdateGrid(Player(Index).Pet.Map, Player(Index).Pet.x, Player(Index).Pet.y, Player(Index).Pet.Map, DirToX(x, Dir), DirToY(y, Dir))
                Player(Index).Pet.Y = DirToY(Y, Dir)
                Player(Index).Pet.X = DirToX(X, Dir)
                Packet = "PETMOVE" & SEP_CHAR & Index & SEP_CHAR & DirToX(X, Dir) & SEP_CHAR & DirToY(Y, Dir) & SEP_CHAR & Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                Call SendDataToMap(Player(Index).Pet.Map, Packet)
           ' End If
        End If
    End If

    ' Remove SP if the player is running.
    If SP_RUNNING = 1 Then
        If Movement = MOVING_RUNNING Then
            If GetPlayerSP(Index) > 0 Then
                Call SetPlayerSP(Index, GetPlayerSP(Index) - 1)
                Call SendSP(Index)
            Else
                Call PlayerMsg(Index, "Te sientes cansado para correr.", Blue)
            End If
        End If
    End If

    Moved = NO


' save the current location
    Xold = GetPlayerX(Index)
    Yold = GetPlayerY(Index)
    pmap = GetPlayerMap(Index)
   
' validate map number
   
    If pmap <= 0 Or pmap > MAX_MAPS Then
        Call HackingAttempt(Index, vbNullString)
        Exit Sub
    End If
   
   
' update it to match client - this will be correct 99% of the time
    Call SetPlayerX(Index, Xpos)
    Call SetPlayerY(Index, Ypos)
   
' next check to see if we have gone outside of map boundries
'   if we have, need to try to warp to next map if there is one

    If Dir = DIR_UP And Ypos < 0 And Map(pmap).Up > 0 Then
        Call PlayerWarp(Index, Map(pmap).Up, Xpos, MAX_MAPY)
        Moved = YES
    ElseIf Dir = DIR_DOWN And Ypos > MAX_MAPY And Map(pmap).Down > 0 Then
        Call PlayerWarp(Index, Map(pmap).Down, Xpos, 0)
        Moved = YES
    ElseIf Dir = DIR_LEFT And Xpos < 0 And Map(pmap).Left > 0 Then
        Call PlayerWarp(Index, Map(pmap).Left, MAX_MAPX, Ypos)
        Moved = YES
    ElseIf Dir = DIR_RIGHT And Xpos > MAX_MAPX And Map(pmap).Right > 0 Then
        Call PlayerWarp(Index, Map(pmap).Right, 0, Ypos)
        Moved = YES
    End If
   
' restore values in case we got warped

    Xpos = GetPlayerX(Index)
    Ypos = GetPlayerY(Index)
    pmap = GetPlayerMap(Index)

' check to make sure new position is on the map

    If Xpos < 0 Or Ypos < 0 Or Xpos > MAX_MAPX Or Ypos > MAX_MAPY Then
        Call HackingAttempt(Index, vbNullString)
        Exit Sub
    End If

' Check to make sure that the tile is walkable
    If Map(pmap).Tile(Xpos, Ypos).Type <> TILE_TYPE_BLOCKED Then
' Check to see if the tile is a key and if it is check if its opened
        If (Map(pmap).Tile(Xpos, Ypos).Type <> TILE_TYPE_KEY And Map(pmap).Tile(Xpos, Ypos).Type <> TILE_TYPE_DOOR) Or ((Map(pmap).Tile(Xpos, Ypos).Type = TILE_TYPE_DOOR Or Map(pmap).Tile(Xpos, Ypos).Type = TILE_TYPE_KEY) And TempTile(pmap).DoorOpen(Xpos, Ypos) = YES) Then
            Packet = "playermove" & SEP_CHAR & Index & SEP_CHAR & Xpos & SEP_CHAR & Ypos & SEP_CHAR & Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMapBut(Index, pmap, Packet)
            Moved = YES
        End If
    End If

' at this point we have either moved or there is a problem with the new location
'   if we didn't move, we need to reset to previous locations and quit

    If Moved <> YES Then
        Call SetPlayerX(Index, Xold)
        Call SetPlayerY(Index, Yold)
        Call SendPlayerNewXY(Index)
        Exit Sub
    End If

    ' healing tiles code
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_HEAL Then
        Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
        Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
        Call SendHP(Index)
        Call SendMP(Index)
        Call PlayerMsg(Index, "Sientes como una fuerza sanadora entra dentro de ti!", BRIGHTGREEN)
    End If

    ' Check for kill tile, and if so kill them
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_KILL Then
        Call SetPlayerHP(Index, 0)
        Call PlayerMsg(Index, "Sientes como la muerte penetra en tu cuerpo.", BRIGHTRED)

        ' Warp player away
        If scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & Index
        Else
            If Map(GetPlayerMap(Index)).BootMap > 0 Then
                Call PlayerWarp(Index, Map(GetPlayerMap(Index)).BootMap, Map(GetPlayerMap(Index)).BootX, Map(GetPlayerMap(Index)).BootY)
            Else
                Call PlayerWarp(Index, START_MAP, START_X, START_Y)
            End If
        End If
        Call SetPlayerHP(Index, GetPlayerMaxHP(Index))
        Call SetPlayerMP(Index, GetPlayerMaxMP(Index))
        Call SetPlayerSP(Index, GetPlayerMaxSP(Index))
        Call SendHP(Index)
        Call SendMP(Index)
        Call SendSP(Index)
        Moved = YES
    End If

    If GetPlayerX(Index) + 1 <= MAX_MAPX Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) + 1, GetPlayerY(Index)).Type = TILE_TYPE_DOOR Then
            X = GetPlayerX(Index) + 1
            Y = GetPlayerY(Index)

            If TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount

                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "door" & END_CHAR)
            End If
        End If
    End If
    If GetPlayerX(Index) - 1 >= 0 Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index) - 1, GetPlayerY(Index)).Type = TILE_TYPE_DOOR Then
            X = GetPlayerX(Index) - 1
            Y = GetPlayerY(Index)

            If TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount

                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "door" & END_CHAR)
            End If
        End If
    End If
    If GetPlayerY(Index) - 1 >= 0 Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) - 1).Type = TILE_TYPE_DOOR Then
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) - 1

            If TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount

                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "door" & END_CHAR)
            End If
        End If
    End If
    If GetPlayerY(Index) + 1 <= MAX_MAPY Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index) + 1).Type = TILE_TYPE_DOOR Then
            X = GetPlayerX(Index)
            Y = GetPlayerY(Index) + 1

            If TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
                TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
                TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount

                Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "door" & END_CHAR)
            End If
        End If
    End If

    ' Check to see if the tile is a warp tile, and if so warp them
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_WARP Then
        mapnum = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        X = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2
        Y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3

        Call PlayerWarp(Index, mapnum, X, Y)
        Moved = YES
    End If

    ' Check for key trigger open
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_KEYOPEN Then
        X = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        Y = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2

        If Map(GetPlayerMap(Index)).Tile(X, Y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = NO Then
            TempTile(GetPlayerMap(Index)).DoorOpen(X, Y) = YES
            TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount

            Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & 1 & END_CHAR)
            If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) = vbNullString Then
                Call MapMsg(GetPlayerMap(Index), "La puerta ha sido abierta!", WHITE)
            Else
                Call MapMsg(GetPlayerMap(Index), Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), WHITE)
            End If
            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "key" & END_CHAR)
        End If
    End If

    ' Check for shop
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SHOP Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 > 0 Then
            Call SendTrade(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)
        Else
            Call PlayerMsg(Index, "No hay una tienda aqui.", BRIGHTRED)
        End If
    End If

    ' Check if player stepped on sprite changing tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SPRITE_CHANGE Then
        If GetPlayerSprite(Index) = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 Then
            Call PlayerMsg(Index, "Ya tienes este sprite!", BRIGHTRED)
            Exit Sub
        Else
            If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 = 0 Then
                Call SendDataTo(Index, "spritechange" & SEP_CHAR & 0 & END_CHAR)
            Else
                If Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2).Type = ITEM_TYPE_CURRENCY Then
                    Call PlayerMsg(Index, "Este sprite te costara " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3 & " " & Trim$(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2).Name) & "!", YELLOW)
                Else
                    Call PlayerMsg(Index, "Este sprite te costara " & Trim$(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2).Name) & "!", YELLOW)
                End If
                Call SendDataTo(Index, "spritechange" & SEP_CHAR & 1 & END_CHAR)
            End If
        End If
    End If

    ' Check if player stepped on house buying tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_HOUSE Then
        If Len(Map(GetPlayerMap(Index)).Owner) < 2 Then
            If GetPlayerName(Index) = Map(GetPlayerMap(Index)).Owner Then
                Call PlayerMsg(Index, "Ya tienes esta casa!", BRIGHTRED)
                Exit Sub
            Else
                If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 = 0 Then
                    Call SendDataTo(Index, "housebuy" & SEP_CHAR & 0 & END_CHAR)
                Else
                    If Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).Type = ITEM_TYPE_CURRENCY Then
                        Call PlayerMsg(Index, "Esta casa te costara " & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 & " " & Trim$(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).Name) & "!", YELLOW)
                    Else
                        Call PlayerMsg(Index, "Esta casa te costara un " & Trim$(Item(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).Name) & "!", YELLOW)
                    End If
                    Call SendDataTo(Index, "housebuy" & SEP_CHAR & 1 & END_CHAR)
                End If
            End If
        Else
            Call PlayerMsg(Index, "Esta casa no esta en venta!", BRIGHTRED)
            Exit Sub
        End If
    End If

    ' Check if player stepped on sprite changing tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_CLASS_CHANGE Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 > -1 Then
            If GetPlayerClass(Index) <> Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2 Then
                Call PlayerMsg(Index, "No eres la clase requerida!", BRIGHTRED)
                Exit Sub
            End If
        End If

        If GetPlayerClass(Index) = Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1 Then
            Call PlayerMsg(Index, "Ya eres esta clase!", BRIGHTRED)
        Else
            If Player(Index).Char(Player(Index).CharNum).Sex = 0 Then
                If GetPlayerSprite(Index) = ClassData(GetPlayerClass(Index)).MaleSprite Then
                    Call SetPlayerSprite(Index, ClassData(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).MaleSprite)
                End If
            Else
                If GetPlayerSprite(Index) = ClassData(GetPlayerClass(Index)).FemaleSprite Then
                    Call SetPlayerSprite(Index, ClassData(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1).FemaleSprite)
                End If
            End If

            Call SetPlayerSTR(Index, (Player(Index).Char(Player(Index).CharNum).STR - ClassData(GetPlayerClass(Index)).STR))
            Call SetPlayerDEF(Index, (Player(Index).Char(Player(Index).CharNum).DEF - ClassData(GetPlayerClass(Index)).DEF))
            Call SetPlayerMAGI(Index, (Player(Index).Char(Player(Index).CharNum).Magi - ClassData(GetPlayerClass(Index)).Magi))
            Call SetPlayerSPEED(Index, (Player(Index).Char(Player(Index).CharNum).Speed - ClassData(GetPlayerClass(Index)).Speed))

            Call SetPlayerClassData(Index, Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1)

            Call SetPlayerSTR(Index, (Player(Index).Char(Player(Index).CharNum).STR + ClassData(GetPlayerClass(Index)).STR))
            Call SetPlayerDEF(Index, (Player(Index).Char(Player(Index).CharNum).DEF + ClassData(GetPlayerClass(Index)).DEF))
            Call SetPlayerMAGI(Index, (Player(Index).Char(Player(Index).CharNum).Magi + ClassData(GetPlayerClass(Index)).Magi))
            Call SetPlayerSPEED(Index, (Player(Index).Char(Player(Index).CharNum).Speed + ClassData(GetPlayerClass(Index)).Speed))


            Call PlayerMsg(Index, "Tu nueva clase es " & Trim$(ClassData(GetPlayerClass(Index)).Name) & "!", BRIGHTGREEN)

            Call SendStats(Index)
            Call SendHP(Index)
            Call SendMP(Index)
            Call SendSP(Index)
            Player(Index).Char(Player(Index).CharNum).MAXHP = GetPlayerMaxHP(Index)
            Player(Index).Char(Player(Index).CharNum).MAXMP = GetPlayerMaxMP(Index)
            Player(Index).Char(Player(Index).CharNum).MAXSP = GetPlayerMaxSP(Index)
            Call SendDataToMap(GetPlayerMap(Index), "checksprite" & SEP_CHAR & Index & SEP_CHAR & GetPlayerSprite(Index) & END_CHAR)
        End If
    End If

    ' Check if player stepped on notice tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_NOTICE Then
        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) <> vbNullString Then
            Call PlayerMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), BLACK)
        End If
        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String2) <> vbNullString Then
            Call PlayerMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String2), GREY)
        End If
        If Not Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String3 = vbNullString Or Not Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String3 = vbNullString Then
            Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String3 & END_CHAR)
        End If
    End If

    ' Check if player steppted on minus stat tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_LOWER_STAT Then
        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1) <> vbNullString Then
            Call PlayerMsg(Index, Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1), BLACK)
        End If
        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1) <> 0 Then
            Call SetPlayerHP(Index, GetPlayerHP(Index) - Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1))
        End If
        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2) <> 0 Then
            Call SetPlayerMP(Index, GetPlayerMP(Index) - Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data2))
        End If
        If Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3) <> 0 Then
            Call SetPlayerSP(Index, GetPlayerSP(Index) - Trim$(Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data3))
        End If
    End If

    ' Check if player stepped on sound tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SOUND Then
        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).String1 & END_CHAR)
    End If

    If scripting = 1 Then
        If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_SCRIPTED Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & Index & "," & Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Data1
        End If
    End If

    ' Check if player stepped on Bank tile
    If Map(GetPlayerMap(Index)).Tile(GetPlayerX(Index), GetPlayerY(Index)).Type = TILE_TYPE_BANK Then
        Call SendDataTo(Index, "openbank" & END_CHAR)
    End If

End Sub

Function CanNpcMove(ByVal mapnum As Long, ByVal MapNpcNum As Long, ByVal Dir) As Boolean
    Dim I As Long
    Dim TileType As Long
    Dim X As Long
    Dim Y As Long

    ' Check for sub-script out of range.
    If mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Function
    End If

    ' Check for sub-script out of range.
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check for sub-script out of range.
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Function
    End If

    X = MapNPC(mapnum, MapNpcNum).X
    Y = MapNPC(mapnum, MapNpcNum).Y

    CanNpcMove = True

    Select Case Dir
        Case DIR_UP
            If Y > 0 Then
                ' Get the attribute on the tile.
                TileType = Map(mapnum).Tile(X, Y - 1).Type
                                
                ' Check to make sure that the tile is walkable.
                If TileType = TILE_TYPE_BLOCKED Or TileType = TILE_TYPE_NPCAVOID Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way.
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) Then
                        If GetPlayerMap(I) = mapnum Then
                            If GetPlayerX(I) = MapNPC(mapnum, MapNpcNum).X Then
                                If GetPlayerY(I) = (MapNPC(mapnum, MapNpcNum).Y - 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next I

                ' Check to make sure that there is not another npc in the way.
                For I = 1 To MAX_MAP_NPCS
                    If I <> MapNpcNum Then
                        If MapNPC(mapnum, I).num > 0 Then
                            If MapNPC(mapnum, I).X = MapNPC(mapnum, MapNpcNum).X Then
                                If MapNPC(mapnum, I).Y = (MapNPC(mapnum, MapNpcNum).Y - 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next I
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN
            If Y < MAX_MAPY Then
                ' Get the attribute on the tile.
                TileType = Map(mapnum).Tile(X, Y + 1).Type
                
                ' Check to make sure that the tile is walkable.
                If TileType = TILE_TYPE_BLOCKED Or TileType = TILE_TYPE_NPCAVOID Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way.
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) Then
                        If GetPlayerMap(I) = mapnum Then
                            If GetPlayerX(I) = MapNPC(mapnum, MapNpcNum).X Then
                                If GetPlayerY(I) = (MapNPC(mapnum, MapNpcNum).Y + 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next I

                ' Check to make sure that there is not another npc in the way.
                For I = 1 To MAX_MAP_NPCS
                    If I <> MapNpcNum Then
                        If MapNPC(mapnum, I).num > 0 Then
                            If MapNPC(mapnum, I).X = MapNPC(mapnum, MapNpcNum).X Then
                                If MapNPC(mapnum, I).Y = (MapNPC(mapnum, MapNpcNum).Y + 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next I
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT
            If X > 0 Then
                ' Get the attribute on the tile.
                TileType = Map(mapnum).Tile(X - 1, Y).Type

                ' Check to make sure that the tile is walkable.
                If TileType = TILE_TYPE_BLOCKED Or TileType = TILE_TYPE_NPCAVOID Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way.
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) Then
                        If GetPlayerMap(I) = mapnum Then
                            If GetPlayerX(I) = (MapNPC(mapnum, MapNpcNum).X - 1) Then
                                If GetPlayerY(I) = MapNPC(mapnum, MapNpcNum).Y Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next I

                ' Check to make sure that there is not another npc in the way.
                For I = 1 To MAX_MAP_NPCS
                    If I <> MapNpcNum Then
                        If MapNPC(mapnum, I).num > 0 Then
                            If MapNPC(mapnum, I).X = (MapNPC(mapnum, MapNpcNum).X - 1) Then
                                If MapNPC(mapnum, I).Y = MapNPC(mapnum, MapNpcNum).Y Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next I
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT
            If X < MAX_MAPX Then
                ' Get the attribute on the tile.
                TileType = Map(mapnum).Tile(X + 1, Y).Type
                
                ' Check to make sure that the tile is walkable.
                If TileType = TILE_TYPE_BLOCKED Or TileType = TILE_TYPE_NPCAVOID Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way.
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) Then
                        If GetPlayerMap(I) = mapnum Then
                            If GetPlayerX(I) = (MapNPC(mapnum, MapNpcNum).X + 1) Then
                                If GetPlayerY(I) = MapNPC(mapnum, MapNpcNum).Y Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next I

                ' Check to make sure that there is not another npc in the way.
                For I = 1 To MAX_MAP_NPCS
                    If I <> MapNpcNum Then
                        If MapNPC(mapnum, I).num > 0 Then
                            If MapNPC(mapnum, I).X = (MapNPC(mapnum, MapNpcNum).X + 1) Then
                                If MapNPC(mapnum, I).Y = MapNPC(mapnum, MapNpcNum).Y Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next I
            Else
                CanNpcMove = False
            End If
    End Select
End Function

Public Sub NpcMove(ByVal mapnum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long, ByVal Movement As Long)
    ' Check to make sure it's a valid map.
    If mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check to make sure it's a valid NPC.
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Sub
    End If

    ' Check to make sure it's a valid direction.
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    ' Check to make sure it's a valid movement speed.
    If Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If

    MapNPC(mapnum, MapNpcNum).Dir = Dir

    Select Case Dir
        Case DIR_UP
            MapNPC(mapnum, MapNpcNum).Y = MapNPC(mapnum, MapNpcNum).Y - 1

        Case DIR_DOWN
            MapNPC(mapnum, MapNpcNum).Y = MapNPC(mapnum, MapNpcNum).Y + 1

        Case DIR_LEFT
            MapNPC(mapnum, MapNpcNum).X = MapNPC(mapnum, MapNpcNum).X - 1

        Case DIR_RIGHT
            MapNPC(mapnum, MapNpcNum).X = MapNPC(mapnum, MapNpcNum).X + 1
    End Select

    Call SendDataToMap(mapnum, "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(mapnum, MapNpcNum).X & SEP_CHAR & MapNPC(mapnum, MapNpcNum).Y & SEP_CHAR & MapNPC(mapnum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR)
End Sub

Public Sub NpcDir(ByVal mapnum As Long, ByVal MapNpcNum As Long, ByVal Dir As Long)
    ' Check to make sure it's a valid map.
    If mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Check to make sure it's a valid NPC.
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Sub
    End If

    ' Check to make sure it's a valid direction.
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then
        Exit Sub
    End If

    MapNPC(mapnum, MapNpcNum).Dir = Dir

    Call SendDataToMap(mapnum, "NPCDIR" & SEP_CHAR & MapNpcNum & SEP_CHAR & Dir & END_CHAR)
End Sub
Public Sub JoinGame(ByVal Index As Long)
    Dim MOTD As String
    Dim FileData As String

    ' Set the flag so we know the person is in the game
    Player(Index).InGame = True

    ' Send an ok to client to start receiving in game data
    Call SendDataTo(Index, "loginok" & SEP_CHAR & Index & END_CHAR)

    Call CheckEquippedItems(Index)
    Call SendClasses(Index)
    Call SendItems(Index)
    Call SendEmoticons(Index)
    Call SendElements(Index)
    Call SendArrows(Index)
    Call SendNpcs(Index)
    Call SendShops(Index)
    Call SendQuest(Index)
    Call SendSpells(Index)
    Call SendInventory(Index)
    Call SendBank(Index)
    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
    Call SendPTS(Index)
    Call SendStats(Index)
    Call SendWeatherTo(Index)
    Call SendTimeTo(Index)
    Call SendGameClockTo(Index)
    Call DisabledTimeTo(Index)
    Call SendSprite(Index, Index)
    Call SendPlayerSpells(Index)
    Call SendOnlineList
    Call SendPlayerQuestFlags(Index)

    ' Warp the player to his saved location
    Call PlayerWarp(Index, GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index))

    If scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "JoinGame " & Index
    Else
        ' Send a global message that he/she joined.
        If GetPlayerAccess(Index) = 0 Then
            Call GlobalMsg(GetPlayerName(Index) & " ha entrado en " & GAME_NAME & "!", 7)
        Else
            Call GlobalMsg(GetPlayerName(Index) & " ha entrado en " & GAME_NAME & "!", 15)
        End If

        Call PlayerMsg(Index, "Bienvenido a " & GAME_NAME & "!", 15)

        ' Send the player the welcome message.
        MOTD = Trim$(GetVar(App.Path & "\MOTD.ini", "MOTD", "Msg"))
        If LenB(MOTD) <> 0 Then
            Call PlayerMsg(Index, "MOTD: " & MOTD, 11)
        End If

        ' Update all clients with the player.
        Call SendWhosOnline(Index)
    End If

    ' Tell the client the player is in-game.
    Call SendDataTo(Index, "ingame" & END_CHAR)

    ' Update the server console.
    Call ShowPLR(Index)
    
    If IsPetAliveOnLogin(Index) > 0 Then
       Call SpawnPet(Index)
    End If
    
    FileData = ReadINI("SK1", "sid", App.Path & "\Scripts\db\" & GetPlayerName(Index) & ".ini", vbNullString)
    Call SendDataTo(Index, "getspell1" & SEP_CHAR & FileData & END_CHAR)
    FileData = ReadINI("SK2", "sid", App.Path & "\Scripts\db\" & GetPlayerName(Index) & ".ini", vbNullString)
    Call SendDataTo(Index, "getspell2" & SEP_CHAR & FileData & END_CHAR)
    FileData = ReadINI("SK3", "sid", App.Path & "\Scripts\db\" & GetPlayerName(Index) & ".ini", vbNullString)
    Call SendDataTo(Index, "getspell3" & SEP_CHAR & FileData & END_CHAR)
    FileData = ReadINI("SK4", "sid", App.Path & "\Scripts\db\" & GetPlayerName(Index) & ".ini", vbNullString)
    Call SendDataTo(Index, "getspell4" & SEP_CHAR & FileData & END_CHAR)
    FileData = ReadINI("SK5", "sid", App.Path & "\Scripts\db\" & GetPlayerName(Index) & ".ini", vbNullString)
    Call SendDataTo(Index, "getspell5" & SEP_CHAR & FileData & END_CHAR)
    FileData = ReadINI("SK6", "sid", App.Path & "\Scripts\db\" & GetPlayerName(Index) & ".ini", vbNullString)
    Call SendDataTo(Index, "getspell6" & SEP_CHAR & FileData & END_CHAR)
    FileData = ReadINI("SK7", "sid", App.Path & "\Scripts\db\" & GetPlayerName(Index) & ".ini", vbNullString)
    Call SendDataTo(Index, "getspell7" & SEP_CHAR & FileData & END_CHAR)
    FileData = ReadINI("SK8", "sid", App.Path & "\Scripts\db\" & GetPlayerName(Index) & ".ini", vbNullString)
    Call SendDataTo(Index, "getspell8" & SEP_CHAR & FileData & END_CHAR)
    FileData = ReadINI("SK9", "sid", App.Path & "\Scripts\db\" & GetPlayerName(Index) & ".ini", vbNullString)
    Call SendDataTo(Index, "getspell9" & SEP_CHAR & FileData & END_CHAR)
    FileData = ReadINI("SK10", "sid", App.Path & "\Scripts\db\" & GetPlayerName(Index) & ".ini", vbNullString)
    Call SendDataTo(Index, "getspell10" & SEP_CHAR & FileData & END_CHAR)
End Sub

Public Sub LeftGame(ByVal Index As Long)
    Dim N As Long

    If Player(Index).InGame Then
        Player(Index).InGame = False
        If GetPlayerParty(Index) > 0 Then Call PartyRemoval(Index, GetPlayerParty(Index), Trim$(GetPlayerName(Index)))

        ' Stop processing NPCs if no one is on it.
        If GetTotalMapPlayers(GetPlayerMap(Index)) = 0 Then
            PlayersOnMap(GetPlayerMap(Index)) = NO
        End If
        
                ' Check if the player was in a party, and if so cancel it out so the other player doesn't continue to get half exp
        If Player(Index).InParty = YES Then
            N = Player(Index).PartyPlayer
            
            Call PlayerMsg(N, GetPlayerName(Index) & " ha salido de " & GAME_NAME & ", deshaciendose el grupo actual.", PINK)
            Player(N).InParty = NO
            Player(N).PartyPlayer = 0
        End If

        
        If Player(Index).Pet.Alive = YES Then
          ' Call TakeFromGrid(Player(Index).Pet.Map, Player(Index).Pet.x, Player(Index).Pet.y)
           Call savepet(Index)
        End If
        

        If scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "LeftGame " & Index
        Else
            ' Check to see if there is any boot map data.
            If Map(GetPlayerMap(Index)).BootMap > 0 Then
                Call SetPlayerX(Index, Map(GetPlayerMap(Index)).BootX)
                Call SetPlayerY(Index, Map(GetPlayerMap(Index)).BootY)
                Call SetPlayerMap(Index, Map(GetPlayerMap(Index)).BootMap)
            End If

            ' Inform the server that the player logged off.
            If GetPlayerAccess(Index) = 0 Then
                Call GlobalMsg(GetPlayerName(Index) & " ha salido de " & GAME_NAME & "!", 7)
            Else
                Call GlobalMsg(GetPlayerName(Index) & " ha salido de " & GAME_NAME & "!", 15)
            End If
        End If

        Call SavePlayer(Index)
        Call SendLeftGame(Index)

        Call TextAdd(frmServer.txtText(0), GetPlayerName(Index) & " se ha desconectado de " & GAME_NAME & ".", True)

        Call RemovePLR(Index)
    End If

    Call ClearPlayer(Index)
    Call SendOnlineList
End Sub

Function GetTotalMapPlayers(ByVal mapnum As Long) As Long
    Dim I As Long

    If mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Function
    End If

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            If GetPlayerMap(I) = mapnum Then
                GetTotalMapPlayers = GetTotalMapPlayers + 1
            End If
        End If
    Next I
End Function

Function GetNpcMaxHP(ByVal npcnum As Long) As Long
    If npcnum < 1 Or npcnum > MAX_NPCS Then
        Exit Function
    End If

    GetNpcMaxHP = NPC(npcnum).MAXHP
End Function

Function GetNpcMaxMP(ByVal npcnum As Long) As Long
    If npcnum < 1 Or npcnum > MAX_NPCS Then
        Exit Function
    End If

    GetNpcMaxMP = NPC(npcnum).Magi * 2
End Function

Function GetNpcMaxSP(ByVal npcnum As Long) As Long
    If npcnum < 1 Or npcnum > MAX_NPCS Then
        Exit Function
    End If

    GetNpcMaxSP = NPC(npcnum).Speed * 2
End Function

Function GetPlayerHPRegen(ByVal Index As Long) As Integer
    Dim Total As Integer

    If HP_REGEN = 1 Then
        If Index < 1 Or Index > MAX_PLAYERS Then
            Exit Function
        End If

        If Not IsPlaying(Index) Then
            Exit Function
        End If

        Total = Int(GetPlayerDEF(Index) / 2)
        If Total < 2 Then
            Total = 2
        End If

        GetPlayerHPRegen = Total
    End If
End Function

Function GetPlayerMPRegen(ByVal Index As Long) As Integer
    Dim Total As Integer

    If MP_REGEN = 1 Then
        If Index < 1 Or Index > MAX_PLAYERS Then
            Exit Function
        End If

        If Not IsPlaying(Index) Then
            Exit Function
        End If

        Total = Int(GetPlayerMAGI(Index) / 2)
        If Total < 2 Then
            Total = 2
        End If

        GetPlayerMPRegen = Total
    End If
End Function

Function GetPlayerSPRegen(ByVal Index As Long) As Integer
    Dim Total As Integer

    If SP_REGEN = 1 Then
        If Index < 1 Or Index > MAX_PLAYERS Then
            Exit Function
        End If

        If Not IsPlaying(Index) Then
            Exit Function
        End If

        Total = Int(GetPlayerSPEED(Index) / 2)
        If Total < 2 Then
            Total = 2
        End If

        GetPlayerSPRegen = Total
    End If
End Function

Function GetNpcHPRegen(ByVal npcnum As Long) As Integer
    Dim Total As Integer

    If NPC_REGEN = 1 Then
        If npcnum < 1 Or npcnum > MAX_NPCS Then
            Exit Function
        End If
    
        Total = Int(NPC(npcnum).DEF / 3)
        If Total < 1 Then
            Total = 1
        End If
    
        GetNpcHPRegen = Total
    End If
End Function

Sub CheckPlayerLevelUp(ByVal Index As Long)
    Dim I As Long
    Dim d As Long
    Dim c As Long
    c = 0

    If GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then
        If GetPlayerLevel(Index) < MAX_LEVEL Then
            If scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerLevelUp " & Index
            Else
                Do Until GetPlayerExp(Index) < GetPlayerNextLevel(Index)
                    DoEvents
                    If GetPlayerLevel(Index) < MAX_LEVEL Then
                        If GetPlayerExp(Index) >= GetPlayerNextLevel(Index) Then
                            d = GetPlayerExp(Index) - GetPlayerNextLevel(Index)
                            Call SetPlayerLevel(Index, GetPlayerLevel(Index) + 1)
                            I = Int(GetPlayerSPEED(Index) / 10)
                            If I < 1 Then
                                I = 1
                            End If
                            If I > 3 Then
                                I = 3
                            End If

                            Call SetPlayerPOINTS(Index, GetPlayerPOINTS(Index) + I)
                            Call SetPlayerExp(Index, d)
                            c = c + 1
                        End If
                    End If
                Loop
                If c > 1 Then
                    Call GlobalMsg(GetPlayerName(Index) & " ha ganado " & c & " niveles!", 6)
                Else
                    Call GlobalMsg(GetPlayerName(Index) & " ha ganado un nivel!", 6)
                End If
                Call BattleMsg(Index, "Tu tienes " & GetPlayerPOINTS(Index) & " puntos de estado", 9, 0)
            End If
            Call SendDataToMap(GetPlayerMap(Index), "levelup" & SEP_CHAR & Index & END_CHAR)
            Call SendPlayerLevelToAll(Index)
        End If

        If GetPlayerLevel(Index) = MAX_LEVEL Then
            Call SetPlayerExp(Index, Experience(MAX_LEVEL))
        End If
    End If

    Call SendHP(Index)
    Call SendMP(Index)
    Call SendSP(Index)
    Call SendPTS(Index)

    Player(Index).Char(Player(Index).CharNum).MAXHP = GetPlayerMaxHP(Index)
    Player(Index).Char(Player(Index).CharNum).MAXMP = GetPlayerMaxMP(Index)
    Player(Index).Char(Player(Index).CharNum).MAXSP = GetPlayerMaxSP(Index)

    Call SendStats(Index)
End Sub

Sub CastSpell(ByVal Index As Long, ByVal SpellSlot As Long)
    Dim spellnum As Long, I As Long, N As Long, Damage As Long
    Dim Casted As Boolean
    Casted = False
        

    ' Prevent player from using spells if they have been script locked
    If Player(Index).LockedSpells = True Then
        Exit Sub
    End If

    ' Prevent subscript out of range
    If SpellSlot <= 0 Or SpellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If

    spellnum = GetPlayerSpell(Index, SpellSlot)

    ' Make sure player has the spell
    If Not HasSpell(Index, spellnum) Then
        Call BattleMsg(Index, "Tu no tienes este hechizo!", BRIGHTRED, 0)
        Exit Sub
    End If

    I = GetSpellReqLevel(spellnum)

    ' Check if they have enough MP
    If GetPlayerMP(Index) < Spell(spellnum).MPCost Then
        Call BattleMsg(Index, "No tienes suficiente mana!", BRIGHTRED, 0)
        Exit Sub
    End If

    ' Make sure they are the right level
    If I > GetPlayerLevel(Index) Then
        Call BattleMsg(Index, "Necesitas ser mayor del nivel " & I & " para realizar este hechizo.", BRIGHTRED, 0)
        Exit Sub
    End If

    ' Check if timer is ok
    If GetTickCount < Player(Index).AttackTimer + Spell(spellnum).TimeToCast * 1000 Then
        Exit Sub
    End If

    ' Check if the spell is scripted and do that instead of a stat modification
    If Spell(spellnum).Type = SPELL_TYPE_SCRIPTED Then

        MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedSpell " & Index & "," & Spell(spellnum).Data1

        Exit Sub
    End If
' End If

    Dim X As Long, Y As Long

    If Spell(spellnum).AE = 1 Then
        For Y = GetPlayerY(Index) - Spell(spellnum).Range To GetPlayerY(Index) + Spell(spellnum).Range
            For X = GetPlayerX(Index) - Spell(spellnum).Range To GetPlayerX(Index) + Spell(spellnum).Range
                N = -1
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) = True Then
                        If GetPlayerMap(Index) = GetPlayerMap(I) Then
                            If GetPlayerX(I) = X And GetPlayerY(I) = Y Then
                                If I = Index Then
                                    If Spell(spellnum).Type = SPELL_TYPE_ADDHP Or Spell(spellnum).Type = SPELL_TYPE_ADDMP Or Spell(spellnum).Type = SPELL_TYPE_ADDSP Then
                                        Player(Index).Target = I
                                        Player(Index).TargetType = TARGET_TYPE_PLAYER
                                        N = Player(Index).Target
                                    End If
                                Else
                                    Player(Index).Target = I
                                    Player(Index).TargetType = TARGET_TYPE_PLAYER
                                    N = Player(Index).Target
                                End If
                            End If
                        End If
                    End If
                Next I

                For I = 1 To MAX_MAP_NPCS
                    If MapNPC(GetPlayerMap(Index), I).num > 0 Then
                        If NPC(MapNPC(GetPlayerMap(Index), I).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(MapNPC(GetPlayerMap(Index), I).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            If MapNPC(GetPlayerMap(Index), I).X = X And MapNPC(GetPlayerMap(Index), I).Y = Y Then
                                Player(Index).Target = I
                                Player(Index).TargetType = TARGET_TYPE_NPC
                                N = Player(Index).Target
                            End If
                        End If
                    End If
                Next I

                Casted = False
                If N > 0 Then
                    If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
                        If IsPlaying(N) Then
                            If N = Index Then
                                Select Case Spell(spellnum).Type

                                    Case SPELL_TYPE_ADDHP
                                        ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        Call SetPlayerHP(N, GetPlayerHP(N) + Spell(spellnum).Data1)
                                        Call SendHP(N)

                                    Case SPELL_TYPE_ADDMP
                                        ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        Call SetPlayerMP(N, GetPlayerMP(N) + Spell(spellnum).Data1)
                                        Call SendMP(N)

                                    Case SPELL_TYPE_ADDSP
                                        ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                        Call SetPlayerMP(N, GetPlayerSP(N) + Spell(spellnum).Data1)
                                        Call SendMP(N)
                                End Select

                                Casted = True
                            Else
                                Call PlayerMsg(Index, "No puedes realizar el hechizo!", BRIGHTRED)
                            End If
                            If N <> Index Then
                                Player(Index).TargetType = TARGET_TYPE_PLAYER
                                If GetPlayerHP(N) > 0 And GetPlayerMap(Index) = GetPlayerMap(N) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(N) >= 10 And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(N) <= 0 Then
' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)

                                    Select Case Spell(spellnum).Type
                                        Case SPELL_TYPE_SUBHP

                                            Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(spellnum).Data1) - GetPlayerProtection(N)
                                            If Damage > 0 Then
                                                Call AttackPlayer(Index, N, Damage)
                                            Else
                                                Call BattleMsg(Index, "El hechizo era muy debil para dañar a " & GetPlayerName(N) & "!", BRIGHTRED, 0)
                                            End If

                                        Case SPELL_TYPE_SUBMP
                                            Call SetPlayerMP(N, GetPlayerMP(N) - Spell(spellnum).Data1)
                                            Call SendMP(N)

                                        Case SPELL_TYPE_SUBSP
                                            Call SetPlayerSP(N, GetPlayerSP(N) - Spell(spellnum).Data1)
                                            Call SendSP(N)
                                    End Select

                                    Casted = True
                                Else
                                    If GetPlayerMap(Index) = GetPlayerMap(N) And Spell(spellnum).Type >= SPELL_TYPE_ADDHP And Spell(spellnum).Type <= SPELL_TYPE_ADDSP Then
                                        Select Case Spell(spellnum).Type

                                            Case SPELL_TYPE_ADDHP
                                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                Call SetPlayerHP(N, GetPlayerHP(N) + Spell(spellnum).Data1)
                                                Call SendHP(N)

                                            Case SPELL_TYPE_ADDMP
                                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                Call SetPlayerMP(N, GetPlayerMP(N) + Spell(spellnum).Data1)
                                                Call SendMP(N)

                                            Case SPELL_TYPE_ADDSP
                                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                Call SetPlayerMP(N, GetPlayerSP(N) + Spell(spellnum).Data1)
                                                Call SendMP(N)
                                        End Select

                                        Casted = True
                                    Else
                                        Call PlayerMsg(Index, "No puedes realizar el hechizo!", BRIGHTRED)
                                    End If
                                End If
                            Else
                                Player(Index).TargetType = TARGET_TYPE_PLAYER
                                If N = Index Then
                                    Select Case Spell(spellnum).Type

                                        Case SPELL_TYPE_ADDHP
                                            ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                            Call SetPlayerHP(N, GetPlayerHP(N) + Spell(spellnum).Data1)
                                            Call SendHP(N)

                                        Case SPELL_TYPE_ADDMP
                                            ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                            Call SetPlayerMP(N, GetPlayerMP(N) + Spell(spellnum).Data1)
                                            Call SendMP(N)

                                        Case SPELL_TYPE_ADDSP
                                            ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                            Call SetPlayerMP(N, GetPlayerSP(N) + Spell(spellnum).Data1)
                                            Call SendMP(N)
                                    End Select

                                    Casted = True
                                Else
                                    Call PlayerMsg(Index, "No puedes realizar el hechizo!", BRIGHTRED)
                                End If
                                If GetPlayerHP(N) > 0 And GetPlayerMap(Index) = GetPlayerMap(N) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(N) >= 10 And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(N) <= 0 Then
                                Else
                                    If GetPlayerMap(Index) = GetPlayerMap(N) And Spell(spellnum).Type >= SPELL_TYPE_ADDHP And Spell(spellnum).Type <= SPELL_TYPE_ADDSP Then
                                        Select Case Spell(spellnum).Type

                                            Case SPELL_TYPE_ADDHP
                                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                Call SetPlayerHP(N, GetPlayerHP(N) + Spell(spellnum).Data1)
                                                Call SendHP(N)

                                            Case SPELL_TYPE_ADDMP
                                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                Call SetPlayerMP(N, GetPlayerMP(N) + Spell(spellnum).Data1)
                                                Call SendMP(N)

                                            Case SPELL_TYPE_ADDSP
                                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                                Call SetPlayerMP(N, GetPlayerSP(N) + Spell(spellnum).Data1)
                                                Call SendMP(N)
                                        End Select

                                        Casted = True
                                    Else
                                        Call BattleMsg(Index, "No puedes realizar el hechizo!", BRIGHTRED, 0)
                                    End If
                                End If
                            End If
                        Else
                            Call BattleMsg(Index, "No puedes realizar el hechizo!", BRIGHTRED, 0)
                        End If
                    Else
                        Player(Index).TargetType = TARGET_TYPE_NPC
                        If NPC(MapNPC(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(MapNPC(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(MapNPC(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_QUEST Then
                            If Spell(spellnum).Type >= SPELL_TYPE_SUBHP And Spell(spellnum).Type <= SPELL_TYPE_SUBSP Then
                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on a " & Trim$(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & ".", BrightBlue)
                                Select Case Spell(spellnum).Type

                                    Case SPELL_TYPE_SUBHP
                                        Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(spellnum).Data1) - Int(NPC(MapNPC(GetPlayerMap(Index), N).num).DEF / 2)

                                        If Damage > 0 Then
                                            If Spell(spellnum).Element <> 0 And NPC(MapNPC(GetPlayerMap(Index), N).num).Element <> 0 Then
                                                If Element(Spell(spellnum).Element).Strong = NPC(MapNPC(GetPlayerMap(Index), N).num).Element Or Element(NPC(MapNPC(GetPlayerMap(Index), N).num).Element).Weak = Spell(spellnum).Element Then
                                                    Call BattleMsg(Index, "Una mezcla mortal de los elementos daña a " & Trim$(NPC(MapNPC(GetPlayerMap(Index), N).num).Name) & "!", Blue, 0)
                                                    Damage = Int(Damage * 1.25)
                                                    If Element(Spell(spellnum).Element).Strong = NPC(MapNPC(GetPlayerMap(Index), N).num).Element And Element(NPC(MapNPC(GetPlayerMap(Index), N).num).Element).Weak = Spell(spellnum).Element Then
                                                        Damage = Int(Damage * 1.2)
                                                    End If
                                                End If

                                                If Element(Spell(spellnum).Element).Weak = NPC(MapNPC(GetPlayerMap(Index), N).num).Element Or Element(NPC(MapNPC(GetPlayerMap(Index), N).num).Element).Strong = Spell(spellnum).Element Then
                                                    Call BattleMsg(Index, "" & Trim$(NPC(MapNPC(GetPlayerMap(Index), N).num).Name) & " absorbe la fuerza elemental y se cura!", Red, 0)
                                                    Damage = Int(Damage * 0.75)
                                                    If Element(Spell(spellnum).Element).Weak = NPC(MapNPC(GetPlayerMap(Index), N).num).Element And Element(NPC(MapNPC(GetPlayerMap(Index), N).num).Element).Strong = Spell(spellnum).Element Then
                                                        Damage = Int(Damage * (2 / 3))
                                                    End If
                                                End If
                                            End If
                                            Call AttackNpc(Index, N, Damage)
                                        Else
                                            Call BattleMsg(Index, "El hechizo es muy debil para dañar a " & Trim$(NPC(MapNPC(GetPlayerMap(Index), N).num).Name) & "!", BRIGHTRED, 0)
                                        End If

                                    Case SPELL_TYPE_SUBMP
                                        MapNPC(GetPlayerMap(Index), N).MP = MapNPC(GetPlayerMap(Index), N).MP - Spell(spellnum).Data1

                                    Case SPELL_TYPE_SUBSP
                                        MapNPC(GetPlayerMap(Index), N).SP = MapNPC(GetPlayerMap(Index), N).SP - Spell(spellnum).Data1
                                End Select

                                Casted = True
                            Else
                                Select Case Spell(spellnum).Type
                                    Case SPELL_TYPE_ADDHP
' MapNpc(GetPlayerMap(Index), n).HP = MapNpc(GetPlayerMap(Index), n).HP + Spell(SpellNum).Data1

                                    Case SPELL_TYPE_ADDMP
' MapNpc(GetPlayerMap(Index), n).MP = MapNpc(GetPlayerMap(Index), n).MP + Spell(SpellNum).Data1

                                    Case SPELL_TYPE_ADDSP
                                ' MapNpc(GetPlayerMap(Index), n).SP = MapNpc(GetPlayerMap(Index), n).SP + Spell(SpellNum).Data1
                                End Select
                                Casted = False
                            End If
                        Else
                            Call BattleMsg(Index, "No puedes realizar el hechizo!", BRIGHTRED, 0)
                        End If
                    End If
                End If
                If Casted = True Then
                    Call SendDataToMap(GetPlayerMap(Index), "spellanim" & SEP_CHAR & spellnum & SEP_CHAR & Spell(spellnum).SpellAnim & SEP_CHAR & Spell(spellnum).SpellTime & SEP_CHAR & Spell(spellnum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & Player(Index).Target & SEP_CHAR & Player(Index).CastedSpell & SEP_CHAR & Spell(spellnum).Big & END_CHAR)
                    Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "magic" & SEP_CHAR & Spell(spellnum).Sound & END_CHAR)
                End If
            Next X
        Next Y

        Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(spellnum).MPCost)
        Call SendMP(Index)
    Else
        N = Player(Index).Target
        If Player(Index).TargetType = TARGET_TYPE_PLAYER Then
            If IsPlaying(N) Then
                If GetPlayerName(N) <> GetPlayerName(Index) Then
                    If CInt(Sqr((GetPlayerX(Index) - GetPlayerX(N)) ^ 2 + ((GetPlayerY(Index) - GetPlayerY(N)) ^ 2))) > Spell(spellnum).Range Then
                        Call BattleMsg(Index, "Estas muy alejado para dañar al objetivo.", BRIGHTRED, 0)
                        Exit Sub
                    End If
                End If
                Player(Index).TargetType = TARGET_TYPE_PLAYER
                If GetPlayerHP(N) > 0 And GetPlayerMap(Index) = GetPlayerMap(N) And GetPlayerLevel(Index) >= 10 And GetPlayerLevel(N) >= 10 And (Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(Index) <= 0 And GetPlayerAccess(N) <= 0 Then
' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)

                    Select Case Spell(spellnum).Type
                        Case SPELL_TYPE_SUBHP

                            Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(spellnum).Data1) - GetPlayerProtection(N)
                            If Damage > 0 Then
                                Call AttackPlayer(Index, N, Damage)
                            Else
                                Call BattleMsg(Index, "El hechizo es muy debil para dañar a " & GetPlayerName(N) & "!", BRIGHTRED, 0)
                            End If

                        Case SPELL_TYPE_SUBMP
                            Call SetPlayerMP(N, GetPlayerMP(N) - Spell(spellnum).Data1)
                            Call SendMP(N)

                        Case SPELL_TYPE_SUBSP
                            Call SetPlayerSP(N, GetPlayerSP(N) - Spell(spellnum).Data1)
                            Call SendSP(N)
                    End Select

                    ' Take away the mana points
                    Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(spellnum).MPCost)
                    Call SendMP(Index)
                    Casted = True
                Else
                    If GetPlayerMap(Index) = GetPlayerMap(N) And Spell(spellnum).Type >= SPELL_TYPE_ADDHP And Spell(spellnum).Type <= SPELL_TYPE_ADDSP Then
                        Select Case Spell(spellnum).Type

                            Case SPELL_TYPE_ADDHP
                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                Call SetPlayerHP(N, GetPlayerHP(N) + Spell(spellnum).Data1)
                                Call SendHP(N)

                            Case SPELL_TYPE_ADDMP
                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                Call SetPlayerMP(N, GetPlayerMP(N) + Spell(spellnum).Data1)
                                Call SendMP(N)

                            Case SPELL_TYPE_ADDSP
                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)
                                Call SetPlayerMP(N, GetPlayerSP(N) + Spell(spellnum).Data1)
                                Call SendMP(N)
                        End Select

                        ' Take away the mana points
                        Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(spellnum).MPCost)
                        Call SendMP(Index)
                        Casted = True
                    Else
                        Call BattleMsg(Index, "No puedes realizar el hechizo!", BRIGHTRED, 0)
                    End If
                End If
            Else
                Call PlayerMsg(Index, "No puedes realizar el hechizo!", BRIGHTRED)
            End If
        Else
                N = Player(Index).TargetNPC
            If CInt(Sqr((GetPlayerX(Index) - MapNPC(GetPlayerMap(Index), N).X) ^ 2 + ((GetPlayerY(Index) - MapNPC(GetPlayerMap(Index), N).Y) ^ 2))) > Spell(spellnum).Range Then
                Call BattleMsg(Index, "Estás muy alejado para dañar a tu objetivo.", BRIGHTRED, 0)
                Exit Sub
            End If

            Player(Index).TargetType = TARGET_TYPE_NPC

            If NPC(MapNPC(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(MapNPC(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(MapNPC(GetPlayerMap(Index), N).num).Behavior <> NPC_BEHAVIOR_QUEST Then
' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on a " & Trim$(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & ".", BrightBlue)

                Select Case Spell(spellnum).Type
                    Case SPELL_TYPE_ADDHP
                        MapNPC(GetPlayerMap(Index), N).HP = MapNPC(GetPlayerMap(Index), N).HP + Spell(spellnum).Data1

                    Case SPELL_TYPE_SUBHP

                        Damage = (Int(GetPlayerMAGI(Index) / 4) + Spell(spellnum).Data1) - Int(NPC(MapNPC(GetPlayerMap(Index), N).num).DEF / 2)
                        If Damage > 0 Then
                            If Spell(spellnum).Element <> 0 And NPC(MapNPC(GetPlayerMap(Index), N).num).Element <> 0 Then
                                If Element(Spell(spellnum).Element).Strong = NPC(MapNPC(GetPlayerMap(Index), N).num).Element Or Element(NPC(MapNPC(GetPlayerMap(Index), N).num).Element).Weak = Spell(spellnum).Element Then
                                    Call BattleMsg(Index, "Una mezcla mortal de los elementos daña a " & Trim$(NPC(MapNPC(GetPlayerMap(Index), N).num).Name) & "!", Blue, 0)
                                    Damage = Int(Damage * 1.25)
                                    If Element(Spell(spellnum).Element).Strong = NPC(MapNPC(GetPlayerMap(Index), N).num).Element And Element(NPC(MapNPC(GetPlayerMap(Index), N).num).Element).Weak = Spell(spellnum).Element Then
                                        Damage = Int(Damage * 1.2)
                                    End If
                                End If

                                If Element(Spell(spellnum).Element).Weak = NPC(MapNPC(GetPlayerMap(Index), N).num).Element Or Element(NPC(MapNPC(GetPlayerMap(Index), N).num).Element).Strong = Spell(spellnum).Element Then
                                    Call BattleMsg(Index, "" & Trim$(NPC(MapNPC(GetPlayerMap(Index), N).num).Name) & " absorbe gran cantidad elemental y se cura!", Red, 0)
                                    Damage = Int(Damage * 0.75)
                                    If Element(Spell(spellnum).Element).Weak = NPC(MapNPC(GetPlayerMap(Index), N).num).Element And Element(NPC(MapNPC(GetPlayerMap(Index), N).num).Element).Strong = Spell(spellnum).Element Then
                                        Damage = Int(Damage * (2 / 3))
                                    End If
                                End If
                            End If
                            Call AttackNpc(Index, N, Damage)
                        Else
                            Call BattleMsg(Index, "El hechizo es muy debil para dañar a " & Trim$(NPC(MapNPC(GetPlayerMap(Index), N).num).Name) & "!", BRIGHTRED, 0)
                        End If

                    Case SPELL_TYPE_ADDMP
                        MapNPC(GetPlayerMap(Index), N).MP = MapNPC(GetPlayerMap(Index), N).MP + Spell(spellnum).Data1

                    Case SPELL_TYPE_SUBMP
                        MapNPC(GetPlayerMap(Index), N).MP = MapNPC(GetPlayerMap(Index), N).MP - Spell(spellnum).Data1

                    Case SPELL_TYPE_ADDSP
                        MapNPC(GetPlayerMap(Index), N).SP = MapNPC(GetPlayerMap(Index), N).SP + Spell(spellnum).Data1

                    Case SPELL_TYPE_SUBSP
                        MapNPC(GetPlayerMap(Index), N).SP = MapNPC(GetPlayerMap(Index), N).SP - Spell(spellnum).Data1
                End Select

                ' Take away the mana points
                Call SetPlayerMP(Index, GetPlayerMP(Index) - Spell(spellnum).MPCost)
                Call SendMP(Index)
                Casted = True
            Else
                Call BattleMsg(Index, "No puedes realizar el hechizo!", BRIGHTRED, 0)
            End If
        End If
    End If
' Solución para animación de hechizos
' Fixed #1 por Stream / Cambiado timer de hechizos
If Casted = True Then
        Player(Index).AttackTimer = GetTickCount
        Player(Index).CastedSpell = YES
        Call SendDataToMap(GetPlayerMap(Index), "spellanim" & SEP_CHAR & spellnum & SEP_CHAR & Spell(spellnum).SpellAnim & SEP_CHAR & Spell(spellnum).SpellTime & SEP_CHAR & Spell(spellnum).SpellDone & SEP_CHAR & Index & SEP_CHAR & Player(Index).TargetType & SEP_CHAR & N & SEP_CHAR & Player(Index).CastedSpell & SEP_CHAR & Spell(spellnum).Big & END_CHAR)
        Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "magic" & SEP_CHAR & Spell(spellnum).Sound & END_CHAR)
    End If
End Sub

Function CanPlayerCriticalHit(ByVal Index As Long) As Boolean
    Dim I As Long
    Dim N As Long

    If GetPlayerWeaponSlot(Index) > 0 Then
        N = Int(Rnd * 2)

        If N = 1 Then
            I = Int(GetPlayerSTR(Index) / 2) + Int(GetPlayerLevel(Index) / 2)

            N = Int(Rnd * 100) + 1
            If N <= I Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If
End Function

Function CanPlayerBlockHit(ByVal Index As Long) As Boolean
    Dim I As Long
    Dim N As Long

    If GetPlayerShieldSlot(Index) > 0 Then
        N = Int(Rnd * 2)

        If N = 1 Then
            I = Int(GetPlayerDEF(Index) / 2) + Int(GetPlayerLevel(Index) / 2)

            N = Int(Rnd * 100) + 1
            If N <= I Then
                CanPlayerBlockHit = True
            End If
        End If
    End If
End Function

Public Sub CheckEquippedItems(ByVal Index As Long)
    Dim ItemNum As Long

    ' Check to make sure the weapon exists.
    ItemNum = GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then
            If Item(ItemNum).Type <> ITEM_TYPE_TWO_HAND Then
                Call SetPlayerWeaponSlot(Index, 0)
            End If
        End If
    Else
        Call SetPlayerWeaponSlot(Index, 0)
    End If

    ' Check to make sure the chest armor exists.
    ItemNum = GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_ARMOR Then
            Call SetPlayerArmorSlot(Index, 0)
        End If
    Else
        Call SetPlayerArmorSlot(Index, 0)
    End If

    ' Check to make sure the helmet exists.
    ItemNum = GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_HELMET Then
            Call SetPlayerHelmetSlot(Index, 0)
        End If
    Else
        Call SetPlayerHelmetSlot(Index, 0)
    End If

    ' Check to make sure the shield exists.
    ItemNum = GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_SHIELD Then
            Call SetPlayerShieldSlot(Index, 0)
        End If
    Else
        Call SetPlayerShieldSlot(Index, 0)
    End If

    ' Check to make sure the leggings exists.
    ItemNum = GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_LEGS Then
            Call SetPlayerLegsSlot(Index, 0)
        End If
    Else
        Call SetPlayerLegsSlot(Index, 0)
    End If

    ' Check to make sure the ring exists.
    ItemNum = GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_RING Then
            Call SetPlayerRingSlot(Index, 0)
        End If
    Else
        Call SetPlayerRingSlot(Index, 0)
    End If

    ' Check to make sure the necklace exists.
    ItemNum = GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_NECKLACE Then
            Call SetPlayerNecklaceSlot(Index, 0)
        End If
    Else
        Call SetPlayerNecklaceSlot(Index, 0)
    End If
End Sub



' This sub-routine needs a re-write. [Mellowz]

Public Sub ShowPLR(ByVal Index As Long)
    Dim LS As ListItem

    On Error Resume Next

    If frmServer.lvUsers.ListItems.Count > 0 And IsPlaying(Index) Then
        frmServer.lvUsers.ListItems.Remove Index
    End If

    Set LS = frmServer.lvUsers.ListItems.Add(Index, , Index)

    If IsPlaying(Index) Then
        LS.SubItems(1) = GetPlayerLogin(Index)
        LS.SubItems(2) = GetPlayerName(Index)
        LS.SubItems(3) = GetPlayerLevel(Index)
        LS.SubItems(4) = GetPlayerSprite(Index)
        LS.SubItems(5) = GetPlayerAccess(Index)
    End If
End Sub

Public Sub RemovePLR(ByVal Index As Long)
    Dim LS As ListItem
    
    On Error Resume Next

    If Not IsPlaying(Index) Then
        frmServer.lvUsers.ListItems.Remove Index
    
        Set LS = frmServer.lvUsers.ListItems.Add(Index, , Index)
        
        LS.SubItems(1) = vbNullString
        LS.SubItems(2) = vbNullString
        LS.SubItems(3) = vbNullString
        LS.SubItems(4) = vbNullString
        LS.SubItems(5) = vbNullString
    End If
End Sub

Function CanAttackPlayerWithArrow(ByVal Attacker As Long, ByVal Victim As Long) As Boolean
    ' Check If map Is attackable
    If Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(Attacker)).Moral = MAP_MORAL_NO_PENALTY Or GetPlayerPK(Victim) = YES Then
        ' Make sure they are high enough level
        If GetPlayerLevel(Attacker) < 10 Then
            Call PlayerMsg(Attacker, "Tu nivel es menor de 10, no podras atacar a nadie hasta que superes este nivel.", BRIGHTRED)
        Else
            If GetPlayerLevel(Victim) < 10 Then
                Call PlayerMsg(Attacker, GetPlayerName(Victim) & " es menor del nivel 10 por lo que no puedes atacarle.", BRIGHTRED)
            Else
                If Trim$(GetPlayerGuild(Attacker)) <> vbNullString And GetPlayerGuild(Victim) <> vbNullString Then
                    If Trim$(GetPlayerGuild(Attacker)) <> Trim$(GetPlayerGuild(Victim)) Then
                        CanAttackPlayerWithArrow = True
                    Else
                        Call PlayerMsg(Attacker, "Esta en el mismo clan que tu por lo que no puedes atacarlo.", BRIGHTRED)
                    End If
                Else
                    CanAttackPlayerWithArrow = True
                End If
            End If
        End If
    Else
        Call PlayerMsg(Attacker, "Esto es una zona segura!", BRIGHTRED)
    End If
End Function

Function CanAttackNpcWithArrow(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
    Dim mapnum As Long, npcnum As Long
    Dim AttackSpeed As Long

    If GetPlayerWeaponSlot(Attacker) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(Attacker, GetPlayerWeaponSlot(Attacker))).AttackSpeed
    Else
        AttackSpeed = 1000
    End If

    CanAttackNpcWithArrow = False

    ' Check For subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Check For subscript out of range
    If MapNPC(GetPlayerMap(Attacker), MapNpcNum).num <= 0 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(Attacker)
    npcnum = MapNPC(mapnum, MapNpcNum).num

    ' Make sure the npc isn't already dead
    If MapNPC(mapnum, MapNpcNum).HP <= 0 Then
        Exit Function
    End If

    ' Make sure they are On the same map
    If IsPlaying(Attacker) Then
        If npcnum > 0 And GetTickCount > Player(Attacker).AttackTimer + AttackSpeed Then
            ' Check If at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    If NPC(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(npcnum).Behavior <> NPC_BEHAVIOR_SCRIPTED And NPC(npcnum).Behavior <> NPC_BEHAVIOR_QUEST Then
                        CanAttackNpcWithArrow = True
                    Else
                        If NPC(npcnum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & NPC(npcnum).SpawnSecs
                        ElseIf NPC(npcnum).Behavior = NPC_BEHAVIOR_QUEST Then
                              Call DoQuest(NPC(npcnum).Quest, Attacker, npcnum)
                        ElseIf NPC(npcnum).Behavior = NPC_BEHAVIOR_CHAOSKNIGHT Then
                        Call DoQuestNpcKillsQuest(Attacker, npcnum)
                       Else
                            Call PlayerMsg(Attacker, Trim$(NPC(npcnum).Name) & " :" & Trim$(NPC(npcnum).AttackSay), Green)
                        End If
                    End If

                Case DIR_DOWN
                    If NPC(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(npcnum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
                        CanAttackNpcWithArrow = True
                    Else
                        If NPC(npcnum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & NPC(npcnum).SpawnSecs
                        Else
                            Call PlayerMsg(Attacker, Trim$(NPC(npcnum).Name) & " :" & Trim$(NPC(npcnum).AttackSay), Green)
                        End If
                    End If

                Case DIR_LEFT
                    If NPC(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(npcnum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
                        CanAttackNpcWithArrow = True
                    Else
                        If NPC(npcnum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & NPC(npcnum).SpawnSecs
                        Else
                            Call PlayerMsg(Attacker, Trim$(NPC(npcnum).Name) & " :" & Trim$(NPC(npcnum).AttackSay), Green)
                        End If
                    End If

                Case DIR_RIGHT
                    If NPC(npcnum).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(npcnum).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(npcnum).Behavior <> NPC_BEHAVIOR_SCRIPTED Then
                        CanAttackNpcWithArrow = True
                    Else
                        If NPC(npcnum).Behavior = NPC_BEHAVIOR_SCRIPTED Then
                            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedNPC " & Attacker & "," & NPC(npcnum).SpawnSecs
                        Else
                            Call PlayerMsg(Attacker, Trim$(NPC(npcnum).Name) & " :" & Trim$(NPC(npcnum).AttackSay), Green)
                        End If
                    End If
            End Select
        End If
    End If
End Function

Sub SendIndexWornEquipment(ByVal Index As Long)
    Dim Armor As Long
    Dim Helmet As Long
    Dim Shield As Long
    Dim Weapon As Long
    Dim Legs As Long
    Dim Ring As Long
    Dim Necklace As Long

    If GetPlayerArmorSlot(Index) > 0 Then
        Armor = GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))
    End If

    If GetPlayerHelmetSlot(Index) > 0 Then
        Helmet = GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))
    End If

    If GetPlayerShieldSlot(Index) > 0 Then
        Shield = GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))
    End If

    If GetPlayerWeaponSlot(Index) > 0 Then
        Weapon = GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))
    End If

    If GetPlayerLegsSlot(Index) > 0 Then
        Legs = GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))
    End If

    If GetPlayerRingSlot(Index) > 0 Then
        Ring = GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))
    End If

    If GetPlayerNecklaceSlot(Index) > 0 Then
        Necklace = GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))
    End If

    Call SendDataToMap(GetPlayerMap(Index), "itemworn" & SEP_CHAR & Index & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & END_CHAR)
End Sub

Sub SendIndexWornEquipmentto(ByVal Index As Long, ByVal From As Long)
    Dim Armor As Long
    Dim Helmet As Long
    Dim Shield As Long
    Dim Weapon As Long
    Dim Legs As Long
    Dim Ring As Long
    Dim Necklace As Long

    If GetPlayerArmorSlot(From) > 0 Then
        Armor = GetPlayerInvItemNum(From, GetPlayerArmorSlot(From))
    End If

    If GetPlayerHelmetSlot(From) > 0 Then
        Helmet = GetPlayerInvItemNum(From, GetPlayerHelmetSlot(From))
    End If

    If GetPlayerShieldSlot(From) > 0 Then
        Shield = GetPlayerInvItemNum(From, GetPlayerShieldSlot(Index))
    End If

    If GetPlayerWeaponSlot(From) > 0 Then
        Weapon = GetPlayerInvItemNum(From, GetPlayerWeaponSlot(From))
    End If

    If GetPlayerLegsSlot(From) > 0 Then
        Legs = GetPlayerInvItemNum(From, GetPlayerLegsSlot(From))
    End If

    If GetPlayerRingSlot(From) > 0 Then
        Ring = GetPlayerInvItemNum(From, GetPlayerRingSlot(From))
    End If

    If GetPlayerNecklaceSlot(From) > 0 Then
        Necklace = GetPlayerInvItemNum(From, GetPlayerNecklaceSlot(From))
    End If

    Call SendDataTo(Index, "itemworn" & SEP_CHAR & From & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & END_CHAR)
End Sub


Sub SendIndexWornEquipmentFromMap(ByVal Index As Long)
    Dim Packet As String
    Dim I As Long
    Dim Armor As Long
    Dim Helmet As Long
    Dim Shield As Long
    Dim Weapon As Long
    Dim Legs As Long
    Dim Ring As Long
    Dim Necklace As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) = True Then
            If GetPlayerMap(I) = GetPlayerMap(Index) Then

                Armor = 0
                Helmet = 0
                Shield = 0
                Weapon = 0
                Legs = 0
                Ring = 0
                Necklace = 0

                If GetPlayerArmorSlot(I) > 0 Then
                    Armor = GetPlayerInvItemNum(I, GetPlayerArmorSlot(I))
                End If
                If GetPlayerHelmetSlot(I) > 0 Then
                    Helmet = GetPlayerInvItemNum(I, GetPlayerHelmetSlot(I))
                End If
                If GetPlayerShieldSlot(I) > 0 Then
                    Shield = GetPlayerInvItemNum(I, GetPlayerShieldSlot(I))
                End If
                If GetPlayerWeaponSlot(I) > 0 Then
                    Weapon = GetPlayerInvItemNum(I, GetPlayerWeaponSlot(I))
                End If
                If GetPlayerLegsSlot(I) > 0 Then
                    Legs = GetPlayerInvItemNum(I, GetPlayerLegsSlot(I))
                End If
                If GetPlayerRingSlot(I) > 0 Then
                    Ring = GetPlayerInvItemNum(I, GetPlayerRingSlot(I))
                End If
                If GetPlayerNecklaceSlot(I) > 0 Then
                    Necklace = GetPlayerInvItemNum(I, GetPlayerNecklaceSlot(I))
                End If

                Packet = "itemworn" & SEP_CHAR & I & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & END_CHAR
                Call SendDataTo(Index, Packet)
            End If
        End If
    Next I
End Sub

Sub AddNewTimer(ByVal Name As String, ByVal Interval As Long)
    On Error Resume Next
    Dim TmpTimer As clsCTimers
    Set TmpTimer = New clsCTimers
    TmpTimer.Name = Name
    TmpTimer.Interval = Interval
    TmpTimer.tmrWait = GetTickCount + Interval
    CTimers.Add TmpTimer, Name
    If Err.Number > 0 Then
        Debug.Print "Err: " & Err.Number
        CTimers.Item(Name).Name = Name
        CTimers.Item(Name).Interval = Interval
        CTimers.Item(Name).tmrWait = GetTickCount + Interval
        Err.Clear
    End If
End Sub

Function GetTimeLeft(ByVal Name As String) As Long
    On Error GoTo Hell
    GetTimeLeft = CTimers.Item(Name).tmrWait - GetTickCount
    Exit Function
Hell:
    GetTimeLeft = -1
End Function

Sub GetRidOfTimer(ByVal Name As String)
    Call CTimers.Remove(Name)
End Sub
Sub ScriptSetTile(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal setx As Long, ByVal sety As Long, ByVal tileset As Long, ByVal layer As Long)
    Dim Packet As String
    Packet = "tilecheck" & SEP_CHAR & mapper & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & layer & SEP_CHAR

    Select Case layer

        Case 0
            Map(mapper).Tile(X, Y).Ground = sety * 14 + setx
            Map(mapper).Tile(X, Y).GroundSet = tileset
            Packet = Packet & Map(mapper).Tile(X, Y).Ground & SEP_CHAR & Map(mapper).Tile(X, Y).GroundSet

        Case 1
            Map(mapper).Tile(X, Y).Mask = sety * 14 + setx
            Map(mapper).Tile(X, Y).MaskSet = tileset
            Packet = Packet & Map(mapper).Tile(X, Y).Mask & SEP_CHAR & Map(mapper).Tile(X, Y).MaskSet

        Case 2
            Map(mapper).Tile(X, Y).Anim = sety * 14 + setx
            Map(mapper).Tile(X, Y).AnimSet = tileset
            Packet = Packet & Map(mapper).Tile(X, Y).Anim & SEP_CHAR & Map(mapper).Tile(X, Y).AnimSet

        Case 3
            Map(mapper).Tile(X, Y).Mask2 = sety * 14 + setx
            Map(mapper).Tile(X, Y).Mask2Set = tileset
            Packet = Packet & Map(mapper).Tile(X, Y).Mask2 & SEP_CHAR & Map(mapper).Tile(X, Y).Mask2Set

        Case 4
            Map(mapper).Tile(X, Y).M2Anim = sety * 14 + setx
            Map(mapper).Tile(X, Y).M2AnimSet = tileset
            Packet = Packet & Map(mapper).Tile(X, Y).M2Anim & SEP_CHAR & Map(mapper).Tile(X, Y).M2AnimSet

        Case 5
            Map(mapper).Tile(X, Y).Fringe = sety * 14 + setx
            Map(mapper).Tile(X, Y).FringeSet = tileset
            Packet = Packet & Map(mapper).Tile(X, Y).Fringe & SEP_CHAR & Map(mapper).Tile(X, Y).FringeSet

        Case 6
            Map(mapper).Tile(X, Y).FAnim = sety * 14 + setx
            Map(mapper).Tile(X, Y).FAnimSet = tileset
            Packet = Packet & Map(mapper).Tile(X, Y).FAnim & SEP_CHAR & Map(mapper).Tile(X, Y).FAnimSet

        Case 7
            Map(mapper).Tile(X, Y).Fringe2 = sety * 14 + setx
            Map(mapper).Tile(X, Y).Fringe2Set = tileset
            Packet = Packet & Map(mapper).Tile(X, Y).Fringe2 & SEP_CHAR & Map(mapper).Tile(X, Y).Fringe2Set

        Case 8
            Map(mapper).Tile(X, Y).F2Anim = sety * 14 + setx
            Map(mapper).Tile(X, Y).F2AnimSet = tileset
            Packet = Packet & Map(mapper).Tile(X, Y).F2Anim & SEP_CHAR & Map(mapper).Tile(X, Y).F2AnimSet
    End Select

    Call SaveMap(mapper)
    Call SendDataToAll(Packet & END_CHAR)
End Sub

Sub ScriptSetAttribute(ByVal mapper As Long, ByVal X As Long, ByVal Y As Long, ByVal Attrib As Long, ByVal Data1 As Long, ByVal Data2 As Long, ByVal Data3 As Long, ByVal String1 As String, ByVal String2 As String, ByVal String3 As String)
    Dim Packet As String
    
    With Map(mapper).Tile(X, Y)
        .Type = Attrib
        .Data1 = Data1
        .Data2 = Data2
        .Data3 = Data3
        .String1 = String1
        .String2 = String2
        .String3 = String3
    End With

    Packet = "tilecheckattribute" & SEP_CHAR & mapper & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR
    With Map(mapper).Tile(X, Y)
        Packet = Packet & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR
    End With
    
    Call SaveMap(mapper)
    Call SendDataToAll(Packet & END_CHAR)
End Sub

Function ItemIsUsable(ByVal Index As Long, ByVal InvNum As Long) As Boolean
    ' Check if the player meets the class requirement.
    If Item(GetPlayerInvItemNum(Index, InvNum)).ClassReq > -1 Then
        If GetPlayerClass(Index) <> Item(GetPlayerInvItemNum(Index, InvNum)).ClassReq Then
            Call PlayerMsg(Index, "Necesitas ser la clase " & GetClassName(Item(GetPlayerInvItemNum(Index, InvNum)).ClassReq) & " para usar este objeto!", BRIGHTRED)
            Exit Function
        End If
    End If

    ' Check if the player meets the access requirement.
    If GetPlayerAccess(Index) < Item(GetPlayerInvItemNum(Index, InvNum)).AccessReq Then
        Call PlayerMsg(Index, "Tu privilegio necesita ser mayor de " & Item(GetPlayerInvItemNum(Index, InvNum)).AccessReq & "!", BRIGHTRED)
        Exit Function
    End If

    ' Check if the player meets the strength requirement.
    If GetPlayerSTR(Index) < Item(GetPlayerInvItemNum(Index, InvNum)).StrReq Then
        Call PlayerMsg(Index, "No tienes la suficiente fuerza para equiparte esto!", BRIGHTRED)
        Exit Function
    End If

    ' Check if the player meets the defense requirement.
    If GetPlayerDEF(Index) < Item(GetPlayerInvItemNum(Index, InvNum)).DefReq Then
        Call PlayerMsg(Index, "No tienes la suficiente defensa para equiparte esto!", BRIGHTRED)
        Exit Function
    End If

    ' Check if the player meets the magic requirement.
    If GetPlayerMAGI(Index) < Item(GetPlayerInvItemNum(Index, InvNum)).MagicReq Then
        Call PlayerMsg(Index, "No tienes la suficiente magia para equiparte esto!", BRIGHTRED)
        Exit Function
    End If

    ' Check if the player meets the speed requirement.
    If GetPlayerSPEED(Index) < Item(GetPlayerInvItemNum(Index, InvNum)).SpeedReq Then
        Call PlayerMsg(Index, "No tienes la suficiente velocidad para equiparte esto!", BRIGHTRED)
        Exit Function
    End If

    ItemIsUsable = True
End Function

Function ItemIsEquipped(ByVal Index As Long, ByVal ItemNum As Long) As Boolean
    If GetPlayerWeaponSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerArmorSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerShieldSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerHelmetSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerLegsSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerRingSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerNecklaceSlot(Index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If
End Function


Public Function DirToX(ByVal X As Long, _
   ByVal Dir As Byte) As Long
    DirToX = X

    If Dir = DIR_UP Or Dir = DIR_DOWN Then Exit Function

    ' LEFT = 2, RIGHT = 3
    ' 2 * 2 = 4, 4 - 5 = -1
    ' 3 * 2 = 6, 6 - 5 = 1
    DirToX = X + ((Dir * 2) - 5)
End Function

Public Function DirToY(ByVal Y As Long, _
   ByVal Dir As Byte) As Long
    DirToY = Y

    If Dir = DIR_LEFT Or Dir = DIR_RIGHT Then Exit Function

    ' UP = 0, DOWN = 1
    ' 0 * 2 = 0, 0 - 1 = -1
    ' 1 * 2 = 2, 2 - 1 = 1
    DirToY = Y + ((Dir * 2) - 1)
End Function

Sub GMGiveItem(ByVal N As Long, ByVal I As Integer)
Dim q As Long

                q = 1
               
                Do While q < 25
                   If GetPlayerInvItemNum(N, q) = 0 Then
                        Call SetPlayerInvItemNum(N, q, I)
                        Call SetPlayerInvItemValue(N, q, 1)
                        Call SetPlayerInvItemDur(N, q, Item(I).Data1)
                        Call SendInventoryUpdate(N, q)
                        Call PlayerMsg(N, "Un objeto ha sido añadido en tu inventario por un administrador.", Green)
                        Call SendPlayerData(N)
                        q = 25
                    End If
                    q = q + 1
                Loop
        Exit Sub

End Sub

Sub GMTakeItem(ByVal N As Long, ByVal I As Integer)
Dim TakeText
Dim TakeTextAmount
Dim N1 As Long

   
    Call TakeItem(N, I, 1)
    Call PlayerMsg(N, "Un objeto de tu inventario ha sido borrado por un administrador.", Red)
    Call SendPlayerData(N)
    Call SendStats(N)
   
End Sub

Sub ClearParties()
Dim I, o As Long

    For I = 1 To MAX_PARTIES
        For o = 1 To MAX_PARTY_MEMBERS
            Party(I).Member(o) = 0
        Next
    Next
End Sub

