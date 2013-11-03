Attribute VB_Name = "modGameLogic"
Option Explicit

Function GetPlayerDamage(ByVal index As Long) As Long
    Dim WeaponSlot As Long, RingSlot As Long, NecklaceSlot As Long
    GetPlayerDamage = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If

    ' GetPlayerDamage in script - TODO LATER - Can't get it to work. :(
    ' If Scripting = 1 Then
    ' GetPlayerDamage = MyScript.RunCodeReturn("Scripts\Main.txt", "GetPlayerDamage ", index)
    ' Else
    GetPlayerDamage = Int(GetPlayerSTR(index) / 2)
' End If

    If GetPlayerDamage <= 0 Then
        GetPlayerDamage = 1
    End If

    If GetPlayerWeaponSlot(index) > 0 Then
        WeaponSlot = GetPlayerWeaponSlot(index)

        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(index, WeaponSlot)).Data2

        If GetPlayerInvItemDur(index, WeaponSlot) > -1 Then
            Call SetPlayerInvItemDur(index, WeaponSlot, GetPlayerInvItemDur(index, WeaponSlot) - 1)

            If GetPlayerInvItemDur(index, WeaponSlot) = 0 Then
                Call BattleMsg(index, "Tu " & Trim$(Item(GetPlayerInvItemNum(index, WeaponSlot)).Name) & " se ha roto.", YELLOW, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, WeaponSlot), 0)
            Else
                If GetPlayerInvItemDur(index, WeaponSlot) <= 10 Then
                    Call BattleMsg(index, "Tu " & Trim$(Item(GetPlayerInvItemNum(index, WeaponSlot)).Name) & " esta apunto de romperse. Duración: " & GetPlayerInvItemDur(index, WeaponSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(index, WeaponSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If

    If GetPlayerRingSlot(index) > 0 Then
        RingSlot = GetPlayerRingSlot(index)

        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(index, RingSlot)).Data2

        If GetPlayerInvItemDur(index, RingSlot) > -1 Then
            Call SetPlayerInvItemDur(index, RingSlot, GetPlayerInvItemDur(index, RingSlot) - 1)

            If GetPlayerInvItemDur(index, RingSlot) = 0 Then
                Call BattleMsg(index, "Tu " & Trim$(Item(GetPlayerInvItemNum(index, RingSlot)).Name) & " se ha roto.", YELLOW, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, RingSlot), 0)
            Else
                If GetPlayerInvItemDur(index, RingSlot) <= 10 Then
                    Call BattleMsg(index, "Tu " & Trim$(Item(GetPlayerInvItemNum(index, RingSlot)).Name) & " esta apunto de romperse. Duración: " & GetPlayerInvItemDur(index, RingSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(index, RingSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If

    If GetPlayerNecklaceSlot(index) > 0 Then
        NecklaceSlot = GetPlayerNecklaceSlot(index)

        GetPlayerDamage = GetPlayerDamage + Item(GetPlayerInvItemNum(index, NecklaceSlot)).Data2

        If GetPlayerInvItemDur(index, NecklaceSlot) > -1 Then
            Call SetPlayerInvItemDur(index, NecklaceSlot, GetPlayerInvItemDur(index, NecklaceSlot) - 1)

            If GetPlayerInvItemDur(index, NecklaceSlot) = 0 Then
                Call BattleMsg(index, "Tu " & Trim$(Item(GetPlayerInvItemNum(index, NecklaceSlot)).Name) & " se ha roto.", YELLOW, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, NecklaceSlot), 0)
            Else
                If GetPlayerInvItemDur(index, NecklaceSlot) <= 10 Then
                    Call BattleMsg(index, "Tu " & Trim$(Item(GetPlayerInvItemNum(index, NecklaceSlot)).Name) & " esta apunto de romperse. Duración: " & GetPlayerInvItemDur(index, NecklaceSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(index, NecklaceSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If



    If GetPlayerDamage < 0 Then
        GetPlayerDamage = 0
    End If
End Function

Function GetPlayerProtection(ByVal index As Long) As Long
    Dim ArmorSlot As Long, HelmSlot As Long, ShieldSlot As Long, LegsSlot As Long

    GetPlayerProtection = 0

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > MAX_PLAYERS Then
        Exit Function
    End If

    ArmorSlot = GetPlayerArmorSlot(index)
    HelmSlot = GetPlayerHelmetSlot(index)
    ShieldSlot = GetPlayerShieldSlot(index)
    LegsSlot = GetPlayerLegsSlot(index)
    GetPlayerProtection = Int(GetPlayerDEF(index) / 5)

    If ArmorSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(index, ArmorSlot)).Data2
        If GetPlayerInvItemDur(index, ArmorSlot) > -1 Then
            Call SetPlayerInvItemDur(index, ArmorSlot, GetPlayerInvItemDur(index, ArmorSlot) - 1)

            If GetPlayerInvItemDur(index, ArmorSlot) = 0 Then
                Call BattleMsg(index, "Tu " & Trim$(Item(GetPlayerInvItemNum(index, ArmorSlot)).Name) & " se ha roto.", YELLOW, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, ArmorSlot), 0)
            Else
                If GetPlayerInvItemDur(index, ArmorSlot) <= 10 Then
                    Call BattleMsg(index, "Tu " & Trim$(Item(GetPlayerInvItemNum(index, ArmorSlot)).Name) & " esta apunto de romperse. Duración: " & GetPlayerInvItemDur(index, ArmorSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(index, ArmorSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If

    If HelmSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(index, HelmSlot)).Data2
        If GetPlayerInvItemDur(index, HelmSlot) > -1 Then
            Call SetPlayerInvItemDur(index, HelmSlot, GetPlayerInvItemDur(index, HelmSlot) - 1)

            If GetPlayerInvItemDur(index, HelmSlot) <= 0 Then
                Call BattleMsg(index, "Tu " & Trim$(Item(GetPlayerInvItemNum(index, HelmSlot)).Name) & " se ha roto.", YELLOW, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, HelmSlot), 0)
            Else
                If GetPlayerInvItemDur(index, HelmSlot) <= 10 Then
                    Call BattleMsg(index, "Tu " & Trim$(Item(GetPlayerInvItemNum(index, HelmSlot)).Name) & " " & Trim$(Item(GetPlayerInvItemNum(index, ArmorSlot)).Name) & " esta apunto de romperse. Duración: " & GetPlayerInvItemDur(index, HelmSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(index, HelmSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If

    If ShieldSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(index, ShieldSlot)).Data2
        If GetPlayerInvItemDur(index, ShieldSlot) > -1 Then
            Call SetPlayerInvItemDur(index, ShieldSlot, GetPlayerInvItemDur(index, ShieldSlot) - 1)

            If GetPlayerInvItemDur(index, ShieldSlot) <= 0 Then
                Call BattleMsg(index, "Tu " & Trim$(Item(GetPlayerInvItemNum(index, ShieldSlot)).Name) & " se ha roto.", YELLOW, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, ShieldSlot), 0)
            Else
                If GetPlayerInvItemDur(index, ShieldSlot) <= 10 Then
                    Call BattleMsg(index, "Tu " & Trim$(Item(GetPlayerInvItemNum(index, ShieldSlot)).Name) & " esta apunto de romperse. Duración: " & GetPlayerInvItemDur(index, ShieldSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(index, ShieldSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If

    If LegsSlot > 0 Then
        GetPlayerProtection = GetPlayerProtection + Item(GetPlayerInvItemNum(index, LegsSlot)).Data2
        If GetPlayerInvItemDur(index, LegsSlot) > -1 Then
            Call SetPlayerInvItemDur(index, LegsSlot, GetPlayerInvItemDur(index, LegsSlot) - 1)

            If GetPlayerInvItemDur(index, LegsSlot) <= 0 Then
                Call BattleMsg(index, "Tu " & Trim$(Item(GetPlayerInvItemNum(index, LegsSlot)).Name) & " se ha roto.", YELLOW, 0)
                Call TakeItem(index, GetPlayerInvItemNum(index, LegsSlot), 0)
            Else
                If GetPlayerInvItemDur(index, LegsSlot) <= 10 Then
                    Call BattleMsg(index, "Tu " & Trim$(Item(GetPlayerInvItemNum(index, LegsSlot)).Name) & " " & Trim$(Item(GetPlayerInvItemNum(index, ArmorSlot)).Name) & " esta apunto de romperse. Duración: " & GetPlayerInvItemDur(index, LegsSlot) & "/" & Trim$(Item(GetPlayerInvItemNum(index, LegsSlot)).Data1), YELLOW, 0)
                End If
            End If
        End If
    End If
End Function

Function FindOpenPlayerSlot() As Long
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If
    Next i
End Function

Public Function FindOpenInvSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(index) Then
        Exit Function
    End If

    ' Check to see if they already have the item.
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If
        Next i
    End If

    ' Try to find an open inventory slot.
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If
    Next i
End Function

Function FindOpenBankSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(index) Then
        Exit Function
    End If

    ' Check to see if they already have the item.
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(index, i) = ItemNum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next i
    End If

    ' Try to find an open bank slot.
    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next i
End Function

Function FindOpenMapItemSlot(ByVal mapnum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range.
    If mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_MAP_ITEMS
        If MapItem(mapnum, i).num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If
    Next i
End Function

Function FindOpenSpellSlot(ByVal index As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If
    Next i
End Function

Function HasSpell(ByVal index As Long, ByVal spellnum As Long) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) = spellnum Then
            HasSpell = True
            Exit Function
        End If
    Next i
End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If
    Next i
End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    Name = LCase$(Name)

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If Len(GetPlayerName(i)) >= Len(Name) Then
                If LCase$(GetPlayerName(i)) = Name Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

Function HasItem(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(index) Then
        Exit Function
    End If

    ' Check to see if the player has the item.
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                HasItem = GetPlayerInvItemValue(index, i)
            Else
                HasItem = 1
            End If

            Exit Function
        End If
    Next i
End Function

Sub TakeItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
    Dim i As Long, N As Long
    Dim TakeItem As Boolean

    TakeItem = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    Call SendInventoryUpdate(index, i)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerInvItemNum(index, i)).Type
                    Case ITEM_TYPE_WEAPON
                        If GetPlayerWeaponSlot(index) > 0 Then
                            If i = GetPlayerWeaponSlot(index) Then
                                Call SetPlayerWeaponSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If

                    Case ITEM_TYPE_ARMOR
                        If GetPlayerArmorSlot(index) > 0 Then
                            If i = GetPlayerArmorSlot(index) Then
                                Call SetPlayerArmorSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerArmorSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If

                    Case ITEM_TYPE_HELMET
                        If GetPlayerHelmetSlot(index) > 0 Then
                            If i = GetPlayerHelmetSlot(index) Then
                                Call SetPlayerHelmetSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If

                    Case ITEM_TYPE_SHIELD
                        If GetPlayerShieldSlot(index) > 0 Then
                            If i = GetPlayerShieldSlot(index) Then
                                Call SetPlayerShieldSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerShieldSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If

                    Case ITEM_TYPE_LEGS
                        If GetPlayerLegsSlot(index) > 0 Then
                            If i = GetPlayerLegsSlot(index) Then
                                Call SetPlayerLegsSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerLegsSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If

                    Case ITEM_TYPE_RING
                        If GetPlayerRingSlot(index) > 0 Then
                            If i = GetPlayerRingSlot(index) Then
                                Call SetPlayerRingSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerRingSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If

                    Case ITEM_TYPE_NECKLACE
                        If GetPlayerNecklaceSlot(index) > 0 Then
                            If i = GetPlayerNecklaceSlot(index) Then
                                Call SetPlayerNecklaceSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                End Select


                N = Item(GetPlayerInvItemNum(index, i)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (N <> ITEM_TYPE_WEAPON) And (N <> ITEM_TYPE_ARMOR) And (N <> ITEM_TYPE_HELMET) And (N <> ITEM_TYPE_SHIELD) And (N <> ITEM_TYPE_LEGS) And (N <> ITEM_TYPE_RING) And (N <> ITEM_TYPE_NECKLACE) Then
                    TakeItem = True
                End If
            End If

            If TakeItem = True Then
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                Call SetPlayerInvItemDur(index, i, 0)

                ' Send the inventory update
                Call SendInventoryUpdate(index, i)
                Exit Sub
            End If
        End If
    Next i
End Sub

Sub GiveItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
    Dim i As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(index) Then
        Exit Sub
    End If

    i = FindOpenInvSlot(index, ItemNum)

    ' Check to see if inventory is full
    If i > 0 Then
        Call SetPlayerInvItemNum(index, i, ItemNum)
        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)

        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Or (Item(ItemNum).Type = ITEM_TYPE_LEGS) Or (Item(ItemNum).Type = ITEM_TYPE_RING) Or (Item(ItemNum).Type = ITEM_TYPE_NECKLACE) Then
            Call SetPlayerInvItemDur(index, i, Item(ItemNum).Data1)
        End If

        Call SendInventoryUpdate(index, i)
    Else
        Call PlayerMsg(index, "Tu inventario esta lleno!", BRIGHTRED)
    End If
End Sub

Sub TakeBankItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
    Dim i As Long, N As Long
    Dim TakeBankItem As Boolean

    TakeBankItem = False

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    For i = 1 To MAX_BANK
        ' Check to see if the player has the item
        If GetPlayerBankItemNum(index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Or Item(ItemNum).Stackable = 1 Then
                ' Is what we are trying to take away more then what they have? If so just set it to zero
                If ItemVal >= GetPlayerBankItemValue(index, i) Then
                    TakeBankItem = True
                Else
                    Call SetPlayerBankItemValue(index, i, GetPlayerBankItemValue(index, i) - ItemVal)
                    Call SendBankUpdate(index, i)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerBankItemNum(index, i)).Type
                    Case ITEM_TYPE_WEAPON
                        If GetPlayerWeaponSlot(index) > 0 Then
                            If i = GetPlayerWeaponSlot(index) Then
                                Call SetPlayerWeaponSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(index, GetPlayerWeaponSlot(index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If

                    Case ITEM_TYPE_ARMOR
                        If GetPlayerArmorSlot(index) > 0 Then
                            If i = GetPlayerArmorSlot(index) Then
                                Call SetPlayerArmorSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(index, GetPlayerArmorSlot(index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If

                    Case ITEM_TYPE_HELMET
                        If GetPlayerHelmetSlot(index) > 0 Then
                            If i = GetPlayerHelmetSlot(index) Then
                                Call SetPlayerHelmetSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(index, GetPlayerHelmetSlot(index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If

                    Case ITEM_TYPE_SHIELD
                        If GetPlayerShieldSlot(index) > 0 Then
                            If i = GetPlayerShieldSlot(index) Then
                                Call SetPlayerShieldSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(index, GetPlayerShieldSlot(index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If

                    Case ITEM_TYPE_LEGS
                        If GetPlayerLegsSlot(index) > 0 Then
                            If i = GetPlayerLegsSlot(index) Then
                                Call SetPlayerLegsSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(index, GetPlayerLegsSlot(index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If

                    Case ITEM_TYPE_RING
                        If GetPlayerRingSlot(index) > 0 Then
                            If i = GetPlayerRingSlot(index) Then
                                Call SetPlayerRingSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(index, GetPlayerRingSlot(index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If

                    Case ITEM_TYPE_NECKLACE
                        If GetPlayerNecklaceSlot(index) > 0 Then
                            If i = GetPlayerNecklaceSlot(index) Then
                                Call SetPlayerNecklaceSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeBankItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerBankItemNum(index, GetPlayerNecklaceSlot(index)) Then
                                    TakeBankItem = True
                                End If
                            End If
                        Else
                            TakeBankItem = True
                        End If
                End Select


                N = Item(GetPlayerBankItemNum(index, i)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (N <> ITEM_TYPE_WEAPON) And (N <> ITEM_TYPE_ARMOR) And (N <> ITEM_TYPE_HELMET) And (N <> ITEM_TYPE_SHIELD) And (N <> ITEM_TYPE_LEGS) And (N <> ITEM_TYPE_RING) And (N <> ITEM_TYPE_NECKLACE) Then
                    TakeBankItem = True
                End If
            End If

            If TakeBankItem = True Then
                Call SetPlayerBankItemNum(index, i, 0)
                Call SetPlayerBankItemValue(index, i, 0)
                Call SetPlayerBankItemDur(index, i, 0)

                ' Send the Bank update
                Call SendBankUpdate(index, i)
                Exit Sub
            End If
        End If
    Next i
End Sub

Sub GiveBankItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal BankSlot As Long)
    Dim i As Long

    ' Check for subscript out of range.
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If Not IsPlaying(index) Then
        Exit Sub
    End If

    i = BankSlot

    ' Check to see if Bankentory is full
    If i > 0 Then
        Call SetPlayerBankItemNum(index, i, ItemNum)
        Call SetPlayerBankItemValue(index, i, GetPlayerBankItemValue(index, i) + ItemVal)

        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Or (Item(ItemNum).Type = ITEM_TYPE_LEGS) Or (Item(ItemNum).Type = ITEM_TYPE_RING) Or (Item(ItemNum).Type = ITEM_TYPE_NECKLACE) Then
            Call SetPlayerBankItemDur(index, i, Item(ItemNum).Data1)
        End If
    Else
        Call SendDataTo(index, "bankmsg" & SEP_CHAR & "Banco lleno!" & END_CHAR)
    End If
End Sub

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim i As Long

    ' Check for subscript out of range.
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Find open map item slot.
    i = FindOpenMapItemSlot(mapnum)

    Call SpawnItemSlot(i, ItemNum, ItemVal, Item(ItemNum).Data1, mapnum, x, y)
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim i As Long

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

    i = MapItemSlot

    If i > 0 Then
        MapItem(mapnum, i).num = ItemNum
        MapItem(mapnum, i).Value = ItemVal

        If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_NECKLACE) Then
            MapItem(mapnum, i).Dur = ItemDur
        Else
            MapItem(mapnum, i).Dur = 0
        End If

        MapItem(mapnum, i).x = x
        MapItem(mapnum, i).y = y

        Call SendDataToMap(mapnum, "SPAWNITEM" & SEP_CHAR & i & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(mapnum, i).Dur & SEP_CHAR & x & SEP_CHAR & y & END_CHAR)
    End If
End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next i
End Sub

Sub SpawnMapItems(ByVal mapnum As Long)
    Dim x As Integer
    Dim y As Integer

    ' Check for subscript out of range.
    If mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Spawn all the mapped items on their specified tile.
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            If Map(mapnum).Tile(x, y).Type = TILE_TYPE_ITEM Then
                If (Item(Map(mapnum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY Or Item(Map(mapnum).Tile(x, y).Data1).Stackable = 1) And Map(mapnum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(mapnum).Tile(x, y).Data1, 1, mapnum, x, y)
                Else
                    Call SpawnItem(Map(mapnum).Tile(x, y).Data1, Map(mapnum).Tile(x, y).Data2, mapnum, x, y)
                End If
            End If
        Next x
    Next y
End Sub

Sub PlayerMapGetItem(ByVal index As Long)
    Dim i As Long
    Dim N As Long
    Dim mapnum As Long
    Dim Msg As String

    If IsPlaying(index) = False Then
        Exit Sub
    End If

    mapnum = GetPlayerMap(index)

    For i = 1 To MAX_MAP_ITEMS
        ' See if theres even an item here
        If (MapItem(mapnum, i).num > 0) Then
            If (MapItem(mapnum, i).num <= MAX_ITEMS) Then
        
                ' Check if item is at the same location as the player
                If (MapItem(mapnum, i).x = GetPlayerX(index)) Then
                    If (MapItem(mapnum, i).y = GetPlayerY(index)) Then
                    
                        ' Find open slot
                        N = FindOpenInvSlot(index, MapItem(mapnum, i).num)
        
                        ' Open slot available?
                        If N <> 0 Then
                            ' Set item in players inventory
                            Call SetPlayerInvItemNum(index, N, MapItem(mapnum, i).num)
                            If Item(GetPlayerInvItemNum(index, N)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, N)).Stackable = 1 Then
                                Call SetPlayerInvItemValue(index, N, GetPlayerInvItemValue(index, N) + MapItem(mapnum, i).Value)
                                Msg = "Tu obtienes " & MapItem(mapnum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(index, N)).Name) & "."
                            Else
                                Call SetPlayerInvItemValue(index, N, 0)
                                Msg = "Tu obtienes " & Trim$(Item(GetPlayerInvItemNum(index, N)).Name) & "."
                            End If
                            Call SetPlayerInvItemDur(index, N, MapItem(mapnum, i).Dur)
        
                            ' Borra todos los objetos del mapa.
                            Call ClearMapItem(i, mapnum)
        
                            Call SendInventoryUpdate(index, N)
                            Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
                            Call PlayerMsg(index, Msg, YELLOW)
                            Exit Sub
                        Else
                            Call PlayerMsg(index, "Tu inventario está lleno!", BRIGHTRED)
                            Exit Sub
                        End If
                    End If
                End If
                
            End If
        End If
    Next i
End Sub

Sub PlayerMapDropItem(ByVal index As Long, ByVal InvNum As Long, ByVal Amount As Long)
    Dim i As Long
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If

    If (GetPlayerInvItemNum(index, InvNum) > 0) Then
        If (GetPlayerInvItemNum(index, InvNum) <= MAX_ITEMS) Then
            i = FindOpenMapItemSlot(GetPlayerMap(index))
    
            If i <> 0 Then
                MapItem(GetPlayerMap(index), i).Dur = 0
    
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerInvItemNum(index, InvNum)).Type
                    Case ITEM_TYPE_ARMOR
                        If InvNum = GetPlayerArmorSlot(index) Then
                            Call SetPlayerArmorSlot(index, 0)
                            Call SendWornEquipment(index)
                            Call SendIndexWornEquipment(index)
                        End If
                        MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
    
                    Case ITEM_TYPE_WEAPON
                        If InvNum = GetPlayerWeaponSlot(index) Then
                            Call SetPlayerWeaponSlot(index, 0)
                            Call SendWornEquipment(index)
                            Call SendIndexWornEquipment(index)
                        End If
                        MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
    
                    Case ITEM_TYPE_HELMET
                        If InvNum = GetPlayerHelmetSlot(index) Then
                            Call SetPlayerHelmetSlot(index, 0)
                            Call SendWornEquipment(index)
                            Call SendIndexWornEquipment(index)
                        End If
                        MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
    
                    Case ITEM_TYPE_SHIELD
                        If InvNum = GetPlayerShieldSlot(index) Then
                            Call SetPlayerShieldSlot(index, 0)
                            Call SendWornEquipment(index)
                            Call SendIndexWornEquipment(index)
                        End If
                        MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                    Case ITEM_TYPE_LEGS
                        If InvNum = GetPlayerLegsSlot(index) Then
                            Call SetPlayerLegsSlot(index, 0)
                            Call SendWornEquipment(index)
                            Call SendIndexWornEquipment(index)
                        End If
                        MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                    Case ITEM_TYPE_RING
                        If InvNum = GetPlayerRingSlot(index) Then
                            Call SetPlayerRingSlot(index, 0)
                            Call SendWornEquipment(index)
                            Call SendIndexWornEquipment(index)
                        End If
                        MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                    Case ITEM_TYPE_NECKLACE
                        If InvNum = GetPlayerNecklaceSlot(index) Then
                            Call SetPlayerNecklaceSlot(index, 0)
                            Call SendWornEquipment(index)
                            Call SendIndexWornEquipment(index)
                        End If
                        MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                End Select
    
                MapItem(GetPlayerMap(index), i).num = GetPlayerInvItemNum(index, InvNum)
                MapItem(GetPlayerMap(index), i).x = GetPlayerX(index)
                MapItem(GetPlayerMap(index), i).y = GetPlayerY(index)
    
                If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
                    ' Check if its more then they have and if so drop it all
                    If Amount >= GetPlayerInvItemValue(index, InvNum) Then
                        MapItem(GetPlayerMap(index), i).Value = GetPlayerInvItemValue(index, InvNum)
                        ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, InvNum) & " " & Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemNum(index, InvNum, 0)
                        Call SetPlayerInvItemValue(index, InvNum, 0)
                        Call SetPlayerInvItemDur(index, InvNum, 0)
                    Else
                        MapItem(GetPlayerMap(index), i).Value = Amount
                        ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & Amount & " " & Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                        Call SetPlayerInvItemValue(index, InvNum, GetPlayerInvItemValue(index, InvNum) - Amount)
                    End If
                Else
                    ' Its not a currency object so this is easy
                    MapItem(GetPlayerMap(index), i).Value = 0
    
                    ' Normally messages for item drops would go here but it's scripted now
    
                    Call SetPlayerInvItemNum(index, InvNum, 0)
                    Call SetPlayerInvItemValue(index, InvNum, 0)
                    Call SetPlayerInvItemDur(index, InvNum, 0)
                End If
    
                ' Send inventory update
                Call SendInventoryUpdate(index, InvNum)
                ' Spawn the item before we set the num or we'll get a different free map item slot
                Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).num, Amount, MapItem(GetPlayerMap(index), i).Dur, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    
                If scripting = 1 Then
                    MyScript.ExecuteStatement "Scripts\Main.txt", "onitemdrop " & index & "," & GetPlayerMap(index) & "," & MapItem(GetPlayerMap(index), i).num & "," & Amount & "," & MapItem(GetPlayerMap(index), i).Dur & "," & i & "," & InvNum
                End If
    
            Else
                Call PlayerMsg(index, "Hay demasiado objetos en el suelo.", BRIGHTRED)
            End If
        End If
        
    End If
End Sub

Sub SpawnNpc(ByVal MapNpcNum As Long, ByVal mapnum As Long)
    Dim Packet As String
    Dim npcnum As Long
    Dim i As Long
    Dim x As Long
    Dim y As Long
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
            For i = 1 To 100
                x = Int(Rnd * MAX_MAPX)
                y = Int(Rnd * MAX_MAPY)
    
                If Map(mapnum).Tile(x, y).Type = TILE_TYPE_WALKABLE Then
                    MapNPC(mapnum, MapNpcNum).x = x
                    MapNPC(mapnum, MapNpcNum).y = y
                    Spawned = True
                    Exit For
                End If
            Next i

            ' Didn't spawn, so now we'll just try to find a free tile
            If Not Spawned Then
                For y = 0 To MAX_MAPY
                    For x = 0 To MAX_MAPX
                        If Map(mapnum).Tile(x, y).Type = TILE_TYPE_WALKABLE Then
                            MapNPC(mapnum, MapNpcNum).x = x
                            MapNPC(mapnum, MapNpcNum).y = y
                            Spawned = True
                        End If
                    Next x
                Next y
            End If
        Else
            ' We subtract one because Rand is ListIndex 0. [Mellowz]
            MapNPC(mapnum, MapNpcNum).x = Map(mapnum).SpawnX(MapNpcNum) - 1
            MapNPC(mapnum, MapNpcNum).y = Map(mapnum).SpawnY(MapNpcNum) - 1
            Spawned = True
        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Packet = "SPAWNNPC" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(mapnum, MapNpcNum).num & SEP_CHAR & MapNPC(mapnum, MapNpcNum).x & SEP_CHAR & MapNPC(mapnum, MapNpcNum).y & SEP_CHAR & MapNPC(mapnum, MapNpcNum).Dir & SEP_CHAR & NPC(MapNPC(mapnum, MapNpcNum).num).Big & END_CHAR
            Call SendDataToMap(mapnum, Packet)
        End If
    End If

    ' Enable this to display HP when monsters spawn.
    ' Call SendDataToMap(MapNum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(MapNum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNPC(MapNum, MapNpcNum).num) & END_CHAR)
End Sub

Sub SpawnMapNpcs(ByVal mapnum As Long)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        If Map(mapnum).NPC(i) > 0 Then
            Call SpawnNpc(i, mapnum)
        End If
    Next i
End Sub

Sub SpawnAllMapNpcs()
    Dim i As Long

    For i = 1 To MAX_MAPS
        If PlayersOnMap(i) = YES Then
            Call SpawnMapNpcs(i)
        End If
    Next i
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
                If (MapNPC(mapnum, MapNpcNum).y + 1 = GetPlayerY(Attacker)) And (MapNPC(mapnum, MapNpcNum).x = GetPlayerX(Attacker)) Then
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
                If (MapNPC(mapnum, MapNpcNum).y - 1 = GetPlayerY(Attacker)) And (MapNPC(mapnum, MapNpcNum).x = GetPlayerX(Attacker)) Then
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
                If (MapNPC(mapnum, MapNpcNum).y = GetPlayerY(Attacker)) And (MapNPC(mapnum, MapNpcNum).x + 1 = GetPlayerX(Attacker)) Then
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
                If (MapNPC(mapnum, MapNpcNum).y = GetPlayerY(Attacker)) And (MapNPC(mapnum, MapNpcNum).x - 1 = GetPlayerX(Attacker)) Then
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


Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long) As Boolean
    Dim mapnum As Long
    Dim npcnum As Long

    If Not IsPlaying(index) Then
        Exit Function
    End If

    ' Make sure the NPC map number isn't out-of-range.
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then
        Exit Function
    End If

    ' Make sure that it's a valid NPC.
    If MapNPC(GetPlayerMap(index), MapNpcNum).num < 1 Then
        Exit Function
    End If

    mapnum = GetPlayerMap(index)
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
    If Player(index).GettingMap = YES Then
        Exit Function
    End If

    MapNPC(mapnum, MapNpcNum).AttackTimer = GetTickCount

    If IsPlaying(index) Then
        If npcnum > 0 Then
            If (GetPlayerY(index) + 1 = MapNPC(mapnum, MapNpcNum).y) And (GetPlayerX(index) = MapNPC(mapnum, MapNpcNum).x) Then
                CanNpcAttackPlayer = True
            Else
                If (GetPlayerY(index) - 1 = MapNPC(mapnum, MapNpcNum).y) And (GetPlayerX(index) = MapNPC(mapnum, MapNpcNum).x) Then
                    CanNpcAttackPlayer = True
                Else
                    If (GetPlayerY(index) = MapNPC(mapnum, MapNpcNum).y) And (GetPlayerX(index) + 1 = MapNPC(mapnum, MapNpcNum).x) Then
                        CanNpcAttackPlayer = True
                    Else
                        If (GetPlayerY(index) = MapNPC(mapnum, MapNpcNum).y) And (GetPlayerX(index) - 1 = MapNPC(mapnum, MapNpcNum).x) Then
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
    Dim N As Long, i As Long, q As Integer, x As Long
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

        If DoubleExp Then EXP = EXP * 2

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
            For i = 1 To MAX_PARTY_MEMBERS

                If Party(Player(Attacker).PartyID).Member(i) <> Attacker Then
                    If Party(Player(Attacker).PartyID).Member(i) <> 0 Then
                        If GetPlayerMap(Attacker) = GetPlayerMap(Party(Player(Attacker).PartyID).Member(i)) Then
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
                For i = 1 To o

                    If Party(Player(Attacker).PartyID).Member(i) <> Attacker And Party(Player(Attacker).PartyID).Member(i) <> 0 Then
                        If GetPlayerLevel(Party(Player(Attacker).PartyID).Member(i)) = MAX_LEVEL Then
                            Call SetPlayerExp(Party(Player(Attacker).PartyID).Member(i), Experience(MAX_LEVEL))
                            Call BattleMsg(Party(Player(Attacker).PartyID).Member(i), "No puedes ganar más experiencia!", BRIGHTBLUE, 0)
                        Else
                            Call SetPlayerExp(Party(Player(Attacker).PartyID).Member(i), Player(Party(Player(Attacker).PartyID).Member(i)).Char(Player(Party(Player(Attacker).PartyID).Member(i)).CharNum).EXP + Int(EXP * (0.25 / o)))
                            Call BattleMsg(Party(Player(Attacker).PartyID).Member(i), "Has ganado " & Int(EXP * (0.25 / o)) & " puntos de experiencia de tu grupo.", BRIGHTBLUE, 0)
                            'Call BattleMsg(Party(Player(Attacker).PartyID).Member(I), "PUTA", Red, 0)
                            Call SendStats(Party(Player(Attacker).PartyID).Member(i))
                            Call SendPlayerData(Party(Player(Attacker).PartyID).Member(i))
                        End If
                        
                    End If
                Next
            End If
        End If
        ' Drop the items if they earn it.
        For i = 1 To MAX_NPC_DROPS
            If NPC(npcnum).ItemNPC(i).ItemNum > 0 Then
                N = Int(Rnd * NPC(npcnum).ItemNPC(i).chance) + 1
                If N = 1 Then
                    Call SpawnItem(NPC(npcnum).ItemNPC(i).ItemNum, NPC(npcnum).ItemNPC(i).ItemValue, mapnum, MapNPC(mapnum, MapNpcNum).x, MapNPC(mapnum, MapNpcNum).y)
                End If
            End If
        Next i

        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNPC(mapnum, MapNpcNum).num = 0
        MapNPC(mapnum, MapNpcNum).SpawnWait = GetTickCount
        MapNPC(mapnum, MapNpcNum).HP = 0
        Call SendDataToMap(mapnum, "NPCDEAD" & SEP_CHAR & MapNpcNum & END_CHAR)

        ' Check for level up
        Call CheckPlayerLevelUp(Attacker)

        
                ' Check for level up party member
         If Player(Attacker).InParty = YES Then
            For x = 1 To MAX_PARTY_MEMBERS

                If Party(Player(Attacker).PartyID).Member(x) <> 0 Then
                    Call CheckPlayerLevelUp(Party(Player(Attacker).PartyID).Member(x))
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
            For i = 1 To MAX_MAP_NPCS
                If MapNPC(mapnum, i).num = MapNPC(mapnum, MapNpcNum).num Then
                    MapNPC(mapnum, i).Target = Attacker
                End If
            Next i
        End If
    End If

    Call SendDataToMap(mapnum, "npchp" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(mapnum, MapNpcNum).HP & SEP_CHAR & GetNpcMaxHP(MapNPC(mapnum, MapNpcNum).num) & END_CHAR)

    ' Reset attack timer
    Player(Attacker).AttackTimer = GetTickCount
End Sub

Sub JoinWarp(ByVal index As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim OldMap As Long

    ' Check for subscript out of range.
    If Not IsPlaying(index) Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Save current number map the player is on.
    OldMap = GetPlayerMap(index)

    Call SendLeaveMap(index, OldMap)

    Call SetPlayerMap(index, mapnum)
    Call SetPlayerX(index, x)
    Call SetPlayerY(index, y)

    ' Check to see if anyone is on the map.
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(mapnum) = YES

    Player(index).GettingMap = YES

    Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & mapnum & SEP_CHAR & Map(mapnum).Revision & END_CHAR)

    Call SendInventory(index)
    Call SendIndexWornEquipmentFromMap(index)
End Sub

Sub PlayerWarp(ByVal index As Long, ByVal mapnum As Long, ByVal x As Long, ByVal y As Long)
    Dim OldMap As Long

    On Error GoTo WarpErr

    ' Check for subscript out of range.
    If Not IsPlaying(index) Then
        Exit Sub
    End If

    ' Check for subscript out of range.
    If mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Sub
    End If

    ' Save current number map the player is on.
    OldMap = GetPlayerMap(index)

    If Not OldMap = mapnum Then
        Call SendLeaveMap(index, OldMap)
    End If

    Call SetPlayerMap(index, mapnum)
    Call SetPlayerX(index, x)
    Call SetPlayerY(index, y)

    ' Check to see if anyone is on the map.
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If
    
    If Player(index).Pet.Alive = YES Then
        Player(index).Pet.MapToGo = mapnum
        Player(index).Pet.Map = mapnum
        Player(index).Pet.x = x
        Player(index).Pet.y = y
        Call SetPlayerPetX(index, x)
        Call SetPlayerPetY(index, y)
        Call SetPlayerPetMap(index, mapnum)
    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(mapnum) = YES

    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "warp" & END_CHAR)

    Player(index).GettingMap = YES

    Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & mapnum & SEP_CHAR & Map(mapnum).Revision & END_CHAR)

    Call SendInventory(index)
    Call SendIndexInventoryFromMap(index)
    Call SendIndexWornEquipmentFromMap(index)

    If scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "OnMapLoad " & index & "," & OldMap & "," & mapnum
    End If
    
    Exit Sub

WarpErr:
    Call AddLog("PlayerWarp error for player index " & index & " on map " & GetPlayerMap(index) & ".", "logs\ErrorLog.txt")
End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal Movement As Long, Xpos As Integer, Ypos As Integer)
    Dim Packet As String
    Dim mapnum As Long
    Dim x As Long
    Dim y As Long
    Dim i As Long
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
    If IsPlaying(index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
        Exit Sub
    End If
    
    If Player(index).GettingMap = True Then
        Exit Sub
    End If

    ' Check for scrolling to prevent RTE 9
    If GetPlayerX(index) > MAX_MAPX Or GetPlayerY(index) > MAX_MAPY Then
        Call PlayerWarp(index, GetPlayerMap(index), 0, 0)
        Exit Sub
    End If

    Call SetPlayerDir(index, Dir)
    
    If Player(index).Pet.Alive = YES Then

    
        If Player(index).Pet.Map = GetPlayerMap(index) And Player(index).Pet.x = x And Player(index).Pet.y = y Then
           ' If Grid(GetPlayerMap(Index)).Loc(DirToX(x, Dir), DirToY(y, Dir)).Blocked = False Then
             '   Call UpdateGrid(Player(Index).Pet.Map, Player(Index).Pet.x, Player(Index).Pet.y, Player(Index).Pet.Map, DirToX(x, Dir), DirToY(y, Dir))
                Player(index).Pet.y = DirToY(y, Dir)
                Player(index).Pet.x = DirToX(x, Dir)
                Packet = "PETMOVE" & SEP_CHAR & index & SEP_CHAR & DirToX(x, Dir) & SEP_CHAR & DirToY(y, Dir) & SEP_CHAR & Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
                Call SendDataToMap(Player(index).Pet.Map, Packet)
           ' End If
        End If
    End If

    ' Remove SP if the player is running.
    If SP_RUNNING = 1 Then
        If Movement = MOVING_RUNNING Then
            If GetPlayerSP(index) > 0 Then
                Call SetPlayerSP(index, GetPlayerSP(index) - 1)
                Call SendSP(index)
            Else
                Call PlayerMsg(index, "Te sientes cansado para correr.", Blue)
            End If
        End If
    End If

    Moved = NO


' save the current location
    Xold = GetPlayerX(index)
    Yold = GetPlayerY(index)
    pmap = GetPlayerMap(index)
   
' validate map number
   
    If pmap <= 0 Or pmap > MAX_MAPS Then
        Call HackingAttempt(index, vbNullString)
        Exit Sub
    End If
   
   
' update it to match client - this will be correct 99% of the time
    Call SetPlayerX(index, Xpos)
    Call SetPlayerY(index, Ypos)
   
' next check to see if we have gone outside of map boundries
'   if we have, need to try to warp to next map if there is one

    If Dir = DIR_UP And Ypos < 0 And Map(pmap).Up > 0 Then
        Call PlayerWarp(index, Map(pmap).Up, Xpos, MAX_MAPY)
        Moved = YES
    ElseIf Dir = DIR_DOWN And Ypos > MAX_MAPY And Map(pmap).Down > 0 Then
        Call PlayerWarp(index, Map(pmap).Down, Xpos, 0)
        Moved = YES
    ElseIf Dir = DIR_LEFT And Xpos < 0 And Map(pmap).Left > 0 Then
        Call PlayerWarp(index, Map(pmap).Left, MAX_MAPX, Ypos)
        Moved = YES
    ElseIf Dir = DIR_RIGHT And Xpos > MAX_MAPX And Map(pmap).Right > 0 Then
        Call PlayerWarp(index, Map(pmap).Right, 0, Ypos)
        Moved = YES
    End If
   
' restore values in case we got warped

    Xpos = GetPlayerX(index)
    Ypos = GetPlayerY(index)
    pmap = GetPlayerMap(index)

' check to make sure new position is on the map

    If Xpos < 0 Or Ypos < 0 Or Xpos > MAX_MAPX Or Ypos > MAX_MAPY Then
        Call HackingAttempt(index, vbNullString)
        Exit Sub
    End If

' Check to make sure that the tile is walkable
    If Map(pmap).Tile(Xpos, Ypos).Type <> TILE_TYPE_BLOCKED Then
' Check to see if the tile is a key and if it is check if its opened
        If (Map(pmap).Tile(Xpos, Ypos).Type <> TILE_TYPE_KEY And Map(pmap).Tile(Xpos, Ypos).Type <> TILE_TYPE_DOOR) Or ((Map(pmap).Tile(Xpos, Ypos).Type = TILE_TYPE_DOOR Or Map(pmap).Tile(Xpos, Ypos).Type = TILE_TYPE_KEY) And TempTile(pmap).DoorOpen(Xpos, Ypos) = YES) Then
            Packet = "playermove" & SEP_CHAR & index & SEP_CHAR & Xpos & SEP_CHAR & Ypos & SEP_CHAR & Dir & SEP_CHAR & Movement & END_CHAR
            Call SendDataToMapBut(index, pmap, Packet)
            Moved = YES
        End If
    End If

' at this point we have either moved or there is a problem with the new location
'   if we didn't move, we need to reset to previous locations and quit

    If Moved <> YES Then
        Call SetPlayerX(index, Xold)
        Call SetPlayerY(index, Yold)
        Call SendPlayerNewXY(index)
        Exit Sub
    End If

    ' healing tiles code
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_HEAL Then
        Call SetPlayerHP(index, GetPlayerMaxHP(index))
        Call SetPlayerMP(index, GetPlayerMaxMP(index))
        Call SendHP(index)
        Call SendMP(index)
        Call PlayerMsg(index, "Sientes como una fuerza sanadora entra dentro de ti!", BRIGHTGREEN)
    End If

    ' Check for kill tile, and if so kill them
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_KILL Then
        Call SetPlayerHP(index, 0)
        Call PlayerMsg(index, "Sientes como la muerte penetra en tu cuerpo.", BRIGHTRED)

        ' Warp player away
        If scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "OnDeath " & index
        Else
            If Map(GetPlayerMap(index)).BootMap > 0 Then
                Call PlayerWarp(index, Map(GetPlayerMap(index)).BootMap, Map(GetPlayerMap(index)).BootX, Map(GetPlayerMap(index)).BootY)
            Else
                Call PlayerWarp(index, START_MAP, START_X, START_Y)
            End If
        End If
        Call SetPlayerHP(index, GetPlayerMaxHP(index))
        Call SetPlayerMP(index, GetPlayerMaxMP(index))
        Call SetPlayerSP(index, GetPlayerMaxSP(index))
        Call SendHP(index)
        Call SendMP(index)
        Call SendSP(index)
        Moved = YES
    End If

    If GetPlayerX(index) + 1 <= MAX_MAPX Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Then
            x = GetPlayerX(index) + 1
            y = GetPlayerY(index)

            If TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount

                Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "door" & END_CHAR)
            End If
        End If
    End If
    If GetPlayerX(index) - 1 >= 0 Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type = TILE_TYPE_DOOR Then
            x = GetPlayerX(index) - 1
            y = GetPlayerY(index)

            If TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount

                Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "door" & END_CHAR)
            End If
        End If
    End If
    If GetPlayerY(index) - 1 >= 0 Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type = TILE_TYPE_DOOR Then
            x = GetPlayerX(index)
            y = GetPlayerY(index) - 1

            If TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount

                Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "door" & END_CHAR)
            End If
        End If
    End If
    If GetPlayerY(index) + 1 <= MAX_MAPY Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type = TILE_TYPE_DOOR Then
            x = GetPlayerX(index)
            y = GetPlayerY(index) + 1

            If TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
                TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
                TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount

                Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
                Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "door" & END_CHAR)
            End If
        End If
    End If

    ' Check to see if the tile is a warp tile, and if so warp them
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_WARP Then
        mapnum = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        x = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2
        y = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3

        Call PlayerWarp(index, mapnum, x, y)
        Moved = YES
    End If

    ' Check for key trigger open
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_KEYOPEN Then
        x = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        y = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2

        If Map(GetPlayerMap(index)).Tile(x, y).Type = TILE_TYPE_KEY And TempTile(GetPlayerMap(index)).DoorOpen(x, y) = NO Then
            TempTile(GetPlayerMap(index)).DoorOpen(x, y) = YES
            TempTile(GetPlayerMap(index)).DoorTimer = GetTickCount

            Call SendDataToMap(GetPlayerMap(index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & END_CHAR)
            If Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1) = vbNullString Then
                Call MapMsg(GetPlayerMap(index), "La puerta ha sido abierta!", WHITE)
            Else
                Call MapMsg(GetPlayerMap(index), Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1), WHITE)
            End If
            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "key" & END_CHAR)
        End If
    End If

    ' Check for shop
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SHOP Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 > 0 Then
            Call SendTrade(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)
        Else
            Call PlayerMsg(index, "No hay una tienda aqui.", BRIGHTRED)
        End If
    End If

    ' Check if player stepped on sprite changing tile
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SPRITE_CHANGE Then
        If GetPlayerSprite(index) = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 Then
            Call PlayerMsg(index, "Ya tienes este sprite!", BRIGHTRED)
            Exit Sub
        Else
            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 = 0 Then
                Call SendDataTo(index, "spritechange" & SEP_CHAR & 0 & END_CHAR)
            Else
                If Item(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2).Type = ITEM_TYPE_CURRENCY Then
                    Call PlayerMsg(index, "Este sprite te costara " & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3 & " " & Trim$(Item(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2).Name) & "!", YELLOW)
                Else
                    Call PlayerMsg(index, "Este sprite te costara " & Trim$(Item(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2).Name) & "!", YELLOW)
                End If
                Call SendDataTo(index, "spritechange" & SEP_CHAR & 1 & END_CHAR)
            End If
        End If
    End If

    ' Check if player stepped on house buying tile
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_HOUSE Then
        If Len(Map(GetPlayerMap(index)).Owner) < 2 Then
            If GetPlayerName(index) = Map(GetPlayerMap(index)).Owner Then
                Call PlayerMsg(index, "Ya tienes esta casa!", BRIGHTRED)
                Exit Sub
            Else
                If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 = 0 Then
                    Call SendDataTo(index, "housebuy" & SEP_CHAR & 0 & END_CHAR)
                Else
                    If Item(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1).Type = ITEM_TYPE_CURRENCY Then
                        Call PlayerMsg(index, "Esta casa te costara " & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 & " " & Trim$(Item(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1).Name) & "!", YELLOW)
                    Else
                        Call PlayerMsg(index, "Esta casa te costara un " & Trim$(Item(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1).Name) & "!", YELLOW)
                    End If
                    Call SendDataTo(index, "housebuy" & SEP_CHAR & 1 & END_CHAR)
                End If
            End If
        Else
            Call PlayerMsg(index, "Esta casa no esta en venta!", BRIGHTRED)
            Exit Sub
        End If
    End If

    ' Check if player stepped on sprite changing tile
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_CLASS_CHANGE Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 > -1 Then
            If GetPlayerClass(index) <> Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2 Then
                Call PlayerMsg(index, "No eres la clase requerida!", BRIGHTRED)
                Exit Sub
            End If
        End If

        If GetPlayerClass(index) = Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1 Then
            Call PlayerMsg(index, "Ya eres esta clase!", BRIGHTRED)
        Else
            If Player(index).Char(Player(index).CharNum).Sex = 0 Then
                If GetPlayerSprite(index) = ClassData(GetPlayerClass(index)).MaleSprite Then
                    Call SetPlayerSprite(index, ClassData(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1).MaleSprite)
                End If
            Else
                If GetPlayerSprite(index) = ClassData(GetPlayerClass(index)).FemaleSprite Then
                    Call SetPlayerSprite(index, ClassData(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1).FemaleSprite)
                End If
            End If

            Call SetPlayerSTR(index, (Player(index).Char(Player(index).CharNum).STR - ClassData(GetPlayerClass(index)).STR))
            Call SetPlayerDEF(index, (Player(index).Char(Player(index).CharNum).DEF - ClassData(GetPlayerClass(index)).DEF))
            Call SetPlayerMAGI(index, (Player(index).Char(Player(index).CharNum).Magi - ClassData(GetPlayerClass(index)).Magi))
            Call SetPlayerSPEED(index, (Player(index).Char(Player(index).CharNum).Speed - ClassData(GetPlayerClass(index)).Speed))

            Call SetPlayerClassData(index, Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1)

            Call SetPlayerSTR(index, (Player(index).Char(Player(index).CharNum).STR + ClassData(GetPlayerClass(index)).STR))
            Call SetPlayerDEF(index, (Player(index).Char(Player(index).CharNum).DEF + ClassData(GetPlayerClass(index)).DEF))
            Call SetPlayerMAGI(index, (Player(index).Char(Player(index).CharNum).Magi + ClassData(GetPlayerClass(index)).Magi))
            Call SetPlayerSPEED(index, (Player(index).Char(Player(index).CharNum).Speed + ClassData(GetPlayerClass(index)).Speed))


            Call PlayerMsg(index, "Tu nueva clase es " & Trim$(ClassData(GetPlayerClass(index)).Name) & "!", BRIGHTGREEN)

            Call SendStats(index)
            Call SendHP(index)
            Call SendMP(index)
            Call SendSP(index)
            Player(index).Char(Player(index).CharNum).MAXHP = GetPlayerMaxHP(index)
            Player(index).Char(Player(index).CharNum).MAXMP = GetPlayerMaxMP(index)
            Player(index).Char(Player(index).CharNum).MAXSP = GetPlayerMaxSP(index)
            Call SendDataToMap(GetPlayerMap(index), "checksprite" & SEP_CHAR & index & SEP_CHAR & GetPlayerSprite(index) & END_CHAR)
        End If
    End If

    ' Check if player stepped on notice tile
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_NOTICE Then
        If Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1) <> vbNullString Then
            Call PlayerMsg(index, Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1), BLACK)
        End If
        If Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String2) <> vbNullString Then
            Call PlayerMsg(index, Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String2), GREY)
        End If
        If Not Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String3 = vbNullString Or Not Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String3 = vbNullString Then
            Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String3 & END_CHAR)
        End If
    End If

    ' Check if player steppted on minus stat tile
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_LOWER_STAT Then
        If Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1) <> vbNullString Then
            Call PlayerMsg(index, Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1), BLACK)
        End If
        If Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1) <> 0 Then
            Call SetPlayerHP(index, GetPlayerHP(index) - Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1))
        End If
        If Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2) <> 0 Then
            Call SetPlayerMP(index, GetPlayerMP(index) - Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data2))
        End If
        If Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3) <> 0 Then
            Call SetPlayerSP(index, GetPlayerSP(index) - Trim$(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data3))
        End If
    End If

    ' Check if player stepped on sound tile
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SOUND Then
        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).String1 & END_CHAR)
    End If

    If scripting = 1 Then
        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_SCRIPTED Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedTile " & index & "," & Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Data1
        End If
    End If

    ' Check if player stepped on Bank tile
    If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).Type = TILE_TYPE_BANK Then
        Call SendDataTo(index, "openbank" & END_CHAR)
    End If

End Sub

Function CanNpcMove(ByVal mapnum As Long, ByVal MapNpcNum As Long, ByVal Dir) As Boolean
    Dim i As Long
    Dim TileType As Long
    Dim x As Long
    Dim y As Long

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

    x = MapNPC(mapnum, MapNpcNum).x
    y = MapNPC(mapnum, MapNpcNum).y

    CanNpcMove = True

    Select Case Dir
        Case DIR_UP
            If y > 0 Then
                ' Get the attribute on the tile.
                TileType = Map(mapnum).Tile(x, y - 1).Type
                                
                ' Check to make sure that the tile is walkable.
                If TileType = TILE_TYPE_BLOCKED Or TileType = TILE_TYPE_NPCAVOID Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way.
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If GetPlayerMap(i) = mapnum Then
                            If GetPlayerX(i) = MapNPC(mapnum, MapNpcNum).x Then
                                If GetPlayerY(i) = (MapNPC(mapnum, MapNpcNum).y - 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next i

                ' Check to make sure that there is not another npc in the way.
                For i = 1 To MAX_MAP_NPCS
                    If i <> MapNpcNum Then
                        If MapNPC(mapnum, i).num > 0 Then
                            If MapNPC(mapnum, i).x = MapNPC(mapnum, MapNpcNum).x Then
                                If MapNPC(mapnum, i).y = (MapNPC(mapnum, MapNpcNum).y - 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next i
            Else
                CanNpcMove = False
            End If

        Case DIR_DOWN
            If y < MAX_MAPY Then
                ' Get the attribute on the tile.
                TileType = Map(mapnum).Tile(x, y + 1).Type
                
                ' Check to make sure that the tile is walkable.
                If TileType = TILE_TYPE_BLOCKED Or TileType = TILE_TYPE_NPCAVOID Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way.
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If GetPlayerMap(i) = mapnum Then
                            If GetPlayerX(i) = MapNPC(mapnum, MapNpcNum).x Then
                                If GetPlayerY(i) = (MapNPC(mapnum, MapNpcNum).y + 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next i

                ' Check to make sure that there is not another npc in the way.
                For i = 1 To MAX_MAP_NPCS
                    If i <> MapNpcNum Then
                        If MapNPC(mapnum, i).num > 0 Then
                            If MapNPC(mapnum, i).x = MapNPC(mapnum, MapNpcNum).x Then
                                If MapNPC(mapnum, i).y = (MapNPC(mapnum, MapNpcNum).y + 1) Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next i
            Else
                CanNpcMove = False
            End If

        Case DIR_LEFT
            If x > 0 Then
                ' Get the attribute on the tile.
                TileType = Map(mapnum).Tile(x - 1, y).Type

                ' Check to make sure that the tile is walkable.
                If TileType = TILE_TYPE_BLOCKED Or TileType = TILE_TYPE_NPCAVOID Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way.
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If GetPlayerMap(i) = mapnum Then
                            If GetPlayerX(i) = (MapNPC(mapnum, MapNpcNum).x - 1) Then
                                If GetPlayerY(i) = MapNPC(mapnum, MapNpcNum).y Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next i

                ' Check to make sure that there is not another npc in the way.
                For i = 1 To MAX_MAP_NPCS
                    If i <> MapNpcNum Then
                        If MapNPC(mapnum, i).num > 0 Then
                            If MapNPC(mapnum, i).x = (MapNPC(mapnum, MapNpcNum).x - 1) Then
                                If MapNPC(mapnum, i).y = MapNPC(mapnum, MapNpcNum).y Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next i
            Else
                CanNpcMove = False
            End If

        Case DIR_RIGHT
            If x < MAX_MAPX Then
                ' Get the attribute on the tile.
                TileType = Map(mapnum).Tile(x + 1, y).Type
                
                ' Check to make sure that the tile is walkable.
                If TileType = TILE_TYPE_BLOCKED Or TileType = TILE_TYPE_NPCAVOID Then
                    CanNpcMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way.
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) Then
                        If GetPlayerMap(i) = mapnum Then
                            If GetPlayerX(i) = (MapNPC(mapnum, MapNpcNum).x + 1) Then
                                If GetPlayerY(i) = MapNPC(mapnum, MapNpcNum).y Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next i

                ' Check to make sure that there is not another npc in the way.
                For i = 1 To MAX_MAP_NPCS
                    If i <> MapNpcNum Then
                        If MapNPC(mapnum, i).num > 0 Then
                            If MapNPC(mapnum, i).x = (MapNPC(mapnum, MapNpcNum).x + 1) Then
                                If MapNPC(mapnum, i).y = MapNPC(mapnum, MapNpcNum).y Then
                                    CanNpcMove = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                Next i
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
            MapNPC(mapnum, MapNpcNum).y = MapNPC(mapnum, MapNpcNum).y - 1

        Case DIR_DOWN
            MapNPC(mapnum, MapNpcNum).y = MapNPC(mapnum, MapNpcNum).y + 1

        Case DIR_LEFT
            MapNPC(mapnum, MapNpcNum).x = MapNPC(mapnum, MapNpcNum).x - 1

        Case DIR_RIGHT
            MapNPC(mapnum, MapNpcNum).x = MapNPC(mapnum, MapNpcNum).x + 1
    End Select

    Call SendDataToMap(mapnum, "NPCMOVE" & SEP_CHAR & MapNpcNum & SEP_CHAR & MapNPC(mapnum, MapNpcNum).x & SEP_CHAR & MapNPC(mapnum, MapNpcNum).y & SEP_CHAR & MapNPC(mapnum, MapNpcNum).Dir & SEP_CHAR & Movement & END_CHAR)
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
Public Sub JoinGame(ByVal index As Long)
    Dim MOTD As String
    Dim FileData As String

    ' Set the flag so we know the person is in the game
    Player(index).InGame = True

    ' Send an ok to client to start receiving in game data
    Call SendDataTo(index, "loginok" & SEP_CHAR & index & END_CHAR)

    Call CheckEquippedItems(index)
    Call SendClasses(index)
    Call SendItems(index)
    Call SendEmoticons(index)
    Call SendElements(index)
    Call SendArrows(index)
    Call SendNpcs(index)
    Call SendShops(index)
    Call SendQuest(index)
    Call SendSpells(index)
    Call SendInventory(index)
    Call SendBank(index)
    Call SendHP(index)
    Call SendMP(index)
    Call SendSP(index)
    Call SendPTS(index)
    Call SendStats(index)
    Call SendWeatherTo(index)
    Call SendTimeTo(index)
    Call SendGameClockTo(index)
    Call DisabledTimeTo(index)
    Call SendSprite(index, index)
    Call SendPlayerSpells(index)
    Call SendOnlineList
    Call SendPlayerQuestFlags(index)

    ' Warp the player to his saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))

    If scripting = 1 Then
        MyScript.ExecuteStatement "Scripts\Main.txt", "JoinGame " & index
    Else
        ' Send a global message that he/she joined.
        If GetPlayerAccess(index) = 0 Then
            Call GlobalMsg(GetPlayerName(index) & " ha entrado en " & GAME_NAME & "!", 7)
        Else
            Call GlobalMsg(GetPlayerName(index) & " ha entrado en " & GAME_NAME & "!", 15)
        End If

        Call PlayerMsg(index, "Bienvenido a " & GAME_NAME & "!", 15)

        ' Send the player the welcome message.
        MOTD = Trim$(GetVar(App.Path & "\MOTD.ini", "MOTD", "Msg"))
        If LenB(MOTD) <> 0 Then
            Call PlayerMsg(index, "MOTD: " & MOTD, 11)
        End If

        ' Update all clients with the player.
        Call SendWhosOnline(index)
    End If

    ' Tell the client the player is in-game.
    Call SendDataTo(index, "ingame" & END_CHAR)

    ' Update the server console.
    Call ShowPLR(index)
    
    If IsPetAliveOnLogin(index) > 0 Then
       Call SpawnPet(index)
    End If
    
    FileData = ReadINI("SK1", "sid", App.Path & "\Scripts\db\" & GetPlayerName(index) & ".ini", vbNullString)
    Call SendDataTo(index, "getspell1" & SEP_CHAR & FileData & END_CHAR)
    FileData = ReadINI("SK2", "sid", App.Path & "\Scripts\db\" & GetPlayerName(index) & ".ini", vbNullString)
    Call SendDataTo(index, "getspell2" & SEP_CHAR & FileData & END_CHAR)
    FileData = ReadINI("SK3", "sid", App.Path & "\Scripts\db\" & GetPlayerName(index) & ".ini", vbNullString)
    Call SendDataTo(index, "getspell3" & SEP_CHAR & FileData & END_CHAR)
    FileData = ReadINI("SK4", "sid", App.Path & "\Scripts\db\" & GetPlayerName(index) & ".ini", vbNullString)
    Call SendDataTo(index, "getspell4" & SEP_CHAR & FileData & END_CHAR)
    FileData = ReadINI("SK5", "sid", App.Path & "\Scripts\db\" & GetPlayerName(index) & ".ini", vbNullString)
    Call SendDataTo(index, "getspell5" & SEP_CHAR & FileData & END_CHAR)
    FileData = ReadINI("SK6", "sid", App.Path & "\Scripts\db\" & GetPlayerName(index) & ".ini", vbNullString)
    Call SendDataTo(index, "getspell6" & SEP_CHAR & FileData & END_CHAR)
    FileData = ReadINI("SK7", "sid", App.Path & "\Scripts\db\" & GetPlayerName(index) & ".ini", vbNullString)
    Call SendDataTo(index, "getspell7" & SEP_CHAR & FileData & END_CHAR)
    FileData = ReadINI("SK8", "sid", App.Path & "\Scripts\db\" & GetPlayerName(index) & ".ini", vbNullString)
    Call SendDataTo(index, "getspell8" & SEP_CHAR & FileData & END_CHAR)
    FileData = ReadINI("SK9", "sid", App.Path & "\Scripts\db\" & GetPlayerName(index) & ".ini", vbNullString)
    Call SendDataTo(index, "getspell9" & SEP_CHAR & FileData & END_CHAR)
    FileData = ReadINI("SK10", "sid", App.Path & "\Scripts\db\" & GetPlayerName(index) & ".ini", vbNullString)
    Call SendDataTo(index, "getspell10" & SEP_CHAR & FileData & END_CHAR)
End Sub

Public Sub LeftGame(ByVal index As Long)
    Dim N As Long

    If Player(index).InGame Then
        Player(index).InGame = False
        If GetPlayerParty(index) > 0 Then Call PartyRemoval(index, GetPlayerParty(index), Trim$(GetPlayerName(index)))

        ' Stop processing NPCs if no one is on it.
        If GetTotalMapPlayers(GetPlayerMap(index)) = 0 Then
            PlayersOnMap(GetPlayerMap(index)) = NO
        End If
        
                ' Check if the player was in a party, and if so cancel it out so the other player doesn't continue to get half exp
        If Player(index).InParty = YES Then
            N = Player(index).PartyPlayer
            
            Call PlayerMsg(N, GetPlayerName(index) & " ha salido de " & GAME_NAME & ", deshaciendose el grupo actual.", PINK)
            Player(N).InParty = NO
            Player(N).PartyPlayer = 0
        End If

        
        If Player(index).Pet.Alive = YES Then
          ' Call TakeFromGrid(Player(Index).Pet.Map, Player(Index).Pet.x, Player(Index).Pet.y)
           Call savepet(index)
        End If
        

        If scripting = 1 Then
            MyScript.ExecuteStatement "Scripts\Main.txt", "LeftGame " & index
        Else
            ' Check to see if there is any boot map data.
            If Map(GetPlayerMap(index)).BootMap > 0 Then
                Call SetPlayerX(index, Map(GetPlayerMap(index)).BootX)
                Call SetPlayerY(index, Map(GetPlayerMap(index)).BootY)
                Call SetPlayerMap(index, Map(GetPlayerMap(index)).BootMap)
            End If

            ' Inform the server that the player logged off.
            If GetPlayerAccess(index) = 0 Then
                Call GlobalMsg(GetPlayerName(index) & " ha salido de " & GAME_NAME & "!", 7)
            Else
                Call GlobalMsg(GetPlayerName(index) & " ha salido de " & GAME_NAME & "!", 15)
            End If
        End If

        Call SavePlayer(index)
        Call SendLeftGame(index)

        Call TextAdd(frmServer.txtText(0), GetPlayerName(index) & " se ha desconectado de " & GAME_NAME & ".", True)

        Call RemovePLR(index)
    End If

    Call ClearPlayer(index)
    Call SendOnlineList
End Sub

Function GetTotalMapPlayers(ByVal mapnum As Long) As Long
    Dim i As Long

    If mapnum < 1 Or mapnum > MAX_MAPS Then
        Exit Function
    End If

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = mapnum Then
                GetTotalMapPlayers = GetTotalMapPlayers + 1
            End If
        End If
    Next i
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

Function GetPlayerHPRegen(ByVal index As Long) As Integer
    Dim Total As Integer

    If HP_REGEN = 1 Then
        If index < 1 Or index > MAX_PLAYERS Then
            Exit Function
        End If

        If Not IsPlaying(index) Then
            Exit Function
        End If

        Total = Int(GetPlayerDEF(index) / 2)
        If Total < 2 Then
            Total = 2
        End If

        GetPlayerHPRegen = Total
    End If
End Function

Function GetPlayerMPRegen(ByVal index As Long) As Integer
    Dim Total As Integer

    If MP_REGEN = 1 Then
        If index < 1 Or index > MAX_PLAYERS Then
            Exit Function
        End If

        If Not IsPlaying(index) Then
            Exit Function
        End If

        Total = Int(GetPlayerMAGI(index) / 2)
        If Total < 2 Then
            Total = 2
        End If

        GetPlayerMPRegen = Total
    End If
End Function

Function GetPlayerSPRegen(ByVal index As Long) As Integer
    Dim Total As Integer

    If SP_REGEN = 1 Then
        If index < 1 Or index > MAX_PLAYERS Then
            Exit Function
        End If

        If Not IsPlaying(index) Then
            Exit Function
        End If

        Total = Int(GetPlayerSPEED(index) / 2)
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

Sub CheckPlayerLevelUp(ByVal index As Long)
    Dim i As Long
    Dim d As Long
    Dim c As Long
    c = 0

    If GetPlayerExp(index) >= GetPlayerNextLevel(index) Then
        If GetPlayerLevel(index) < MAX_LEVEL Then
            If scripting = 1 Then
                MyScript.ExecuteStatement "Scripts\Main.txt", "PlayerLevelUp " & index
            Else
                Do Until GetPlayerExp(index) < GetPlayerNextLevel(index)
                    DoEvents
                    If GetPlayerLevel(index) < MAX_LEVEL Then
                        If GetPlayerExp(index) >= GetPlayerNextLevel(index) Then
                            d = GetPlayerExp(index) - GetPlayerNextLevel(index)
                            Call SetPlayerLevel(index, GetPlayerLevel(index) + 1)
                            i = Int(GetPlayerSPEED(index) / 10)
                            If i < 1 Then
                                i = 1
                            End If
                            If i > 3 Then
                                i = 3
                            End If

                            Call SetPlayerPOINTS(index, GetPlayerPOINTS(index) + i)
                            Call SetPlayerExp(index, d)
                            c = c + 1
                        End If
                    End If
                Loop
                If c > 1 Then
                    Call GlobalMsg(GetPlayerName(index) & " ha ganado " & c & " niveles!", 6)
                Else
                    Call GlobalMsg(GetPlayerName(index) & " ha ganado un nivel!", 6)
                End If
                Call BattleMsg(index, "Tu tienes " & GetPlayerPOINTS(index) & " puntos de estado", 9, 0)
            End If
            Call SendDataToMap(GetPlayerMap(index), "levelup" & SEP_CHAR & index & END_CHAR)
            Call SendPlayerLevelToAll(index)
        End If

        If GetPlayerLevel(index) = MAX_LEVEL Then
            Call SetPlayerExp(index, Experience(MAX_LEVEL))
        End If
    End If

    Call SendHP(index)
    Call SendMP(index)
    Call SendSP(index)
    Call SendPTS(index)

    Player(index).Char(Player(index).CharNum).MAXHP = GetPlayerMaxHP(index)
    Player(index).Char(Player(index).CharNum).MAXMP = GetPlayerMaxMP(index)
    Player(index).Char(Player(index).CharNum).MAXSP = GetPlayerMaxSP(index)

    Call SendStats(index)
End Sub

Sub CastSpell(ByVal index As Long, ByVal SpellSlot As Long)
    Dim spellnum As Long, i As Long, N As Long, Damage As Long
    Dim Casted As Boolean
    Casted = False
        

    ' Prevent player from using spells if they have been script locked
    If Player(index).LockedSpells = True Then
        Exit Sub
    End If

    ' Prevent subscript out of range
    If SpellSlot <= 0 Or SpellSlot > MAX_PLAYER_SPELLS Then
        Exit Sub
    End If

    spellnum = GetPlayerSpell(index, SpellSlot)

    ' Make sure player has the spell
    If Not HasSpell(index, spellnum) Then
        Call BattleMsg(index, "Tu no tienes este hechizo!", BRIGHTRED, 0)
        Exit Sub
    End If

    i = GetSpellReqLevel(spellnum)

    ' Check if they have enough MP
    If GetPlayerMP(index) < Spell(spellnum).MPCost Then
        Call BattleMsg(index, "No tienes suficiente mana!", BRIGHTRED, 0)
        Exit Sub
    End If

    ' Make sure they are the right level
    If i > GetPlayerLevel(index) Then
        Call BattleMsg(index, "Necesitas ser mayor del nivel " & i & " para realizar este hechizo.", BRIGHTRED, 0)
        Exit Sub
    End If

    ' Check if timer is ok
    If GetTickCount < Player(index).AttackTimer + Spell(spellnum).TimeToCast * 1000 Then
        Exit Sub
    End If

    ' Check if the spell is scripted and do that instead of a stat modification
    If Spell(spellnum).Type = SPELL_TYPE_SCRIPTED Then

        MyScript.ExecuteStatement "Scripts\Main.txt", "ScriptedSpell " & index & "," & Spell(spellnum).Data1

        Exit Sub
    End If
' End If

    Dim x As Long, y As Long

    If Spell(spellnum).AE = 1 Then
        For y = GetPlayerY(index) - Spell(spellnum).Range To GetPlayerY(index) + Spell(spellnum).Range
            For x = GetPlayerX(index) - Spell(spellnum).Range To GetPlayerX(index) + Spell(spellnum).Range
                N = -1
                For i = 1 To MAX_PLAYERS
                    If IsPlaying(i) = True Then
                        If GetPlayerMap(index) = GetPlayerMap(i) Then
                            If GetPlayerX(i) = x And GetPlayerY(i) = y Then
                                If i = index Then
                                    If Spell(spellnum).Type = SPELL_TYPE_ADDHP Or Spell(spellnum).Type = SPELL_TYPE_ADDMP Or Spell(spellnum).Type = SPELL_TYPE_ADDSP Then
                                        Player(index).Target = i
                                        Player(index).TargetType = TARGET_TYPE_PLAYER
                                        N = Player(index).Target
                                    End If
                                Else
                                    Player(index).Target = i
                                    Player(index).TargetType = TARGET_TYPE_PLAYER
                                    N = Player(index).Target
                                End If
                            End If
                        End If
                    End If
                Next i

                For i = 1 To MAX_MAP_NPCS
                    If MapNPC(GetPlayerMap(index), i).num > 0 Then
                        If NPC(MapNPC(GetPlayerMap(index), i).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(MapNPC(GetPlayerMap(index), i).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                            If MapNPC(GetPlayerMap(index), i).x = x And MapNPC(GetPlayerMap(index), i).y = y Then
                                Player(index).Target = i
                                Player(index).TargetType = TARGET_TYPE_NPC
                                N = Player(index).Target
                            End If
                        End If
                    End If
                Next i

                Casted = False
                If N > 0 Then
                    If Player(index).TargetType = TARGET_TYPE_PLAYER Then
                        If IsPlaying(N) Then
                            If N = index Then
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
                                Call PlayerMsg(index, "No puedes realizar el hechizo!", BRIGHTRED)
                            End If
                            If N <> index Then
                                Player(index).TargetType = TARGET_TYPE_PLAYER
                                If GetPlayerHP(N) > 0 And GetPlayerMap(index) = GetPlayerMap(N) And GetPlayerLevel(index) >= 10 And GetPlayerLevel(N) >= 10 And (Map(GetPlayerMap(index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(index) <= 0 And GetPlayerAccess(N) <= 0 Then
' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)

                                    Select Case Spell(spellnum).Type
                                        Case SPELL_TYPE_SUBHP

                                            Damage = (Int(GetPlayerMAGI(index) / 4) + Spell(spellnum).Data1) - GetPlayerProtection(N)
                                            If Damage > 0 Then
                                                Call AttackPlayer(index, N, Damage)
                                            Else
                                                Call BattleMsg(index, "El hechizo era muy debil para dañar a " & GetPlayerName(N) & "!", BRIGHTRED, 0)
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
                                    If GetPlayerMap(index) = GetPlayerMap(N) And Spell(spellnum).Type >= SPELL_TYPE_ADDHP And Spell(spellnum).Type <= SPELL_TYPE_ADDSP Then
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
                                        Call PlayerMsg(index, "No puedes realizar el hechizo!", BRIGHTRED)
                                    End If
                                End If
                            Else
                                Player(index).TargetType = TARGET_TYPE_PLAYER
                                If N = index Then
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
                                    Call PlayerMsg(index, "No puedes realizar el hechizo!", BRIGHTRED)
                                End If
                                If GetPlayerHP(N) > 0 And GetPlayerMap(index) = GetPlayerMap(N) And GetPlayerLevel(index) >= 10 And GetPlayerLevel(N) >= 10 And (Map(GetPlayerMap(index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(index) <= 0 And GetPlayerAccess(N) <= 0 Then
                                Else
                                    If GetPlayerMap(index) = GetPlayerMap(N) And Spell(spellnum).Type >= SPELL_TYPE_ADDHP And Spell(spellnum).Type <= SPELL_TYPE_ADDSP Then
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
                                        Call BattleMsg(index, "No puedes realizar el hechizo!", BRIGHTRED, 0)
                                    End If
                                End If
                            End If
                        Else
                            Call BattleMsg(index, "No puedes realizar el hechizo!", BRIGHTRED, 0)
                        End If
                    Else
                        Player(index).TargetType = TARGET_TYPE_NPC
                        If NPC(MapNPC(GetPlayerMap(index), N).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(MapNPC(GetPlayerMap(index), N).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(MapNPC(GetPlayerMap(index), N).num).Behavior <> NPC_BEHAVIOR_QUEST Then
                            If Spell(spellnum).Type >= SPELL_TYPE_SUBHP And Spell(spellnum).Type <= SPELL_TYPE_SUBSP Then
                                ' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on a " & Trim$(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & ".", BrightBlue)
                                Select Case Spell(spellnum).Type

                                    Case SPELL_TYPE_SUBHP
                                        Damage = (Int(GetPlayerMAGI(index) / 4) + Spell(spellnum).Data1) - Int(NPC(MapNPC(GetPlayerMap(index), N).num).DEF / 2)

                                        If Damage > 0 Then
                                            If Spell(spellnum).Element <> 0 And NPC(MapNPC(GetPlayerMap(index), N).num).Element <> 0 Then
                                                If Element(Spell(spellnum).Element).Strong = NPC(MapNPC(GetPlayerMap(index), N).num).Element Or Element(NPC(MapNPC(GetPlayerMap(index), N).num).Element).Weak = Spell(spellnum).Element Then
                                                    Call BattleMsg(index, "Una mezcla mortal de los elementos daña a " & Trim$(NPC(MapNPC(GetPlayerMap(index), N).num).Name) & "!", Blue, 0)
                                                    Damage = Int(Damage * 1.25)
                                                    If Element(Spell(spellnum).Element).Strong = NPC(MapNPC(GetPlayerMap(index), N).num).Element And Element(NPC(MapNPC(GetPlayerMap(index), N).num).Element).Weak = Spell(spellnum).Element Then
                                                        Damage = Int(Damage * 1.2)
                                                    End If
                                                End If

                                                If Element(Spell(spellnum).Element).Weak = NPC(MapNPC(GetPlayerMap(index), N).num).Element Or Element(NPC(MapNPC(GetPlayerMap(index), N).num).Element).Strong = Spell(spellnum).Element Then
                                                    Call BattleMsg(index, "" & Trim$(NPC(MapNPC(GetPlayerMap(index), N).num).Name) & " absorbe la fuerza elemental y se cura!", Red, 0)
                                                    Damage = Int(Damage * 0.75)
                                                    If Element(Spell(spellnum).Element).Weak = NPC(MapNPC(GetPlayerMap(index), N).num).Element And Element(NPC(MapNPC(GetPlayerMap(index), N).num).Element).Strong = Spell(spellnum).Element Then
                                                        Damage = Int(Damage * (2 / 3))
                                                    End If
                                                End If
                                            End If
                                            Call AttackNpc(index, N, Damage)
                                        Else
                                            Call BattleMsg(index, "El hechizo es muy debil para dañar a " & Trim$(NPC(MapNPC(GetPlayerMap(index), N).num).Name) & "!", BRIGHTRED, 0)
                                        End If

                                    Case SPELL_TYPE_SUBMP
                                        MapNPC(GetPlayerMap(index), N).MP = MapNPC(GetPlayerMap(index), N).MP - Spell(spellnum).Data1

                                    Case SPELL_TYPE_SUBSP
                                        MapNPC(GetPlayerMap(index), N).SP = MapNPC(GetPlayerMap(index), N).SP - Spell(spellnum).Data1
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
                            Call BattleMsg(index, "No puedes realizar el hechizo!", BRIGHTRED, 0)
                        End If
                    End If
                End If
                If Casted = True Then
                    Call SendDataToMap(GetPlayerMap(index), "spellanim" & SEP_CHAR & spellnum & SEP_CHAR & Spell(spellnum).SpellAnim & SEP_CHAR & Spell(spellnum).SpellTime & SEP_CHAR & Spell(spellnum).SpellDone & SEP_CHAR & index & SEP_CHAR & Player(index).TargetType & SEP_CHAR & Player(index).Target & SEP_CHAR & Player(index).CastedSpell & SEP_CHAR & Spell(spellnum).Big & END_CHAR)
                    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "magic" & SEP_CHAR & Spell(spellnum).Sound & END_CHAR)
                End If
            Next x
        Next y

        Call SetPlayerMP(index, GetPlayerMP(index) - Spell(spellnum).MPCost)
        Call SendMP(index)
    Else
        N = Player(index).Target
        If Player(index).TargetType = TARGET_TYPE_PLAYER Then
            If IsPlaying(N) Then
                If GetPlayerName(N) <> GetPlayerName(index) Then
                    If CInt(Sqr((GetPlayerX(index) - GetPlayerX(N)) ^ 2 + ((GetPlayerY(index) - GetPlayerY(N)) ^ 2))) > Spell(spellnum).Range Then
                        Call BattleMsg(index, "Estas muy alejado para dañar al objetivo.", BRIGHTRED, 0)
                        Exit Sub
                    End If
                End If
                Player(index).TargetType = TARGET_TYPE_PLAYER
                If GetPlayerHP(N) > 0 And GetPlayerMap(index) = GetPlayerMap(N) And GetPlayerLevel(index) >= 10 And GetPlayerLevel(N) >= 10 And (Map(GetPlayerMap(index)).Moral = MAP_MORAL_NONE Or Map(GetPlayerMap(index)).Moral = MAP_MORAL_NO_PENALTY) And GetPlayerAccess(index) <= 0 And GetPlayerAccess(N) <= 0 Then
' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on " & GetPlayerName(n) & ".", BrightBlue)

                    Select Case Spell(spellnum).Type
                        Case SPELL_TYPE_SUBHP

                            Damage = (Int(GetPlayerMAGI(index) / 4) + Spell(spellnum).Data1) - GetPlayerProtection(N)
                            If Damage > 0 Then
                                Call AttackPlayer(index, N, Damage)
                            Else
                                Call BattleMsg(index, "El hechizo es muy debil para dañar a " & GetPlayerName(N) & "!", BRIGHTRED, 0)
                            End If

                        Case SPELL_TYPE_SUBMP
                            Call SetPlayerMP(N, GetPlayerMP(N) - Spell(spellnum).Data1)
                            Call SendMP(N)

                        Case SPELL_TYPE_SUBSP
                            Call SetPlayerSP(N, GetPlayerSP(N) - Spell(spellnum).Data1)
                            Call SendSP(N)
                    End Select

                    ' Take away the mana points
                    Call SetPlayerMP(index, GetPlayerMP(index) - Spell(spellnum).MPCost)
                    Call SendMP(index)
                    Casted = True
                Else
                    If GetPlayerMap(index) = GetPlayerMap(N) And Spell(spellnum).Type >= SPELL_TYPE_ADDHP And Spell(spellnum).Type <= SPELL_TYPE_ADDSP Then
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
                        Call SetPlayerMP(index, GetPlayerMP(index) - Spell(spellnum).MPCost)
                        Call SendMP(index)
                        Casted = True
                    Else
                        Call BattleMsg(index, "No puedes realizar el hechizo!", BRIGHTRED, 0)
                    End If
                End If
            Else
                Call PlayerMsg(index, "No puedes realizar el hechizo!", BRIGHTRED)
            End If
        Else
                N = Player(index).TargetNPC
            If CInt(Sqr((GetPlayerX(index) - MapNPC(GetPlayerMap(index), N).x) ^ 2 + ((GetPlayerY(index) - MapNPC(GetPlayerMap(index), N).y) ^ 2))) > Spell(spellnum).Range Then
                Call BattleMsg(index, "Estás muy alejado para dañar a tu objetivo.", BRIGHTRED, 0)
                Exit Sub
            End If

            Player(index).TargetType = TARGET_TYPE_NPC

            If NPC(MapNPC(GetPlayerMap(index), N).num).Behavior <> NPC_BEHAVIOR_FRIENDLY And NPC(MapNPC(GetPlayerMap(index), N).num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER And NPC(MapNPC(GetPlayerMap(index), N).num).Behavior <> NPC_BEHAVIOR_QUEST Then
' Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " casts " & Trim$(Spell(SpellNum).Name) & " on a " & Trim$(Npc(MapNpc(GetPlayerMap(index), n).num).Name) & ".", BrightBlue)

                Select Case Spell(spellnum).Type
                    Case SPELL_TYPE_ADDHP
                        MapNPC(GetPlayerMap(index), N).HP = MapNPC(GetPlayerMap(index), N).HP + Spell(spellnum).Data1

                    Case SPELL_TYPE_SUBHP

                        Damage = (Int(GetPlayerMAGI(index) / 4) + Spell(spellnum).Data1) - Int(NPC(MapNPC(GetPlayerMap(index), N).num).DEF / 2)
                        If Damage > 0 Then
                            If Spell(spellnum).Element <> 0 And NPC(MapNPC(GetPlayerMap(index), N).num).Element <> 0 Then
                                If Element(Spell(spellnum).Element).Strong = NPC(MapNPC(GetPlayerMap(index), N).num).Element Or Element(NPC(MapNPC(GetPlayerMap(index), N).num).Element).Weak = Spell(spellnum).Element Then
                                    Call BattleMsg(index, "Una mezcla mortal de los elementos daña a " & Trim$(NPC(MapNPC(GetPlayerMap(index), N).num).Name) & "!", Blue, 0)
                                    Damage = Int(Damage * 1.25)
                                    If Element(Spell(spellnum).Element).Strong = NPC(MapNPC(GetPlayerMap(index), N).num).Element And Element(NPC(MapNPC(GetPlayerMap(index), N).num).Element).Weak = Spell(spellnum).Element Then
                                        Damage = Int(Damage * 1.2)
                                    End If
                                End If

                                If Element(Spell(spellnum).Element).Weak = NPC(MapNPC(GetPlayerMap(index), N).num).Element Or Element(NPC(MapNPC(GetPlayerMap(index), N).num).Element).Strong = Spell(spellnum).Element Then
                                    Call BattleMsg(index, "" & Trim$(NPC(MapNPC(GetPlayerMap(index), N).num).Name) & " absorbe gran cantidad elemental y se cura!", Red, 0)
                                    Damage = Int(Damage * 0.75)
                                    If Element(Spell(spellnum).Element).Weak = NPC(MapNPC(GetPlayerMap(index), N).num).Element And Element(NPC(MapNPC(GetPlayerMap(index), N).num).Element).Strong = Spell(spellnum).Element Then
                                        Damage = Int(Damage * (2 / 3))
                                    End If
                                End If
                            End If
                            Call AttackNpc(index, N, Damage)
                        Else
                            Call BattleMsg(index, "El hechizo es muy debil para dañar a " & Trim$(NPC(MapNPC(GetPlayerMap(index), N).num).Name) & "!", BRIGHTRED, 0)
                        End If

                    Case SPELL_TYPE_ADDMP
                        MapNPC(GetPlayerMap(index), N).MP = MapNPC(GetPlayerMap(index), N).MP + Spell(spellnum).Data1

                    Case SPELL_TYPE_SUBMP
                        MapNPC(GetPlayerMap(index), N).MP = MapNPC(GetPlayerMap(index), N).MP - Spell(spellnum).Data1

                    Case SPELL_TYPE_ADDSP
                        MapNPC(GetPlayerMap(index), N).SP = MapNPC(GetPlayerMap(index), N).SP + Spell(spellnum).Data1

                    Case SPELL_TYPE_SUBSP
                        MapNPC(GetPlayerMap(index), N).SP = MapNPC(GetPlayerMap(index), N).SP - Spell(spellnum).Data1
                End Select

                ' Take away the mana points
                Call SetPlayerMP(index, GetPlayerMP(index) - Spell(spellnum).MPCost)
                Call SendMP(index)
                Casted = True
            Else
                Call BattleMsg(index, "No puedes realizar el hechizo!", BRIGHTRED, 0)
            End If
        End If
    End If
' Solución para animación de hechizos
' Fixed #1 por Stream / Cambiado timer de hechizos
If Casted = True Then
        Player(index).AttackTimer = GetTickCount
        Player(index).CastedSpell = YES
        Call SendDataToMap(GetPlayerMap(index), "spellanim" & SEP_CHAR & spellnum & SEP_CHAR & Spell(spellnum).SpellAnim & SEP_CHAR & Spell(spellnum).SpellTime & SEP_CHAR & Spell(spellnum).SpellDone & SEP_CHAR & index & SEP_CHAR & Player(index).TargetType & SEP_CHAR & N & SEP_CHAR & Player(index).CastedSpell & SEP_CHAR & Spell(spellnum).Big & END_CHAR)
        Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "magic" & SEP_CHAR & Spell(spellnum).Sound & END_CHAR)
    End If
End Sub

Function CanPlayerCriticalHit(ByVal index As Long) As Boolean
    Dim i As Long
    Dim N As Long

    If GetPlayerWeaponSlot(index) > 0 Then
        N = Int(Rnd * 2)

        If N = 1 Then
            i = Int(GetPlayerSTR(index) / 2) + Int(GetPlayerLevel(index) / 2)

            N = Int(Rnd * 100) + 1
            If N <= i Then
                CanPlayerCriticalHit = True
            End If
        End If
    End If
End Function

Function CanPlayerBlockHit(ByVal index As Long) As Boolean
    Dim i As Long
    Dim N As Long

    If GetPlayerShieldSlot(index) > 0 Then
        N = Int(Rnd * 2)

        If N = 1 Then
            i = Int(GetPlayerDEF(index) / 2) + Int(GetPlayerLevel(index) / 2)

            N = Int(Rnd * 100) + 1
            If N <= i Then
                CanPlayerBlockHit = True
            End If
        End If
    End If
End Function

Public Sub CheckEquippedItems(ByVal index As Long)
    Dim ItemNum As Long

    ' Check to make sure the weapon exists.
    ItemNum = GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_WEAPON Then
            If Item(ItemNum).Type <> ITEM_TYPE_TWO_HAND Then
                Call SetPlayerWeaponSlot(index, 0)
            End If
        End If
    Else
        Call SetPlayerWeaponSlot(index, 0)
    End If

    ' Check to make sure the chest armor exists.
    ItemNum = GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_ARMOR Then
            Call SetPlayerArmorSlot(index, 0)
        End If
    Else
        Call SetPlayerArmorSlot(index, 0)
    End If

    ' Check to make sure the helmet exists.
    ItemNum = GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_HELMET Then
            Call SetPlayerHelmetSlot(index, 0)
        End If
    Else
        Call SetPlayerHelmetSlot(index, 0)
    End If

    ' Check to make sure the shield exists.
    ItemNum = GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_SHIELD Then
            Call SetPlayerShieldSlot(index, 0)
        End If
    Else
        Call SetPlayerShieldSlot(index, 0)
    End If

    ' Check to make sure the leggings exists.
    ItemNum = GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_LEGS Then
            Call SetPlayerLegsSlot(index, 0)
        End If
    Else
        Call SetPlayerLegsSlot(index, 0)
    End If

    ' Check to make sure the ring exists.
    ItemNum = GetPlayerInvItemNum(index, GetPlayerRingSlot(index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_RING Then
            Call SetPlayerRingSlot(index, 0)
        End If
    Else
        Call SetPlayerRingSlot(index, 0)
    End If

    ' Check to make sure the necklace exists.
    ItemNum = GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))
    If ItemNum > 0 Then
        If Item(ItemNum).Type <> ITEM_TYPE_NECKLACE Then
            Call SetPlayerNecklaceSlot(index, 0)
        End If
    Else
        Call SetPlayerNecklaceSlot(index, 0)
    End If
End Sub



' This sub-routine needs a re-write. [Mellowz]

Public Sub ShowPLR(ByVal index As Long)
    Dim LS As ListItem

    On Error Resume Next

    If frmServer.lvUsers.ListItems.Count > 0 And IsPlaying(index) Then
        frmServer.lvUsers.ListItems.Remove index
    End If

    Set LS = frmServer.lvUsers.ListItems.Add(index, , index)

    If IsPlaying(index) Then
        LS.SubItems(1) = GetPlayerLogin(index)
        LS.SubItems(2) = GetPlayerName(index)
        LS.SubItems(3) = GetPlayerLevel(index)
        LS.SubItems(4) = GetPlayerSprite(index)
        LS.SubItems(5) = GetPlayerAccess(index)
    End If
End Sub

Public Sub RemovePLR(ByVal index As Long)
    Dim LS As ListItem
    
    On Error Resume Next

    If Not IsPlaying(index) Then
        frmServer.lvUsers.ListItems.Remove index
    
        Set LS = frmServer.lvUsers.ListItems.Add(index, , index)
        
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

Sub SendIndexWornEquipment(ByVal index As Long)
    Dim Armor As Long
    Dim Helmet As Long
    Dim Shield As Long
    Dim Weapon As Long
    Dim Legs As Long
    Dim Ring As Long
    Dim Necklace As Long

    If GetPlayerArmorSlot(index) > 0 Then
        Armor = GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))
    End If

    If GetPlayerHelmetSlot(index) > 0 Then
        Helmet = GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))
    End If

    If GetPlayerShieldSlot(index) > 0 Then
        Shield = GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))
    End If

    If GetPlayerWeaponSlot(index) > 0 Then
        Weapon = GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))
    End If

    If GetPlayerLegsSlot(index) > 0 Then
        Legs = GetPlayerInvItemNum(index, GetPlayerLegsSlot(index))
    End If

    If GetPlayerRingSlot(index) > 0 Then
        Ring = GetPlayerInvItemNum(index, GetPlayerRingSlot(index))
    End If

    If GetPlayerNecklaceSlot(index) > 0 Then
        Necklace = GetPlayerInvItemNum(index, GetPlayerNecklaceSlot(index))
    End If

    Call SendDataToMap(GetPlayerMap(index), "itemworn" & SEP_CHAR & index & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & END_CHAR)
End Sub

Sub SendIndexWornEquipmentto(ByVal index As Long, ByVal From As Long)
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
        Shield = GetPlayerInvItemNum(From, GetPlayerShieldSlot(index))
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

    Call SendDataTo(index, "itemworn" & SEP_CHAR & From & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & END_CHAR)
End Sub


Sub SendIndexWornEquipmentFromMap(ByVal index As Long)
    Dim Packet As String
    Dim i As Long
    Dim Armor As Long
    Dim Helmet As Long
    Dim Shield As Long
    Dim Weapon As Long
    Dim Legs As Long
    Dim Ring As Long
    Dim Necklace As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) = True Then
            If GetPlayerMap(i) = GetPlayerMap(index) Then

                Armor = 0
                Helmet = 0
                Shield = 0
                Weapon = 0
                Legs = 0
                Ring = 0
                Necklace = 0

                If GetPlayerArmorSlot(i) > 0 Then
                    Armor = GetPlayerInvItemNum(i, GetPlayerArmorSlot(i))
                End If
                If GetPlayerHelmetSlot(i) > 0 Then
                    Helmet = GetPlayerInvItemNum(i, GetPlayerHelmetSlot(i))
                End If
                If GetPlayerShieldSlot(i) > 0 Then
                    Shield = GetPlayerInvItemNum(i, GetPlayerShieldSlot(i))
                End If
                If GetPlayerWeaponSlot(i) > 0 Then
                    Weapon = GetPlayerInvItemNum(i, GetPlayerWeaponSlot(i))
                End If
                If GetPlayerLegsSlot(i) > 0 Then
                    Legs = GetPlayerInvItemNum(i, GetPlayerLegsSlot(i))
                End If
                If GetPlayerRingSlot(i) > 0 Then
                    Ring = GetPlayerInvItemNum(i, GetPlayerRingSlot(i))
                End If
                If GetPlayerNecklaceSlot(i) > 0 Then
                    Necklace = GetPlayerInvItemNum(i, GetPlayerNecklaceSlot(i))
                End If

                Packet = "itemworn" & SEP_CHAR & i & SEP_CHAR & Armor & SEP_CHAR & Weapon & SEP_CHAR & Helmet & SEP_CHAR & Shield & SEP_CHAR & Legs & SEP_CHAR & Ring & SEP_CHAR & Necklace & END_CHAR
                Call SendDataTo(index, Packet)
            End If
        End If
    Next i
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
Sub ScriptSetTile(ByVal mapper As Long, ByVal x As Long, ByVal y As Long, ByVal setx As Long, ByVal sety As Long, ByVal tileset As Long, ByVal layer As Long)
    Dim Packet As String
    Packet = "tilecheck" & SEP_CHAR & mapper & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & layer & SEP_CHAR

    Select Case layer

        Case 0
            Map(mapper).Tile(x, y).Ground = sety * 14 + setx
            Map(mapper).Tile(x, y).GroundSet = tileset
            Packet = Packet & Map(mapper).Tile(x, y).Ground & SEP_CHAR & Map(mapper).Tile(x, y).GroundSet

        Case 1
            Map(mapper).Tile(x, y).Mask = sety * 14 + setx
            Map(mapper).Tile(x, y).MaskSet = tileset
            Packet = Packet & Map(mapper).Tile(x, y).Mask & SEP_CHAR & Map(mapper).Tile(x, y).MaskSet

        Case 2
            Map(mapper).Tile(x, y).Anim = sety * 14 + setx
            Map(mapper).Tile(x, y).AnimSet = tileset
            Packet = Packet & Map(mapper).Tile(x, y).Anim & SEP_CHAR & Map(mapper).Tile(x, y).AnimSet

        Case 3
            Map(mapper).Tile(x, y).Mask2 = sety * 14 + setx
            Map(mapper).Tile(x, y).Mask2Set = tileset
            Packet = Packet & Map(mapper).Tile(x, y).Mask2 & SEP_CHAR & Map(mapper).Tile(x, y).Mask2Set

        Case 4
            Map(mapper).Tile(x, y).M2Anim = sety * 14 + setx
            Map(mapper).Tile(x, y).M2AnimSet = tileset
            Packet = Packet & Map(mapper).Tile(x, y).M2Anim & SEP_CHAR & Map(mapper).Tile(x, y).M2AnimSet

        Case 5
            Map(mapper).Tile(x, y).Fringe = sety * 14 + setx
            Map(mapper).Tile(x, y).FringeSet = tileset
            Packet = Packet & Map(mapper).Tile(x, y).Fringe & SEP_CHAR & Map(mapper).Tile(x, y).FringeSet

        Case 6
            Map(mapper).Tile(x, y).FAnim = sety * 14 + setx
            Map(mapper).Tile(x, y).FAnimSet = tileset
            Packet = Packet & Map(mapper).Tile(x, y).FAnim & SEP_CHAR & Map(mapper).Tile(x, y).FAnimSet

        Case 7
            Map(mapper).Tile(x, y).Fringe2 = sety * 14 + setx
            Map(mapper).Tile(x, y).Fringe2Set = tileset
            Packet = Packet & Map(mapper).Tile(x, y).Fringe2 & SEP_CHAR & Map(mapper).Tile(x, y).Fringe2Set

        Case 8
            Map(mapper).Tile(x, y).F2Anim = sety * 14 + setx
            Map(mapper).Tile(x, y).F2AnimSet = tileset
            Packet = Packet & Map(mapper).Tile(x, y).F2Anim & SEP_CHAR & Map(mapper).Tile(x, y).F2AnimSet
    End Select

    Call SaveMap(mapper)
    Call SendDataToAll(Packet & END_CHAR)
End Sub

Sub ScriptSetAttribute(ByVal mapper As Long, ByVal x As Long, ByVal y As Long, ByVal Attrib As Long, ByVal Data1 As Long, ByVal Data2 As Long, ByVal Data3 As Long, ByVal String1 As String, ByVal String2 As String, ByVal String3 As String)
    Dim Packet As String
    
    With Map(mapper).Tile(x, y)
        .Type = Attrib
        .Data1 = Data1
        .Data2 = Data2
        .Data3 = Data3
        .String1 = String1
        .String2 = String2
        .String3 = String3
    End With

    Packet = "tilecheckattribute" & SEP_CHAR & mapper & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR
    With Map(mapper).Tile(x, y)
        Packet = Packet & .Type & SEP_CHAR & .Data1 & SEP_CHAR & .Data2 & SEP_CHAR & .Data3 & SEP_CHAR & .String1 & SEP_CHAR & .String2 & SEP_CHAR & .String3 & SEP_CHAR
    End With
    
    Call SaveMap(mapper)
    Call SendDataToAll(Packet & END_CHAR)
End Sub

Function ItemIsUsable(ByVal index As Long, ByVal InvNum As Long) As Boolean
    ' Check if the player meets the class requirement.
    If Item(GetPlayerInvItemNum(index, InvNum)).ClassReq > -1 Then
        If GetPlayerClass(index) <> Item(GetPlayerInvItemNum(index, InvNum)).ClassReq Then
            Call PlayerMsg(index, "Necesitas ser la clase " & GetClassName(Item(GetPlayerInvItemNum(index, InvNum)).ClassReq) & " para usar este objeto!", BRIGHTRED)
            Exit Function
        End If
    End If

    ' Check if the player meets the access requirement.
    If GetPlayerAccess(index) < Item(GetPlayerInvItemNum(index, InvNum)).AccessReq Then
        Call PlayerMsg(index, "Tu privilegio necesita ser mayor de " & Item(GetPlayerInvItemNum(index, InvNum)).AccessReq & "!", BRIGHTRED)
        Exit Function
    End If

    ' Check if the player meets the strength requirement.
    If GetPlayerSTR(index) < Item(GetPlayerInvItemNum(index, InvNum)).StrReq Then
        Call PlayerMsg(index, "No tienes la suficiente fuerza para equiparte esto!", BRIGHTRED)
        Exit Function
    End If

    ' Check if the player meets the defense requirement.
    If GetPlayerDEF(index) < Item(GetPlayerInvItemNum(index, InvNum)).DefReq Then
        Call PlayerMsg(index, "No tienes la suficiente defensa para equiparte esto!", BRIGHTRED)
        Exit Function
    End If

    ' Check if the player meets the magic requirement.
    If GetPlayerMAGI(index) < Item(GetPlayerInvItemNum(index, InvNum)).MagicReq Then
        Call PlayerMsg(index, "No tienes la suficiente magia para equiparte esto!", BRIGHTRED)
        Exit Function
    End If

    ' Check if the player meets the speed requirement.
    If GetPlayerSPEED(index) < Item(GetPlayerInvItemNum(index, InvNum)).SpeedReq Then
        Call PlayerMsg(index, "No tienes la suficiente velocidad para equiparte esto!", BRIGHTRED)
        Exit Function
    End If

    ItemIsUsable = True
End Function

Function ItemIsEquipped(ByVal index As Long, ByVal ItemNum As Long) As Boolean
    If GetPlayerWeaponSlot(index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerArmorSlot(index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerShieldSlot(index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerHelmetSlot(index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerLegsSlot(index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerRingSlot(index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If

    If GetPlayerNecklaceSlot(index) = ItemNum Then
        ItemIsEquipped = True
        Exit Function
    End If
End Function


Public Function DirToX(ByVal x As Long, _
   ByVal Dir As Byte) As Long
    DirToX = x

    If Dir = DIR_UP Or Dir = DIR_DOWN Then Exit Function

    ' LEFT = 2, RIGHT = 3
    ' 2 * 2 = 4, 4 - 5 = -1
    ' 3 * 2 = 6, 6 - 5 = 1
    DirToX = x + ((Dir * 2) - 5)
End Function

Public Function DirToY(ByVal y As Long, _
   ByVal Dir As Byte) As Long
    DirToY = y

    If Dir = DIR_LEFT Or Dir = DIR_RIGHT Then Exit Function

    ' UP = 0, DOWN = 1
    ' 0 * 2 = 0, 0 - 1 = -1
    ' 1 * 2 = 2, 2 - 1 = 1
    DirToY = y + ((Dir * 2) - 1)
End Function

Sub GMGiveItem(ByVal N As Long, ByVal i As Integer)
Dim q As Long

                q = 1
               
                Do While q < 25
                   If GetPlayerInvItemNum(N, q) = 0 Then
                        Call SetPlayerInvItemNum(N, q, i)
                        Call SetPlayerInvItemValue(N, q, 1)
                        Call SetPlayerInvItemDur(N, q, Item(i).Data1)
                        Call SendInventoryUpdate(N, q)
                        Call PlayerMsg(N, "Un objeto ha sido añadido en tu inventario por un administrador.", Green)
                        Call SendPlayerData(N)
                        q = 25
                    End If
                    q = q + 1
                Loop
        Exit Sub

End Sub

Sub GMTakeItem(ByVal N As Long, ByVal i As Integer)
Dim TakeText
Dim TakeTextAmount
Dim N1 As Long

   
    Call TakeItem(N, i, 1)
    Call PlayerMsg(N, "Un objeto de tu inventario ha sido borrado por un administrador.", Red)
    Call SendPlayerData(N)
    Call SendStats(N)
   
End Sub

Sub ClearParties()
Dim i, o As Long

    For i = 1 To MAX_PARTIES
        For o = 1 To MAX_PARTY_MEMBERS
            Party(i).Member(o) = 0
        Next
    Next
End Sub

