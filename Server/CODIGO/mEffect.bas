Attribute VB_Name = "mEffect"
Option Explicit

Public Sub Effect_Add(ByVal UserIndex As Integer, _
                      ByVal Slot As Byte, _
                      ByVal ObjIndex As Integer)
    
    On Error GoTo ErrHandler
    
    With UserList(UserIndex)

        If .flags.SlotEvent > 0 Then
            Call WriteConsoleMsg(UserIndex, "¡No puedes utilizar Scrolls en eventos automáticos!", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If ObjData(ObjIndex).Time = 0 Then
            Call Effect_Selected(UserIndex, ObjData(ObjIndex).BonusTipe, ObjData(ObjIndex).BonusValue, Slot)
        Else

            If .Counters.TimeBonus > 0 Then
                Call WriteConsoleMsg(UserIndex, "Ya tienes un efecto activo.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

                '
            End If
            
            .Stats.BonusTipe = ObjData(ObjIndex).BonusTipe
            .Stats.BonusValue = ObjData(ObjIndex).BonusValue
            .Counters.TimeBonus = ObjData(ObjIndex).Time
            
            Call WriteConsoleMsg(UserIndex, "Tendrás el efecto elegido durante " & Int(.Counters.TimeBonus / 60) & " minutos.", FontTypeNames.FONTTYPE_INFOGREEN)
            
            'Quitamos del inv el item
            If ObjData(ObjIndex).RemoveObj > 0 Then
                Call QuitarUserInvItem(UserIndex, Slot, ObjData(ObjIndex).RemoveObj)
                Call UpdateUserInv(False, UserIndex, Slot)

            End If

        End If
        
    End With

    Exit Sub

ErrHandler:

End Sub

Private Sub Effect_Selected(ByVal UserIndex As Integer, _
                            ByVal Tipe As eEffectObj, _
                            ByVal Value As Single, _
                            ByVal Slot As Byte)

    '<EhHeader>
    On Error GoTo Effect_Selected_Err

    '</EhHeader>
                            
    With UserList(UserIndex)

        Select Case Tipe

            Case eEffectObj.e_Exp
                .Stats.Exp = .Stats.Exp + Value
                Call CheckUserLevel(UserIndex)
                Call WriteUpdateExp(UserIndex)
                Call WriteConsoleMsg(UserIndex, "¡Has ganado " & Value & " puntos de experiencia!", FontTypeNames.FONTTYPE_INFOGREEN)
                
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call UpdateUserInv(False, UserIndex, Slot)
            
            Case eEffectObj.e_Gld
                      
                .Stats.Gld = .Stats.Gld + Value

                If (.Stats.Gld) > MAXORO Then .Stats.Gld = MAXORO
                Call WriteUpdateGold(UserIndex)
                Call WriteConsoleMsg(UserIndex, "¡Has ganado " & Value & " Monedas de Oro!", FontTypeNames.FONTTYPE_INFOGREEN)
                
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call UpdateUserInv(False, UserIndex, Slot)
                
            Case eEffectObj.e_Revive
                Call RevivirUsuario(UserIndex)
                .Stats.MinHam = 0
                .Stats.MinAGU = 0
                .flags.Hambre = 1
                .flags.Sed = 1
                Call WriteUpdateHungerAndThirst(UserIndex)
                Call WriteConsoleMsg(UserIndex, "¡Has vuelvo al mundo! ¡Quedas sediento!", FontTypeNames.FONTTYPE_INFOGREEN)
                
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call UpdateUserInv(False, UserIndex, Slot)
                
            Case eEffectObj.e_NewHead, eEffectObj.e_NewHeadClassic

                Dim TempHead As Integer

                TempHead = .Char.Head
                      
                Call User_GenerateNewHead(UserIndex, Tipe)
                
                If TempHead <> .Char.Head Then
                    Call QuitarUserInvItem(UserIndex, Slot, 1)
                    Call UpdateUserInv(False, UserIndex, Slot)

                End If
                    
                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
                Call SaveUser(UserList(UserIndex), CharPath & UCase$(.Name) & ".chr", False)
                    
            Case eEffectObj.e_ChangeGenero
                        
                If .Genero = Hombre Then
                    .Genero = Mujer
                Else
                    .Genero = Hombre

                End If
                                                
                If .Invent.ArmourEqpObjIndex > 0 Then
                    If Not SexoPuedeUsarItem(UserIndex, .Invent.ArmourEqpObjIndex) Then
                        Call Desequipar(UserIndex, .Invent.ArmourEqpSlot)

                    End If

                Else
                    Call DarCuerpoDesnudo(UserIndex)

                End If
                        
                Call User_GenerateNewHead(UserIndex, eEffectObj.e_NewHead)
                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
                Call SaveUser(UserList(UserIndex), CharPath & UCase$(.Name) & ".chr", False)
                Call WriteConsoleMsg(UserIndex, "¡Has cambiado tu género!", FontTypeNames.FONTTYPE_INFOGREEN)
                        
                Call QuitarUserInvItem(UserIndex, Slot, 1)
                Call UpdateUserInv(False, UserIndex, Slot)

        End Select
    
    End With

    '<EhFooter>
    Exit Sub

Effect_Selected_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mEffect.Effect_Selected " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub Effect_Remove(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo Effect_Remove_Err

    '</EhHeader>

    With UserList(UserIndex)
        .Stats.BonusTipe = 0
        .Stats.BonusValue = 0
        .Counters.TimeBonus = 0

    End With
    
    '<EhFooter>
    Exit Sub

Effect_Remove_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mEffect.Effect_Remove " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub Effect_Loop(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo Effect_Loop_Err

    '</EhHeader>
    
    With UserList(UserIndex)
        
        ' // NUEVO
        If .Pos.Map = 0 Then Exit Sub
        If MapInfo(.Pos.Map).Pk = False Then Exit Sub
        
        If .Counters.TimeBonus > 0 Then
        
            .Counters.TimeBonus = .Counters.TimeBonus - 1

            If .Counters.TimeBonus = 0 Then
                Effect_Remove (UserIndex)
                Call WriteConsoleMsg(UserIndex, "El efecto se ha ido.", FontTypeNames.FONTTYPE_INFORED)

            End If

        End If
    
    End With
    
    '<EhFooter>
    Exit Sub

Effect_Loop_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mEffect.Effect_Loop " & "at line " & Erl
        
    '</EhFooter>
End Sub

