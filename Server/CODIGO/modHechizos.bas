Attribute VB_Name = "modHechizos"
'Argentum Online 0.12.2
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public Const SUPERANILLO As Integer = 700

Public Sub ChangeSlotSpell(ByVal UserIndex As Integer, _
                           ByVal SlotOld As Byte, _
                           ByVal SlotNew As Byte)

    '<EhHeader>
    On Error GoTo ChangeSlotSpell_Err

    '</EhHeader>
    
    With UserList(UserIndex)

        Dim TempHechizo As Integer
        
        If SlotOld <= 0 Or SlotOld > MAXUSERHECHIZOS Then
            Call Logs_Security(eSecurity, eAntiHack, "El personaje " & UserList(UserIndex).Name & " ha intentado hackear el ChangeSlotSpell")
            Exit Sub

        End If
        
        If SlotNew <= 0 Or SlotNew > MAXUSERHECHIZOS Then
            Call Logs_Security(eSecurity, eAntiHack, "El personaje " & UserList(UserIndex).Name & " ha intentado hackear el ChangeSlotSpell")
            Exit Sub

        End If
        
        TempHechizo = .Stats.UserHechizos(SlotOld)
        .Stats.UserHechizos(SlotOld) = .Stats.UserHechizos(SlotNew)
        .Stats.UserHechizos(SlotNew) = TempHechizo

    End With

    '<EhFooter>
    Exit Sub

ChangeSlotSpell_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.ChangeSlotSpell " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub NpcLanzaSpellSobreTerreno(ByVal NpcIndex As Integer, _
                              ByVal Map As Integer, _
                              ByVal X As Integer, _
                              ByVal Y As Integer)

    '<EhHeader>
    On Error GoTo NpcLanzaSpellSobreTerreno_Err

    '</EhHeader>
                                  
    If Not Intervalo_CriatureAttack(NpcIndex) Then Exit Sub
    If Npclist(NpcIndex).flags.LanzaSpells = 0 Then Exit Sub
    
    Dim TempX      As Integer

    Dim TempY      As Integer

    Dim UserIndex  As Integer

    Dim SpellIndex As Integer

    Dim Random     As Integer
    
    Dim FxBool     As Boolean

    Random = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    SpellIndex = Npclist(NpcIndex).Spells(Random)
    
    With Hechizos(SpellIndex)

        For TempX = X - .TileRange To X + .TileRange
            For TempY = Y - .TileRange To Y + .TileRange
    
                If InMapBounds(Map, TempX, TempY) Then
                    UserIndex = MapData(Map, TempX, TempY).UserIndex
                    
                    If UserIndex > 0 Then

                        ' ¡¡Agregan HP!!
                        If .SubeHP = 1 Then
                        
                            ' Update HP
                            UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MinHp + RandomNumber(.MinHp, .MaxHp)

                            If UserList(UserIndex).Stats.MinHp > UserList(UserIndex).Stats.MaxHp Then UserList(UserIndex).Stats.MinHp = UserList(UserIndex).Stats.MaxHp
                            
                            ' ¡¡Quitan HP!!
                        ElseIf .SubeHP = 2 Then
                            
                        End If

                    End If
                    
                    If RandomNumber(1, 10) <= 2 And Not FxBool Then
                        FxBool = True
                        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFXMap(TempX, TempY, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))

                    End If

                End If
    
            Next TempY
        Next TempX
        
        If Not FxBool Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFXMap(X, Y, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))

        End If

        ' Spell Wav
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayEffect(Hechizos(SpellIndex).WAV, X, Y))
     
        ' Spell Words
        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead(Hechizos(SpellIndex).PalabrasMagicas, Npclist(NpcIndex).Char.charindex, vbCyan))
    
    End With
                
    '<EhFooter>
    Exit Sub

NpcLanzaSpellSobreTerreno_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.NpcLanzaSpellSobreTerreno " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, _
                           ByVal UserIndex As Integer, _
                           ByVal Spell As Integer, _
                           Optional ByVal DecirPalabras As Boolean = False, _
                           Optional ByVal IgnoreVisibilityCheck As Boolean = False)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 11/11/2010
    '13/02/2009: ZaMa - Los npcs que tiren magias, no podran hacerlo en mapas donde no se permita usarla.
    '13/07/2010: ZaMa - Ahora no se contabiliza la muerte de un atacable.
    '21/09/2010: ZaMa - Amplio los tipos de hechizos que pueden lanzar los npcs.
    '21/09/2010: ZaMa - Permito que se ignore el chequeo de visibilidad (pueden atacar a invis u ocultos).
    '11/11/2010: ZaMa - No se envian los efectos del hechizo si no lo castea.
    '***************************************************
    '<EhHeader>
    On Error GoTo NpcLanzaSpellSobreUser_Err

    '</EhHeader>

    If Not Intervalo_CriatureAttack(NpcIndex) Then Exit Sub
    If Not IntervaloPuedeRecibirAtaqueCriature(UserIndex) Then Exit Sub
          
    If Not EsObjetivoValido(NpcIndex, UserIndex) Then Exit Sub
          
    With UserList(UserIndex)
    
        If .flags.Muerto = 1 Then Exit Sub
        If (.flags.Mimetizado = 1) And (MapInfo(.Pos.Map).Pk) Then Exit Sub ' // NUEVO
        If Power.UserIndex = UserIndex Then Exit Sub
        
        ' Doesn't consider if the user is hidden/invisible or not.
        If Not IgnoreVisibilityCheck Then
            If UserList(UserIndex).flags.Invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then Exit Sub

        End If
        
        ' Si no se peude usar magia en el mapa, no le deja hacerlo.
        If MapInfo(UserList(UserIndex).Pos.Map).MagiaSinEfecto > 0 Then Exit Sub

        Dim daño As Integer
    
        ' Heal HP
        If Hechizos(Spell).SubeHP = 1 Then
        
            Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
        
            daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
            
            .Stats.MinHp = .Stats.MinHp + daño

            If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
            
            Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, d_CurarSpell))
            Call WriteUpdateUserStats(UserIndex)
        
            ' Damage
        ElseIf Hechizos(Spell).SubeHP = 2 Then
            
            If .flags.Privilegios And PlayerType.User Then
            
                Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
            
                daño = RandomNumber(Hechizos(Spell).MinHp, Hechizos(Spell).MaxHp)
                daño = daño - (daño * .Stats.UserSkills(eSkill.Resistencia) / 2000)
                
                If .Invent.CascoEqpObjIndex > 0 Then
                    daño = daño - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)

                End If
                
                If .Invent.EscudoEqpObjIndex > 0 Then
                    daño = daño - RandomNumber(ObjData(.Invent.EscudoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.EscudoEqpObjIndex).DefensaMagicaMax)

                End If
                
                If .Invent.ArmourEqpObjIndex > 0 Then
                    daño = daño - RandomNumber(ObjData(.Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.ArmourEqpObjIndex).DefensaMagicaMax)

                End If
                
                If .Invent.AnilloEqpObjIndex > 0 Then
                    daño = daño - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)

                End If
                
                daño = daño - (daño * UserList(UserIndex).Stats.UserSkills(eSkill.Resistencia) / 2000)
                
                If daño < 0 Then daño = 0
            
                .Stats.MinHp = .Stats.MinHp - daño
                
                Call SubirSkill(UserIndex, eSkill.Resistencia, True)
                Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, d_DañoNpc))
                Call WriteUpdateUserStats(UserIndex)
                
                'Muere
                If .Stats.MinHp < 1 Then
                    .Stats.MinHp = 0

                    If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                        RestarCriminalidad (UserIndex)

                    End If
                    
                    Dim MasterIndex As Integer

                    MasterIndex = Npclist(NpcIndex).MaestroUser
                    
                    '[Barrin 1-12-03]
                    If MasterIndex > 0 Then
                        
                        ' No son frags los muertos atacables
                        If .flags.AtacablePor <> MasterIndex Then
                            'Store it!
                            ' Call Statistics.StoreFrag(MasterIndex, UserIndex)
                            
                            Call ContarMuerte(UserIndex, MasterIndex)

                        End If
                        
                        Call ActStats(UserIndex, MasterIndex)

                    End If

                    '[/Barrin]
                    
                    Call UserDie(UserIndex)
                    
                End If
            
            End If
            
        End If
        
        ' Paralisis/Inmobilize
        If Hechizos(Spell).Paraliza = 1 Or Hechizos(Spell).Inmoviliza = 1 Then
            
            If .flags.Paralizado = 0 Then
                
                Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
                
                If .Invent.AnilloEqpObjIndex = SUPERANILLO Then
                    Call WriteConsoleMsg(UserIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                    Exit Sub

                End If
                
                Dim Dividido As Byte
                
                If Hechizos(Spell).Inmoviliza = 1 Then
                    .flags.Inmovilizado = 1

                End If
                  
                .flags.Paralizado = 1
                
                If .Clase = eClass.Warrior Or .Clase = eClass.Hunter Then
                    .Counters.Paralisis = Int(IntervaloParalizado / 2)
                Else
                    .Counters.Paralisis = IntervaloParalizado

                End If
                
                If .Invent.ReliquiaSlot > 0 Then
                    If ObjData(.Invent.ReliquiaObjIndex).EffectUser.AfectaParalisis > 0 Then
                        .Counters.Paralisis = IntervaloParalizado / ObjData(.Invent.ReliquiaObjIndex).EffectUser.AfectaParalisis

                        If .Counters.Paralisis <= 0 Then .Counters.Paralisis = 0
                        
                        WriteConsoleMsg UserIndex, "Tu reliquia ha rechazado el efecto de la parálisis a solo " & Int(.Counters.Paralisis / 40) & " segundos.", FontTypeNames.FONTTYPE_INFO

                    End If

                End If
                
                Call WriteParalizeOK(UserIndex)
                
            End If
            
        End If
        
        ' Stupidity
        If Hechizos(Spell).Estupidez = 1 Then
             
            If .flags.Estupidez = 0 Then
            
                Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
            
                If .Invent.AnilloEqpObjIndex = SUPERANILLO Then
                    Call WriteConsoleMsg(UserIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                    Exit Sub

                End If
                  
                .flags.Estupidez = 1
                .Counters.Ceguera = IntervaloInvisible
                          
                Call WriteDumb(UserIndex)
                
            End If

        End If
        
        ' Blind
        If Hechizos(Spell).Ceguera = 1 Then
             
            If .flags.Ceguera = 0 Then
            
                Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
            
                If .Invent.AnilloEqpObjIndex = SUPERANILLO Then
                    Call WriteConsoleMsg(UserIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                    Exit Sub

                End If
                  
                .flags.Ceguera = 1
                .Counters.Ceguera = IntervaloInvisible
                          
                Call WriteBlind(UserIndex)
                
            End If

        End If
        
        ' Remove Invisibility/Hidden
        If Hechizos(Spell).RemueveInvisibilidadParcial = 1 Then
                 
            Call SendSpellEffects(UserIndex, NpcIndex, Spell, DecirPalabras)
                 
            'Sacamos el efecto de ocultarse
            If .flags.Oculto = 1 Then
                .Counters.TiempoOculto = 0
                .flags.Oculto = 0
                Call SetInvisible(UserIndex, .Char.charindex, False)
                Call WriteConsoleMsg(UserIndex, "¡Has sido detectado!", FontTypeNames.FONTTYPE_VENENO, eMessageType.Combate)
            Else
                'sino, solo lo "iniciamos" en la sacada de invisibilidad.
                Call WriteConsoleMsg(UserIndex, "Comienzas a hacerte visible.", FontTypeNames.FONTTYPE_VENENO, eMessageType.Combate)
                .Counters.Invisibilidad = IntervaloInvisible - 1

            End If
        
        End If
        
    End With
    
    '<EhFooter>
    Exit Sub

NpcLanzaSpellSobreUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.NpcLanzaSpellSobreUser " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub SendSpellEffects(ByVal UserIndex As Integer, _
                             ByVal NpcIndex As Integer, _
                             ByVal Spell As Integer, _
                             ByVal DecirPalabras As Boolean)

    '<EhHeader>
    On Error GoTo SendSpellEffects_Err

    '</EhHeader>

    '***************************************************
    'Author: ZaMa
    'Last Modification: 11/11/2010
    'Sends spell's wav, fx and mgic words to users.
    '***************************************************
    With UserList(UserIndex)
        ' Spell Wav
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(Hechizos(Spell).WAV, .Pos.X, .Pos.Y, .Char.charindex))
            
        ' Spell FX
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, Hechizos(Spell).FXgrh, Hechizos(Spell).loops))
    
        ' Spell Words
        If DecirPalabras Then
                  
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateDamage(Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y - 1, -2, d_AddMagicWord, Hechizos(Spell).PalabrasMagicas))

            'Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead(Hechizos(Spell).PalabrasMagicas, Npclist(NpcIndex).Char.CharIndex, vbCyan))
        End If

    End With

    '<EhFooter>
    Exit Sub

SendSpellEffects_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.SendSpellEffects " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, _
                                 ByVal TargetNPC As Integer, _
                                 ByVal SpellIndex As Integer, _
                                 Optional ByVal DecirPalabras As Boolean = False)

    '***************************************************
    'Author: Unknown
    'Last Modification: 21/09/2010
    '21/09/2010: ZaMa - Now npcs can cast a wider range of spells.
    '***************************************************
    '<EhHeader>
    On Error GoTo NpcLanzaSpellSobreNpc_Err

    '</EhHeader>

    If Not Intervalo_CriatureAttack(NpcIndex) Then Exit Sub
    
    Dim Danio As Integer
    
    With Npclist(TargetNPC)
        
        ' Spell deals damage??
        If Hechizos(SpellIndex).SubeHP = 2 Then
            
            Danio = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
            
            If Npclist(NpcIndex).MaestroUser > 0 Then
                Call CalcularDarExp(Npclist(NpcIndex).MaestroUser, TargetNPC, Danio)
                Call Quests_AddNpc(Npclist(NpcIndex).MaestroUser, TargetNPC, Danio)

            End If
        
            ' Deal damage
            .Stats.MinHp = .Stats.MinHp - Danio
            
            'Muere?
            If .Stats.MinHp < 1 Then
                .Stats.MinHp = 0

                If Npclist(NpcIndex).MaestroUser > 0 Then
                    Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
                Else
                    Call MuereNpc(TargetNPC, 0)

                End If

            End If
            
            ' Spell recovers health??
        ElseIf Hechizos(SpellIndex).SubeHP = 1 Then

            If .Stats.MinHp = .Stats.MaxHp Then Exit Sub
                
            Danio = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
                
            ' Recovers health
            .Stats.MinHp = .Stats.MinHp + Danio
            
            If .Stats.MinHp > .Stats.MaxHp Then
                .Stats.MinHp = .Stats.MaxHp

            End If
            
        End If
        
        ' Spell Adds/Removes poison?
        If Hechizos(SpellIndex).Envenena = 1 Then
            .flags.Envenenado = 1
        ElseIf Hechizos(SpellIndex).CuraVeneno = 1 Then
            .flags.Envenenado = 0

        End If

        ' Spells Adds/Removes Paralisis/Inmobility?
        If Hechizos(SpellIndex).Paraliza = 1 Then
            .flags.Paralizado = 1
            .flags.Inmovilizado = 1
            .Contadores.Paralisis = IntervaloParalizado
            
        ElseIf Hechizos(SpellIndex).Inmoviliza = 1 Then
            .flags.Inmovilizado = 1
            .flags.Paralizado = 0
            .Contadores.Paralisis = IntervaloParalizado
            
        ElseIf Hechizos(SpellIndex).RemoverParalisis = 1 Then

            If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
                .flags.Paralizado = 0
                .flags.Inmovilizado = 0
                .Contadores.Paralisis = 0

            End If

        End If
            
        ' Spell sound and FX
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayEffect(Hechizos(SpellIndex).WAV, .Pos.X, .Pos.Y, .Char.charindex))
            
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(.Char.charindex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
    
        ' Decir las palabras magicas?
        If DecirPalabras Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageChatOverHead(Hechizos(SpellIndex).PalabrasMagicas, Npclist(NpcIndex).Char.charindex, vbCyan))

        End If
    
    End With

    '<EhFooter>
    Exit Sub

NpcLanzaSpellSobreNpc_Err:
    LogError Err.description & vbCrLf & "in NpcLanzaSpellSobreNpc " & "at line " & Erl

    '</EhFooter>
End Sub

Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo TieneHechizo_Err

    '</EhHeader>
    
    Dim j As Integer

    For j = 1 To MAXUSERHECHIZOS

        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True

            Exit Function

        End If

    Next

    '<EhFooter>
    Exit Function

TieneHechizo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.TieneHechizo " & "at line " & Erl
        
    '</EhFooter>
End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, _
                   ByVal Slot As Integer, _
                   Optional ByVal HechizoIndex As Integer = 0)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo AgregarHechizo_Err

    '</EhHeader>

    Dim hIndex As Integer

    Dim j      As Integer

    With UserList(UserIndex)

        If HechizoIndex > 0 Then
            hIndex = HechizoIndex
        Else
            hIndex = ObjData(.Invent.Object(Slot).ObjIndex).HechizoIndex

        End If
    
        If Not TieneHechizo(hIndex, UserIndex) Then

            'Buscamos un slot vacio
            For j = 1 To MAXUSERHECHIZOS

                If .Stats.UserHechizos(j) = 0 Then Exit For
            Next j
            
            If .Stats.UserHechizos(j) <> 0 Then
                Call WriteConsoleMsg(UserIndex, "No tienes espacio para más hechizos.", FontTypeNames.FONTTYPE_INFO)
            Else
                .Stats.UserHechizos(j) = hIndex
                Call UpdateUserHechizos(False, UserIndex, CByte(j))
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "Ya tienes ese hechizo.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

    '<EhFooter>
    Exit Sub

AgregarHechizo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.AgregarHechizo " & "at line " & Erl
        
    '</EhFooter>
End Sub
            
Sub DecirPalabrasMagicas(ByVal SpellWords As String, ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 17/11/2009
    '25/07/2009: ZaMa - Invisible admins don't say any word when casting a spell
    '17/11/2009: ZaMa - Now the user become visible when casting a spell, if it is hidden
    '11/06/2011: CHOTS - Color de dialogos customizables
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        
        If .flags.AdminInvisible <> 1 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatPersonalizado(SpellWords, .Char.charindex, 5))
            'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y - 1, -2, d_AddMagicWord, SpellWords))
            
            ' Si estaba oculto, se vuelve visible
            If .flags.Oculto = 1 Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                
                If .flags.Invisible = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                    Call SetInvisible(UserIndex, .Char.charindex, False)

                End If

            End If

        End If

    End With
    
    Exit Sub
    
ErrHandler:
    Call LogError("Error en DecirPalabrasMagicas. Error: " & Err.number & " - " & Err.description)

End Sub

''
' Check if an user can cast a certain spell
'
' @param UserIndex Specifies reference to user
' @param HechizoIndex Specifies reference to spell
' @return   True if the user can cast the spell, otherwise returns false
Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo PuedeLanzar_Err

    '</EhHeader>

    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010
    'Last Modification By: ZaMa
    '06/11/09 - Corregida la bonificación de maná del mimetismo en el druida con flauta mágica equipada.
    '19/11/2009: ZaMa - Validacion de mana para el Invocar Mascotas
    '12/01/2010: ZaMa - Validacion de mana para hechizos lanzados por druida.
    '***************************************************
    Dim DruidManaBonus As Single

    With UserList(UserIndex)

        If .flags.Muerto Then
            Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos estando muerto.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

            Exit Function

        End If
            
        If Hechizos(HechizoIndex).NeedStaff > 0 Then
            If .Clase = eClass.Mage Then
                If .Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(.Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                        Call WriteConsoleMsg(UserIndex, "No posees un báculo lo suficientemente poderoso para poder lanzar el conjuro.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

                        Exit Function

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes lanzar este conjuro sin la ayuda de un báculo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

                    Exit Function

                End If

            End If

        End If
            
        If Hechizos(HechizoIndex).LvlMin > 0 Then
            If .Stats.Elv < Hechizos(HechizoIndex).LvlMin Then
                Call WriteConsoleMsg(UserIndex, "Necesitas ser Nivel " & Hechizos(HechizoIndex).LvlMin & " para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

                Exit Function
            
            End If
        
        End If
            
        If .Stats.UserSkills(eSkill.Magia) < Hechizos(HechizoIndex).MinSkill Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes puntos de magia para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

            Exit Function

        End If
        
        If .Stats.MinSta < Hechizos(HechizoIndex).StaRequerido Then
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(UserIndex, "Estás muy cansado para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
            Else
                Call WriteConsoleMsg(UserIndex, "Estás muy cansada para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

            End If

            Exit Function

        End If
        
        If .Stats.MinMan < Hechizos(HechizoIndex).ManaRequerido Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente maná.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If
            
        If .Stats.MinHp < Hechizos(HechizoIndex).HpRequerido Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente vida.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If
        
    End With
    
    PuedeLanzar = True
    '<EhFooter>
    Exit Function

PuedeLanzar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.PuedeLanzar " & "at line " & Erl
        
    '</EhFooter>
End Function

Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef B As Boolean)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo HechizoTerrenoEstado_Err

    '</EhHeader>

    Dim PosCasteadaX As Integer

    Dim PosCasteadaY As Integer

    Dim PosCasteadaM As Integer

    Dim H            As Integer

    Dim TempX        As Integer

    Dim TempY        As Integer

    With UserList(UserIndex)
        PosCasteadaX = .flags.TargetX
        PosCasteadaY = .flags.TargetY
        PosCasteadaM = .flags.TargetMap
        
        H = .flags.Hechizo
        
        If Hechizos(H).RemueveInvisibilidadParcial = 1 Then
            B = True

            For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
                For TempY = PosCasteadaY - 8 To PosCasteadaY + 8

                    If InMapBounds(PosCasteadaM, TempX, TempY) Then
                        If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then

                            'hay un user
                            If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.Invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.AdminInvisible = 0 Then
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.charindex, Hechizos(H).FXgrh, Hechizos(H).loops))

                            End If

                        End If

                    End If

                Next TempY
            Next TempX
        
            Call InfoHechizo(UserIndex)

        End If
        
        Dim daño As Long
                            
        If Hechizos(H).SanacionGlobalNpcs = 1 Then
            B = True

            For TempX = PosCasteadaX - 2 To PosCasteadaX + 2
                For TempY = PosCasteadaY - 2 To PosCasteadaY + 2

                    If InMapBounds(PosCasteadaM, TempX, TempY) Then
                        If MapData(PosCasteadaM, TempX, TempY).NpcIndex > 0 Then

                            Dim tNpc As Integer: tNpc = MapData(PosCasteadaM, TempX, TempY).NpcIndex
                            
                            daño = RandomNumber(Hechizos(H).MinHp, Hechizos(H).MaxHp)
                            Npclist(tNpc).Stats.MinHp = Npclist(tNpc).Stats.MinHp + daño
                                
                            If Npclist(tNpc).Stats.MinHp > Npclist(tNpc).Stats.MaxHp Then Npclist(tNpc).Stats.MinHp = Npclist(tNpc).Stats.MaxHp
                                    
                            Call SendData(SendTarget.ToNPCArea, tNpc, PrepareMessageCreateFX(Npclist(tNpc).Char.charindex, Hechizos(H).FXgrh, Hechizos(H).loops))
                            Call SendData(SendTarget.ToNPCArea, tNpc, PrepareMessageCreateDamage(Npclist(tNpc).Pos.X, Npclist(tNpc).Pos.Y, daño, d_CurarSpell))

                        End If

                    End If

                Next TempY
            Next TempX
        
            Call InfoHechizo(UserIndex)

        End If
        
        If Hechizos(H).SanacionGlobal = 1 Then
            If .GuildIndex = 0 Then Exit Sub
            B = True

            For TempX = PosCasteadaX - 2 To PosCasteadaX + 2
                For TempY = PosCasteadaY - 2 To PosCasteadaY + 2

                    If InMapBounds(PosCasteadaM, TempX, TempY) Then
                        If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then

                            Dim tUser As Integer: tUser = MapData(PosCasteadaM, TempX, TempY).UserIndex
                            
                            ' Curamos a nuestro propio CLAN.
                            If .GuildIndex = UserList(tUser).GuildIndex Then
                                daño = RandomNumber(Hechizos(H).MinHp, Hechizos(H).MaxHp)
                                UserList(tUser).Stats.MinHp = UserList(tUser).Stats.MinHp + daño
                                
                                If UserList(tUser).Stats.MinHp > UserList(tUser).Stats.MaxHp Then UserList(tUser).Stats.MinHp = UserList(tUser).Stats.MaxHp
                                    
                                Call WriteUpdateHP(tUser)
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(tUser).Char.charindex, Hechizos(H).FXgrh, Hechizos(H).loops))
                                Call SendData(SendTarget.ToPCArea, tUser, PrepareMessageCreateDamage(UserList(tUser).Pos.X, UserList(tUser).Pos.Y, daño, d_CurarSpell))
                                
                                ' Sanación más REMOVER PARÁLISIS.
                                If Hechizos(H).RemoverParalisis = 1 Then
                                    If UserList(tUser).flags.Paralizado = 1 Or UserList(tUser).flags.Inmovilizado = 1 Then
                                        Call RemoveParalisis(tUser)

                                    End If

                                End If

                            End If

                        End If

                    End If

                Next TempY
            Next TempX
        
            Call InfoHechizo(UserIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

HechizoTerrenoEstado_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.HechizoTerrenoEstado " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef HechizoCasteado As Boolean)
    '***************************************************
    'Author: Uknown
    'Last modification: 18/09/2010
    'Sale del sub si no hay una posición valida.
    '18/11/2009: Optimizacion de codigo.
    '18/09/2010: ZaMa - No se permite invocar en mapas con InvocarSinEfecto.
    '***************************************************

    On Error GoTo error

    With UserList(UserIndex)

        Dim mapa As Integer

        mapa = .Pos.Map
    
        'No permitimos se invoquen criaturas en zonas seguras
        If MapInfo(mapa).Pk = False Or MapData(mapa, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA Then
            Call WriteConsoleMsg(UserIndex, "No puedes invocar criaturas en zona segura.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
    
        'No permitimos se invoquen criaturas en mapas donde esta prohibido hacerlo
        If MapInfo(mapa).InvocarSinEfecto = 1 Then
            Call WriteConsoleMsg(UserIndex, "Invocar no está permitido aquí! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
        
        If .flags.SlotFast > 0 Then
            If RetoFast(.flags.SlotFast).ConfigVale <> ValeTodo Then
                Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite invocar criaturas.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Exit Sub
    
            End If

        End If
        
        Dim SlotEvent As Byte

        SlotEvent = .flags.SlotEvent

        If SlotEvent > 0 Then
            If Events(SlotEvent).config(eConfigEvent.eInvocar) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Invocar no está permitido aquí! Retirate de la Zona del Evento si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If
        
        If .MascotaIndex Then
            Call QuitarPet(UserIndex, .MascotaIndex)
            Exit Sub

        End If
            
        Dim SpellIndex As Integer, NroNpcs As Integer, NpcIndex As Integer, PetIndex As Integer

        Dim targetPos  As WorldPos
        
        Dim Entrenable As Boolean
    
        targetPos.Map = .flags.TargetMap
        targetPos.X = .flags.TargetX
        targetPos.Y = .flags.TargetY
    
        SpellIndex = .flags.Hechizo
        
        If MapData(targetPos.Map, targetPos.X, targetPos.Y).trigger = POSINVALIDA Or MapData(targetPos.Map, targetPos.X, targetPos.Y).TileExit.Map <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Elige una posición válida para realizar la invocación.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
         
        If Hechizos(SpellIndex).Warp = 1 Then
            PetIndex = Hechizos(SpellIndex).NumNpc

            ' Warp de Mascota
            Entrenable = True
            
        Else
            PetIndex = Hechizos(SpellIndex).NumNpc
            ' Invocación de fuego fatuo y demas
            
            ' If PetIndex = 791 Then
            '  Entrenable = True
            ' Else
            '   Entrenable = False
            ' End If
            
        End If
        
        NpcIndex = SpawnNpc(PetIndex, targetPos, False, False)
            
        If NpcIndex > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayEffect(Hechizos(SpellIndex).WAV, .Pos.X, .Pos.Y))
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFX(Npclist(NpcIndex).Char.charindex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
            
            .MascotaIndex = NpcIndex
                
            With Npclist(NpcIndex)
                .MaestroUser = UserIndex
                .Contadores.TiempoExistencia = IntervaloInvocacion
                .GiveGLD = 0
                .MenuIndex = eMenues.iemascota
                .Entrenable = Entrenable

            End With
                
            Call FollowAmo(NpcIndex)
        Else

            Exit Sub

        End If

    End With

    Call InfoHechizo(UserIndex)
    HechizoCasteado = True

    Exit Sub

error:

    With UserList(UserIndex)
        LogError ("[" & Err.number & "] " & Err.description & " por el usuario " & .Name & "(" & UserIndex & ") en (" & .Pos.Map & ", " & .Pos.X & ", " & .Pos.Y & "). Tratando de tirar el hechizo " & Hechizos(SpellIndex).Nombre & "(" & SpellIndex & ") en la posicion ( " & .flags.TargetX & ", " & .flags.TargetY & ")")

    End With

End Sub

''
' Le da propiedades al nuevo npc
'
' @param UserIndex  Indice del usuario que invoca.
' @param b  Indica si se termino la operación.

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal SpellIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 18/11/2009
    '18/11/2009: ZaMa - Optimizacion de codigo.
    '***************************************************
    '<EhHeader>
    On Error GoTo HandleHechizoTerreno_Err

    '</EhHeader>
    
    Dim HechizoCasteado As Boolean

    Dim ManaRequerida   As Integer
    
    Select Case Hechizos(SpellIndex).Tipo

        Case TipoHechizo.uInvocacion
            Call HechizoInvocacion(UserIndex, HechizoCasteado)
            
        Case TipoHechizo.uEstado
            Call HechizoTerrenoEstado(UserIndex, HechizoCasteado)

    End Select

    If HechizoCasteado Then

        With UserList(UserIndex)
            Call SubirSkill(UserIndex, eSkill.Magia, True)
            
            ManaRequerida = Hechizos(SpellIndex).ManaRequerido
            
            ' Bonificaciones en hechizos
            If .Clase = eClass.Druid Then

                ' Solo con flauta equipada
                If .Invent.MagicObjIndex = ANILLOMAGICO Then
                    ' 30% menos de mana para invocaciones
                    ManaRequerida = ManaRequerida * 0.7

                End If
                    
            ElseIf .Clase = eClass.Paladin Or .Clase = eClass.Assasin Then
                ' 25% menos de mana para invocaciones
                ManaRequerida = ManaRequerida * 0.75

            End If
            
            ' Quito la mana requerida
            .Stats.MinMan = .Stats.MinMan - ManaRequerida

            If .Stats.MinMan < 0 Then .Stats.MinMan = 0
            
            ' Quito la estamina requerida
            .Stats.MinSta = .Stats.MinSta - Hechizos(SpellIndex).StaRequerido

            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
            ' Update user stats
            Call WriteUpdateUserStats(UserIndex)

        End With

    End If
    
    '<EhFooter>
    Exit Sub

HandleHechizoTerreno_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.HandleHechizoTerreno " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal SpellIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010
    '18/11/2009: ZaMa - Optimizacion de codigo.
    '12/01/2010: ZaMa - Optimizacion y agrego bonificaciones al druida.
    '***************************************************
    '<EhHeader>
    On Error GoTo HandleHechizoUsuario_Err

    '</EhHeader>
    
    Dim HechizoCasteado As Boolean

    Dim ManaRequerida   As Integer
    
    Select Case Hechizos(SpellIndex).Tipo

        Case TipoHechizo.uEstado
            ' Afectan estados (por ejem : Envenenamiento)
            Call HechizoEstadoUsuario(UserIndex, HechizoCasteado)
        
        Case TipoHechizo.uPropiedades
            ' Afectan HP,MANA,STAMINA,ETC
            HechizoCasteado = HechizoPropUsuario(UserIndex)

    End Select

    If HechizoCasteado Then

        With UserList(UserIndex)
            Call SubirSkill(UserIndex, eSkill.Magia, True)
            
            ManaRequerida = Hechizos(SpellIndex).ManaRequerido
            
            ' Bonificaciones para druida
            If .Clase = eClass.Druid Then

                ' Solo con flauta magica
                If .Invent.MagicObjIndex = ANILLOMAGICO Then
                    If Hechizos(SpellIndex).Mimetiza = 1 Then
                        ' 50% menos de mana para mimetismo
                        ' ManaRequerida = ManaRequerida * 0.5
                        
                    ElseIf SpellIndex <> APOCALIPSIS_SPELL_INDEX Or SpellIndex <> DESCARGA_SPELL_INDEX Then

                        ' 10% menos de mana para todo menos apoca y descarga
                        'ManaRequerida = ManaRequerida * 0.9
                    End If

                End If
                    
            ElseIf .Clase = eClass.Paladin Or .Clase = eClass.Assasin Then
                ' 15% menos de mana  hechizos contra usuarios incluyéndose.
                ManaRequerida = ManaRequerida * 0.85

            End If
            
            ' Quito la mana requerida
            .Stats.MinMan = .Stats.MinMan - ManaRequerida

            If .Stats.MinMan < 0 Then .Stats.MinMan = 0
            
            ' Quito la estamina requerida
            .Stats.MinSta = .Stats.MinSta - Hechizos(SpellIndex).StaRequerido

            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
            ' Update user stats
            Call WriteUpdateUserStats(UserIndex)
            Call WriteUpdateUserStats(.flags.TargetUser)
            .flags.TargetUser = 0
            
        End With

    End If

    '<EhFooter>
    Exit Sub

HandleHechizoUsuario_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.HandleHechizoUsuario " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer)

    '<EhHeader>
    On Error GoTo HandleHechizoNPC_Err

    '</EhHeader>

    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010
    '13/02/2009: ZaMa - Agregada 50% bonificacion en coste de mana a mimetismo para druidas
    '17/11/2009: ZaMa - Optimizacion de codigo.
    '12/01/2010: ZaMa - Bonificacion para druidas de 10% para todos hechizos excepto apoca y descarga.
    '12/01/2010: ZaMa - Los druidas mimetizados con npcs ahora son ignorados.
    '***************************************************
    Dim HechizoCasteado As Boolean

    Dim ManaRequerida   As Long
    
    With UserList(UserIndex)

        If Npclist(.flags.TargetNPC).AntiMagia > 0 Then
            Call WriteConsoleMsg(UserIndex, "¡El efecto de Magia ha sido rechazado!", FontTypeNames.FONTTYPE_INFORED)
            Exit Sub

        End If
            
        Select Case Hechizos(HechizoIndex).Tipo

            Case TipoHechizo.uEstado
                ' Afectan estados (por ejem : Envenenamiento)
                Call HechizoEstadoNPC(.flags.TargetNPC, HechizoIndex, HechizoCasteado, UserIndex)
                
            Case TipoHechizo.uPropiedades
                ' Afectan HP,MANA,STAMINA,ETC
                Call HechizoPropNPC(HechizoIndex, .flags.TargetNPC, UserIndex, HechizoCasteado)

        End Select
        
        If HechizoCasteado Then
            Call SubirSkill(UserIndex, eSkill.Magia, True)
            
            ManaRequerida = Hechizos(HechizoIndex).ManaRequerido
            
            ' Bonificación para druidas.
            If .Clase = eClass.Druid Then
                ' Se mostró como usuario, puede ser atacado por npcs
                .flags.Ignorado = False
                
                ' Solo con flauta equipada
                If .Invent.MagicObjIndex = ANILLOMAGICO Then
                    If Hechizos(HechizoIndex).Mimetiza = 1 Then
                        ' 50% menos de mana para mimetismo
                        ManaRequerida = ManaRequerida * 0.5
                        ' Será ignorado hasta que pierda el efecto del mimetismo o ataque un npc
                        .flags.Ignorado = True
                    Else

                        ' 10% menos de mana para hechizos
                        If HechizoIndex <> APOCALIPSIS_SPELL_INDEX Or HechizoIndex <> DESCARGA_SPELL_INDEX Then

                            ' ManaRequerida = ManaRequerida * 0.9
                        End If

                    End If

                End If
                    
            ElseIf .Clase = eClass.Paladin Or .Clase = eClass.Assasin Then
                ManaRequerida = ManaRequerida * 0.85

            End If
            
            ' Quito la mana requerida
            .Stats.MinMan = .Stats.MinMan - ManaRequerida

            If .Stats.MinMan < 0 Then .Stats.MinMan = 0
            
            ' Quito la estamina requerida
            .Stats.MinSta = .Stats.MinSta - Hechizos(HechizoIndex).StaRequerido

            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
            ' Update user stats
            Call WriteUpdateUserStats(UserIndex)
            .flags.TargetNPC = 0

        End If

    End With

    '<EhFooter>
    Exit Sub

HandleHechizoNPC_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.HandleHechizoNPC " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub LanzarHechizo(ByVal SpellIndex As Integer, ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo LanzarHechizo_Err

    '</EhHeader>

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 02/16/2010
    '24/01/2007 ZaMa - Optimizacion de codigo.
    '02/16/2010: Marco - Now .flags.hechizo makes reference to global spell index instead of user's spell index
    '***************************************************

    With UserList(UserIndex)
    
        If .flags.EnConsulta Then
            Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos si estás en consulta.", FontTypeNames.FONTTYPE_INFO)

            Exit Sub

        End If
    
        If .flags.GmSeguidor > 0 Then

            Dim Temp As Long, TiempoActual As Long

            TiempoActual = GetTime
            Temp = TiempoActual - .interval(0).ISpell
                        
            Call WriteUpdateInfoIntervals(.flags.GmSeguidor, 3, Temp, .flags.MenuCliente)
            
            'If .flags.TargetUser > 0 Then
            'Call WriteUpdateInfoIntervals(.flags.GmSeguidor, 5, "Lanzó hechizo sobre " & UserList(.flags.TargetUser).Name, .flags.MenuCliente)
            'ElseIf .flags.TargetNPC > 0 Then
            'Call WriteUpdateInfoIntervals(.flags.GmSeguidor, 5, "Lanzó hechizo sobre " & Npclist(.flags.TargetNPC).Name, .flags.MenuCliente)
            'End If
            
            .interval(0).ISpell = TiempoActual

        End If
        
        'Chequeamos que no esté desnudo
        'If .flags.Desnudo Then
        'If .Genero = eGenero.Hombre Then
        'Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos si estás desnudo.", FontTypeNames.FONTTYPE_INFO)
        'Else
        'Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos si estás desnuda.", FontTypeNames.FONTTYPE_INFO)
        'End If
        'Exit Sub
        'End If
    
        If PuedeLanzar(UserIndex, SpellIndex) Then

            Select Case Hechizos(SpellIndex).Target

                Case TargetType.uUsuarios

                    If .flags.TargetUser > 0 Then
                        If Abs(UserList(.flags.TargetUser).Pos.Y - .Pos.Y) <= RANGO_VISION_y Then
                            Call HandleHechizoUsuario(UserIndex, SpellIndex)
                        Else
                            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING, eMessageType.Combate)

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "Este hechizo actúa sólo sobre usuarios.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

                        'Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageCreateDamage(.flags.TargetX, .flags.TargetY - 1, -1, eDamageType.d_Fallas, "Fallas"))
                    End If
            
                Case TargetType.uNPC

                    If .flags.TargetNPC > 0 Then
                        If Abs(Npclist(.flags.TargetNPC).Pos.Y - .Pos.Y) <= RANGO_VISION_y Then
                            Call HandleHechizoNPC(UserIndex, SpellIndex)
                        Else
                            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING, eMessageType.Combate)

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "Este hechizo sólo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

                        'Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageCreateDamage(.flags.TargetX, .flags.TargetY - 1, -1, eDamageType.d_Fallas, "Fallas"))
                    End If
            
                Case TargetType.uUsuariosYnpc

                    If .flags.TargetUser > 0 Then
                        If Abs(UserList(.flags.TargetUser).Pos.Y - .Pos.Y) <= RANGO_VISION_y Then
                            Call HandleHechizoUsuario(UserIndex, SpellIndex)
                        Else
                            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING, eMessageType.Combate)

                        End If

                    ElseIf .flags.TargetNPC > 0 Then

                        If Abs(Npclist(.flags.TargetNPC).Pos.Y - .Pos.Y) <= RANGO_VISION_y Then
                            Call HandleHechizoNPC(UserIndex, SpellIndex)
                        Else
                            Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING, eMessageType.Combate)

                        End If

                    Else
                        'Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageCreateDamage(.flags.TargetX, .flags.TargetY - 1, -1, eDamageType.d_Fallas, "Fallas"))
                        Call WriteConsoleMsg(UserIndex, "Target inválido.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

                    End If
            
                Case TargetType.uTerreno
                    Call HandleHechizoTerreno(UserIndex, SpellIndex)

                Case TargetType.uArea
                    Call HandleHechizoArea(UserIndex, SpellIndex)
                          
            End Select
        
        End If
    
        If .Counters.Trabajando Then .Counters.Trabajando = .Counters.Trabajando - 1
    
        If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1
       
    End With
    
    '<EhFooter>
    Exit Sub

LanzarHechizo_Err:
        
    LogError "Error en LanzarHechizo. Error " & Err.number & " : " & Err.description & " Hechizo: " & Hechizos(SpellIndex).Nombre & "(" & SpellIndex & "). Casteado por: " & UserList(UserIndex).Name & "(" & UserIndex & "). at line " & Erl
        
    '</EhFooter>
End Sub

Sub HandleHechizoArea(ByVal UserIndex As Integer, ByVal SpellIndex As Integer)

    On Error GoTo HandleHechizoArea_Err

    Dim HechizoCasteado As Boolean

    Dim ManaRequerida   As Integer
    
    Select Case Hechizos(SpellIndex).Tipo

        Case TipoHechizo.uInvocacion
            'Call HechizoInvocacion(UserIndex, HechizoCasteado)

        Case TipoHechizo.uPropiedades
            ' If esnpc then
            'HechizoCasteado = HechizoPropAreaNPC(UserIndex)
            ' else
            HechizoCasteado = HechizoPropAreaUsuario(UserIndex)
            ' end if
                   
    End Select

    If HechizoCasteado Then

        With UserList(UserIndex)
            Call SubirSkill(UserIndex, eSkill.Magia, True)
            
            ManaRequerida = Hechizos(SpellIndex).ManaRequerido
            
            ' Bonificaciones en hechizos
            If .Clase = eClass.Druid Then

                ' Solo con flauta equipada
                If .Invent.MagicObjIndex = ANILLOMAGICO Then
                    ' 30% menos de mana para invocaciones
                    ManaRequerida = ManaRequerida * 0.7

                End If
                    
            ElseIf .Clase = eClass.Paladin Or .Clase = eClass.Assasin Then
                ' 25% menos de mana para invocaciones
                ManaRequerida = ManaRequerida * 0.75

            End If
            
            ' Quito la mana requerida
            .Stats.MinMan = .Stats.MinMan - ManaRequerida

            If .Stats.MinMan < 0 Then .Stats.MinMan = 0
            
            ' Quito la estamina requerida
            .Stats.MinSta = .Stats.MinSta - Hechizos(SpellIndex).StaRequerido

            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
            ' Update user stats
            Call WriteUpdateUserStats(UserIndex)

        End With

    End If
    
    '<EhFooter>
    Exit Sub

HandleHechizoArea_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.HandleHechizoArea " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef HechizoCasteado As Boolean)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 28/04/2010
    'Handles the Spells that afect the Stats of an User
    '24/01/2007 Pablo (ToxicWaste) - Invisibilidad no permitida en Mapas con InviSinEfecto
    '26/01/2007 Pablo (ToxicWaste) - Cambios que permiten mejor manejo de ataques en los rings.
    '26/01/2007 Pablo (ToxicWaste) - Revivir no permitido en Mapas con ResuSinEfecto
    '02/01/2008 Marcos (ByVal) - Curar Veneno no permitido en usuarios muertos.
    '06/28/2008 NicoNZ - Agregué que se le de valor al flag Inmovilizado.
    '17/11/2008: NicoNZ - Agregado para quitar la penalización de vida en el ring y cambio de ecuacion.
    '13/02/2009: ZaMa - Arreglada ecuacion para quitar vida tras resucitar en rings.
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '28/04/2010: ZaMa - Agrego Restricciones para ciudas respecto al estado atacable.
    '16/09/2010: ZaMa - Solo se hace invi para los clientes si no esta navegando.
    '***************************************************
    '<EhHeader>
    On Error GoTo HechizoEstadoUsuario_Err

    '</EhHeader>

    Dim HechizoIndex As Integer

    Dim TargetIndex  As Integer

    With UserList(UserIndex)
        HechizoIndex = .flags.Hechizo
        TargetIndex = .flags.TargetUser
    
        ' <-------- Agrega Invisibilidad ---------->
        If Hechizos(HechizoIndex).Invisibilidad = 1 Then
            If UserList(TargetIndex).flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡El usuario está muerto!", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                HechizoCasteado = False

                Exit Sub

            End If
        
            If UserList(TargetIndex).Counters.Saliendo Then
                If UserIndex <> TargetIndex Then
                    Call WriteConsoleMsg(UserIndex, "¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                    HechizoCasteado = False

                    Exit Sub

                Else
                    Call WriteConsoleMsg(UserIndex, "¡No puedes hacerte invisible mientras te encuentras saliendo!", FontTypeNames.FONTTYPE_WARNING)
                    HechizoCasteado = False

                    Exit Sub

                End If

            End If
        
            'No usar invi mapas InviSinEfecto
            If MapInfo(UserList(TargetIndex).Pos.Map).InviSinEfecto > 0 Then
                Call WriteConsoleMsg(UserIndex, "¡La invisibilidad no funciona aquí!", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False

                Exit Sub

            End If
            
            If .flags.SlotEvent > 0 Then
                If Events(.flags.SlotEvent).config(eConfigEvent.eInvisibilidad) = 0 Then
                    Call WriteConsoleMsg(UserIndex, "¡La invisibilidad no funciona aquí! Retirate de la Zona del Evento si deseas utilizar el hechizo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If

            End If
                
            If .flags.SlotFast > 0 Then
                If RetoFast(.flags.SlotFast).ConfigVale <> ValeTodo Then
                    Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite este hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                    Exit Sub

                End If

            End If
        
            ' No invi en zona segura
            If Not MapInfo(.Pos.Map).Pk Then
                Call WriteConsoleMsg(UserIndex, "El hechizo no tiene efecto en zona segura.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            If UserList(TargetIndex).flags.Mimetizado = 1 Then
                Call WriteConsoleMsg(UserIndex, "El hechizo no tiene efecto estando mimetizado.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
                
            If Power.UserIndex = TargetIndex Then
                Call WriteConsoleMsg(UserIndex, "El personaje posee un poder superior.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            If UserList(TargetIndex).flags.Invisible = 1 Or UserList(TargetIndex).flags.Oculto = 1 Then
                Call WriteConsoleMsg(UserIndex, "El personaje ya se encuentra invisible.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            ' Chequea si el status permite ayudar al otro usuario
            HechizoCasteado = CanSupportUser(UserIndex, TargetIndex, True)

            If Not HechizoCasteado Then Exit Sub
        
            'Si sos user, no uses este hechizo con GMS.
            If .flags.Privilegios And PlayerType.User Then
                If Not UserList(TargetIndex).flags.Privilegios And PlayerType.User Then
                    HechizoCasteado = False

                    Exit Sub

                End If

            End If
            
            UserList(TargetIndex).flags.Invisible = 1
            
            ' Solo se hace invi para los clientes si no esta navegando
            If UserList(TargetIndex).flags.Navegando = 0 Then
                Call SetInvisible(TargetIndex, UserList(TargetIndex).Char.charindex, True)
                UserList(TargetIndex).Counters.DrawersCount = RandomNumberPower(1, 200)

            End If
        
            Call InfoHechizo(UserIndex)
            
            HechizoCasteado = True

        End If
    
        ' <-------- Agrega Mimetismo ---------->
        If Hechizos(HechizoIndex).Mimetiza = 1 Then
            If TargetIndex = UserIndex Then Exit Sub
            
            If UserList(TargetIndex).flags.Muerto = 1 Then

                Exit Sub

            End If
            
            If UserList(UserIndex).flags.Navegando = 1 Then

                Exit Sub

            End If
        
            If UserList(TargetIndex).flags.Navegando = 1 Then

                Exit Sub

            End If
        
            If UserList(TargetIndex).flags.Transform = 1 Then

                Exit Sub

            End If

            If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Sub
        
            If UserList(TargetIndex).flags.Transform = 1 Then

                Exit Sub

            End If
            
            If UserList(TargetIndex).flags.TransformVIP = 1 Then

                Exit Sub

            End If
            
            If Not MapInfo(.Pos.Map).Pk Then
                Call WriteConsoleMsg(UserIndex, "El hechizo no tiene efecto en zona segura.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            If MapInfo(.Pos.Map).MimetismoSinEfecto = 1 Then
                Call WriteConsoleMsg(UserIndex, "El mapa no permite el efecto mimetismo.", FontTypeNames.FONTTYPE_INFO)
            
                Exit Sub

            End If
            
            'Si sos user, no uses este hechizo con GMS.
            If .flags.Privilegios And PlayerType.User Then
                If Not UserList(TargetIndex).flags.Privilegios And PlayerType.User Then

                    Exit Sub

                End If

            End If
        
            If .flags.Mimetizado = 1 Then
                Call WriteConsoleMsg(UserIndex, "Ya te encuentras mimetizado. El hechizo no ha tenido efecto.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            If .flags.AdminInvisible = 1 Then Exit Sub

            If UserList(TargetIndex).flags.Invisible = 1 Or UserList(TargetIndex).flags.Oculto = 1 Then
                Call WriteConsoleMsg(UserIndex, "El hechizo no tiene efecto estando invisible.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            'copio el char original al mimetizado
        
            .CharMimetizado.Body = .Char.Body
            .CharMimetizado.Head = .Char.Head
            .CharMimetizado.CascoAnim = .Char.CascoAnim
            .CharMimetizado.ShieldAnim = .Char.ShieldAnim
            .CharMimetizado.WeaponAnim = .Char.WeaponAnim
        
            .flags.Mimetizado = 1
            .flags.Ignorado = True
            
            'ahora pongo local el del enemigo
            .Char.Body = UserList(TargetIndex).Char.Body
            .Char.Head = UserList(TargetIndex).Char.Head
            .Char.CascoAnim = UserList(TargetIndex).Char.CascoAnim
            .Char.ShieldAnim = UserList(TargetIndex).Char.ShieldAnim
            .Char.WeaponAnim = UserList(TargetIndex).Char.WeaponAnim
        
            Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
       
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True

        End If
    
        ' <-------- Agrega Envenenamiento ---------->
        If Hechizos(HechizoIndex).Envenena = 1 Then
            If UserIndex = TargetIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                Exit Sub

            End If
        
            If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Sub
        
            If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Sub
            If UserIndex <> TargetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

            End If

            UserList(TargetIndex).flags.Envenenado = 1
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True

        End If
    
        ' <-------- Cura Envenenamiento ---------->
        If Hechizos(HechizoIndex).CuraVeneno = 1 Then
            
            If UserList(TargetIndex).flags.Envenenado = 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje no está envenenado.", FontTypeNames.FONTTYPE_INFORED)
                HechizoCasteado = False
                Exit Sub

            End If
            
            'Verificamos que el usuario no este muerto
            If UserList(TargetIndex).flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡El usuario está muerto!", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                HechizoCasteado = False

                Exit Sub

            End If
        
            ' Chequea si el status permite ayudar al otro usuario
            HechizoCasteado = CanSupportUser(UserIndex, TargetIndex)

            If Not HechizoCasteado Then Exit Sub
            
            'Si sos user, no uses este hechizo con GMS.
            If .flags.Privilegios And PlayerType.User Then
                If Not UserList(TargetIndex).flags.Privilegios And PlayerType.User Then

                    Exit Sub

                End If

            End If
            
            UserList(TargetIndex).flags.Envenenado = 0
            Call WriteUpdateEffect(TargetIndex)
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True

        End If
    
        ' <-------- Agrega Maldicion ---------->
        If Hechizos(HechizoIndex).Maldicion = 1 Then
            If UserIndex = TargetIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                Exit Sub

            End If
        
            If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Sub
        
            If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Sub
            If UserIndex <> TargetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

            End If

            UserList(TargetIndex).flags.Maldicion = 1
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True

        End If
    
        ' <-------- Remueve Maldicion ---------->
        If Hechizos(HechizoIndex).RemoverMaldicion = 1 Then
            UserList(TargetIndex).flags.Maldicion = 0
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True

        End If
    
        ' <-------- Agrega Bendicion ---------->
        If Hechizos(HechizoIndex).Bendicion = 1 Then
            UserList(TargetIndex).flags.Bendicion = 1
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True

        End If
    
        ' <-------- Agrega Paralisis/Inmobilidad ---------->
        If Hechizos(HechizoIndex).Paraliza = 1 Or Hechizos(HechizoIndex).Inmoviliza = 1 Then
            If UserIndex = TargetIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                Exit Sub

            End If
            
            If .flags.SlotReto > 0 Then
                If Retos(.flags.SlotReto).config(eRetoConfig.eInmovilizar) = 0 Then
                    Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite este hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                    Exit Sub
                
                End If
            
            End If
            
            If .flags.SlotEvent > 0 Then
                If Events(.flags.SlotEvent).config(eConfigEvent.eUseParalizar) = 0 Then
                    Call WriteConsoleMsg(UserIndex, "No tienes permitido utilizar esta clase de hechizos en el evento.", FontTypeNames.FONTTYPE_INFORED)
                    Exit Sub

                End If

            End If
            
            If UserList(TargetIndex).flags.Paralizado = 0 Then
                If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Sub
            
                If UserIndex <> TargetIndex Then
                    Call checkHechizosEfectividad(UserIndex, TargetIndex)
                    Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

                End If
            
                Call InfoHechizo(UserIndex)
                HechizoCasteado = True

                If UserList(TargetIndex).Invent.AnilloEqpObjIndex = SUPERANILLO Then
                    Call WriteConsoleMsg(TargetIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                    Call WriteConsoleMsg(UserIndex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                    Call FlushBuffer(TargetIndex)

                    Exit Sub

                End If
                
                If Power.UserIndex = TargetIndex Then
                    Call WriteConsoleMsg(TargetIndex, "¡Te han querido inmovilizar!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                    Call WriteConsoleMsg(UserIndex, " ¡Ingenuo! El poder de las medusas es superior al tuyo", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                    Call FlushBuffer(TargetIndex)

                    Exit Sub

                End If
            
                If Hechizos(HechizoIndex).Inmoviliza = 1 Then UserList(TargetIndex).flags.Inmovilizado = 1
                UserList(TargetIndex).flags.Paralizado = 1
                UserList(TargetIndex).Counters.Paralisis = IIf(.Stats.MaxMan = 0, (IntervaloParalizado / 2), IntervaloParalizado)
            
                UserList(TargetIndex).flags.ParalizedByIndex = UserIndex
                UserList(TargetIndex).flags.ParalizedBy = UserList(UserIndex).Name
                
                If UserList(TargetIndex).flags.SlotEvent = 0 Then
                    Call SendData(SendTarget.ToOne, TargetIndex, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " lanzó " & Hechizos(HechizoIndex).Nombre, FontTypeNames.FONTTYPE_FIGHT))
                
                    Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageConsoleMsg(Hechizos(HechizoIndex).Nombre & " sobre " & UserList(TargetIndex).Name, FontTypeNames.FONTTYPE_FIGHT))

                End If
                
                Call WriteParalizeOK(TargetIndex)
                Call FlushBuffer(TargetIndex)

            End If

        End If
    
        ' <-------- Remueve Paralisis/Inmobilidad ---------->
        If Hechizos(HechizoIndex).RemoverParalisis = 1 Then
        
            ' Remueve si esta en ese estado
            If UserList(TargetIndex).flags.Paralizado = 1 Then
        
                ' Chequea si el status permite ayudar al otro usuario
                HechizoCasteado = CanSupportUser(UserIndex, TargetIndex, True)

                If Not HechizoCasteado Then Exit Sub
                      
                Call RemoveParalisis(TargetIndex)
                Call InfoHechizo(UserIndex)
                Call WriteConsoleMsg(TargetIndex, "¡" & .Name & " te ha devuelto la movilidad!", FontTypeNames.FONTTYPE_USERPLATA, eMessageType.Combate)
                Call WriteConsoleMsg(UserIndex, "¡Has devuelvo la movilidad a " & UserList(TargetIndex).Name & "!", FontTypeNames.FONTTYPE_USERPLATA, eMessageType.Combate)
        
            End If

        End If
    
        ' <-------- Remueve Estupidez (Aturdimiento) ---------->
        If Hechizos(HechizoIndex).RemoverEstupidez = 1 Then
    
            ' Remueve si esta en ese estado
            If UserList(TargetIndex).flags.Estupidez = 1 Then
        
                ' Chequea si el status permite ayudar al otro usuario
                HechizoCasteado = CanSupportUser(UserIndex, TargetIndex)

                If Not HechizoCasteado Then Exit Sub
        
                UserList(TargetIndex).flags.Estupidez = 0
            
                'no need to crypt this
                Call WriteDumbNoMore(TargetIndex)
                Call FlushBuffer(TargetIndex)
                Call InfoHechizo(UserIndex)
        
            End If

        End If
    
        ' <-------- Revive ---------->
        If Hechizos(HechizoIndex).Revivir = 1 Then
            If UserList(TargetIndex).flags.Muerto = 1 Then
            
                'Seguro de resurreccion (solo afecta a los hechizos, no al sacerdote ni al comando de GM)
                If UserList(TargetIndex).flags.SeguroResu Then
                    Call WriteConsoleMsg(UserIndex, "¡El espíritu no tiene intenciones de regresar al mundo de los vivos!", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False

                    Exit Sub

                End If
        
                'No usar resu en mapas con ResuSinEfecto
                If MapInfo(UserList(TargetIndex).Pos.Map).ResuSinEfecto > 0 Then
                    Call WriteConsoleMsg(UserIndex, "¡Revivir no está permitido aquí! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False

                    Exit Sub

                End If
                
                If .flags.SlotReto > 0 Then
                    If Retos(.flags.SlotReto).config(eRetoConfig.eResucitar) = 0 Then
                        Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite este hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
    
                        Exit Sub
                    
                    End If
                
                End If
                
                If .flags.SlotEvent > 0 Then
                    If Events(.flags.SlotEvent).config(eConfigEvent.eResu) = 0 Then
                        Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite este hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                        Exit Sub

                    End If

                End If
                    
                If .flags.SlotFast > 0 Then
                    If RetoFast(.flags.SlotFast).ConfigVale <> ValeResu And RetoFast(.flags.SlotFast).ConfigVale <> ValeTodo Then
                        Call WriteConsoleMsg(UserIndex, "El evento en el que estás participando no permite este hechizo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                        Exit Sub

                    End If

                End If
            
                'revisamos si necesita vara
                If .Clase = eClass.Mage Then
                    If .Invent.WeaponEqpObjIndex > 0 Then
                        If ObjData(.Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                            Call WriteConsoleMsg(UserIndex, "Necesitas un báculo mejor para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                            HechizoCasteado = False

                            Exit Sub

                        End If

                    End If

                ElseIf .Clase = eClass.Bard Then

                    If .Invent.MagicObjIndex <> LAUDMAGICO Then
                        Call WriteConsoleMsg(UserIndex, "Necesitas un instrumento mágico para devolver la vida.", FontTypeNames.FONTTYPE_INFO)
                        HechizoCasteado = False

                        Exit Sub

                    End If

                ElseIf .Clase = eClass.Druid Then

                    If .Invent.MagicObjIndex <> ANILLOMAGICO Then
                        Call WriteConsoleMsg(UserIndex, "Necesitas un instrumento mágico para devolver la vida.", FontTypeNames.FONTTYPE_INFO)
                        HechizoCasteado = False

                        Exit Sub

                    End If

                End If
            
                ' Chequea si el status permite ayudar al otro usuario
                HechizoCasteado = CanSupportUser(UserIndex, TargetIndex, True)

                If Not HechizoCasteado Then Exit Sub
    
                Dim EraCriminal As Boolean

                EraCriminal = Escriminal(UserIndex)
            
                If Not Escriminal(TargetIndex) Then
                    If TargetIndex <> UserIndex Then
                        .Reputacion.NobleRep = .Reputacion.NobleRep + 500

                        If .Reputacion.NobleRep > MAXREP Then .Reputacion.NobleRep = MAXREP
                        Call WriteConsoleMsg(UserIndex, "¡Los Dioses te sonríen, has ganado 500 puntos de nobleza!", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
            
                If EraCriminal And Not Escriminal(UserIndex) Then
                    Call RefreshCharStatus(UserIndex)

                End If
            
                With UserList(TargetIndex)
                    'Pablo Toxic Waste (GD: 29/04/07)
                    .Stats.MinAGU = 0
                    .flags.Sed = 1
                    .Stats.MinHam = 0
                    .flags.Hambre = 1
                    Call WriteUpdateHungerAndThirst(TargetIndex)
                    Call InfoHechizo(UserIndex)
                    .Stats.MinMan = 0
                    .Stats.MinSta = 0

                End With
            
                'Agregado para quitar la penalización de vida en el ring y cambio de ecuacion. (NicoNZ)
                If (TriggerZonaPelea(UserIndex, TargetIndex) <> TRIGGER6_PERMITE) Then

                    'Solo saco vida si es User. no quiero que exploten GMs por ahi.
                    If .flags.Privilegios And PlayerType.User Then
                        If .Clase <> eClass.Cleric Then
                            .Stats.MinHp = .Stats.MinHp * (1 - (.Stats.Elv) * 0.015)

                        End If

                    End If

                End If
            
                If (.Stats.MinHp <= 0) Then
                    Call UserDie(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar fue demasiado grande.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                Else

                    If .Clase <> eClass.Cleric Then
                        Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar te ha debilitado.", FontTypeNames.FONTTYPE_INFO)

                    End If

                    HechizoCasteado = True

                End If
            
                If UserList(TargetIndex).flags.Traveling = 1 Then
                    Call EndTravel(TargetIndex, True)

                End If
            
                Call RevivirUsuario(TargetIndex)
            Else
                HechizoCasteado = False

            End If
    
        End If
    
        ' <-------- Agrega Ceguera ---------->
        If Hechizos(HechizoIndex).Ceguera = 1 Then
            If UserIndex = TargetIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                Exit Sub

            End If
            
            If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Sub
        
            If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Sub
            If UserIndex <> TargetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

            End If

            UserList(TargetIndex).flags.Ceguera = 1
            UserList(TargetIndex).Counters.Ceguera = IntervaloParalizado / 3
    
            Call WriteBlind(TargetIndex)
            Call FlushBuffer(TargetIndex)
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True

        End If
    
        ' <-------- Agrega Estupidez (Aturdimiento) ---------->
        If Hechizos(HechizoIndex).Estupidez = 1 Then
            If UserIndex = TargetIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                Exit Sub

            End If
            
            If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Sub
        
            If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Sub
            If UserIndex <> TargetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

            End If

            If UserList(TargetIndex).flags.Estupidez = 0 Then
                UserList(TargetIndex).flags.Estupidez = 1
                UserList(TargetIndex).Counters.Ceguera = IntervaloParalizado

            End If

            Call WriteDumb(TargetIndex)
            Call FlushBuffer(TargetIndex)
    
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True

        End If

    End With

    '<EhFooter>
    Exit Sub

HechizoEstadoUsuario_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.HechizoEstadoUsuario " & "at line " & Erl

    '</EhFooter>
End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, _
                     ByVal SpellIndex As Integer, _
                     ByRef HechizoCasteado As Boolean, _
                     ByVal UserIndex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 07/07/2008
    'Handles the Spells that afect the Stats of an NPC
    '04/13/2008 NicoNZ - Guardias Faccionarios pueden ser
    'removidos por users de su misma faccion.
    '07/07/2008: NicoNZ - Solo se puede mimetizar con npcs si es druida
    '***************************************************
    '<EhHeader>
    On Error GoTo HechizoEstadoNPC_Err

    '</EhHeader>

    With Npclist(NpcIndex)

        If Hechizos(SpellIndex).Invisibilidad = 1 Then
            Call InfoHechizo(UserIndex)
            .flags.Invisible = 1
            HechizoCasteado = True

        End If
    
        If Hechizos(SpellIndex).Envenena = 1 Then
            If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
                HechizoCasteado = False

                Exit Sub

            End If

            Call NPCAtacado(NpcIndex, UserIndex)
            Call InfoHechizo(UserIndex)
            .flags.Envenenado = 1
            HechizoCasteado = True

        End If
    
        If Hechizos(SpellIndex).CuraVeneno = 1 Then
            Call InfoHechizo(UserIndex)
            .flags.Envenenado = 0
            HechizoCasteado = True

        End If
    
        If Hechizos(SpellIndex).Maldicion = 1 Then
            If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
                HechizoCasteado = False

                Exit Sub

            End If

            Call NPCAtacado(NpcIndex, UserIndex)
            Call InfoHechizo(UserIndex)
            .flags.Maldicion = 1
            HechizoCasteado = True

        End If
    
        If Hechizos(SpellIndex).RemoverMaldicion = 1 Then
            Call InfoHechizo(UserIndex)
            .flags.Maldicion = 0
            HechizoCasteado = True

        End If
    
        If Hechizos(SpellIndex).Bendicion = 1 Then
            Call InfoHechizo(UserIndex)
            .flags.Bendicion = 1
            HechizoCasteado = True

        End If
    
        If Hechizos(SpellIndex).Paraliza = 1 And Hechizos(SpellIndex).Inmoviliza = 1 Then
            If .flags.AfectaParalisis = 0 Then
                If Not PuedeAtacarNPC(UserIndex, NpcIndex, True) Then
                    HechizoCasteado = False

                    Exit Sub

                End If

                Call NPCAtacado(NpcIndex, UserIndex)
                Call InfoHechizo(UserIndex)
                .flags.Paralizado = 1
                .flags.Inmovilizado = 1
                .Contadores.Paralisis = (IntervaloParalizado * 4)
                Call AnimacionIdle(NpcIndex, False)
                HechizoCasteado = True
            Else
                Call WriteConsoleMsg(UserIndex, "El NPC es inmune a este hechizo.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)
                HechizoCasteado = False

                Exit Sub

            End If

        End If
    
        If Hechizos(SpellIndex).RemoverParalisis = 1 Then
            If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
                If .MaestroUser = UserIndex Then
                    Call InfoHechizo(UserIndex)
                    .flags.Paralizado = 0
                    .Contadores.Paralisis = 0
                    HechizoCasteado = True
                Else

                    If .NPCtype = eNPCType.GuardiaReal Then
                        If esArmada(UserIndex) Then
                            Call InfoHechizo(UserIndex)
                            .flags.Paralizado = 0
                            .Contadores.Paralisis = 0
                            HechizoCasteado = True

                            Exit Sub

                        Else
                            Call WriteConsoleMsg(UserIndex, "Sólo puedes remover la parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
                            HechizoCasteado = False

                            Exit Sub

                        End If
                    
                        Call WriteConsoleMsg(UserIndex, "Solo puedes remover la parálisis de los NPCs que te consideren su amo.", FontTypeNames.FONTTYPE_INFO)
                        HechizoCasteado = False

                        Exit Sub

                    Else

                        If .NPCtype = eNPCType.GuardiasCaos Then
                            If esCaos(UserIndex) Then
                                Call InfoHechizo(UserIndex)
                                .flags.Paralizado = 0
                                .Contadores.Paralisis = 0
                                HechizoCasteado = True

                                Exit Sub

                            Else
                                Call WriteConsoleMsg(UserIndex, "Solo puedes remover la parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
                                HechizoCasteado = False

                                Exit Sub

                            End If

                        End If

                    End If

                End If

            Else
                Call WriteConsoleMsg(UserIndex, "Este NPC no está paralizado", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False

                Exit Sub

            End If

        End If
     
        If Hechizos(SpellIndex).Paraliza = 1 And Hechizos(SpellIndex).Inmoviliza = 0 Then
            If .flags.AfectaParalisis = 0 Then
                If Not PuedeAtacarNPC(UserIndex, NpcIndex, True) Then
                    HechizoCasteado = False

                    Exit Sub

                End If

                Call NPCAtacado(NpcIndex, UserIndex)
                .flags.Inmovilizado = 1
                .flags.Paralizado = 0
                .Contadores.Paralisis = (IntervaloParalizado * 3)
                Call InfoHechizo(UserIndex)
                Call AnimacionIdle(NpcIndex, True)
                HechizoCasteado = True
            Else
                Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

    If Hechizos(SpellIndex).Mimetiza = 1 Then

        With UserList(UserIndex)

            If .flags.Mimetizado = 1 Then
                Call WriteConsoleMsg(UserIndex, "Ya te encuentras mimetizado. El hechizo no ha tenido efecto.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
            
            If .flags.Navegando = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes mimetizarte navegando.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            If .flags.Invisible = 1 Or .flags.Oculto = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes mimetizarte estando invisible.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            If .flags.Transform = 1 Or .flags.TransformVIP Then
                Call WriteConsoleMsg(UserIndex, "No puedes mimetizarte en ese estado.", FontTypeNames.FONTTYPE_INFORED)

                Exit Sub

            End If
            
            If Not MapInfo(.Pos.Map).Pk Then
                Call WriteConsoleMsg(UserIndex, "El hechizo tiene efecto en zonas inseguras", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
        
            If MapInfo(.Pos.Map).MimetismoSinEfecto = 1 Then
                Call WriteConsoleMsg(UserIndex, "El mapa no permite el efecto mimetismo.", FontTypeNames.FONTTYPE_INFO)
            
                Exit Sub

            End If
                
            If Npclist(NpcIndex).Char.Body = 0 Then
                Call WriteConsoleMsg(UserIndex, "¡No puedes tomar la forma de la criatura!", FontTypeNames.FONTTYPE_INFO)
            
                Exit Sub

            End If
                
            If .flags.AdminInvisible = 1 Then Exit Sub
            
            If .Clase = eClass.Druid Then
                'copio el char original al mimetizado
            
                .CharMimetizado.Body = .Char.Body
                .CharMimetizado.Head = .Char.Head
                .CharMimetizado.CascoAnim = .Char.CascoAnim
                .CharMimetizado.ShieldAnim = .Char.ShieldAnim
                .CharMimetizado.WeaponAnim = .Char.WeaponAnim
            
                .flags.Mimetizado = 1
                .ShowName = False
                .flags.Ignorado = True
                
                'ahora pongo lo del NPC.
                .Char.Body = Npclist(NpcIndex).Char.Body
                .Char.Head = Npclist(NpcIndex).Char.Head
                .Char.CascoAnim = NingunCasco
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                      
                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraIndex)
                Call RefreshCharStatus(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "Sólo los druidas pueden mimetizarse con criaturas.", FontTypeNames.FONTTYPE_INFO)

                Exit Sub

            End If
    
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True

        End With

    End If

    '<EhFooter>
    Exit Sub

HechizoEstadoNPC_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.HechizoEstadoNPC " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub HechizoPropNPC(ByVal SpellIndex As Integer, _
                   ByVal NpcIndex As Integer, _
                   ByVal UserIndex As Integer, _
                   ByRef HechizoCasteado As Boolean)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 18/09/2010
    'Handles the Spells that afect the Life NPC
    '14/08/2007 Pablo (ToxicWaste) - Orden general.
    '18/09/2010: ZaMa - Ahora valida si podes ayudar a un npc.
    '***************************************************
    '<EhHeader>
    On Error GoTo HechizoPropNPC_Err

    '</EhHeader>

    Dim daño As Long

    With Npclist(NpcIndex)

        'Salud
        If Hechizos(SpellIndex).SubeHP = 1 Then
        
            HechizoCasteado = CanSupportNpc(UserIndex, NpcIndex)
        
            If HechizoCasteado Then
        
                If .Hostile = 0 Then
                    Call WriteConsoleMsg(UserIndex, "No puedes curar a la criatura", FontTypeNames.FONTTYPE_INFORED)
                    HechizoCasteado = False

                    Exit Sub

                End If
            
                daño = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
                daño = daño + Porcentaje(daño, 3 * (UserList(UserIndex).Stats.Elv))
            
                Call InfoHechizo(UserIndex)
                .Stats.MinHp = .Stats.MinHp + daño

                If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
                Call WriteConsoleMsg(UserIndex, "Has curado " & daño & " puntos de vida a la criatura.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, d_CurarSpell))

            End If
        
        ElseIf Hechizos(SpellIndex).SubeHP = 2 Then

            If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
                HechizoCasteado = False

                Exit Sub

            End If
        
            Call NPCAtacado(NpcIndex, UserIndex)
            daño = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
            daño = daño + Porcentaje(daño, 3 * (UserList(UserIndex).Stats.Elv))
            
            If Hechizos(SpellIndex).StaffAffected Then
                If UserList(UserIndex).Clase = eClass.Mage Then
                    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                        daño = (daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                        'Aumenta daño segun el staff-
                        'Daño = (Daño* (70 + BonifBáculo)) / 100
                    Else
                        daño = daño * 0.7 'Baja daño a 70% del original

                    End If

                End If

            End If
        
            If .NPCtype = DRAGON Then
                If UserList(UserIndex).Invent.WeaponEqpObjIndex = VaraMataDragonesIndex Then
                    daño = daño * 3

                End If

            End If
        
            'Esta con gran poder?
            If Power.UserIndex = UserIndex Then
                daño = daño * 1.2

            End If
            
            If UserList(UserIndex).Invent.MagicObjIndex = LAUDMAGICO Or UserList(UserIndex).Invent.MagicObjIndex = ANILLOMAGICO Then
                daño = daño * 1.04  'laud magico de los bardos 4%

            End If
        
            #If Testeo = 1 Then

                If EsAdmin(UCase$(UserList(UserIndex).Name)) Then
                    daño = .Stats.MaxHp

                End If

            #End If
        
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True
        
            If .flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayEffect(.flags.Snd2, .Pos.X, .Pos.Y, .Char.charindex))

            End If
        
            'Quizas tenga defenza magica el NPC. Pablo (ToxicWaste)
            daño = daño - .Stats.defM
        
            If daño < 0 Then daño = 0
                
            Call CalcularDarExp(UserIndex, NpcIndex, daño)
            Call Quests_AddNpc(UserIndex, NpcIndex, daño)
                  
            .Stats.MinHp = .Stats.MinHp - daño
            Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, d_DañoNpcSpell))

            If .Stats.MinHp < 1 Then
                .Stats.MinHp = 0
                Call MuereNpc(NpcIndex, UserIndex)

            End If

        End If

    End With

    '<EhFooter>
    Exit Sub

HechizoPropNPC_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.HechizoPropNPC " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub InfoHechizo(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo InfoHechizo_Err

    '</EhHeader>

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 25/07/2009
    '25/07/2009: ZaMa - Code improvements.
    '25/07/2009: ZaMa - Now invisible admins magic sounds are not sent to anyone but themselves
    '***************************************************
    Dim SpellIndex As Integer

    Dim tUser      As Integer

    Dim tNpc       As Integer

    Dim Valid      As Boolean: Valid = True
    
    With UserList(UserIndex)
        SpellIndex = .flags.Hechizo
                
        If Hechizos(SpellIndex).AutoLanzar = 1 Then
            tUser = UserIndex
        Else
            tUser = .flags.TargetUser

        End If
                
        tNpc = .flags.TargetNPC
        
        If .flags.SlotEvent > 0 Then
            If Events(.flags.SlotEvent).Modality = eModalityEvent.DeathMatch Then
                Valid = False

            End If

        End If
        
        If Valid Then Call DecirPalabrasMagicas(Hechizos(SpellIndex).PalabrasMagicas, UserIndex)
        
        If tUser > 0 Then
            ' bueno hace eso para todos como primer paso, avismae cuando lo hayas hecho joya avisme por face cuando lo termines sisi
            ' Los admins invisibles no producen sonidos ni fx's
            
            If .flags.AdminInvisible = 1 And UserIndex = tUser Then
                Call SendData(ToOne, UserIndex, PrepareMessageCreateFX(UserList(tUser).Char.charindex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
                Call SendData(ToOne, UserIndex, PrepareMessagePlayEffect(Hechizos(SpellIndex).WAV, UserList(tUser).Pos.X, UserList(tUser).Pos.Y, UserList(tUser).Char.charindex))
            Else
                Call SendData(SendTarget.ToPCArea, tUser, PrepareMessageCreateFX(UserList(tUser).Char.charindex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
                Call SendData(SendTarget.ToPCArea, tUser, PrepareMessagePlayEffect(Hechizos(SpellIndex).WAV, UserList(tUser).Pos.X, UserList(tUser).Pos.Y, UserList(tUser).Char.charindex))
                
            End If

        ElseIf tNpc > 0 Then
            Call SendData(SendTarget.ToNPCArea, tNpc, PrepareMessageCreateFX(Npclist(tNpc).Char.charindex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
            Call SendData(SendTarget.ToNPCArea, tNpc, PrepareMessagePlayEffect(Hechizos(SpellIndex).WAV, Npclist(tNpc).Pos.X, Npclist(tNpc).Pos.Y, Npclist(tNpc).Char.charindex))
            
        End If

    End With

    '<EhFooter>
    Exit Sub

InfoHechizo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.InfoHechizo " & "at line " & Erl
        
    '</EhFooter>
End Sub

' Hechizos que causan efectos sobre el Area
Public Function HechizoPropAreaUsuario(ByVal UserIndex As Integer) As Boolean

    Dim SpellIndex  As Integer

    Dim Damage      As Long

    Dim TargetIndex As Integer
    
    Dim A           As Long
    
    Dim X           As Byte, Y As Byte
    
    Dim Spell       As tHechizo
    
    With UserList(UserIndex)
        SpellIndex = .flags.Hechizo
        Spell = Hechizos(SpellIndex)
        
        For X = .Pos.X - Spell.AreaX To .Pos.X + Spell.AreaX
            For Y = .Pos.Y - Spell.AreaY To .Pos.Y + Spell.AreaY
                TargetIndex = MapData(.Pos.Map, X, Y).UserIndex
                
                If TargetIndex > 0 Then

                    ' @ Quita SALUD
                    If Spell.SubeHP = 2 And TargetIndex <> UserIndex Then
                        If PuedeAtacar(UserIndex, TargetIndex) Then
                            If HechizoUserReceiveDamage(UserIndex, SpellIndex) Then
                                
                                HechizoPropAreaUsuario = True
                                Damage = HechizoUserUpdateDamage(UserIndex, TargetIndex, SpellIndex)
                                
                                If Damage > 0 Then
                                    ' Call InfoHechizo(UserIndex)
                                    
                                    With UserList(TargetIndex)
                                        .Stats.MinHp = .Stats.MinHp - Damage
                                        Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
                                        Call SubirSkill(TargetIndex, eSkill.Resistencia, True)
                                        Call WriteUpdateHP(TargetIndex)
                                        
                                        Call SendData(SendTarget.ToPCArea, TargetIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, Damage, d_DañoUserSpell))
                                        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageUpdateControlPotas(.Char.charindex, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMan, .Stats.MaxMan))
            
                                        'Muere
                                        If .Stats.MinHp < 1 Then
                                            If .flags.AtacablePor <> UserIndex Then Call ContarMuerte(TargetIndex, UserIndex)
                
                                            .Stats.MinHp = 0
                                            Call ActStats(TargetIndex, UserIndex)
                                            Call UserDie(TargetIndex, UserIndex)
    
                                        End If
    
                                    End With

                                End If
                                
                            End If

                        End If

                    End If

                End If

            Next Y
        
        Next X

        ' Effects User
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(Hechizos(SpellIndex).WAV, X, Y))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Hechizos(SpellIndex).PalabrasMagicas, UserList(UserIndex).Char.charindex, vbCyan))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFXMap(.Pos.X, .Pos.Y, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))

    End With
    
End Function

' @ Comprueba que el usuario pueda recibir un ataque de daño mágico
Public Function HechizoUserReceiveDamage(ByVal UserIndex As Integer, _
                                         ByVal SpellIndex As Integer) As Boolean
    
    With UserList(UserIndex)
        
        If .flags.SlotEvent > 0 Then
            If Events(.flags.SlotEvent).config(eConfigEvent.eUseTormenta) = 0 And SpellIndex = eHechizosIndex.eTormenta Then
                Call WriteConsoleMsg(UserIndex, "¡No puedes utilizar este hechizo en el evento!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Exit Function

            End If
                
            If Events(.flags.SlotEvent).config(eConfigEvent.eUseApocalipsis) = 0 And (SpellIndex = eHechizosIndex.eApocalipsis Or SpellIndex = eHechizosIndex.eExplosionAbismal) Then
                Call WriteConsoleMsg(UserIndex, "¡No puedes utilizar este hechizo en el evento!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Exit Function

            End If
                
            If Events(.flags.SlotEvent).config(eConfigEvent.eUseDescarga) = 0 And SpellIndex = eHechizosIndex.eDescarga Then
                Call WriteConsoleMsg(UserIndex, "¡No puedes utilizar este hechizo en el evento!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Exit Function

            End If

        End If
    
    End With

    HechizoUserReceiveDamage = True

End Function

' @ Actualiza el Damage mágico del poder
Public Function HechizoUserUpdateDamage(ByVal UserIndex As Integer, _
                                        ByVal TargetIndex As Integer, _
                                        ByVal SpellIndex As Integer) As Long
    
    Dim Damage As Long
    
    With UserList(TargetIndex)
        Damage = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
        
        Damage = Damage + Porcentaje(Damage, 3 * (UserList(UserIndex).Stats.Elv))
        
        If Hechizos(SpellIndex).StaffAffected Then
            If UserList(UserIndex).Clase = eClass.Mage Then
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    Damage = (Damage * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                Else
                    Damage = Damage * 0.7 'Baja Damage a 70% del original

                End If

            End If

        End If
    
        If UserList(UserIndex).Invent.MagicObjIndex = LAUDMAGICO Then
            Damage = Damage * 1.05  'laud magico de los bardos y anillos de druidas

        End If
                
        If UserList(UserIndex).Invent.MagicObjIndex = ANILLOMAGICO Then
            Damage = Damage * 1.03  'laud magico de los bardos y anillos de druidas

        End If

        'cascos antimagia
        If (.Invent.CascoEqpObjIndex > 0) Then
            Damage = Damage - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)

        End If
        
        'If .Invent.EscudoEqpObjIndex > 0 Then
        'Damage = Damage - RandomNumber(ObjData(.Invent.EscudoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.EscudoEqpObjIndex).DefensaMagicaMax)
        'End If
                
        'If .Invent.ArmourEqpObjIndex > 0 Then
        'Damage = Damage - RandomNumber(ObjData(.Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.ArmourEqpObjIndex).DefensaMagicaMax)
        'End If
                
        'anillos
        If (.Invent.AnilloEqpObjIndex > 0) Then
            Damage = Damage - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)

        End If
            
        'Esta con gran poder?
        If Power.UserIndex = UserIndex Then
            Damage = Damage * 1.05

        End If
        
        ' Bonos
        If .flags.SelectedBono > 0 Then
        
            ' Bonos RM
            If ObjData(.flags.SelectedBono).BonoRm > 0 Then
                Damage = Damage * ObjData(.flags.SelectedBono).BonoRm

            End If

        End If
        
        If UserList(UserIndex).flags.SelectedBono > 0 Then
            
            ' Bonos Damage mágicos
            If ObjData(UserList(UserIndex).flags.SelectedBono).BonoHechizos > 0 Then
                Damage = Damage * ObjData(UserList(UserIndex).flags.SelectedBono).BonoHechizos

            End If
            
        End If
    
        Damage = Damage - (Damage * .Stats.UserSkills(eSkill.Resistencia) / 2000)
        
        If Damage < 0 Then Damage = 0
        
        HechizoUserUpdateDamage = Damage

    End With

End Function

Public Function HechizoPropUsuario(ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 28/04/2010
    '02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
    '28/04/2010: ZaMa - Agrego Restricciones para ciudas respecto al estado atacable.
    '***************************************************
    '<EhHeader>
    On Error GoTo HechizoPropUsuario_Err

    '</EhHeader>

    Dim SpellIndex As Integer

    Dim daño As Long

    Dim TargetIndex As Integer

    SpellIndex = UserList(UserIndex).flags.Hechizo
    TargetIndex = UserList(UserIndex).flags.TargetUser
      
    With UserList(TargetIndex)

        If .flags.Muerto Then
            Call WriteConsoleMsg(UserIndex, "No puedes lanzar este hechizo a un muerto.", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

            Exit Function

        End If
          
        ' <-------- Aumenta Hambre ---------->
        If Hechizos(SpellIndex).SubeHam = 1 Then
        
            Call InfoHechizo(UserIndex)
        
            daño = RandomNumber(Hechizos(SpellIndex).MinHam, Hechizos(SpellIndex).MaxHam)
        
            .Stats.MinHam = .Stats.MinHam + daño

            If .Stats.MinHam > .Stats.MaxHam Then .Stats.MinHam = .Stats.MaxHam
        
            If UserIndex <> TargetIndex Then
                Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de hambre a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Else
                Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

            End If
        
            Call WriteUpdateHungerAndThirst(TargetIndex)
    
            ' <-------- Quita Hambre ---------->
        ElseIf Hechizos(SpellIndex).SubeHam = 2 Then

            If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Function
            If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
            If UserIndex <> TargetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
            Else

                Exit Function

            End If
        
            Call InfoHechizo(UserIndex)
        
            daño = RandomNumber(Hechizos(SpellIndex).MinHam, Hechizos(SpellIndex).MaxHam)
        
            .Stats.MinHam = .Stats.MinHam - daño
        
            If UserIndex <> TargetIndex Then
                Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de hambre a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Else
                Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

            End If
        
            If .Stats.MinHam < 1 Then
                .Stats.MinHam = 0
                .flags.Hambre = 1

            End If
        
            Call WriteUpdateHungerAndThirst(TargetIndex)

        End If
    
        ' <-------- Aumenta Sed ---------->
        If Hechizos(SpellIndex).SubeSed = 1 Then
        
            Call InfoHechizo(UserIndex)
        
            daño = RandomNumber(Hechizos(SpellIndex).MinSed, Hechizos(SpellIndex).MaxSed)
        
            .Stats.MinAGU = .Stats.MinAGU + daño

            If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
        
            Call WriteUpdateHungerAndThirst(TargetIndex)
             
            If UserIndex <> TargetIndex Then
                Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de sed a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Else
                Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

            End If
    
            ' <-------- Quita Sed ---------->
        ElseIf Hechizos(SpellIndex).SubeSed = 2 Then

            If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Function
        
            If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
            If UserIndex <> TargetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

            End If
        
            Call InfoHechizo(UserIndex)
        
            daño = RandomNumber(Hechizos(SpellIndex).MinSed, Hechizos(SpellIndex).MaxSed)
        
            .Stats.MinAGU = .Stats.MinAGU - daño
        
            If UserIndex <> TargetIndex Then
                Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de sed a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Else
                Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

            End If
        
            If .Stats.MinAGU < 1 Then
                .Stats.MinAGU = 0
                .flags.Sed = 1

            End If
        
            Call WriteUpdateHungerAndThirst(TargetIndex)
        
        End If
    
        ' <-------- Aumenta Agilidad ---------->
        If Hechizos(SpellIndex).SubeAgilidad = 1 Then
        
            ' Chequea si el status permite ayudar al otro usuario
            If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
        
            Call InfoHechizo(UserIndex)
            daño = RandomNumber(Hechizos(SpellIndex).MinAgilidad, Hechizos(SpellIndex).MaxAgilidad)
        
            .flags.DuracionEfecto = 1200
            .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + daño

            If .Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Agilidad) * 2) Then .Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Agilidad) * 2)
        
            .flags.TomoPocion = True
            Call WriteUpdateDexterity(TargetIndex)
    
            ' <-------- Quita Agilidad ---------->
        ElseIf Hechizos(SpellIndex).SubeAgilidad = 2 Then

            If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Function
        
            If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
            If UserIndex <> TargetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

            End If
        
            Call InfoHechizo(UserIndex)
        
            .flags.TomoPocion = True
            daño = RandomNumber(Hechizos(SpellIndex).MinAgilidad, Hechizos(SpellIndex).MaxAgilidad)
            .flags.DuracionEfecto = 700
            .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) - daño

            If .Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then .Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
        
            Call WriteUpdateDexterity(TargetIndex)

        End If
    
        ' <-------- Aumenta Fuerza ---------->
        If Hechizos(SpellIndex).SubeFuerza = 1 Then
    
            ' Chequea si el status permite ayudar al otro usuario
            If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
        
            Call InfoHechizo(UserIndex)
            daño = RandomNumber(Hechizos(SpellIndex).MinFuerza, Hechizos(SpellIndex).MaxFuerza)
        
            .flags.DuracionEfecto = 1200
    
            .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + daño

            If .Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Fuerza) * 2) Then .Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Fuerza) * 2)
        
            .flags.TomoPocion = True
            Call WriteUpdateStrenght(TargetIndex)
    
            ' <-------- Quita Fuerza ---------->
        ElseIf Hechizos(SpellIndex).SubeFuerza = 2 Then

            If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Then Exit Function
            If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
            If UserIndex <> TargetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

            End If
        
            Call InfoHechizo(UserIndex)
        
            .flags.TomoPocion = True
        
            daño = RandomNumber(Hechizos(SpellIndex).MinFuerza, Hechizos(SpellIndex).MaxFuerza)
            .flags.DuracionEfecto = 700
            .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) - daño

            If .Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then .Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
        
            Call WriteUpdateStrenght(TargetIndex)

        End If
    
        ' <-------- Cura salud ---------->
        If Hechizos(SpellIndex).SubeHP = 1 Then
        
            'Verifica que el usuario no este muerto
            If .flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "¡El usuario está muerto!", FontTypeNames.FONTTYPE_INFO)

                Exit Function

            End If
            
            If .flags.SlotEvent > 0 Or .flags.SlotFast > 0 Or .flags.SlotReto > 0 Or .flags.Desafiando > 0 Then
                Call WriteConsoleMsg(UserIndex, "¡No se permite curar desde donde estás!", FontTypeNames.FONTTYPE_INFO)

                Exit Function

            End If
        
            ' Chequea si el status permite ayudar al otro usuario
            If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
           
            If .Stats.MinHp = .Stats.MaxHp Then
                Call WriteConsoleMsg(UserIndex, "El personaje está sano", FontTypeNames.FONTTYPE_INFORED)

                Exit Function

            End If
        
            daño = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
            daño = daño + Porcentaje(daño, 3 * (.Stats.Elv))
        
            Call InfoHechizo(UserIndex)
    
            .Stats.MinHp = .Stats.MinHp + daño

            If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
        
            Call WriteUpdateHP(TargetIndex)
        
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageUpdateControlPotas(.Char.charindex, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMan, .Stats.MaxMan))
            Call SendData(SendTarget.ToPCArea, TargetIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, d_CurarSpell))
        
            ' <-------- Quita salud (Daña) ---------->
        ElseIf Hechizos(SpellIndex).SubeHP = 2 Then
        
            If UserIndex = TargetIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

                Exit Function

            End If
            
            ' Chequeo de Eventos (Anti Spells)
            If .flags.SlotEvent > 0 Then
                If Events(.flags.SlotEvent).config(eConfigEvent.eUseTormenta) = 0 And SpellIndex = eHechizosIndex.eTormenta Then
                    Call WriteConsoleMsg(UserIndex, "¡No puedes utilizar este hechizo en el evento!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                    Exit Function

                End If
                
                If Events(.flags.SlotEvent).config(eConfigEvent.eUseApocalipsis) = 0 And (SpellIndex = eHechizosIndex.eApocalipsis Or SpellIndex = eHechizosIndex.eExplosionAbismal) Then
                    Call WriteConsoleMsg(UserIndex, "¡No puedes utilizar este hechizo en el evento!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                    Exit Function

                End If
                
                If Events(.flags.SlotEvent).config(eConfigEvent.eUseDescarga) = 0 And SpellIndex = eHechizosIndex.eDescarga Then
                    Call WriteConsoleMsg(UserIndex, "¡No puedes utilizar este hechizo en el evento!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                    Exit Function

                End If

            End If

            daño = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
        
            daño = daño + Porcentaje(daño, 3 * (UserList(UserIndex).Stats.Elv))
        
            If Hechizos(SpellIndex).StaffAffected Then
                If UserList(UserIndex).Clase = eClass.Mage Then
                    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                        daño = (daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                    Else
                        daño = daño * 0.7 'Baja daño a 70% del original

                    End If

                End If

            End If

            If UserList(UserIndex).Invent.MagicObjIndex = LAUDMAGICO Then
                daño = daño * 1.05  'laud magico de los bardos y anillos de druidas

            End If
                
            If UserList(UserIndex).Invent.MagicObjIndex = ANILLOMAGICO Then
                daño = daño * 1.03  'laud magico de los bardos y anillos de druidas

            End If

            'cascos antimagia
            If (.Invent.CascoEqpObjIndex > 0) Then
                daño = daño - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)

            End If
                
            If .Pos.Map <> 130 And .Pos.Map <> 131 And .Pos.Map <> 132 Then
                
                ' Daño mágico para los clanes con CASTILLO NORTE
                If Castle_CheckBonus(UserList(UserIndex).GuildIndex, eCastle.CASTLE_NORTH) Then
                    daño = daño * 1.02

                End If
                    
                ' Resistencia mágica para los clanes con CASTILLO OESTE
                If Castle_CheckBonus(.GuildIndex, eCastle.CASTLE_WEST) Then
                    daño = daño * 0.98

                End If
                    
            End If
                
            'If .Invent.EscudoEqpObjIndex > 0 Then
            'Daño = Daño - RandomNumber(ObjData(.Invent.EscudoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.EscudoEqpObjIndex).DefensaMagicaMax)
            'End If
                
            'If .Invent.ArmourEqpObjIndex > 0 Then
            'Daño = Daño - RandomNumber(ObjData(.Invent.ArmourEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.ArmourEqpObjIndex).DefensaMagicaMax)
            'End If
                
            'anillos
            If (.Invent.AnilloEqpObjIndex > 0) Then
                daño = daño - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)

            End If
            
            'Esta con gran poder?
            If Power.UserIndex = UserIndex Then
                daño = daño * 1.05

            End If
        
            ' Bonos
            If .flags.SelectedBono > 0 Then
        
                ' Bonos RM
                If ObjData(.flags.SelectedBono).BonoRm > 0 Then
                    daño = daño * ObjData(.flags.SelectedBono).BonoRm

                End If

            End If
        
            If UserList(UserIndex).flags.SelectedBono > 0 Then
            
                ' Bonos Daño mágicos
                If ObjData(UserList(UserIndex).flags.SelectedBono).BonoHechizos > 0 Then
                    daño = daño * ObjData(UserList(UserIndex).flags.SelectedBono).BonoHechizos

                End If
            
            End If
        
            ' ReliquiaDrag equipped
            'If UserList(UserIndex).Invent.ReliquiaSlot > 0 Then
            'Daño = Effect_UpdatePorc(UserIndex, Daño)
            'End If
        
            ' ReliquiaDrag equipped
            'If .Invent.ReliquiaSlot > 0 Then
            ' Daño = Effect_UpdatePorc(TargetIndex, Daño)
            'End If
        
            daño = daño - (daño * UserList(TargetIndex).Stats.UserSkills(eSkill.Resistencia) / 2000)
        
            If daño < 0 Then daño = 0
        
            If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
            If UserIndex <> TargetIndex Then
                Call checkHechizosEfectividad(UserIndex, TargetIndex)
                Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

            End If
            
            Call InfoHechizo(UserIndex)
        
            If UserList(UserIndex).flags.SlotEvent > 0 Then
                Events_Add_Damage UserList(UserIndex).flags.SlotEvent, UserList(UserIndex).flags.SlotUserEvent, daño

            End If
        
            .Stats.MinHp = .Stats.MinHp - daño
        
            Call SubirSkill(TargetIndex, eSkill.Resistencia, True)
            Call WriteUpdateHP(TargetIndex)
        
            Dim Valid As Boolean: Valid = True

            If .flags.SlotEvent > 0 Then
                If Events(.flags.SlotEvent).Modality = eModalityEvent.DeathMatch Then
                    Valid = False

                End If

            End If
        
            If Valid Then
                Call SendData(SendTarget.ToOne, TargetIndex, PrepareMessageConsoleMsg(UserList(UserIndex).Name & " lanzó " & Hechizos(SpellIndex).Nombre & " -" & daño, FontTypeNames.FONTTYPE_FIGHT))
                
                Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageConsoleMsg(Hechizos(SpellIndex).Nombre & " sobre " & UserList(TargetIndex).Name & " -" & daño, FontTypeNames.FONTTYPE_FIGHT))

            End If

            Call SendData(SendTarget.ToPCArea, TargetIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, d_DañoUserSpell))
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageUpdateControlPotas(.Char.charindex, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMan, .Stats.MaxMan))
        
            'Muere
            If .Stats.MinHp < 1 Then
        
                If .flags.AtacablePor <> UserIndex Then
                    'Store it!
                    ' Call Statistics.StoreFrag(UserIndex, TargetIndex)
                    Call ContarMuerte(TargetIndex, UserIndex)

                End If
            
                .Stats.MinHp = 0
                Call ActStats(TargetIndex, UserIndex)
                Call UserDie(TargetIndex, UserIndex)
                Call SendData(SendTarget.ToPCArea, TargetIndex, PrepareMessageCreateFX(UserList(TargetIndex).Char.charindex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))

            End If
        
        End If
    
        ' <-------- Aumenta Mana ---------->
        If Hechizos(SpellIndex).SubeMana = 1 Then
        
            Call InfoHechizo(UserIndex)
            .Stats.MinMan = .Stats.MinMan + daño

            If .Stats.MinMan > .Stats.MaxMan Then .Stats.MinMan = .Stats.MaxMan
        
            Call WriteUpdateMana(TargetIndex)
        
            If UserIndex <> TargetIndex Then
                Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de maná a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Else
                Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

            End If
    
            ' <-------- Quita Mana ---------->
        ElseIf Hechizos(SpellIndex).SubeMana = 2 Then

            If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
            If UserIndex <> TargetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

            End If
        
            Call InfoHechizo(UserIndex)
        
            If UserIndex <> TargetIndex Then
                Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de maná a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Else
                Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

            End If
        
            .Stats.MinMan = .Stats.MinMan - daño

            If .Stats.MinMan < 1 Then .Stats.MinMan = 0
        
            Call WriteUpdateMana(TargetIndex)
        
        End If
    
        ' <-------- Aumenta Stamina ---------->
        If Hechizos(SpellIndex).SubeSta = 1 Then
            Call InfoHechizo(UserIndex)
            .Stats.MinSta = .Stats.MinSta + daño

            If .Stats.MinSta > .Stats.MaxSta Then .Stats.MinSta = .Stats.MaxSta
        
            Call WriteUpdateSta(TargetIndex)
        
            If UserIndex <> TargetIndex Then
                Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de energía a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Else
                Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

            End If
        
            ' <-------- Quita Stamina ---------->
        ElseIf Hechizos(SpellIndex).SubeSta = 2 Then

            If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
        
            If UserIndex <> TargetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)

            End If
        
            Call InfoHechizo(UserIndex)
        
            If UserIndex <> TargetIndex Then
                Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de energía a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            Else
                Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)

            End If
        
            .Stats.MinSta = .Stats.MinSta - daño
        
            If .Stats.MinSta < 1 Then .Stats.MinSta = 0
        
            Call WriteUpdateSta(TargetIndex)
        
        End If

    End With

    HechizoPropUsuario = True

    Call FlushBuffer(TargetIndex)

    '<EhFooter>
    Exit Function

HechizoPropUsuario_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.HechizoPropUsuario " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function CanSupportUser(ByVal CasterIndex As Integer, _
                               ByVal TargetIndex As Integer, _
                               Optional ByVal DoCriminal As Boolean = False) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 28/04/2010
    'Checks if caster can cast support magic on target user.
    '***************************************************
     
    On Error GoTo ErrHandler
 
    With UserList(CasterIndex)
        
        ' Te podes curar a vos mismo
        If CasterIndex = TargetIndex Then
            CanSupportUser = True

            Exit Function

        End If
        
        ' No podes ayudar si estas en consulta
        If .flags.EnConsulta Then
            Call WriteConsoleMsg(CasterIndex, "No puedes ayudar usuarios mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If
        
        ' Si estas en la arena, esta todo permitido
        If TriggerZonaPelea(CasterIndex, TargetIndex) = TRIGGER6_PERMITE Then
            CanSupportUser = True

            Exit Function

        End If
     
        ' Victima criminal?
        If Escriminal(TargetIndex) Then
        
            ' Casteador Ciuda?
            If Not Escriminal(CasterIndex) Then
            
                ' Armadas no pueden ayudar
                If esArmada(CasterIndex) Then
                    Call WriteConsoleMsg(CasterIndex, "Los miembros del ejército real no pueden ayudar a los criminales.", FontTypeNames.FONTTYPE_INFO)

                    Exit Function

                End If
                
                ' Si el ciuda tiene el seguro puesto no puede ayudar
                If .flags.Seguro Then
                    Call WriteConsoleMsg(CasterIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos.", FontTypeNames.FONTTYPE_INFO)

                    Exit Function

                Else

                    ' Penalizacion
                    If DoCriminal Then
                        Call VolverCriminal(CasterIndex)
                    Else
                        Call DisNobAuBan(CasterIndex, .Reputacion.NobleRep * 0.5, 10000)

                    End If

                End If

            End If
            
            ' Victima ciuda o army
        Else

            ' Casteador es caos? => No Pueden ayudar ciudas
            If esCaos(CasterIndex) Then
                Call WriteConsoleMsg(CasterIndex, "Los miembros de la legión oscura no pueden ayudar a los ciudadanos.", FontTypeNames.FONTTYPE_INFO)

                Exit Function
                
                ' Casteador ciuda/army?
            ElseIf Not Escriminal(CasterIndex) Then
                
                ' Esta en estado atacable?
                If UserList(TargetIndex).flags.AtacablePor > 0 Then
                    
                    ' No esta atacable por el casteador?
                    If UserList(TargetIndex).flags.AtacablePor <> CasterIndex Then
                    
                        ' Si es armada no puede ayudar
                        If esArmada(CasterIndex) Then
                            Call WriteConsoleMsg(CasterIndex, "Los miembros del ejército real no pueden ayudar a ciudadanos en estado atacable.", FontTypeNames.FONTTYPE_INFO)

                            Exit Function

                        End If
    
                        ' Seguro puesto?
                        If .flags.Seguro Then
                            Call WriteConsoleMsg(CasterIndex, "Para ayudar ciudadanos en estado atacable debes sacarte el seguro, pero te puedes volver criminal.", FontTypeNames.FONTTYPE_INFO)

                            Exit Function

                        Else
                            Call DisNobAuBan(CasterIndex, .Reputacion.NobleRep * 0.5, 10000)

                        End If

                    End If

                End If
    
            End If

        End If

    End With
    
    CanSupportUser = True

    Exit Function
    
ErrHandler:
    Call LogError("Error en CanSupportUser, Error: " & Err.number & " - " & Err.description & " CasterIndex: " & CasterIndex & ", TargetIndex: " & TargetIndex)

End Function

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, _
                       ByVal UserIndex As Integer, _
                       ByVal Slot As Byte)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo UpdateUserHechizos_Err

    '</EhHeader>

    Dim LoopC As Byte

    With UserList(UserIndex)

        'Actualiza un solo slot
        If Not UpdateAll Then

            'Actualiza el inventario
            If .Stats.UserHechizos(Slot) > 0 Then
                Call ChangeUserHechizo(UserIndex, Slot, .Stats.UserHechizos(Slot))
            Else
                Call ChangeUserHechizo(UserIndex, Slot, 0)

            End If

        Else

            'Actualiza todos los slots
            For LoopC = 1 To MAXUSERHECHIZOS

                'Actualiza el inventario
                If .Stats.UserHechizos(LoopC) > 0 Then
                    Call ChangeUserHechizo(UserIndex, LoopC, .Stats.UserHechizos(LoopC))
                Else
                    Call ChangeUserHechizo(UserIndex, LoopC, 0)

                End If
            
            Next LoopC

        End If

    End With

    '<EhFooter>
    Exit Sub

UpdateUserHechizos_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.UpdateUserHechizos " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function CanSupportNpc(ByVal CasterIndex As Integer, _
                              ByVal TargetIndex As Integer) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 18/09/2010
    'Checks if caster can cast support magic on target Npc.
    '***************************************************
     
    On Error GoTo ErrHandler
 
    Dim OwnerIndex As Integer
 
    With UserList(CasterIndex)
        
        OwnerIndex = Npclist(TargetIndex).Owner
        
        ' Si no tiene dueño puede
        If OwnerIndex = 0 Then
            CanSupportNpc = True

            Exit Function

        End If
        
        ' Puede hacerlo si es su propio npc
        If CasterIndex = OwnerIndex Then
            CanSupportNpc = True

            Exit Function

        End If
        
        ' No podes ayudar si estas en consulta
        If .flags.EnConsulta Then
            Call WriteConsoleMsg(CasterIndex, "No puedes ayudar npcs mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If
        
        ' Si estas en la arena, esta todo permitido
        If TriggerZonaPelea(CasterIndex, OwnerIndex) = TRIGGER6_PERMITE Then
            CanSupportNpc = True

            Exit Function

        End If
     
        ' Victima criminal?
        If Escriminal(OwnerIndex) Then

            ' Victima caos?
            If esCaos(OwnerIndex) Then

                ' Atacante caos?
                If esCaos(CasterIndex) Then
                    ' No podes ayudar a un npc de un caos si sos caos
                    Call WriteConsoleMsg(CasterIndex, "No puedes ayudar npcs que están luchando contra un miembro de tu facción.", FontTypeNames.FONTTYPE_INFO)

                    Exit Function

                End If

            End If
        
            ' Uno es caos y el otro no, o la victima es pk, entonces puede ayudar al npc
            CanSupportNpc = True

            Exit Function
                
            ' Victima ciuda
        Else

            ' Atacante ciuda?
            If Not Escriminal(CasterIndex) Then

                ' Atacante armada?
                If esArmada(CasterIndex) Then

                    ' Victima armada?
                    If esArmada(OwnerIndex) Then
                        ' No podes ayudar a un npc de un armada si sos armada
                        Call WriteConsoleMsg(CasterIndex, "No puedes ayudar npcs que están luchando contra un miembro de tu facción.", FontTypeNames.FONTTYPE_INFO)

                        Exit Function

                    End If

                End If
                
                ' Uno es armada y el otro ciuda, o los dos ciudas, puede atacar si no tiene seguro
                If .flags.Seguro Then
                    Call WriteConsoleMsg(CasterIndex, "Para ayudar a criaturas que luchan contra ciudadanos debes sacarte el seguro.", FontTypeNames.FONTTYPE_INFO)

                    Exit Function

                End If
                
            End If
            
            ' Atacante criminal y victima ciuda, entonces puede ayudar al npc
            CanSupportNpc = True

            Exit Function
            
        End If
    
    End With
    
    CanSupportNpc = True

    Exit Function
    
ErrHandler:
    Call LogError("Error en CanSupportNpc, Error: " & Err.number & " - " & Err.description & " CasterIndex: " & CasterIndex & ", OwnerIndex: " & OwnerIndex)

End Function

Sub ChangeUserHechizo(ByVal UserIndex As Integer, _
                      ByVal Slot As Byte, _
                      ByVal Hechizo As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ChangeUserHechizo_Err

    '</EhHeader>
    
    UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo
    
    If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
        Call WriteChangeSpellSlot(UserIndex, Slot)
    Else
        Call WriteChangeSpellSlot(UserIndex, Slot)

    End If

    '<EhFooter>
    Exit Sub

ChangeUserHechizo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.ChangeUserHechizo " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub DisNobAuBan(ByVal UserIndex As Integer, NoblePts As Long, BandidoPts As Long)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo DisNobAuBan_Err

    '</EhHeader>

    'disminuye la nobleza NoblePts puntos y aumenta el bandido BandidoPts puntos
    Dim EraCriminal As Boolean

    EraCriminal = Escriminal(UserIndex)
    
    With UserList(UserIndex)

        'Si estamos en la arena no hacemos nada
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        If .flags.Privilegios And (PlayerType.User) Then
            'pierdo nobleza...
            .Reputacion.NobleRep = .Reputacion.NobleRep - NoblePts

            If .Reputacion.NobleRep < 0 Then
                .Reputacion.NobleRep = 0

            End If
            
            'gano bandido...
            .Reputacion.BandidoRep = .Reputacion.BandidoRep + BandidoPts

            If .Reputacion.BandidoRep > MAXREP Then .Reputacion.BandidoRep = MAXREP
            Call WriteMultiMessage(UserIndex, eMessages.NobilityLost) 'Call WriteNobilityLost(UserIndex)

            If Escriminal(UserIndex) Then
                If .Faction.Status = r_Armada Then
                    Call mFacciones.Faction_RemoveUser(UserIndex)
                Else
                    Call Guilds_CheckAlineation(UserIndex, a_Neutral)

                End If

            End If

        End If
        
        If Not EraCriminal And Escriminal(UserIndex) Then
            Call RefreshCharStatus(UserIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

DisNobAuBan_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.DisNobAuBan " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub Events_GranBestia_AttackUsers(ByVal NpcIndex As Integer)

    '<EhHeader>
    On Error GoTo Events_GranBestia_AttackUsers_Err

    '</EhHeader>

    Dim TempX        As Integer

    Dim TempY        As Integer
    
    Dim X            As Integer

    Dim Y            As Integer
    
    Dim Damage       As Long
    
    Dim UserIndex    As Integer
    
    Dim Attacks      As Byte
    
    Const SpellIndex As Byte = 51

    Const MAX_ATTACK As Byte = 4
    
    Const Map        As Byte = 65

    Const MIN_X      As Byte = 48

    Const MAX_X      As Byte = 62

    Const MIN_Y      As Byte = 43

    Const MAX_Y      As Byte = 54
    
    X = RandomNumber(MIN_X, MAX_X)
    Y = RandomNumber(MIN_Y, MAX_Y)
    
    With Npclist(NpcIndex)

        For TempX = X To RandomNumber(MAX_X - 3, MAX_X)
            For TempY = Y To RandomNumber(MAX_Y - 3, MAX_Y)

                If InMapBounds(Map, TempX, TempY) Then
                    UserIndex = MapData(Map, TempX, TempY).UserIndex
                    
                    If UserIndex > 0 Then

                        With UserList(UserIndex)

                            If .flags.SlotEvent > 0 Then
                            
                                If RandomNumber(1, 100) <= 10 Then
                                    Damage = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp) + RandomNumber(50, 100)
                                Else
                                    Damage = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)

                                End If
                                
                                .Stats.MinHp = .Stats.MinHp - Damage
                                
                                Call WriteUpdateHP(UserIndex)
                                Call WriteConsoleMsg(UserIndex, "¡La gran bestia te ha quitado " & Damage & " puntos de vida!", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(Hechizos(SpellIndex).WAV, .Pos.X, .Pos.Y, .Char.charindex))
                                
                                If .Stats.MinHp <= 0 Then
                                    .Stats.MinHp = .Stats.MaxHp
                                    Events_GranBestia_MuereUser (UserIndex)

                                    Exit Sub

                                End If
                            
                            End If

                        End With
                        
                        Attacks = Attacks + 1

                    End If
                    
                    If Attacks = MAX_ATTACK Then Exit Sub

                End If

            Next TempY
        Next TempX

    End With

    '<EhFooter>
    Exit Sub

Events_GranBestia_AttackUsers_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.Events_GranBestia_AttackUsers " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub checkHechizosEfectividad(ByVal UserIndex As Integer, ByVal TargetUser As Integer)

    '<EhHeader>
    On Error GoTo checkHechizosEfectividad_Err

    '</EhHeader>
    With UserList(UserIndex)

        If .Pos.Map = 1 Then Exit Sub
        
        If UserList(TargetUser).flags.Inmovilizado + UserList(TargetUser).flags.Paralizado = 0 Then
            .Counters.controlHechizos.HechizosCasteados = .Counters.controlHechizos.HechizosCasteados + 1
        
            Dim efectividad As Double
            
            efectividad = (100 * .Counters.controlHechizos.HechizosCasteados) / .Counters.controlHechizos.HechizosTotales
            
            If efectividad >= 85 And .Counters.controlHechizos.HechizosTotales >= 10 Then
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("El usuario " & .Name & " está lanzando hechizos con una efectividad de " & efectividad & "% (Casteados: " & .Counters.controlHechizos.HechizosCasteados & "/" & .Counters.controlHechizos.HechizosTotales & "), revisar.", FontTypeNames.FONTTYPE_TALK))
                Call Logs_Security(eSecurity, eLogSecurity.eAntiCheat, "El usuario " & .Name & " con IP: " & .IpAddress & " está lanzando hechizos con una efectividad de " & efectividad & "% (Casteados: " & .Counters.controlHechizos.HechizosCasteados & "/" & .Counters.controlHechizos.HechizosTotales & "), revisar.")

            End If
           
        Else
            .Counters.controlHechizos.HechizosTotales = .Counters.controlHechizos.HechizosTotales - 1

        End If

    End With

    '<EhFooter>
    Exit Sub

checkHechizosEfectividad_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.modHechizos.checkHechizosEfectividad " & "at line " & Erl
        
    '</EhFooter>
End Sub

