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

Public Const HELEMENTAL_FUEGO As Integer = 26
Public Const HELEMENTAL_TIERRA As Integer = 28
Public Const SUPERANILLO As Integer = 649

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal Userindex As Integer, ByVal spell As Integer, _
                           Optional ByVal DecirPalabras As Boolean = False, _
                           Optional ByVal IgnoreVisibilityCheck As Boolean = False)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 13/02/2009
'13/02/2009: ZaMa - Los npcs que tiren magias, no podran hacerlo en mapas donde no se permita usarla.
'***************************************************
If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
If UserList(Userindex).flags.invisible = 1 Or UserList(Userindex).flags.Oculto = 1 Then Exit Sub

' Si no se peude usar magia en el mapa, no le deja hacerlo.
If MapInfo(UserList(Userindex).Pos.Map).MagiaSinEfecto > 0 Then Exit Sub

Npclist(NpcIndex).CanAttack = 0
Dim daño As Integer

With UserList(Userindex)
    If Hechizos(spell).SubeHP = 1 Then
    
        daño = RandomNumber(Hechizos(spell).MinHp, Hechizos(spell).MaxHp)
        daño = daño - (daño * UserList(Userindex).Stats.UserSkills(eSkill.Resistencia) / 2000)
        ' daño = daño - Porcentaje(daño, Int(((UserList(UserIndex).Stats.UserSkills(Resistencia) + 1) / 4) + ResistenciaClase(UserList(UserIndex).clase = Druid Or Cleric Or Assasin Or Bard Or Mage Or Paladin Or warrior Or Hunter)))
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(spell).WAV, .Pos.X, .Pos.Y))
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(spell).FXgrh, Hechizos(spell).loops))
    
        .Stats.MinHp = .Stats.MinHp + daño
        If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
        
        Call WriteConsoleMsg(Userindex, Npclist(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteUpdateUserStats(Userindex)
    
    ElseIf Hechizos(spell).SubeHP = 2 Then
        
        If .flags.Privilegios And PlayerType.User Then
         daño = daño - (daño * UserList(Userindex).Stats.UserSkills(eSkill.Resistencia) / 2000)
       ' daño = daño - Porcentaje(daño, Int(((UserList(UserIndex).Stats.UserSkills(Resistencia) + 1) / 4) + ResistenciaClase(UserList(UserIndex).clase = (Druid Or Mage Or Paladin Or Hunter Or Assasin Or Cleric Or Pirat Or Bard))))
    Call SubirSkill(Userindex, eSkill.Resistencia, True)
            daño = RandomNumber(Hechizos(spell).MinHp, Hechizos(spell).MaxHp)
            
            If .Invent.CascoEqpObjIndex > 0 Then
                daño = daño - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)
            End If
            daño = daño - (daño * UserList(Userindex).Stats.UserSkills(eSkill.Resistencia) / 2000)
           '  daño = daño - Porcentaje(daño, Int(((UserList(UserIndex).Stats.UserSkills(Resistencia) + 1) / 4) + ResistenciaClase(UserList(UserIndex).clase = Druid Or Cleric Or Assasin Or Bard Or Mage Or Paladin Or warrior Or Hunter)))
            If .Invent.AnilloEqpObjIndex > 0 Then
                daño = daño - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)
            End If
           daño = daño - (daño * UserList(Userindex).Stats.UserSkills(eSkill.Resistencia) / 2000)
          ' daño = daño - Porcentaje(daño, Int(((UserList(UserIndex).Stats.UserSkills(Resistencia) + 1) / 4) + ResistenciaClase(UserList(UserIndex).clase = Druid Or Cleric Or Assasin Or Bard Or Mage Or Paladin Or warrior Or Hunter)))
            If daño < 0 Then daño = 0
            
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(spell).WAV, .Pos.X, .Pos.Y))
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(spell).FXgrh, Hechizos(spell).loops))
        
            .Stats.MinHp = .Stats.MinHp - daño
            
            Call WriteConsoleMsg(Userindex, Npclist(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
                       SendData SendTarget.ToPCArea, Userindex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, DAMAGE_NORMAL)
            Call WriteUpdateUserStats(Userindex)
            
            'Muere
            If .Stats.MinHp < 1 Then
                .Stats.MinHp = 0
                If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                    RestarCriminalidad (Userindex)
                End If
                Call UserDie(Userindex)
                '[Barrin 1-12-03]
                If Npclist(NpcIndex).MaestroUser > 0 Then
                    'Store it!
                    Call Statistics.StoreFrag(Npclist(NpcIndex).MaestroUser, Userindex)
                    
                    Call ContarMuerte(Userindex, Npclist(NpcIndex).MaestroUser)
                    Call ActStats(Userindex, Npclist(NpcIndex).MaestroUser)
                End If
                '[/Barrin]
            End If
        
        End If
        
    End If
    
    If Hechizos(spell).Paraliza = 1 Or Hechizos(spell).Inmoviliza = 1 Then
        If .flags.Paralizado = 0 Then
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(spell).WAV, .Pos.X, .Pos.Y))
            Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(spell).FXgrh, Hechizos(spell).loops))
              
            If .Invent.AnilloEqpObjIndex = SUPERANILLO Then
                Call WriteConsoleMsg(Userindex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            
            If Hechizos(spell).Inmoviliza = 1 Then
                .flags.Inmovilizado = 1
            End If
              
            .flags.Paralizado = 1
            .Counters.Paralisis = IntervaloParalizado
              
            Call WriteParalizeOK(Userindex)
        End If
    End If
    
    If Hechizos(spell).Estupidez = 1 Then   ' turbacion
         If .flags.Estupidez = 0 Then
              Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(spell).WAV, .Pos.X, .Pos.Y))
              Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(spell).FXgrh, Hechizos(spell).loops))
              
                If .Invent.AnilloEqpObjIndex = SUPERANILLO Then
                    Call WriteConsoleMsg(Userindex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
              
              .flags.Estupidez = 1
              .Counters.Ceguera = IntervaloInvisible
                      
            Call WriteDumb(Userindex)
         End If
    End If
End With

End Sub

Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal spell As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

'solo hechizos ofensivos!

If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
Npclist(NpcIndex).CanAttack = 0

Dim daño As Integer

If Hechizos(spell).SubeHP = 2 Then
    
    daño = RandomNumber(Hechizos(spell).MinHp, Hechizos(spell).MaxHp)
    Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(spell).WAV, Npclist(TargetNPC).Pos.X, Npclist(TargetNPC).Pos.Y))
    Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(Npclist(TargetNPC).Char.CharIndex, Hechizos(spell).FXgrh, Hechizos(spell).loops))
    
    Npclist(TargetNPC).Stats.MinHp = Npclist(TargetNPC).Stats.MinHp - daño
    
    'Muere
    If Npclist(TargetNPC).Stats.MinHp < 1 Then
        Npclist(TargetNPC).Stats.MinHp = 0
        If Npclist(NpcIndex).MaestroUser > 0 Then
            Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
        Else
            Call MuereNpc(TargetNPC, 0)
        End If
    End If
    
End If
    
End Sub

Function TieneHechizo(ByVal i As Integer, ByVal Userindex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo Errhandler
    
    Dim j As Integer
    For j = 1 To MAXUSERHECHIZOS
        If UserList(Userindex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function
        End If
    Next

Exit Function
Errhandler:

End Function

Sub AgregarHechizo(ByVal Userindex As Integer, ByVal Slot As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim hIndex As Integer
Dim j As Integer
Dim i As Integer
Dim NoLoUsa As Integer

With UserList(Userindex)
    hIndex = ObjData(.Invent.Object(Slot).objindex).HechizoIndex
    
      For i = 1 To NUMCLASES
         If ObjData(.Invent.Object(Slot).objindex).ClaseProhibida(i) = UserList(Userindex).clase Then
             NoLoUsa = 1
             Call WriteConsoleMsg(Userindex, "Tu clase no puede aprender este hechizo.", FontTypeNames.FONTTYPE_INFO)
         End If
       Next i
    
    If Not TieneHechizo(hIndex, Userindex) Then
        'Buscamos un slot vacio
        For j = 1 To MAXUSERHECHIZOS
            If .Stats.UserHechizos(j) = 0 Then Exit For
        Next j
            
       If .Stats.UserHechizos(j) <> 0 Then
            Call WriteConsoleMsg(Userindex, "No tienes espacio para más hechizos.", FontTypeNames.FONTTYPE_INFO)
        Else
            If NoLoUsa = 0 Then
                .Stats.UserHechizos(j) = hIndex
                Call UpdateUserHechizos(False, Userindex, CByte(j))
                'Quitamos del inv el item
                Call QuitarUserInvItem(Userindex, CByte(Slot), 1)
            End If
        End If
    Else
        Call WriteConsoleMsg(Userindex, "Ya tienes ese hechizo.", FontTypeNames.FONTTYPE_INFO)
    End If
End With

End Sub
            
Sub DecirPalabrasMagicas(ByVal SpellWords As String, ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 17/11/2009
'25/07/2009: ZaMa - Invisible admins don't say any word when casting a spell
'17/11/2009: ZaMa - Now the user become visible when casting a spell, if it is hidden
'***************************************************
On Error Resume Next
With UserList(Userindex)
    If .flags.AdminInvisible <> 1 Then
        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(SpellWords, .Char.CharIndex, vbCyan))
        
        ' Si estaba oculto, se vuelve visible
        If .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.invisible = 0 Then
                Call WriteConsoleMsg(Userindex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                Call SetInvisible(Userindex, .Char.CharIndex, False)
            End If
        End If
    End If
End With
    Exit Sub
End Sub

''
' Check if an user can cast a certain spell
'
' @param UserIndex Specifies reference to user
' @param HechizoIndex Specifies reference to spell
' @return   True if the user can cast the spell, otherwise returns false
Function PuedeLanzar(ByVal Userindex As Integer, ByVal HechizoIndex As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2010
'Last Modification By: ZaMa
'06/11/09 - Corregida la bonificación de maná del mimetismo en el druida con flauta mágica equipada.
'19/11/2009: ZaMa - Validacion de mana para el Invocar Mascotas
'12/01/2010: ZaMa - Validacion de mana para hechizos lanzados por druida.
'***************************************************
Dim DruidManaBonus As Single

    With UserList(Userindex)
        If .flags.Muerto Then
            Call WriteConsoleMsg(Userindex, "No puedes lanzar hechizos estando muerto.", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
            
        If Hechizos(HechizoIndex).NeedStaff > 0 Then
        If UserList(Userindex).clase = eClass.Mage Then
            If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
                If ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                    Call WriteConsoleMsg(Userindex, "No posees un báculo lo suficientemente poderoso para que puedas lanzar el conjuro.", FontTypeNames.FONTTYPE_INFO)
                    PuedeLanzar = False
                    Exit Function
                End If
            Else
                Call WriteConsoleMsg(Userindex, "No puedes lanzar este conjuro sin la ayuda de un báculo.", FontTypeNames.FONTTYPE_INFO)
                PuedeLanzar = False
                Exit Function
            End If
        End If
    End If
        
    If UserList(Userindex).Stats.UserSkills(eSkill.Magia) < Hechizos(HechizoIndex).MinSkill Then
        Call WriteConsoleMsg(Userindex, "No tenes suficientes puntos de magia para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
        PuedeLanzar = False
        Exit Function
    End If
    
    If UserList(Userindex).Stats.MinSta < Hechizos(HechizoIndex).StaRequerido Then
        If UserList(Userindex).Genero = eGenero.Hombre Then
            Call WriteConsoleMsg(Userindex, "Estás muy cansado para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(Userindex, "Estás muy cansada para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
        End If
        PuedeLanzar = False
        Exit Function
    End If

    If UserList(Userindex).clase = eClass.Druid Then
        If UserList(Userindex).Invent.MunicionEqpObjIndex = FLAUTAELFICA And UserList(Userindex).Invent.MunicionEqpObjIndex = FLAUTAANTIGUA And UserList(Userindex).Invent.MunicionEqpObjIndex = AnilloBronce And UserList(Userindex).Invent.MunicionEqpObjIndex = AnilloPlata Then
            If Hechizos(HechizoIndex).Mimetiza Then
                DruidManaBonus = 0.5
            ElseIf Hechizos(HechizoIndex).tipo = uInvocacion Then
                DruidManaBonus = 0.7
            Else
                DruidManaBonus = 1
            End If
        Else
            DruidManaBonus = 1
        End If
    Else
        DruidManaBonus = 1
    End If
    
    If UserList(Userindex).Stats.MinMAN < Hechizos(HechizoIndex).ManaRequerido * DruidManaBonus Then
        Call WriteConsoleMsg(Userindex, "No tenes suficiente mana.", FontTypeNames.FONTTYPE_INFO)
        PuedeLanzar = False
        Exit Function
    End If
        End With
    PuedeLanzar = True
End Function

Sub HechizoTerrenoEstado(ByVal Userindex As Integer, ByRef b As Boolean)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim PosCasteadaX As Integer
Dim PosCasteadaY As Integer
Dim PosCasteadaM As Integer
Dim h As Integer
Dim TempX As Integer
Dim TempY As Integer

    With UserList(Userindex)
        PosCasteadaX = .flags.TargetX
        PosCasteadaY = .flags.TargetY
        PosCasteadaM = .flags.TargetMap
        
        h = .flags.Hechizo
        
        If Hechizos(h).RemueveInvisibilidadParcial = 1 Then
            b = True
            For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
                For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
                    If InMapBounds(PosCasteadaM, TempX, TempY) Then
                        If MapData(PosCasteadaM, TempX, TempY).Userindex > 0 Then
                            'hay un user
                            If UserList(MapData(PosCasteadaM, TempX, TempY).Userindex).flags.invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).Userindex).flags.AdminInvisible = 0 Then
                                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(UserList(MapData(PosCasteadaM, TempX, TempY).Userindex).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))
                            End If
                        End If
                    End If
                Next TempY
            Next TempX
        
            Call InfoHechizo(Userindex)
        End If
    End With
End Sub

''
' Le da propiedades al nuevo npc
'
' @param UserIndex  Indice del usuario que invoca.
' @param b  Indica si se termino la operación.

Sub HechizoInvocacion(ByVal Userindex As Integer, ByRef HechizoCasteado As Boolean)
'***************************************************
'Author: Uknown
'Last modification: 18/11/2009
'Sale del sub si no hay una posición valida.
'18/11/2009: Optimizacion de codigo.
'***************************************************
On Error GoTo error

With UserList(Userindex)
        'No permitimos se invoquen criaturas en zonas seguras
    If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then
        Call WriteConsoleMsg(Userindex, "No puedes invocar criaturas en zona segura.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If

 If (Hechizos(.flags.Hechizo).NumNpc = 111 Or Hechizos(.flags.Hechizo).NumNpc = 110 Or Hechizos(.flags.Hechizo).NumNpc = ELEMENTALFUEGO Or Hechizos(.flags.Hechizo).NumNpc = ELEMENTALTIERRA Or Hechizos(.flags.Hechizo).NumNpc = LOBO Or Hechizos(.flags.Hechizo).NumNpc = ZOMBIE Or Hechizos(.flags.Hechizo).NumNpc = ELEMENTALAGUA Or Hechizos(.flags.Hechizo).NumNpc = OSOS) And (MapInfo(.Pos.Map).Pk = False Or MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA) Then
    WriteConsoleMsg Userindex, "No puedes invocar criaturas en zona segura.", FontTypeNames.FONTTYPE_INFO
    Exit Sub
    End If

   ' If .Pos.Map = 298 Or .Pos.Map = 293 Or .Pos.Map = 276 Or .Pos.Map = 300 Then Exit Sub
    'No deja invocar mas de 1 fatuo o 1 espiritu indomable
       Dim SpellIndex As Integer, NroNpcs As Integer, NpcIndex As Integer, PetIndex As Integer
    Dim TargetPos As WorldPos
    
    
    TargetPos.Map = .flags.TargetMap
    TargetPos.X = .flags.TargetX
    TargetPos.Y = .flags.TargetY
    
    SpellIndex = .flags.Hechizo

     
    ' If Hechizos(SpellIndex).NumNpc = 110 And UserList(UserIndex).NroMascotas = 1 Then Exit Sub
   ' If Hechizos(SpellIndex).NumNpc = ELEMENTALFUEGO And UserList(UserIndex).MascotasIndex = 1 Then Exit Sub
    ' Warp de mascotas
    If Hechizos(SpellIndex).Warp = 1 Then
        PetIndex = FarthestPet(Userindex)
        
        ' La invoco cerca mio
        If Npclist(.MascotasType(.NroMascotas)).Contadores.TiempoExistencia = 0 Then
        If .NroMascotas > 0 Then
    WarpMascotas Userindex, False
    .NroMascotas = 0
    ElseIf .NroMascotas <= 0 Then
    WarpMascotas Userindex, True
    .NroMascotas = 3
        End If
        End If
    ' Invocacion normal
    Else
'solo 1 fuego fatuo puede ser invocadooooo wachin
    If PetIndex <= 0 Then
        If .NroMascotas >= MAXMASCOTAS Then Exit Sub
        If .NroMascotas > 0 Then
       If .MascotasType(.NroMascotas) = 111 And Hechizos(SpellIndex).NumNpc <> 111 Then Exit Sub
      If .MascotasType(.NroMascotas) = 111 And Hechizos(SpellIndex).NumNpc = 111 Then Exit Sub
    End If
    End If
         If Hechizos(SpellIndex).NumNpc = 111 And .NroMascotas >= 1 Then Exit Sub

'  .NroMascotas = 0
         
  '   If Hechizos(SpellIndex).NumNpc = ELEMENTALFUEGO And Npclist(PetIndex).Numero = 89 Then Exit Sub
        For NroNpcs = 1 To Hechizos(SpellIndex).cant
            
            If .NroMascotas < MAXMASCOTAS Then
                NpcIndex = SpawnNpc(Hechizos(SpellIndex).NumNpc, TargetPos, True, False)
                If NpcIndex > 0 Then
                    .NroMascotas = .NroMascotas + 1
                    
                    PetIndex = FreeMascotaIndex(Userindex)
                    
                    .MascotasIndex(PetIndex) = NpcIndex
                    .MascotasType(PetIndex) = Npclist(NpcIndex).Numero
                    
                    With Npclist(NpcIndex)
                        .MaestroUser = Userindex
                        .Contadores.TiempoExistencia = IntervaloInvocacion
                        .GiveGLD = 0
                    End With
                    
                    Call FollowAmo(NpcIndex)
                Else
                    Exit Sub
                End If
            Else
                Exit For
            End If
        
        Next NroNpcs
    End If
End With

Call InfoHechizo(Userindex)
HechizoCasteado = True

Exit Sub

error:
    With UserList(Userindex)
        LogError ("[" & Err.Number & "] " & Err.description & " por el usuario " & .Name & "(" & Userindex & _
                ") en (" & .Pos.Map & ", " & .Pos.X & ", " & .Pos.Y & "). Tratando de tirar el hechizo " & _
                Hechizos(SpellIndex).Nombre & "(" & SpellIndex & ") en la posicion ( " & .flags.TargetX & ", " & .flags.TargetY & ")")
    End With

End Sub
Sub HandleHechizoTerreno(ByVal Userindex As Integer, ByVal SpellIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 18/11/2009
'18/11/2009: ZaMa - Optimizacion de codigo.
'***************************************************

    If Not UserList(Userindex).flags.ModoCombate Then
    Call WriteConsoleMsg(Userindex, "Debes estar en modo de combate para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
    End If
     
    Dim HechizoCasteado As Boolean
    Dim ManaRequerida As Integer
    
    Select Case Hechizos(SpellIndex).tipo
        Case TipoHechizo.uInvocacion
            Call HechizoInvocacion(Userindex, HechizoCasteado)
            
        Case TipoHechizo.uEstado
            Call HechizoTerrenoEstado(Userindex, HechizoCasteado)
    End Select

       If HechizoCasteado Then
        With UserList(Userindex)
            Call SubirSkill(Userindex, eSkill.Magia, True)
            
            ManaRequerida = Hechizos(SpellIndex).ManaRequerido
            
            If Hechizos(SpellIndex).Warp = 1 Then ' Invocó una mascota
            ' Consume toda la mana
                ManaRequerida = 1000
            End If
            
            ' Quito la mana requerida
            .Stats.MinMAN = .Stats.MinMAN - ManaRequerida
            If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
            
            ' Quito la estamina requerida
            .Stats.MinSta = .Stats.MinSta - Hechizos(SpellIndex).StaRequerido
            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
            ' Update user stats
            Call WriteUpdateUserStats(Userindex)
        End With
    End If
    
End Sub

Sub HandleHechizoUsuario(ByVal Userindex As Integer, ByVal SpellIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2010
'18/11/2009: ZaMa - Optimizacion de codigo.
'12/01/2010: ZaMa - Optimizacion y agrego bonificaciones al druida.
'***************************************************
    
    If Not UserList(Userindex).flags.ModoCombate Then
    Call WriteConsoleMsg(Userindex, "Debes estar en modo de combate para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
    End If
    
    Dim HechizoCasteado As Boolean
    Dim ManaRequerida As Integer
    
    Select Case Hechizos(SpellIndex).tipo
        Case TipoHechizo.uEstado
            ' Afectan estados (por ejem : Envenenamiento)
            Call HechizoEstadoUsuario(Userindex, HechizoCasteado)
        
        Case TipoHechizo.uPropiedades
            ' Afectan HP,MANA,STAMINA,ETC
            HechizoCasteado = HechizoPropUsuario(Userindex)
    End Select

    If HechizoCasteado Then
        With UserList(Userindex)
            Call SubirSkill(Userindex, eSkill.Magia, True)
            
            ManaRequerida = Hechizos(SpellIndex).ManaRequerido
            
            ' Bonificaciones para druida
            If .clase = eClass.Druid Then
                ' Solo con flauta magica
                If .Invent.MunicionEqpObjIndex = FLAUTAELFICA And .Invent.MunicionEqpObjIndex = FLAUTAANTIGUA And UserList(Userindex).Invent.MunicionEqpObjIndex = AnilloBronce And UserList(Userindex).Invent.MunicionEqpObjIndex = AnilloPlata Then
                    If Hechizos(SpellIndex).Mimetiza = 1 Then
                        ' 50% menos de mana para mimetismo
                        ManaRequerida = ManaRequerida * 0.5
                        
                    ElseIf SpellIndex <> APOCALIPSIS_SPELL_INDEX Then
                        ' 10% menos de mana para todo menos apoca y descarga
                        ManaRequerida = ManaRequerida * 0.9
                    End If
                End If
            End If
            
            ' Quito la mana requerida
            .Stats.MinMAN = .Stats.MinMAN - ManaRequerida
            If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
            
            ' Quito la estamina requerida
            .Stats.MinSta = .Stats.MinSta - Hechizos(SpellIndex).StaRequerido
            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
            ' Update user stats
            Call WriteUpdateUserStats(Userindex)
            Call WriteUpdateUserStats(.flags.TargetUser)
            .flags.TargetUser = 0
        End With
    End If

End Sub

Sub HandleHechizoNPC(ByVal Userindex As Integer, ByVal HechizoIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2010
'13/02/2009: ZaMa - Agregada 50% bonificacion en coste de mana a mimetismo para druidas
'17/11/2009: ZaMa - Optimizacion de codigo.
'12/01/2010: ZaMa - Bonificacion para druidas de 10% para todos hechizos excepto apoca y descarga.
'12/01/2010: ZaMa - Los druidas mimetizados con npcs ahora son ignorados.
'***************************************************
    Dim HechizoCasteado As Boolean
    Dim ManaRequerida As Long
    
    With UserList(Userindex)
        Select Case Hechizos(HechizoIndex).tipo
            Case TipoHechizo.uEstado
                ' Afectan estados (por ejem : Envenenamiento)
                Call HechizoEstadoNPC(.flags.TargetNPC, HechizoIndex, HechizoCasteado, Userindex)
                
            Case TipoHechizo.uPropiedades
                ' Afectan HP,MANA,STAMINA,ETC
                Call HechizoPropNPC(HechizoIndex, .flags.TargetNPC, Userindex, HechizoCasteado)
        End Select
        
        
        If HechizoCasteado Then
            Call SubirSkill(Userindex, eSkill.Magia, True)
            
            ManaRequerida = Hechizos(HechizoIndex).ManaRequerido
            
            ' Bonificación para druidas.
            If .clase = eClass.Druid Then
                ' Se mostró como usuario, puede ser atacado por npcs
                .flags.Ignorado = False
                
                ' Solo con flauta equipada
                If .Invent.MunicionEqpObjIndex = FLAUTAELFICA And .Invent.MunicionEqpObjIndex = FLAUTAANTIGUA And UserList(Userindex).Invent.MunicionEqpObjIndex = AnilloBronce And UserList(Userindex).Invent.MunicionEqpObjIndex = AnilloPlata Then
                    If Hechizos(HechizoIndex).Mimetiza = 1 Then
                        ' 50% menos de mana para mimetismo
                        ManaRequerida = ManaRequerida * 0.5
                        ' Será ignorado hasta que pierda el efecto del mimetismo o ataque un npc
                        .flags.Ignorado = True
                    Else
                        ' 10% menos de mana para hechizos
                        If HechizoIndex <> APOCALIPSIS_SPELL_INDEX Then
                             ManaRequerida = ManaRequerida * 0.9
                        End If
                    End If
                End If
            End If
            
            ' Quito la mana requerida
            .Stats.MinMAN = .Stats.MinMAN - ManaRequerida
            If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
            
            ' Quito la estamina requerida
            .Stats.MinSta = .Stats.MinSta - Hechizos(HechizoIndex).StaRequerido
            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
            ' Update user stats
            Call WriteUpdateUserStats(Userindex)
            .flags.TargetNPC = 0
        End If
    End With
End Sub


Sub LanzarHechizo(ByVal SpellIndex As Integer, ByVal Userindex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 02/16/2010
'24/01/2007 ZaMa - Optimizacion de codigo.
'02/16/2010: Marco - Now .flags.hechizo makes reference to global spell index instead of user's spell index
'***************************************************
On Error GoTo Errhandler

With UserList(Userindex)
    
 If Hechizos(SpellIndex).Nombre = "Implorar Ayuda" And .clase <> eClass.Druid Then
WriteConsoleMsg Userindex, "No eres un druida para invocar este hechizos!", FontTypeNames.FONTTYPE_INFO
Exit Sub
End If
    
    
      If .death = True And DeathMatch.Cuenta > 0 Then
            WriteConsoleMsg Userindex, "¡No puedes atacar antes de la cuenta regresiva!", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
    If (DeathMatch.Ingresaron < DeathMatch.Cupos And .death = True) Then
    WriteConsoleMsg Userindex, "No puedes atacar si no se llenaron los cupos", FontTypeNames.FONTTYPE_WARNING
    Exit Sub
    End If
    
                    If .hungry = True And JDH.Cuenta > 0 Then
            WriteConsoleMsg Userindex, "¡No puedes atacar antes de la cuenta regresiva!", FontTypeNames.FONTTYPE_INFO
            Exit Sub
        End If
        
    If (JDH.Ingresaron < JDH.Cupos And .hungry = True) Then
    WriteConsoleMsg Userindex, "No puedes atacar si no se llenaron los cupos", FontTypeNames.FONTTYPE_WARNING
    Exit Sub
    End If
    
    If .flags.EnConsulta Then
        Call WriteConsoleMsg(Userindex, "No puedes lanzar hechizos si estás en consulta.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If PuedeLanzar(Userindex, SpellIndex) Then
        Select Case Hechizos(SpellIndex).Target
            Case TargetType.uUsuarios
                If .flags.TargetUser > 0 Then
                    If Abs(UserList(.flags.TargetUser).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        Call HandleHechizoUsuario(Userindex, SpellIndex)
                    Else
                        Call WriteConsoleMsg(Userindex, "Estás demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "Este hechizo actúa sólo sobre usuarios.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case TargetType.uNPC
                If .flags.TargetNPC > 0 Then
                    If Abs(Npclist(.flags.TargetNPC).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        Call HandleHechizoNPC(Userindex, SpellIndex)
                    Else
                        Call WriteConsoleMsg(Userindex, "Estás demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "Este hechizo sólo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case TargetType.uUsuariosYnpc
                If .flags.TargetUser > 0 Then
                    If Abs(UserList(.flags.TargetUser).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        Call HandleHechizoUsuario(Userindex, SpellIndex)
                    Else
                        Call WriteConsoleMsg(Userindex, "Estás demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                    End If
                ElseIf .flags.TargetNPC > 0 Then
                    If Abs(Npclist(.flags.TargetNPC).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        Call HandleHechizoNPC(Userindex, SpellIndex)
                    Else
                        Call WriteConsoleMsg(Userindex, "Estás demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
                    End If
                Else
                    Call WriteConsoleMsg(Userindex, "Target inválido.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case TargetType.uTerreno
                Call HandleHechizoTerreno(Userindex, SpellIndex)
        End Select
        
    End If
    
    If .Counters.Trabajando Then _
        .Counters.Trabajando = .Counters.Trabajando - 1
    
    If .Counters.Ocultando Then _
        .Counters.Ocultando = .Counters.Ocultando - 1

End With

Exit Sub

Errhandler:
    Call LogError("Error en LanzarHechizo. Error " & Err.Number & " : " & Err.description & _
        " Hechizo: " & Hechizos(SpellIndex).Nombre & "(" & SpellIndex & _
        "). Casteado por: " & UserList(Userindex).Name & "(" & Userindex & ").")
    
End Sub

Sub HechizoEstadoUsuario(ByVal Userindex As Integer, ByRef HechizoCasteado As Boolean)
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
'***************************************************


Dim HechizoIndex As Integer
Dim TargetIndex As Integer

With UserList(Userindex)
    HechizoIndex = .flags.Hechizo
    TargetIndex = .flags.TargetUser
    
    ' <-------- Agrega Invisibilidad ---------->
    If Hechizos(HechizoIndex).Invisibilidad = 1 Then
        If UserList(TargetIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡El usuario está muerto!", FontTypeNames.FONTTYPE_INFO)
            HechizoCasteado = False
            Exit Sub
        End If
        
        If UserList(TargetIndex).Counters.Saliendo Then
            If Userindex <> TargetIndex Then
                Call WriteConsoleMsg(Userindex, "¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
                Exit Sub
            Else
                Call WriteConsoleMsg(Userindex, "¡No puedes hacerte invisible mientras te encuentras saliendo!", FontTypeNames.FONTTYPE_WARNING)
                HechizoCasteado = False
                Exit Sub
            End If
        End If
        
        'No usar invi mapas InviSinEfecto
        If MapInfo(UserList(TargetIndex).Pos.Map).InviSinEfecto > 0 Then
            Call WriteConsoleMsg(Userindex, "¡La invisibilidad no funciona aquí!", FontTypeNames.FONTTYPE_INFO)
            HechizoCasteado = False
            Exit Sub
        End If
        
        ' Chequea si el status permite ayudar al otro usuario
        HechizoCasteado = CanSupportUser(Userindex, TargetIndex, True)
        If Not HechizoCasteado Then Exit Sub
        
        'Si sos user, no uses este hechizo con GMS.
        If .flags.Privilegios And PlayerType.User Then
            If Not UserList(TargetIndex).flags.Privilegios And PlayerType.User Then
                HechizoCasteado = False
                Exit Sub
            End If
        End If
       
        UserList(TargetIndex).flags.invisible = 1
        Call SetInvisible(TargetIndex, UserList(TargetIndex).Char.CharIndex, True)
    
        Call InfoHechizo(Userindex)
        HechizoCasteado = True
    End If
    
    ' <-------- Agrega Mimetismo ---------->
    If Hechizos(HechizoIndex).Mimetiza = 1 Then
        If UserList(TargetIndex).flags.Muerto = 1 Then
            Exit Sub
        End If
        
        If UserList(TargetIndex).flags.Navegando = 1 Then
            Exit Sub
        End If
        If .flags.Navegando = 1 Then
            Exit Sub
        End If
        
        'Si sos user, no uses este hechizo con GMS.
        If .flags.Privilegios And PlayerType.User Then
            If Not UserList(TargetIndex).flags.Privilegios And PlayerType.User Then
                Exit Sub
            End If
        End If
        
        If .flags.Mimetizado = 1 Then
            Call WriteConsoleMsg(Userindex, "Ya te encuentras mimetizado. El hechizo no ha tenido efecto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.AdminInvisible = 1 Then Exit Sub
        
        'copio el char original al mimetizado
        
        .CharMimetizado.body = .Char.body
        .CharMimetizado.Head = .Char.Head
        .CharMimetizado.CascoAnim = .Char.CascoAnim
        .CharMimetizado.ShieldAnim = .Char.ShieldAnim
        .CharMimetizado.WeaponAnim = .Char.WeaponAnim
        
        .flags.Mimetizado = 1
        
        'ahora pongo local el del enemigo
        .Char.body = UserList(TargetIndex).Char.body
        .Char.Head = UserList(TargetIndex).Char.Head
        .Char.CascoAnim = UserList(TargetIndex).Char.CascoAnim
        .Char.ShieldAnim = UserList(TargetIndex).Char.ShieldAnim
        .Char.WeaponAnim = GetWeaponAnim(Userindex, UserList(TargetIndex).Invent.WeaponEqpObjIndex)
        
        Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
       
       Call InfoHechizo(Userindex)
       HechizoCasteado = True
    End If
    
    ' <-------- Agrega Envenenamiento ---------->
    If Hechizos(HechizoIndex).Envenena = 1 Then
        If Userindex = TargetIndex Then
            Call WriteConsoleMsg(Userindex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
        If Not PuedeAtacar(Userindex, TargetIndex) Then Exit Sub
        If Userindex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(Userindex, TargetIndex)
        End If
        UserList(TargetIndex).flags.Envenenado = 1
        Call InfoHechizo(Userindex)
        HechizoCasteado = True
    End If
    
    ' <-------- Cura Envenenamiento ---------->
    If Hechizos(HechizoIndex).CuraVeneno = 1 Then
    
        'Verificamos que el usuario no este muerto
        If UserList(TargetIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡El usuario está muerto!", FontTypeNames.FONTTYPE_INFO)
            HechizoCasteado = False
            Exit Sub
        End If
        
                If UserList(TargetIndex).flags.Envenenado = 0 Then
            Call WriteConsoleMsg(Userindex, "¡El usuario no está envenenado!", FontTypeNames.FONTTYPE_INFO)
            HechizoCasteado = False
            Exit Sub
        End If
        
        ' Chequea si el status permite ayudar al otro usuario
        HechizoCasteado = CanSupportUser(Userindex, TargetIndex)
        If Not HechizoCasteado Then Exit Sub
            
        'Si sos user, no uses este hechizo con GMS.
        If .flags.Privilegios And PlayerType.User Then
            If Not UserList(TargetIndex).flags.Privilegios And PlayerType.User Then
                Exit Sub
            End If
        End If
            
        UserList(TargetIndex).flags.Envenenado = 0
        Call InfoHechizo(Userindex)
        HechizoCasteado = True
    End If
    
    ' <-------- Agrega Maldicion ---------->
    If Hechizos(HechizoIndex).Maldicion = 1 Then
        If Userindex = TargetIndex Then
            Call WriteConsoleMsg(Userindex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
        If Not PuedeAtacar(Userindex, TargetIndex) Then Exit Sub
        If Userindex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(Userindex, TargetIndex)
        End If
        UserList(TargetIndex).flags.Maldicion = 1
        Call InfoHechizo(Userindex)
        HechizoCasteado = True
    End If
    
    ' <-------- Remueve Maldicion ---------->
    If Hechizos(HechizoIndex).RemoverMaldicion = 1 Then
            UserList(TargetIndex).flags.Maldicion = 0
            Call InfoHechizo(Userindex)
            HechizoCasteado = True
    End If
    
    ' <-------- Agrega Bendicion ---------->
    If Hechizos(HechizoIndex).Bendicion = 1 Then
            UserList(TargetIndex).flags.Bendicion = 1
            Call InfoHechizo(Userindex)
            HechizoCasteado = True
    End If
    
    ' <-------- Agrega Paralisis/Inmobilidad ---------->
    If Hechizos(HechizoIndex).Paraliza = 1 Or Hechizos(HechizoIndex).Inmoviliza = 1 Then
        If Userindex = TargetIndex Then
            Call WriteConsoleMsg(Userindex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
         If UserList(TargetIndex).flags.Paralizado = 0 Then
            If Not PuedeAtacar(Userindex, TargetIndex) Then Exit Sub
            
            If Userindex <> TargetIndex Then
                Call UsuarioAtacadoPorUsuario(Userindex, TargetIndex)
            End If
            
            Call InfoHechizo(Userindex)
            HechizoCasteado = True
            If UserList(TargetIndex).Invent.AnilloEqpObjIndex = SUPERANILLO Then
                Call WriteConsoleMsg(TargetIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(Userindex, " ¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
                Call FlushBuffer(TargetIndex)
                Exit Sub
            End If
            
            If Hechizos(HechizoIndex).Inmoviliza = 1 Then UserList(TargetIndex).flags.Inmovilizado = 1
            UserList(TargetIndex).flags.Paralizado = 1
            UserList(TargetIndex).Counters.Paralisis = IntervaloParalizado
            
            Call WriteParalizeOK(TargetIndex)
            Call FlushBuffer(TargetIndex)
        End If
    End If
    
    ' <-------- Remueve Paralisis/Inmobilidad ---------->
    If Hechizos(HechizoIndex).RemoverParalisis = 1 Then
        
        ' Remueve si esta en ese estado
        If UserList(TargetIndex).flags.Paralizado = 1 Then
        
            ' Chequea si el status permite ayudar al otro usuario
            HechizoCasteado = CanSupportUser(Userindex, TargetIndex, True)
            If Not HechizoCasteado Then Exit Sub
            
            UserList(TargetIndex).flags.Inmovilizado = 0
            UserList(TargetIndex).flags.Paralizado = 0
            
            'no need to crypt this
            Call WriteParalizeOK(TargetIndex)
            Call InfoHechizo(Userindex)
        
        End If
    End If
    
    ' <-------- Remueve Estupidez (Aturdimiento) ---------->
    If Hechizos(HechizoIndex).RemoverEstupidez = 1 Then
    
        ' Remueve si esta en ese estado
        If UserList(TargetIndex).flags.Estupidez = 1 Then
        
            ' Chequea si el status permite ayudar al otro usuario
            HechizoCasteado = CanSupportUser(Userindex, TargetIndex)
            If Not HechizoCasteado Then Exit Sub
        
            UserList(TargetIndex).flags.Estupidez = 0
            
            'no need to crypt this
            Call WriteDumbNoMore(TargetIndex)
            Call FlushBuffer(TargetIndex)
            Call InfoHechizo(Userindex)
        
        End If
    End If
    
    ' <-------- Revive ---------->
    If Hechizos(HechizoIndex).Revivir = 1 Then
        If UserList(TargetIndex).flags.Muerto = 1 Then
            
            'Seguro de resurreccion (solo afecta a los hechizos, no al sacerdote ni al comando de GM)
            If UserList(TargetIndex).flags.ModoCombate Then
                Call WriteConsoleMsg(Userindex, "El usuario esta en Modo Combate. No puedes revivirlo.", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
                Exit Sub
            End If
        
            'No usar resu en mapas con ResuSinEfecto
            If MapInfo(UserList(TargetIndex).Pos.Map).ResuSinEfecto > 0 Then
                Call WriteConsoleMsg(Userindex, "¡Revivir no está permitido aquí! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
                Exit Sub
            End If
            
           'No podemos resucitar si nuestra barra de energía no está llena. (GD: 29/04/07)
            If .Stats.MinSta = 500 Then
                Call WriteConsoleMsg(Userindex, "No puedes resucitar si no tienes 500 puntos de energía.", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
                Exit Sub
            End If
            
            
        'revisamos si necesita vara
            If .clase = eClass.Mage Then
                If .Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(.Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                        Call WriteConsoleMsg(Userindex, "Necesitas un báculo mejor para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
                        HechizoCasteado = False
                        Exit Sub
                    End If
                End If
            ElseIf .clase = eClass.Bard Then
                If .Invent.MunicionEqpObjIndex <> LAUDELFICO And .Invent.MunicionEqpObjIndex <> LAUDSUPERMAGICO <> LaudBronce <> LaudPlata Then
                    Call WriteConsoleMsg(Userindex, "Necesitas un instrumento mágico para devolver la vida.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                    Exit Sub
                End If
            ElseIf .clase = eClass.Druid Then
                If .Invent.MunicionEqpObjIndex <> FLAUTAELFICA And .Invent.MunicionEqpObjIndex <> FLAUTAANTIGUA And .Invent.MunicionEqpObjIndex <> AnilloBronce And .Invent.MunicionEqpObjIndex <> AnilloPlata Then
                    Call WriteConsoleMsg(Userindex, "Necesitas un instrumento mágico para devolver la vida.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                    Exit Sub
                End If
            End If
            
            ' Chequea si el status permite ayudar al otro usuario
            HechizoCasteado = CanSupportUser(Userindex, TargetIndex, True)
            If Not HechizoCasteado Then Exit Sub
    
            Dim EraCriminal As Boolean
            EraCriminal = criminal(Userindex)
            
            If Not criminal(TargetIndex) Then
                If TargetIndex <> Userindex Then
                    .Reputacion.NobleRep = .Reputacion.NobleRep + 500
                    If .Reputacion.NobleRep > MAXREP Then _
                        .Reputacion.NobleRep = MAXREP
                    Call WriteConsoleMsg(Userindex, "¡Los Dioses te sonríen, has ganado 500 puntos de nobleza!", FontTypeNames.FONTTYPE_INFO)
                End If
            End If
            
            If EraCriminal And Not criminal(Userindex) Then
                Call RefreshCharStatus(Userindex)
            End If
            
            With UserList(TargetIndex)
                'Pablo Toxic Waste (GD: 29/04/07)
                .Stats.MinAGU = 0
                .flags.Sed = 1
                .Stats.MinHam = 0
                .flags.Hambre = 1
                Call WriteUpdateHungerAndThirst(TargetIndex)
                Call InfoHechizo(Userindex)
                .Stats.MinMAN = 0
                .Stats.MinSta = 0
            End With
            
            'Agregado para quitar la penalización de vida en el ring y cambio de ecuacion. (NicoNZ)
            If (TriggerZonaPelea(Userindex, TargetIndex) <> TRIGGER6_PERMITE) Then
                'Solo saco vida si es User. no quiero que exploten GMs por ahi.
                If .flags.Privilegios And PlayerType.User Then
                    .Stats.MinHp = .Stats.MinHp * (1 - UserList(TargetIndex).Stats.ELV * 0.015)
                End If
            End If
            
            If (.Stats.MinHp <= 0) Then
                Call UserDie(Userindex)
                'Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar fue demasiado grande.", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
            Else
                'Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar te ha debilitado.", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = True
            End If
            
            If UserList(TargetIndex).flags.Traveling = 1 Then
                UserList(TargetIndex).Counters.goHome = 0
                UserList(TargetIndex).flags.Traveling = 0
                'Call WriteConsoleMsg(TargetIndex, "Tu viaje ha sido cancelado.", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteMultiMessage(TargetIndex, eMessages.CancelHome)
            End If
            
            Call RevivirUsuario(TargetIndex)
        Else
            HechizoCasteado = False
        End If
    
    End If
    
    ' <-------- Agrega Ceguera ---------->
    If Hechizos(HechizoIndex).Ceguera = 1 Then
        If Userindex = TargetIndex Then
            Call WriteConsoleMsg(Userindex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
            If Not PuedeAtacar(Userindex, TargetIndex) Then Exit Sub
            If Userindex <> TargetIndex Then
                Call UsuarioAtacadoPorUsuario(Userindex, TargetIndex)
            End If
            UserList(TargetIndex).flags.Ceguera = 1
            UserList(TargetIndex).Counters.Ceguera = IntervaloParalizado / 3
    
            Call WriteBlind(TargetIndex)
            Call FlushBuffer(TargetIndex)
            Call InfoHechizo(Userindex)
            HechizoCasteado = True
    End If
    
    ' <-------- Agrega Estupidez (Aturdimiento) ---------->
    If Hechizos(HechizoIndex).Estupidez = 1 Then
        If Userindex = TargetIndex Then
            Call WriteConsoleMsg(Userindex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
            If Not PuedeAtacar(Userindex, TargetIndex) Then Exit Sub
            If Userindex <> TargetIndex Then
                Call UsuarioAtacadoPorUsuario(Userindex, TargetIndex)
            End If
            If UserList(TargetIndex).flags.Estupidez = 0 Then
                UserList(TargetIndex).flags.Estupidez = 1
                UserList(TargetIndex).Counters.Ceguera = IntervaloParalizado
            End If
            Call WriteDumb(TargetIndex)
            Call FlushBuffer(TargetIndex)
    
            Call InfoHechizo(Userindex)
            HechizoCasteado = True
    End If
End With

End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal SpellIndex As Integer, ByRef HechizoCasteado As Boolean, ByVal Userindex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 07/07/2008
'Handles the Spells that afect the Stats of an NPC
'04/13/2008 NicoNZ - Guardias Faccionarios pueden ser
'removidos por users de su misma faccion.
'07/07/2008: NicoNZ - Solo se puede mimetizar con npcs si es druida
'***************************************************

With Npclist(NpcIndex)
    If Hechizos(SpellIndex).Invisibilidad = 1 Then
        Call InfoHechizo(Userindex)
        .flags.invisible = 1
        HechizoCasteado = True
    End If
    
    If Hechizos(SpellIndex).Envenena = 1 Then
        If Not PuedeAtacarNPC(Userindex, NpcIndex) Then
            HechizoCasteado = False
            Exit Sub
        End If
        Call NPCAtacado(NpcIndex, Userindex)
        Call InfoHechizo(Userindex)
        .flags.Envenenado = 1
        HechizoCasteado = True
    End If
    
    If Hechizos(SpellIndex).CuraVeneno = 1 Then
        Call InfoHechizo(Userindex)
        .flags.Envenenado = 0
        HechizoCasteado = True
    End If
    
    If Hechizos(SpellIndex).Maldicion = 1 Then
        If Not PuedeAtacarNPC(Userindex, NpcIndex) Then
            HechizoCasteado = False
            Exit Sub
        End If
        Call NPCAtacado(NpcIndex, Userindex)
        Call InfoHechizo(Userindex)
        .flags.Maldicion = 1
        HechizoCasteado = True
    End If
    
    If Hechizos(SpellIndex).RemoverMaldicion = 1 Then
        Call InfoHechizo(Userindex)
        .flags.Maldicion = 0
        HechizoCasteado = True
    End If
    
    If Hechizos(SpellIndex).Bendicion = 1 Then
        Call InfoHechizo(Userindex)
        .flags.Bendicion = 1
        HechizoCasteado = True
    End If
    
    If Hechizos(SpellIndex).Paraliza = 1 Then
        If .flags.AfectaParalisis = 0 Then
            If Not PuedeAtacarNPC(Userindex, NpcIndex, True) Then
                HechizoCasteado = False
                Exit Sub
            End If
            Call NPCAtacado(NpcIndex, Userindex)
            Call InfoHechizo(Userindex)
            .flags.Paralizado = 1
            .flags.Inmovilizado = 0
            .Contadores.Paralisis = IntervaloParalizado
            HechizoCasteado = True
        Else
            Call WriteConsoleMsg(Userindex, "El NPC es inmune a este hechizo.", FontTypeNames.FONTTYPE_INFO)
            HechizoCasteado = False
            Exit Sub
        End If
    End If
    
    If Hechizos(SpellIndex).RemoverParalisis = 1 Then
        If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
            If .MaestroUser = Userindex Then
                Call InfoHechizo(Userindex)
                .flags.Paralizado = 0
                .Contadores.Paralisis = 0
                HechizoCasteado = True
            Else
                If .NPCtype = eNPCType.GuardiaReal Then
                    If esArmada(Userindex) Then
                        Call InfoHechizo(Userindex)
                        .flags.Paralizado = 0
                        .Contadores.Paralisis = 0
                        HechizoCasteado = True
                        Exit Sub
                    Else
                        Call WriteConsoleMsg(Userindex, "Sólo puedes remover la parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
                        HechizoCasteado = False
                        Exit Sub
                    End If
                    
                    Call WriteConsoleMsg(Userindex, "Solo puedes remover la parálisis de los NPCs que te consideren su amo.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                    Exit Sub
                Else
                    If .NPCtype = eNPCType.Guardiascaos Then
                        If esCaos(Userindex) Then
                            Call InfoHechizo(Userindex)
                            .flags.Paralizado = 0
                            .Contadores.Paralisis = 0
                            HechizoCasteado = True
                            Exit Sub
                        Else
                            Call WriteConsoleMsg(Userindex, "Solo puedes remover la parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
                            HechizoCasteado = False
                            Exit Sub
                        End If
                    End If
                End If
            End If
       Else
          Call WriteConsoleMsg(Userindex, "Este NPC no está paralizado", FontTypeNames.FONTTYPE_INFO)
          HechizoCasteado = False
          Exit Sub
       End If
    End If
     
    If Hechizos(SpellIndex).Inmoviliza = 1 Then
        If .flags.AfectaParalisis = 0 Then
            If Not PuedeAtacarNPC(Userindex, NpcIndex, True) Then
                HechizoCasteado = False
                Exit Sub
            End If
            Call NPCAtacado(NpcIndex, Userindex)
            .flags.Inmovilizado = 1
            .flags.Paralizado = 0
            .Contadores.Paralisis = IntervaloParalizado
            Call InfoHechizo(Userindex)
            HechizoCasteado = True
        Else
            Call WriteConsoleMsg(Userindex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
        End If
    End If
End With

If Hechizos(SpellIndex).Mimetiza = 1 Then
    With UserList(Userindex)
        If .flags.Mimetizado = 1 Then
            Call WriteConsoleMsg(Userindex, "Ya te encuentras mimetizado. El hechizo no ha tenido efecto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.AdminInvisible = 1 Then Exit Sub
        
            
        If .clase = eClass.Druid Then
            'copio el char original al mimetizado
            
            .CharMimetizado.body = .Char.body
            .CharMimetizado.Head = .Char.Head
            .CharMimetizado.CascoAnim = .Char.CascoAnim
            .CharMimetizado.ShieldAnim = .Char.ShieldAnim
            .CharMimetizado.WeaponAnim = .Char.WeaponAnim
            
            .flags.Mimetizado = 1
            
            'ahora pongo lo del NPC.
            .Char.body = Npclist(NpcIndex).Char.body
            .Char.Head = Npclist(NpcIndex).Char.Head
            .Char.CascoAnim = NingunCasco
            .Char.ShieldAnim = NingunEscudo
            .Char.WeaponAnim = NingunArma
        
            Call ChangeUserChar(Userindex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
            
        Else
            Call WriteConsoleMsg(Userindex, "Sólo los druidas pueden mimetizarse con criaturas.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
       Call InfoHechizo(Userindex)
       HechizoCasteado = True
    End With
End If

End Sub

Sub HechizoPropNPC(ByVal SpellIndex As Integer, ByVal NpcIndex As Integer, ByVal Userindex As Integer, ByRef HechizoCasteado As Boolean)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 14/08/2007
'Handles the Spells that afect the Life NPC
'14/08/2007 Pablo (ToxicWaste) - Orden general.
'***************************************************

Dim daño As Long

With Npclist(NpcIndex)
    
    If .flags.Paralizado And Hechizos(SpellIndex).Inmoviliza Then
HechizoCasteado = False
End If

If UserList(Userindex).flags.Oculto = 1 Then
DecirPalabrasMagicas Hechizos(SpellIndex).PalabrasMagicas, Userindex
UserList(Userindex).flags.Oculto = 0
SetInvisible Userindex, UserList(Userindex).Char.CharIndex, False
End If

    'Salud
    
    
    If Hechizos(SpellIndex).SubeHP = 1 Then

        If MapInfo(.Pos.Map).Pk = False Then
  Call WriteConsoleMsg(Userindex, "¡No puedes curar a este npc!", FontTypeNames.FONTTYPE_INFO)
   Exit Sub
   End If
   
            If Npclist(NpcIndex).Stats.MinHp = .Stats.MaxHp Then
    Call WriteConsoleMsg(Userindex, "¡La criatura no está herida!", FontTypeNames.FONTTYPE_FIGHT)
    Exit Sub
    End If

        daño = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
        daño = daño + Porcentaje(daño, 3 * UserList(Userindex).Stats.ELV)
        
        Call InfoHechizo(Userindex)
        .Stats.MinHp = .Stats.MinHp + daño
        If .Stats.MinHp > .Stats.MaxHp Then _
            .Stats.MinHp = .Stats.MaxHp
        Call WriteConsoleMsg(Userindex, "Has curado " & daño & " puntos de vida a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
        HechizoCasteado = True
        
    ElseIf Hechizos(SpellIndex).SubeHP = 2 Then

        If Not PuedeAtacarNPC(Userindex, NpcIndex) Then
            HechizoCasteado = False
            Exit Sub
        End If
        Call NPCAtacado(NpcIndex, Userindex)
        daño = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
        daño = daño + Porcentaje(daño, 3 * UserList(Userindex).Stats.ELV)
    
        If Hechizos(SpellIndex).StaffAffected Then
            If UserList(Userindex).clase = eClass.Mage Then
                If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
                    daño = (daño * (ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                    'Aumenta daño segun el staff-
                    'Daño = (Daño* (70 + BonifBáculo)) / 100
                Else
                    daño = daño * 0.7 'Baja daño a 70% del original
                End If
            End If
        End If
          If UserList(Userindex).Invent.MunicionEqpObjIndex = LAUDELFICO Or UserList(Userindex).Invent.MunicionEqpObjIndex = FLAUTAELFICA Then
            daño = daño * 1.04  'laud magico de los bardos
        End If
                  If UserList(Userindex).Invent.MunicionEqpObjIndex = LAUDSUPERMAGICO Or UserList(Userindex).Invent.MunicionEqpObjIndex = FLAUTAANTIGUA Then
            daño = daño * 1.09  'laud magico de los bardos
        End If
        
        
    
        Call InfoHechizo(Userindex)
        HechizoCasteado = True
        
        If .flags.Snd2 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(.flags.Snd2, .Pos.X, .Pos.Y))
        End If
        
        'Quizas tenga defenza magica el NPC. Pablo (ToxicWaste)
        daño = daño - .Stats.defM
        If daño < 0 Then daño = 0
        
        .Stats.MinHp = .Stats.MinHp - daño
        SendData SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, DAMAGE_NORMAL)
        Call WriteConsoleMsg(Userindex, "¡Le has quitado " & daño & " puntos de vida a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
        Call CalcularDarExp(Userindex, NpcIndex, daño)

        
        If .Stats.MinHp < 1 Then
            .Stats.MinHp = 0
            Call MuereNpc(NpcIndex, Userindex)
        End If
        End If
End With

End Sub

Sub InfoHechizo(ByVal Userindex As Integer)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 25/07/2009
'25/07/2009: ZaMa - Code improvements.
'25/07/2009: ZaMa - Now invisible admins magic sounds are not sent to anyone but themselves
'***************************************************
    Dim SpellIndex As Integer
    Dim tUser As Integer
    Dim tNpc As Integer
    
    With UserList(Userindex)
        SpellIndex = .flags.Hechizo
        tUser = .flags.TargetUser
        tNpc = .flags.TargetNPC
        
        Call DecirPalabrasMagicas(Hechizos(SpellIndex).PalabrasMagicas, Userindex)
        
        If tUser > 0 Then
            ' Los admins invisibles no producen sonidos ni fx's
            If .flags.AdminInvisible = 1 And Userindex = tUser Then
                Call EnviarDatosASlot(Userindex, PrepareMessageCreateFX(UserList(tUser).Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
                Call EnviarDatosASlot(Userindex, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, UserList(tUser).Pos.X, UserList(tUser).Pos.Y))
            Else
                Call SendData(SendTarget.ToPCArea, tUser, PrepareMessageCreateFX(UserList(tUser).Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
                Call SendData(SendTarget.ToPCArea, tUser, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, UserList(tUser).Pos.X, UserList(tUser).Pos.Y)) 'Esta linea faltaba. Pablo (ToxicWaste)
            End If
        ElseIf tNpc > 0 Then
            Call SendData(SendTarget.ToNPCArea, tNpc, PrepareMessageCreateFX(Npclist(tNpc).Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
            Call SendData(SendTarget.ToNPCArea, tNpc, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, Npclist(tNpc).Pos.X, Npclist(tNpc).Pos.Y))
        End If
        
        If tUser > 0 Then
            If Userindex <> tUser Then
                If .showName Then
                    Call WriteConsoleMsg(Userindex, Hechizos(SpellIndex).HechizeroMsg & " " & UserList(tUser).Name, FontTypeNames.FONTTYPE_FIGHT)
                Else
                    Call WriteConsoleMsg(Userindex, Hechizos(SpellIndex).HechizeroMsg & " alguien.", FontTypeNames.FONTTYPE_FIGHT)
                End If
                Call WriteConsoleMsg(tUser, .Name & " " & Hechizos(SpellIndex).TargetMsg, FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(Userindex, Hechizos(SpellIndex).PropioMsg, FontTypeNames.FONTTYPE_FIGHT)
            End If
        ElseIf tNpc > 0 Then
            Call WriteConsoleMsg(Userindex, Hechizos(SpellIndex).HechizeroMsg & " " & "la criatura.", FontTypeNames.FONTTYPE_FIGHT)
        End If
    End With

End Sub

Public Function HechizoPropUsuario(ByVal Userindex As Integer) As Boolean
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 28/04/2010
'02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
'28/04/2010: ZaMa - Agrego Restricciones para ciudas respecto al estado atacable.
'***************************************************

Dim SpellIndex As Integer
Dim daño As Long
Dim TargetIndex As Integer

SpellIndex = UserList(Userindex).flags.Hechizo
TargetIndex = UserList(Userindex).flags.TargetUser
      
With UserList(TargetIndex)
    If .flags.Muerto Then
        Call WriteConsoleMsg(Userindex, "No puedes lanzar este hechizo a un muerto.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
          
    ' <-------- Aumenta Hambre ---------->
    If Hechizos(SpellIndex).SubeHam = 1 Then
        
        Call InfoHechizo(Userindex)
        
        daño = RandomNumber(Hechizos(SpellIndex).MinHam, Hechizos(SpellIndex).MaxHam)
        
        .Stats.MinHam = .Stats.MinHam + daño
        If .Stats.MinHam > .Stats.MaxHam Then _
            .Stats.MinHam = .Stats.MaxHam
        
        If Userindex <> TargetIndex Then
            Call WriteConsoleMsg(Userindex, "Le has restaurado " & daño & " puntos de hambre a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(TargetIndex, UserList(Userindex).Name & " te ha restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(Userindex, "Te has restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
        End If
        
        Call WriteUpdateHungerAndThirst(TargetIndex)
    
    ' <-------- Quita Hambre ---------->
    ElseIf Hechizos(SpellIndex).SubeHam = 2 Then
        If Not PuedeAtacar(Userindex, TargetIndex) Then Exit Function
        
        If Userindex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(Userindex, TargetIndex)
        Else
            Exit Function
        End If
        
        Call InfoHechizo(Userindex)
        
        daño = RandomNumber(Hechizos(SpellIndex).MinHam, Hechizos(SpellIndex).MaxHam)
        
        .Stats.MinHam = .Stats.MinHam - daño
        
        If Userindex <> TargetIndex Then
            Call WriteConsoleMsg(Userindex, "Le has quitado " & daño & " puntos de hambre a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(TargetIndex, UserList(Userindex).Name & " te ha quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(Userindex, "Te has quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
        End If
        
        If .Stats.MinHam < 1 Then
            .Stats.MinHam = 0
            .flags.Hambre = 1
        End If
        
        Call WriteUpdateHungerAndThirst(TargetIndex)
    End If
    
    ' <-------- Aumenta Sed ---------->
    If Hechizos(SpellIndex).SubeSed = 1 Then
        
        Call InfoHechizo(Userindex)
        
        daño = RandomNumber(Hechizos(SpellIndex).MinSed, Hechizos(SpellIndex).MaxSed)
        
        .Stats.MinAGU = .Stats.MinAGU + daño
        If .Stats.MinAGU > .Stats.MaxAGU Then _
            .Stats.MinAGU = .Stats.MaxAGU
        
        Call WriteUpdateHungerAndThirst(TargetIndex)
             
        If Userindex <> TargetIndex Then
          Call WriteConsoleMsg(Userindex, "Le has restaurado " & daño & " puntos de sed a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
          Call WriteConsoleMsg(TargetIndex, UserList(Userindex).Name & " te ha restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
        Else
          Call WriteConsoleMsg(Userindex, "Te has restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
        End If
        
    
    ' <-------- Quita Sed ---------->
    ElseIf Hechizos(SpellIndex).SubeSed = 2 Then
        
        If Not PuedeAtacar(Userindex, TargetIndex) Then Exit Function
        
        If Userindex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(Userindex, TargetIndex)
        End If
        
        Call InfoHechizo(Userindex)
        
        daño = RandomNumber(Hechizos(SpellIndex).MinSed, Hechizos(SpellIndex).MaxSed)
        
        .Stats.MinAGU = .Stats.MinAGU - daño
        
        If Userindex <> TargetIndex Then
            Call WriteConsoleMsg(Userindex, "Le has quitado " & daño & " puntos de sed a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(TargetIndex, UserList(Userindex).Name & " te ha quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(Userindex, "Te has quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
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
        If Not CanSupportUser(Userindex, TargetIndex) Then Exit Function
        
        Call InfoHechizo(Userindex)
        daño = RandomNumber(Hechizos(SpellIndex).MinAgilidad, Hechizos(SpellIndex).MaxAgilidad)
        
        .flags.DuracionEfecto = 1200
        .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + daño
        If .Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Agilidad) * 2) Then _
            .Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Agilidad) * 2)
        
        .flags.TomoPocion = True
        Call WriteUpdateDexterity(TargetIndex)
    
    ' <-------- Quita Agilidad ---------->
    ElseIf Hechizos(SpellIndex).SubeAgilidad = 2 Then
        
        If Not PuedeAtacar(Userindex, TargetIndex) Then Exit Function
        
        If Userindex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(Userindex, TargetIndex)
        End If
        
        Call InfoHechizo(Userindex)
        
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
        If Not CanSupportUser(Userindex, TargetIndex) Then Exit Function
        
        Call InfoHechizo(Userindex)
        daño = RandomNumber(Hechizos(SpellIndex).MinFuerza, Hechizos(SpellIndex).MaxFuerza)
        
        .flags.DuracionEfecto = 1200
    
        .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + daño
        If .Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Fuerza) * 2) Then _
            .Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Fuerza) * 2)
        
        .flags.TomoPocion = True
        Call WriteUpdateStrenght(TargetIndex)
    
    ' <-------- Quita Fuerza ---------->
    ElseIf Hechizos(SpellIndex).SubeFuerza = 2 Then
    
        If Not PuedeAtacar(Userindex, TargetIndex) Then Exit Function
        
        If Userindex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(Userindex, TargetIndex)
        End If
        
        Call InfoHechizo(Userindex)
        
        .flags.TomoPocion = True
        
        daño = RandomNumber(Hechizos(SpellIndex).MinFuerza, Hechizos(SpellIndex).MaxFuerza)
        .flags.DuracionEfecto = 700
        .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) - daño
        If .Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then .Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
        
        Call WriteUpdateStrenght(TargetIndex)
    End If
    
    ' <-------- Cura salud ---------->
    If Hechizos(SpellIndex).SubeHP = 1 Then
        
             If UserList(Userindex).Stats.MinHp = .Stats.MaxHp Then
    Call WriteConsoleMsg(Userindex, "¡No estás herido!", FontTypeNames.FONTTYPE_FIGHT)
    Exit Function
    End If
    
         If UserList(TargetIndex).Stats.MinHp = .Stats.MaxHp Then
    Call WriteConsoleMsg(Userindex, "¡No está herido!", FontTypeNames.FONTTYPE_FIGHT)
    Exit Function
    End If
        
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then
        Call WriteConsoleMsg(Userindex, "No puedes curar a este usuario.", FontTypeNames.FONTTYPE_INFO)
        Exit Function
    End If
        
        'Verifica que el usuario no este muerto
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(Userindex, "¡El usuario está muerto!", FontTypeNames.FONTTYPE_INFO)
            Exit Function
        End If
        
        ' Chequea si el status permite ayudar al otro usuario
        If Not CanSupportUser(Userindex, TargetIndex) Then Exit Function
           
        daño = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
        daño = daño + Porcentaje(daño, 3 * UserList(Userindex).Stats.ELV)
        
        Call InfoHechizo(Userindex)
    
        .Stats.MinHp = .Stats.MinHp + daño
        If .Stats.MinHp > .Stats.MaxHp Then _
            .Stats.MinHp = .Stats.MaxHp
        
        Call WriteUpdateHP(TargetIndex)
        Call WriteUpdateFollow(TargetIndex)
        
        If Userindex <> TargetIndex Then
            Call WriteConsoleMsg(Userindex, "Le has restaurado " & daño & " puntos de vida a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(TargetIndex, UserList(Userindex).Name & " te ha restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(Userindex, "Te has restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        End If
        
    ' <-------- Quita salud (Daña) ---------->
    ElseIf Hechizos(SpellIndex).SubeHP = 2 Then
        
        If Userindex = TargetIndex Then
            Call WriteConsoleMsg(Userindex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Function
        End If
        
        daño = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
        
        daño = daño + Porcentaje(daño, 3 * UserList(Userindex).Stats.ELV)
        
        If Hechizos(SpellIndex).StaffAffected Then
            If UserList(Userindex).clase = eClass.Mage Then
                If UserList(Userindex).Invent.WeaponEqpObjIndex > 0 Then
                    daño = (daño * (ObjData(UserList(Userindex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                Else
                    daño = daño * 0.7 'Baja daño a 70% del original
                End If
            End If
        End If
        
        If UserList(Userindex).Invent.MunicionEqpObjIndex = LAUDELFICO Or UserList(Userindex).Invent.MunicionEqpObjIndex = FLAUTAELFICA Then
            daño = daño * 1.04  'laud magico de los bardos
        End If
        
                 If UserList(Userindex).Invent.MunicionEqpObjIndex = LaudBronce Or UserList(Userindex).Invent.MunicionEqpObjIndex = AnilloBronce Then
            daño = daño * 1.05  'laud magico de los bardos
        End If
        
                         If UserList(Userindex).Invent.MunicionEqpObjIndex = LaudPlata Or UserList(Userindex).Invent.MunicionEqpObjIndex = AnilloPlata Then
            daño = daño * 1.06  'laud magico de los bardos
        End If
        
         If UserList(Userindex).Invent.MunicionEqpObjIndex = LAUDSUPERMAGICO Or UserList(Userindex).Invent.MunicionEqpObjIndex = FLAUTAANTIGUA Then
            daño = daño * 1.09  'laud magico de los bardos
        End If
        
        'cascos antimagia
        If (.Invent.CascoEqpObjIndex > 0) Then
            daño = daño - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)
        End If
        
        'anillos
        If (.Invent.AnilloEqpObjIndex > 0) Then
            daño = daño - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)
        End If
        
        daño = daño - (daño * UserList(TargetIndex).Stats.UserSkills(eSkill.Resistencia) / 2000)
        
        If daño < 0 Then daño = 0
        
        If Not PuedeAtacar(Userindex, TargetIndex) Then Exit Function
        
        If Userindex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(Userindex, TargetIndex)
        End If
        
        Call InfoHechizo(Userindex)
        
        .Stats.MinHp = .Stats.MinHp - daño
        SendData SendTarget.ToPCArea, TargetIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, DAMAGE_NORMAL)
        Call WriteUpdateHP(TargetIndex)
        Call WriteUpdateFollow(TargetIndex)
        
        Call WriteConsoleMsg(Userindex, "Le has quitado " & daño & " puntos de vida a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(TargetIndex, UserList(Userindex).Name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        
        'Muere
        If .Stats.MinHp < 1 Then
        
            If .flags.AtacablePor <> Userindex Then
                'Store it!
                Call Statistics.StoreFrag(Userindex, TargetIndex)
                Call ContarMuerte(TargetIndex, Userindex)
            End If
            
            .Stats.MinHp = 0
            Call ActStats(TargetIndex, Userindex)
            Call UserDie(TargetIndex)
        End If
        
    End If
    
    ' <-------- Aumenta Mana ---------->
    If Hechizos(SpellIndex).SubeMana = 1 Then
        
        Call InfoHechizo(Userindex)
        .Stats.MinMAN = .Stats.MinMAN + daño
        If .Stats.MinMAN > .Stats.MaxMAN Then _
            .Stats.MinMAN = .Stats.MaxMAN
        
        Call WriteUpdateMana(TargetIndex)
        Call WriteUpdateFollow(TargetIndex)
        
        If Userindex <> TargetIndex Then
            Call WriteConsoleMsg(Userindex, "Le has restaurado " & daño & " puntos de maná a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(TargetIndex, UserList(Userindex).Name & " te ha restaurado " & daño & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(Userindex, "Te has restaurado " & daño & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT)
        End If
        
    
    ' <-------- Quita Mana ---------->
    ElseIf Hechizos(SpellIndex).SubeMana = 2 Then
        If Not PuedeAtacar(Userindex, TargetIndex) Then Exit Function
        
        If Userindex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(Userindex, TargetIndex)
        End If
        
        Call InfoHechizo(Userindex)
        
        If Userindex <> TargetIndex Then
            Call WriteConsoleMsg(Userindex, "Le has quitado " & daño & " puntos de maná a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(TargetIndex, UserList(Userindex).Name & " te ha quitado " & daño & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(Userindex, "Te has quitado " & daño & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT)
        End If
        
        .Stats.MinMAN = .Stats.MinMAN - daño
        If .Stats.MinMAN < 1 Then .Stats.MinMAN = 0
        
        Call WriteUpdateMana(TargetIndex)
        Call WriteUpdateFollow(TargetIndex)
        
    End If
    
    ' <-------- Aumenta Stamina ---------->
    If Hechizos(SpellIndex).SubeSta = 1 Then
        Call InfoHechizo(Userindex)
        .Stats.MinSta = .Stats.MinSta + daño
        If .Stats.MinSta > .Stats.MaxSta Then _
            .Stats.MinSta = .Stats.MaxSta
        
        Call WriteUpdateSta(TargetIndex)
        
        If Userindex <> TargetIndex Then
            Call WriteConsoleMsg(Userindex, "Le has restaurado " & daño & " puntos de energía a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(TargetIndex, UserList(Userindex).Name & " te ha restaurado " & daño & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(Userindex, "Te has restaurado " & daño & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT)
        End If
        
    ' <-------- Quita Stamina ---------->
    ElseIf Hechizos(SpellIndex).SubeSta = 2 Then
        If Not PuedeAtacar(Userindex, TargetIndex) Then Exit Function
        
        If Userindex <> TargetIndex Then
            Call UsuarioAtacadoPorUsuario(Userindex, TargetIndex)
        End If
        
        Call InfoHechizo(Userindex)
        
        If Userindex <> TargetIndex Then
            Call WriteConsoleMsg(Userindex, "Le has quitado " & daño & " puntos de energía a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(TargetIndex, UserList(Userindex).Name & " te ha quitado " & daño & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(Userindex, "Te has quitado " & daño & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT)
        End If
        
        .Stats.MinSta = .Stats.MinSta - daño
        
        If .Stats.MinSta < 1 Then .Stats.MinSta = 0
        
        Call WriteUpdateSta(TargetIndex)
        
    End If
End With

HechizoPropUsuario = True

Call FlushBuffer(TargetIndex)

End Function

Public Function CanSupportUser(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer, _
                               Optional ByVal DoCriminal As Boolean = False) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 28/04/2010
'Checks if caster can cast support magic on target user.
'***************************************************
     
 On Error GoTo Errhandler
 
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
        If criminal(TargetIndex) Then
        
            ' Casteador Ciuda?
            If Not criminal(CasterIndex) Then
            
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
            ElseIf Not criminal(CasterIndex) Then
                
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
    
Errhandler:
    Call LogError("Error en CanSupportUser, Error: " & Err.Number & " - " & Err.description & _
                  " CasterIndex: " & CasterIndex & ", TargetIndex: " & TargetIndex)

End Function

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal Userindex As Integer, ByVal Slot As Byte)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

Dim LoopC As Byte

With UserList(Userindex)
    'Actualiza un solo slot
    If Not UpdateAll Then
        'Actualiza el inventario
        If .Stats.UserHechizos(Slot) > 0 Then
            Call ChangeUserHechizo(Userindex, Slot, .Stats.UserHechizos(Slot))
        Else
            Call ChangeUserHechizo(Userindex, Slot, 0)
        End If
    Else
        'Actualiza todos los slots
        For LoopC = 1 To MAXUSERHECHIZOS
            'Actualiza el inventario
            If .Stats.UserHechizos(LoopC) > 0 Then
                Call ChangeUserHechizo(Userindex, LoopC, .Stats.UserHechizos(LoopC))
            Else
                Call ChangeUserHechizo(Userindex, LoopC, 0)
            End If
        Next LoopC
    End If
End With

End Sub

Sub ChangeUserHechizo(ByVal Userindex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************
    
    UserList(Userindex).Stats.UserHechizos(Slot) = Hechizo
    
    If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
        Call WriteChangeSpellSlot(Userindex, Slot)
    Else
        Call WriteChangeSpellSlot(Userindex, Slot)
    End If

End Sub


Public Sub DesplazarHechizo(ByVal Userindex As Integer, ByVal Dire As Integer, ByVal HechizoDesplazado As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

If (Dire <> 1 And Dire <> -1) Then Exit Sub
If Not (HechizoDesplazado >= 1 And HechizoDesplazado <= MAXUSERHECHIZOS) Then Exit Sub

Dim TempHechizo As Integer

With UserList(Userindex)
    If Dire = 1 Then 'Mover arriba
        If HechizoDesplazado = 1 Then
            Call WriteConsoleMsg(Userindex, "No puedes mover el hechizo en esa dirección.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            TempHechizo = .Stats.UserHechizos(HechizoDesplazado)
            .Stats.UserHechizos(HechizoDesplazado) = .Stats.UserHechizos(HechizoDesplazado - 1)
            .Stats.UserHechizos(HechizoDesplazado - 1) = TempHechizo
        End If
    Else 'mover abajo
        If HechizoDesplazado = MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(Userindex, "No puedes mover el hechizo en esa dirección.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        Else
            TempHechizo = .Stats.UserHechizos(HechizoDesplazado)
            .Stats.UserHechizos(HechizoDesplazado) = .Stats.UserHechizos(HechizoDesplazado + 1)
            .Stats.UserHechizos(HechizoDesplazado + 1) = TempHechizo
        End If
    End If
End With

End Sub

Public Sub DisNobAuBan(ByVal Userindex As Integer, NoblePts As Long, BandidoPts As Long)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    'disminuye la nobleza NoblePts puntos y aumenta el bandido BandidoPts puntos
    Dim EraCriminal As Boolean
    EraCriminal = criminal(Userindex)
    
    With UserList(Userindex)
        'Si estamos en la arena no hacemos nada
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
            'pierdo nobleza...
            .Reputacion.NobleRep = .Reputacion.NobleRep - NoblePts
            If .Reputacion.NobleRep < 0 Then
                .Reputacion.NobleRep = 0
            End If
            
            'gano bandido...
            .Reputacion.BandidoRep = .Reputacion.BandidoRep + BandidoPts
            If .Reputacion.BandidoRep > MAXREP Then _
                .Reputacion.BandidoRep = MAXREP
            Call WriteMultiMessage(Userindex, eMessages.NobilityLost) 'Call WriteNobilityLost(UserIndex)
            If criminal(Userindex) Then If .Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(Userindex)
        End If
        
        If Not EraCriminal And criminal(Userindex) Then
            Call RefreshCharStatus(Userindex)
        End If
    End With
End Sub

Public Function CanSupportNpc(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 18/09/2010
'Checks if caster can cast support magic on target Npc.
'***************************************************
     
 On Error GoTo Errhandler
 
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
        If criminal(OwnerIndex) Then
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
            If Not criminal(CasterIndex) Then
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
                    
                ' ayudo al npc sin seguro, se convierte en atacable
                Else
                    Call ToogleToAtackable(CasterIndex, OwnerIndex, True)
                    CanSupportNpc = True
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
    
Errhandler:
    Call LogError("Error en CanSupportNpc, Error: " & Err.Number & " - " & Err.description & _
                  " CasterIndex: " & CasterIndex & ", OwnerIndex: " & OwnerIndex)

End Function

Function ResistenciaClase(clase As String) As Integer
Dim Cuan As Integer
Select Case UCase$(clase)
    Case "MAGO"
        Cuan = 3
    Case "DRUIDA"
        Cuan = 2
    Case "CLERIGO"
        Cuan = 2
    Case "BARDO"
        Cuan = 1
    Case Else
        Cuan = 0
End Select
ResistenciaClase = Cuan
End Function

