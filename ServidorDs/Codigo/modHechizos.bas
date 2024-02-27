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

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByVal spell As Integer, _
                           Optional ByVal DecirPalabras As Boolean = False, _
                           Optional ByVal IgnoreVisibilityCheck As Boolean = False)
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 13/02/2009
      '13/02/2009: ZaMa - Los npcs que tiren magias, no podran hacerlo en mapas donde no se permita usarla.
      '***************************************************
10    If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
20    If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then Exit Sub

      ' Si no se peude usar magia en el mapa, no le deja hacerlo.
30    If MapInfo(UserList(UserIndex).Pos.map).MagiaSinEfecto > 0 Then Exit Sub

40    Npclist(NpcIndex).CanAttack = 0
      Dim daño As Integer

50    With UserList(UserIndex)
60        If Hechizos(spell).SubeHP = 1 Then
          
70            daño = RandomNumber(Hechizos(spell).MinHp, Hechizos(spell).MaxHp)
80            daño = daño - (daño * UserList(UserIndex).Stats.UserSkills(eSkill.Resistencia) / 2000)
              ' daño = daño - Porcentaje(daño, Int(((UserList(UserIndex).Stats.UserSkills(Resistencia) + 1) / 4) + ResistenciaClase(UserList(UserIndex).clase = Druid Or Cleric Or Assasin Or Bard Or Mage Or Paladin Or warrior Or Hunter)))
90            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(spell).WAV, .Pos.X, .Pos.Y))
100           Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(spell).FXgrh, Hechizos(spell).loops))
          
110           .Stats.MinHp = .Stats.MinHp + daño
120           If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
              
130           Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
140           Call WriteUpdateUserStats(UserIndex)
          
150       ElseIf Hechizos(spell).SubeHP = 2 Then
              
160           If .flags.Privilegios And PlayerType.User Then
170            daño = daño - (daño * UserList(UserIndex).Stats.UserSkills(eSkill.Resistencia) / 2000)
             ' daño = daño - Porcentaje(daño, Int(((UserList(UserIndex).Stats.UserSkills(Resistencia) + 1) / 4) + ResistenciaClase(UserList(UserIndex).clase = (Druid Or Mage Or Paladin Or Hunter Or Assasin Or Cleric Or Pirat Or Bard))))
180       Call SubirSkill(UserIndex, eSkill.Resistencia, True)
190               daño = RandomNumber(Hechizos(spell).MinHp, Hechizos(spell).MaxHp)
                  
200               If .Invent.CascoEqpObjIndex > 0 Then
210                   daño = daño - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)
220               End If
230               daño = daño - (daño * UserList(UserIndex).Stats.UserSkills(eSkill.Resistencia) / 2000)
                 '  daño = daño - Porcentaje(daño, Int(((UserList(UserIndex).Stats.UserSkills(Resistencia) + 1) / 4) + ResistenciaClase(UserList(UserIndex).clase = Druid Or Cleric Or Assasin Or Bard Or Mage Or Paladin Or warrior Or Hunter)))
240               If .Invent.AnilloEqpObjIndex > 0 Then
250                   daño = daño - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)
260               End If
270              daño = daño - (daño * UserList(UserIndex).Stats.UserSkills(eSkill.Resistencia) / 2000)
                ' daño = daño - Porcentaje(daño, Int(((UserList(UserIndex).Stats.UserSkills(Resistencia) + 1) / 4) + ResistenciaClase(UserList(UserIndex).clase = Druid Or Cleric Or Assasin Or Bard Or Mage Or Paladin Or warrior Or Hunter)))
280               If daño < 0 Then daño = 0
                  
290               Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(spell).WAV, .Pos.X, .Pos.Y))
300               Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(spell).FXgrh, Hechizos(spell).loops))
              
310               .Stats.MinHp = .Stats.MinHp - daño
                  
320               Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
330                          SendData SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, DAMAGE_NORMAL)
340               Call WriteUpdateUserStats(UserIndex)
                  
                  'Muere
350               If .Stats.MinHp < 1 Then
360                   .Stats.MinHp = 0
370                   If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
380                       RestarCriminalidad (UserIndex)
390                   End If
400                   Call UserDie(UserIndex)
                      '[Barrin 1-12-03]
410                   If Npclist(NpcIndex).MaestroUser > 0 Then
                          'Store it!
                          'Call Statistics.StoreFrag(Npclist(NpcIndex).MaestroUser, Userindex)
                          
420                       Call ContarMuerte(UserIndex, Npclist(NpcIndex).MaestroUser)
430                       Call ActStats(UserIndex, Npclist(NpcIndex).MaestroUser)
440                   End If
                      '[/Barrin]
450               End If
              
460           End If
              
470       End If
          
480       If Hechizos(spell).Paraliza = 1 Or Hechizos(spell).Inmoviliza = 1 Then
490           If .flags.Paralizado = 0 Then
500               Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(spell).WAV, .Pos.X, .Pos.Y))
510               Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(spell).FXgrh, Hechizos(spell).loops))
                    
520               If .Invent.AnilloEqpObjIndex = SUPERANILLO Then
530                   Call WriteConsoleMsg(UserIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
540                   Exit Sub
550               End If
                  
560               If Hechizos(spell).Inmoviliza = 1 Then
570                   .flags.Inmovilizado = 1
580               End If
                    
590               .flags.Paralizado = 1
600               .Counters.Paralisis = IntervaloParalizado
                    
610               Call WriteParalizeOK(UserIndex)
620           End If
630       End If
          
640       If Hechizos(spell).Estupidez = 1 Then   ' turbacion
650            If .flags.Estupidez = 0 Then
660                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(spell).WAV, .Pos.X, .Pos.Y))
670                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(spell).FXgrh, Hechizos(spell).loops))
                    
680                   If .Invent.AnilloEqpObjIndex = SUPERANILLO Then
690                       Call WriteConsoleMsg(UserIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
700                       Exit Sub
710                   End If
                    
720                 .flags.Estupidez = 1
730                 .Counters.Ceguera = IntervaloInvisible
                            
740               Call WriteDumb(UserIndex)
750            End If
760       End If
770   End With

End Sub

Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer, ByVal spell As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      'solo hechizos ofensivos!

10    If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
20    Npclist(NpcIndex).CanAttack = 0

      Dim daño As Integer

30    If Hechizos(spell).SubeHP = 2 Then
          
40        daño = RandomNumber(Hechizos(spell).MinHp, Hechizos(spell).MaxHp)
50        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(spell).WAV, Npclist(TargetNPC).Pos.X, Npclist(TargetNPC).Pos.Y))
60        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(Npclist(TargetNPC).Char.CharIndex, Hechizos(spell).FXgrh, Hechizos(spell).loops))
          
70        Npclist(TargetNPC).Stats.MinHp = Npclist(TargetNPC).Stats.MinHp - daño
          
          'Muere
80        If Npclist(TargetNPC).Stats.MinHp < 1 Then
90            Npclist(TargetNPC).Stats.MinHp = 0
100           If Npclist(NpcIndex).MaestroUser > 0 Then
110               Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
120           Else
130               Call MuereNpc(TargetNPC, 0)
140           End If
150       End If
          
160   End If
          
End Sub

Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo Errhandler
          
          Dim j As Integer
20        For j = 1 To MAXUSERHECHIZOS
30            If UserList(UserIndex).Stats.UserHechizos(j) = i Then
40                TieneHechizo = True
50                Exit Function
60            End If
70        Next

80    Exit Function
Errhandler:

End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal Slot As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim hIndex As Integer
      Dim j As Integer
      Dim i As Integer
      Dim NoLoUsa As Integer

10    With UserList(UserIndex)
20        hIndex = ObjData(.Invent.Object(Slot).ObjIndex).HechizoIndex
          
30          For i = 1 To NUMCLASES
40             If ObjData(.Invent.Object(Slot).ObjIndex).ClaseProhibida(i) = UserList(UserIndex).clase Then
50                 NoLoUsa = 1
60                 Call WriteConsoleMsg(UserIndex, "Tu clase no puede aprender este hechizo.", FontTypeNames.FONTTYPE_INFO)
70             End If
80           Next i
          
90        If Not TieneHechizo(hIndex, UserIndex) Then
              'Buscamos un slot vacio
100           For j = 1 To MAXUSERHECHIZOS
110               If .Stats.UserHechizos(j) = 0 Then Exit For
120           Next j
                  
130          If .Stats.UserHechizos(j) <> 0 Then
140               Call WriteConsoleMsg(UserIndex, "No tienes espacio para más hechizos.", FontTypeNames.FONTTYPE_INFO)
150           Else
160               If NoLoUsa = 0 Then
170                   .Stats.UserHechizos(j) = hIndex
180                   Call UpdateUserHechizos(False, UserIndex, CByte(j))
                      'Quitamos del inv el item
190                   Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)
200               End If
210           End If
220       Else
230           Call WriteConsoleMsg(UserIndex, "Ya tienes ese hechizo.", FontTypeNames.FONTTYPE_INFO)
240       End If
250   End With

End Sub
            
Sub DecirPalabrasMagicas(ByVal SpellIndex As Integer, ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 17/11/2009
      '25/07/2009: ZaMa - Invisible admins don't say any word when casting a spell
      '17/11/2009: ZaMa - Now the user become visible when casting a spell, if it is hidden
      '***************************************************
10    On Error Resume Next
20    With UserList(UserIndex)
30        If .flags.AdminInvisible <> 1 Then
              'Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageChatOverHead(SpellWords, .Char.CharIndex, vbCyan))
40            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePalabrasMagicas(.Char.CharIndex, SpellIndex, vbCyan))
              
              ' Si estaba oculto, se vuelve visible
50            If .flags.Oculto = 1 Then
60                .flags.Oculto = 0
70                .Counters.TiempoOculto = 0
                  
80                If .flags.invisible = 0 Then
90                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
100                   Call SetInvisible(UserIndex, .Char.CharIndex, False)
110               End If
120           End If
130       End If
140   End With
150       Exit Sub
End Sub

''
' Check if an user can cast a certain spell
'
' @param UserIndex Specifies reference to user
' @param HechizoIndex Specifies reference to spell
' @return   True if the user can cast the spell, otherwise returns false
Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: 12/01/2010
      'Last Modification By: ZaMa
      '06/11/09 - Corregida la bonificación de maná del mimetismo en el druida con flauta mágica equipada.
      '19/11/2009: ZaMa - Validacion de mana para el Invocar Mascotas
      '12/01/2010: ZaMa - Validacion de mana para hechizos lanzados por druida.
      '***************************************************
      Dim DruidManaBonus As Single

10        With UserList(UserIndex)
20            If .flags.Muerto Then
30                Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos estando muerto.", FontTypeNames.FONTTYPE_INFO)
40                Exit Function
50            End If
                  
60            If Hechizos(HechizoIndex).NeedStaff > 0 Then
70            If UserList(UserIndex).clase = eClass.Mage Then
80                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
90                    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
100                       Call WriteConsoleMsg(UserIndex, "No posees un báculo lo suficientemente poderoso para que puedas lanzar el conjuro.", FontTypeNames.FONTTYPE_INFO)
110                       PuedeLanzar = False
120                       Exit Function
130                   End If
140               Else
150                   Call WriteConsoleMsg(UserIndex, "No puedes lanzar este conjuro sin la ayuda de un báculo.", FontTypeNames.FONTTYPE_INFO)
160                   PuedeLanzar = False
170                   Exit Function
180               End If
190           End If
200       End If
              
210       If UserList(UserIndex).Stats.UserSkills(eSkill.Magia) < Hechizos(HechizoIndex).MinSkill Then
220           Call WriteConsoleMsg(UserIndex, "No tenes suficientes puntos de magia para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
230           PuedeLanzar = False
240           Exit Function
250       End If
          
260       If UserList(UserIndex).Stats.MinSta < Hechizos(HechizoIndex).StaRequerido Then
270           If UserList(UserIndex).Genero = eGenero.Hombre Then
280               Call WriteConsoleMsg(UserIndex, "Estás muy cansado para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
290           Else
300               Call WriteConsoleMsg(UserIndex, "Estás muy cansada para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
310           End If
320           PuedeLanzar = False
330           Exit Function
340       End If

350       If UserList(UserIndex).clase = eClass.Druid Then
360           If UserList(UserIndex).Invent.MunicionEqpObjIndex = FLAUTAELFICA And UserList(UserIndex).Invent.MunicionEqpObjIndex = FLAUTAANTIGUA And UserList(UserIndex).Invent.MunicionEqpObjIndex = AnilloBronce And UserList(UserIndex).Invent.MunicionEqpObjIndex = AnilloPlata Then
370               If Hechizos(HechizoIndex).Mimetiza Then
380                   DruidManaBonus = 0.5
390               ElseIf Hechizos(HechizoIndex).Tipo = uInvocacion Then
400                   DruidManaBonus = 0.7
410               Else
420                   DruidManaBonus = 1
430               End If
440           Else
450               DruidManaBonus = 1
460           End If
470       Else
480           DruidManaBonus = 1
490       End If
          
500       If UserList(UserIndex).Stats.MinMAN < Hechizos(HechizoIndex).ManaRequerido * DruidManaBonus Then
510           Call WriteConsoleMsg(UserIndex, "No tenes suficiente mana.", FontTypeNames.FONTTYPE_INFO)
520           PuedeLanzar = False
530           Exit Function
540       End If
550           End With
560       PuedeLanzar = True
End Function

Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef B As Boolean)
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

10        With UserList(UserIndex)
20            PosCasteadaX = .flags.TargetX
30            PosCasteadaY = .flags.TargetY
40            PosCasteadaM = .flags.TargetMap
              
50            h = .flags.Hechizo
              
60            If Hechizos(h).RemueveInvisibilidadParcial = 1 Then
70                B = True
80                For TempX = PosCasteadaX - 8 To PosCasteadaX + 8
90                    For TempY = PosCasteadaY - 8 To PosCasteadaY + 8
100                       If InMapBounds(PosCasteadaM, TempX, TempY) Then
110                           If MapData(PosCasteadaM, TempX, TempY).UserIndex > 0 Then
                                  'hay un user
120                               If UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.invisible = 1 And UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).flags.AdminInvisible = 0 Then
130                                   Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(MapData(PosCasteadaM, TempX, TempY).UserIndex).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))
140                               End If
150                           End If
160                       End If
170                   Next TempY
180               Next TempX
              
190               Call InfoHechizo(UserIndex)
200           End If
210       End With
End Sub

''
' Le da propiedades al nuevo npc
'
' @param UserIndex  Indice del usuario que invoca.
' @param b  Indica si se termino la operación.

Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef HechizoCasteado As Boolean)
      '***************************************************
      'Author: Uknown
      'Last modification: 18/11/2009
      'Sale del sub si no hay una posición valida.
      '18/11/2009: Optimizacion de codigo.
      '***************************************************
10    On Error GoTo error

20    With UserList(UserIndex)
              'No permitimos se invoquen criaturas en zonas seguras
30        If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then
40            Call WriteConsoleMsg(UserIndex, "No puedes invocar criaturas en zona segura.", FontTypeNames.FONTTYPE_INFO)
50            Exit Sub
60        End If

70     If (Hechizos(.flags.Hechizo).NumNpc = 111 Or Hechizos(.flags.Hechizo).NumNpc = 110 Or Hechizos(.flags.Hechizo).NumNpc = ELEMENTALFUEGO Or Hechizos(.flags.Hechizo).NumNpc = ELEMENTALTIERRA Or Hechizos(.flags.Hechizo).NumNpc = LOBO Or Hechizos(.flags.Hechizo).NumNpc = ZOMBIE Or Hechizos(.flags.Hechizo).NumNpc = ELEMENTALAGUA Or Hechizos(.flags.Hechizo).NumNpc = OSOS) And (MapInfo(.Pos.map).Pk = False Or MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONASEGURA) Then
80        WriteConsoleMsg UserIndex, "No puedes invocar criaturas en zona segura.", FontTypeNames.FONTTYPE_INFO
90        Exit Sub
100       End If

         ' If .Pos.Map = 298 Or .Pos.Map = 293 Or .Pos.Map = 276 Or .Pos.Map = 300 Then Exit Sub
          'No deja invocar mas de 1 fatuo o 1 espiritu indomable
             Dim SpellIndex As Integer, NroNpcs As Integer, NpcIndex As Integer, PetIndex As Integer
          Dim TargetPos As WorldPos
          
          
110       TargetPos.map = .flags.TargetMap
120       TargetPos.X = .flags.TargetX
130       TargetPos.Y = .flags.TargetY
          
140       SpellIndex = .flags.Hechizo

           
          ' If Hechizos(SpellIndex).NumNpc = 110 And UserList(UserIndex).NroMascotas = 1 Then Exit Sub
         ' If Hechizos(SpellIndex).NumNpc = ELEMENTALFUEGO And UserList(UserIndex).MascotasIndex = 1 Then Exit Sub
          ' Warp de mascotas
150       If Hechizos(SpellIndex).Warp = 1 Then
160           PetIndex = FarthestPet(UserIndex)
              
              ' La invoco cerca mio
170           If Npclist(.MascotasType(.NroMascotas)).Contadores.TiempoExistencia = 0 Then
180           If .NroMascotas > 0 Then
190       WarpMascotas UserIndex, False
200       .NroMascotas = 0
210       ElseIf .NroMascotas <= 0 Then
220       WarpMascotas UserIndex, True
230       .NroMascotas = 3
240           End If
250           End If
          ' Invocacion normal
260       Else
      'solo 1 fuego fatuo puede ser invocadooooo wachin
270       If PetIndex <= 0 Then
280           If .NroMascotas >= MAXMASCOTAS Then Exit Sub
              
290           If .NroMascotas > 0 Then
300               If .MascotasType(.NroMascotas) = 111 And Hechizos(SpellIndex).NumNpc <> 111 Then Exit Sub
310               If .MascotasType(.NroMascotas) = 111 And Hechizos(SpellIndex).NumNpc = 111 Then Exit Sub
320               If .MascotasType(.NroMascotas) = 110 And Hechizos(SpellIndex).NumNpc <> 110 Then Exit Sub
330               If .MascotasType(.NroMascotas) = 110 And Hechizos(SpellIndex).NumNpc = 110 Then Exit Sub
340           End If
350       End If
          
360       If Hechizos(SpellIndex).NumNpc = 111 And .NroMascotas >= 1 Then Exit Sub
370       If Hechizos(SpellIndex).NumNpc = 110 And .NroMascotas >= 1 Then Exit Sub
      '  .NroMascotas = 0
               
        '   If Hechizos(SpellIndex).NumNpc = ELEMENTALFUEGO And Npclist(PetIndex).Numero = 89 Then Exit Sub
380           For NroNpcs = 1 To Hechizos(SpellIndex).cant
                  
390               If .NroMascotas < MAXMASCOTAS Then
400                   NpcIndex = SpawnNpc(Hechizos(SpellIndex).NumNpc, TargetPos, True, False)
410                   If NpcIndex > 0 Then
420                       .NroMascotas = .NroMascotas + 1
                          
430                       PetIndex = FreeMascotaIndex(UserIndex)
                          
440                       .MascotasIndex(PetIndex) = NpcIndex
450                       .MascotasType(PetIndex) = Npclist(NpcIndex).Numero
                          
460                       With Npclist(NpcIndex)
470                           .MaestroUser = UserIndex
480                           .Contadores.TiempoExistencia = IntervaloInvocacion
490                           .GiveGLD = 0
500                       End With
                          
510                       Call FollowAmo(NpcIndex)
520                   Else
530                       Exit Sub
540                   End If
550               Else
560                   Exit For
570               End If
              
580           Next NroNpcs
590       End If
600   End With

610   Call InfoHechizo(UserIndex)
620   HechizoCasteado = True

630   Exit Sub

error:
640       With UserList(UserIndex)
650           LogError ("[" & Err.Number & "] " & Err.Description & " por el usuario " & .Name & "(" & UserIndex & _
                      ") en (" & .Pos.map & ", " & .Pos.X & ", " & .Pos.Y & "). Tratando de tirar el hechizo " & _
                      Hechizos(SpellIndex).Nombre & "(" & SpellIndex & ") en la posicion ( " & .flags.TargetX & ", " & .flags.TargetY & ")")
660       End With

End Sub
Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal SpellIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 18/11/2009
      '18/11/2009: ZaMa - Optimizacion de codigo.
      '***************************************************

10        If Not UserList(UserIndex).flags.ModoCombate Then
20        Call WriteConsoleMsg(UserIndex, "Debes estar en modo de combate para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
30        Exit Sub
40        End If
           
          Dim HechizoCasteado As Boolean
          Dim ManaRequerida As Integer
          
50        Select Case Hechizos(SpellIndex).Tipo
              Case TipoHechizo.uInvocacion
60                Call HechizoInvocacion(UserIndex, HechizoCasteado)
                  
70            Case TipoHechizo.uEstado
80                Call HechizoTerrenoEstado(UserIndex, HechizoCasteado)
90        End Select

100          If HechizoCasteado Then
110           With UserList(UserIndex)
120               Call SubirSkill(UserIndex, eSkill.Magia, True)
                  
130               ManaRequerida = Hechizos(SpellIndex).ManaRequerido
                  
140               If Hechizos(SpellIndex).Warp = 1 Then ' Invocó una mascota
                  ' Consume toda la mana
150                   ManaRequerida = 1000
160               End If
                  
                  ' Quito la mana requerida
170               .Stats.MinMAN = .Stats.MinMAN - ManaRequerida
180               If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
                  
                  ' Quito la estamina requerida
190               .Stats.MinSta = .Stats.MinSta - Hechizos(SpellIndex).StaRequerido
200               If .Stats.MinSta < 0 Then .Stats.MinSta = 0
                  
                  ' Update user stats
210               Call WriteUpdateUserStats(UserIndex)
220           End With
230       End If
          
End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal SpellIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 12/01/2010
      '18/11/2009: ZaMa - Optimizacion de codigo.
      '12/01/2010: ZaMa - Optimizacion y agrego bonificaciones al druida.
      '***************************************************
          
10        If Not UserList(UserIndex).flags.ModoCombate Then
20        Call WriteConsoleMsg(UserIndex, "Debes estar en modo de combate para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
30        Exit Sub
40        End If
          
          Dim HechizoCasteado As Boolean
          Dim ManaRequerida As Integer
          
50        Select Case Hechizos(SpellIndex).Tipo
              Case TipoHechizo.uEstado
                  ' Afectan estados (por ejem : Envenenamiento)
60                Call HechizoEstadoUsuario(UserIndex, HechizoCasteado)
              
70            Case TipoHechizo.uPropiedades
                  ' Afectan HP,MANA,STAMINA,ETC
80                HechizoCasteado = HechizoPropUsuario(UserIndex)
90        End Select

100       If HechizoCasteado Then
110           With UserList(UserIndex)
120               Call SubirSkill(UserIndex, eSkill.Magia, True)
                  
130               ManaRequerida = Hechizos(SpellIndex).ManaRequerido
                  
                  ' Bonificaciones para druida
140               If .clase = eClass.Druid Then
                      ' Solo con flauta magica
150                   If .Invent.MunicionEqpObjIndex = FLAUTAELFICA And .Invent.MunicionEqpObjIndex = FLAUTAANTIGUA And UserList(UserIndex).Invent.MunicionEqpObjIndex = AnilloBronce And UserList(UserIndex).Invent.MunicionEqpObjIndex = AnilloPlata Then
160                       If Hechizos(SpellIndex).Mimetiza = 1 Then
                              ' 50% menos de mana para mimetismo
170                           ManaRequerida = ManaRequerida * 0.5
                              
180                       ElseIf SpellIndex <> APOCALIPSIS_SPELL_INDEX Then
                              ' 10% menos de mana para todo menos apoca y descarga
190                           ManaRequerida = ManaRequerida * 0.9
200                       End If
210                   End If
220               End If
                  
                  ' Quito la mana requerida
230               .Stats.MinMAN = .Stats.MinMAN - ManaRequerida
240               If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
                  
                  ' Quito la estamina requerida
250               .Stats.MinSta = .Stats.MinSta - Hechizos(SpellIndex).StaRequerido
260               If .Stats.MinSta < 0 Then .Stats.MinSta = 0
                  
                  ' Update user stats
270               Call WriteUpdateUserStats(UserIndex)
280               Call WriteUpdateUserStats(.flags.TargetUser)
290               .flags.TargetUser = 0
300           End With
310       End If

End Sub

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer)
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
          
10        With UserList(UserIndex)
20            Select Case Hechizos(HechizoIndex).Tipo
                  Case TipoHechizo.uEstado
                      ' Afectan estados (por ejem : Envenenamiento)
30                    Call HechizoEstadoNPC(.flags.TargetNPC, HechizoIndex, HechizoCasteado, UserIndex)
                      
40                Case TipoHechizo.uPropiedades
                      ' Afectan HP,MANA,STAMINA,ETC
50                    Call HechizoPropNPC(HechizoIndex, .flags.TargetNPC, UserIndex, HechizoCasteado)
60            End Select
              
              
70            If HechizoCasteado Then
80                Call SubirSkill(UserIndex, eSkill.Magia, True)
                  
90                ManaRequerida = Hechizos(HechizoIndex).ManaRequerido
                  
                  ' Bonificación para druidas.
100               If .clase = eClass.Druid Then
                      ' Se mostró como usuario, puede ser atacado por npcs
110                   .flags.Ignorado = False
                      
                      ' Solo con flauta equipada
120                   If .Invent.MunicionEqpObjIndex = FLAUTAELFICA And .Invent.MunicionEqpObjIndex = FLAUTAANTIGUA And UserList(UserIndex).Invent.MunicionEqpObjIndex = AnilloBronce And UserList(UserIndex).Invent.MunicionEqpObjIndex = AnilloPlata Then
130                       If Hechizos(HechizoIndex).Mimetiza = 1 Then
                              ' 50% menos de mana para mimetismo
140                           ManaRequerida = ManaRequerida * 0.5
                              ' Será ignorado hasta que pierda el efecto del mimetismo o ataque un npc
150                           .flags.Ignorado = True
160                       Else
                              ' 10% menos de mana para hechizos
170                           If HechizoIndex <> APOCALIPSIS_SPELL_INDEX Then
180                                ManaRequerida = ManaRequerida * 0.9
190                           End If
200                       End If
210                   End If
220               End If
                  
                  ' Quito la mana requerida
230               .Stats.MinMAN = .Stats.MinMAN - ManaRequerida
240               If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
                  
                  ' Quito la estamina requerida
250               .Stats.MinSta = .Stats.MinSta - Hechizos(HechizoIndex).StaRequerido
260               If .Stats.MinSta < 0 Then .Stats.MinSta = 0
                  
                  ' Update user stats
270               Call WriteUpdateUserStats(UserIndex)
280               .flags.TargetNPC = 0
290           End If
300       End With
End Sub


Sub LanzarHechizo(ByVal SpellIndex As Integer, ByVal UserIndex As Integer)
10        On Error GoTo LanzarHechizo_Error
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 02/16/2010
      '24/01/2007 ZaMa - Optimizacion de codigo.
      '02/16/2010: Marco - Now .flags.hechizo makes reference to global spell index instead of user's spell index
      '***************************************************

20    With UserList(UserIndex)
          
30     If Hechizos(SpellIndex).Nombre = "Implorar Ayuda" And .clase <> eClass.Druid Then
40          WriteConsoleMsg UserIndex, "No eres un druida para invocar este hechizos!", FontTypeNames.FONTTYPE_INFO
50          Exit Sub
60      End If
          
70        If .flags.EnConsulta Then
80            Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos si estás en consulta.", FontTypeNames.FONTTYPE_INFO)
90            Exit Sub
100       End If
          
110       If PuedeLanzar(UserIndex, SpellIndex) Then
120           Select Case Hechizos(SpellIndex).Target
                  Case TargetType.uUsuarios
130                   If .flags.TargetUser > 0 Then
140                       If Abs(UserList(.flags.TargetUser).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
150                           Call HandleHechizoUsuario(UserIndex, SpellIndex)
160                       Else
170                           Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
180                       End If
190                   Else
200                       Call WriteConsoleMsg(UserIndex, "Este hechizo actúa sólo sobre usuarios.", FontTypeNames.FONTTYPE_INFO)
210                   End If
                  
220               Case TargetType.uNPC
230                   If .flags.TargetNPC > 0 Then
240                       If Abs(Npclist(.flags.TargetNPC).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
250                           Call HandleHechizoNPC(UserIndex, SpellIndex)
260                       Else
270                           Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
280                       End If
290                   Else
300                       Call WriteConsoleMsg(UserIndex, "Este hechizo sólo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO)
310                   End If
                  
320               Case TargetType.uUsuariosYnpc
330                   If .flags.TargetUser > 0 Then
340                       If Abs(UserList(.flags.TargetUser).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
350                           Call HandleHechizoUsuario(UserIndex, SpellIndex)
360                       Else
370                           Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
380                       End If
390                   ElseIf .flags.TargetNPC > 0 Then
400                       If Abs(Npclist(.flags.TargetNPC).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
410                           Call HandleHechizoNPC(UserIndex, SpellIndex)
420                       Else
430                           Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para lanzar este hechizo.", FontTypeNames.FONTTYPE_WARNING)
440                       End If
450                   Else
460                       Call WriteConsoleMsg(UserIndex, "Target inválido.", FontTypeNames.FONTTYPE_INFO)
470                   End If
                  
480               Case TargetType.uTerreno
490                   Call HandleHechizoTerreno(UserIndex, SpellIndex)
500           End Select
              
510       End If
          
520       If .Counters.Trabajando Then _
              .Counters.Trabajando = .Counters.Trabajando - 1
          
530       If .Counters.Ocultando Then _
              .Counters.Ocultando = .Counters.Ocultando - 1

540   End With

550   Exit Sub

    
560       On Error GoTo 0
570       Exit Sub

LanzarHechizo_Error:

580       Call LogError("Error en LanzarHechizo. Error " & Err.Number & " : " & Err.Description & _
        " Hechizo: " & Hechizos(SpellIndex).Nombre & "(" & SpellIndex & _
        "). Casteado por: " & UserList(UserIndex).Name & "(" & UserIndex & ")." & " Erl: " & Erl)
    
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
      '***************************************************


      Dim HechizoIndex As Integer
      Dim TargetIndex As Integer

10    With UserList(UserIndex)
20        HechizoIndex = .flags.Hechizo
30        TargetIndex = .flags.TargetUser
          
          ' <-------- Agrega Invisibilidad ---------->
40        If Hechizos(HechizoIndex).Invisibilidad = 1 Then
50            If UserList(TargetIndex).flags.Muerto = 1 Then
60                Call WriteConsoleMsg(UserIndex, "¡El usuario está muerto!", FontTypeNames.FONTTYPE_INFO)
70                HechizoCasteado = False
80                Exit Sub
90            End If
              
100           If UserList(TargetIndex).Counters.Saliendo Then
110               If UserIndex <> TargetIndex Then
120                   Call WriteConsoleMsg(UserIndex, "¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
130                   HechizoCasteado = False
140                   Exit Sub
150               Else
160                   Call WriteConsoleMsg(UserIndex, "¡No puedes hacerte invisible mientras te encuentras saliendo!", FontTypeNames.FONTTYPE_WARNING)
170                   HechizoCasteado = False
180                   Exit Sub
190               End If
200           End If
              
              'No usar invi mapas InviSinEfecto
210           If MapInfo(UserList(TargetIndex).Pos.map).InviSinEfecto > 0 Then
220               Call WriteConsoleMsg(UserIndex, "¡La invisibilidad no funciona aquí!", FontTypeNames.FONTTYPE_INFO)
230               HechizoCasteado = False
240               Exit Sub
250           End If
              
              ' Chequea si el status permite ayudar al otro usuario
260           HechizoCasteado = CanSupportUser(UserIndex, TargetIndex, True)
270           If Not HechizoCasteado Then Exit Sub
              
              'Si sos user, no uses este hechizo con GMS.
280           If .flags.Privilegios And PlayerType.User Then
290               If Not UserList(TargetIndex).flags.Privilegios And PlayerType.User Then
300                   HechizoCasteado = False
310                   Exit Sub
320               End If
330           End If
             
340           UserList(TargetIndex).flags.invisible = 1
350           Call SetInvisible(TargetIndex, UserList(TargetIndex).Char.CharIndex, True)
          
360           Call InfoHechizo(UserIndex)
370           HechizoCasteado = True
380       End If
          
          ' <-------- Agrega Mimetismo ---------->
390       If Hechizos(HechizoIndex).Mimetiza = 1 Then
400           Exit Sub
              
410           If UserList(TargetIndex).flags.SlotEvent > 0 Then
420               Exit Sub
430           End If
              
440           If UserList(TargetIndex).flags.Muerto = 1 Then
450               Exit Sub
460           End If
              
470           If UserList(TargetIndex).flags.Navegando = 1 Then
480               Exit Sub
490           End If
500           If .flags.Navegando = 1 Then
510               Exit Sub
520           End If
              
              'Si sos user, no uses este hechizo con GMS.
530           If .flags.Privilegios And PlayerType.User Then
540               If Not UserList(TargetIndex).flags.Privilegios And PlayerType.User Then
550                   Exit Sub
560               End If
570           End If
              
580           If .flags.Mimetizado = 1 Then
590               Call WriteConsoleMsg(UserIndex, "Ya te encuentras mimetizado. El hechizo no ha tenido efecto.", FontTypeNames.FONTTYPE_INFO)
600               Exit Sub
610           End If
              
620           If .flags.AdminInvisible = 1 Then Exit Sub
              
              'copio el char original al mimetizado
              
630           .CharMimetizado.body = .Char.body
640           .CharMimetizado.Head = .Char.Head
650           .CharMimetizado.CascoAnim = .Char.CascoAnim
660           .CharMimetizado.ShieldAnim = .Char.ShieldAnim
670           .CharMimetizado.WeaponAnim = .Char.WeaponAnim
              
680           .flags.Mimetizado = 1
              
              'ahora pongo local el del enemigo
690           .Char.body = UserList(TargetIndex).Char.body
700           .Char.Head = UserList(TargetIndex).Char.Head
710           .Char.CascoAnim = UserList(TargetIndex).Char.CascoAnim
720           .Char.ShieldAnim = UserList(TargetIndex).Char.ShieldAnim
730           .Char.WeaponAnim = GetWeaponAnim(UserIndex, UserList(TargetIndex).Invent.WeaponEqpObjIndex)
              
740           Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
             
750          Call InfoHechizo(UserIndex)
760          HechizoCasteado = True
770       End If
          
          ' <-------- Agrega Envenenamiento ---------->
780       If Hechizos(HechizoIndex).Envenena = 1 Then
790           If UserIndex = TargetIndex Then
800               Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
810               Exit Sub
820           End If
              
830           If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Sub
840           If UserIndex <> TargetIndex Then
850               Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
860           End If
870           UserList(TargetIndex).flags.Envenenado = 1
880           Call InfoHechizo(UserIndex)
890           HechizoCasteado = True
900       End If
          
          ' <-------- Cura Envenenamiento ---------->
910       If Hechizos(HechizoIndex).CuraVeneno = 1 Then
          
              'Verificamos que el usuario no este muerto
920           If UserList(TargetIndex).flags.Muerto = 1 Then
930               Call WriteConsoleMsg(UserIndex, "¡El usuario está muerto!", FontTypeNames.FONTTYPE_INFO)
940               HechizoCasteado = False
950               Exit Sub
960           End If
              
970                   If UserList(TargetIndex).flags.Envenenado = 0 Then
980               Call WriteConsoleMsg(UserIndex, "¡El usuario no está envenenado!", FontTypeNames.FONTTYPE_INFO)
990               HechizoCasteado = False
1000              Exit Sub
1010          End If
              
              ' Chequea si el status permite ayudar al otro usuario
1020          HechizoCasteado = CanSupportUser(UserIndex, TargetIndex)
1030          If Not HechizoCasteado Then Exit Sub
                  
              'Si sos user, no uses este hechizo con GMS.
1040          If .flags.Privilegios And PlayerType.User Then
1050              If Not UserList(TargetIndex).flags.Privilegios And PlayerType.User Then
1060                  Exit Sub
1070              End If
1080          End If
                  
1090          UserList(TargetIndex).flags.Envenenado = 0
1100          Call InfoHechizo(UserIndex)
1110          HechizoCasteado = True
1120      End If
          
          ' <-------- Agrega Maldicion ---------->
1130      If Hechizos(HechizoIndex).Maldicion = 1 Then
1140          If UserIndex = TargetIndex Then
1150              Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
1160              Exit Sub
1170          End If
              
1180          If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Sub
1190          If UserIndex <> TargetIndex Then
1200              Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
1210          End If
1220          UserList(TargetIndex).flags.Maldicion = 1
1230          Call InfoHechizo(UserIndex)
1240          HechizoCasteado = True
1250      End If
          
          ' <-------- Remueve Maldicion ---------->
1260      If Hechizos(HechizoIndex).RemoverMaldicion = 1 Then
1270              UserList(TargetIndex).flags.Maldicion = 0
1280              Call InfoHechizo(UserIndex)
1290              HechizoCasteado = True
1300      End If
          
          ' <-------- Agrega Bendicion ---------->
1310      If Hechizos(HechizoIndex).Bendicion = 1 Then
1320              UserList(TargetIndex).flags.Bendicion = 1
1330              Call InfoHechizo(UserIndex)
1340              HechizoCasteado = True
1350      End If
          
          ' <-------- Agrega Paralisis/Inmobilidad ---------->
1360      If Hechizos(HechizoIndex).Paraliza = 1 Or Hechizos(HechizoIndex).Inmoviliza = 1 Then
1370          If UserIndex = TargetIndex Then
1380              Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
1390              Exit Sub
1400          End If
              
1410          If UserList(TargetIndex).flags.SlotEvent > 0 Then
1420              If Events(UserList(TargetIndex).flags.SlotEvent).Users(UserList(TargetIndex).flags.SlotUserEvent).Selected = 1 Then
1430                  Call WriteConsoleMsg(UserIndex, "El personaje no puede ser paralizado en este evento.", FontTypeNames.FONTTYPE_INFO)
1440                  Exit Sub
1450              End If
1460          End If
                      
1470           If UserList(TargetIndex).flags.Paralizado = 0 Then
1480              If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Sub
                  
1490              If UserIndex <> TargetIndex Then
1500                  Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
1510              End If
                  
1520              Call InfoHechizo(UserIndex)
1530              HechizoCasteado = True
1540              If UserList(TargetIndex).Invent.AnilloEqpObjIndex = SUPERANILLO Then
1550                  Call WriteConsoleMsg(TargetIndex, " Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
1560                  Call WriteConsoleMsg(UserIndex, "¡El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
1570                  Call FlushBuffer(TargetIndex)
1580                  Exit Sub
1590              End If
                  
1600              If UserList(TargetIndex).flags.IsDios Then
1610                  If RandomNumber(1, 100) <= 40 Then
1620                      Call WriteConsoleMsg(TargetIndex, "¡Te han paralizado, pero tu poder es tan garnde que el efecto ha caducado!", FontTypeNames.FONTTYPE_FIGHT)
1630                      Call WriteConsoleMsg(UserIndex, "¡Ingenuo! Tengo el poder de los Dioses. No podrás inmovilizarme.", FontTypeNames.FONTTYPE_FIGHT)
1640                      Call FlushBuffer(TargetIndex)
1650                      Exit Sub
1660                  End If
1670              End If
                  
1680              If Hechizos(HechizoIndex).Inmoviliza = 1 Then UserList(TargetIndex).flags.Inmovilizado = 1
1690              UserList(TargetIndex).flags.Paralizado = 1
1700              UserList(TargetIndex).Counters.Paralisis = IntervaloParalizado
                  
1710              Call WriteParalizeOK(TargetIndex)
1720              Call FlushBuffer(TargetIndex)
1730          End If
1740      End If
          
          ' <-------- Remueve Paralisis/Inmobilidad ---------->
1750      If Hechizos(HechizoIndex).RemoverParalisis = 1 Then
              
              ' Remueve si esta en ese estado
1760          If UserList(TargetIndex).flags.Paralizado = 1 Then
              
                  ' Chequea si el status permite ayudar al otro usuario
1770              HechizoCasteado = CanSupportUser(UserIndex, TargetIndex, True)
1780              If Not HechizoCasteado Then Exit Sub
                  
1790              UserList(TargetIndex).flags.Inmovilizado = 0
1800              UserList(TargetIndex).flags.Paralizado = 0
                  
                  'no need to crypt this
1810              Call WriteParalizeOK(TargetIndex)
1820              Call InfoHechizo(UserIndex)
              
1830          End If
1840      End If
          
          ' <-------- Remueve Estupidez (Aturdimiento) ---------->
1850      If Hechizos(HechizoIndex).RemoverEstupidez = 1 Then
          
              ' Remueve si esta en ese estado
1860          If UserList(TargetIndex).flags.Estupidez = 1 Then
              
                  ' Chequea si el status permite ayudar al otro usuario
1870              HechizoCasteado = CanSupportUser(UserIndex, TargetIndex)
1880              If Not HechizoCasteado Then Exit Sub
              
1890              UserList(TargetIndex).flags.Estupidez = 0
                  
                  'no need to crypt this
1900              Call WriteDumbNoMore(TargetIndex)
1910              Call FlushBuffer(TargetIndex)
1920              Call InfoHechizo(UserIndex)
              
1930          End If
1940      End If
          
          ' <-------- Revive ---------->
1950      If Hechizos(HechizoIndex).Revivir = 1 Then
1960          If UserList(TargetIndex).flags.Muerto = 1 Then
                  
                  'Seguro de resurreccion (solo afecta a los hechizos, no al sacerdote ni al comando de GM)
1970              If UserList(TargetIndex).flags.ModoCombate Then
1980                  Call WriteConsoleMsg(UserIndex, "El usuario esta en Modo Combate. No puedes revivirlo.", FontTypeNames.FONTTYPE_INFO)
1990                  HechizoCasteado = False
2000                  Exit Sub
2010              End If
              
                  'No usar resu en mapas con ResuSinEfecto
2020              If MapInfo(UserList(TargetIndex).Pos.map).ResuSinEfecto > 0 Then
2030                  Call WriteConsoleMsg(UserIndex, "¡Revivir no está permitido aquí! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
2040                  HechizoCasteado = False
2050                  Exit Sub
2060              End If
                  
                 'No podemos resucitar si nuestra barra de energía no está llena. (GD: 29/04/07)
2070              If .Stats.MinSta = 500 Then
2080                  Call WriteConsoleMsg(UserIndex, "No puedes resucitar si no tienes 500 puntos de energía.", FontTypeNames.FONTTYPE_INFO)
2090                  HechizoCasteado = False
2100                  Exit Sub
2110              End If
                  
                  
              'revisamos si necesita vara
2120              If .clase = eClass.Mage Then
2130                  If .Invent.WeaponEqpObjIndex > 0 Then
2140                      If ObjData(.Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
2150                          Call WriteConsoleMsg(UserIndex, "Necesitas un báculo mejor para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
2160                          HechizoCasteado = False
2170                          Exit Sub
2180                      End If
2190                  End If
2200              ElseIf .clase = eClass.Bard Then
2210                  If .Invent.MunicionEqpObjIndex <> LAUDELFICO And _
                          .Invent.MunicionEqpObjIndex <> LAUDSUPERMAGICO And _
                          .Invent.MunicionEqpObjIndex <> LaudBronce And _
                          .Invent.MunicionEqpObjIndex <> LaudPlata Then
2220                      Call WriteConsoleMsg(UserIndex, "Necesitas un instrumento mágico para devolver la vida.", FontTypeNames.FONTTYPE_INFO)
2230                      HechizoCasteado = False
2240                      Exit Sub
2250                  End If
2260              ElseIf .clase = eClass.Druid Then
2270                  If .Invent.MunicionEqpObjIndex <> FLAUTAELFICA And .Invent.MunicionEqpObjIndex <> FLAUTAANTIGUA And .Invent.MunicionEqpObjIndex <> AnilloBronce And .Invent.MunicionEqpObjIndex <> AnilloPlata Then
2280                      Call WriteConsoleMsg(UserIndex, "Necesitas un instrumento mágico para devolver la vida.", FontTypeNames.FONTTYPE_INFO)
2290                      HechizoCasteado = False
2300                      Exit Sub
2310                  End If
2320              End If
                  
                  ' Chequea si el status permite ayudar al otro usuario
2330              HechizoCasteado = CanSupportUser(UserIndex, TargetIndex, True)
2340              If Not HechizoCasteado Then Exit Sub
          
                  Dim EraCriminal As Boolean
2350              EraCriminal = criminal(UserIndex)
                  
2360              If Not criminal(TargetIndex) Then
2370                  If TargetIndex <> UserIndex Then
2380                      .Reputacion.NobleRep = .Reputacion.NobleRep + 500
2390                      If .Reputacion.NobleRep > MAXREP Then _
                              .Reputacion.NobleRep = MAXREP
2400                      Call WriteConsoleMsg(UserIndex, "¡Los Dioses te sonríen, has ganado 500 puntos de nobleza!", FontTypeNames.FONTTYPE_INFO)
2410                  End If
2420              End If
                  
2430              If EraCriminal And Not criminal(UserIndex) Then
2440                  Call RefreshCharStatus(UserIndex)
2450              End If
                  
2460              With UserList(TargetIndex)
                      'Pablo Toxic Waste (GD: 29/04/07)
2470                  .Stats.MinAGU = 0
2480                  .flags.Sed = 1
2490                  .Stats.MinHam = 0
2500                  .flags.Hambre = 1
2510                  Call WriteUpdateHungerAndThirst(TargetIndex)
2520                  Call InfoHechizo(UserIndex)
2530                  .Stats.MinMAN = 0
2540                  .Stats.MinSta = 0
2550              End With
                  
                  'Agregado para quitar la penalización de vida en el ring y cambio de ecuacion. (NicoNZ)
2560              'If (TriggerZonaPelea(UserIndex, TargetIndex) <> TRIGGER6_PERMITE) Then
                      'Solo saco vida si es User. no quiero que exploten GMs por ahi.
2570                  'If .flags.Privilegios And PlayerType.User Then
2580                      '.Stats.MinHp = .Stats.MinHp * (1 - UserList(TargetIndex).Stats.ELV * 0.015)
2590                  'End If
2600              'End If
                  
2610             ' If (.Stats.MinHp <= 0) Then
2620                 ' Call UserDie(UserIndex)
                      'Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar fue demasiado grande.", FontTypeNames.FONTTYPE_INFO)
2630                 ' HechizoCasteado = False
2640             ' Else
                     ' 'Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar te ha debilitado.", FontTypeNames.FONTTYPE_INFO)
2650                  HechizoCasteado = True
2660              'End If
                  
2670              If UserList(TargetIndex).flags.Traveling = 1 Then
2680                  UserList(TargetIndex).Counters.goHome = 0
2690                  UserList(TargetIndex).flags.Traveling = 0
                      'Call WriteConsoleMsg(TargetIndex, "Tu viaje ha sido cancelado.", FontTypeNames.FONTTYPE_FIGHT)
2700                  Call WriteMultiMessage(TargetIndex, eMessages.CancelHome)
2710              End If
                  
2720              Call RevivirUsuario(TargetIndex)
2730          Else
2740              HechizoCasteado = False
2750          End If
          
2760      End If
          
          ' <-------- Agrega Ceguera ---------->
2770      If Hechizos(HechizoIndex).Ceguera = 1 Then
2780          If UserIndex = TargetIndex Then
2790              Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
2800              Exit Sub
2810          End If
              
2820              If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Sub
2830              If UserIndex <> TargetIndex Then
2840                  Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
2850              End If
2860              UserList(TargetIndex).flags.Ceguera = 1
2870              UserList(TargetIndex).Counters.Ceguera = IntervaloParalizado / 3
          
2880              Call WriteBlind(TargetIndex)
2890              Call FlushBuffer(TargetIndex)
2900              Call InfoHechizo(UserIndex)
2910              HechizoCasteado = True
2920      End If
          
          ' <-------- Agrega Estupidez (Aturdimiento) ---------->
2930      If Hechizos(HechizoIndex).Estupidez = 1 Then
2940          If UserIndex = TargetIndex Then
2950              Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
2960              Exit Sub
2970          End If
2980              If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Sub
2990              If UserIndex <> TargetIndex Then
3000                  Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
3010              End If
3020              If UserList(TargetIndex).flags.Estupidez = 0 Then
3030                  UserList(TargetIndex).flags.Estupidez = 1
3040                  UserList(TargetIndex).Counters.Ceguera = IntervaloParalizado
3050              End If
3060              Call WriteDumb(TargetIndex)
3070              Call FlushBuffer(TargetIndex)
          
3080              Call InfoHechizo(UserIndex)
3090              HechizoCasteado = True
3100      End If
3110  End With

End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, ByVal SpellIndex As Integer, ByRef HechizoCasteado As Boolean, ByVal UserIndex As Integer)
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 07/07/2008
      'Handles the Spells that afect the Stats of an NPC
      '04/13/2008 NicoNZ - Guardias Faccionarios pueden ser
      'removidos por users de su misma faccion.
      '07/07/2008: NicoNZ - Solo se puede mimetizar con npcs si es druida
      '***************************************************

10    With Npclist(NpcIndex)
20        If Hechizos(SpellIndex).Invisibilidad = 1 Then
30            Call InfoHechizo(UserIndex)
40            .flags.invisible = 1
50            HechizoCasteado = True
60        End If
          
70        If Hechizos(SpellIndex).Envenena = 1 Then
80            If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
90                HechizoCasteado = False
100               Exit Sub
110           End If
120           Call NPCAtacado(NpcIndex, UserIndex)
130           Call InfoHechizo(UserIndex)
140           .flags.Envenenado = 1
150           HechizoCasteado = True
160       End If
          
170       If Hechizos(SpellIndex).CuraVeneno = 1 Then
180           Call InfoHechizo(UserIndex)
190           .flags.Envenenado = 0
200           HechizoCasteado = True
210       End If
          
220       If Hechizos(SpellIndex).Maldicion = 1 Then
230           If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
240               HechizoCasteado = False
250               Exit Sub
260           End If
270           Call NPCAtacado(NpcIndex, UserIndex)
280           Call InfoHechizo(UserIndex)
290           .flags.Maldicion = 1
300           HechizoCasteado = True
310       End If
          
320       If Hechizos(SpellIndex).RemoverMaldicion = 1 Then
330           Call InfoHechizo(UserIndex)
340           .flags.Maldicion = 0
350           HechizoCasteado = True
360       End If
          
370       If Hechizos(SpellIndex).Bendicion = 1 Then
380           Call InfoHechizo(UserIndex)
390           .flags.Bendicion = 1
400           HechizoCasteado = True
410       End If
          
420       If Hechizos(SpellIndex).Paraliza = 1 Then
430           If .flags.AfectaParalisis = 0 Then
440               If Not PuedeAtacarNPC(UserIndex, NpcIndex, True) Then
450                   HechizoCasteado = False
460                   Exit Sub
470               End If
480               Call NPCAtacado(NpcIndex, UserIndex)
490               Call InfoHechizo(UserIndex)
500               .flags.Paralizado = 1
510               .flags.Inmovilizado = 0
520               .Contadores.Paralisis = IntervaloParalizado
530               HechizoCasteado = True
540           Else
550               Call WriteConsoleMsg(UserIndex, "El NPC es inmune a este hechizo.", FontTypeNames.FONTTYPE_INFO)
560               HechizoCasteado = False
570               Exit Sub
580           End If
590       End If
          
600       If Hechizos(SpellIndex).RemoverParalisis = 1 Then
610           If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
620               If .MaestroUser = UserIndex Then
630                   Call InfoHechizo(UserIndex)
640                   .flags.Paralizado = 0
650                   .Contadores.Paralisis = 0
660                   HechizoCasteado = True
670               Else
680                   If .NPCtype = eNPCType.GuardiaReal Then
690                       If esArmada(UserIndex) Then
700                           Call InfoHechizo(UserIndex)
710                           .flags.Paralizado = 0
720                           .Contadores.Paralisis = 0
730                           HechizoCasteado = True
740                           Exit Sub
750                       Else
760                           Call WriteConsoleMsg(UserIndex, "Sólo puedes remover la parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
770                           HechizoCasteado = False
780                           Exit Sub
790                       End If
                          
800                       Call WriteConsoleMsg(UserIndex, "Solo puedes remover la parálisis de los NPCs que te consideren su amo.", FontTypeNames.FONTTYPE_INFO)
810                       HechizoCasteado = False
820                       Exit Sub
830                   Else
840                       If .NPCtype = eNPCType.Guardiascaos Then
850                           If esCaos(UserIndex) Then
860                               Call InfoHechizo(UserIndex)
870                               .flags.Paralizado = 0
880                               .Contadores.Paralisis = 0
890                               HechizoCasteado = True
900                               Exit Sub
910                           Else
920                               Call WriteConsoleMsg(UserIndex, "Solo puedes remover la parálisis de los Guardias si perteneces a su facción.", FontTypeNames.FONTTYPE_INFO)
930                               HechizoCasteado = False
940                               Exit Sub
950                           End If
960                       End If
970                   End If
980               End If
990          Else
1000            Call WriteConsoleMsg(UserIndex, "Este NPC no está paralizado", FontTypeNames.FONTTYPE_INFO)
1010            HechizoCasteado = False
1020            Exit Sub
1030         End If
1040      End If
           
1050      If Hechizos(SpellIndex).Inmoviliza = 1 Then
1060          If .flags.AfectaParalisis = 0 Then
1070              If Not PuedeAtacarNPC(UserIndex, NpcIndex, True) Then
1080                  HechizoCasteado = False
1090                  Exit Sub
1100              End If
1110              Call NPCAtacado(NpcIndex, UserIndex)
1120              .flags.Inmovilizado = 1
1130              .flags.Paralizado = 0
1140              .Contadores.Paralisis = IntervaloParalizado
1150              Call InfoHechizo(UserIndex)
1160              HechizoCasteado = True
1170          Else
1180              Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
1190          End If
1200      End If
1210  End With

1220  If Hechizos(SpellIndex).Mimetiza = 1 Then
1230      With UserList(UserIndex)
1240          HechizoCasteado = False
1250          Exit Sub
              
1260          If MapInfo(.Pos.map).Pk = False Then
1270              WriteConsoleMsg UserIndex, "No puedes mimetizarte en zona segura.", FontTypeNames.FONTTYPE_INFO
1280              Exit Sub
1290          End If
              
1300          If .flags.Mimetizado = 1 Then
1310              Call WriteConsoleMsg(UserIndex, "Ya te encuentras mimetizado. El hechizo no ha tenido efecto.", FontTypeNames.FONTTYPE_INFO)
1320              Exit Sub
1330          End If
              
1340          If .flags.AdminInvisible = 1 Then Exit Sub
              
1350          If .Char.body = 0 Then
1360              WriteConsoleMsg UserIndex, "Esa criatura no tiene ningún cuerpo.", FontTypeNames.FONTTYPE_INFO
1370              Exit Sub
1380          End If
                  
1390          If .clase = eClass.Druid Then
                  'copio el char original al mimetizado
                  
1400              .CharMimetizado.body = .Char.body
1410              .CharMimetizado.Head = .Char.Head
1420              .CharMimetizado.CascoAnim = .Char.CascoAnim
1430              .CharMimetizado.ShieldAnim = .Char.ShieldAnim
1440              .CharMimetizado.WeaponAnim = .Char.WeaponAnim
                  
1450              .flags.Mimetizado = 1
                  
                  'ahora pongo lo del NPC.
1460              .Char.body = Npclist(NpcIndex).Char.body
1470              .Char.Head = Npclist(NpcIndex).Char.Head
1480              .Char.CascoAnim = NingunCasco
1490              .Char.ShieldAnim = NingunEscudo
1500              .Char.WeaponAnim = NingunArma
              
1510              Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
                  
1520          Else
1530              Call WriteConsoleMsg(UserIndex, "Sólo los druidas pueden mimetizarse con criaturas.", FontTypeNames.FONTTYPE_INFO)
1540              Exit Sub
1550          End If
          
1560         Call InfoHechizo(UserIndex)
1570         HechizoCasteado = True
1580      End With
1590  End If

End Sub

Sub HechizoPropNPC(ByVal SpellIndex As Integer, ByVal NpcIndex As Integer, ByVal UserIndex As Integer, ByRef HechizoCasteado As Boolean)
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 14/08/2007
      'Handles the Spells that afect the Life NPC
      '14/08/2007 Pablo (ToxicWaste) - Orden general.
      '***************************************************

      Dim daño As Long

10    With Npclist(NpcIndex)
          
20        If .flags.Paralizado And Hechizos(SpellIndex).Inmoviliza Then
30    HechizoCasteado = False
40    End If

50    If UserList(UserIndex).flags.Oculto = 1 Then
60    DecirPalabrasMagicas SpellIndex, UserIndex
70    UserList(UserIndex).flags.Oculto = 0
80    SetInvisible UserIndex, UserList(UserIndex).Char.CharIndex, False
90    End If

          'Salud
          
          
100       If Hechizos(SpellIndex).SubeHP = 1 Then

110           If MapInfo(.Pos.map).Pk = False Then
120     Call WriteConsoleMsg(UserIndex, "¡No puedes curar a este npc!", FontTypeNames.FONTTYPE_INFO)
130      Exit Sub
140      End If
         
150               If Npclist(NpcIndex).Stats.MinHp = .Stats.MaxHp Then
160       Call WriteConsoleMsg(UserIndex, "¡La criatura no está herida!", FontTypeNames.FONTTYPE_FIGHT)
170       Exit Sub
180       End If

190           daño = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
200           daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
              
210           Call InfoHechizo(UserIndex)
220           .Stats.MinHp = .Stats.MinHp + daño
230           If .Stats.MinHp > .Stats.MaxHp Then _
                  .Stats.MinHp = .Stats.MaxHp
240           Call WriteConsoleMsg(UserIndex, "Has curado " & daño & " puntos de vida a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
250           HechizoCasteado = True
              
260       ElseIf Hechizos(SpellIndex).SubeHP = 2 Then

270           If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
280               HechizoCasteado = False
290               Exit Sub
300           End If
310           Call NPCAtacado(NpcIndex, UserIndex)
320           daño = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
330           daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
          
340           If Hechizos(SpellIndex).StaffAffected Then
350               If UserList(UserIndex).clase = eClass.Mage Then
360                   If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
370                       daño = (daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                          'Aumenta daño segun el staff-
                          'Daño = (Daño* (70 + BonifBáculo)) / 100
380                   Else
390                       daño = daño * 0.7 'Baja daño a 70% del original
400                   End If
410               End If
420           End If
430             If UserList(UserIndex).Invent.MunicionEqpObjIndex = LAUDELFICO Or UserList(UserIndex).Invent.MunicionEqpObjIndex = FLAUTAELFICA Then
440               daño = daño * 1.04  'laud magico de los bardos
450           End If
460                            If UserList(UserIndex).Invent.MunicionEqpObjIndex = LaudBronce Or UserList(UserIndex).Invent.MunicionEqpObjIndex = AnilloBronce Then
470               daño = daño * 1.05  'laud magico de los bardos
480           End If
              
490                            If UserList(UserIndex).Invent.MunicionEqpObjIndex = LaudPlata Or UserList(UserIndex).Invent.MunicionEqpObjIndex = AnilloPlata Then
500               daño = daño * 1.06  'laud magico de los bardos
510           End If
520                     If UserList(UserIndex).Invent.MunicionEqpObjIndex = LAUDSUPERMAGICO Or UserList(UserIndex).Invent.MunicionEqpObjIndex = FLAUTAANTIGUA Then
530               daño = daño * 1.1   'laud magico de los bardos
540           End If
          
              ' 25% de aumento de daño mágico y 35% de aumento de daño mágico
580           If UserPoderoso(UserIndex) = 262 Then
590               daño = daño * 1.25
600           ElseIf UserPoderoso(UserIndex) = 263 Then
610               daño = daño * 1.35
620           End If
              
630           Call InfoHechizo(UserIndex)
640           HechizoCasteado = True
              
650           If .flags.Snd2 > 0 Then
660               Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(.flags.Snd2, .Pos.X, .Pos.Y))
670           End If
              
              'Quizas tenga defenza magica el NPC. Pablo (ToxicWaste)
680           daño = daño - .Stats.defM
690           If daño < 0 Then daño = 0
              
700           .Stats.MinHp = .Stats.MinHp - daño
710           SendData SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, DAMAGE_NORMAL)
720           Call WriteConsoleMsg(UserIndex, "¡Le has quitado " & daño & " puntos de vida a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
730           Call CalcularDarExp(UserIndex, NpcIndex, daño)

              
740           If .Stats.MinHp < 1 Then
750               .Stats.MinHp = 0
760               Call MuereNpc(NpcIndex, UserIndex)
770           End If
780           End If
790   End With

End Sub

Sub InfoHechizo(ByVal UserIndex As Integer)
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 25/07/2009
      '25/07/2009: ZaMa - Code improvements.
      '25/07/2009: ZaMa - Now invisible admins magic sounds are not sent to anyone but themselves
      '***************************************************
          Dim SpellIndex As Integer
          Dim tUser As Integer
          Dim tNpc As Integer
          
10        With UserList(UserIndex)
20            SpellIndex = .flags.Hechizo
30            tUser = .flags.TargetUser
40            tNpc = .flags.TargetNPC
              
50            Call DecirPalabrasMagicas(SpellIndex, UserIndex)
              
60            If tUser > 0 Then
                  ' Los admins invisibles no producen sonidos ni fx's
70                If .flags.AdminInvisible = 1 And UserIndex = tUser Then
80                    Call EnviarDatosASlot(UserIndex, PrepareMessageCreateFX(UserList(tUser).Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
90                    Call EnviarDatosASlot(UserIndex, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, UserList(tUser).Pos.X, UserList(tUser).Pos.Y))
100               Else
110                   Call SendData(SendTarget.ToPCArea, tUser, PrepareMessageCreateFX(UserList(tUser).Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
120                   Call SendData(SendTarget.ToPCArea, tUser, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, UserList(tUser).Pos.X, UserList(tUser).Pos.Y)) 'Esta linea faltaba. Pablo (ToxicWaste)
130               End If
140           ElseIf tNpc > 0 Then
150               Call SendData(SendTarget.ToNPCArea, tNpc, PrepareMessageCreateFX(Npclist(tNpc).Char.CharIndex, Hechizos(SpellIndex).FXgrh, Hechizos(SpellIndex).loops))
160               Call SendData(SendTarget.ToNPCArea, tNpc, PrepareMessagePlayWave(Hechizos(SpellIndex).WAV, Npclist(tNpc).Pos.X, Npclist(tNpc).Pos.Y))
170           End If
              
180           If tUser > 0 Then
190               If UserIndex <> tUser Then
200                   If .showName And Not UserList(tUser).flags.SlotEvent > 0 Then
                          'Call WriteConsoleMsg(Userindex, Hechizos(SpellIndex).HechizeroMsg & " " & UserList(tUser).Name, FontTypeNames.FONTTYPE_FIGHT)
210                       Call WriteShortMsj(UserIndex, 0, FontTypeNames.FONTTYPE_FIGHT, SpellIndex, , , , UserList(tUser).Name)
220                   Else
                          'Call WriteConsoleMsg(Userindex, Hechizos(SpellIndex).HechizeroMsg & " alguien.", FontTypeNames.FONTTYPE_FIGHT)
230                       Call WriteShortMsj(UserIndex, 1, FontTypeNames.FONTTYPE_FIGHT, SpellIndex)
240                   End If
                      
                      'Call WriteConsoleMsg(tUser, .Name & " " & Hechizos(SpellIndex).TargetMsg, FontTypeNames.FONTTYPE_FIGHT)
250                   Call WriteShortMsj(tUser, 2, FontTypeNames.FONTTYPE_FIGHT, SpellIndex, , , , .Name)
260               Else
                      'Call WriteConsoleMsg(Userindex, Hechizos(SpellIndex).PropioMsg, FontTypeNames.FONTTYPE_FIGHT)
270                   Call WriteShortMsj(UserIndex, 3, FontTypeNames.FONTTYPE_FIGHT, SpellIndex)
280               End If
290           ElseIf tNpc > 0 Then
                  'Call WriteConsoleMsg(Userindex, Hechizos(SpellIndex).HechizeroMsg & " " & "la criatura.", FontTypeNames.FONTTYPE_FIGHT)
300               Call WriteShortMsj(UserIndex, 4, FontTypeNames.FONTTYPE_FIGHT, SpellIndex)
310           End If
320       End With

End Sub

Public Function HechizoPropUsuario(ByVal UserIndex As Integer) As Boolean
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 28/04/2010
      '02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
      '28/04/2010: ZaMa - Agrego Restricciones para ciudas respecto al estado atacable.
      '***************************************************

      Dim SpellIndex As Integer
      Dim daño As Long
      Dim TargetIndex As Integer

10    SpellIndex = UserList(UserIndex).flags.Hechizo
20    TargetIndex = UserList(UserIndex).flags.TargetUser
            
30    With UserList(TargetIndex)
40        If .flags.Muerto Then
50            Call WriteConsoleMsg(UserIndex, "No puedes lanzar este hechizo a un muerto.", FontTypeNames.FONTTYPE_INFO)
60            Exit Function
70        End If
                
          ' <-------- Aumenta Hambre ---------->
80        If Hechizos(SpellIndex).SubeHam = 1 Then
              
90            Call InfoHechizo(UserIndex)
              
100           daño = RandomNumber(Hechizos(SpellIndex).MinHam, Hechizos(SpellIndex).MaxHam)
              
110           .Stats.MinHam = .Stats.MinHam + daño
120           If .Stats.MinHam > .Stats.MaxHam Then _
                  .Stats.MinHam = .Stats.MaxHam
              
130           If UserIndex <> TargetIndex Then
140               Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de hambre a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
150               Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
160           Else
170               Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
180           End If
              
190           Call WriteUpdateHungerAndThirst(TargetIndex)
          
          ' <-------- Quita Hambre ---------->
200       ElseIf Hechizos(SpellIndex).SubeHam = 2 Then
210           If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
              
              If .flags.SlotEvent > 0 Or .flags.SlotReto > 0 Then Exit Function
              
220           If UserIndex <> TargetIndex Then
230               Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
240           Else
250               Exit Function
260           End If
              
270           Call InfoHechizo(UserIndex)
              
280           daño = RandomNumber(Hechizos(SpellIndex).MinHam, Hechizos(SpellIndex).MaxHam)
              
290           .Stats.MinHam = .Stats.MinHam - daño
              
300           If UserIndex <> TargetIndex Then
310               Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de hambre a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
320               Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
330           Else
340               Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
350           End If
              
360           If .Stats.MinHam < 1 Then
370               .Stats.MinHam = 0
380               .flags.Hambre = 1
390           End If
              
400           Call WriteUpdateHungerAndThirst(TargetIndex)
410       End If
          
          ' <-------- Aumenta Sed ---------->
420       If Hechizos(SpellIndex).SubeSed = 1 Then
              
430           Call InfoHechizo(UserIndex)
              
440           daño = RandomNumber(Hechizos(SpellIndex).MinSed, Hechizos(SpellIndex).MaxSed)
              
450           .Stats.MinAGU = .Stats.MinAGU + daño
460           If .Stats.MinAGU > .Stats.MaxAGU Then _
                  .Stats.MinAGU = .Stats.MaxAGU
              
470           Call WriteUpdateHungerAndThirst(TargetIndex)
                   
480           If UserIndex <> TargetIndex Then
490             Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de sed a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
500             Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
510           Else
520             Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
530           End If
              
          
          ' <-------- Quita Sed ---------->
540       ElseIf Hechizos(SpellIndex).SubeSed = 2 Then
              
550           If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
              
              If .flags.SlotEvent > 0 Or .flags.SlotReto > 0 Then Exit Function
              
560           If UserIndex <> TargetIndex Then
570               Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
580           End If
              
590           Call InfoHechizo(UserIndex)
              
600           daño = RandomNumber(Hechizos(SpellIndex).MinSed, Hechizos(SpellIndex).MaxSed)
              
610           .Stats.MinAGU = .Stats.MinAGU - daño
              
620           If UserIndex <> TargetIndex Then
630               Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de sed a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
640               Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
650           Else
660               Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
670           End If
              
680           If .Stats.MinAGU < 1 Then
690               .Stats.MinAGU = 0
700               .flags.Sed = 1
710           End If
              
720           Call WriteUpdateHungerAndThirst(TargetIndex)
              
730       End If
          
          ' <-------- Aumenta Agilidad ---------->
740       If Hechizos(SpellIndex).SubeAgilidad = 1 Then
              
              ' Chequea si el status permite ayudar al otro usuario
750           If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
              
760           Call InfoHechizo(UserIndex)
770           daño = RandomNumber(Hechizos(SpellIndex).MinAgilidad, Hechizos(SpellIndex).MaxAgilidad)
              
780           .flags.DuracionEfecto = 1200
790           .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + daño
800           If .Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Agilidad) * 2) Then _
                  .Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Agilidad) * 2)
              
810           .flags.TomoPocion = True
820           Call WriteUpdateDexterity(TargetIndex)
          
          ' <-------- Quita Agilidad ---------->
830       ElseIf Hechizos(SpellIndex).SubeAgilidad = 2 Then
              
840           If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
              
850           If UserIndex <> TargetIndex Then
860               Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
870           End If
              
880           Call InfoHechizo(UserIndex)
              
890           .flags.TomoPocion = True
900           daño = RandomNumber(Hechizos(SpellIndex).MinAgilidad, Hechizos(SpellIndex).MaxAgilidad)
910           .flags.DuracionEfecto = 700
920           .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) - daño
930           If .Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then .Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
              
940           Call WriteUpdateDexterity(TargetIndex)
950       End If
          
          ' <-------- Aumenta Fuerza ---------->
960       If Hechizos(SpellIndex).SubeFuerza = 1 Then
          
              ' Chequea si el status permite ayudar al otro usuario
970           If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
              
980           Call InfoHechizo(UserIndex)
990           daño = RandomNumber(Hechizos(SpellIndex).MinFuerza, Hechizos(SpellIndex).MaxFuerza)
              
1000          .flags.DuracionEfecto = 1200
          
1010          .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + daño
1020          If .Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Fuerza) * 2) Then _
                  .Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Fuerza) * 2)
              
1030          .flags.TomoPocion = True
1040          Call WriteUpdateStrenght(TargetIndex)
          
          ' <-------- Quita Fuerza ---------->
1050      ElseIf Hechizos(SpellIndex).SubeFuerza = 2 Then
          
1060          If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
              
1070          If UserIndex <> TargetIndex Then
1080              Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
1090          End If
              
1100          Call InfoHechizo(UserIndex)
              
1110          .flags.TomoPocion = True
              
1120          daño = RandomNumber(Hechizos(SpellIndex).MinFuerza, Hechizos(SpellIndex).MaxFuerza)
1130          .flags.DuracionEfecto = 700
1140          .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) - daño
1150          If .Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then .Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
              
1160          Call WriteUpdateStrenght(TargetIndex)
1170      End If
          
          ' <-------- Cura salud ---------->
1180      If Hechizos(SpellIndex).SubeHP = 1 Then
              
1190          If UserList(UserIndex).Stats.MinHp = .Stats.MaxHp Then
1200              Call WriteConsoleMsg(UserIndex, "¡No estás herido!", FontTypeNames.FONTTYPE_FIGHT)
1210              Exit Function
1220              End If
          
1230           If UserList(TargetIndex).Stats.MinHp = .Stats.MaxHp Then
1240              Call WriteConsoleMsg(UserIndex, "¡No está herido!", FontTypeNames.FONTTYPE_FIGHT)
1250              Exit Function
1260              End If
              
1270          If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = eTrigger.ZONAPELEA Then
1280              Call WriteConsoleMsg(UserIndex, "No puedes curar a este usuario.", FontTypeNames.FONTTYPE_INFO)
1290              Exit Function
1300          End If
              
              'Verifica que el usuario no este muerto
1310          If .flags.Muerto = 1 Then
1320              Call WriteConsoleMsg(UserIndex, "¡El usuario está muerto!", FontTypeNames.FONTTYPE_INFO)
1330              Exit Function
1340          End If
              
              ' Chequea si el status permite ayudar al otro usuario
1350          If Not CanSupportUser(UserIndex, TargetIndex) Then Exit Function
                 
1360          daño = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
1370          daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
              
1380          Call InfoHechizo(UserIndex)
          
1390          .Stats.MinHp = .Stats.MinHp + daño
1400          If .Stats.MinHp > .Stats.MaxHp Then _
                  .Stats.MinHp = .Stats.MaxHp
              
1410          Call WriteUpdateHP(TargetIndex)
1420          Call WriteUpdateFollow(TargetIndex)
              
1430          If UserIndex <> TargetIndex Then
1440              Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de vida a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
1450              Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
1460          Else
1470              Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
1480          End If
              
          ' <-------- Quita salud (Daña) ---------->
1490      ElseIf Hechizos(SpellIndex).SubeHP = 2 Then
              
1500          If UserIndex = TargetIndex Then
1510              Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
1520              Exit Function
1530          End If
              
1540          daño = RandomNumber(Hechizos(SpellIndex).MinHp, Hechizos(SpellIndex).MaxHp)
              
1550          daño = daño + Porcentaje(daño, 3 * UserList(UserIndex).Stats.ELV)
              
1560          If UserList(UserIndex).flags.SlotEvent > 0 Then
1570              If Events(UserList(UserIndex).flags.SlotEvent).Modality = HombreLobo Then
1580                  If Events(UserList(UserIndex).flags.SlotEvent).Users(UserList(UserIndex).flags.SlotUserEvent).Selected = 1 Then
1590                      daño = daño * 1.5
1600                  End If
1610              End If
1620          End If
              
1630          If Hechizos(SpellIndex).StaffAffected Then
1640              If UserList(UserIndex).clase = eClass.Mage Then
1650                  If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
1660                      daño = (daño * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
1670                  Else
1680                      daño = daño * 0.7 'Baja daño a 70% del original
1690                  End If
1700              End If
1710          End If
              
1720          If UserList(UserIndex).Invent.MunicionEqpObjIndex = LAUDELFICO Or UserList(UserIndex).Invent.MunicionEqpObjIndex = FLAUTAELFICA Then
1730              daño = daño * 1.04  'laud magico de los bardos
1740          End If
              
1750                   If UserList(UserIndex).Invent.MunicionEqpObjIndex = LaudBronce Or UserList(UserIndex).Invent.MunicionEqpObjIndex = AnilloBronce Then
1760              daño = daño * 1.05  'laud magico de los bardos
1770          End If
              
1780                           If UserList(UserIndex).Invent.MunicionEqpObjIndex = LaudPlata Or UserList(UserIndex).Invent.MunicionEqpObjIndex = AnilloPlata Then
1790              daño = daño * 1.06  'laud magico de los bardos
1800          End If
              
1810           If UserList(UserIndex).Invent.MunicionEqpObjIndex = LAUDSUPERMAGICO Or UserList(UserIndex).Invent.MunicionEqpObjIndex = FLAUTAANTIGUA Then
1820              daño = daño * 1.08  'laud magico de los bardos
1830          End If
              
              'cascos antimagia
1840          If (.Invent.CascoEqpObjIndex > 0) Then
1850              daño = daño - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)
1860          End If
              
              'anillos
1870          If (.Invent.AnilloEqpObjIndex > 0) Then
1880              daño = daño - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)
1890          End If
              
              ' Efectos atacante [DIOS]
1960          If UserPoderoso(UserIndex) = 263 Then
1970              daño = daño * 1.05
1980          End If
              
              ' Efectos VICTIMA [Dios]
1990          If UserPoderoso(TargetIndex) = 262 Then
2000              daño = daño * 0.95
2010          End If
              
2020          daño = daño - (daño * UserList(TargetIndex).Stats.UserSkills(eSkill.Resistencia) / 2000)
              
2030          If daño < 0 Then daño = 0
              
2040          If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
              
2050          If UserIndex <> TargetIndex Then
2060              Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
2070          End If
              
2080          Call InfoHechizo(UserIndex)
              
2090          .Stats.MinHp = .Stats.MinHp - daño
2100          SendData SendTarget.ToPCArea, TargetIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, DAMAGE_NORMAL)
2110          Call WriteUpdateHP(TargetIndex)
2120          Call WriteUpdateFollow(TargetIndex)
              
2130          Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de vida a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
2140          Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
              
              'Muere
2150          If .Stats.MinHp < 1 Then
              
2160              If .flags.AtacablePor <> UserIndex Then
                      'Store it!
                      'Call Statistics.StoreFrag(Userindex, TargetIndex)
2170                  Call ContarMuerte(TargetIndex, UserIndex)
2180              End If
                  
2190              .Stats.MinHp = 0
2200              Call ActStats(TargetIndex, UserIndex)
2210              Call UserDie(TargetIndex, UserIndex)
2220          End If
              
2230      End If
          
          ' <-------- Aumenta Mana ---------->
2240      If Hechizos(SpellIndex).SubeMana = 1 Then
              
2250          Call InfoHechizo(UserIndex)
2260          .Stats.MinMAN = .Stats.MinMAN + daño
2270          If .Stats.MinMAN > .Stats.MaxMAN Then _
                  .Stats.MinMAN = .Stats.MaxMAN
              
2280          Call WriteUpdateMana(TargetIndex)
2290          Call WriteUpdateFollow(TargetIndex)
              
2300          If UserIndex <> TargetIndex Then
2310              Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de maná a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
2320              Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT)
2330          Else
2340              Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT)
2350          End If
              
          
          ' <-------- Quita Mana ---------->
2360      ElseIf Hechizos(SpellIndex).SubeMana = 2 Then
2370          If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
              
2380          If UserIndex <> TargetIndex Then
2390              Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
2400          End If
              
2410          Call InfoHechizo(UserIndex)
              
2420          If UserIndex <> TargetIndex Then
2430              Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de maná a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
2440              Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT)
2450          Else
2460              Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de maná.", FontTypeNames.FONTTYPE_FIGHT)
2470          End If
              
2480          .Stats.MinMAN = .Stats.MinMAN - daño
2490          If .Stats.MinMAN < 1 Then .Stats.MinMAN = 0
              
2500          Call WriteUpdateMana(TargetIndex)
2510          Call WriteUpdateFollow(TargetIndex)
              
2520      End If
          
          ' <-------- Aumenta Stamina ---------->
2530      If Hechizos(SpellIndex).SubeSta = 1 Then
2540          Call InfoHechizo(UserIndex)
2550          .Stats.MinSta = .Stats.MinSta + daño
2560          If .Stats.MinSta > .Stats.MaxSta Then _
                  .Stats.MinSta = .Stats.MaxSta
              
2570          Call WriteUpdateSta(TargetIndex)
              
2580          If UserIndex <> TargetIndex Then
2590              Call WriteConsoleMsg(UserIndex, "Le has restaurado " & daño & " puntos de energía a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
2600              Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha restaurado " & daño & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT)
2610          Else
2620              Call WriteConsoleMsg(UserIndex, "Te has restaurado " & daño & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT)
2630          End If
              
          ' <-------- Quita Stamina ---------->
2640      ElseIf Hechizos(SpellIndex).SubeSta = 2 Then
2650          If Not PuedeAtacar(UserIndex, TargetIndex) Then Exit Function
              
2660          If UserIndex <> TargetIndex Then
2670              Call UsuarioAtacadoPorUsuario(UserIndex, TargetIndex)
2680          End If
              
2690          Call InfoHechizo(UserIndex)
              
2700          If UserIndex <> TargetIndex Then
2710              Call WriteConsoleMsg(UserIndex, "Le has quitado " & daño & " puntos de energía a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
2720              Call WriteConsoleMsg(TargetIndex, UserList(UserIndex).Name & " te ha quitado " & daño & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT)
2730          Else
2740              Call WriteConsoleMsg(UserIndex, "Te has quitado " & daño & " puntos de energía.", FontTypeNames.FONTTYPE_FIGHT)
2750          End If
              
2760          .Stats.MinSta = .Stats.MinSta - daño
              
2770          If .Stats.MinSta < 1 Then .Stats.MinSta = 0
              
2780          Call WriteUpdateSta(TargetIndex)
              
2790      End If
2800  End With

2810  HechizoPropUsuario = True

2820  Call FlushBuffer(TargetIndex)

End Function

Public Function CanSupportUser(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer, _
                               Optional ByVal DoCriminal As Boolean = False) As Boolean
      '***************************************************
      'Author: ZaMa
      'Last Modification: 28/04/2010
      'Checks if caster can cast support magic on target user.
      '***************************************************
           
10     On Error GoTo Errhandler
       
20        With UserList(CasterIndex)
              
              ' Te podes curar a vos mismo
30            If CasterIndex = TargetIndex Then
40                CanSupportUser = True
50                Exit Function
60            End If
              
               ' No podes ayudar si estas en consulta
70            If .flags.EnConsulta Then
80                Call WriteConsoleMsg(CasterIndex, "No puedes ayudar usuarios mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
90                Exit Function
100           End If
              
              ' Si estas en la arena, esta todo permitido
110           If TriggerZonaPelea(CasterIndex, TargetIndex) = TRIGGER6_PERMITE Then
120               CanSupportUser = True
130               Exit Function
140           End If
           
              ' Victima criminal?
150           If criminal(TargetIndex) Then
              
                  ' Casteador Ciuda?
160               If Not criminal(CasterIndex) Then
                  
                      ' Armadas no pueden ayudar
170                   If esArmada(CasterIndex) Then
180                       Call WriteConsoleMsg(CasterIndex, "Los miembros del ejército real no pueden ayudar a los criminales.", FontTypeNames.FONTTYPE_INFO)
190                       Exit Function
200                   End If
                      
                      ' Si el ciuda tiene el seguro puesto no puede ayudar
210                   If .flags.Seguro Then
220                       Call WriteConsoleMsg(CasterIndex, "Para ayudar criminales debes sacarte el seguro ya que te volverás criminal como ellos.", FontTypeNames.FONTTYPE_INFO)
230                       Exit Function
240                   Else
                          ' Penalizacion
250                       If DoCriminal Then
260                           Call VolverCriminal(CasterIndex)
270                       Else
280                           Call DisNobAuBan(CasterIndex, .Reputacion.NobleRep * 0.5, 10000)
290                       End If
300                   End If
310               End If
                  
              ' Victima ciuda o army
320           Else
                  ' Casteador es caos? => No Pueden ayudar ciudas
330               If esCaos(CasterIndex) Then
340                   Call WriteConsoleMsg(CasterIndex, "Los miembros de la legión oscura no pueden ayudar a los ciudadanos.", FontTypeNames.FONTTYPE_INFO)
350                   Exit Function
                      
                  ' Casteador ciuda/army?
360               ElseIf Not criminal(CasterIndex) Then
                      
                      ' Esta en estado atacable?
370                   If UserList(TargetIndex).flags.AtacablePor > 0 Then
                          
                          ' No esta atacable por el casteador?
380                       If UserList(TargetIndex).flags.AtacablePor <> CasterIndex Then
                          
                              ' Si es armada no puede ayudar
390                           If esArmada(CasterIndex) Then
400                               Call WriteConsoleMsg(CasterIndex, "Los miembros del ejército real no pueden ayudar a ciudadanos en estado atacable.", FontTypeNames.FONTTYPE_INFO)
410                               Exit Function
420                           End If
          
                              ' Seguro puesto?
430                           If .flags.Seguro Then
440                               Call WriteConsoleMsg(CasterIndex, "Para ayudar ciudadanos en estado atacable debes sacarte el seguro, pero te puedes volver criminal.", FontTypeNames.FONTTYPE_INFO)
450                               Exit Function
460                           Else
470                               Call DisNobAuBan(CasterIndex, .Reputacion.NobleRep * 0.5, 10000)
480                           End If
490                       End If
500                   End If
          
510               End If
520           End If
530       End With
          
540       CanSupportUser = True

550       Exit Function
          
Errhandler:
560       Call LogError("Error en CanSupportUser, Error: " & Err.Number & " - " & Err.Description & _
                        " CasterIndex: " & CasterIndex & ", TargetIndex: " & TargetIndex)

End Function

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, ByVal UserIndex As Integer, ByVal Slot As Byte)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      Dim LoopC As Byte

10    With UserList(UserIndex)
          'Actualiza un solo slot
20        If Not UpdateAll Then
              'Actualiza el inventario
30            If .Stats.UserHechizos(Slot) > 0 Then
40                Call ChangeUserHechizo(UserIndex, Slot, .Stats.UserHechizos(Slot))
50            Else
60                Call ChangeUserHechizo(UserIndex, Slot, 0)
70            End If
80        Else
              'Actualiza todos los slots
90            For LoopC = 1 To MAXUSERHECHIZOS
                  'Actualiza el inventario
100               If .Stats.UserHechizos(LoopC) > 0 Then
110                   Call ChangeUserHechizo(UserIndex, LoopC, .Stats.UserHechizos(LoopC))
120               Else
130                   Call ChangeUserHechizo(UserIndex, LoopC, 0)
140               End If
150           Next LoopC
160       End If
170   End With

End Sub

Sub ChangeUserHechizo(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal Hechizo As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************
          
10        UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo
          
20        If Hechizo > 0 And Hechizo < NumeroHechizos + 1 Then
30            Call WriteChangeSpellSlot(UserIndex, Slot)
40        Else
50            Call WriteChangeSpellSlot(UserIndex, Slot)
60        End If

End Sub


Public Sub DesplazarHechizo(ByVal UserIndex As Integer, ByVal Dire As Integer, ByVal HechizoDesplazado As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    If (Dire <> 1 And Dire <> -1) Then Exit Sub
20    If Not (HechizoDesplazado >= 1 And HechizoDesplazado <= MAXUSERHECHIZOS) Then Exit Sub

      Dim TempHechizo As Integer

30    With UserList(UserIndex)
40        If Dire = 1 Then 'Mover arriba
50            If HechizoDesplazado = 1 Then
60                Call WriteConsoleMsg(UserIndex, "No puedes mover el hechizo en esa dirección.", FontTypeNames.FONTTYPE_INFO)
70                Exit Sub
80            Else
90                TempHechizo = .Stats.UserHechizos(HechizoDesplazado)
100               .Stats.UserHechizos(HechizoDesplazado) = .Stats.UserHechizos(HechizoDesplazado - 1)
110               .Stats.UserHechizos(HechizoDesplazado - 1) = TempHechizo
120           End If
130       Else 'mover abajo
140           If HechizoDesplazado = MAXUSERHECHIZOS Then
150               Call WriteConsoleMsg(UserIndex, "No puedes mover el hechizo en esa dirección.", FontTypeNames.FONTTYPE_INFO)
160               Exit Sub
170           Else
180               TempHechizo = .Stats.UserHechizos(HechizoDesplazado)
190               .Stats.UserHechizos(HechizoDesplazado) = .Stats.UserHechizos(HechizoDesplazado + 1)
200               .Stats.UserHechizos(HechizoDesplazado + 1) = TempHechizo
210           End If
220       End If
230   End With

End Sub

Public Sub DisNobAuBan(ByVal UserIndex As Integer, NoblePts As Long, BandidoPts As Long)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          'disminuye la nobleza NoblePts puntos y aumenta el bandido BandidoPts puntos
          Dim EraCriminal As Boolean
10        EraCriminal = criminal(UserIndex)
          
20        With UserList(UserIndex)
              'Si estamos en la arena no hacemos nada
30            If MapData(.Pos.map, .Pos.X, .Pos.Y).trigger = 6 Then Exit Sub
              
40            If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
                  'pierdo nobleza...
50                .Reputacion.NobleRep = .Reputacion.NobleRep - NoblePts
60                If .Reputacion.NobleRep < 0 Then
70                    .Reputacion.NobleRep = 0
80                End If
                  
                  'gano bandido...
90                .Reputacion.BandidoRep = .Reputacion.BandidoRep + BandidoPts
100               If .Reputacion.BandidoRep > MAXREP Then _
                      .Reputacion.BandidoRep = MAXREP
110               Call WriteMultiMessage(UserIndex, eMessages.NobilityLost) 'Call WriteNobilityLost(UserIndex)
120               If criminal(UserIndex) Then If .Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)
130           End If
              
140           If Not EraCriminal And criminal(UserIndex) Then
150               Call RefreshCharStatus(UserIndex)
160           End If
170       End With
End Sub

Public Function CanSupportNpc(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer) As Boolean
      '***************************************************
      'Author: ZaMa
      'Last Modification: 18/09/2010
      'Checks if caster can cast support magic on target Npc.
      '***************************************************
           
10     On Error GoTo Errhandler
       
          Dim OwnerIndex As Integer
       
20        With UserList(CasterIndex)
              
30            OwnerIndex = Npclist(TargetIndex).Owner
              
              ' Si no tiene dueño puede
40            If OwnerIndex = 0 Then
50                CanSupportNpc = True
60                Exit Function
70            End If
              
              ' Puede hacerlo si es su propio npc
80            If CasterIndex = OwnerIndex Then
90                CanSupportNpc = True
100               Exit Function
110           End If
              
               ' No podes ayudar si estas en consulta
120           If .flags.EnConsulta Then
130               Call WriteConsoleMsg(CasterIndex, "No puedes ayudar npcs mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
140               Exit Function
150           End If
              
              ' Si estas en la arena, esta todo permitido
160           If TriggerZonaPelea(CasterIndex, OwnerIndex) = TRIGGER6_PERMITE Then
170               CanSupportNpc = True
180               Exit Function
190           End If
           
              ' Victima criminal?
200           If criminal(OwnerIndex) Then
                  ' Victima caos?
210               If esCaos(OwnerIndex) Then
                      ' Atacante caos?
220                   If esCaos(CasterIndex) Then
                          ' No podes ayudar a un npc de un caos si sos caos
230                       Call WriteConsoleMsg(CasterIndex, "No puedes ayudar npcs que están luchando contra un miembro de tu facción.", FontTypeNames.FONTTYPE_INFO)
240                       Exit Function
250                   End If
260               End If
              
                  ' Uno es caos y el otro no, o la victima es pk, entonces puede ayudar al npc
270               CanSupportNpc = True
280               Exit Function
                      
              ' Victima ciuda
290           Else
                  ' Atacante ciuda?
300               If Not criminal(CasterIndex) Then
                      ' Atacante armada?
310                   If esArmada(CasterIndex) Then
                          ' Victima armada?
320                       If esArmada(OwnerIndex) Then
                              ' No podes ayudar a un npc de un armada si sos armada
330                           Call WriteConsoleMsg(CasterIndex, "No puedes ayudar npcs que están luchando contra un miembro de tu facción.", FontTypeNames.FONTTYPE_INFO)
340                           Exit Function
350                       End If
360                   End If
                      
                      ' Uno es armada y el otro ciuda, o los dos ciudas, puede atacar si no tiene seguro
370                   If .flags.Seguro Then
380                       Call WriteConsoleMsg(CasterIndex, "Para ayudar a criaturas que luchan contra ciudadanos debes sacarte el seguro.", FontTypeNames.FONTTYPE_INFO)
390                       Exit Function
                          
                      ' ayudo al npc sin seguro, se convierte en atacable
400                   Else
410                       Call ToogleToAtackable(CasterIndex, OwnerIndex, True)
420                       CanSupportNpc = True
430                       Exit Function
440                   End If
                      
450               End If
                  
                  ' Atacante criminal y victima ciuda, entonces puede ayudar al npc
460               CanSupportNpc = True
470               Exit Function
                  
480           End If
          
490       End With
          
500       CanSupportNpc = True

510       Exit Function
          
Errhandler:
520       Call LogError("Error en CanSupportNpc, Error: " & Err.Number & " - " & Err.Description & _
                        " CasterIndex: " & CasterIndex & ", OwnerIndex: " & OwnerIndex)

End Function

Function ResistenciaClase(clase As String) As Integer
      Dim Cuan As Integer
10    Select Case UCase$(clase)
          Case "MAGO"
20            Cuan = 3
30        Case "DRUIDA"
40            Cuan = 2
50        Case "CLERIGO"
60            Cuan = 2
70        Case "BARDO"
80            Cuan = 1
90        Case Else
100           Cuan = 0
110   End Select
120   ResistenciaClase = Cuan
End Function

