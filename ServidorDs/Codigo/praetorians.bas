Attribute VB_Name = "PraetoriansCoopNPC"
'**************************************************************
' PraetoriansCoopNPC.bas - Handles the Praeorians NPCs.
'
' Implemented by Mariano Barrou (El Oso)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit
'''''''''''''''''''''''''''''''''''''''''
'' DECLARACIONES DEL MODULO PRETORIANO ''
'''''''''''''''''''''''''''''''''''''''''
'' Estas constantes definen que valores tienen
'' los NPCs pretorianos en el NPC-HOSTILES.DAT
'' Son FIJAS, pero se podria hacer una rutina que
'' las lea desde el npcshostiles.dat
Public Const PRCLER_NPC As Integer = 900   ''"Sacerdote Pretoriano"
Public Const PRGUER_NPC As Integer = 901   ''"Guerrero  Pretoriano"
Public Const PRMAGO_NPC As Integer = 902   ''"Mago Pretoriano"
Public Const PRCAZA_NPC As Integer = 903   ''"Cazador Pretoriano"
Public Const PRKING_NPC As Integer = 904   ''"Rey Pretoriano"
Public Const PRDROP_NPC As Integer = 905
Public Const PRDROP2_NPC As Integer = 906
Public Const PRDROP3_NPC As Integer = 907
Public Const PRDROP4_NPC As Integer = 908
''''''''''''''''''''''''''''''''''''''''''''''
''Esta constante identifica en que mapa esta
''la fortaleza pretoriana (no es lo mismo de
''donde estan los NPCs!).
''Se extrae el dato del server.ini en sub LoadSIni
Public MAPA_PRETORIANO As Integer
''''''''''''''''''''''''''''''''''''''''''''''
''Estos numeros son necesarios por cuestiones de
''sonido. Son los numeros de los wavs del cliente.
Private Const SONIDO_Dragon_VIVO As Integer = 30
''ALCOBAS REALES
''OJO LOS BICHOS TAN HARDCODEADOS, NO CAMBIAR EL MAPA DONDE
''ESTÁN UBICADOS!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
''MUCHO MENOS LA COORDENADA Y DE LAS ALCOBAS YA QUE DEBE SER LA MISMA!!!
''(HAY FUNCIONES Q CUENTAN CON QUE ES LA MISMA!)
Public Const ALCOBA1_X As Integer = 35
Public Const ALCOBA1_Y As Integer = 25
Public Const ALCOBA2_X As Integer = 67
Public Const ALCOBA2_Y As Integer = 25

'Added by Nacho
'Cuantos pretorianos vivos quedan. Uno por cada alcoba
Public pretorianosVivos As Integer

'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'/\/\/\/\/\/\/\/\ MODULO DE COMBATE PRETORIANO /\/\/\/\/\/\/\/\/\
'/\/\/\/\/\/\/\/\ (NPCS COOPERATIVOS TIPO CLAN)/\/\/\/\/\/\/\/\/\
'/\/\/\/\/\/\/\/\         por EL OSO           /\/\/\/\/\/\/\/\/\
'/\/\/\/\/\/\/\/\       mbarrou@dc.uba.ar      /\/\/\/\/\/\/\/\/\
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\

Public Function esPretoriano(ByVal NpcIndex As Integer) As Integer
10    On Error GoTo errorh

          Dim N As Integer
          Dim i As Integer
20        N = Npclist(NpcIndex).Numero
30        i = Npclist(NpcIndex).Char.CharIndex
      '    Call SendData(SendTarget.ToNPCArea, NpcIndex, Npclist(NpcIndex).Pos.Map, "||" & vbGreen & "° Soy Pretoriano °" & Str(ind))
40        Select Case Npclist(NpcIndex).Numero
          Case PRCLER_NPC
50            esPretoriano = 1
60        Case PRMAGO_NPC
70            esPretoriano = 2
80        Case PRCAZA_NPC
90            esPretoriano = 3
100       Case PRKING_NPC
110           esPretoriano = 4
120       Case PRGUER_NPC
130           esPretoriano = 5
140       End Select

150   Exit Function

errorh:
160       LogError ("Error en NPCAI.EsPretoriano? " & Npclist(NpcIndex).Name)
          'do nothing

End Function


Sub CrearClanPretoriano(ByVal X As Integer)
      '********************************************************
      'Author: EL OSO
      'Inicializa el clan Pretoriano.
      'Last Modify Date: 22/6/06: (Nacho) Seteamos cuantos NPCs creamos
      '********************************************************
10    On Error GoTo errorh

          ''------------------------------------------------------
          ''recibe el X,Y donde EL REY ANTERIOR ESTABA POSICIONADO.
          ''------------------------------------------------------
          ''35,25 y 67,25 son las posiciones del rey
          
          ''Sub CrearNPC(NroNPC As Integer, mapa As Integer, OrigPos As WorldPos)
          ''Public Const PRCLER_NPC = 900
          ''Public Const PRGUER_NPC = 901
          ''Public Const PRMAGO_NPC = 902
          ''Public Const PRCAZA_NPC = 903
          ''Public Const PRKING_NPC = 904
          Dim wp As WorldPos
          Dim wp2 As WorldPos
          Dim TeleFrag As Integer
          
20        wp.map = MAPA_PRETORIANO
30        If X < 50 Then ''forma burda de ver que alcoba es
40            wp.X = ALCOBA2_X
50            wp.Y = ALCOBA2_Y
60        Else
70            wp.X = ALCOBA1_X
80            wp.Y = ALCOBA1_Y
90        End If
100       pretorianosVivos = 7 'Hay 7 + el Rey.
110       TeleFrag = MapData(wp.map, wp.X, wp.Y).NpcIndex
          
120       If TeleFrag > 0 Then
              ''El rey va a pisar a un npc de antiguo rey
              ''Obtengo en WP2 la mejor posicion cercana
130           Call ClosestLegalPos(wp, wp2)
140           If (LegalPos(wp2.map, wp2.X, wp2.Y)) Then
                  ''mover al actual
                  
150               Call SendData(SendTarget.ToNPCArea, TeleFrag, PrepareMessageCharacterMove(Npclist(TeleFrag).Char.CharIndex, wp2.X, wp2.Y))
                  'Update map and user pos
160               MapData(wp.map, wp.X, wp.Y).NpcIndex = 0
170               Npclist(TeleFrag).Pos = wp2
180               MapData(wp2.map, wp2.X, wp2.Y).NpcIndex = TeleFrag
190           Else
                  ''TELEFRAG!!!
200               Call QuitarNPC(TeleFrag)
210           End If
220       End If
          ''ya limpié el lugar para el rey (wp)
          ''Los otros no necesitan este caso ya que respawnan lejos
          Dim nPos As WorldPos
          'Busco la posicion legal mas cercana aca, aun que creo que tendría que ir en el crearnpc. (NicoNZ)
230       Call ClosestLegalPos(wp, nPos, False, True)
240       Call CrearNPC(PRKING_NPC, MAPA_PRETORIANO, nPos)
          
250       wp.X = wp.X + 3
260       Call ClosestLegalPos(wp, nPos, False, True)
270       Call CrearNPC(PRCLER_NPC, MAPA_PRETORIANO, nPos)
          
280       wp.X = wp.X - 6
290       Call ClosestLegalPos(wp, nPos, False, True)
300       Call CrearNPC(PRCLER_NPC, MAPA_PRETORIANO, nPos)
          
310       wp.Y = wp.Y + 3
320       Call ClosestLegalPos(wp, nPos, False, True)
330       Call CrearNPC(PRGUER_NPC, MAPA_PRETORIANO, nPos)
          
340       wp.X = wp.X + 3
350       Call ClosestLegalPos(wp, nPos, False, True)
360       Call CrearNPC(PRGUER_NPC, MAPA_PRETORIANO, nPos)
          
370       wp.X = wp.X + 3
380       Call ClosestLegalPos(wp, nPos, False, True)
390       Call CrearNPC(PRGUER_NPC, MAPA_PRETORIANO, nPos)
          
400       wp.Y = wp.Y - 6
410       wp.X = wp.X - 1
420       Call ClosestLegalPos(wp, nPos, False, True)
430       Call CrearNPC(PRCAZA_NPC, MAPA_PRETORIANO, nPos)
          
440       wp.X = wp.X - 4
450       Call ClosestLegalPos(wp, nPos, False, True)
460       Call CrearNPC(PRMAGO_NPC, MAPA_PRETORIANO, nPos)
          
470   Exit Sub

errorh:
480       LogError ("Error en NPCAI.CrearClanPretoriano ")
          'do nothing

End Sub

Sub PRCAZA_AI(ByVal npcind As Integer)
10    On Error GoTo errorh
          '' NO CAMBIAR:
          '' HECHIZOS: 1- FLECHA
          

          Dim X As Integer
          Dim Y As Integer
          Dim NPCPosX As Integer
          Dim NPCPosY As Integer
          Dim NPCPosM As Integer
          Dim BestTarget As Integer
          Dim NPCAlInd As Integer
          Dim PJEnInd  As Integer
          
          Dim PJBestTarget As Boolean
          Dim BTx As Integer
          Dim BTy As Integer
          Dim Xc As Integer
          Dim Yc As Integer
          Dim azar As Integer
          Dim azar2 As Integer
          
          Dim quehacer As Byte
              ''1- Ataca usuarios
          
20        NPCPosX = Npclist(npcind).Pos.X
30        NPCPosY = Npclist(npcind).Pos.Y
40        NPCPosM = Npclist(npcind).Pos.map
          
50        PJBestTarget = False
60        X = 0
70        Y = 0
80        quehacer = 0
          
          
90        azar = Sgn(RandomNumber(-1, 1))
          'azar = Sgn(azar)
100       If azar = 0 Then azar = 1
110       azar2 = Sgn(RandomNumber(-1, 1))
          'azar2 = Sgn(azar2)
120       If azar2 = 0 Then azar2 = 1
          
          
          'pick the best target according to the following criteria:
          '1) magues ARE dangerous, but they are weak too, they're
          '   our primary target
          '2) in any other case, our nearest enemy will be attacked
          
130       For X = NPCPosX + (azar * 8) To NPCPosX + (azar * -8) Step -azar
140           For Y = NPCPosY + (azar2 * 7) To NPCPosY + (azar2 * -7) Step -azar2
150               NPCAlInd = MapData(NPCPosM, X, Y).NpcIndex  ''por si implementamos algo contra NPCs
160               PJEnInd = MapData(NPCPosM, X, Y).Userindex
170               If (PJEnInd > 0) And (Npclist(npcind).CanAttack = 1) Then
180                   If (UserList(PJEnInd).flags.invisible = 0 Or UserList(PJEnInd).flags.Oculto = 0) And Not (UserList(PJEnInd).flags.Muerto = 1) And Not UserList(PJEnInd).flags.AdminInvisible = 1 And UserList(PJEnInd).flags.AdminPerseguible Then
                      'ToDo: Borrar los GMs
190                       If (EsMagoOClerigo(PJEnInd)) Then
                              ''say no more, atacar a este
200                           PJBestTarget = True
210                           BestTarget = PJEnInd
220                           quehacer = 1
                              'Call NpcLanzaSpellSobreUser(npcind, PJEnInd, Npclist(npcind).Spells(1)) ''flecha pasa como spell
230                           X = NPCPosX + (azar * -8)
240                           Y = NPCPosY + (azar2 * -7)
                              ''forma espantosa de zafar del for
250                        Else
260                           If (BestTarget > 0) Then
                                  ''ver el mas cercano a mi
270                               If Sqr((X - NPCPosX) ^ 2 + (Y - NPCPosY) ^ 2) < Sqr((NPCPosX - UserList(BestTarget).Pos.X) ^ 2 + (NPCPosY - UserList(BestTarget).Pos.Y) ^ 2) Then
                                      ''el nuevo esta mas cerca
280                                   PJBestTarget = True
290                                   BestTarget = PJEnInd
300                                   quehacer = 1
310                               End If
320                           Else
330                               PJBestTarget = True
340                               BestTarget = PJEnInd
350                               quehacer = 1
360                           End If
370                       End If
380                   End If
390               End If  ''Fin analisis del tile
400           Next Y
410       Next X
          
420   Select Case quehacer
          Case 1  ''nearest target
430           If (Npclist(npcind).CanAttack = 1) Then
440               Call NpcLanzaSpellSobreUser(npcind, BestTarget, Npclist(npcind).Spells(1))
450           End If
          ''case 2: not yet implemented
460   End Select
          
      ''  Vamos a setear el hold on del cazador en el medio entre el rey
      ''  y el atacante. De esta manera se lo podra atacar aun asi está lejos
      ''  pero sin alejarse del rango de los an hoax vorps de los
      ''  clerigos o rey. A menos q este paralizado, claro

470   If Npclist(npcind).flags.Paralizado = 1 Then Exit Sub

480   If Not NPCPosM = MAPA_PRETORIANO Then Exit Sub


      'MEJORA: Si quedan solos, se van con el resto del ejercito
490   If Npclist(npcind).Invent.ArmourEqpSlot <> 0 Then
          'si me estoy yendo a alguna alcoba
500       Call CambiarAlcoba(npcind)
510       Exit Sub
520   End If




530   If EstoyMuyLejos(npcind) Then
540       VolverAlCentro (npcind)
550       Exit Sub
560   End If

570   If (BestTarget > 0) Then

580       BTx = UserList(BestTarget).Pos.X
590       BTy = UserList(BestTarget).Pos.Y
          
600       If NPCPosX < 50 Then
              
610           Call GreedyWalkTo(npcind, MAPA_PRETORIANO, ALCOBA1_X + ((BTx - ALCOBA1_X) \ 2), ALCOBA1_Y + ((BTy - ALCOBA1_Y) \ 2))
              'GreedyWalkTo npcind, MAPA_PRETORIANO, ALCOBA1_X + ((BTx - ALCOBA1_X) \ 2), ALCOBA1_Y + ((BTy - ALCOBA1_Y) \ 2)
620       Else
630           Call GreedyWalkTo(npcind, MAPA_PRETORIANO, ALCOBA2_X + ((BTx - ALCOBA2_X) \ 2), ALCOBA2_Y + ((BTy - ALCOBA2_Y) \ 2))
              'GreedyWalkTo npcind, MAPA_PRETORIANO, ALCOBA2_X + ((BTx - ALCOBA2_X) \ 2), ALCOBA2_Y + ((BTy - ALCOBA2_Y) \ 2)
640       End If
650   Else
          ''2do Loop. Busca gente acercandose por otros frentes para frenarla
660       If NPCPosX < 50 Then Xc = ALCOBA1_X Else Xc = ALCOBA2_X
670       Yc = ALCOBA1_Y
          
680       For X = Xc - 16 To Xc + 16
690           For Y = Yc - 14 To Yc + 14
700               If Not (X <= NPCPosX + 8 And X >= NPCPosX - 8 And Y >= NPCPosY - 7 And Y <= NPCPosY + 7) Then
                      ''si es un tile no analizado
710                   PJEnInd = MapData(NPCPosM, X, Y).Userindex    ''por si implementamos algo contra NPCs
720                   If (PJEnInd > 0) Then
730                       If Not (UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1 Or UserList(PJEnInd).flags.Muerto = 1) Then
                              ''si no esta muerto.., ya encontro algo para ir a buscar
740                           Call GreedyWalkTo(npcind, MAPA_PRETORIANO, UserList(PJEnInd).Pos.X, UserList(PJEnInd).Pos.Y)
750                           Exit Sub
760                       End If
770                   End If
780               End If
790           Next Y
800       Next X
          
          ''vuelve si no esta en proceso de ataque a usuarios
810       If (Npclist(npcind).CanAttack = 1) Then Call VolverAlCentro(npcind)

820   End If
          
830   Exit Sub
errorh:
840       LogError ("Error en NPCAI.PRCAZA_AI ")
          'do nothing

End Sub

Sub PRMAGO_AI(ByVal npcind As Integer)

10        On Error GoTo PRMAGO_AI_Error

      'HECHIZOS: NO CAMBIAR ACA
      'REPRESENTAN LA UBICACION DE LOS SPELLS EN NPC_HOSTILES.DAT y si se los puede cambiar en ese archivo
      '1- APOCALIPSIS 'modificable
      '2- REMOVER INVISIBILIDAD 'NO MODIFICABLE
          Dim DAT_APOCALIPSIS As Integer
          Dim DAT_REMUEVE_INVI As Integer
20        DAT_APOCALIPSIS = 1
30        DAT_REMUEVE_INVI = 2

          ''EL mago pretoriano guarda  el index al NPC Rey en el
          ''inventario.barcoobjind parameter. Ese no es usado nunca.
          ''EL objetivo es no modificar al TAD NPC utilizando una propiedad
          ''que nunca va a ser utilizada por un NPC (espero)
          Dim X      As Integer
          Dim Y      As Integer
          Dim NPCPosX As Integer
          Dim NPCPosY As Integer
          Dim NPCPosM As Integer
          Dim BestTarget As Integer
          Dim NPCAlInd As Integer
          Dim PJEnInd As Integer
          Dim PJBestTarget As Boolean
          Dim bs     As Byte
          Dim azar   As Integer
          Dim azar2  As Integer

          Dim quehacer As Byte
          ''1- atacar a enemigos
          ''2- remover invisibilidades
          ''3- rotura de vara

40        NPCPosX = Npclist(npcind).Pos.X   ''store current position
50        NPCPosY = Npclist(npcind).Pos.Y   ''for direct access
60        NPCPosM = Npclist(npcind).Pos.map

70        PJBestTarget = False
80        BestTarget = 0
90        quehacer = 0
100       X = 0
110       Y = 0


120       If (Npclist(npcind).Stats.MinHp < 750) Then   ''Dying
130           quehacer = 3        ''va a romper su vara en 5 segundos
140       Else
150           If Not (Npclist(npcind).Invent.BarcoSlot = 6) Then
160               Npclist(npcind).Invent.BarcoSlot = 6    ''restore wand break counter
170               Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageCreateFX(Npclist(npcind).Char.CharIndex, 0, 0))
180           End If

              'pick the best target according to the following criteria:
              '1) invisible enemies can be detected sometimes
              '2) a wizard's mission is background spellcasting attack

190           azar = Sgn(RandomNumber(-1, 1))
              'azar = Sgn(azar)
200           If azar = 0 Then azar = 1
210           azar2 = Sgn(RandomNumber(-1, 1))
              'azar2 = Sgn(azar2)
220           If azar2 = 0 Then azar2 = 1

              ''esto fue para rastrear el combat field al azar
              ''Si no se hace asi, los NPCs Pretorianos "combinan" ataques, y cada
              ''ataque puede sumar hasta 700 Hit Points, lo cual los vuelve
              ''invulnerables

              '        azar = 1

230           For X = NPCPosX + (azar * 8) To NPCPosX + (azar * -8) Step -azar
240               For Y = NPCPosY + (azar2 * 7) To NPCPosY + (azar2 * -7) Step -azar2
250                   NPCAlInd = MapData(NPCPosM, X, Y).NpcIndex  ''por si implementamos algo contra NPCs
260                   PJEnInd = MapData(NPCPosM, X, Y).Userindex

270                   If (PJEnInd > 0) And (Npclist(npcind).CanAttack = 1) Then
280                       If Not (UserList(PJEnInd).flags.Muerto = 1) And Not (UserList(PJEnInd).flags.AdminInvisible = 1) And UserList(PJEnInd).flags.AdminPerseguible Then
290                           If (UserList(PJEnInd).flags.invisible = 1) Or (UserList(PJEnInd).flags.Oculto = 1) Then
                                  ''usuario invisible, vamos a ver si se la podemos sacar

300                               If (RandomNumber(1, 100) <= 35) Then
                                      ''mago detecta invisiblidad
310                                   Npclist(npcind).CanAttack = 0
320                                   Call NPCRemueveInvisibilidad(npcind, PJEnInd, DAT_REMUEVE_INVI)
330                                   Exit Sub    ''basta, SUFICIENTE!, jeje
340                               End If
350                               If UserList(PJEnInd).flags.Paralizado = 1 Then
                                      ''los usuarios invisibles y paralizados son un buen target!
360                                   BestTarget = PJEnInd
370                                   PJBestTarget = True
380                                   quehacer = 2
390                               End If
400                           ElseIf (UserList(PJEnInd).flags.Paralizado = 1) Then
410                               If (BestTarget > 0) Then
420                                   If Not (UserList(BestTarget).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1) Then
                                          ''encontre un paralizado visible, y no hay un besttarget invisible (paralizado invisible)
430                                       BestTarget = PJEnInd
440                                       PJBestTarget = True
450                                       quehacer = 2
460                                   End If
470                               Else
480                                   BestTarget = PJEnInd
490                                   PJBestTarget = True
500                                   quehacer = 2
510                               End If
520                           ElseIf BestTarget = 0 Then
                                  ''movil visible
530                               BestTarget = PJEnInd
540                               PJBestTarget = True
550                               quehacer = 2
560                           End If  ''
570                       End If  ''endif:    not muerto
580                   End If  ''endif: es un tile con PJ y puede atacar
590               Next Y
600           Next X
610       End If  ''endif esta muriendo


620       Select Case quehacer
              ''case 1 esta "harcodeado" en el doble for
              ''es remover invisibilidades
          Case 2          ''apocalipsis Rahma Nañarak O'al
630           Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead(Hechizos(Npclist(npcind).Spells(DAT_APOCALIPSIS)).PalabrasMagicas, Npclist(npcind).Char.CharIndex, vbCyan))
640           Call NpcLanzaSpellSobreUser2(npcind, BestTarget, Npclist(npcind).Spells(DAT_APOCALIPSIS))    ''SPELL 1 de Mago: Apocalipsis
650       Case 3

660           Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageCreateFX(Npclist(npcind).Char.CharIndex, FXIDs.FXMEDITARGRANDE, INFINITE_LOOPS))
              ''UserList(UserIndex).Char.FX = FXIDs.FXMEDITARGRANDE

670           If Npclist(npcind).CanAttack = 1 Then
680               Npclist(npcind).CanAttack = 0
690               bs = Npclist(npcind).Invent.BarcoSlot
700               bs = Npclist(npcind).Invent.MonturaSlot
710               bs = bs - 1
720               Call MagoDestruyeWand(npcind, bs, DAT_APOCALIPSIS)
730               If bs = 0 Then
740                   Call MuereNpc(npcind, 0)
750               Else
760                   Npclist(npcind).Invent.BarcoSlot = bs
770                   Npclist(npcind).Invent.MonturaSlot = bs
780               End If
790           End If
800       End Select


          ''movimiento (si puede)
          ''El mago no se mueve a menos q tenga alguien al lado

810       If Npclist(npcind).flags.Paralizado = 1 Then Exit Sub

820       If Not (quehacer = 3) Then      ''si no ta matandose
              ''alejarse si tiene un PJ cerca
              ''pero alejarse sin alejarse del rey
830           If Not (NPCPosM = MAPA_PRETORIANO) Then Exit Sub

              ''Si no hay nadie cerca, o no tengo nada que hacer...
840           If (BestTarget = 0) And (Npclist(npcind).CanAttack = 1) Then Call VolverAlCentro(npcind)

850           PJEnInd = MapData(NPCPosM, NPCPosX - 1, NPCPosY).Userindex
860           If (PJEnInd > 0) Then
870               If Not (UserList(PJEnInd).flags.Muerto = 1) And Not (UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1) Then
                      ''esta es una forma muy facil de matar 2 pajaros
                      ''de un tiro. Se aleja del usuario pq el centro va a
                      ''estar ocupado, y a la vez se aproxima al rey, manteniendo
                      ''una linea de defensa compacta
880                   Call VolverAlCentro(npcind)
890                   Exit Sub
900               End If
910           End If

920           PJEnInd = MapData(NPCPosM, NPCPosX + 1, NPCPosY).Userindex
930           If PJEnInd > 0 Then
940               If Not (UserList(PJEnInd).flags.Muerto = 1) And Not (UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1) Then
950                   Call VolverAlCentro(npcind)
960                   Exit Sub
970               End If
980           End If

990           PJEnInd = MapData(NPCPosM, NPCPosX, NPCPosY - 1).Userindex
1000          If PJEnInd > 0 Then
1010              If Not (UserList(PJEnInd).flags.Muerto = 1) And Not (UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1) Then
1020                  Call VolverAlCentro(npcind)
1030                  Exit Sub
1040              End If
1050          End If

1060          PJEnInd = MapData(NPCPosM, NPCPosX, NPCPosY + 1).Userindex
1070          If PJEnInd > 0 Then
1080              If Not (UserList(PJEnInd).flags.Muerto = 1) And Not (UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1) Then
1090                  Call VolverAlCentro(npcind)
1100                  Exit Sub
1110              End If
1120          End If


1130      End If  ''end if not matandose

1140      Exit Sub

PRMAGO_AI_Error:

1150      LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure PRMAGO_AI, line " & Erl & "."
    
End Sub


Sub PRREY_AI(ByVal npcind As Integer)
10    On Error GoTo errorh
          'HECHIZOS: NO CAMBIAR ACA
          'REPRESENTAN LA UBICACION DE LOS SPELLS EN NPC_HOSTILES.DAT y si se los puede cambiar en ese archivo
          '1- CURAR_LEVES 'NO MODIFICABLE
          '2- REMOVER PARALISIS 'NO MODIFICABLE
          '3- CEUGERA - 'NO MODIFICABLE
          '4- ESTUPIDEZ - 'NO MODIFICABLE
          '5- CURARVENENO - 'NO MODIFICABLE
          Dim DAT_CURARLEVES As Integer
          Dim DAT_REMUEVEPARALISIS As Integer
          Dim DAT_CEGUERA As Integer
          Dim DAT_ESTUPIDEZ As Integer
          Dim DAT_CURARVENENO As Integer
20        DAT_CURARLEVES = 1
30        DAT_REMUEVEPARALISIS = 2
40        DAT_CEGUERA = 3
50        DAT_ESTUPIDEZ = 4
60        DAT_CURARVENENO = 5
          
          
          Dim UI As Integer
          Dim X As Integer
          Dim Y As Integer
          Dim NPCPosX As Integer
          Dim NPCPosY As Integer
          Dim NPCPosM As Integer
          Dim NPCAlInd As Integer
          Dim PJEnInd As Integer
          Dim BestTarget As Integer
          Dim distBestTarget As Integer
          Dim dist As Integer
          Dim e_p As Integer
          Dim hayPretorianos As Boolean
          Dim headingloop As Byte
          Dim nPos As WorldPos
          ''Dim quehacer As Integer
              ''1- remueve paralisis con un minimo % de efecto
              ''2- remueve veneno
              ''3- cura
          
70        NPCPosM = Npclist(npcind).Pos.map
80        NPCPosX = Npclist(npcind).Pos.X
90        NPCPosY = Npclist(npcind).Pos.Y
100       BestTarget = 0
110       distBestTarget = 0
120       hayPretorianos = False
          
          'pick the best target according to the following criteria:
          'King won't fight. Since praetorians' mission is to keep him alive
          'he will stay as far as possible from combat environment, but close enought
          'as to aid his loyal army.
          'If his army has been annihilated, the king will pick the
          'closest enemy an chase it using his special 'weapon speedhack' ability
130       For X = NPCPosX - 8 To NPCPosX + 8
140           For Y = NPCPosY - 7 To NPCPosY + 7
                  'scan combat field
150               NPCAlInd = MapData(NPCPosM, X, Y).NpcIndex
160               PJEnInd = MapData(NPCPosM, X, Y).Userindex
170               If (Npclist(npcind).CanAttack = 1) Then   ''saltea el analisis si no puede atacar para evitar cuentas
180                   If (NPCAlInd > 0) Then
190                       e_p = esPretoriano(NPCAlInd)
200                       If e_p > 0 And e_p < 6 And (Not (NPCAlInd = npcind)) Then
210                           hayPretorianos = True
                              
                              'Me curo mientras haya pretorianos (no es lo ideal, debería no dar experiencia tampoco, pero por ahora es lo que hay)
220                           Npclist(npcind).Stats.MinHp = Npclist(npcind).Stats.MaxHp
230                       End If
                          
240                       If (Npclist(NPCAlInd).flags.Paralizado = 1 And e_p > 0 And e_p < 6) Then
                              ''el rey puede desparalizar con una efectividad del 20%
250                           If (RandomNumber(1, 100) < 21) Then
260                               Call NPCRemueveParalisisNPC(npcind, NPCAlInd, DAT_REMUEVEPARALISIS)
270                               Npclist(npcind).CanAttack = 0
280                               Exit Sub
290                           End If
                          
                          ''failed to remove
300                       ElseIf (Npclist(NPCAlInd).flags.Envenenado = 1) Then    ''un chiche :D
310                           If esPretoriano(NPCAlInd) Then
320                               Call NPCRemueveVenenoNPC(npcind, NPCAlInd, DAT_CURARVENENO)
330                               Npclist(npcind).CanAttack = 0
340                               Exit Sub
350                           End If
360                       ElseIf (Npclist(NPCAlInd).Stats.MaxHp > Npclist(NPCAlInd).Stats.MinHp) Then
370                           If esPretoriano(NPCAlInd) And Not (NPCAlInd = npcind) Then
                                  ''cura, salvo q sea yo mismo. Eso lo hace 'despues'
380                               Call NPCCuraLevesNPC(npcind, NPCAlInd, DAT_CURARLEVES)
390                               Npclist(npcind).CanAttack = 0
                                  ''Exit Sub
400                           End If
410                       End If
420                   End If

430                   If PJEnInd > 0 And Not hayPretorianos Then
440                       If Not (UserList(PJEnInd).flags.Muerto = 1 Or UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1 Or UserList(PJEnInd).flags.Ceguera = 1) And UserList(PJEnInd).flags.AdminPerseguible Then
                              ''si no esta muerto o invisible o ciego... o tiene el /ignorando
450                           dist = Sqr((UserList(PJEnInd).Pos.X - NPCPosX) ^ 2 + (UserList(PJEnInd).Pos.Y - NPCPosY) ^ 2)
460                           If (dist < distBestTarget Or BestTarget = 0) Then
470                               BestTarget = PJEnInd
480                               distBestTarget = dist
490                           End If
500                       End If
510                   End If
520               End If  ''canattack = 1
530           Next Y
540       Next X
          
550       If Not hayPretorianos Then
              ''si estoy aca es porque no hay pretorianos cerca!!!
              ''Todo mi ejercito fue asesinado
              ''Salgo a atacar a todos a lo loco a espadazos
560           If BestTarget > 0 Then
570               If EsAlcanzable(npcind, BestTarget) Then
580                   Call GreedyWalkTo(npcind, UserList(BestTarget).Pos.map, UserList(BestTarget).Pos.X, UserList(BestTarget).Pos.Y)
                      'GreedyWalkTo npcind, UserList(BestTarget).Pos.Map, UserList(BestTarget).Pos.X, UserList(BestTarget).Pos.Y
590               Else
                      ''el chabon es piola y ataca desde lejos entonces lo castigamos!
600                   Call NPCLanzaEstupidezPJ(npcind, BestTarget, DAT_ESTUPIDEZ)
610                   Call NPCLanzaCegueraPJ(npcind, BestTarget, DAT_CEGUERA)
620               End If
                  
                  ''heading loop de ataque
                  ''teclavolaespada
630               For headingloop = eHeading.NORTH To eHeading.WEST
640                   nPos = Npclist(npcind).Pos
650                   Call HeadtoPos(headingloop, nPos)
660                   If InMapBounds(nPos.map, nPos.X, nPos.Y) Then
670                       UI = MapData(nPos.map, nPos.X, nPos.Y).Userindex
680                       If UI > 0 Then
690                           If NpcAtacaUser(npcind, UI) Then
700                               Call ChangeNPCChar(npcind, Npclist(npcind).Char.body, Npclist(npcind).Char.Head, headingloop)
710                           End If
                              
                              ''special speed ability for praetorian king ---------
720                           Npclist(npcind).CanAttack = 1   ''this is NOT a bug!!
                              '----------------------------------------------------
                          
730                       End If
740                   End If
750               Next headingloop
              
760           Else    ''no hay targets cerca
770               Call VolverAlCentro(npcind)
780               If (Npclist(npcind).Stats.MinHp < Npclist(npcind).Stats.MaxHp) And (Npclist(npcind).CanAttack = 1) Then
                      ''si no hay ndie y estoy daniado me curo
790                   Call NPCCuraLevesNPC(npcind, npcind, DAT_CURARLEVES)
800                   Npclist(npcind).CanAttack = 0
810               End If
              
820           End If
830       End If
840   Exit Sub

errorh:
850       LogError ("Error en NPCAI.PRREY_AI? ")
          
End Sub

Sub PRGUER_AI(ByVal npcind As Integer)
10    On Error GoTo errorh

          Dim headingloop As Byte
          Dim nPos As WorldPos
          Dim X As Integer
          Dim Y As Integer
          Dim dist As Integer
          Dim distBestTarget As Integer
          Dim NPCPosX As Integer
          Dim NPCPosY As Integer
          Dim NPCPosM As Integer
          Dim NPCAlInd As Integer
          Dim UI As Integer
          Dim PJEnInd As Integer
          Dim BestTarget As Integer
20        NPCPosM = Npclist(npcind).Pos.map
30        NPCPosX = Npclist(npcind).Pos.X
40        NPCPosY = Npclist(npcind).Pos.Y
50        BestTarget = 0
60        dist = 0
70        distBestTarget = 0
          
80        For X = NPCPosX - 8 To NPCPosX + 8
90            For Y = NPCPosY - 7 To NPCPosY + 7
100               PJEnInd = MapData(NPCPosM, X, Y).Userindex
110               If (PJEnInd > 0) Then
120                   If (Not (UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1 Or UserList(PJEnInd).flags.Muerto = 1)) And EsAlcanzable(npcind, PJEnInd) And UserList(PJEnInd).flags.AdminPerseguible Then
                          ''caluclo la distancia al PJ, si esta mas cerca q el actual
                          ''mejor besttarget entonces ataco a ese.
130                       If (BestTarget > 0) Then
140                           dist = Sqr((UserList(PJEnInd).Pos.X - NPCPosX) ^ 2 + (UserList(PJEnInd).Pos.Y - NPCPosY) ^ 2)
150                           If (dist < distBestTarget) Then
160                               BestTarget = PJEnInd
170                               distBestTarget = dist
180                           End If
190                       Else
200                           distBestTarget = Sqr((UserList(PJEnInd).Pos.X - NPCPosX) ^ 2 + (UserList(PJEnInd).Pos.Y - NPCPosY) ^ 2)
210                           BestTarget = PJEnInd
220                       End If
230                   End If
240               End If
250           Next Y
260       Next X
          
          ''LLamo a esta funcion si lo llevaron muy lejos.
          ''La idea es que no lo "alejen" del rey y despues queden
          ''lejos de la batalla cuando matan a un enemigo o este
          ''sale del area de combate (tipica forma de separar un clan)
270       If Npclist(npcind).flags.Paralizado = 0 Then

              'MEJORA: Si quedan solos, se van con el resto del ejercito
280           If Npclist(npcind).Invent.ArmourEqpSlot <> 0 Then
290               Call CambiarAlcoba(npcind)
                  'si me estoy yendo a alguna alcoba
300           ElseIf BestTarget = 0 Or EstoyMuyLejos(npcind) Then
310               Call VolverAlCentro(npcind)
320           ElseIf BestTarget > 0 Then
330               Call GreedyWalkTo(npcind, UserList(BestTarget).Pos.map, UserList(BestTarget).Pos.X, UserList(BestTarget).Pos.Y)
340           End If
350       End If

      ''teclavolaespada
360   For headingloop = eHeading.NORTH To eHeading.WEST
370       nPos = Npclist(npcind).Pos
380       Call HeadtoPos(headingloop, nPos)
390       If InMapBounds(nPos.map, nPos.X, nPos.Y) Then
400           UI = MapData(nPos.map, nPos.X, nPos.Y).Userindex
410           If UI > 0 Then
420               If Not (UserList(UI).flags.Muerto = 1) Then
430                   If NpcAtacaUser(npcind, UI) Then
440                       Call ChangeNPCChar(npcind, Npclist(npcind).Char.body, Npclist(npcind).Char.Head, headingloop)
450                   End If
460                   Npclist(npcind).CanAttack = 0
470               End If
480           End If
490       End If
500   Next headingloop


510   Exit Sub

errorh:
520       LogError ("Error en NPCAI.PRGUER_AI? ")
          

End Sub

Sub PRCLER_AI(ByVal npcind As Integer)
10    On Error GoTo errorh
          
          'HECHIZOS: NO CAMBIAR ACA
          'REPRESENTAN LA UBICACION DE LOS SPELLS EN NPC_HOSTILES.DAT y si se los puede cambiar en ese archivo
          '1- PARALIZAR PJS 'MODIFICABLE
          '2- REMOVER PARALISIS 'NO MODIFICABLE
          '3- CURARGRAVES - 'NO MODIFICABLE
          '4- PARALIZAR MASCOTAS - 'NO MODIFICABLE
          '5- CURARVENENO - 'NO MODIFICABLE
          Dim DAT_PARALIZARPJ As Integer
          Dim DAT_REMUEVEPARALISIS As Integer
          Dim DAT_CURARGRAVES As Integer
          Dim DAT_PARALIZAR_NPC As Integer
          Dim DAT_TORMENTAAVANZADA As Integer
20        DAT_PARALIZARPJ = 1
30        DAT_REMUEVEPARALISIS = 2
40        DAT_PARALIZAR_NPC = 3
50        DAT_CURARGRAVES = 4
60        DAT_TORMENTAAVANZADA = 5

          Dim X As Integer
          Dim Y As Integer
          Dim NPCPosX As Integer
          Dim NPCPosY As Integer
          Dim NPCPosM As Integer
          Dim NPCAlInd As Integer
          Dim PJEnInd As Integer
          Dim centroX As Integer
          Dim centroY As Integer
          Dim BestTarget As Integer
          Dim PJBestTarget As Boolean
          Dim azar, azar2 As Integer
          Dim quehacer As Byte
              ''1- paralizar enemigo,
              ''2- bombardear enemigo
              ''3- ataque a mascotas
              ''4- curar aliado
70        quehacer = 0
80        NPCPosM = Npclist(npcind).Pos.map
90        NPCPosX = Npclist(npcind).Pos.X
100       NPCPosY = Npclist(npcind).Pos.Y
110       PJBestTarget = False
120       BestTarget = 0
          
130       azar = Sgn(RandomNumber(-1, 1))
140       If azar = 0 Then azar = 1
150       azar2 = Sgn(RandomNumber(-1, 1))
160       If azar2 = 0 Then azar2 = 1
          
          'pick the best target according to the following criteria:
          '1) "hoaxed" friends MUST be released
          '2) enemy shall be annihilated no matter what
          '3) party healing if no threats
170       For X = NPCPosX + (azar * 8) To NPCPosX + (azar * -8) Step -azar
180           For Y = NPCPosY + (azar2 * 7) To NPCPosY + (azar2 * -7) Step -azar2
                  'scan combat field
190               NPCAlInd = MapData(NPCPosM, X, Y).NpcIndex
200               PJEnInd = MapData(NPCPosM, X, Y).Userindex
210               If (Npclist(npcind).CanAttack = 1) Then   ''saltea el analisis si no puede atacar para evitar cuentas
220                   If (NPCAlInd > 0) Then  ''allie?
230                       If (esPretoriano(NPCAlInd) = 0) Then
240                           If (Npclist(NPCAlInd).MaestroUser > 0) And (Not (Npclist(NPCAlInd).flags.Paralizado > 0)) Then
250                               Call NPCparalizaNPC(npcind, NPCAlInd, DAT_PARALIZAR_NPC)
260                               Npclist(npcind).CanAttack = 0
270                               Exit Sub
280                           End If
290                       Else    'es un PJ aliado en combate
300                           If (Npclist(NPCAlInd).flags.Paralizado = 1) Then
                                  ' amigo paralizado, an hoax vorp YA
310                               Call NPCRemueveParalisisNPC(npcind, NPCAlInd, DAT_REMUEVEPARALISIS)
320                               Npclist(npcind).CanAttack = 0
330                               Exit Sub
340                           ElseIf (BestTarget = 0) Then ''si no tiene nada q hacer..
350                               If (Npclist(NPCAlInd).Stats.MaxHp > Npclist(NPCAlInd).Stats.MinHp) Then
360                                   BestTarget = NPCAlInd   ''cura heridas
370                                   PJBestTarget = False
380                                   quehacer = 4
390                               End If
400                           End If
410                       End If
420                   ElseIf (PJEnInd > 0) Then ''aggressor
430                       If Not (UserList(PJEnInd).flags.Muerto = 1) And UserList(PJEnInd).flags.AdminPerseguible Then
440                           If (UserList(PJEnInd).flags.Paralizado = 0) Then
450                               If (Not (UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1)) Then
                                      ''PJ movil y visible, jeje, si o si es target
460                                   BestTarget = PJEnInd
470                                   PJBestTarget = True
480                                   quehacer = 1
490                               End If
500                           Else    ''PJ paralizado, ataca este invisible o no
510                               If Not (BestTarget > 0) Or Not (PJBestTarget) Then ''a menos q tenga algo mejor
520                                   BestTarget = PJEnInd
530                                   PJBestTarget = True
540                                   quehacer = 2
550                               End If
560                           End If  ''endif paralizado
570                       End If  ''end if not muerto
580                   End If  ''listo el analisis del tile
590               End If  ''saltea el analisis si no puede atacar, en realidad no es lo "mejor" pero evita cuentas inútiles
600           Next Y
610       Next X
                  
          ''aqui (si llego) tiene el mejor target
620       Select Case quehacer
          Case 0
              ''nada que hacer. Buscar mas alla del campo de visión algun aliado, a menos
              ''que este paralizado pq no puedo ir
630           If Npclist(npcind).flags.Paralizado = 1 Then Exit Sub
              
640           If Not NPCPosM = MAPA_PRETORIANO Then Exit Sub
              
650           If NPCPosX < 50 Then centroX = ALCOBA1_X Else centroX = ALCOBA2_X
660           centroY = ALCOBA1_Y
              ''aca establecí el lugar de las alcobas
              
              ''Este doble for busca amigos paralizados lejos para ir a rescatarlos
              ''Entra aca solo si en el area cercana al rey no hay algo mejor que
              ''hacer.
670           For X = centroX - 16 To centroX + 16
680               For Y = centroY - 15 To centroY + 15
690                   If Not (X < NPCPosX + 8 And X > NPCPosX + 8 And Y < NPCPosY + 7 And Y > NPCPosY - 7) Then
                      ''si no es un tile ya analizado... (evito cuentas)
700                       NPCAlInd = MapData(NPCPosM, X, Y).NpcIndex
710                       If NPCAlInd > 0 Then
720                           If (esPretoriano(NPCAlInd) > 0 And Npclist(NPCAlInd).flags.Paralizado = 1) Then
                                  ''si esta paralizado lo va a rescatar, sino
                                  ''ya va a volver por su cuenta
730                               Call GreedyWalkTo(npcind, NPCPosM, Npclist(NPCAlInd).Pos.X, Npclist(NPCAlInd).Pos.Y)
      '                            GreedyWalkTo npcind, NPCPosM, Npclist(NPCAlInd).Pos.X, Npclist(NPCAlInd).Pos.Y
740                               Exit Sub
750                           End If
760                       End If  ''endif npc
770                   End If  ''endif tile analizado
780               Next Y
790           Next X
              
              ''si estoy aca esta totalmente al cuete el clerigo o mal posicionado por rescate anterior
800           If Npclist(npcind).Invent.ArmourEqpSlot = 0 Then
810               Call VolverAlCentro(npcind)
820               Exit Sub
830           End If
              ''fin quehacer = 0 (npc al cuete)
              
840       Case 1  '' paralizar enemigo PJ
850           Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead(Hechizos(Npclist(npcind).Spells(DAT_PARALIZARPJ)).PalabrasMagicas, Npclist(npcind).Char.CharIndex, vbCyan))
860           Call NpcLanzaSpellSobreUser(npcind, BestTarget, Npclist(npcind).Spells(DAT_PARALIZARPJ)) ''SPELL 1 de Clerigo es PARALIZAR
870           Exit Sub
880       Case 2  '' ataque a usuarios (invisibles tambien)
890           Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead(Hechizos(Npclist(npcind).Spells(DAT_TORMENTAAVANZADA)).PalabrasMagicas, Npclist(npcind).Char.CharIndex, vbCyan))
900           Call NpcLanzaSpellSobreUser2(npcind, BestTarget, Npclist(npcind).Spells(DAT_TORMENTAAVANZADA)) ''SPELL 2 de Clerigo es Vax On Tar avanzado
910           Exit Sub
920       Case 3  '' ataque a mascotas
930           If Not (Npclist(BestTarget).flags.Paralizado = 1) Then
940               Call NPCparalizaNPC(npcind, BestTarget, DAT_PARALIZAR_NPC)
950               Npclist(npcind).CanAttack = 0
960           End If  ''TODO: vax on tar sobre mascotas
970       Case 4  '' party healing
980           Call NPCcuraNPC(npcind, BestTarget, DAT_CURARGRAVES)
990           Npclist(npcind).CanAttack = 0
1000      End Select
          
          
          
          ''movimientos
          ''EL clerigo no tiene un movimiento fijo, pero es esperable
          ''que no se aleje mucho del rey... y si se aleje de espaderos
          
1010      If Npclist(npcind).flags.Paralizado = 1 Then Exit Sub
          
1020      If Not (NPCPosM = MAPA_PRETORIANO) Then Exit Sub
          
          'MEJORA: Si quedan solos, se van con el resto del ejercito
1030      If Npclist(npcind).Invent.ArmourEqpSlot <> 0 Then
1040          Call CambiarAlcoba(npcind)
1050          Exit Sub
1060      End If
          
          
1070      PJEnInd = MapData(NPCPosM, NPCPosX - 1, NPCPosY).Userindex
1080      If PJEnInd > 0 Then
1090          If Not (UserList(PJEnInd).flags.Muerto = 1) And Not (UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1) Then
                  ''esta es una forma muy facil de matar 2 pajaros
                  ''de un tiro. Se aleja del usuario pq el centro va a
                  ''estar ocupado, y a la vez se aproxima al rey, manteniendo
                  ''una linea de defensa compacta
1100              Call VolverAlCentro(npcind)
1110              Exit Sub
1120          End If
1130      End If
          
1140      PJEnInd = MapData(NPCPosM, NPCPosX + 1, NPCPosY).Userindex
1150      If PJEnInd > 0 Then
1160          If Not (UserList(PJEnInd).flags.Muerto = 1) And Not (UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1) Then
1170              Call VolverAlCentro(npcind)
1180              Exit Sub
1190          End If
1200      End If
          
1210      PJEnInd = MapData(NPCPosM, NPCPosX, NPCPosY - 1).Userindex
1220      If PJEnInd > 0 Then
1230          If Not (UserList(PJEnInd).flags.Muerto = 1) And Not (UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1) Then
1240              Call VolverAlCentro(npcind)
1250              Exit Sub
1260          End If
1270      End If
          
1280      PJEnInd = MapData(NPCPosM, NPCPosX, NPCPosY + 1).Userindex
1290      If PJEnInd > 0 Then
1300          If Not (UserList(PJEnInd).flags.Muerto = 1) And Not (UserList(PJEnInd).flags.invisible = 1 Or UserList(PJEnInd).flags.Oculto = 1) Then
1310              Call VolverAlCentro(npcind)
1320              Exit Sub
1330          End If
1340      End If
          
1350  Exit Sub

errorh:
1360      LogError ("Error en NPCAI.PRCLER_AI? ")
          
End Sub

Function EsMagoOClerigo(ByVal PJEnInd As Integer) As Boolean
10    On Error GoTo errorh

20        EsMagoOClerigo = UserList(PJEnInd).clase = eClass.Mage Or _
                              UserList(PJEnInd).clase = eClass.Cleric Or _
                              UserList(PJEnInd).clase = eClass.Druid Or _
                              UserList(PJEnInd).clase = eClass.Bard
30    Exit Function

errorh:
40        LogError ("Error en NPCAI.EsMagoOClerigo? ")
End Function

Sub NPCRemueveVenenoNPC(ByVal npcind As Integer, ByVal NPCAlInd As Integer, ByVal indice As Integer)
10    On Error GoTo errorh
          Dim indireccion As Integer
          
20        indireccion = Npclist(npcind).Spells(indice)
          '' Envia las palabras magicas, fx y wav del indice-esimo hechizo del npc-hostiles.dat
30        Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead(Hechizos(indireccion).PalabrasMagicas, Npclist(npcind).Char.CharIndex, vbCyan))
40        Call SendData(SendTarget.ToNPCArea, NPCAlInd, PrepareMessageCreateFX(Npclist(NPCAlInd).Char.CharIndex, Hechizos(indireccion).FXgrh, Hechizos(indireccion).loops))
50        Call SendData(SendTarget.ToNPCArea, NPCAlInd, PrepareMessagePlayWave(Hechizos(indireccion).WAV, Npclist(NPCAlInd).Pos.X, Npclist(NPCAlInd).Pos.Y))
60        Npclist(NPCAlInd).Veneno = 0
70        Npclist(NPCAlInd).flags.Envenenado = 0

80    Exit Sub

errorh:
90        LogError ("Error en NPCAI.NPCRemueveVenenoNPC? ")

End Sub

Sub NPCCuraLevesNPC(ByVal npcind As Integer, ByVal NPCAlInd As Integer, ByVal indice As Integer)
10    On Error GoTo errorh
          Dim indireccion As Integer
          
20        indireccion = Npclist(npcind).Spells(indice)
          '' Envia las palabras magicas, fx y wav del indice-esimo hechizo del npc-hostiles.dat
30        Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead(Hechizos(indireccion).PalabrasMagicas, Npclist(npcind).Char.CharIndex, vbCyan))
40        Call SendData(SendTarget.ToNPCArea, NPCAlInd, PrepareMessagePlayWave(Hechizos(indireccion).WAV, Npclist(NPCAlInd).Pos.X, Npclist(NPCAlInd).Pos.Y))
50        Call SendData(SendTarget.ToNPCArea, NPCAlInd, PrepareMessageCreateFX(Npclist(NPCAlInd).Char.CharIndex, Hechizos(indireccion).FXgrh, Hechizos(indireccion).loops))
          
60        If (Npclist(NPCAlInd).Stats.MinHp + 5 < Npclist(NPCAlInd).Stats.MaxHp) Then
70            Npclist(NPCAlInd).Stats.MinHp = Npclist(NPCAlInd).Stats.MinHp + 5
80        Else
90            Npclist(NPCAlInd).Stats.MinHp = Npclist(NPCAlInd).Stats.MaxHp
100       End If
          
110   Exit Sub

errorh:
120       LogError ("Error en NPCAI.NPCCuraLevesNPC? ")
          
End Sub


Sub NPCRemueveParalisisNPC(ByVal npcind As Integer, ByVal NPCAlInd As Integer, ByVal indice As Integer)
10    On Error GoTo errorh
          Dim indireccion As Integer
          
20        indireccion = Npclist(npcind).Spells(indice)
          '' Envia las palabras magicas, fx y wav del indice-esimo hechizo del npc-hostiles.dat
30        Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead(Hechizos(indireccion).PalabrasMagicas, Npclist(npcind).Char.CharIndex, vbCyan))
40        Call SendData(SendTarget.ToNPCArea, NPCAlInd, PrepareMessagePlayWave(Hechizos(indireccion).WAV, Npclist(NPCAlInd).Pos.X, Npclist(NPCAlInd).Pos.Y))
50        Call SendData(SendTarget.ToNPCArea, NPCAlInd, PrepareMessageCreateFX(Npclist(NPCAlInd).Char.CharIndex, Hechizos(indireccion).FXgrh, Hechizos(indireccion).loops))
60        Npclist(NPCAlInd).Contadores.Paralisis = 0
70        Npclist(NPCAlInd).flags.Paralizado = 0
80    Exit Sub

errorh:
90        LogError ("Error en NPCAI.NPCRemueveParalisisNPC? ")

End Sub

Sub NPCparalizaNPC(ByVal paralizador As Integer, ByVal Paralizado As Integer, ByVal indice)
10    On Error GoTo errorh
          Dim indireccion As Integer
          
20        indireccion = Npclist(paralizador).Spells(indice)
          '' Envia las palabras magicas, fx y wav del indice-esimo hechizo del npc-hostiles.dat
30        Call SendData(SendTarget.ToNPCArea, paralizador, PrepareMessageChatOverHead(Hechizos(indireccion).PalabrasMagicas, Npclist(paralizador).Char.CharIndex, vbCyan))
40        Call SendData(SendTarget.ToNPCArea, Paralizado, PrepareMessagePlayWave(Hechizos(indireccion).WAV, Npclist(Paralizado).Pos.X, Npclist(Paralizado).Pos.Y))
50        Call SendData(SendTarget.ToNPCArea, Paralizado, PrepareMessageCreateFX(Npclist(Paralizado).Char.CharIndex, Hechizos(indireccion).FXgrh, Hechizos(indireccion).loops))
          
60        Npclist(Paralizado).flags.Paralizado = 1
70        Npclist(Paralizado).Contadores.Paralisis = IntervaloParalizado * 2

80    Exit Sub

errorh:
90        LogError ("Error en NPCAI.NPCParalizaNPC? ")

End Sub

Sub NPCcuraNPC(ByVal curador As Integer, ByVal curado As Integer, ByVal indice As Integer)
10    On Error GoTo errorh
          Dim indireccion As Integer
          

20        indireccion = Npclist(curador).Spells(indice)
          '' Envia las palabras magicas, fx y wav del indice-esimo hechizo del npc-hostiles.dat
30        Call SendData(SendTarget.ToNPCArea, curador, PrepareMessageChatOverHead(Hechizos(indireccion).PalabrasMagicas, Npclist(curador).Char.CharIndex, vbCyan))
40        Call SendData(SendTarget.ToNPCArea, curado, PrepareMessagePlayWave(Hechizos(indireccion).WAV, Npclist(curado).Pos.X, Npclist(curado).Pos.Y))
50        Call SendData(SendTarget.ToNPCArea, curado, PrepareMessageCreateFX(Npclist(curado).Char.CharIndex, Hechizos(indireccion).FXgrh, Hechizos(indireccion).loops))
60        If Npclist(curado).Stats.MinHp + 30 > Npclist(curado).Stats.MaxHp Then
70            Npclist(curado).Stats.MinHp = Npclist(curado).Stats.MaxHp
80        Else
90            Npclist(curado).Stats.MinHp = Npclist(curado).Stats.MinHp + 30
100       End If
110   Exit Sub

errorh:
120       LogError ("Error en NPCAI.NPCcuraNPC? ")

End Sub

Sub NPCLanzaCegueraPJ(ByVal npcind As Integer, ByVal PJEnInd As Integer, ByVal indice As Integer)
10    On Error GoTo errorh
          Dim indireccion As Integer
          
20        indireccion = Npclist(npcind).Spells(indice)
          '' Envia las palabras magicas, fx y wav del indice-esimo hechizo del npc-hostiles.dat
30        Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead(Hechizos(indireccion).PalabrasMagicas, Npclist(npcind).Char.CharIndex, vbCyan))
40        Call SendData(SendTarget.ToNPCArea, PJEnInd, PrepareMessagePlayWave(Hechizos(indireccion).WAV, UserList(PJEnInd).Pos.X, UserList(PJEnInd).Pos.Y))
50        Call SendData(SendTarget.ToPCArea, PJEnInd, PrepareMessageCreateFX(UserList(PJEnInd).Char.CharIndex, Hechizos(indireccion).FXgrh, Hechizos(indireccion).loops))
          
60        UserList(PJEnInd).flags.Ceguera = 1
70        UserList(PJEnInd).Counters.Ceguera = IntervaloInvisible
          ''Envia ceguera
80        Call WriteBlind(PJEnInd)
          ''bardea si es el rey
90        If Npclist(npcind).Name = "Rey Pretoriano" Then
100           Call WriteConsoleMsg(PJEnInd, "El rey pretoriano te ha vuelto ciego ", FontTypeNames.FONTTYPE_FIGHT)
110           Call WriteConsoleMsg(PJEnInd, "A la distancia escuchas las siguientes palabras: ¡Cobarde, no eres digno de luchar conmigo si escapas! ", FontTypeNames.FONTTYPE_VENENO)
120       End If

130   Exit Sub

errorh:
140       LogError ("Error en NPCAI.NPCLanzaCegueraPJ? ")
End Sub

Sub NPCLanzaEstupidezPJ(ByVal npcind As Integer, ByVal PJEnInd As Integer, ByVal indice As Integer)
10    On Error GoTo errorh
          Dim indireccion As Integer
          

20        indireccion = Npclist(npcind).Spells(indice)
          '' Envia las palabras magicas, fx y wav del indice-esimo hechizo del npc-hostiles.dat
30        Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead(Hechizos(indireccion).PalabrasMagicas, Npclist(npcind).Char.CharIndex, vbCyan))
40        Call SendData(SendTarget.ToNPCArea, PJEnInd, PrepareMessagePlayWave(Hechizos(indireccion).WAV, UserList(PJEnInd).Pos.X, UserList(PJEnInd).Pos.Y))
50        Call SendData(SendTarget.ToPCArea, PJEnInd, PrepareMessageCreateFX(UserList(PJEnInd).Char.CharIndex, Hechizos(indireccion).FXgrh, Hechizos(indireccion).loops))
60        UserList(PJEnInd).flags.Estupidez = 1
70        UserList(PJEnInd).Counters.Estupidez = IntervaloInvisible
          'manda estupidez
80        Call WriteDumb(PJEnInd)

          'bardea si es el rey
90        If Npclist(npcind).Name = "Rey Pretoriano" Then
100           Call WriteConsoleMsg(PJEnInd, "El rey pretoriano te ha vuelto estúpido.", FontTypeNames.FONTTYPE_FIGHT)
110       End If
120   Exit Sub

errorh:
130       LogError ("Error en NPCAI.NPCLanzaEstupidezPJ? ")
End Sub

Sub NPCRemueveInvisibilidad(ByVal npcind As Integer, ByVal PJEnInd As Integer, ByVal indice As Integer)
10    On Error GoTo errorh
          Dim indireccion As Integer
          
20        indireccion = Npclist(npcind).Spells(indice)
          '' Envia las palabras magicas, fx y wav del indice-esimo hechizo del npc-hostiles.dat
30        Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead(Hechizos(indireccion).PalabrasMagicas, Npclist(npcind).Char.CharIndex, vbCyan))
40        Call SendData(SendTarget.ToNPCArea, PJEnInd, PrepareMessagePlayWave(Hechizos(indireccion).WAV, UserList(PJEnInd).Pos.X, UserList(PJEnInd).Pos.Y))
50        Call SendData(SendTarget.ToPCArea, PJEnInd, PrepareMessageCreateFX(UserList(PJEnInd).Char.CharIndex, Hechizos(indireccion).FXgrh, Hechizos(indireccion).loops))
          
          'Sacamos el efecto de ocultarse
60        If UserList(PJEnInd).flags.Oculto = 1 Then
70            UserList(PJEnInd).Counters.TiempoOculto = 0
80            UserList(PJEnInd).flags.Oculto = 0
90            Call SetInvisible(PJEnInd, UserList(PJEnInd).Char.CharIndex, False)
              'Call SendData(SendTarget.ToPCArea, PJEnInd, PrepareMessageSetInvisible(UserList(PJEnInd).Char.CharIndex, False))
100           Call WriteConsoleMsg(PJEnInd, "¡Has sido detectado!", FontTypeNames.FONTTYPE_VENENO)
110       Else
          'sino, solo lo "iniciamos" en la sacada de invisibilidad.
120           Call WriteConsoleMsg(PJEnInd, "Comienzas a hacerte visible.", FontTypeNames.FONTTYPE_VENENO)
130           UserList(PJEnInd).Counters.Invisibilidad = IntervaloInvisible - 1
140       End If

          
150   Exit Sub

errorh:
160       LogError ("Error en NPCAI.NPCRemueveInvisibilidad ")

End Sub

Sub NpcLanzaSpellSobreUser2(ByVal NpcIndex As Integer, ByVal Userindex As Integer, ByVal spell As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 05/09/09
      '05/09/09: Pato - Ahora actualiza la vida del usuario atacado
      '***************************************************
10    On Error GoTo errorh
      ''  Igual a la otra pero ataca invisibles!!!
      '' (malditos controles de casos imposibles...)

20    If Npclist(NpcIndex).CanAttack = 0 Then Exit Sub
      'If UserList(UserIndex).Flags.Invisible = 1 Then Exit Sub

30    Npclist(NpcIndex).CanAttack = 0
      Dim daño As Integer

40    If Hechizos(spell).SubeHP = 1 Then

50        daño = RandomNumber(Hechizos(spell).MinHp, Hechizos(spell).MaxHp)
60        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(spell).WAV, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
70        Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(UserList(Userindex).Char.CharIndex, Hechizos(spell).FXgrh, Hechizos(spell).loops))

80        UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MinHp + daño
90        If UserList(Userindex).Stats.MinHp > UserList(Userindex).Stats.MaxHp Then UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MaxHp
          
100       Call WriteConsoleMsg(Userindex, Npclist(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
          
          'SendData SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.x, .Pos.Y, daño, DAMAGE_NORMAL)
110       Call WriteUpdateHP(Userindex)
120       Call WriteUpdateFollow(Userindex)
130   ElseIf Hechizos(spell).SubeHP = 2 Then
          
140       daño = RandomNumber(Hechizos(spell).MinHp, Hechizos(spell).MaxHp)
150       Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(spell).WAV, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
160       Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(UserList(Userindex).Char.CharIndex, Hechizos(spell).FXgrh, Hechizos(spell).loops))

170       If UserList(Userindex).flags.Privilegios And PlayerType.User Then UserList(Userindex).Stats.MinHp = UserList(Userindex).Stats.MinHp - daño
          
180       Call WriteConsoleMsg(Userindex, Npclist(NpcIndex).Name & " te ha quitado " & daño & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
          
          'Muere
190       If UserList(Userindex).Stats.MinHp < 1 Then
200           UserList(Userindex).Stats.MinHp = 0
210           Call UserDie(Userindex)
220       End If
          
230       Call WriteUpdateHP(Userindex)
240       Call WriteUpdateFollow(Userindex)
250   End If

260   If Hechizos(spell).Paraliza = 1 Then
270        If UserList(Userindex).flags.Paralizado = 0 Then
280             UserList(Userindex).flags.Paralizado = 1
290             UserList(Userindex).Counters.Paralisis = IntervaloParalizado
300             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(Hechizos(spell).WAV, UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y))
310             Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessageCreateFX(UserList(Userindex).Char.CharIndex, Hechizos(spell).FXgrh, Hechizos(spell).loops))

320             Call WriteParalizeOK(Userindex)

330        End If
340   End If

350   Exit Sub

errorh:
360       LogError ("Error en NPCAI.NPCLanzaSpellSobreUser2 ")


End Sub



Sub MagoDestruyeWand(ByVal npcind As Integer, ByVal bs As Byte, ByVal indice As Integer)
10    On Error GoTo errorh
          ''sonidos: 30 y 32, y no los cambien sino termina siendo muy chistoso...
          ''Para el FX utiliza el del hechizos(indice)
          Dim X As Integer
          Dim Y As Integer
          Dim PJInd As Integer
          Dim NPCPosX As Integer
          Dim NPCPosY As Integer
          Dim NPCPosM As Integer
          Dim danio As Double
          Dim dist As Double
          Dim danioI As Integer
          Dim MascotaInd As Integer
          Dim indireccion As Integer
          
20        Select Case bs
              Case 5
30                Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead("Rahma", Npclist(npcind).Char.CharIndex, vbGreen))
40                Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessagePlayWave(SONIDO_Dragon_VIVO, Npclist(npcind).Pos.X, Npclist(npcind).Pos.Y))
50            Case 4
60                Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead("vôrtax", Npclist(npcind).Char.CharIndex, vbGreen))
70                Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessagePlayWave(SONIDO_Dragon_VIVO, Npclist(npcind).Pos.X, Npclist(npcind).Pos.Y))
80            Case 3
90                Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead("Zill", Npclist(npcind).Char.CharIndex, vbGreen))
100               Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessagePlayWave(SONIDO_Dragon_VIVO, Npclist(npcind).Pos.X, Npclist(npcind).Pos.Y))
110           Case 2
120               Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead("yäkà E'nta", Npclist(npcind).Char.CharIndex, vbGreen))
130               Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessagePlayWave(SONIDO_Dragon_VIVO, Npclist(npcind).Pos.X, Npclist(npcind).Pos.Y))
140           Case 1
150               Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead("¡¡Koràtá!!", Npclist(npcind).Char.CharIndex, vbGreen))
160               Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessagePlayWave(SONIDO_Dragon_VIVO, Npclist(npcind).Pos.X, Npclist(npcind).Pos.Y))
170           Case 0
180               Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead(vbNullString, Npclist(npcind).Char.CharIndex, vbGreen))
190               Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessagePlayWave(SONIDO_Dragon_VIVO, Npclist(npcind).Pos.X, Npclist(npcind).Pos.Y))
200               NPCPosX = Npclist(npcind).Pos.X
210               NPCPosY = Npclist(npcind).Pos.Y
220               NPCPosM = Npclist(npcind).Pos.map
230               PJInd = 0
240               indireccion = Npclist(npcind).Spells(indice)
                  ''Daño masivo por destruccion de wand
250               For X = 8 To 95
260                   For Y = 8 To 95
270                       PJInd = MapData(NPCPosM, X, Y).Userindex
280                       MascotaInd = MapData(NPCPosM, X, Y).NpcIndex
290                       If PJInd > 0 Then
300                           dist = Sqr((UserList(PJInd).Pos.X - NPCPosX) ^ 2 + (UserList(PJInd).Pos.Y - NPCPosY) ^ 2)
310                           danio = 880 / (dist ^ (3 / 7))
320                           danioI = Abs(Int(danio))
                              ''efectiviza el danio
330                           If UserList(PJInd).flags.Privilegios And PlayerType.User Then UserList(PJInd).Stats.MinHp = UserList(PJInd).Stats.MinHp - danioI
                              
340                           Call WriteConsoleMsg(PJInd, Npclist(npcind).Name & " te ha quitado " & danioI & " puntos de vida al romper su vara.", FontTypeNames.FONTTYPE_FIGHT)
350                           Call SendData(SendTarget.ToPCArea, PJInd, PrepareMessagePlayWave(Hechizos(indireccion).WAV, UserList(PJInd).Pos.X, UserList(PJInd).Pos.Y))
360                           Call SendData(SendTarget.ToPCArea, PJInd, PrepareMessageCreateFX(UserList(PJInd).Char.CharIndex, Hechizos(indireccion).FXgrh, Hechizos(indireccion).loops))
                              
370                           If UserList(PJInd).Stats.MinHp < 1 Then
380                               UserList(PJInd).Stats.MinHp = 0
390                               Call UserDie(PJInd)
400                           End If
                          
410                       ElseIf (MascotaInd > 0) Then
420                           If (Npclist(MascotaInd).MaestroUser > 0) Then
                              
430                               dist = Sqr((Npclist(MascotaInd).Pos.X - NPCPosX) ^ 2 + (Npclist(MascotaInd).Pos.Y - NPCPosY) ^ 2)
440                               danio = 880 / (dist ^ (3 / 7))
450                               danioI = Abs(Int(danio))
                                  ''efectiviza el danio
460                               Npclist(MascotaInd).Stats.MinHp = Npclist(MascotaInd).Stats.MinHp - danioI
                                  
470                               Call SendData(SendTarget.ToNPCArea, MascotaInd, PrepareMessagePlayWave(Hechizos(indireccion).WAV, Npclist(MascotaInd).Pos.X, Npclist(MascotaInd).Pos.Y))
480                               Call SendData(SendTarget.ToNPCArea, MascotaInd, PrepareMessageCreateFX(Npclist(MascotaInd).Char.CharIndex, Hechizos(indireccion).FXgrh, Hechizos(indireccion).loops))
                                  
490                               If Npclist(MascotaInd).Stats.MinHp < 1 Then
500                                   Npclist(MascotaInd).Stats.MinHp = 0
510                                   Call MuereNpc(MascotaInd, 0)
520                               End If
530                           End If  ''es mascota
540                       End If  ''hay npc
                          
550                   Next Y
560               Next X
570       End Select

580   Exit Sub

errorh:
590       LogError ("Error en NPCAI.MagoDestruyeWand ")

End Sub


Sub GreedyWalkTo(ByVal npcorig As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
10    On Error GoTo errorh
      ''  Este procedimiento es llamado cada vez que un NPC deba ir
      ''  a otro lugar en el mismo mapa. Utiliza una técnica
      ''  de programación greedy no determinística.
      ''  Cada paso azaroso que me acerque al destino, es un buen paso.
      ''  Si no hay mejor paso válido, entonces hay que volver atrás y reintentar.
      ''  Si no puedo moverme, me considero piketeado
      ''  La funcion es larga, pero es O(1) - orden algorítmico temporal constante

      'Rapsodius - Changed Mod by And for speed

      Dim NPCx As Integer
      Dim NPCy As Integer
      Dim USRx As Integer
      Dim USRy As Integer
      Dim dual As Integer
      Dim Mapa As Integer

20    If Not (Npclist(npcorig).Pos.map = map) Then Exit Sub   ''si son distintos mapas abort

30    NPCx = Npclist(npcorig).Pos.X
40    NPCy = Npclist(npcorig).Pos.Y

50    If (NPCx = X And NPCy = Y) Then Exit Sub    ''ya llegué!!


      ''  Levanto las coordenadas del destino
60    USRx = X
70    USRy = Y
80    Mapa = map

      ''  moverse
90        If (NPCx > USRx) Then
100           If (NPCy < USRy) Then
                  ''NPC esta arriba a la derecha
110               dual = RandomNumber(0, 10)
120               If ((dual And 1) = 0) Then ''move down
130                   If LegalPos(Mapa, NPCx, NPCy + 1) Then
140                       Call MoverAba(npcorig)
150                       Exit Sub
160                   ElseIf LegalPos(Mapa, NPCx - 1, NPCy) Then
170                       Call MoverIzq(npcorig)
180                       Exit Sub
190                   ElseIf LegalPos(Mapa, NPCx + 1, NPCy) Then
200                       Call MoverDer(npcorig)
210                       Exit Sub
220                   ElseIf LegalPos(Mapa, NPCx, NPCy - 1) Then
230                       Call MoverArr(npcorig)
240                       Exit Sub
250                   Else
                          ''aqui no puedo ir a ningun lado. Hay q ver si me bloquean caspers
260                       If CasperBlock(npcorig) Then Call LiberarCasperBlock(npcorig)
270                   End If
                      
280               Else        ''random first move
290                   If LegalPos(Mapa, NPCx - 1, NPCy) Then
300                       Call MoverIzq(npcorig)
310                       Exit Sub
320                   ElseIf LegalPos(Mapa, NPCx, NPCy + 1) Then
330                       Call MoverAba(npcorig)
340                       Exit Sub
350                   ElseIf LegalPos(Mapa, NPCx + 1, NPCy) Then
360                       Call MoverDer(npcorig)
370                       Exit Sub
380                   ElseIf LegalPos(Mapa, NPCx, NPCy - 1) Then
390                       Call MoverArr(npcorig)
400                       Exit Sub
410                   Else
420                       If CasperBlock(npcorig) Then Call LiberarCasperBlock(npcorig)
430                   End If
440               End If  ''checked random first move
450           ElseIf (NPCy > USRy) Then   ''NPC esta abajo a la derecha
460               dual = RandomNumber(0, 10)
470               If ((dual And 1) = 0) Then ''move up
480                   If LegalPos(Mapa, NPCx, NPCy - 1) Then  ''U
490                       Call MoverArr(npcorig)
500                       Exit Sub
510                   ElseIf LegalPos(Mapa, NPCx - 1, NPCy) Then  ''L
520                       Call MoverIzq(npcorig)
530                       Exit Sub
540                   ElseIf LegalPos(Mapa, NPCx, NPCy + 1) Then  ''D
550                       Call MoverAba(npcorig)
560                       Exit Sub
570                   ElseIf LegalPos(Mapa, NPCx + 1, NPCy) Then  ''R
580                       Call MoverDer(npcorig)
590                       Exit Sub
600                   Else
610                       If CasperBlock(npcorig) Then Call LiberarCasperBlock(npcorig)
620                   End If
630               Else    ''random first move
640                   If LegalPos(Mapa, NPCx - 1, NPCy) Then  ''L
650                       Call MoverIzq(npcorig)
660                       Exit Sub
670                   ElseIf LegalPos(Mapa, NPCx, NPCy - 1) Then ''U
680                       Call MoverArr(npcorig)
690                       Exit Sub
700                   ElseIf LegalPos(Mapa, NPCx, NPCy + 1) Then  ''D
710                       Call MoverAba(npcorig)
720                       Exit Sub
730                   ElseIf LegalPos(Mapa, NPCx + 1, NPCy) Then  ''R
740                       Call MoverDer(npcorig)
750                       Exit Sub
760                   Else
770                       If CasperBlock(npcorig) Then Call LiberarCasperBlock(npcorig)
780                   End If
790               End If  ''endif random first move
800           Else    ''x completitud, esta en la misma Y
810               If LegalPos(Mapa, NPCx - 1, NPCy) Then  ''L
820                   Call MoverIzq(npcorig)
830                   Exit Sub
840               ElseIf LegalPos(Mapa, NPCx, NPCy + 1) Then  ''D
850                   Call MoverAba(npcorig)
860                   Exit Sub
870               ElseIf LegalPos(Mapa, NPCx, NPCy - 1) Then  ''U
880                   Call MoverArr(npcorig)
890                   Exit Sub
900               Else
                      ''si me muevo abajo entro en loop. Aca el algoritmo falla
910                   If Npclist(npcorig).CanAttack = 1 And (RandomNumber(1, 100) > 95) Then
920                       Call SendData(SendTarget.ToNPCArea, npcorig, PrepareMessageChatOverHead("Maldito bastardo, ¡Ven aquí!", Str(Npclist(npcorig).Char.CharIndex), vbYellow))
930                       Npclist(npcorig).CanAttack = 0
940                   End If
950               End If
960           End If
          
970       ElseIf (NPCx < USRx) Then
              
980           If (NPCy < USRy) Then
                  ''NPC esta arriba a la izquierda
990               dual = RandomNumber(0, 10)
1000              If ((dual And 1) = 0) Then ''move down
1010                  If LegalPos(Mapa, NPCx, NPCy + 1) Then  ''ABA
1020                      Call MoverAba(npcorig)
1030                      Exit Sub
1040                  ElseIf LegalPos(Mapa, NPCx + 1, NPCy) Then  ''R
1050                      Call MoverDer(npcorig)
1060                      Exit Sub
1070                  ElseIf LegalPos(Mapa, NPCx - 1, NPCy) Then
1080                      Call MoverIzq(npcorig)
1090                      Exit Sub
1100                  ElseIf LegalPos(Mapa, NPCx, NPCy - 1) Then
1110                      Call MoverArr(npcorig)
1120                      Exit Sub
1130                  Else
1140                      If CasperBlock(npcorig) Then Call LiberarCasperBlock(npcorig)
1150                  End If
1160              Else    ''random first move
1170                  If LegalPos(Mapa, NPCx + 1, NPCy) Then  ''DER
1180                      Call MoverDer(npcorig)
1190                      Exit Sub
1200                  ElseIf LegalPos(Mapa, NPCx, NPCy + 1) Then  ''ABA
1210                      Call MoverAba(npcorig)
1220                      Exit Sub
1230                  ElseIf LegalPos(Mapa, NPCx - 1, NPCy) Then
1240                      Call MoverIzq(npcorig)
1250                      Exit Sub
1260                  ElseIf LegalPos(Mapa, NPCx, NPCy - 1) Then
1270                      Call MoverArr(npcorig)
1280                      Exit Sub
1290                  Else
1300                      If CasperBlock(npcorig) Then Call LiberarCasperBlock(npcorig)
1310                  End If
1320              End If
              
1330          ElseIf (NPCy > USRy) Then   ''NPC esta abajo a la izquierda
1340              dual = RandomNumber(0, 10)
1350              If ((dual And 1) = 0) Then ''move up
1360                  If LegalPos(Mapa, NPCx, NPCy - 1) Then  ''U
1370                      Call MoverArr(npcorig)
1380                      Exit Sub
1390                  ElseIf LegalPos(Mapa, NPCx + 1, NPCy) Then  ''R
1400                      Call MoverDer(npcorig)
1410                      Exit Sub
1420                  ElseIf LegalPos(Mapa, NPCx - 1, NPCy) Then  ''L
1430                      Call MoverIzq(npcorig)
1440                      Exit Sub
1450                  ElseIf LegalPos(Mapa, NPCx, NPCy + 1) Then  ''D
1460                      Call MoverAba(npcorig)
1470                      Exit Sub
1480                  Else
1490                      If CasperBlock(npcorig) Then Call LiberarCasperBlock(npcorig)
1500                  End If
1510              Else
1520                  If LegalPos(Mapa, NPCx + 1, NPCy) Then  ''R
1530                      Call MoverDer(npcorig)
1540                      Exit Sub
1550                  ElseIf LegalPos(Mapa, NPCx, NPCy - 1) Then  ''U
1560                      Call MoverArr(npcorig)
1570                      Exit Sub
1580                  ElseIf LegalPos(Mapa, NPCx, NPCy + 1) Then  ''D
1590                      Call MoverAba(npcorig)
1600                      Exit Sub
1610                  ElseIf LegalPos(Mapa, NPCx - 1, NPCy) Then  ''L
1620                      Call MoverIzq(npcorig)
1630                      Exit Sub
1640                  Else
1650                      If CasperBlock(npcorig) Then Call LiberarCasperBlock(npcorig)
1660                  End If
1670              End If
1680          Else    ''x completitud, esta en la misma Y
1690              If LegalPos(Mapa, NPCx + 1, NPCy) Then  ''R
1700                  Call MoverDer(npcorig)
1710                  Exit Sub
1720              ElseIf LegalPos(Mapa, NPCx, NPCy + 1) Then  ''D
1730                  Call MoverAba(npcorig)
1740                  Exit Sub
1750              ElseIf LegalPos(Mapa, NPCx, NPCy - 1) Then  ''U
1760                  Call MoverArr(npcorig)
1770                  Exit Sub
1780              Else
                      ''si me muevo loopeo. aca falla el algoritmo
1790                  If Npclist(npcorig).CanAttack = 1 And (RandomNumber(1, 100) > 95) Then
1800                      Call SendData(SendTarget.ToNPCArea, npcorig, PrepareMessageChatOverHead("Maldito bastardo, ¡Ven aquí!", Npclist(npcorig).Char.CharIndex, vbYellow))
1810                      Npclist(npcorig).CanAttack = 0
1820                  End If
1830              End If
1840          End If
          
          
1850      Else ''igual X
1860          If (NPCy > USRy) Then    ''NPC ESTA ABAJO
1870              If LegalPos(Mapa, NPCx, NPCy - 1) Then  ''U
1880                  Call MoverArr(npcorig)
1890                  Exit Sub
1900              ElseIf LegalPos(Mapa, NPCx + 1, NPCy) Then  ''R
1910                  Call MoverDer(npcorig)
1920                  Exit Sub
1930              ElseIf LegalPos(Mapa, NPCx - 1, NPCy) Then  ''L
1940                  Call MoverIzq(npcorig)
1950                  Exit Sub
1960              Else
                      ''aca tambien falla el algoritmo
1970                  If Npclist(npcorig).CanAttack = 1 And (RandomNumber(1, 100) > 95) Then
1980                      Call SendData(SendTarget.ToNPCArea, npcorig, PrepareMessageChatOverHead("Maldito bastardo, ¡Ven aquí!", Npclist(npcorig).Char.CharIndex, vbYellow))
1990                      Npclist(npcorig).CanAttack = 0
2000                  End If
2010              End If
2020          Else    ''NPC ESTA ARRIBA
2030              If LegalPos(Mapa, NPCx, NPCy + 1) Then  ''ABA
2040                  Call MoverAba(npcorig)
2050                  Exit Sub
2060              ElseIf LegalPos(Mapa, NPCx + 1, NPCy) Then  ''R
2070                  Call MoverDer(npcorig)
2080                  Exit Sub
2090              ElseIf LegalPos(Mapa, NPCx - 1, NPCy) Then  ''L
2100                  Call MoverIzq(npcorig)
2110                  Exit Sub
2120              Else
                      ''posible loop
2130                  If Npclist(npcorig).CanAttack = 1 And (RandomNumber(1, 100) > 95) Then
2140                      Call SendData(SendTarget.ToNPCArea, npcorig, PrepareMessageChatOverHead("Maldito bastardo, ¡Ven aquí!", Npclist(npcorig).Char.CharIndex, vbYellow))
2150                      Npclist(npcorig).CanAttack = 0
2160                  End If
2170              End If
2180          End If
2190      End If

2200  Exit Sub

errorh:
2210      LogError ("Error en NPCAI.GreedyWalkTo")

End Sub

Sub MoverAba(ByVal npcorig As Integer)
10    On Error GoTo errorh

          Dim Mapa As Integer
          Dim NPCx As Integer
          Dim NPCy As Integer
20        Mapa = Npclist(npcorig).Pos.map
30        NPCx = Npclist(npcorig).Pos.X
40        NPCy = Npclist(npcorig).Pos.Y
          
50        Call SendData(SendTarget.ToNPCArea, npcorig, PrepareMessageCharacterMove(Npclist(npcorig).Char.CharIndex, NPCx, NPCy + 1))
          'Update map and npc pos
60        MapData(Mapa, NPCx, NPCy).NpcIndex = 0
70        Npclist(npcorig).Pos.Y = NPCy + 1
80        Npclist(npcorig).Char.Heading = eHeading.SOUTH
90        MapData(Mapa, NPCx, NPCy + 1).NpcIndex = npcorig
          
          'Revisamos sidebemos cambair el área
100       Call ModAreas.CheckUpdateNeededNpc(npcorig, SOUTH)
110   Exit Sub

errorh:
120       LogError ("Error en NPCAI.MoverAba ")

End Sub

Sub MoverArr(ByVal npcorig As Integer)
10    On Error GoTo errorh

          Dim Mapa As Integer
          Dim NPCx As Integer
          Dim NPCy As Integer
20        Mapa = Npclist(npcorig).Pos.map
30        NPCx = Npclist(npcorig).Pos.X
40        NPCy = Npclist(npcorig).Pos.Y
          
50        Call SendData(SendTarget.ToNPCArea, npcorig, PrepareMessageCharacterMove(Npclist(npcorig).Char.CharIndex, NPCx, NPCy - 1))
          'Update map and npc pos
60        MapData(Mapa, NPCx, NPCy).NpcIndex = 0
70        Npclist(npcorig).Pos.Y = NPCy - 1
80        Npclist(npcorig).Char.Heading = eHeading.NORTH
90        MapData(Mapa, NPCx, NPCy - 1).NpcIndex = npcorig
          
          'Revisamos sidebemos cambair el área
100       Call ModAreas.CheckUpdateNeededNpc(npcorig, NORTH)
110   Exit Sub

errorh:
120       LogError ("Error en NPCAI.MoverArr")
End Sub

Sub MoverIzq(ByVal npcorig As Integer)
10    On Error GoTo errorh

          Dim Mapa As Integer
          Dim NPCx As Integer
          Dim NPCy As Integer
20        Mapa = Npclist(npcorig).Pos.map
30        NPCx = Npclist(npcorig).Pos.X
40        NPCy = Npclist(npcorig).Pos.Y

50        Call SendData(SendTarget.ToNPCArea, npcorig, PrepareMessageCharacterMove(Npclist(npcorig).Char.CharIndex, NPCx - 1, NPCy))
          'Update map and npc pos
60        MapData(Mapa, NPCx, NPCy).NpcIndex = 0
70        Npclist(npcorig).Pos.X = NPCx - 1
80        Npclist(npcorig).Char.Heading = eHeading.WEST
90        MapData(Mapa, NPCx - 1, NPCy).NpcIndex = npcorig
          
          'Revisamos sidebemos cambair el área
100       Call ModAreas.CheckUpdateNeededNpc(npcorig, WEST)
110   Exit Sub

errorh:
120       LogError ("Error en NPCAI.MoverIzq")

End Sub

Sub MoverDer(ByVal npcorig As Integer)
10    On Error GoTo errorh

          Dim Mapa As Integer
          Dim NPCx As Integer
          Dim NPCy As Integer
20        Mapa = Npclist(npcorig).Pos.map
30        NPCx = Npclist(npcorig).Pos.X
40        NPCy = Npclist(npcorig).Pos.Y
          
50        Call SendData(SendTarget.ToNPCArea, npcorig, PrepareMessageCharacterMove(Npclist(npcorig).Char.CharIndex, NPCx + 1, NPCy))
          'Update map and npc pos
60        MapData(Mapa, NPCx, NPCy).NpcIndex = 0
70        Npclist(npcorig).Pos.X = NPCx + 1
80        Npclist(npcorig).Char.Heading = eHeading.EAST
90        MapData(Mapa, NPCx + 1, NPCy).NpcIndex = npcorig
          
          'Revisamos sidebemos cambair el área
100       Call ModAreas.CheckUpdateNeededNpc(npcorig, EAST)
110   Exit Sub

errorh:
120       LogError ("Error en NPCAI.MoverDer")

End Sub


Sub VolverAlCentro(ByVal npcind As Integer)
10    On Error GoTo errorh
          
          Dim NPCPosX As Integer
          Dim NPCPosY As Integer
          Dim NpcMap As Integer
20        NPCPosX = Npclist(npcind).Pos.X
30        NPCPosY = Npclist(npcind).Pos.Y
40        NpcMap = Npclist(npcind).Pos.map
          
50        If NpcMap = MAPA_PRETORIANO Then
              ''35,25 y 67,25 son las posiciones del rey
60            If NPCPosX < 50 Then    ''esta a la izquierda
70                Call GreedyWalkTo(npcind, NpcMap, ALCOBA1_X, ALCOBA1_Y)
                  'GreedyWalkTo npcind, NpcMap, 35, 25
80            Else
90                Call GreedyWalkTo(npcind, NpcMap, ALCOBA2_X, ALCOBA2_Y)
                  'GreedyWalkTo npcind, NpcMap, 67, 25
100           End If
110       End If

120   Exit Sub

errorh:
130       LogError ("Error en NPCAI.VolverAlCentro")

End Sub

Function EstoyMuyLejos(ByVal npcind) As Boolean
      ''me dice si estoy fuera del anillo exterior de proteccion
      ''de los clerigos
          
          Dim retvalue As Boolean
          
          'If Npclist(npcind).Pos.X < 50 Then
          '    retvalue = Npclist(npcind).Pos.X < 43 And Npclist(npcind).Pos.X > 27
          'Else
          '    retvalue = Npclist(npcind).Pos.X < 80 And Npclist(npcind).Pos.X > 59
          'End If
          
10        retvalue = Npclist(npcind).Pos.Y > 39
          
20        If Not Npclist(npcind).Pos.map = MAPA_PRETORIANO Then
30            EstoyMuyLejos = False
40        Else
50            EstoyMuyLejos = retvalue
60        End If

70    Exit Function

errorh:
80        LogError ("Error en NPCAI.EstoymUYLejos")

End Function

Function EstoyLejos(ByVal npcind) As Boolean
10    On Error GoTo errorh

          ''35,25 y 67,25 son las posiciones del rey
          ''esta fction me indica si estoy lejos del rango de vision
          
          
          Dim retvalue As Boolean
          
20        If Npclist(npcind).Pos.X < 50 Then
30            retvalue = Npclist(npcind).Pos.X < 43 And Npclist(npcind).Pos.X > 27
40        Else
50            retvalue = Npclist(npcind).Pos.X < 75 And Npclist(npcind).Pos.X > 59
60        End If
          
70        retvalue = retvalue And Npclist(npcind).Pos.Y > 19 And Npclist(npcind).Pos.Y < 31
          
80        If Not Npclist(npcind).Pos.map = MAPA_PRETORIANO Then
90            EstoyLejos = False
100       Else
110           EstoyLejos = Not retvalue
120       End If

130   Exit Function

errorh:
140       LogError ("Error en NPCAI.EstoyLejos")

End Function

Function EsAlcanzable(ByVal npcind As Integer, ByVal PJEnInd As Integer) As Boolean
10    On Error GoTo errorh
          
          ''esta funcion es especialmente hecha para el mapa pretoriano
          ''Está diseñada para que se ignore a los PJs que estan demasiado lejos
          ''evitando asi que los "lockeen" en la pelea sacandolos de combate
          ''sin matarlos. La fcion es totalmente inutil si los NPCs estan en otro mapa.
          ''Chequea la posibilidad que les hagan /racc desde otro mapa para evitar
          ''malos comportamientos
          ''35,25 y 67,25 son las posiciones del rey
      ''On Error Resume Next


          Dim retvalue As Boolean
          Dim retValue2 As Boolean
          
          Dim PJPosX As Integer
          Dim PJPosY As Integer
          Dim NPCPosX As Integer
          Dim NPCPosY As Integer
          
20        PJPosX = UserList(PJEnInd).Pos.X
30        PJPosY = UserList(PJEnInd).Pos.Y
40        NPCPosX = Npclist(npcind).Pos.X
50        NPCPosY = Npclist(npcind).Pos.Y
          
60        If (Npclist(npcind).Pos.map = MAPA_PRETORIANO) And (UserList(PJEnInd).Pos.map = MAPA_PRETORIANO) Then
              ''los bounds del mapa pretoriano son fijos.
              ''Esta en una posicion alcanzable si esta dentro del
              ''espacio de las alcobas reales del mapa diseñado por mi.
              ''Y dentro de la alcoba en el rango del perimetro de defensa
              '' 8+8+8+8 x 7+7+7+7
70            retvalue = PJPosX > 18 And PJPosX < 49 And NPCPosX <= 51 'And NPCPosX < 49
80            retvalue = retvalue And (PJPosY > 14 And PJPosY < 40) 'And NPCPosY > 14 And NPCPosY < 50)
90            retValue2 = PJPosX > 52 And PJPosX < 81 And NPCPosX > 51 'And NPCPosX < 81
100           retValue2 = retValue2 And (PJPosY > 14 And PJPosY < 40) 'And NPCPosY > 14 And NPCPosY < 50)
              ''rv dice si estan en la alcoba izquierda los 2 y en zona valida de combate
              ''rv2 dice si estan en la derecha
110           retvalue = retvalue Or retValue2
              'If retvalue = False Then
              '    If Npclist(npcind).CanAttack = 1 Then
              '        Call SendData(SendTarget.ToNPCArea, npcind, Npclist(npcind).Pos.Map, "||" & vbYellow & "°¡ Cobarde !°" & str(Npclist(npcind).Char.CharIndex))
              '        Npclist(npcind).CanAttack = 0
              '    End If
              'End If
120       Else
130           retvalue = False
140       End If
          
150       EsAlcanzable = retvalue
           
160   Exit Function

errorh:
170       LogError ("Error en NPCAI.EsAlcanzable")
       
       
End Function



Function CasperBlock(ByVal Npc As Integer) As Boolean
10    On Error GoTo errorh
          
          Dim NPCPosM As Integer
          Dim NPCPosX As Integer
          Dim NPCPosY As Integer
          Dim PJ As Integer
          
          Dim retvalue As Boolean
          
20        NPCPosX = Npclist(Npc).Pos.X
30        NPCPosY = Npclist(Npc).Pos.Y
40        NPCPosM = Npclist(Npc).Pos.map
          
50        retvalue = Not (LegalPos(NPCPosM, NPCPosX + 1, NPCPosY) Or _
                      LegalPos(NPCPosM, NPCPosX - 1, NPCPosY) Or _
                      LegalPos(NPCPosM, NPCPosX, NPCPosY + 1) Or _
                      LegalPos(NPCPosM, NPCPosX, NPCPosY - 1))
                      
60        If retvalue Then
              ''si son todas invalidas
              ''busco que algun casper sea causante de piketeo
70            retvalue = False

80            PJ = MapData(NPCPosM, NPCPosX + 1, NPCPosY).Userindex
90            If PJ > 0 Then
100               retvalue = UserList(PJ).flags.Muerto = 1
110           End If
              
120           PJ = MapData(NPCPosM, NPCPosX - 1, NPCPosY).Userindex
130           If PJ > 0 Then
140               retvalue = retvalue Or UserList(PJ).flags.Muerto = 1
150           End If
              
160           PJ = MapData(NPCPosM, NPCPosX, NPCPosY + 1).Userindex
170           If PJ > 0 Then
180               retvalue = retvalue Or UserList(PJ).flags.Muerto = 1
190           End If
              
200           PJ = MapData(NPCPosM, NPCPosX, NPCPosY - 1).Userindex
210           If PJ > 0 Then
220               retvalue = retvalue Or UserList(PJ).flags.Muerto = 1
230           End If
              
240       Else
250           retvalue = False
          
260       End If
          
270       CasperBlock = retvalue
280       Exit Function

errorh:
      '    MsgBox ("ERROR!!")
290       CasperBlock = False
300       LogError ("Error en NPCAI.CasperBlock")

End Function


Sub LiberarCasperBlock(ByVal npcind As Integer)
10    On Error GoTo errorh

          Dim NPCPosX As Integer
          Dim NPCPosY As Integer
          Dim NPCPosM As Integer
          
20        NPCPosX = Npclist(npcind).Pos.X
30        NPCPosY = Npclist(npcind).Pos.Y
40        NPCPosM = Npclist(npcind).Pos.map
          
50        If LegalPos(NPCPosM, NPCPosX + 1, NPCPosY + 1) Then
60            Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageCharacterMove(Npclist(npcind).Char.CharIndex, NPCPosX + 1, NPCPosY + 1))
              'Update map and npc pos
70            MapData(NPCPosM, NPCPosX, NPCPosY).NpcIndex = 0
80            Npclist(npcind).Pos.Y = NPCPosY + 1
90            Npclist(npcind).Pos.X = NPCPosX + 1
100           Npclist(npcind).Char.Heading = eHeading.SOUTH
110           MapData(NPCPosM, NPCPosX + 1, NPCPosY + 1).NpcIndex = npcind
120           Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead("¡¡JA JA JA JA!!", Npclist(npcind).Char.CharIndex, vbYellow))
130           Exit Sub
140       End If

150       If LegalPos(NPCPosM, NPCPosX - 1, NPCPosY - 1) Then
160           Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageCharacterMove(Npclist(npcind).Char.CharIndex, NPCPosX - 1, NPCPosY - 1))
              'Update map and npc pos
170           MapData(NPCPosM, NPCPosX, NPCPosY).NpcIndex = 0
180           Npclist(npcind).Pos.Y = NPCPosY - 1
190           Npclist(npcind).Pos.X = NPCPosX - 1
200           Npclist(npcind).Char.Heading = eHeading.NORTH
210           MapData(NPCPosM, NPCPosX - 1, NPCPosY - 1).NpcIndex = npcind
220           Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead("¡¡JA JA JA JA!!", Npclist(npcind).Char.CharIndex, vbYellow))
230           Exit Sub
240       End If

250       If LegalPos(NPCPosM, NPCPosX + 1, NPCPosY - 1) Then
260           Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageCharacterMove(Npclist(npcind).Char.CharIndex, NPCPosX + 1, NPCPosY - 1))
              'Update map and npc pos
270           MapData(NPCPosM, NPCPosX, NPCPosY).NpcIndex = 0
280           Npclist(npcind).Pos.Y = NPCPosY - 1
290           Npclist(npcind).Pos.X = NPCPosX + 1
300           Npclist(npcind).Char.Heading = eHeading.EAST
310           MapData(NPCPosM, NPCPosX + 1, NPCPosY - 1).NpcIndex = npcind
320           Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead("¡¡JA JA JA JA!!", Npclist(npcind).Char.CharIndex, vbYellow))
330           Exit Sub
340       End If
          
350       If LegalPos(NPCPosM, NPCPosX - 1, NPCPosY + 1) Then
360           Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageCharacterMove(Npclist(npcind).Char.CharIndex, NPCPosX - 1, NPCPosY + 1))
              'Update map and npc pos
370           MapData(NPCPosM, NPCPosX, NPCPosY).NpcIndex = 0
380           Npclist(npcind).Pos.Y = NPCPosY + 1
390           Npclist(npcind).Pos.X = NPCPosX - 1
400           Npclist(npcind).Char.Heading = eHeading.WEST
410           MapData(NPCPosM, NPCPosX - 1, NPCPosY + 1).NpcIndex = npcind
420           Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead("¡¡JA JA JA JA!!", Npclist(npcind).Char.CharIndex, vbYellow))
430           Exit Sub
440       End If
          
          ''si esta aca, estamos fritos!
450       If Npclist(npcind).CanAttack = 1 Then
460           Call SendData(SendTarget.ToNPCArea, npcind, PrepareMessageChatOverHead("¡Por las barbas de los antiguos reyes! ¡Alejáos endemoniados espectros o sufriréis la furia de los dioses!", Npclist(npcind).Char.CharIndex, vbYellow))
470           Npclist(npcind).CanAttack = 0
480       End If
          
490   Exit Sub

errorh:
500       LogError ("Error en NPCAI.LiberarCasperBlock")

End Sub

Public Sub CambiarAlcoba(ByVal npcind As Integer)
10    On Error GoTo errorh

20        Select Case Npclist(npcind).Invent.ArmourEqpSlot
              Case 2
30                Call GreedyWalkTo(npcind, MAPA_PRETORIANO, 48, 70)
40                If Npclist(npcind).Pos.X = 48 And Npclist(npcind).Pos.Y = 70 Then Npclist(npcind).Invent.ArmourEqpSlot = Npclist(npcind).Invent.ArmourEqpSlot + 1
50            Case 6
60                Call GreedyWalkTo(npcind, MAPA_PRETORIANO, 52, 71)
70                If Npclist(npcind).Pos.X = 52 And Npclist(npcind).Pos.Y = 71 Then Npclist(npcind).Invent.ArmourEqpSlot = Npclist(npcind).Invent.ArmourEqpSlot + 1
80            Case 1
90                Call GreedyWalkTo(npcind, MAPA_PRETORIANO, 73, 56)
100               If Npclist(npcind).Pos.X = 73 And Npclist(npcind).Pos.Y = 56 Then Npclist(npcind).Invent.ArmourEqpSlot = Npclist(npcind).Invent.ArmourEqpSlot + 1
110           Case 7
120               Call GreedyWalkTo(npcind, MAPA_PRETORIANO, 73, 48)
130               If Npclist(npcind).Pos.X = 73 And Npclist(npcind).Pos.Y = 48 Then Npclist(npcind).Invent.ArmourEqpSlot = Npclist(npcind).Invent.ArmourEqpSlot + 1
140           Case 5
150               Call GreedyWalkTo(npcind, MAPA_PRETORIANO, 31, 56)
160               If Npclist(npcind).Pos.X = 31 And Npclist(npcind).Pos.Y = 56 Then Npclist(npcind).Invent.ArmourEqpSlot = Npclist(npcind).Invent.ArmourEqpSlot + 1
170           Case 3
180               Call GreedyWalkTo(npcind, MAPA_PRETORIANO, 31, 48)
190               If Npclist(npcind).Pos.X = 31 And Npclist(npcind).Pos.Y = 48 Then Npclist(npcind).Invent.ArmourEqpSlot = Npclist(npcind).Invent.ArmourEqpSlot + 1
200           Case 4, 8
210               Npclist(npcind).Invent.ArmourEqpSlot = 0
220               Exit Sub
230       End Select

240   Exit Sub
errorh:
250   Call LogError("Error en CambiarAlcoba " & Err.Description)
End Sub
