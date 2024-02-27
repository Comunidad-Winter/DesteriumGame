Attribute VB_Name = "NPCs"
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


'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Contiene todas las rutinas necesarias para cotrolar los
'NPCs meno la rutina de AI que se encuentra en el modulo
'AI_NPCs para su mejor comprension.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Option Explicit

Sub QuitarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim i As Integer
          
10        For i = 1 To MAXMASCOTAS
20          If UserList(UserIndex).MascotasIndex(i) = NpcIndex Then
30             UserList(UserIndex).MascotasIndex(i) = 0
40             UserList(UserIndex).MascotasType(i) = 0
               
50             UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas - 1
60             Exit For
70          End If
80        Next i
End Sub

Sub QuitarMascotaNpc(ByVal Maestro As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1
End Sub

Sub MuereNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
        '********************************************************
        'Author: Unknown
        'Llamado cuando la vida de un NPC llega a cero.
        'Last Modify Date: 24/01/2007
        '22/06/06: (Nacho) Chequeamos si es pretoriano
        '24/01/2007: Pablo (ToxicWaste): Agrego para actualización de tag si cambia de status.
        '********************************************************
     
        '<EhHeader>
10      On Error GoTo MuereNpc_Err
 
        '</EhHeader>
        Dim MiNPC As Npc
        Dim i As Long
        
20      MiNPC = Npclist(NpcIndex)
        Dim EraCriminal  As Boolean
        Dim IsPretoriano As Boolean
        
30       If MiNPC.Numero = 697 Then
40          If UserList(UserIndex).flags.SlotEvent > 0 Then
50              FinishCastleMode UserList(UserIndex).flags.SlotEvent, UserList(UserIndex).flags.SlotUserEvent
60          End If
70      End If

75        If MiNPC.Invasion = 1 Then
            mInvasiones.MuereNpcInvasion UserIndex, MiNPC.Invasion, MiNPC.DropIndex
            Npclist(NpcIndex).DropIndex = 0
            Npclist(NpcIndex).Invasion = 0
        End If
        
        ' Npc de invocacion
80         If MiNPC.flags.Invocacion = 1 Then
90           For i = 1 To NumInvocaciones
100              If Invocaciones(i).NpcIndex = Npclist(NpcIndex).Numero Then
110                  Invocaciones(i).Activo = 0
120              End If
130          Next i
140     End If
         
        
        
150     If (esPretoriano(NpcIndex) = 4) Then
            'Solo nos importa si fue matado en el mapa pretoriano.
160         IsPretoriano = True
 
170         If Npclist(NpcIndex).Pos.map = MAPA_PRETORIANO Then
                'seteamos todos estos 'flags' acorde para que cambien solos de alcoba
                Dim j    As Integer
                Dim NPCI As Integer
     
180             For i = 8 To 90
190                 For j = 8 To 90
             
200                     NPCI = MapData(Npclist(NpcIndex).Pos.map, i, j).NpcIndex
 
210                     If NPCI > 0 Then
220                         If esPretoriano(NPCI) > 0 And NPCI <> NpcIndex Then
230                             If Npclist(NpcIndex).Pos.X > 50 Then
240                                 If Npclist(NPCI).Pos.X > 50 Then Npclist(NPCI).Invent.ArmourEqpSlot = 1
250                             Else
 
260                                 If Npclist(NPCI).Pos.X <= 50 Then Npclist(NPCI).Invent.ArmourEqpSlot = 5
 
270                             End If
 
280                         End If
 
290                     End If
 
300                 Next j
310             Next i
 
320             Call CrearClanPretoriano(Npclist(NpcIndex).Pos.X)
 
330         End If
 
340     ElseIf esPretoriano(NpcIndex) > 0 Then
350         IsPretoriano = True
 
360         If Npclist(NpcIndex).Pos.map = MAPA_PRETORIANO Then
370             Npclist(NpcIndex).Invent.ArmourEqpSlot = 0
380             pretorianosVivos = pretorianosVivos - 1
 
390         End If
 
400     End If
        'Quitamos el npc
410     Call QuitarNPC(NpcIndex)
 
420     If UserIndex > 0 Then ' Lo mato un usuario?
 
430         With UserList(UserIndex)
     
440             If MiNPC.flags.Snd3 > 0 Then
450                 Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MiNPC.flags.Snd3, _
                            MiNPC.Pos.X, MiNPC.Pos.Y))
 
460             End If
 
470             .flags.TargetNPC = 0
480             .flags.TargetNpcTipo = eNPCType.Comun
         
                'El user que lo mato tiene mascotas?
490             If .NroMascotas > 0 Then
                    Dim t As Integer
 
500                 For t = 1 To MAXMASCOTAS
 
510                     If .MascotasIndex(t) > 0 Then
520                         If Npclist(.MascotasIndex(t)).TargetNPC = NpcIndex Then
530                             Call FollowAmo(.MascotasIndex(t))
 
540                         End If
 
550                     End If
 
560                 Next t
 
570             End If
         
                '[KEVIN]
580    If MiNPC.flags.ExpCount > 0 Then
590             If .GroupIndex > 0 Then
600                 'Call mdParty.ObtenerExito(UserIndex, MiNPC.flags.ExpCount, MiNPC.Pos.map, MiNPC.Pos.X, MiNPC.Pos.Y)
                    Call mGroup.AddExpGroup(.GroupIndex, MiNPC.flags.ExpCount)
610             Else
620                 .Stats.Exp = .Stats.Exp + MiNPC.flags.ExpCount
630                 If .Stats.Exp > MAXEXP Then _
                        .Stats.Exp = MAXEXP
640                 Call WriteConsoleMsg(UserIndex, "Has ganado " & MiNPC.flags.ExpCount & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
650                  If UserList(UserIndex).flags.Oro = 1 Then
660         UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + (MiNPC.flags.ExpCount * 0.4)
670              WriteConsoleMsg UserIndex, "Aumento de exp 40%> Has ganado " & (MiNPC.flags.ExpCount * 0.4) & " puntos de experiencia.", FontTypeNames.fonttype_dios
680     End If
690             End If
700             MiNPC.flags.ExpCount = 0
710         End If
         
                '[/KEVIN]
720             Call WriteConsoleMsg(UserIndex, "¡Has matado a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
 
730             If .Stats.NPCsMuertos < 32000 Then .Stats.NPCsMuertos = .Stats.NPCsMuertos + 1
         
740             EraCriminal = criminal(UserIndex)
         
750             If MiNPC.Stats.Alineacion = 0 Then
         
760                 If MiNPC.Numero = Guardias Then
770                     .Reputacion.NobleRep = 0
780                     .Reputacion.PlebeRep = 0
790                     .Reputacion.AsesinoRep = .Reputacion.AsesinoRep + 500
 
800                     If .Reputacion.AsesinoRep > MAXREP Then .Reputacion.AsesinoRep = MAXREP
 
810                 End If
             
820                 If MiNPC.MaestroUser = 0 Then
830                     .Reputacion.AsesinoRep = .Reputacion.AsesinoRep + vlASESINO
 
840                     If .Reputacion.AsesinoRep > MAXREP Then .Reputacion.AsesinoRep = MAXREP
 
850                 End If
 
860             ElseIf MiNPC.Stats.Alineacion = 1 Then
870                 .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlCAZADOR
 
880                 If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
                 
890             ElseIf MiNPC.Stats.Alineacion = 2 Then
900                 .Reputacion.NobleRep = .Reputacion.NobleRep + vlASESINO / 2
 
910                 If .Reputacion.NobleRep > MAXREP Then .Reputacion.NobleRep = MAXREP
                 
920             ElseIf MiNPC.Stats.Alineacion = 4 Then
930                 .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlCAZADOR
 
940                 If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
                 
950             End If
         
960             If criminal(UserIndex) And esArmada(UserIndex) Then Call ExpulsarFaccionReal(UserIndex)
970             If Not criminal(UserIndex) And esCaos(UserIndex) Then Call ExpulsarFaccionCaos(UserIndex)
         
980             If EraCriminal And Not criminal(UserIndex) Then
990                 Call RefreshCharStatus(UserIndex)
1000            ElseIf Not EraCriminal And criminal(UserIndex) Then
1010                Call RefreshCharStatus(UserIndex)
 
1020            End If
         
1030            Call CheckUserLevel(UserIndex)
         
1040        End With
            
1050        For i = 1 To MAXUSERQUESTS
 
1060        With UserList(UserIndex).QuestStats.Quests(i)
 
1070            If .QuestIndex Then
1080                If QuestList(.QuestIndex).RequiredNPCs Then
 
1090                    For j = 1 To QuestList(.QuestIndex).RequiredNPCs
 
1100                        If QuestList(.QuestIndex).RequiredNPC(j).NpcIndex = MiNPC.Numero Then
1110                            If QuestList(.QuestIndex).RequiredNPC(j).Amount > .NPCsKilled(j) Then
1120                                .NPCsKilled(j) = .NPCsKilled(j) + 1
 
1130                            End If
 
1140                        End If
 
1150                    Next j
 
1160                End If
 
1170            End If
 
1180        End With
 
1190    Next i
 
1200    End If ' Userindex > 0
        
1210    If MiNPC.MaestroUser = 0 Then
            'Tiramos el oro
1220        Call NPCTirarOro(UserIndex, MiNPC)
            'Tiramos el inventario
1230        Call NPC_TIRAR_ITEMS(MiNPC, IsPretoriano)
            'ReSpawn o no
1240        Call ReSpawnNpc(MiNPC)
 
1250    End If
        
1260

        '<EhFooter>
1270    Exit Sub
 
MuereNpc_Err:
1280    LogError Err.Description & vbCrLf & "MuereNpc " & "at line " & Erl
 
        '</EhFooter>
End Sub
Private Sub ResetNpcFlags(ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          'Clear the npc's flags
          
10        With Npclist(NpcIndex).flags
20            .InscribedPrevio = 0
30            .TeamEvent = 0
40            .SlotEvent = 0
50            .Invocacion = 0
60            .AfectaParalisis = 0
70            .AguaValida = 0
80            .AttackedBy = vbNullString
90            .AttackedFirstBy = vbNullString
100           .BackUp = 0
110           .Bendicion = 0
120           .Domable = 0
130           .Envenenado = 0
140           .Faccion = 0
150           .Follow = False
160           .AtacaDoble = 0
170           .LanzaSpells = 0
180           .invisible = 0
190           .Maldicion = 0
200           .OldHostil = 0
210           .OldMovement = 0
220           .Paralizado = 0
230           .Inmovilizado = 0
240           .Respawn = 0
250           .RespawnOrigPos = 0
260           .Snd1 = 0
270           .Snd2 = 0
280           .Snd3 = 0
290           .TierraInvalida = 0
300       End With
End Sub

Private Sub ResetNpcCounters(ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With Npclist(NpcIndex).Contadores
20            .Paralisis = 0
30            .TiempoExistencia = 0
40        End With
End Sub

Private Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With Npclist(NpcIndex).Char
20            .body = 0
30            .CascoAnim = 0
40            .CharIndex = 0
50            .FX = 0
60            .Head = 0
70            .Heading = 0
80            .loops = 0
90            .ShieldAnim = 0
100           .WeaponAnim = 0
110       End With
End Sub

Private Sub ResetNpcCriatures(ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim j As Long
          
10        With Npclist(NpcIndex)
20            For j = 1 To .NroCriaturas
30                .Criaturas(j).NpcIndex = 0
40                .Criaturas(j).NpcName = vbNullString
50            Next j
              
60            .NroCriaturas = 0
70        End With
End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim j As Long
          
10        With Npclist(NpcIndex)
20            For j = 1 To .NroExpresiones
30                .Expresiones(j) = vbNullString
40            Next j
              
50            .NroExpresiones = 0
60        End With
End Sub

Private Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With Npclist(NpcIndex)
20            .Attackable = 0
30            .CanAttack = 0
40            .Comercia = 0
50            .GiveEXP = 0
60            .GiveGLD = 0
70            .Hostile = 0
80            .InvReSpawn = 0
90            .QuestNumber = 0
              
100           If .MaestroUser > 0 Then Call QuitarMascota(.MaestroUser, NpcIndex)
110           If .MaestroNpc > 0 Then Call QuitarMascotaNpc(.MaestroNpc)
120           If .Owner > 0 Then Call PerdioNpc(.Owner)
              
130           .MaestroUser = 0
140           .MaestroNpc = 0
              
150           .Mascotas = 0
160           .Movement = 0
170           .Name = vbNullString
180           .NPCtype = 0
190           .Numero = 0
200           .Orig.map = 0
210           .Orig.X = 0
220           .Orig.Y = 0
230           .PoderAtaque = 0
240           .PoderEvasion = 0
250           .Pos.map = 0
260           .Pos.X = 0
270           .Pos.Y = 0
280           .SkillDomar = 0
290           .Target = 0
300           .TargetNPC = 0
310           .TipoItems = 0
320           .Veneno = 0
330           .desc = vbNullString
              
              
              Dim j As Long
340           For j = 1 To .NroSpells
350               .Spells(j) = 0
360           Next j
370       End With
          
380       Call ResetNpcCharInfo(NpcIndex)
390       Call ResetNpcCriatures(NpcIndex)
400       Call ResetExpresiones(NpcIndex)
End Sub

Public Sub QuitarNPC(ByVal NpcIndex As Integer)
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 16/11/2009
      '16/11/2009: ZaMa - Now npcs lose their owner
      '***************************************************
10    On Error GoTo ErrHandler

20        With Npclist(NpcIndex)
30            .flags.NPCActive = False
              
40            .Owner = 0 ' Murio, no necesita mas dueños :P.
              
50            If InMapBounds(.Pos.map, .Pos.X, .Pos.Y) Then
60                Call EraseNPCChar(NpcIndex)
70            End If
80        End With
              
          'Nos aseguramos de que el inventario sea removido...
          'asi los lobos no volveran a tirar armaduras ;))
90        Call ResetNpcInv(NpcIndex)
100       Call ResetNpcFlags(NpcIndex)
110       Call ResetNpcCounters(NpcIndex)
          
120       Call ResetNpcMainInfo(NpcIndex)
          
130       If NpcIndex = LastNPC Then
140           Do Until Npclist(LastNPC).flags.NPCActive
150               LastNPC = LastNPC - 1
160               If LastNPC < 1 Then Exit Do
170           Loop
180       End If
              
            
190       If NumNPCs <> 0 Then
200           NumNPCs = NumNPCs - 1
210       End If
220   Exit Sub

ErrHandler:
230       Call LogError("Error en QuitarNPC")
End Sub

Public Sub QuitarPet(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
      '***************************************************
      'Autor: ZaMa
      'Last Modification: 18/11/2009
      'Kills a pet
      '***************************************************
10    On Error GoTo ErrHandler

          Dim i As Integer
          Dim PetIndex As Integer

20        With UserList(UserIndex)
              
              ' Busco el indice de la mascota
30            For i = 1 To MAXMASCOTAS
40                If .MascotasIndex(i) = NpcIndex Then PetIndex = i
50            Next i
              
              ' Poco probable que pase, pero por las dudas..
60            If PetIndex = 0 Then Exit Sub
              
              ' Limpio el slot de la mascota
70            .NroMascotas = .NroMascotas - 1
80            .MascotasIndex(PetIndex) = 0
90            .MascotasType(PetIndex) = 0
              
              ' Elimino la mascota
100           Call QuitarNPC(NpcIndex)
110       End With
          
120       Exit Sub

ErrHandler:
130       Call LogError("Error en QuitarPet. Error: " & Err.Number & " Desc: " & Err.Description & " NpcIndex: " & NpcIndex & " UserIndex: " & UserIndex & " PetIndex: " & PetIndex)
End Sub

Private Function TestSpawnTrigger(Pos As WorldPos, Optional PuedeAgua As Boolean = False) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************
          
10        If LegalPos(Pos.map, Pos.X, Pos.Y, PuedeAgua) Then
20            TestSpawnTrigger = _
              MapData(Pos.map, Pos.X, Pos.Y).trigger <> 3 And _
              MapData(Pos.map, Pos.X, Pos.Y).trigger <> 2 And _
              MapData(Pos.map, Pos.X, Pos.Y).trigger <> 1
30        End If
          
End Function

Sub CrearNPC(NroNPC As Integer, Mapa As Integer, OrigPos As WorldPos)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      'Crea un NPC del tipo NRONPC

      Dim Pos As WorldPos
      Dim newpos As WorldPos
      Dim altpos As WorldPos
      Dim nIndex As Integer
      Dim PosicionValida As Boolean
      Dim Iteraciones As Long
      Dim PuedeAgua As Boolean
      Dim PuedeTierra As Boolean


      Dim map As Integer
      Dim X As Integer
      Dim Y As Integer

10        nIndex = OpenNPC(NroNPC) 'Conseguimos un indice
          
20        If nIndex > MAXNPCS Then Exit Sub
30        PuedeAgua = Npclist(nIndex).flags.AguaValida
40        PuedeTierra = IIf(Npclist(nIndex).flags.TierraInvalida = 1, False, True)
          
          'Necesita ser respawned en un lugar especifico
50        If InMapBounds(OrigPos.map, OrigPos.X, OrigPos.Y) Then
              
60            map = OrigPos.map
70            X = OrigPos.X
80            Y = OrigPos.Y
90            Npclist(nIndex).Orig = OrigPos
100           Npclist(nIndex).Pos = OrigPos
             
110       Else
              
120           Pos.map = Mapa 'mapa
130           altpos.map = Mapa
              
140           Do While Not PosicionValida
150               Pos.X = RandomNumber(MinXBorder, MaxXBorder)    'Obtenemos posicion al azar en x
160               Pos.Y = RandomNumber(MinYBorder, MaxYBorder)    'Obtenemos posicion al azar en y
                  
170               Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
180               If newpos.X <> 0 And newpos.Y <> 0 Then
190                   altpos.X = newpos.X
200                   altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn, pero intentando qeu si tenía que ser en el agua, sea en el agua.)
210               Else
220                   Call ClosestLegalPos(Pos, newpos, PuedeAgua)
230                   If newpos.X <> 0 And newpos.Y <> 0 Then
240                       altpos.X = newpos.X
250                       altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn)
260                   End If
270               End If
                  'Si X e Y son iguales a 0 significa que no se encontro posicion valida
280               If LegalPosNPC(newpos.map, newpos.X, newpos.Y, PuedeAgua) And _
                     Not HayPCarea(newpos) And TestSpawnTrigger(newpos, PuedeAgua) Then
                      'Asignamos las nuevas coordenas solo si son validas
290                   Npclist(nIndex).Pos.map = newpos.map
300                   Npclist(nIndex).Pos.X = newpos.X
310                   Npclist(nIndex).Pos.Y = newpos.Y
320                   PosicionValida = True
330               Else
340                   newpos.X = 0
350                   newpos.Y = 0
                  
360               End If
                      
                      
                      
                  'for debug
370               Iteraciones = Iteraciones + 1
380               If Iteraciones > MAXSPAWNATTEMPS Then
390                   If altpos.X <> 0 And altpos.Y <> 0 Then
400                       map = altpos.map
410                       X = altpos.X
420                       Y = altpos.Y
430                       Npclist(nIndex).Pos.map = map
440                       Npclist(nIndex).Pos.X = X
450                       Npclist(nIndex).Pos.Y = Y
460                       Call MakeNPCChar(True, map, nIndex, map, X, Y)
470                       Exit Sub
480                   Else
490                       altpos.X = 50
500                       altpos.Y = 50
510                       Call ClosestLegalPos(altpos, newpos)
520                       If newpos.X <> 0 And newpos.Y <> 0 Then
530                           Npclist(nIndex).Pos.map = newpos.map
540                           Npclist(nIndex).Pos.X = newpos.X
550                           Npclist(nIndex).Pos.Y = newpos.Y
560                           Call MakeNPCChar(True, newpos.map, nIndex, newpos.map, newpos.X, newpos.Y)
570                           Exit Sub
580                       Else
590                           Call QuitarNPC(nIndex)
600                           Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Mapa:" & Mapa & " NroNpc:" & NroNPC)
610                           Exit Sub
620                       End If
630                   End If
640               End If
650           Loop
                  
              'asignamos las nuevas coordenas
660           map = newpos.map
670           X = Npclist(nIndex).Pos.X
680           Y = Npclist(nIndex).Pos.Y
690       End If
                  
          'Crea el NPC
700       Call MakeNPCChar(True, map, nIndex, map, X, Y)
                  
End Sub

Public Sub MakeNPCChar(ByVal toMap As Boolean, sndIndex As Integer, NpcIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

On Error GoTo ErrHandler
          Dim CharIndex As Integer
          Dim shield As Integer
          Dim Helm As Integer
          Dim weapon As Integer
          
10        If Npclist(NpcIndex).Char.CharIndex = 0 Then
20            CharIndex = NextOpenCharIndex
30            Npclist(NpcIndex).Char.CharIndex = CharIndex
40            CharList(CharIndex) = NpcIndex
50        End If
          
60        MapData(map, X, Y).NpcIndex = NpcIndex
          
70        If Npclist(NpcIndex).Numero = 704 Then
80            shield = ObjData(Npclist(NpcIndex).Char.ShieldAnim).ShieldAnim
90            weapon = ObjData(Npclist(NpcIndex).Char.WeaponAnim).WeaponAnim
100           Helm = ObjData(Npclist(NpcIndex).Char.CascoAnim).CascoAnim
110       End If
          
120       If Not toMap Then
130           Call WriteCharacterCreate(sndIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.Head, _
              Npclist(NpcIndex).Char.Heading, Npclist(NpcIndex).Char.CharIndex, X, Y, weapon, shield, 0, 0, Helm, vbNullString, 0, 0)
              'Call WriteCharacterCreate(sndIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.Head, Npclist(NpcIndex).Char.Heading, Npclist(NpcIndex).Char.CharIndex, X, Y, 0, 0, 0, 0, 0, vbNullString, 0, 0)

             
             ' Call WriteCharacterCreate(sndIndex, Npclist(NpcIndex).Char.body, Npclist(NpcIndex).Char.Head, Npclist(NpcIndex).Char.Heading,
             'Npclist(NpcIndex).Chanim), IIf((Npclist(NpcIndex).Char.(NpcIndex).Char.CascoAnim = 0), 0, ObjData(Npclist(NpcIndex).Char.CascoAnim).CascoAnim)), vbNullString, 0, 0)
140           Call FlushBuffer(sndIndex)
150       Else
160           Call AgregarNpc(NpcIndex)
170       End If

Exit Sub
ErrHandler:
LogError "Error en linea " & Erl & " Procedimiento MakeNpcChar"
End Sub

Public Sub ChangeNPCChar(ByVal NpcIndex As Integer, ByVal body As Integer, ByVal Head As Integer, ByVal Heading As eHeading)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If NpcIndex > 0 Then
20            With Npclist(NpcIndex).Char
30                .body = body
40                .Head = Head
50                .Heading = Heading
                  
60                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterChange(body, Head, Heading, .CharIndex, 0, 0, 0, 0, 0))
70            End With
80        End If
End Sub

Private Sub EraseNPCChar(ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    If Npclist(NpcIndex).Char.CharIndex <> 0 Then CharList(Npclist(NpcIndex).Char.CharIndex) = 0

20    If Npclist(NpcIndex).Char.CharIndex = LastChar Then
30        Do Until CharList(LastChar) > 0
40            LastChar = LastChar - 1
50            If LastChar <= 1 Then Exit Do
60        Loop
70    End If

      'Quitamos del mapa
80    MapData(Npclist(NpcIndex).Pos.map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0

      'Actualizamos los clientes
90    Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterRemove(Npclist(NpcIndex).Char.CharIndex))

      'Update la lista npc
100   Npclist(NpcIndex).Char.CharIndex = 0


      'update NumChars
110   NumChars = NumChars - 1


End Sub

Public Sub MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte)
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 06/04/2009
      '06/04/2009: ZaMa - Now npcs can force to change position with dead character
      '01/08/2009: ZaMa - Now npcs can't force to chance position with a dead character if that means to change the terrain the character is in
      '***************************************************

10    On Error GoTo errh
          Dim nPos As WorldPos
          Dim UserIndex As Integer
          
20        With Npclist(NpcIndex)
30            nPos = .Pos
40            Call HeadtoPos(nHeading, nPos)
              
              ' es una posicion legal
50            If LegalPosNPC(.Pos.map, nPos.X, nPos.Y, .flags.AguaValida = 1, .MaestroUser <> 0) Then
                  
60                If .flags.AguaValida = 0 And HayAgua(.Pos.map, nPos.X, nPos.Y) Then Exit Sub
70                If .flags.TierraInvalida = 1 And Not HayAgua(.Pos.map, nPos.X, nPos.Y) Then Exit Sub
                  
80                UserIndex = MapData(.Pos.map, nPos.X, nPos.Y).UserIndex
                  ' Si hay un usuario a donde se mueve el npc, entonces esta muerto
90                If UserIndex > 0 Then
                      
                      ' No se traslada caspers de agua a tierra
100                   If HayAgua(.Pos.map, nPos.X, nPos.Y) And Not HayAgua(.Pos.map, .Pos.X, .Pos.Y) Then Exit Sub
                      ' No se traslada caspers de tierra a agua
110                   If Not HayAgua(.Pos.map, nPos.X, nPos.Y) And HayAgua(.Pos.map, .Pos.X, .Pos.Y) Then Exit Sub
                      
120                   With UserList(UserIndex)
                          ' Actualizamos posicion y mapa
130                       MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex = 0
140                       .Pos.X = Npclist(NpcIndex).Pos.X
150                       .Pos.Y = Npclist(NpcIndex).Pos.Y
160                       MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex = UserIndex
                              
                          ' Avisamos a los usuarios del area, y al propio usuario lo forzamos a moverse
170                       Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, .Pos.X, .Pos.Y))
180                       Call WriteForceCharMove(UserIndex, InvertHeading(nHeading))
190                   End With
200               End If
                  
210               Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.X, nPos.Y))

                  'Update map and user pos
220               MapData(.Pos.map, .Pos.X, .Pos.Y).NpcIndex = 0
230               .Pos = nPos
240               .Char.Heading = nHeading
250               MapData(.Pos.map, nPos.X, nPos.Y).NpcIndex = NpcIndex
260               Call CheckUpdateNeededNpc(NpcIndex, nHeading)
              
270           ElseIf .MaestroUser = 0 Then
280               If .Movement = TipoAI.NpcPathfinding Then
                      'Someone has blocked the npc's way, we must to seek a new path!
290                   .PFINFO.PathLenght = 0
300               End If
310           End If
320       End With
330   Exit Sub

errh:
340       LogError ("Error en move npc " & NpcIndex)
End Sub

Function NextOpenNPC() As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10    On Error GoTo ErrHandler
          Dim LoopC As Long
            
20        For LoopC = 1 To MAXNPCS + 1
30            If LoopC > MAXNPCS Then Exit For
40            If Not Npclist(LoopC).flags.NPCActive Then Exit For
50        Next LoopC
            
60        NextOpenNPC = LoopC
70    Exit Function

ErrHandler:
80        Call LogError("Error en NextOpenNPC")
End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 10/07/2010
      '10/07/2010: ZaMa - Now npcs can't poison dead users.
      '***************************************************
       
          Dim N As Integer
         
10        With UserList(UserIndex)
20            If .flags.Muerto = 1 Then Exit Sub
             
30            N = RandomNumber(1, 100)
40            If N < 30 Then
50                .flags.Envenenado = 1
60                Call WriteConsoleMsg(UserIndex, "¡¡La criatura te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
70            End If
80        End With
         
End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, Pos As WorldPos, ByVal FX As Boolean, ByVal Respawn As Boolean) As Integer
      '***************************************************
      'Autor: Unknown (orginal version)
      'Last Modification: 06/15/2008
      '23/01/2007 -> Pablo (ToxicWaste): Creates an NPC of the type Npcindex
      '06/15/2008 -> Optimizé el codigo. (NicoNZ)
      '***************************************************
      Dim newpos As WorldPos
      Dim altpos As WorldPos
      Dim nIndex As Integer
      Dim PosicionValida As Boolean
      Dim PuedeAgua As Boolean
      Dim PuedeTierra As Boolean


      Dim map As Integer
      Dim X As Integer
      Dim Y As Integer

10    nIndex = OpenNPC(NpcIndex, Respawn)   'Conseguimos un indice

20    If nIndex > MAXNPCS Then
30        SpawnNpc = 0
40        Exit Function
50    End If

60    PuedeAgua = Npclist(nIndex).flags.AguaValida
70    PuedeTierra = Not Npclist(nIndex).flags.TierraInvalida = 1
              
80    Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
90    Call ClosestLegalPos(Pos, altpos, PuedeAgua)
      'Si X e Y son iguales a 0 significa que no se encontro posicion valida

100   If newpos.X <> 0 And newpos.Y <> 0 Then
          'Asignamos las nuevas coordenas solo si son validas
110       Npclist(nIndex).Pos.map = newpos.map
120       Npclist(nIndex).Pos.X = newpos.X
130       Npclist(nIndex).Pos.Y = newpos.Y
140       PosicionValida = True
150   Else
160       If altpos.X <> 0 And altpos.Y <> 0 Then
170           Npclist(nIndex).Pos.map = altpos.map
180           Npclist(nIndex).Pos.X = altpos.X
190           Npclist(nIndex).Pos.Y = altpos.Y
200           PosicionValida = True
210       Else
220           PosicionValida = False
230       End If
240   End If

250   If Not PosicionValida Then
260       Call QuitarNPC(nIndex)
270       SpawnNpc = 0
280       Exit Function
290   End If

      'asignamos las nuevas coordenas
300   map = newpos.map
310   X = Npclist(nIndex).Pos.X
320   Y = Npclist(nIndex).Pos.Y

      'Crea el NPC
330   Call MakeNPCChar(True, map, nIndex, map, X, Y)

340   If FX Then
350       Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
360       Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessageCreateFX(Npclist(nIndex).Char.CharIndex, FXIDs.FXWARP, 0))
370   End If

380   SpawnNpc = nIndex

End Function

Sub ReSpawnNpc(MiNPC As Npc)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.Numero, MiNPC.Pos.map, MiNPC.Orig)

End Sub

Private Sub NPCTirarOro(ByVal UserIndex As Integer, ByRef MiNPC As Npc)
          Dim MiAux As Long
          
10        If MiNPC.GiveGLD > 0 Then
              If UserList(UserIndex).GroupIndex > 0 Then
                    If Groups(UserList(UserIndex).GroupIndex).Members > 1 Then
                        mGroup.AddGldGroup UserList(UserIndex).GroupIndex, MiNPC.GiveGLD * Oroc
                        Exit Sub
                    End If
              End If
            
20            MiAux = MiNPC.GiveGLD * Oroc
                  Dim MiObj As Obj

40                Do While MiAux > MAX_INVENTORY_OBJS
50                    MiObj.Amount = MAX_INVENTORY_OBJS
60                    MiObj.ObjIndex = iORO
70                    Call TirarItemAlPiso(MiNPC.Pos, MiObj)
80                    MiAux = MiAux - MAX_INVENTORY_OBJS
90                Loop
100               If MiAux > 0 Then
110                   MiObj.Amount = MiAux
120                   MiObj.ObjIndex = iORO
130                   Call TirarItemAlPiso(MiNPC.Pos, MiObj)
140               End If
150
220       End If
End Sub
Public Function OpenNPC(ByVal NpcNumber As Integer, Optional ByVal Respawn = True) As Integer
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      '###################################################
      '#               ATENCION PELIGRO                  #
      '###################################################
      '
      '    ¡¡¡¡ NO USAR GetVar PARA LEER LOS NPCS !!!!
      '
      'El que ose desafiar esta LEY, se las tendrá que ver
      'conmigo. Para leer los NPCS se deberá usar la
      'nueva clase clsinimanager.
      '
      'Alejo
      '
      '###################################################
          Dim NpcIndex As Integer
          Dim Leer As clsIniManager
          Dim LoopC As Long
          Dim ln As String
          Dim aux As String
          
10        Set Leer = LeerNPCs
          
          'If requested index is invalid, abort
20        If Not Leer.KeyExists("NPC" & NpcNumber) Then
30            OpenNPC = MAXNPCS + 1
40            Exit Function
50        End If
          
60        NpcIndex = NextOpenNPC
          
70        If NpcIndex > MAXNPCS Then 'Limite de npcs
80            OpenNPC = NpcIndex
90            Exit Function
100       End If
          
110       With Npclist(NpcIndex)
          
120           .Numero = NpcNumber
130           .Name = Leer.GetValue("NPC" & NpcNumber, "Name")
140           .desc = Leer.GetValue("NPC" & NpcNumber, "Desc")
              
150           .Movement = val(Leer.GetValue("NPC" & NpcNumber, "Movement"))
160           .flags.OldMovement = .Movement
              
170           .flags.AguaValida = val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
180           .flags.TierraInvalida = val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
190           .flags.Faccion = val(Leer.GetValue("NPC" & NpcNumber, "Faccion"))
200           .flags.AtacaDoble = val(Leer.GetValue("NPC" & NpcNumber, "AtacaDoble"))
              
210           .NPCtype = val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))
              
220           .Char.body = val(Leer.GetValue("NPC" & NpcNumber, "Body"))
230           .Char.Head = val(Leer.GetValue("NPC" & NpcNumber, "Head"))
240           .Char.ShieldAnim = val(Leer.GetValue("NPC" & NpcNumber, "EscudoAnim"))
250           .Char.WeaponAnim = val(Leer.GetValue("NPC" & NpcNumber, "ArmaAnim"))
260           .Char.CascoAnim = val(Leer.GetValue("NPC" & NpcNumber, "CascoAnim"))
270           .Char.Heading = val(Leer.GetValue("NPC" & NpcNumber, "Heading"))
              
280           .Attackable = val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
290           .Comercia = val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
300           .Hostile = val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
310           .flags.OldHostil = .Hostile
              
320           .GiveEXP = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP")) * Expc
              
330           .flags.ExpCount = .GiveEXP
              
340           .Veneno = val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))
              
350           .flags.Domable = val(Leer.GetValue("NPC" & NpcNumber, "Domable"))
              
360           .GiveGLD = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLD"))
              
370           .QuestNumber = val(Leer.GetValue("NPC" & NpcNumber, "QuestNumber"))
              
380           .PoderAtaque = val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
390           .PoderEvasion = val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))
              
400           .InvReSpawn = val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))
              
410           With .Stats
420               .MaxHp = val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
430               .MinHp = val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
440               .MaxHIT = val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
450               .MinHIT = val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
460               .def = val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
470               .defM = val(Leer.GetValue("NPC" & NpcNumber, "DEFm"))
480               .Alineacion = val(Leer.GetValue("NPC" & NpcNumber, "Alineacion"))
490           End With
              
500           .Invent.NroItems = val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))
510           For LoopC = 1 To .Invent.NroItems
520               ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & LoopC)
530               .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
540               .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
550           Next LoopC
              
560           For LoopC = 1 To MAX_NPC_DROPS
570               ln = Leer.GetValue("NPC" & NpcNumber, "Drop" & LoopC)
580               .Drop(LoopC).ObjIndex = val(ReadField(1, ln, 45))
590               .Drop(LoopC).Amount = val(ReadField(2, ln, 45))
600               .Drop(LoopC).Probability = val(ReadField(3, ln, 45))
610           Next LoopC

              
620           .flags.LanzaSpells = val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))
630           If .flags.LanzaSpells > 0 Then ReDim .Spells(1 To .flags.LanzaSpells)
640           For LoopC = 1 To .flags.LanzaSpells
650               .Spells(LoopC) = val(Leer.GetValue("NPC" & NpcNumber, "Sp" & LoopC))
660           Next LoopC
              
670           If .NPCtype = eNPCType.Entrenador Then
680               .NroCriaturas = val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))
690               ReDim .Criaturas(1 To .NroCriaturas) As tCriaturasEntrenador
700               For LoopC = 1 To .NroCriaturas
710                   .Criaturas(LoopC).NpcIndex = Leer.GetValue("NPC" & NpcNumber, "CI" & LoopC)
720                   .Criaturas(LoopC).NpcName = Leer.GetValue("NPC" & NpcNumber, "CN" & LoopC)
730               Next LoopC
740           End If
              
750           With .flags
760               .NPCActive = True
                  
770               If Respawn Then
780                   .Respawn = val(Leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
790               Else
800                   .Respawn = 1
810               End If
                  
820               .BackUp = val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
830               .RespawnOrigPos = val(Leer.GetValue("NPC" & NpcNumber, "OrigPos"))
840               .AfectaParalisis = val(Leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
                  
850               .Snd1 = val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
860               .Snd2 = val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
870               .Snd3 = val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))
880           End With
              
              '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
890           .NroExpresiones = val(Leer.GetValue("NPC" & NpcNumber, "NROEXP"))
900           If .NroExpresiones > 0 Then ReDim .Expresiones(1 To .NroExpresiones) As String
910           For LoopC = 1 To .NroExpresiones
920               .Expresiones(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Exp" & LoopC)
930           Next LoopC
              '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
              
              'Tipo de items con los que comercia
940           .TipoItems = val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))
              
950           .Ciudad = val(Leer.GetValue("NPC" & NpcNumber, "Ciudad"))
960       End With
          
          'Update contadores de NPCs
970       If NpcIndex > LastNPC Then LastNPC = NpcIndex
980       NumNPCs = NumNPCs + 1
          
          'Devuelve el nuevo Indice
990       OpenNPC = NpcIndex
End Function

Public Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With Npclist(NpcIndex)
20            If .flags.Follow Then
30                .flags.AttackedBy = vbNullString
40                .flags.Follow = False
50                .Movement = .flags.OldMovement
60                .Hostile = .flags.OldHostil
70            Else
80                .flags.AttackedBy = UserName
90                .flags.Follow = True
100               .Movement = TipoAI.NPCDEFENSA
110               .Hostile = 0
120           End If
130       End With
End Sub

Public Sub FollowAmo(ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        With Npclist(NpcIndex)
20            .flags.Follow = True
30            .Movement = TipoAI.SigueAmo
40            .Hostile = 0
50            .Target = 0
60            .TargetNPC = 0
70        End With
End Sub

Public Sub ValidarPermanenciaNpc(ByVal NpcIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      'Chequea si el npc continua perteneciendo a algún usuario
      '***************************************************

10        With Npclist(NpcIndex)
20            If IntervaloPerdioNpc(.Owner) Then Call PerdioNpc(.Owner)
30        End With
End Sub
