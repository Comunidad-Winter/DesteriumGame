Attribute VB_Name = "SistemaCombate"

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
'
'Diseño y corrección del modulo de combate por
'Gerardo Saiz, gerardosaiz@yahoo.com
'

'9/01/2008 Pablo (ToxicWaste) - Ahora TODOS los modificadores de Clase se controlan desde Balance.dat


Option Explicit

Public Const MAXDISTANCIAARCO As Byte = 18
Public Const MAXDISTANCIAMAGIA As Byte = 18

Public Function MinimoInt(ByVal A As Integer, ByVal B As Integer) As Integer
10        If A > B Then
20            MinimoInt = B
30        Else
40            MinimoInt = A
50        End If
End Function

Public Function MaximoInt(ByVal A As Integer, ByVal B As Integer) As Integer
10        If A > B Then
20            MaximoInt = A
30        Else
40            MaximoInt = B
50        End If
End Function

Private Function PoderEvasionEscudo(ByVal userindex As Integer) As Long
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

10        PoderEvasionEscudo = (UserList(userindex).Stats.UserSkills(eSkill.Defensa) * ModClase(UserList(userindex).clase).Escudo) / 2
End Function

Private Function PoderEvasion(ByVal userindex As Integer) As Long
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************
          Dim lTemp As Long
10        With UserList(userindex)
20            lTemp = (.Stats.UserSkills(eSkill.Tacticas) + _
                .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).Evasion
             
30            PoderEvasion = (lTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
40        End With
End Function

Private Function PoderAtaqueArma(ByVal userindex As Integer) As Long
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim PoderAtaqueTemp As Long
          
10        With UserList(userindex)
20            If .Stats.UserSkills(eSkill.Armas) < 31 Then
30                PoderAtaqueTemp = .Stats.UserSkills(eSkill.Armas) * ModClase(.clase).AtaqueArmas
40            ElseIf .Stats.UserSkills(eSkill.Armas) < 61 Then
50                PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueArmas
60            ElseIf .Stats.UserSkills(eSkill.Armas) < 91 Then
70                PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueArmas
80            Else
90               PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueArmas
100           End If
              
110           PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
120       End With
End Function

Private Function PoderAtaqueProyectil(ByVal userindex As Integer) As Long
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim PoderAtaqueTemp As Long
          
10        With UserList(userindex)
20            If .Stats.UserSkills(eSkill.Proyectiles) < 31 Then
30                PoderAtaqueTemp = .Stats.UserSkills(eSkill.Proyectiles) * ModClase(.clase).AtaqueProyectiles
40            ElseIf .Stats.UserSkills(eSkill.Proyectiles) < 61 Then
50                PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueProyectiles
60            ElseIf .Stats.UserSkills(eSkill.Proyectiles) < 91 Then
70                PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueProyectiles
80            Else
90                PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Proyectiles) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueProyectiles
100           End If
              
110           PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
120       End With
End Function

Private Function PoderAtaqueWrestling(ByVal userindex As Integer) As Long
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim PoderAtaqueTemp As Long
          
10        With UserList(userindex)
20            If .Stats.UserSkills(eSkill.Wrestling) < 31 Then
30                PoderAtaqueTemp = .Stats.UserSkills(eSkill.Wrestling) * ModClase(.clase).AtaqueWrestling
40            ElseIf .Stats.UserSkills(eSkill.Wrestling) < 61 Then
50                PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueWrestling
60            ElseIf .Stats.UserSkills(eSkill.Wrestling) < 91 Then
70                PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueWrestling
80            Else
90                PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Wrestling) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * ModClase(.clase).AtaqueWrestling
100           End If
              
110           PoderAtaqueWrestling = (PoderAtaqueTemp + (2.5 * MaximoInt(.Stats.ELV - 12, 0)))
120       End With
End Function

Public Function UserImpactoNpc(ByVal userindex As Integer, ByVal NpcIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim PoderAtaque As Long
          Dim Arma As Integer
          Dim Skill As eSkill
          Dim ProbExito As Long
          
10        Arma = UserList(userindex).Invent.WeaponEqpObjIndex
          
20        If Arma > 0 Then 'Usando un arma
30            If ObjData(Arma).proyectil = 1 Then
40                PoderAtaque = PoderAtaqueProyectil(userindex)
50                Skill = eSkill.Proyectiles
60            Else
70                PoderAtaque = PoderAtaqueArma(userindex)
80                Skill = eSkill.Armas
90            End If
100       Else 'Peleando con puños
110           PoderAtaque = PoderAtaqueWrestling(userindex)
120           Skill = eSkill.Wrestling
130       End If
          
          ' Chances are rounded
140       ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))
          
150       UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
          
160       If UserImpactoNpc Then
170           Call SubirSkill(userindex, Skill, True)
180       Else
190           Call SubirSkill(userindex, Skill, False)
200       End If
End Function
Public Function NpcImpacto(ByVal NpcIndex As Integer, ByVal userindex As Integer) As Boolean
10        On Error GoTo NpcImpacto_Error
      '*************************************************
      'Author: Unknown
      'Last modified: 03/15/2006
      'Revisa si un NPC logra impactar a un user o no
      '03/15/2006 Maraxus - Evité una división por cero que eliminaba NPCs
      '*************************************************
          Dim Rechazo As Boolean
          Dim ProbRechazo As Long
          Dim ProbExito As Long
          Dim UserEvasion As Long
          Dim NpcPoderAtaque As Long
          Dim PoderEvasioEscudo As Long
          Dim SkillTacticas As Long
          Dim SkillDefensa As Long
          
'Agregué un Manejador de error que me dice _
en qué linea va a tirar el Error.

20        UserEvasion = PoderEvasion(userindex)
30        NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
40        PoderEvasioEscudo = PoderEvasionEscudo(userindex)
          
50        SkillTacticas = UserList(userindex).Stats.UserSkills(eSkill.Tacticas)
60        SkillDefensa = UserList(userindex).Stats.UserSkills(eSkill.Defensa)
          
          'Esta usando un escudo ???
70        If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then UserEvasion = UserEvasion + PoderEvasioEscudo
          
          ' Chances are rounded
80        ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))
          
90        NpcImpacto = (RandomNumber(1, 100) <= ProbExito)
          
          ' el usuario esta usando un escudo ???
100       If UserList(userindex).Invent.EscudoEqpObjIndex > 0 Then
110           If Not NpcImpacto Then
120               If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
                      ' Chances are rounded
130                   ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
140                   Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
                      
150                   If Rechazo Then
                          'Se rechazo el ataque con el escudo
160                       Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_ESCUDO, UserList(userindex).Pos.X, UserList(userindex).Pos.Y))
170                       Call WriteMultiMessage(userindex, eMessages.BlockedWithShieldUser) 'Call WriteBlockedWithShieldUser(UserIndex)
180                       SendData SendTarget.ToPCArea, userindex, PrepareMessageMovimientSW(UserList(userindex).Char.CharIndex, 2)
190                       Call SubirSkill(userindex, eSkill.Defensa, True)
200                   Else
210                       Call SubirSkill(userindex, eSkill.Defensa, False)
220                   End If
230               End If
240           End If
250       End If
    
260       On Error GoTo 0
270       Exit Function

NpcImpacto_Error:

280        LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure NpcImpacto, line " & Erl & "."
    
End Function




Public Function CalcularDaño(ByVal userindex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long
      '***************************************************
      'Author: Unknown
      'Last Modification: 01/04/2010 (ZaMa)
      '01/04/2010: ZaMa - Modifico el daño de wrestling.
      '01/04/2010: ZaMa - Agrego bonificadores de wrestling para los guantes.
      '***************************************************
          Dim DañoArma As Long
          Dim DañoUsuario As Long
          Dim Arma As ObjData
          Dim ModifClase As Single
          Dim proyectil As ObjData
          Dim DañoMaxArma As Long
          Dim DañoMinArma As Long
          Dim ObjIndex As Integer
          
          ''sacar esto si no queremos q la matadracos mate el Dragon si o si
          Dim matoDragon As Boolean
10        matoDragon = False
          
20        With UserList(userindex)
30            If .Invent.WeaponEqpObjIndex > 0 Then
40                Arma = ObjData(.Invent.WeaponEqpObjIndex)
                  
                  ' Ataca a un npc?
50                If NpcIndex > 0 Then
60                    If Arma.proyectil = 1 Then
70                        ModifClase = ModClase(.clase).DañoProyectiles
80                        DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
90                        DañoMaxArma = Arma.MaxHIT
                          
100                       If Arma.Municion = 1 Then
110                           proyectil = ObjData(.Invent.MunicionEqpObjIndex)
120                           DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                              ' For some reason this isn't done...
                              'DañoMaxArma = DañoMaxArma + proyectil.MaxHIT
130                       End If
140                   Else
150                       ModifClase = ModClase(.clase).DañoArmas
                          
160                       If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then ' Usa la mata Dragones?
170                           If Npclist(NpcIndex).NPCtype = Dragon Then 'Ataca Dragon?
180                               DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
190                               DañoMaxArma = Arma.MaxHIT
200                                matoDragon = True
210                           Else ' Sino es Dragon daño es 1
220                               DañoArma = 1
230                               DañoMaxArma = 1
240                           End If
250                       Else
260                           DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
270                           DañoMaxArma = Arma.MaxHIT
280                       End If
290                   End If
300               Else ' Ataca usuario
310                   If Arma.proyectil = 1 Then
320                       ModifClase = ModClase(.clase).DañoProyectiles
330                       DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
340                       DañoMaxArma = Arma.MaxHIT
                           
350                       If Arma.Municion = 1 Then
360                           proyectil = ObjData(.Invent.MunicionEqpObjIndex)
370                           DañoArma = DañoArma + RandomNumber(proyectil.MinHIT, proyectil.MaxHIT)
                              ' For some reason this isn't done...
                              'DañoMaxArma = DañoMaxArma + proyectil.MaxHIT
380                       End If
390                   Else
400                       ModifClase = ModClase(.clase).DañoArmas
                          
410                       If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
420                           ModifClase = ModClase(.clase).DañoArmas
430                           DañoArma = 1 ' Si usa la espada mataDragones daño es 1
440                           DañoMaxArma = 1
450                       Else
460                           DañoArma = RandomNumber(Arma.MinHIT, Arma.MaxHIT)
470                           DañoMaxArma = Arma.MaxHIT
480                       End If
490                   End If
500               End If
510           Else
520               ModifClase = ModClase(.clase).DañoWrestling
                  
                  ' Daño sin guantes
530               DañoMinArma = 4
540               DañoMaxArma = 9
                  
                  ' Plus de guantes (en slot de anillo)
550               ObjIndex = .Invent.AnilloEqpObjIndex
560               If ObjIndex > 0 Then
570                   If ObjData(ObjIndex).Guante = 1 Then
580                       DañoMinArma = DañoMinArma + ObjData(ObjIndex).MinHIT
590                       DañoMaxArma = DañoMaxArma + ObjData(ObjIndex).MaxHIT
600                   End If
610               End If
                  
620               DañoArma = RandomNumber(DañoMinArma, DañoMaxArma)
                  
630           End If
              
640           DañoUsuario = RandomNumber(.Stats.MinHIT, .Stats.MaxHIT)
              
              ''sacar esto si no queremos q la matadracos mate el Dragon si o si
650           If matoDragon Then
660               CalcularDaño = Npclist(NpcIndex).Stats.MinHp + Npclist(NpcIndex).Stats.def
670           Else
680               CalcularDaño = (3 * DañoArma + ((DañoMaxArma / 5) * MaximoInt(0, .Stats.UserAtributos(eAtributos.Fuerza) - 15)) + DañoUsuario) * ModifClase
690           End If
700       End With
End Function

Public Sub UserDañoNpc(ByVal userindex As Integer, ByVal NpcIndex As Integer)
10        On Error GoTo UserDañoNpc_Error
      '***************************************************
      'Author: Unknown
      'Last Modification: 07/04/2010 (ZaMa)
      '25/01/2010: ZaMa - Agrego poder acuchillar npcs.
      '07/04/2010: ZaMa - Los asesinos apuñalan acorde al daño base sin descontar la defensa del npc.
      '***************************************************

          Dim daño As Long
          Dim DañoBase As Long
          
20        DañoBase = CalcularDaño(userindex, NpcIndex)
          
          'esta navegando? si es asi le sumamos el daño del barco
30        If UserList(userindex).flags.Navegando = 1 Then
40            If UserList(userindex).Invent.BarcoObjIndex > 0 Then
50                DañoBase = DañoBase + RandomNumber(ObjData(UserList(userindex).Invent.BarcoObjIndex).MinHIT, _
                                              ObjData(UserList(userindex).Invent.BarcoObjIndex).MaxHIT)
60            End If
70        End If
          
          
80        With Npclist(NpcIndex)
90            daño = DañoBase - .Stats.def
            
130         If UserPoderoso(userindex) = 263 Then
140             daño = daño * 1.35
150         ElseIf UserPoderoso(userindex) = 262 Then
160             daño = daño * 1.25
170         End If
        
180           If daño < 0 Then daño = 0
              
              'Call WriteUserHitNPC(UserIndex, daño)
190           Call WriteMultiMessage(userindex, eMessages.UserHitNPC, daño)
200           Call CalcularDarExp(userindex, NpcIndex, daño)
210           .Stats.MinHp = .Stats.MinHp - daño
              
220           SendData SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, DAMAGE_NORMAL)
                     
230           If .Stats.MinHp > 0 Then
                  'Trata de apuñalar por la espalda al enemigo
240               If UserList(userindex).clase = eClass.Hunter Then
250               DoGolpeArco userindex, NpcIndex, 0, daño
260               End If
                  
270               If PuedeApuñalar(userindex) Then
280               UserList(userindex).Dañoapu = daño
290                  Call DoApuñalar(userindex, NpcIndex, 0, DañoBase)
300               End If
                  
                  'trata de dar golpe crítico
310               Call DoGolpeCritico(userindex, NpcIndex, 0, daño)
                  
320               If PuedeAcuchillar(userindex) Then
330                   Call DoAcuchillar(userindex, NpcIndex, 0, daño)
340               End If
350           End If
              
              
360           If .Stats.MinHp <= 0 Then
                  ' Si era un Dragon perdemos la espada mataDragones
370               If .NPCtype = Dragon Then
                      'Si tiene equipada la matadracos se la sacamos
380                   If UserList(userindex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
390                       Call QuitarObjetos(EspadaMataDragonesIndex, 1, userindex)
400                   End If
410                   If .Stats.MaxHp > 100000 Then Call LogDesarrollo(UserList(userindex).Name & " mató un dragón")
420               End If
                  
                  ' Para que las mascotas no sigan intentando luchar y
                  ' comiencen a seguir al amo
                  Dim j As Integer
430               For j = 1 To MAXMASCOTAS
440                   If UserList(userindex).MascotasIndex(j) > 0 Then
450                       If Npclist(UserList(userindex).MascotasIndex(j)).TargetNPC = NpcIndex Then
460                           Npclist(UserList(userindex).MascotasIndex(j)).TargetNPC = 0
470                           Npclist(UserList(userindex).MascotasIndex(j)).Movement = TipoAI.SigueAmo
480                       End If
490                   End If
500               Next j
                  
510               Call MuereNpc(NpcIndex, userindex)
520           End If
530       End With
    
540       On Error GoTo 0
550       Exit Sub

UserDañoNpc_Error:

560       LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure UserDañoNpc, line " & Erl & "."
    
End Sub



Public Sub NpcDaño(ByVal NpcIndex As Integer, ByVal userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************
       
          Dim daño As Integer
          Dim Lugar As Integer
          Dim absorbido As Integer
          Dim defbarco As Integer
          Dim defmontura As Integer
          Dim Obj As ObjData
         
10        daño = RandomNumber(Npclist(NpcIndex).Stats.MinHIT, Npclist(NpcIndex).Stats.MaxHIT)
          
20        With UserList(userindex)
30            If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
40                Obj = ObjData(.Invent.BarcoObjIndex)
50                defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
60            End If
             
70            If .flags.Montando = 1 Then
80                Obj = ObjData(.Invent.MonturaObjIndex)
90                defmontura = RandomNumber(Obj.MinDef, Obj.MaxDef)
100           End If
             
110           Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
             
120           Select Case Lugar
                  Case PartesCuerpo.bCabeza
                      'Si tiene casco absorbe el golpe
130                   If .Invent.CascoEqpObjIndex > 0 Then
140                      Obj = ObjData(.Invent.CascoEqpObjIndex)
150                      absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
160                   End If
170             Case Else
                      'Si tiene armadura absorbe el golpe
180                   If .Invent.ArmourEqpObjIndex > 0 Then
                          Dim Obj2 As ObjData
190                       Obj = ObjData(.Invent.ArmourEqpObjIndex)
200                       If .Invent.EscudoEqpObjIndex Then
210                           Obj2 = ObjData(.Invent.EscudoEqpObjIndex)
220                           absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
230                       Else
240                           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
250                      End If
260                   End If
270           End Select
             
280           absorbido = absorbido + defbarco + defmontura
290           daño = daño - absorbido
300           If daño < 1 Then daño = 1
             
310           Call WriteMultiMessage(userindex, eMessages.NPCHitUser, Lugar, daño)
              'Call WriteNPCHitUser(UserIndex, Lugar, daño)
             
320           If .flags.Privilegios And PlayerType.User Then .Stats.MinHp = .Stats.MinHp - daño
             
330           If .flags.Meditando Then
340               If daño > Fix(.Stats.MinHp / 100 * .Stats.UserAtributos(eAtributos.Inteligencia) * .Stats.UserSkills(eSkill.Meditar) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
350                   .flags.Meditando = False
360                   Call WriteMeditateToggle(userindex)
370                   Call WriteConsoleMsg(userindex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
380                   .Char.FX = 0
390                   .Char.loops = 0
400                   Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
410               End If
420           End If
             
              'Muere el usuario
430           If .Stats.MinHp <= 0 Then
440               Call WriteMultiMessage(userindex, eMessages.NPCKillUser) 'Call WriteNPCKillUser(UserIndex) ' Le informamos que ha muerto
                 
                  'Si lo mato un guardia
450               If criminal(userindex) And Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
460                   Call RestarCriminalidad(userindex)
470                   If Not criminal(userindex) And .Faccion.FuerzasCaos = 1 Then Call ExpulsarFaccionCaos(userindex)
480               End If
                 
490               If Npclist(NpcIndex).MaestroUser > 0 Then
500                   Call AllFollowAmo(Npclist(NpcIndex).MaestroUser)
510               Else
                      'Al matarlo no lo sigue mas
520                   If Npclist(NpcIndex).Stats.Alineacion = 0 Then
530                       Npclist(NpcIndex).Movement = Npclist(NpcIndex).flags.OldMovement
540                       Npclist(NpcIndex).Hostile = Npclist(NpcIndex).flags.OldHostil
550                       Npclist(NpcIndex).flags.AttackedBy = vbNullString
560                   End If
570               End If
                 
580               Call UserDie(userindex)
590           End If
600       End With
End Sub


Public Sub RestarCriminalidad(ByVal userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim EraCriminal As Boolean
10        EraCriminal = criminal(userindex)
          
20        With UserList(userindex).Reputacion
30            If .BandidoRep > 0 Then
40                 .BandidoRep = .BandidoRep - vlASALTO
50                 If .BandidoRep < 0 Then .BandidoRep = 0
60            ElseIf .LadronesRep > 0 Then
70                 .LadronesRep = .LadronesRep - (vlCAZADOR * 10)
80                 If .LadronesRep < 0 Then .LadronesRep = 0
90            End If
100       End With
          
110       If EraCriminal And Not criminal(userindex) Then
120           Call RefreshCharStatus(userindex)
130       End If
End Sub

Public Sub CheckPets(ByVal NpcIndex As Integer, ByVal userindex As Integer, Optional ByVal CheckElementales As Boolean = True)
      '***************************************************
      'Author: Unknown
      'Last Modification: 15/04/2010
      '15/04/2010: ZaMa - Las mascotas no se apropian de npcs.
      '***************************************************

          Dim j As Integer
          
          ' Si no tengo mascotas, para que cheaquear lo demas?
10        If UserList(userindex).NroMascotas = 0 Then Exit Sub
          
20        If Not PuedeAtacarNPC(userindex, NpcIndex, , True) Then Exit Sub
          
30        With UserList(userindex)
40            For j = 1 To MAXMASCOTAS
50                If .MascotasIndex(j) > 0 Then
60                   If .MascotasIndex(j) <> NpcIndex Then
70                    If CheckElementales Or (Npclist(.MascotasIndex(j)).Numero <> ELEMENTALFUEGO And Npclist(.MascotasIndex(j)).Numero <> ELEMENTALTIERRA) Then
                          
80                        If Npclist(.MascotasIndex(j)).TargetNPC = 0 Then Npclist(.MascotasIndex(j)).TargetNPC = NpcIndex
90                        Npclist(.MascotasIndex(j)).Movement = TipoAI.NpcAtacaNpc
100                   End If
110                  End If
120               End If
130           Next j
140       End With
End Sub

Public Sub AllFollowAmo(ByVal userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim j As Integer
          
10        For j = 1 To MAXMASCOTAS
20            If UserList(userindex).MascotasIndex(j) > 0 Then
30                Call FollowAmo(UserList(userindex).MascotasIndex(j))
40            End If
50        Next j
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, ByVal userindex As Integer) As Boolean
      '*************************************************
      'Author: Unknown
      'Last modified: -
      '
      '*************************************************

10        With UserList(userindex)
20            If .flags.AdminInvisible = 1 Then Exit Function
30            If (Not .flags.Privilegios And PlayerType.User) <> 0 And Not .flags.AdminPerseguible Then Exit Function
40            If Not CanAttackReyCastle(userindex, NpcIndex) Then Exit Function
50        End With
          
60        With Npclist(NpcIndex)
              ' El npc puede atacar ???
70            If .CanAttack = 1 Then
80                NpcAtacaUser = True
90                Call CheckPets(NpcIndex, userindex, False)
                  
100               If .Target = 0 Then .Target = userindex
                  
110               If UserList(userindex).flags.AtacadoPorNpc = 0 And UserList(userindex).flags.AtacadoPorUser = 0 Then
120                   UserList(userindex).flags.AtacadoPorNpc = NpcIndex
130               End If
140           Else
150               NpcAtacaUser = False
160               Exit Function
170           End If
              
180           .CanAttack = 0
              
190           If .flags.Snd1 > 0 Then
200               Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))
210           End If
220       End With
          
230       If NpcImpacto(NpcIndex, userindex) Then
240           With UserList(userindex)
250               Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
                  
260               If .flags.Meditando = False Then
270                   If .flags.Navegando = 0 And .flags.Montando = 0 Then
280                       Call SendData(SendTarget.ToPCArea, userindex, PrepareMessageCreateFX(.Char.CharIndex, FXSANGRE, 0))
290                   End If
300               End If
                  
310               Call NpcDaño(NpcIndex, userindex)
320               Call WriteUpdateHP(userindex)
330               Call WriteUpdateFollow(userindex)
                  
                  '¿Puede envenenar?
340               If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(userindex)
350           End With
              
360           Call SubirSkill(userindex, eSkill.Tacticas, False)
370       Else
380           Call WriteMultiMessage(userindex, eMessages.NPCSwing)
390           Call SubirSkill(userindex, eSkill.Tacticas, True)
400       End If
          
          'Controla el nivel del usuario
410       Call CheckUserLevel(userindex)
End Function

Private Function NpcImpactoNpc(ByVal Atacante As Integer, ByVal Victima As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim PoderAtt As Long
          Dim PoderEva As Long
          Dim ProbExito As Long
          
10        PoderAtt = Npclist(Atacante).PoderAtaque
20        PoderEva = Npclist(Victima).PoderEvasion
          
          ' Chances are rounded
30        ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtt - PoderEva) * 0.4))
40        NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
End Function

Public Sub NpcDañoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim daño As Integer
          
10        With Npclist(Atacante)
20            daño = RandomNumber(.Stats.MinHIT, .Stats.MaxHIT)
30            Npclist(Victima).Stats.MinHp = Npclist(Victima).Stats.MinHp - daño
40            CalcularDarExp .MaestroUser, Victima, daño
              
50            If Npclist(Victima).Stats.MinHp < 1 Then
60                .Movement = .flags.OldMovement
                  
70                If LenB(.flags.AttackedBy) <> 0 Then
80                    .Hostile = .flags.OldHostil
90                End If
                  
100               If .MaestroUser > 0 Then
110                   Call FollowAmo(Atacante)
120               End If
                  
130               Call MuereNpc(Victima, .MaestroUser)
140           End If
150       End With
End Sub
        

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, ByVal Victima As Integer, Optional ByVal cambiarMOvimiento As Boolean = True)
      '*************************************************
      'Author: Unknown
      'Last modified: 01/03/2009
      '01/03/2009: ZaMa - Las mascotas no pueden atacar al rey si quedan pretorianos vivos.
      '23/05/2010: ZaMa - Ahora los elementales renuevan el tiempo de pertencia del npc que atacan si pertenece a su amo.
      '*************************************************
         
          Dim MasterIndex As Integer
         
10        With Npclist(Atacante)
             
               'Es el Rey Preatoriano?
20            If Npclist(Victima).Numero = PRKING_NPC Then
30                If pretorianosVivos > 0 Then
40                    Call WriteConsoleMsg(.MaestroUser, "Debes matar al resto del ejército antes de atacar al rey!", FontTypeNames.FONTTYPE_FIGHT)
50                    .TargetNPC = 0
60                    Exit Sub
70                End If
80            End If
             
              ' El npc puede atacar ???
90            If .CanAttack = 1 Then
100               .CanAttack = 0
110               If cambiarMOvimiento Then
120                   Npclist(Victima).TargetNPC = Atacante
130                   Npclist(Victima).Movement = TipoAI.NpcAtacaNpc
140               End If
150           Else
160               Exit Sub
170           End If
             
180           If .flags.Snd1 > 0 Then
190               Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(.flags.Snd1, .Pos.X, .Pos.Y))
200           End If
             
210           MasterIndex = .MaestroUser
             
              ' Tiene maestro?
220           If MasterIndex > 0 Then
                  ' Su maestro es dueño del npc al que ataca?
230               If Npclist(Victima).Owner = MasterIndex Then
                      ' Renuevo el timer de pertenencia
240                   Call IntervaloPerdioNpc(MasterIndex, True)
250               End If
260           End If
             
             
270           If NpcImpactoNpc(Atacante, Victima) Then
280               If Npclist(Victima).flags.Snd2 > 0 Then
290                   Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(Npclist(Victima).flags.Snd2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
300               Else
310                   Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
320               End If
             
330               If MasterIndex > 0 Then
340                   Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
350               Else
360                   Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_IMPACTO, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
370               End If
                 
380               Call NpcDañoNpc(Atacante, Victima)
390           Else
400               If MasterIndex > 0 Then
410                   Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
420               Else
430                   Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayWave(SND_SWING, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y))
440               End If
450           End If
460       End With
End Sub
Public Function UsuarioAtacaNpc(ByVal userindex As Integer, _
                                ByVal NpcIndex As Integer) As Boolean
 
        '***************************************************
        'Author: Unknown
        'Last Modification: 14/01/2010 (ZaMa)
        '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados por npcs cuando los atacan.
        '14/01/2010: ZaMa - Lo transformo en función, para que no se pierdan municiones al atacar targets inválidos.
        '***************************************************
        '<EhHeader>
10      On Error GoTo UsuarioAtacaNpc_Err
 
        '</EhHeader>
 
20      If Not PuedeAtacarNPC(userindex, NpcIndex) Then Exit Function
 
30      Call NPCAtacado(NpcIndex, userindex)
 
40      If UserImpactoNpc(userindex, NpcIndex) Then
50          If Npclist(NpcIndex).flags.Snd2 > 0 Then
60              Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(Npclist(NpcIndex).flags.Snd2, _
                        Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
70          Else
80              Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_IMPACTO2, Npclist( _
                        NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y))
 
90          End If
     
100         Call UserDañoNpc(userindex, NpcIndex)
110     Else
120         Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_SWING, UserList( _
                    userindex).Pos.X, UserList(userindex).Pos.Y))
130         Call WriteMultiMessage(userindex, eMessages.UserSwing)
 
140     End If
 
        ' Reveló su condición de usuario al atacar, los npcs lo van a atacar
150     UserList(userindex).flags.Ignorado = False
 
160     UsuarioAtacaNpc = True
 
        '<EhFooter>
170     Exit Function
 
UsuarioAtacaNpc_Err:
180     LogError Err.Description & vbCrLf & "UsuarioAtacaNpc " & NpcIndex & " " & "at line " & Erl
             
        '</EhFooter>
End Function
Public Sub UsuarioAtaca(ByVal userindex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim Index As Integer
          Dim AttackPos As WorldPos
          
          'Check bow's interval
10        If Not IntervaloPermiteUsarArcos(userindex, False) Then Exit Sub
          
          'Check Spell-Magic interval
20        If Not IntervaloPermiteMagiaGolpe(userindex) Then
              'Check Attack interval
30            If Not IntervaloPermiteAtacar(userindex) Then
40                Exit Sub
50            End If
60        End If
          
              Dim loquebaja As Byte
          
70        With UserList(userindex)
80       If .Invent.WeaponEqpObjIndex > 0 Then
90                If ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).QuitaEnergia = 0 Then
100               loquebaja = RandomNumber(1, 10)
110                   If .Stats.MinSta - loquebaja <= 0 Then
120                   Call WriteConsoleMsg(userindex, "Estás muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
130                   Exit Sub
140                   Else
150                   Call QuitarSta(userindex, loquebaja)
160                   End If
170               Else
180                   If UserList(userindex).Stats.MinSta >= ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).QuitaEnergia Then
190                   Call QuitarSta(userindex, ObjData(UserList(userindex).Invent.WeaponEqpObjIndex).QuitaEnergia)
200                   Else
210                   Call WriteConsoleMsg(userindex, "Estás muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
220                   Exit Sub
230                   End If
240               End If
250           Else
260           Call QuitarSta(userindex, RandomNumber(1, 10))
270           End If
280           SendData SendTarget.ToPCArea, userindex, PrepareMessageMovimientSW(.Char.CharIndex, 1)
              
290           AttackPos = .Pos
300           Call HeadtoPos(.Char.Heading, AttackPos)

              'Exit if not legal
310           If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
320               Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
330               Exit Sub
340           End If
              
350           Index = MapData(AttackPos.map, AttackPos.X, AttackPos.Y).userindex
              
              'Look for user
360           If Index > 0 Then
370               Call UsuarioAtacaUsuario(userindex, Index)
380               Call WriteUpdateUserStats(userindex)
390               Call WriteUpdateUserStats(Index)
400               Exit Sub
410           End If
              
420           Index = MapData(AttackPos.map, AttackPos.X, AttackPos.Y).NpcIndex
              
             'Look for NPC
430           If Index > 0 Then
440               If Npclist(Index).Attackable Then
450                   If Npclist(Index).MaestroUser > 0 And MapInfo(Npclist(Index).Pos.map).Pk = False Then
460                       Call WriteConsoleMsg(userindex, "No puedes atacar mascotas en zona segura.", FontTypeNames.FONTTYPE_FIGHT)
470                       Exit Sub
480                   End If
                     
490                   Call UsuarioAtacaNpc(userindex, Index)
500               Else
510                   Call WriteConsoleMsg(userindex, "No puedes atacar a este NPC.", FontTypeNames.FONTTYPE_FIGHT)
520               End If
                 
530               Call WriteUpdateUserStats(userindex)
                 
540               Exit Sub
550           End If
             
560           Call SendData(SendTarget.ToPCArea, userindex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
570           Call WriteUpdateUserStats(userindex)
             
580           If .Counters.Trabajando Then .Counters.Trabajando = .Counters.Trabajando - 1
                 
590           If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1
600       End With
End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
10        On Error GoTo UsuarioImpacto_Error
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim ProbRechazo As Long
          Dim Rechazo As Boolean
          Dim ProbExito As Long
          Dim PoderAtaque As Long
          Dim UserPoderEvasion As Long
          Dim UserPoderEvasionEscudo As Long
          Dim Arma As Integer
          Dim SkillTacticas As Long
          Dim SkillDefensa As Long
          Dim ProbEvadir As Long
          Dim Skill As eSkill
          
20        SkillTacticas = UserList(VictimaIndex).Stats.UserSkills(eSkill.Tacticas)
30        SkillDefensa = UserList(VictimaIndex).Stats.UserSkills(eSkill.Defensa)
          
40        Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
          
          'Calculamos el poder de evasion...
50        UserPoderEvasion = PoderEvasion(VictimaIndex)
          
60        If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
70           UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
80           UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo
90        Else
100           UserPoderEvasionEscudo = 0
110       End If
          
          'Esta usando un arma ???
120       If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
130           If ObjData(Arma).proyectil = 1 Then
140               PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
150               Skill = eSkill.Proyectiles
160           Else
170               PoderAtaque = PoderAtaqueArma(AtacanteIndex)
180               Skill = eSkill.Armas
190           End If
200       Else
210           PoderAtaque = PoderAtaqueWrestling(AtacanteIndex)
220           Skill = eSkill.Wrestling
230       End If
          
          ' Chances are rounded
240       ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtaque - UserPoderEvasion) * 0.4))
          
          ' Se reduce la evasion un 25%
250       If UserList(VictimaIndex).flags.Meditando = True Then
260           ProbEvadir = (100 - ProbExito) * 0.75
270           ProbExito = MinimoInt(90, 100 - ProbEvadir)
280       End If
          
290       UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)
          
          ' el usuario esta usando un escudo ???
300       If UserList(VictimaIndex).Invent.EscudoEqpObjIndex > 0 Then
              'Fallo ???
310           If Not UsuarioImpacto Then
                  ' Chances are rounded
                  'El +1 sirve para que nunca divida por cero.
320               ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (1 + (SkillDefensa + SkillTacticas))))
330               Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
340               If Rechazo Then
                      'Se rechazo el ataque con el escudo
350                   Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayWave(SND_ESCUDO, UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y))
360                                    SendData SendTarget.ToPCArea, VictimaIndex, PrepareMessageMovimientSW(UserList(VictimaIndex).Char.CharIndex, 2)
370                   Call WriteMultiMessage(AtacanteIndex, eMessages.BlockedWithShieldother)
380                   Call WriteMultiMessage(VictimaIndex, eMessages.BlockedWithShieldUser)
                      
390                   Call SubirSkill(VictimaIndex, eSkill.Defensa, True)
400               Else
410                   Call SubirSkill(VictimaIndex, eSkill.Defensa, False)
420               End If
430           End If
440       End If
          
450       If Not UsuarioImpacto Then
460           Call SubirSkill(AtacanteIndex, Skill, False)
470       End If
          
480       Call FlushBuffer(VictimaIndex)
          
490       Exit Function
          
    
500       On Error GoTo 0
510       Exit Function

UsuarioImpacto_Error:

520       LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure UsuarioImpacto, line " & Erl & "."

End Function

Public Function UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer) As Boolean
      '***************************************************
      'Author: Unknown
      'Last Modification: 14/01/2010 (ZaMa)
      '14/01/2010: ZaMa - Lo transformo en función, para que no se pierdan municiones al atacar targets
      '                    inválidos, y evitar un doble chequeo innecesario
      '***************************************************

10    On Error GoTo Errhandler

20        If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Function
30        With UserList(AtacanteIndex)
          
40            If Distancia(.Pos, UserList(VictimaIndex).Pos) > MAXDISTANCIAARCO Then
50               Call WriteConsoleMsg(AtacanteIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
60               Exit Function
70            End If
              
80            Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)
              
90            If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
100               Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_IMPACTO, .Pos.X, .Pos.Y))
                  
110               If UserList(VictimaIndex).flags.Navegando = 0 Then
120                   Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.CharIndex, FXSANGRE, 0))
130               End If
                  
                  'Pablo (ToxicWaste): Guantes de Hurto del Bandido en acción
                      
                  'y ahora, el ladrón puede llegar a paralizar con el golpe.
140               If .clase = eClass.Thief Then
150                   Call DoHandInmo(AtacanteIndex, VictimaIndex)
160               End If
                  
170               Call SubirSkill(VictimaIndex, eSkill.Tacticas, False)
180               Call UserDañoUser(AtacanteIndex, VictimaIndex)
190           Else
                  ' Invisible admins doesn't make sound to other clients except itself
200               If .flags.AdminInvisible = 1 Then
210                   Call EnviarDatosASlot(AtacanteIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
220               Else
230                   Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayWave(SND_SWING, .Pos.X, .Pos.Y))
240               End If
                  
250               Call WriteMultiMessage(AtacanteIndex, eMessages.UserSwing)
260               Call WriteMultiMessage(VictimaIndex, eMessages.UserAttackedSwing, AtacanteIndex)
270               Call SubirSkill(VictimaIndex, eSkill.Tacticas, True)
280           End If
              
290           If .clase = eClass.Thief Then Call Desarmar(AtacanteIndex, VictimaIndex)
              
300           If .flags.SlotEvent > 0 Then
310               If Events(.flags.SlotEvent).Modality = Aracnus Then
320                   Aracnus_Veneno AtacanteIndex, VictimaIndex
330               ElseIf Events(.flags.SlotEvent).Modality = Minotauro Then
340                   Minotauro_Veneno AtacanteIndex, VictimaIndex
350               End If
360           End If
              
370       End With
          
380       UsuarioAtacaUsuario = True
          
390       Exit Function
          
Errhandler:
400       Call LogError("Error en UsuarioAtacaUsuario. Error " & Err.Number & " : " & Err.Description)
End Function

Public Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 12/01/2010 (ZaMa)
      '12/01/2010: ZaMa - Implemento armas arrojadizas y probabilidad de acuchillar
      '11/03/2010: ZaMa - Ahora no cuenta la muerte si estaba en estado atacable, y no se vuelve criminal
      '***************************************************
         
10    On Error GoTo Errhandler
       
          Dim daño As Long
          Dim Lugar As Byte
          Dim absorbido As Long
          Dim defbarco As Integer
          Dim defmontura As Integer
          Dim Obj As ObjData
          Dim Resist As Byte
         
20        daño = CalcularDaño(AtacanteIndex)
         
30        Call UserEnvenena(AtacanteIndex, VictimaIndex)
         
40        With UserList(AtacanteIndex)
50            If .flags.Navegando = 1 And .Invent.BarcoObjIndex > 0 Then
60                 Obj = ObjData(.Invent.BarcoObjIndex)
70                 daño = daño + RandomNumber(Obj.MinHIT, Obj.MaxHIT)
80            End If
              
90            If .flags.Montando = 1 Then
100                Obj = ObjData(.Invent.MonturaObjIndex)
110                daño = daño + RandomNumber(Obj.MinHIT, Obj.MaxHIT)
120           End If
              
              ' Usuario transformado en HOMBRE LOBO ,tiene 50% mas de daño.
130           If .flags.SlotEvent > 0 Then
140               If Events(.flags.SlotEvent).Modality = HombreLobo Then
150                   If Events(.flags.SlotEvent).Users(.flags.SlotUserEvent).Selected = 1 Then
160                       daño = daño * 1.5
170                   End If
180               End If
190           End If
            
              
              ' Efectos atacante [DIOS]
260           If UserPoderoso(AtacanteIndex) = 263 Then
270               daño = daño * 1.04
280           End If
              
              ' Efectos VICTIMA [Dios]
290           If UserPoderoso(VictimaIndex) = 262 Then
300               daño = daño * 0.95
310           End If
             
320           If UserList(VictimaIndex).flags.Navegando = 1 And UserList(VictimaIndex).Invent.BarcoObjIndex > 0 Then
330                Obj = ObjData(UserList(VictimaIndex).Invent.BarcoObjIndex)
340                defbarco = RandomNumber(Obj.MinDef, Obj.MaxDef)
350           End If
360           If UserList(VictimaIndex).flags.Montando = 1 And UserList(VictimaIndex).Invent.MonturaObjIndex > 0 Then
370                Obj = ObjData(UserList(VictimaIndex).Invent.MonturaObjIndex)
380                defmontura = RandomNumber(Obj.MinDef, Obj.MaxDef)
390           End If
             
400           If .Invent.WeaponEqpObjIndex > 0 Then
410               Resist = ObjData(.Invent.WeaponEqpObjIndex).Refuerzo
420           End If
             
430           Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
             
440           Select Case Lugar
                  Case PartesCuerpo.bCabeza
                      'Si tiene casco absorbe el golpe
450                   If UserList(VictimaIndex).Invent.CascoEqpObjIndex > 0 Then
460                       Obj = ObjData(UserList(VictimaIndex).Invent.CascoEqpObjIndex)
470                       absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
480                       absorbido = absorbido + defbarco + defmontura - Resist
490                       daño = daño - absorbido
500                       If daño < 0 Then daño = 1
510                   End If
                 
520               Case Else
                      'Si tiene armadura absorbe el golpe
530                   If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
540                       Obj = ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex)
                          Dim Obj2 As ObjData
550                       If UserList(VictimaIndex).Invent.EscudoEqpObjIndex Then
560                           Obj2 = ObjData(UserList(VictimaIndex).Invent.EscudoEqpObjIndex)
570                           absorbido = RandomNumber(Obj.MinDef + Obj2.MinDef, Obj.MaxDef + Obj2.MaxDef)
580                       Else
590                           absorbido = RandomNumber(Obj.MinDef, Obj.MaxDef)
600                       End If
610                       absorbido = absorbido + defbarco + defmontura - Resist
620                       daño = daño - absorbido
630                       If daño < 0 Then daño = 1
640                   End If
650           End Select
660           Call WriteMultiMessage(AtacanteIndex, eMessages.UserHittedUser, UserList(VictimaIndex).Char.CharIndex, Lugar, daño)
670           Call WriteMultiMessage(VictimaIndex, eMessages.UserHittedByUser, .Char.CharIndex, Lugar, daño)
680           UserList(VictimaIndex).Stats.MinHp = UserList(VictimaIndex).Stats.MinHp - daño
              
690           SendData SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateDamage(UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y, daño, DAMAGE_NORMAL)
              
700           If .flags.Hambre = 0 And .flags.Sed = 0 Then
                  'Si usa un arma quizas suba "Combate con armas"
710               If .Invent.WeaponEqpObjIndex > 0 Then
720                   If ObjData(.Invent.WeaponEqpObjIndex).proyectil Then
                          'es un Arco. Sube Armas a Distancia
730                       Call SubirSkill(AtacanteIndex, eSkill.Proyectiles, True)
                         
                          ' Si es arma arrojadiza..
740                       If ObjData(.Invent.WeaponEqpObjIndex).Municion = 0 Then
                              ' Si acuchilla
750                           If ObjData(.Invent.WeaponEqpObjIndex).Acuchilla = 1 Then
760                               Call DoAcuchillar(AtacanteIndex, 0, VictimaIndex, daño)
770                           End If
780                       End If
790                   Else
                          'Sube combate con armas.
800                       Call SubirSkill(AtacanteIndex, eSkill.Armas, True)
810                   End If
820               Else
                      'sino tal vez lucha libre
830                   Call SubirSkill(AtacanteIndex, eSkill.Wrestling, True)
840               End If
                         
                  'Trata de apuñalar por la espalda al enemigo
850                If PuedeApuñalar(AtacanteIndex) Then
860               UserList(AtacanteIndex).Dañoapu = daño
870                   Call DoApuñalar(AtacanteIndex, 0, VictimaIndex, daño)
880               End If
                  'e intenta dar un golpe crítico [Pablo (ToxicWaste)]
890                           Call DoGolpeCritico(AtacanteIndex, 0, VictimaIndex, daño)
900           End If
                  
                         'Doble chekeo innecesario, pero bueno..
              'Hecho para que no envie apu + golpe normal.
910           If Not PuedeApuñalar(AtacanteIndex) Then
920                  SendData SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateDamage(UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y, daño, DAMAGE_NORMAL)
930           End If
              
940           If UserList(VictimaIndex).Stats.MinHp <= 0 Then
                 
                  ' No cuenta la muerte si estaba en estado atacable
950               If UserList(VictimaIndex).flags.AtacablePor <> AtacanteIndex Then
                      'Store it!
                      'Call Statistics.StoreFrag(AtacanteIndex, VictimaIndex)
                      
960                   Call ContarMuerte(VictimaIndex, AtacanteIndex)
970               End If
                  
                  ' Para que las mascotas no sigan intentando luchar y
                  ' comiencen a seguir al amo
                  Dim j As Integer
980               For j = 1 To MAXMASCOTAS
990                   If .MascotasIndex(j) > 0 Then
1000                      If Npclist(.MascotasIndex(j)).Target = VictimaIndex Then
1010                          Npclist(.MascotasIndex(j)).Target = 0
1020                          Call FollowAmo(.MascotasIndex(j))
1030                      End If
1040                  End If
1050              Next j
                 
1060              Call ActStats(VictimaIndex, AtacanteIndex)
1070              Call UserDie(VictimaIndex, AtacanteIndex)
1080          Else
                  'Está vivo - Actualizamos el HP
1090              Call WriteUpdateHP(VictimaIndex)
1100              Call WriteUpdateFollow(VictimaIndex)
1110          End If
1120      End With
         
          'Controla el nivel del usuario
1130      Call CheckUserLevel(AtacanteIndex)
         
1140      Call FlushBuffer(VictimaIndex)
         
1150      Exit Sub
         
Errhandler:
          Dim AtacanteNick As String
          Dim VictimaNick As String
         
1160      If AtacanteIndex > 0 Then AtacanteNick = UserList(AtacanteIndex).Name
1170      If VictimaIndex > 0 Then VictimaNick = UserList(VictimaIndex).Name
         
1180      Call LogError("Error en UserDañoUser. Error " & Err.Number & " : " & Err.Description & " AtacanteIndex: " & _
                   AtacanteIndex & " Nick: " & AtacanteNick & " VictimaIndex: " & VictimaIndex & " Nick: " & VictimaNick & " LINEA: " & Erl)
End Sub

Sub UsuarioAtacadoPorUsuario(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)
      '***************************************************
      'Autor: Unknown
      'Last Modification: 05/05/2010
      'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
      '10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
      '05/05/2010: ZaMa - Ahora no suma puntos de bandido al atacar a alguien en estado atacable.
      '***************************************************

10        If TriggerZonaPelea(AttackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub
          
          Dim EraCriminal As Boolean
          Dim VictimaEsAtacable As Boolean
          
20        If Not criminal(AttackerIndex) Then
30            If Not criminal(VictimIndex) Then
                  ' Si la victima no es atacable por el agresor, entonces se hace pk
40                VictimaEsAtacable = UserList(VictimIndex).flags.AtacablePor = AttackerIndex
50                If Not VictimaEsAtacable Then Call VolverCriminal(AttackerIndex)
60            End If
70        End If
          
80        With UserList(VictimIndex)
90            If .flags.Meditando Then
100               .flags.Meditando = False
110               Call WriteMeditateToggle(VictimIndex)
120               Call WriteConsoleMsg(VictimIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
130               .Char.FX = 0
140               .Char.loops = 0
150               Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
160           End If
170       End With
          
180       EraCriminal = criminal(AttackerIndex)
          
          ' Si ataco a un atacable, no suma puntos de bandido
190       If Not VictimaEsAtacable Then
200           With UserList(AttackerIndex).Reputacion
210               If Not criminal(VictimIndex) Then
220                   .BandidoRep = .BandidoRep + vlASALTO
230                   If .BandidoRep > MAXREP Then .BandidoRep = MAXREP
                      
240                   .NobleRep = .NobleRep * 0.5
250                   If .NobleRep < 0 Then .NobleRep = 0
260               Else
270                   .NobleRep = .NobleRep + vlNoble
280                   If .NobleRep > MAXREP Then .NobleRep = MAXREP
290               End If
300           End With
310       End If
          
320       If criminal(AttackerIndex) Then
330           If UserList(AttackerIndex).Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(AttackerIndex)
              
340           If Not EraCriminal Then Call RefreshCharStatus(AttackerIndex)
350       ElseIf EraCriminal Then
360           Call RefreshCharStatus(AttackerIndex)
370       End If
          
380       Call AllMascotasAtacanUser(AttackerIndex, VictimIndex)
390       Call AllMascotasAtacanUser(VictimIndex, AttackerIndex)
          
          'Si la victima esta saliendo se cancela la salida
400       Call CancelExit(VictimIndex)
410       Call FlushBuffer(VictimIndex)
End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************
          'Reaccion de las mascotas
          Dim iCount As Integer
          
10        For iCount = 1 To MAXMASCOTAS
20            If UserList(Maestro).MascotasIndex(iCount) > 0 Then
30                Npclist(UserList(Maestro).MascotasIndex(iCount)).flags.AttackedBy = UserList(victim).Name
40                Npclist(UserList(Maestro).MascotasIndex(iCount)).Movement = TipoAI.NPCDEFENSA
50                Npclist(UserList(Maestro).MascotasIndex(iCount)).Hostile = 1
60            End If
70        Next iCount
End Sub

Public Function PuedeAtacar(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer) As Boolean
'***************************************************
'Autor: Unknown
'Last Modification: 02/04/2010
'Returns true if the AttackerIndex is allowed to attack the VictimIndex.
'24/01/2007 Pablo (ToxicWaste) - Ordeno todo y agrego situacion de Defensa en ciudad Armada y Caos.
'24/02/2009: ZaMa - Los usuarios pueden atacarse entre si.
'02/04/2010: ZaMa - Los armadas no pueden atacar nunca a los ciudas, salvo que esten atacables.
'***************************************************
10    On Error GoTo Errhandler

    'MUY importante el orden de estos "IF"...
    
    'Estas muerto no podes atacar
20    If UserList(AttackerIndex).flags.Muerto = 1 Then
30      Call WriteShortMsj(AttackerIndex, 5, FontTypeNames.FONTTYPE_INFO)
        'Call WriteConsoleMsg(AttackerIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
40      PuedeAtacar = False
50      Exit Function
60        End If
    
    'No podes atacar a alguien muerto
70    If UserList(VictimIndex).flags.Muerto = 1 Then
80      Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a un espíritu.", FontTypeNames.FONTTYPE_INFO)
90      PuedeAtacar = False
100     Exit Function
110       End If
    
    ' No podes atacar si estas en consulta
120   If UserList(AttackerIndex).flags.EnConsulta Then
130     Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
140     Exit Function
150       End If
    
    ' No podes atacar si esta en consulta
160   If UserList(VictimIndex).flags.EnConsulta Then
170     Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios mientras estan en consulta.", FontTypeNames.FONTTYPE_INFO)
180     Exit Function
190       End If
    
    'Retos
200   If UserList(VictimIndex).flags.SlotReto > 0 Then
210     With Retos(UserList(VictimIndex).flags.SlotReto)
220         If .Users(UserList(AttackerIndex).flags.SlotRetoUser).Team = .Users(UserList(VictimIndex).flags.SlotRetoUser).Team Then
230             PuedeAtacar = False
240             Exit Function
250         End If
260     End With
270       End If
    
    
    
280   If UserList(AttackerIndex).Counters.TimeFight > 0 Then
290     WriteConsoleMsg AttackerIndex, "No puedes atacar hasta que no termine la cuenta regresiva.", FontTypeNames.FONTTYPE_INFO
300     PuedeAtacar = False
310     Exit Function
320       End If
    
    
330   If UserList(AttackerIndex).flags.SlotEvent > 0 Then
            If UserList(VictimIndex).flags.SlotEvent <= 0 Then
                PuedeAtacar = False
                Exit Function
            End If
            
            If UserList(VictimIndex).flags.SlotEvent <= 0 Then
                PuedeAtacar = False
                Exit Function
            End If
            
340         If Events(UserList(AttackerIndex).flags.SlotEvent).Modality = DagaRusa Then
350             WriteConsoleMsg AttackerIndex, "No puedes atacar en este tipo de eventos.", FontTypeNames.FONTTYPE_INFO
360             PuedeAtacar = False
370             Exit Function
380         End If
        
390         If Events(UserList(AttackerIndex).flags.SlotEvent).TimeCount > 0 Then
400             WriteConsoleMsg AttackerIndex, "No puedes atacar hasta que no termine la cuenta regresiva.", FontTypeNames.FONTTYPE_INFO
410             PuedeAtacar = False
420             Exit Function
430         End If
        
            
440         If Events(UserList(AttackerIndex).flags.SlotEvent).Run Then
450             If UserList(AttackerIndex).flags.SlotUserEvent > 0 Then
460                 If Events(UserList(AttackerIndex).flags.SlotEvent).Users(UserList(AttackerIndex).flags.SlotUserEvent).Team > 0 Then
470                       If Not CanAttackUserEvent(AttackerIndex, VictimIndex) Then
480                            WriteConsoleMsg AttackerIndex, "No puedes atacar a tu compañero", FontTypeNames.FONTTYPE_INFO
490                            PuedeAtacar = False
500                            Exit Function
510                        End If
                        
520                    End If
530                End If
540            End If
550    End If

    'Estamos en una Arena? o un trigger zona segura?
560    Select Case TriggerZonaPelea(AttackerIndex, VictimIndex)
        Case eTrigger6.TRIGGER6_PERMITE
570         PuedeAtacar = (UserList(VictimIndex).flags.AdminInvisible = 0)
580         Exit Function
        
590     Case eTrigger6.TRIGGER6_PROHIBE
600         PuedeAtacar = False
610         Exit Function
        
620     Case eTrigger6.TRIGGER6_AUSENTE
            'Si no estamos en el Trigger 6 entonces es imposible atacar un gm
630         If (UserList(VictimIndex).flags.Privilegios And PlayerType.User) = 0 Then
640             If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(AttackerIndex, "El ser es demasiado poderoso.", FontTypeNames.FONTTYPE_WARNING)
650             PuedeAtacar = False
660             Exit Function
670         End If
680       End Select
    
    'Ataca un ciudadano?
690    If Not criminal(VictimIndex) Then
        ' El atacante es ciuda?
700     If Not criminal(AttackerIndex) Then
            ' El atacante es armada?
710         If esArmada(AttackerIndex) Then
                ' La victima es armada?
720             If esArmada(VictimIndex) Then
                    ' No puede
730                 Call WriteConsoleMsg(AttackerIndex, "Los soldados del ejército real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
740                 Exit Function
750             End If
760         End If
            
            ' Ciuda (o army) atacando a otro ciuda (o army)
770            If UserList(VictimIndex).flags.AtacablePor = AttackerIndex Then
                ' Se vuelve atacable.
780             If ToogleToAtackable(AttackerIndex, VictimIndex, False) Then
790                 PuedeAtacar = True
800                 Exit Function
810             End If
820         End If
830     End If
    ' Ataca a un criminal
840       Else
        'Sos un Caos atacando otro caos?
850        If esCaos(VictimIndex) Then
860         If esCaos(AttackerIndex) Then
870             Call WriteConsoleMsg(AttackerIndex, "Los miembros de la legión oscura tienen prohibido atacarse entre sí.", FontTypeNames.FONTTYPE_WARNING)
880             Exit Function
890         End If
900     End If
910       End If
    
    'Tenes puesto el seguro?
920    If UserList(AttackerIndex).flags.Seguro Then
930     If Not criminal(VictimIndex) Then
940         Call WriteConsoleMsg(AttackerIndex, "No puedes atacar ciudadanos, para hacerlo debes desactivar el seguro.", FontTypeNames.FONTTYPE_WARNING)
950         PuedeAtacar = False
960         Exit Function
970     End If
980       Else
        ' Un ciuda es atacado
990        If Not criminal(VictimIndex) Then
            ' Por un armada sin seguro
1000        If esArmada(AttackerIndex) Then
                ' No puede
1010            Call WriteConsoleMsg(AttackerIndex, "Los soldados del ejército real tienen prohibido atacar ciudadanos.", FontTypeNames.FONTTYPE_WARNING)
1020            PuedeAtacar = False
1030            Exit Function
1040        End If
1050    End If
1060   End If

    
    'Seguro de clanes?
    If UserList(AttackerIndex).GuildIndex <> 0 Then
        If UserList(AttackerIndex).SeguroClan = True Then
            If UserList(AttackerIndex).GuildIndex = UserList(VictimIndex).GuildIndex Then
                Call WriteConsoleMsg(AttackerIndex, "Debes desactivar el seguro de clan para atacar a miembros de tu mismo clan. Presiona la letra Z para desactivarlo.", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = False
                Exit Function
            End If
        End If
    End If
    
    
    'Estas en un Mapa Seguro?
1070   If MapInfo(UserList(VictimIndex).Pos.map).Pk = False Then
1080    If esArmada(AttackerIndex) Then
1090        If UserList(AttackerIndex).Faccion.RecompensasReal > 11 Then
1100            If UserList(VictimIndex).Pos.map = 58 Or UserList(VictimIndex).Pos.map = 59 Or UserList(VictimIndex).Pos.map = 60 Then
1110            Call WriteConsoleMsg(VictimIndex, "¡Huye de la ciudad! Estás siendo atacado y no podrás defenderte.", FontTypeNames.FONTTYPE_WARNING)
1120            PuedeAtacar = True 'Beneficio de Armadas que atacan en su ciudad.
1130            Exit Function
1140            End If
1150        End If
1160    End If
1170       If esCaos(AttackerIndex) Then
1180        If UserList(AttackerIndex).Faccion.RecompensasCaos > 11 Then
1190            If UserList(VictimIndex).Pos.map = 151 Or UserList(VictimIndex).Pos.map = 156 Then
1200            Call WriteConsoleMsg(VictimIndex, "¡Huye de la ciudad! Estás siendo atacado y no podrás defenderte.", FontTypeNames.FONTTYPE_WARNING)
1210            PuedeAtacar = True 'Beneficio de Caos que atacan en su ciudad.
1220            Exit Function
1230            End If
1240        End If
1250    End If
1260    Call WriteConsoleMsg(AttackerIndex, "Esta es una zona segura, aquí no puedes atacar a otros usuarios.", FontTypeNames.FONTTYPE_WARNING)
1270    PuedeAtacar = False
1280    Exit Function
1290      End If
    
    'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
1300   If MapData(UserList(VictimIndex).Pos.map, UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Or _
        MapData(UserList(AttackerIndex).Pos.map, UserList(AttackerIndex).Pos.X, UserList(AttackerIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
1310    Call WriteConsoleMsg(AttackerIndex, "No puedes pelear aquí.", FontTypeNames.FONTTYPE_WARNING)
1320    PuedeAtacar = False
1330    Exit Function
1340      End If
    
1350      PuedeAtacar = True
1360  Exit Function

Errhandler:
1370      Call LogError("Error en PuedeAtacar. Error " & Err.Number & " : " & Err.Description & " At line " & Erl)
End Function

Public Function PuedeAtacarNPC(ByVal AttackerIndex As Integer, ByVal NpcIndex As Integer, _
                Optional ByVal Paraliza As Boolean = False, Optional ByVal IsPet As Boolean = False) As Boolean
      '***************************************************
      'Autor: Unknown Author (Original version)
      'Returns True if AttackerIndex can attack the NpcIndex
      'Last Modification: 04/07/2010
      '24/01/2007 Pablo (ToxicWaste) - Orden y corrección de ataque sobre una mascota y guardias
      '14/08/2007 Pablo (ToxicWaste) - Reescribo y agrego TODOS los casos posibles cosa de usar
      'esta función para todo lo referente a ataque a un NPC. Ya sea Magia, Físico o a Distancia.
      '16/11/2009: ZaMa - Agrego validacion de pertenencia de npc.
      '02/04/2010: ZaMa - Los armadas ya no peuden atacar npcs no hotiles.
      '23/05/2010: ZaMa - El inmo/para renuevan el timer de pertenencia si el ataque fue a un npc propio.
      '04/07/2010: ZaMa - Ahora no se puede apropiar del dragon de dd.
      '***************************************************
       
10    On Error GoTo Errhandler
       
20        With Npclist(NpcIndex)
         
              'Estas muerto?
30            If UserList(AttackerIndex).flags.Muerto = 1 Then
                  'Call WriteConsoleMsg(AttackerIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
40                Call WriteShortMsj(AttackerIndex, 5, FontTypeNames.FONTTYPE_INFO)
50                Exit Function
60            End If
             
              'Sos consejero?
70            If UserList(AttackerIndex).flags.Privilegios And PlayerType.Consejero Then
                  'No pueden atacar NPC los Consejeros.
80                Exit Function
90            End If
             
            'Estas en modo Combate?
100       If Not UserList(AttackerIndex).flags.ModoCombate Then
110           Call WriteConsoleMsg(AttackerIndex, "Para realizar esta accion debes activar el modo combate, puedes hacerlo con la tecla ""C""", FontTypeNames.FONTTYPE_INFO)
120           PuedeAtacarNPC = False
130           Exit Function
140       End If
          
             
              ' No podes atacar si estas en consulta
150           If UserList(AttackerIndex).flags.EnConsulta Then
160               Call WriteConsoleMsg(AttackerIndex, "No puedes atacar npcs mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
170               Exit Function
180           End If
             
              'Es una criatura atacable?
190           If .Attackable = 0 Then
200               Call WriteConsoleMsg(AttackerIndex, "No puedes atacar esta criatura.", FontTypeNames.FONTTYPE_INFO)
210               Exit Function
220           End If
             
              'Es valida la distancia a la cual estamos atacando?
230           If Distancia(UserList(AttackerIndex).Pos, .Pos) >= MAXDISTANCIAARCO Then
240              Call WriteConsoleMsg(AttackerIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)
250              Exit Function
260           End If
             
              'Es una criatura No-Hostil?
270           If .Hostile = 0 Then
                  'Es Guardia del Caos?
280               If .NPCtype = eNPCType.Guardiascaos Then
                      'Lo quiere atacar un caos?
290                   If esCaos(AttackerIndex) Then
300                       Call WriteConsoleMsg(AttackerIndex, "No puedes atacar Guardias del Caos siendo de la legión oscura.", FontTypeNames.FONTTYPE_INFO)
310                       Exit Function
320                   End If
                  'Es guardia Real?
330               ElseIf .NPCtype = eNPCType.GuardiaReal Then
                      'Lo quiere atacar un Armada?
340                   If esArmada(AttackerIndex) Then
350                       Call WriteConsoleMsg(AttackerIndex, "No puedes atacar Guardias Reales siendo del ejército real.", FontTypeNames.FONTTYPE_INFO)
360                       Exit Function
370                   End If
                      'Tienes el seguro puesto?
380                   If UserList(AttackerIndex).flags.Seguro Then
390                       Call WriteConsoleMsg(AttackerIndex, "Para poder atacar Guardias Reales debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
400                       Exit Function
410                   Else
420                       Call WriteConsoleMsg(AttackerIndex, "¡Atacaste un Guardia Real! Eres un criminal.", FontTypeNames.FONTTYPE_INFO)
430                       Call VolverCriminal(AttackerIndex)
440                       PuedeAtacarNPC = True
450                       Exit Function
460                   End If
             
                  'No era un Guardia, asi que es una criatura No-Hostil común.
                  'Para asegurarnos que no sea una Mascota:
470               ElseIf .MaestroUser = 0 Then
                      'Si sos ciudadano tenes que quitar el seguro para atacarla.
480                   If Not criminal(AttackerIndex) Then
                         
                          ' Si sos armada no podes atacarlo directamente
490                       If esArmada(AttackerIndex) Then
500                           Call WriteConsoleMsg(AttackerIndex, "Los miembros del ejército real no pueden atacar npcs no hostiles.", FontTypeNames.FONTTYPE_INFO)
510                           Exit Function
520                       End If
                     
                          'Sos ciudadano, tenes el seguro puesto?
530                       If UserList(AttackerIndex).flags.Seguro Then
540                           Call WriteConsoleMsg(AttackerIndex, "Para atacar a este NPC debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
550                           Exit Function
560                       Else
                              'No tiene seguro puesto. Puede atacar pero es penalizado.
570                           Call WriteConsoleMsg(AttackerIndex, "Atacaste un NPC no-hostil. Continúa haciéndolo y te podrás convertir en criminal.", FontTypeNames.FONTTYPE_INFO)
                              'NicoNZ: Cambio para que al atacar npcs no hostiles no bajen puntos de nobleza
580                           Call DisNobAuBan(AttackerIndex, 0, 1000)
590                           PuedeAtacarNPC = True
600                           Exit Function
610                       End If
620                   End If
630               End If
640           End If
         
         
              Dim MasterIndex As Integer
650           MasterIndex = .MaestroUser
             
              'Es el NPC mascota de alguien?
660           If MasterIndex > 0 Then
                 
                  ' Dueño de la mascota ciuda?
670               If Not criminal(MasterIndex) Then
                     
                      ' Atacante ciuda?
680                   If Not criminal(AttackerIndex) Then
                         
                          ' Si esta en estado atacable puede atacar su mascota sin problemas
690                       If UserList(MasterIndex).flags.AtacablePor = AttackerIndex Then
                              ' Toogle to atacable and restart the timer
700                           Call ToogleToAtackable(AttackerIndex, MasterIndex)
710                           PuedeAtacarNPC = True
720                           Exit Function
730                       End If
                         
                          'Atacante armada?
740                       If esArmada(AttackerIndex) Then
                              'El atacante es Armada y esta intentando atacar mascota de un Ciudadano
750                           Call WriteConsoleMsg(AttackerIndex, "Los miembros del ejército real no pueden atacar mascotas de ciudadanos.", FontTypeNames.FONTTYPE_INFO)
760                           Exit Function
770                       End If
                         
                          'El atacante es Ciudadano y esta intentando atacar mascota de un Ciudadano.
780                       If UserList(AttackerIndex).flags.Seguro Then
                              'El atacante tiene el seguro puesto. No puede atacar.
790                           Call WriteConsoleMsg(AttackerIndex, "Para atacar mascotas de ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
800                           Exit Function
810                       Else
                              'El atacante no tiene el seguro puesto. Recibe penalización.
820                           Call WriteConsoleMsg(AttackerIndex, "Has atacado la Mascota de un ciudadano. Eres un criminal.", FontTypeNames.FONTTYPE_INFO)
830                           Call VolverCriminal(AttackerIndex)
840                           PuedeAtacarNPC = True
850                           Exit Function
860                       End If
870                   Else
                          'El atacante es criminal y quiere atacar un elemental ciuda, pero tiene el seguro puesto (NicoNZ)
880                       If UserList(AttackerIndex).flags.Seguro Then
890                           Call WriteConsoleMsg(AttackerIndex, "Para atacar mascotas de ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
900                           Exit Function
910                       End If
920                   End If
                 
                  'Es mascota de un caos?
930               ElseIf esCaos(MasterIndex) Then
                      'Es Caos el Dueño.
940                   If esCaos(AttackerIndex) Then
                          'Un Caos intenta atacar una criatura de un Caos. No puede atacar.
950                       Call WriteConsoleMsg(AttackerIndex, "Los miembros de la legión oscura no pueden atacar mascotas de otros legionarios. ", FontTypeNames.FONTTYPE_INFO)
960                       Exit Function
970                   End If
980               End If
                 
              ' No es mascota de nadie, le pertenece a alguien?
990           ElseIf .Owner > 0 Then
             
                  Dim OwnerUserIndex As Integer
1000              OwnerUserIndex = .Owner
                 
                  ' Puede atacar a su propia criatura!
1010              If OwnerUserIndex = AttackerIndex Then
1020                  PuedeAtacarNPC = True
1030                  Call IntervaloPerdioNpc(OwnerUserIndex, True) ' Renuevo el timer
1040                  Exit Function
1050              End If
                 
                  ' Esta compartiendo el npc con el atacante? => Puede atacar!
1060              If UserList(OwnerUserIndex).flags.ShareNpcWith = AttackerIndex Then
1070                  PuedeAtacarNPC = True
1080                  Exit Function
1090              End If
                 
                  ' Si son del mismo clan o party, pueden atacar (No renueva el timer)
1100              If Not SameClan(OwnerUserIndex, AttackerIndex) And Not SameParty(OwnerUserIndex, AttackerIndex) Then
                 
                      ' Si se le agoto el tiempo
1110                  If IntervaloPerdioNpc(OwnerUserIndex) Then ' Se lo roba :P
1120                      Call PerdioNpc(OwnerUserIndex)
1130                      Call ApropioNpc(AttackerIndex, NpcIndex)
1140                      PuedeAtacarNPC = True
1150                      Exit Function
                         
                      ' Si lanzo un hechizo de para o inmo
1160                  ElseIf Paraliza Then
                     
                          ' Si ya esta paralizado o inmobilizado, no puedo inmobilizarlo de nuevo
1170                      If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then
                             
                              'TODO_ZAMA: Si dejo esto asi, los pks con seguro peusto van a poder inmobilizar criaturas con dueño
                              ' Si es pk neutral, puede hacer lo que quiera :P.
1180                          If Not criminal(AttackerIndex) And Not criminal(OwnerUserIndex) Then
                             
                                   'El atacante es Armada
1190                              If esArmada(AttackerIndex) Then
                                     
                                       'Intententa paralizar un npc de un armada?
1200                                  If esArmada(OwnerUserIndex) Then
                                          'El atacante es Armada y esta intentando paralizar un npc de un armada: No puede
1210                                      Call WriteConsoleMsg(AttackerIndex, "Los miembros del Ejército Real no pueden paralizar criaturas ya paralizadas pertenecientes a otros miembros del Ejército Real", FontTypeNames.FONTTYPE_INFO)
1220                                      Exit Function
                                     
                                      'El atacante es Armada y esta intentando paralizar un npc de un ciuda
1230                                  Else
                                          ' Si tiene seguro no puede
1240                                      If UserList(AttackerIndex).flags.Seguro Then
1250                                          Call WriteConsoleMsg(AttackerIndex, "Para paralizar criaturas ya paralizadas pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
1260                                          Exit Function
1270                                      Else
                                              ' Si ya estaba atacable, no podrá atacar a un npc perteneciente a otro ciuda
1280                                          If ToogleToAtackable(AttackerIndex, OwnerUserIndex) Then
1290                                              Call WriteConsoleMsg(AttackerIndex, "Has paralizado la criatura de un ciudadano, ahora eres atacable por él.", FontTypeNames.FONTTYPE_INFO)
1300                                              PuedeAtacarNPC = True
1310                                          End If
                                             
1320                                          Exit Function
                                             
1330                                      End If
1340                                  End If
                                     
                                  ' El atacante es ciuda
1350                              Else
                                      'El atacante tiene el seguro puesto, no puede paralizar
1360                                  If UserList(AttackerIndex).flags.Seguro Then
1370                                      Call WriteConsoleMsg(AttackerIndex, "Para paralizar criaturas ya paralizadas pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
1380                                      Exit Function
                                         
                                      'El atacante no tiene el seguro puesto, ataca.
1390                                  Else
                                          ' Si ya estaba atacable, no podrá atacar a un npc perteneciente a otro ciuda
1400                                      If ToogleToAtackable(AttackerIndex, OwnerUserIndex) Then
1410                                          Call WriteConsoleMsg(AttackerIndex, "Has paralizado la criatura de un ciudadano, ahora eres atacable por él.", FontTypeNames.FONTTYPE_INFO)
1420                                          PuedeAtacarNPC = True
1430                                      End If
                                         
1440                                      Exit Function
1450                                  End If
1460                              End If
                                 
                              ' Al menos uno de los dos es criminal
1470                          Else
                                  ' Si ambos son caos
1480                              If esCaos(AttackerIndex) And esCaos(OwnerUserIndex) Then
                                      'El atacante es Caos y esta intentando paralizar un npc de un Caos
1490                                  Call WriteConsoleMsg(AttackerIndex, "Los miembros de la legión oscura no pueden paralizar criaturas ya paralizadas por otros legionarios.", FontTypeNames.FONTTYPE_INFO)
1500                                  Exit Function
1510                              End If
1520                          End If
                         
                          ' El npc no esta inmobilizado ni paralizado
1530                      Else
                              ' Si no tiene dueño, puede apropiarselo
1540                          If OwnerUserIndex = 0 Then
                             
                                  ' Siempre que no posea uno ya (el inmo/para no cambia pertenencia de npcs).
1550                              If UserList(AttackerIndex).flags.OwnedNpc = 0 Then
1560                                  Call ApropioNpc(AttackerIndex, NpcIndex)
1570                              End If
                                 
                              ' Si inmobiliza a su propio npc, renueva el timer
1580                          ElseIf OwnerUserIndex = AttackerIndex Then
1590                              Call IntervaloPerdioNpc(OwnerUserIndex, True) ' Renuevo el timer
1600                          End If
                             
                              ' Siempre se pueden paralizar/inmobilizar npcs con o sin dueño
                              ' que no tengan ese estado
1610                          PuedeAtacarNPC = True
1620                          Exit Function
       
1630                      End If
                         
                      ' No lanzó hechizos inmobilizantes
1640                  Else
                         
                          ' El npc le pertenece a un ciudadano
1650                      If Not criminal(OwnerUserIndex) Then
                             
                              'El atacante es Armada y esta intentando atacar un npc de un Ciudadano
1660                          If esArmada(AttackerIndex) Then
                              
1670                              If Not .flags.TeamEvent > 0 Then
                                      'Intententa atacar un npc de un armada?
1680                                  If esArmada(OwnerUserIndex) Then
                                          'El atacante es Armada y esta intentando atacar el npc de un armada: No puede
1690                                      Call WriteConsoleMsg(AttackerIndex, "Los miembros del Ejército Real no pueden atacar criaturas pertenecientes a otros miembros del Ejército Real", FontTypeNames.FONTTYPE_INFO)
1700                                      Exit Function
                                     
                                      'El atacante es Armada y esta intentando atacar un npc de un ciuda
1710                                  Else
                                         
                                          ' Si tiene seguro no puede
1720                                      If UserList(AttackerIndex).flags.Seguro Then
1730                                          Call WriteConsoleMsg(AttackerIndex, "Para atacar criaturas ya pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
1740                                          Exit Function
1750                                      Else
                                  
                                             
1760                                          Exit Function
1770                                      End If
1780                                  End If
1790                              End If
                                 
                              ' No es aramda, puede ser criminal o ciuda
1800                          Else
                                 
                                  'El atacante es Ciudadano y esta intentando atacar un npc de un Ciudadano.
1810                              If Not criminal(AttackerIndex) Then
                                     
1820                                  If UserList(AttackerIndex).flags.Seguro Then
                                          'El atacante tiene el seguro puesto. No puede atacar.
1830                                      Call WriteConsoleMsg(AttackerIndex, "Para atacar criaturas pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
1840                                      Exit Function
                                     
                                      'El atacante no tiene el seguro puesto, ataca.
1850                                  Else
1860                                      If ToogleToAtackable(AttackerIndex, OwnerUserIndex) Then
1870                                          Call WriteConsoleMsg(AttackerIndex, "Has atacado a la criatura de un ciudadano, ahora eres atacable por él.", FontTypeNames.FONTTYPE_INFO)
1880                                          PuedeAtacarNPC = True
1890                                      End If
                                         
1900                                      Exit Function
1910                                  End If
                                     
                                  'El atacante es criminal y esta intentando atacar un npc de un Ciudadano.
1920                              Else
                                      ' Es criminal atacando un npc de un ciuda, con seguro puesto.
1930                                  If UserList(AttackerIndex).flags.Seguro Then
1940                                      Call WriteConsoleMsg(AttackerIndex, "Para atacar criaturas pertenecientes a ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)
1950                                      Exit Function
1960                                  End If
                                     
1970                                  PuedeAtacarNPC = True
1980                              End If
1990                          End If
                             
                          ' Es npc de un criminal
2000                      Else
2010                          If Not .flags.TeamEvent > 0 Then
2020                              If esCaos(OwnerUserIndex) Then
                                      'Es Caos el Dueño.
2030                                  If esCaos(AttackerIndex) Then
                                          'Un Caos intenta atacar una npc de un Caos. No puede atacar.
2040                                      Call WriteConsoleMsg(AttackerIndex, "Los miembros de la Legión Oscura no pueden atacar criaturas de otros legionarios. ", FontTypeNames.FONTTYPE_INFO)
2050                                      Exit Function
2060                                  End If
2070                              End If
2080                          End If
2090                      End If
2100                  End If
2110              End If
                 
              ' Si no tiene dueño el npc, se lo apropia
2120          Else
                  ' Solo pueden apropiarse de npcs los caos, armadas o ciudas.
2130              If Not criminal(AttackerIndex) Or esCaos(AttackerIndex) Then
                      ' No puede apropiarse de los pretos!
2140                  If Npclist(NpcIndex).NPCtype <> eNPCType.pretoriano Then
                          ' No puede apropiarse del dragon de dd!
2150                      If Npclist(NpcIndex).NPCtype <> Dragon Then
                              ' Si es una mascota atacando, no se apropia del npc
2160                          If Not IsPet Then
                                  ' No es dueño de ningun npc => Se lo apropia.
2170                              If UserList(AttackerIndex).flags.OwnedNpc = 0 Then
2180                                  Call ApropioNpc(AttackerIndex, NpcIndex)
                                  ' Es dueño de un npc, pero no puede ser de este porque no tiene propietario.
2190                              Else
                                      ' Se va a adueñar del npc (y perder el otro) solo si no inmobiliza/paraliza
2200                                  If Not Paraliza Then Call ApropioNpc(AttackerIndex, NpcIndex)
2210                              End If
2220                          End If
2230                      End If
2240                  End If
2250              End If
2260          End If
              
              
2270          If (UserList(AttackerIndex).flags.SlotEvent) > 0 And (.flags.TeamEvent > 0) Then
2280              If Events(UserList(AttackerIndex).flags.SlotEvent).Modality = CastleMode Then
2290                  If Not EventosDS.CanAttackReyCastle(AttackerIndex, NpcIndex) Then
2300                      WriteConsoleMsg AttackerIndex, "No puedes atacar a tu rey", FontTypeNames.FONTTYPE_FIGHT
2310                      Exit Function
2320                  End If
                  
2330              End If
              
2340          End If
              
2350      End With
         
          'Es el Rey Preatoriano?
2360      If esPretoriano(NpcIndex) = 4 Then
2370          If pretorianosVivos > 0 Then
2380              Call WriteConsoleMsg(AttackerIndex, "Debes matar al resto del ejército antes de atacar al rey.", FontTypeNames.FONTTYPE_FIGHT)
2390              Exit Function
2400          End If
2410      End If
         
2420      PuedeAtacarNPC = True
             
2430      Exit Function
             
Errhandler:
         
          Dim AtckName As String
          Dim OwnerName As String
       
2440      If AttackerIndex > 0 Then AtckName = UserList(AttackerIndex).Name
2450      If OwnerUserIndex > 0 Then OwnerName = UserList(OwnerUserIndex).Name
         
2460      Call LogError("Error en PuedeAtacarNpc. Erorr: " & Err.Number & " - " & Err.Description & " Atacante: " & _
                         AttackerIndex & "-> " & AtckName & ". Owner: " & OwnerUserIndex & "-> " & OwnerName & _
                         ". NpcIndex: " & NpcIndex & ".")
End Function

Private Function SameClan(ByVal userindex As Integer, ByVal OtherUserIndex As Integer) As Boolean
      '***************************************************
      'Autor: ZaMa
      'Returns True if both players belong to the same clan.
      'Last Modification: 16/11/2009
      '***************************************************
10        SameClan = (UserList(userindex).GuildIndex = UserList(OtherUserIndex).GuildIndex) And _
                      UserList(userindex).GuildIndex <> 0
End Function

Private Function SameParty(ByVal userindex As Integer, ByVal OtherUserIndex As Integer) As Boolean
      '***************************************************
      'Autor: ZaMa
      'Returns True if both players belong to the same party.
      'Last Modification: 16/11/2009
      '***************************************************
10        SameParty = UserList(userindex).GroupIndex = UserList(OtherUserIndex).GroupIndex And _
                      UserList(userindex).GroupIndex <> 0
End Function

Sub CalcularDarExp(ByVal userindex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)
      '***************************************************
      'Autor: Nacho (Integer)
      'Last Modification: 03/09/06 Nacho
      'Reescribi gran parte del Sub
      'Ahora, da toda la experiencia del npc mientras este vivo.
      '***************************************************
          Dim ExpaDar As Long
          
          '[Nacho] Chekeamos que las variables sean validas para las operaciones
10        If ElDaño <= 0 Then ElDaño = 0
20        If Npclist(NpcIndex).Stats.MaxHp <= 0 Then Exit Sub
30        If ElDaño > Npclist(NpcIndex).Stats.MinHp Then ElDaño = Npclist(NpcIndex).Stats.MinHp
          
          
          'Npclist(NpcIndex).Stats.MinHp = 1
40        If ElDaño < 0 Then ElDaño = Npclist(NpcIndex).Stats.MinHp
          '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
50        ExpaDar = CLng(ElDaño * (Npclist(NpcIndex).GiveEXP / Npclist(NpcIndex).Stats.MaxHp))
60        If ExpaDar <= 0 Then Exit Sub
          
          '[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
                  'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
                  'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
70        If ExpaDar > Npclist(NpcIndex).flags.ExpCount Then
80            ExpaDar = Npclist(NpcIndex).flags.ExpCount
90            Npclist(NpcIndex).flags.ExpCount = 0
100       Else
110           Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).flags.ExpCount - ExpaDar
120       End If
          
          '[Nacho] Le damos la exp al user
130       If ExpaDar > 0 Then
140           If userindex <> 0 Then ' FIX CHOTO A LOS OSOS
150               If UserList(userindex).GroupIndex > 0 Then
160                   Call mGroup.AddExpGroup(UserList(userindex).GroupIndex, ExpaDar)
170               Else
180                   UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + ExpaDar
190                   If UserList(userindex).Stats.Exp > MAXEXP Then _
                          UserList(userindex).Stats.Exp = MAXEXP
200                   Call WriteConsoleMsg(userindex, "Has ganado " & ExpaDar & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT)
210                   If UserList(userindex).flags.Oro = 1 Then
220                   UserList(userindex).Stats.Exp = UserList(userindex).Stats.Exp + (ExpaDar * 0.35)
230                        WriteConsoleMsg userindex, "Aumento de exp 35% > Has ganado " & (ExpaDar * 0.35) & " puntos de experiencia.", FontTypeNames.fonttype_dios
240               End If
250           End If
              
              
260           Call CheckUserLevel(userindex)
270           End If
280       End If
End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, ByVal Destino As Integer) As eTrigger6
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

      'TODO: Pero que rebuscado!!
      'Nigo:  Te lo rediseñe, pero no te borro el TODO para que lo revises.
10    On Error GoTo Errhandler
          Dim tOrg As eTrigger
          Dim tDst As eTrigger
          
20        tOrg = MapData(UserList(Origen).Pos.map, UserList(Origen).Pos.X, UserList(Origen).Pos.Y).trigger
30        tDst = MapData(UserList(Destino).Pos.map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).trigger
          
40        If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
50            If tOrg = tDst Then
60                TriggerZonaPelea = TRIGGER6_PERMITE
70            Else
80                TriggerZonaPelea = TRIGGER6_PROHIBE
90            End If
100       Else
110           TriggerZonaPelea = TRIGGER6_AUSENTE
120       End If

130   Exit Function
Errhandler:
140       TriggerZonaPelea = TRIGGER6_AUSENTE
150       LogError ("Error en TriggerZonaPelea Origen: " & Origen & " y Destino " & Destino & "-" & Err.Description)
End Function

Sub UserEnvenena(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim ObjInd As Integer
          
10        ObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
          
20        If ObjInd > 0 Then
30            If ObjData(ObjInd).proyectil = 1 Then
40                ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex
50            End If
              
60            If ObjInd > 0 Then
70                If ObjData(ObjInd).Envenena = 1 Then
                      
80                    If RandomNumber(1, 100) < 60 Then
90                        UserList(VictimaIndex).flags.Envenenado = 1
100                       Call WriteConsoleMsg(VictimaIndex, "¡¡" & UserList(AtacanteIndex).Name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
110                       Call WriteConsoleMsg(AtacanteIndex, "¡¡Has envenenado a " & UserList(VictimaIndex).Name & "!!", FontTypeNames.FONTTYPE_FIGHT)
120                   End If
130               End If
140           End If
150       End If
          
160       Call FlushBuffer(VictimaIndex)
End Sub
