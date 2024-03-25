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

Public Const MAXDISTANCIAARCO  As Byte = 18

Public Const MAXDISTANCIAMAGIA As Byte = 18

Public Function MinimoInt(ByVal A As Integer, ByVal B As Integer) As Integer

    If A > B Then
        MinimoInt = B
    Else
        MinimoInt = A

    End If

End Function

Public Function MaximoInt(ByVal A As Integer, ByVal B As Integer) As Integer

    If A > B Then
        MaximoInt = A
    Else
        MaximoInt = B

    End If

End Function

Private Function PoderEvasionEscudo(ByVal UserIndex As Integer) As Long

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo PoderEvasionEscudo_Err

    '</EhHeader>

    PoderEvasionEscudo = (UserList(UserIndex).Stats.UserSkills(eSkill.Defensa) * Balance.ModClase(UserList(UserIndex).Clase).Escudo) / 2
    '<EhFooter>
    Exit Function

PoderEvasionEscudo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.SistemaCombate.PoderEvasionEscudo " & "at line " & Erl
        
    '</EhFooter>
End Function

Private Function PoderEvasion(ByVal UserIndex As Integer) As Long

    '<EhHeader>
    On Error GoTo PoderEvasion_Err

    '</EhHeader>

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Dim lTemp As Long

    With UserList(UserIndex)
        lTemp = (.Stats.UserSkills(eSkill.Tacticas) + .Stats.UserSkills(eSkill.Tacticas) / 33 * .Stats.UserAtributos(eAtributos.Agilidad)) * Balance.ModClase(.Clase).Evasion
       
        PoderEvasion = (lTemp + (2.5 * MaximoInt((.Stats.Elv) - 12, 0)))

        ' # Reduce Evasion en EVENTOS MISMA CLASE
        If .flags.SlotEvent > 0 Then
            If Events(.flags.SlotEvent).ChangeClass > 0 Then
                PoderEvasion = PoderEvasion * 0.75

            End If

        End If
              
        If .Pos.Map <> 130 And .Pos.Map <> 131 And .Pos.Map <> 132 Then

            ' # Evasión para los clanes con CASTILLO SUR
            If Castle_CheckBonus(.GuildIndex, eCastle.CASTLE_SOUTH) Then
                PoderEvasion = PoderEvasion * 1.02

            End If

        End If
            
    End With

    '<EhFooter>
    Exit Function

PoderEvasion_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.SistemaCombate.PoderEvasion " & "at line " & Erl
        
    '</EhFooter>
End Function

Private Function PoderAtaqueArma(ByVal UserIndex As Integer) As Long

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo PoderAtaqueArma_Err

    '</EhHeader>

    Dim PoderAtaqueTemp As Long
    
    With UserList(UserIndex)

        If .Stats.UserSkills(eSkill.Armas) < 31 Then
            PoderAtaqueTemp = .Stats.UserSkills(eSkill.Armas) * Balance.ModClase(.Clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Armas) < 61 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + .Stats.UserAtributos(eAtributos.Agilidad)) * Balance.ModClase(.Clase).AtaqueArmas
        ElseIf .Stats.UserSkills(eSkill.Armas) < 91 Then
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * Balance.ModClase(.Clase).AtaqueArmas
        Else
            PoderAtaqueTemp = (.Stats.UserSkills(eSkill.Armas) + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * Balance.ModClase(.Clase).AtaqueArmas

        End If
        
        PoderAtaqueArma = (PoderAtaqueTemp + (2.5 * MaximoInt((.Stats.Elv) - 12, 0)))

    End With

    '<EhFooter>
    Exit Function

PoderAtaqueArma_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.SistemaCombate.PoderAtaqueArma " & "at line " & Erl
        
    '</EhFooter>
End Function

Private Function PoderAtaqueProyectil(ByVal UserIndex As Integer) As Long

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo PoderAtaqueProyectil_Err

    '</EhHeader>

    Dim PoderAtaqueTemp  As Long

    Dim SkillProyectiles As Integer
    
    With UserList(UserIndex)
     
        SkillProyectiles = .Stats.UserSkills(eSkill.Proyectiles)
    
        If SkillProyectiles < 31 Then
            PoderAtaqueTemp = SkillProyectiles * Balance.ModClase(.Clase).AtaqueProyectiles
        ElseIf SkillProyectiles < 61 Then
            PoderAtaqueTemp = (SkillProyectiles + .Stats.UserAtributos(eAtributos.Agilidad)) * Balance.ModClase(.Clase).AtaqueProyectiles
        ElseIf SkillProyectiles < 91 Then
            PoderAtaqueTemp = (SkillProyectiles + 2 * .Stats.UserAtributos(eAtributos.Agilidad)) * Balance.ModClase(.Clase).AtaqueProyectiles
        Else
            PoderAtaqueTemp = (SkillProyectiles + 3 * .Stats.UserAtributos(eAtributos.Agilidad)) * Balance.ModClase(.Clase).AtaqueProyectiles

        End If
        
        PoderAtaqueProyectil = (PoderAtaqueTemp + (2.5 * MaximoInt((.Stats.Elv) - 12, 0)))

    End With

    '<EhFooter>
    Exit Function

PoderAtaqueProyectil_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.SistemaCombate.PoderAtaqueProyectil " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function UserImpactoNpc(ByVal UserIndex As Integer, _
                               ByVal NpcIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo UserImpactoNpc_Err

    '</EhHeader>

    Dim PoderAtaque As Long

    Dim Arma        As Integer

    Dim Skill       As eSkill

    Dim ProbExito   As Long
    
    Arma = UserList(UserIndex).Invent.WeaponEqpObjIndex
    
    If Arma > 0 Then 'Usando un arma
        If ObjData(Arma).proyectil = 1 Then
            PoderAtaque = PoderAtaqueProyectil(UserIndex)
            Skill = eSkill.Proyectiles
        Else
            PoderAtaque = PoderAtaqueArma(UserIndex)
            Skill = eSkill.Armas

        End If

    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((PoderAtaque - Npclist(NpcIndex).PoderEvasion) * 0.4)))
    
    UserImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
    
    If UserImpactoNpc Then
        If Skill Then Call SubirSkill(UserIndex, Skill, True)
    Else

        If Skill Then Call SubirSkill(UserIndex, Skill, False)

    End If
        
    Npclist(NpcIndex).Target = UserIndex

    '<EhFooter>
    Exit Function

UserImpactoNpc_Err:
    LogError Err.description & vbCrLf & "in UserImpactoNpc " & "at line " & Erl

    '</EhFooter>
End Function

Public Function NpcImpacto(ByVal NpcIndex As Integer, _
                           ByVal UserIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo NpcImpacto_Err

    '</EhHeader>

    '*************************************************
    'Author: Unknown
    'Last modified: 03/15/2006
    'Revisa si un NPC logra impactar a un user o no
    '03/15/2006 Maraxus - Evité una división por cero que eliminaba NPCs
    '*************************************************
    Dim Rechazo           As Boolean

    Dim ProbRechazo       As Long

    Dim ProbExito         As Long

    Dim UserEvasion       As Long

    Dim NpcPoderAtaque    As Long

    Dim PoderEvasioEscudo As Long

    Dim SkillTacticas     As Long

    Dim SkillDefensa      As Long
    
    UserEvasion = PoderEvasion(UserIndex)
    NpcPoderAtaque = Npclist(NpcIndex).PoderAtaque
    PoderEvasioEscudo = PoderEvasionEscudo(UserIndex)
    
    SkillTacticas = UserList(UserIndex).Stats.UserSkills(eSkill.Tacticas)
    SkillDefensa = UserList(UserIndex).Stats.UserSkills(eSkill.Defensa)
    
    'Esta usando un escudo ???
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).NoShield = 0 Then
            UserEvasion = UserEvasion + PoderEvasioEscudo

        End If

    End If
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + ((NpcPoderAtaque - UserEvasion) * 0.4)))
    
    NpcImpacto = (RandomNumber(1, 100) <= ProbExito)
    
    ' el usuario esta usando un escudo ???
    If UserList(UserIndex).Invent.EscudoEqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).Invent.EscudoEqpObjIndex).NoShield = 0 Then
        
            If Not NpcImpacto Then
                If SkillDefensa + SkillTacticas > 0 Then  'Evitamos división por cero
                    ' Chances are rounded
                    ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / (SkillDefensa + SkillTacticas)))
                Else
                    ProbRechazo = 10 'Si no tiene skills le dejamos el 10% mínimo

                End If
                
                Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
                    
                If Rechazo Then
                    'Se rechazo el ataque con el escudo
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_ESCUDO, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(UserIndex).Char.charindex))
                    Call WriteMultiMessage(UserIndex, eMessages.BlockedWithShieldUser) 'Call WriteBlockedWithShieldUser(UserIndex)
                    Call SubirSkill(UserIndex, eSkill.Defensa, True)
                Else
                    Call SubirSkill(UserIndex, eSkill.Defensa, False)

                End If

            End If

        End If

    End If

    '<EhFooter>
    Exit Function

NpcImpacto_Err:
    LogError Err.description & vbCrLf & "in NpcImpacto " & "at line " & Erl

    '</EhFooter>
End Function

Public Function CalcularDaño(ByVal UserIndex As Integer, Optional ByVal NpcIndex As Integer = 0) As Long

    '<EhHeader>
    On Error GoTo CalcularDaño_Err

    '</EhHeader>

    '***************************************************
    'Author: Unknown
    'Last Modification: 01/04/2010 (ZaMa)
    '01/04/2010: ZaMa - Modaño de wrestling.
    '01/04/2010: ZaMa - Agrego bonificadores de wrestling para los guantes.
    '***************************************************
    Dim DañoArma As Long

    Dim DañoUsuario As Long

    Dim Arma       As ObjData

    Dim ModifClase As Single

    Dim proyectil  As ObjData

    Dim DañoMaxArma As Long

    Dim DañoMinArma As Long

    Dim ObjIndex   As Integer
    
    ''sacar esto si no queremos q la matadracos mate el Dragon si o si
    Dim matoDragon As Boolean

    matoDragon = False
    
    With UserList(UserIndex)

        If .Invent.WeaponEqpObjIndex > 0 Then
            Arma = ObjData(.Invent.WeaponEqpObjIndex)
            
            ' Ataca a un npc?
            If NpcIndex > 0 Then
                If Arma.proyectil = 1 Then
                    ModifClase = Balance.ModClase(.Clase).DañoProyectiles
                    DañoArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                    DañoMaxArma = Arma.MaxHit
                    
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        DañoArma = DañoArma + RandomNumber(proyectil.MinHit, proyectil.MaxHit)

                        ' For some reason this isn't done...
                        'DañoMaxArma = DañoMaxArma + proyectil.MaxHIT
                    End If

                Else
                    ModifClase = Balance.ModClase(.Clase).DañoArmas
                    
                    If (.Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex) Then ' Usa la mata Dragones?
                        If Npclist(NpcIndex).NPCtype = DRAGON Then 'Ataca Dragon?

                            Dim Porc As Long

                            Porc = Int(Npclist(NpcIndex).Stats.MaxHp * 0.01)
                            
                            DañoArma = RandomNumber(Arma.MinHit, Porc)
                            DañoMaxArma = Arma.MaxHit
                            'matoDragon = True ''sacar esto si no queremos q la matadracos mate el Dragon si o si
                        Else ' Sino es Dragon daño es 1
                            DañoArma = 1
                            DañoMaxArma = 1

                        End If

                    Else
                        DañoArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                        DañoMaxArma = Arma.MaxHit

                    End If

                End If

            Else ' Ataca usuario

                If Arma.proyectil = 1 Then
                    ModifClase = Balance.ModClase(.Clase).DañoProyectiles
                    DañoArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                    DañoMaxArma = Arma.MaxHit
                     
                    If Arma.Municion = 1 Then
                        proyectil = ObjData(.Invent.MunicionEqpObjIndex)
                        DañoArma = DañoArma + RandomNumber(proyectil.MinHit, proyectil.MaxHit)

                        ' For some reason this isn't done...
                        'DañoMaxArma = DañoMaxArma + proyectil.MaxHIT
                    End If

                Else
                    ModifClase = Balance.ModClase(.Clase).DañoArmas
                    
                    If .Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Or .Invent.WeaponEqpObjIndex = EspadaDiablo Then
                        ModifClase = Balance.ModClase(.Clase).DañoArmas
                        DañoArma = 1 ' Si usa la espada mataDragones daño es 1
                        DañoMaxArma = 1
                    Else
                        DañoArma = RandomNumber(Arma.MinHit, Arma.MaxHit)
                        DañoMaxArma = Arma.MaxHit

                    End If

                End If

            End If

        Else
            ModifClase = Balance.ModClase(.Clase).DañoWrestling
            
            ' Daño sin guantes
            DañoMinArma = 4
            DañoMaxArma = 9
            
            ' Plus de guantes (en slot de anillo)
            ObjIndex = .Invent.AnilloEqpObjIndex

            If ObjIndex > 0 Then
                If ObjData(ObjIndex).Guante = 1 Then
                    DañoMinArma = DañoMinArma + ObjData(ObjIndex).MinHit
                    DañoMaxArma = DañoMaxArma + ObjData(ObjIndex).MaxHit

                End If

            End If
            
            DañoArma = RandomNumber(DañoMinArma, DañoMaxArma)
            
        End If
        
        DañoUsuario = RandomNumber(.Stats.MinHit, .Stats.MaxHit)
        
        ''sacar esto si no queremos q la matadracos mate el Dragon si o si
        If matoDragon Then
            CalcularDaño = Npclist(NpcIndex).Stats.MinHp + Npclist(NpcIndex).Stats.def
        Else
            CalcularDaño = (3 * DañoArma + ((DañoMaxArma / 5) * MaximoInt(0, .Stats.UserAtributos(eAtributos.Fuerza) - 15)) + DañoUsuario) * ModifClase

        End If

    End With

    '<EhFooter>
    Exit Function

CalcularDaño_Err:
    LogError Err.description & vbCrLf & "in CalcularDaño " & "at line " & Erl

    '</EhFooter>
End Function

Public Sub UserDañoNpc(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByRef Dead As Boolean, ByVal DosManos As Boolean)

    '***************************************************
    'Author: Unknown
    'Last Modification: 07/04/2010 (Pato)
    '25/01/2010: ZaMa - Agrego poder acuchillar npcs.
    '07/04/2010: ZaMa - Los asesinos apuñalan acorde al daño base sin descontar la defensa del npc.
    '07/04/2010: Pato - Si se mata al dragón en party se loguean los miembros de la misma.
    '11/07/2010: ZaMa - Ahora la defensa es solo ignorada para asesinos.
    '***************************************************
    '<EhHeader>
    On Error GoTo UserDañoNpc_Err

    '</EhHeader>

    Dim daño As Long

    Dim DañoBase As Long

    Dim PI        As Integer

    Dim Text      As String

    Dim i         As Integer
    
    Dim BoatIndex As Integer
    
    DañoBase = CalcularDaño(UserIndex, NpcIndex)
    
    'esta navegando? si es asi le sumamos el daño del barco
    If UserList(UserIndex).flags.Navegando = 1 Then
    
        BoatIndex = UserList(UserIndex).Invent.BarcoObjIndex

        If BoatIndex > 0 Then
            DañoBase = DañoBase + RandomNumber(ObjData(BoatIndex).MinHit, ObjData(BoatIndex).MaxHit)

        End If

    End If
    
    'Esta con gran poder?
    If Power.UserIndex = UserIndex Then
        DañoBase = DañoBase * 1.15

    End If
    
    ' ReliquiaDrag equipped
    ' If UserList(UserIndex).Invent.ReliquiaSlot > 0 Then
    'DañoBase = Effect_UpdatePorc(UserIndex, DañoBase)
    'End If
        
    Dim Y As Byte
        
    If DosManos Then ' Identificacion dos manos
        Y = 1

    End If
        
    With Npclist(NpcIndex)
               
        daño = DañoBase - .Stats.def
        
        If daño < 0 Then daño = 0
        
        'Call WriteMultiMessage(UserIndex, eMessages.UserHitNPC, daño)
        Call CalcularDarExp(UserIndex, NpcIndex, daño)
        Call Quests_AddNpc(UserIndex, NpcIndex, daño)
        .Stats.MinHp = .Stats.MinHp - daño
              
        Dim exito As Boolean
        
        If .Stats.MinHp > 0 Then
            Call DoGolpeCritico_Npcs(UserIndex, NpcIndex, daño)
            
            If PuedeAcuchillar(UserIndex) Then
                Call DoAcuchillar(UserIndex, NpcIndex, 0, daño)

            End If
            
            'Trata de apuñalar por la espalda al enemigo
            If PuedeApuñalar(UserIndex) Then
                
                UserList(UserIndex).DañoApu = daño
                
                ' La defensa se ignora solo en asesinos
                If UserList(UserIndex).Clase <> eClass.Assasin Then
                    DañoBase = daño

                End If
                
                Call DoApuñalar(UserIndex, NpcIndex, 0, DañoBase, exito)
                
            End If

        End If

        If Not exito Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y + Y, daño, d_DañoNpc))

        End If
            
        If .Stats.MinHp <= 0 Then

            ' Si era un Dragon perdemos la espada mataDragones
            If .NPCtype = DRAGON Then

                'Si tiene equipada la matadracos se la sacamos
                If UserList(UserIndex).Invent.WeaponEqpObjIndex = EspadaMataDragonesIndex Then
                    Call QuitarObjetos(EspadaMataDragonesIndex, 1, UserIndex)

                End If

            End If
            
            If UserList(UserIndex).MascotaIndex Then
                If Npclist(UserList(UserIndex).MascotaIndex).TargetNPC = NpcIndex Then
                    Npclist(UserList(UserIndex).MascotaIndex).TargetNPC = 0
                    Npclist(UserList(UserIndex).MascotaIndex).Movement = TipoAI.SigueAmo

                End If
            
            End If
            
            Dead = True
            Call MuereNpc(NpcIndex, UserIndex)

        End If
        
    End With

    '<EhFooter>
    Exit Sub

UserDañoNpc_Err:
    LogError Err.description & vbCrLf & "in UserDañoNpc " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub NpcDaño(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 03/06/2011 (Amraphen)
    '18/09/2010: ZaMa - Ahora se considera siempre la defensa del barco y el escudo.
    '03/06/2011: Amraphen - Agrego defensa adicional de armadura de segunda jerarquía.
    '***************************************************
    '<EhHeader>
    On Error GoTo NpcDaño_Err

    '</EhHeader>

    Dim daño As Integer

    Dim Lugar       As Integer

    Dim Obj         As ObjData
    
    Dim BoatDefense As Integer
        
    Dim HeadDefense As Integer

    Dim BodyDefense As Integer
    
    Dim BoatIndex   As Integer

    Dim HelmetIndex As Integer

    Dim ArmourIndex As Integer

    Dim ShieldIndex As Integer
    
    daño = RandomNumber(Npclist(NpcIndex).Stats.MinHit, Npclist(NpcIndex).Stats.MaxHit)
    
    With UserList(UserIndex)

        ' Navega?
        If .flags.Navegando = 1 Then
            ' En barca suma defensa
            BoatIndex = .Invent.BarcoObjIndex

            If BoatIndex > 0 Then
                Obj = ObjData(BoatIndex)
                BoatDefense = RandomNumber(Obj.MinDef, Obj.MaxDef)

            End If

        End If
        
        Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
        
        Select Case Lugar
        
            Case PartesCuerpo.bCabeza
            
                'Si tiene casco absorbe el golpe
                HelmetIndex = .Invent.CascoEqpObjIndex

                If HelmetIndex > 0 Then
                    Obj = ObjData(HelmetIndex)
                    HeadDefense = RandomNumber(Obj.MinDef, Obj.MaxDef)

                End If
                
            Case Else
                
                Dim MinDef As Integer

                Dim MaxDef As Integer
            
                'Si tiene armadura absorbe el golpe
                ArmourIndex = .Invent.ArmourEqpObjIndex

                If ArmourIndex > 0 Then
                    Obj = ObjData(ArmourIndex)
                    MinDef = Obj.MinDef
                    MaxDef = Obj.MaxDef

                End If
                
                'Si tiene armadura de segunda jerarquía obtiene un porcentaje de defensa adicional.
                If .Invent.FactionArmourEqpObjIndex > 0 Then
                    If .Faction.Status > 0 Then
                        MinDef = MinDef + InfoFaction(.Faction.Status).Range(.Faction.Range).MinDef
                        MaxDef = MaxDef + InfoFaction(.Faction.Status).Range(.Faction.Range).MaxDef

                    End If

                End If
                
                ' Si tiene escudo absorbe el golpe
                ShieldIndex = .Invent.EscudoEqpObjIndex

                If ShieldIndex > 0 Then
                    Obj = ObjData(ShieldIndex)
                    MinDef = MinDef + Obj.MinDef
                    MaxDef = MaxDef + Obj.MaxDef

                End If
                
                BodyDefense = RandomNumber(MinDef, MaxDef)
        
        End Select
        
        ' Daño final
        daño = daño - HeadDefense - BodyDefense - BoatDefense

        If daño < 1 Then daño = 1
        
        Call WriteMultiMessage(UserIndex, eMessages.NPCHitUser, Lugar, daño)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, daño, d_DañoNpc))
        
        If .flags.Privilegios And PlayerType.User Then .Stats.MinHp = .Stats.MinHp - daño
        
        If .flags.Meditando Then
            If daño > Fix(.Stats.MinHp / 100 * .Stats.UserAtributos(eAtributos.Inteligencia) * .Stats.UserSkills(eSkill.Magia) / 100 * 12 / (RandomNumber(0, 5) + 7)) Then
                .flags.Meditando = False
                .Char.FX = 0
                .Char.loops = 0
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, 0))

            End If

        End If
        
        'Muere el usuario
        If .Stats.MinHp <= 0 Then
            Call WriteMultiMessage(UserIndex, eMessages.NPCKillUser)  'Le informamos que ha muerto ;)
            
            'Si lo mato un guardia
            If Escriminal(UserIndex) Then
                If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                    Call RestarCriminalidad(UserIndex)

                End If

            End If
            
            If UserList(UserIndex).MascotaIndex > 0 Then
                Call FollowAmo(UserList(UserIndex).MascotaIndex)
            Else

                'Al matarlo no lo sigue mas
                With Npclist(NpcIndex)

                    If .flags.AIAlineacion = 0 Then
                        .Movement = .flags.OldMovement
                        .Hostile = .flags.OldHostil
                        .flags.AttackedBy = vbNullString
                        Npclist(NpcIndex).Target = 0

                    End If

                End With
                
            End If
            
            Call UserDie(UserIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

NpcDaño_Err:
    LogError Err.description & vbCrLf & "in NpcDaño " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub RestarCriminalidad(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo RestarCriminalidad_Err

    '</EhHeader>

    Dim EraCriminal As Boolean

    EraCriminal = Escriminal(UserIndex)
    
    With UserList(UserIndex).Reputacion

        If .BandidoRep > 0 Then
            .BandidoRep = .BandidoRep - vlASALTO

            If .BandidoRep < 0 Then .BandidoRep = 0
        ElseIf .LadronesRep > 0 Then
            .LadronesRep = .LadronesRep - (vlCAZADOR * 10)

            If .LadronesRep < 0 Then .LadronesRep = 0

        End If
    
        If EraCriminal And Not Escriminal(UserIndex) Then
        
            If UserList(UserIndex).Faction.Status = r_Caos Then
                Call mFacciones.Faction_RemoveUser(UserIndex)
            Else
                Call Guilds_CheckAlineation(UserIndex, a_Neutral)

            End If
            
            Call RefreshCharStatus(UserIndex)

        End If
    
    End With

    '<EhFooter>
    Exit Sub

RestarCriminalidad_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.SistemaCombate.RestarCriminalidad " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function NpcAtacaUser(ByVal NpcIndex As Integer, _
                             ByVal UserIndex As Integer, _
                             Optional ByVal Heading As eHeading = 0) As Boolean

    '*************************************************
    'Author: Unknown
    'Last modified: -
    '
    '*************************************************
    '<EhHeader>
    On Error GoTo NpcAtacaUser_Err

    '</EhHeader>

    With UserList(UserIndex)
            
        If .flags.AdminInvisible = 1 Then Exit Function
        If (Not .flags.Privilegios And PlayerType.User) <> 0 And Not .flags.AdminPerseguible Then Exit Function
        If Npclist(NpcIndex).NPCtype = eNPCType.Mascota Then Exit Function
        If Not CanAttackReyCastle(UserIndex, NpcIndex) Then Exit Function
        If (.flags.Mimetizado = 1) And (MapInfo(.Pos.Map).Pk) Then Exit Function ' // NUEVO
        If Npclist(NpcIndex).GiveResource.ObjIndex > 0 Then Exit Function ' Los npcs que se usan para extraer recursos no atacan a los usuarios.
        If Not IntervaloPuedeRecibirAtaqueCriature(UserIndex) Then Exit Function
            
        If Npclist(NpcIndex).CastleIndex > 0 Then
            If Castle(Npclist(NpcIndex).CastleIndex).GuildIndex = UserList(UserIndex).GuildIndex Then
                Exit Function
            
            End If

        End If

    End With
    
    With Npclist(NpcIndex)

        If ((MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).Blocked And 2 ^ (Heading - 1)) <> 0) Then
            NpcAtacaUser = False
            Exit Function

        End If

        ' Chequeos de mascotas/monturas
        If .MonturaIndex > 0 Then Exit Function
        
        ' El npc puede atacar ???
        If Intervalo_CriatureAttack(NpcIndex) Then
            NpcAtacaUser = True
            Call AllMascotasAtacanNPC(NpcIndex, UserIndex)
            
            If .Target = 0 Then .Target = UserIndex
            
            If UserList(UserIndex).flags.AtacadoPorNpc = 0 And UserList(UserIndex).flags.AtacadoPorUser = 0 Then
                UserList(UserIndex).flags.AtacadoPorNpc = NpcIndex

            End If

        Else
            NpcAtacaUser = False

            Exit Function

        End If
        
        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayEffect(.flags.Snd1, .Pos.X, .Pos.Y, .Char.charindex))

        End If
        
        'Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterChange(.Char.Body, .Char.BodyAttack, .Char.Head, .Char.Heading, .Char.CharIndex, .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, .Char.loops, .Char.CascoAnim))
    End With
    
    '  Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterAttackNpc(Npclist(NpcIndex).Char.CharIndex, Npclist(NpcIndex).Char.BodyAttack))
    
    If NpcImpacto(NpcIndex, UserIndex) Then

        With UserList(UserIndex)
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_IMPACTO, .Pos.X, .Pos.Y, .Char.charindex))
            
            If .flags.Meditando = False Then
                If .flags.Navegando = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.charindex, FXSANGRE, 0))

                End If

            End If
            
            Call NpcDaño(NpcIndex, UserIndex)
            Call WriteUpdateHP(UserIndex)
            
            '¿Puede envenenar?
            If Npclist(NpcIndex).Veneno = 1 Then Call NpcEnvenenarUser(UserIndex)

        End With
        
        Call SubirSkill(UserIndex, eSkill.Tacticas, False)
    Else
        Call WriteMultiMessage(UserIndex, eMessages.NPCSwing)
        Call SubirSkill(UserIndex, eSkill.Tacticas, True)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.charindex, FXSWING, 0))

    End If
    
    'Controla el nivel del usuario
    Call CheckUserLevel(UserIndex)
    '<EhFooter>
    Exit Function

NpcAtacaUser_Err:
    LogError Err.description & vbCrLf & "in NpcAtacaUser " & "at line " & Erl

    '</EhFooter>
End Function

Private Function NpcImpactoNpc(ByVal Atacante As Integer, _
                               ByVal Victima As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo NpcImpactoNpc_Err

    '</EhHeader>

    Dim PoderAtt  As Long

    Dim PoderEva  As Long

    Dim ProbExito As Long
    
    PoderAtt = Npclist(Atacante).PoderAtaque
    PoderEva = Npclist(Victima).PoderEvasion
    
    ' Chances are rounded
    ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtt - PoderEva) * 0.4))
    NpcImpactoNpc = (RandomNumber(1, 100) <= ProbExito)
    '<EhFooter>
    Exit Function

NpcImpactoNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.SistemaCombate.NpcImpactoNpc " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Sub NpcDañoNpc(ByVal Atacante As Integer, ByVal Victima As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo NpcDañoNpc_Err

    '</EhHeader>

    Dim daño As Integer

    Dim MasterIndex As Integer
    
    With Npclist(Atacante)
        daño = RandomNumber(.Stats.MinHit, .Stats.MaxHit)
        Npclist(Victima).Stats.MinHp = Npclist(Victima).Stats.MinHp - daño
        
        If .MaestroUser > 0 Then
            Call CalcularDarExp(.MaestroUser, Victima, daño)
            Call Quests_AddNpc(.MaestroUser, Victima, daño)

        End If
        
        If Npclist(Victima).Stats.MinHp < 1 Then
            .Movement = .flags.OldMovement
            .TargetNPC = 0
            
            If LenB(.flags.AttackedBy) <> 0 Then
                .Hostile = .flags.OldHostil

            End If
            
            MasterIndex = .MaestroUser

            If MasterIndex > 0 Then
                Call FollowAmo(Atacante)

            End If
            
            Call MuereNpc(Victima, MasterIndex)

        End If

    End With

    '<EhFooter>
    Exit Sub

NpcDañoNpc_Err:
    LogError Err.description & vbCrLf & "in NpcDañoNpc " & "at line " & Erl

    '</EhFooter>
End Sub

Public Sub NpcAtacaNpc(ByVal Atacante As Integer, _
                       ByVal Victima As Integer, _
                       Optional ByVal cambiarMOvimiento As Boolean = True)

    '<EhHeader>
    On Error GoTo NpcAtacaNpc_Err

    '</EhHeader>

    '*************************************************
    'Author: Unknown
    'Last modified: 01/03/2009
    '01/03/2009: ZaMa - Las mascotas no pueden atacar al rey si quedan pretorianos vivos.
    '23/05/2010: ZaMa - Ahora los elementales renuevan el tiempo de pertencia del npc que atacan si pertenece a su amo.
    '*************************************************
    
    Dim MasterIndex As Integer
    
    With Npclist(Atacante)
        
        'Es el Rey Preatoriano?
        If Npclist(Victima).NPCtype = eNPCType.Pretoriano Then
            If Not ClanPretoriano(Npclist(Victima).ClanIndex).CanAtackMember(Victima) Then
                Call WriteConsoleMsg(.MaestroUser, "Debes matar al resto del ejército antes de atacar al rey!", FontTypeNames.FONTTYPE_FIGHT)
                .TargetNPC = 0

                Exit Sub

            End If

        End If
        
        ' El npc puede atacar ???
        If Intervalo_CriatureAttack(Atacante) Then

            If cambiarMOvimiento Then
                Npclist(Victima).TargetNPC = Atacante
                Npclist(Victima).Movement = TipoAI.eNpcAtacaNpc

            End If

        Else

            Exit Sub

        End If
            
        Dim Heading As eHeading

        Heading = GetHeadingFromWorldPos(Npclist(Atacante).Pos, Npclist(Victima).Pos)

        If Heading <> Npclist(Atacante).Char.Heading And Npclist(Atacante).flags.Inmovilizado = 1 Then
            Npclist(Atacante).TargetNPC = 0
            Npclist(Atacante).Movement = TipoAI.MueveAlAzar
            Exit Sub

        End If
            
        Call ChangeNPCChar(Atacante, Npclist(Atacante).Char.Body, Npclist(Atacante).Char.Head, Heading)
        
        Heading = GetHeadingFromWorldPos(Npclist(Victima).Pos, Npclist(Atacante).Pos)

        If Heading <> Npclist(Victima).Char.Heading Then
            If Npclist(Victima).flags.Inmovilizado > 0 Then
                cambiarMOvimiento = False

            End If

        End If
                
        If cambiarMOvimiento Then
            Npclist(Victima).TargetNPC = Atacante
            Npclist(Victima).Movement = TipoAI.eNpcAtacaNpc

        End If

        If .flags.Snd1 > 0 Then
            Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayEffect(.flags.Snd1, .Pos.X, .Pos.Y, .Char.charindex))

        End If
        
        MasterIndex = .MaestroUser
        
        ' Tiene maestro?
        If MasterIndex > 0 Then

            ' Su maestro es dueño del npc al que ataca?
            If Npclist(Victima).Owner = MasterIndex Then
                ' Renuevo el timer de pertenencia
                Call IntervaloPerdioNpc(MasterIndex, True)

            End If

        End If
        
        If NpcImpactoNpc(Atacante, Victima) Then
            If Npclist(Victima).flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayEffect(Npclist(Victima).flags.Snd2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y, Npclist(Victima).Char.charindex))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayEffect(SND_IMPACTO2, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y, Npclist(Victima).Char.charindex))

            End If
        
            If MasterIndex > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayEffect(SND_IMPACTO, .Pos.X, .Pos.Y, .Char.charindex))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayEffect(SND_IMPACTO, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y, Npclist(Victima).Char.charindex))

            End If
            
            Call NpcDañoNpc(Atacante, Victima)
        Else

            If MasterIndex > 0 Then
                Call SendData(SendTarget.ToNPCArea, Atacante, PrepareMessagePlayEffect(SND_SWING, .Pos.X, .Pos.Y, .Char.charindex))
            Else
                Call SendData(SendTarget.ToNPCArea, Victima, PrepareMessagePlayEffect(SND_SWING, Npclist(Victima).Pos.X, Npclist(Victima).Pos.Y, Npclist(Victima).Char.charindex))

            End If

        End If

    End With

    '<EhFooter>
    Exit Sub

NpcAtacaNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.SistemaCombate.NpcAtacaNpc " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Public Sub AllMascotasAtacanNPC(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

    On Error GoTo AllMascotasAtacanNPC_Err
        
    Dim mascotaIdx As Integer

    mascotaIdx = UserList(UserIndex).MascotaIndex
            
    If mascotaIdx > 0 And mascotaIdx <> NpcIndex Then

        With Npclist(mascotaIdx)
                    
            .flags.AtacaNPCs = True

            If .flags.AtacaNPCs And .TargetNPC = 0 Then
                .TargetNPC = NpcIndex
                .Movement = TipoAI.eNpcAtacaNpc

            End If
            
        End With

    End If
        
    Exit Sub

AllMascotasAtacanNPC_Err:

End Sub

Public Function UsuarioAtacaNpc(ByVal UserIndex As Integer, _
                                ByVal NpcIndex As Integer, _
                                ByRef Dead As Boolean) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: 24/05/2011 (Amraphen)
    '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados por npcs cuando los atacan.
    '14/01/2010: ZaMa - Lo transformo en función, para que no se pierdan municiones al atacar targets inválidos.
    '13/02/2011: Amraphen - Ahora la stamina es quitada cuando efectivamente se ataca al NPC.
    '24/05/2011: Amraphen - Ahora se envía la animación del pj al golpear.
    '***************************************************
    '<EhHeader>
    On Error GoTo UsuarioAtacaNpc_Err

    '</EhHeader>

    Dim MunicionIndex As Integer
        
    Static DosManos   As Boolean
        
    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil > 0 Then
            MunicionIndex = UserList(UserIndex).Invent.MunicionEqpObjIndex

        End If
        
        'If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus > 0 Then
        'Call WriteConsoleMsg(UserIndex, "No puedes usar así estos objetos mágicos.", FontTypeNames.FONTTYPE_INFORED)
        'Exit Function
        'End If
            
        If Npclist(NpcIndex).RequiredWeapon > 0 And UserList(UserIndex).Invent.WeaponEqpObjIndex <> Npclist(NpcIndex).RequiredWeapon Then
            Call WriteConsoleMsg(UserIndex, "¡Solo puedes extraer el recurso teniendo equipado " & ObjData(Npclist(NpcIndex).RequiredWeapon).Name & ".", FontTypeNames.FONTTYPE_INFORED)
            Exit Function

        End If
            
        If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).DosManos > 0 Then
            DosManos = Not DosManos

        End If
            
    Else

        If Npclist(NpcIndex).RequiredWeapon > 0 Then
            Call WriteConsoleMsg(UserIndex, "¡Solo puedes extraer el recurso teniendo equipado " & ObjData(Npclist(NpcIndex).RequiredWeapon).Name & ".", FontTypeNames.FONTTYPE_INFORED)
            Exit Function

        End If
            
    End If
        
    If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then Exit Function

    Call NPCAtacado(NpcIndex, UserIndex)
    
    If UserImpactoNpc(UserIndex, NpcIndex) Then
        'Send animation
        Call SendCharacterSwing(UserIndex)
            
        If Npclist(NpcIndex).flags.Snd2 > 0 Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayEffect(Npclist(NpcIndex).flags.Snd2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, Npclist(NpcIndex).Char.charindex))
        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_IMPACTO2, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, Npclist(NpcIndex).Char.charindex))

        End If
            
        Call UserDañoNpc(UserIndex, NpcIndex, Dead, DosManos)
            
        If MunicionIndex > 0 Then
            If ObjData(MunicionIndex).VictimAnim > 0 Then
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFX(Npclist(NpcIndex).Char.charindex, ObjData(MunicionIndex).VictimAnim, 0))
            Else
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFX(Npclist(NpcIndex).Char.charindex, FXSANGRE, 0))

            End If
            
            If RandomNumber(1, 100) <= ObjData(MunicionIndex).Incineracion And Npclist(NpcIndex).flags.Incinerado = 0 Then
                Npclist(NpcIndex).Contadores.Incinerado = IntervaloFrio
                Npclist(NpcIndex).flags.Incinerado = 1
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayEffect(eSound.sFogata, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, Npclist(NpcIndex).Char.charindex, True))
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFX(Npclist(NpcIndex).Char.charindex, FXIDs.FX_INCINERADO, -1))

            End If

        Else
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFX(Npclist(NpcIndex).Char.charindex, FXSANGRE, 0))

        End If
        
    Else
        'Send animation
        Call SendCharacterSwing(UserIndex)
            
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_SWING, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(UserIndex).Char.charindex))
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(Npclist(NpcIndex).Char.charindex, FXSWING, 0))

        'Call WriteMultiMessage(UserIndex, eMessages.UserSwing)
    End If
    
    'Quitamos stamina
    Call QuitarSta(UserIndex, RandomNumber(1, 10))
    
    ' Reveló su condición de usuario al atacar, los npcs lo van a atacar
    UserList(UserIndex).flags.Ignorado = False
    
    UsuarioAtacaNpc = True

    '<EhFooter>
    Exit Function

UsuarioAtacaNpc_Err:
    LogError Err.description & vbCrLf & "in UsuarioAtacaNpc " & "at line " & Erl

    '</EhFooter>
End Function

Public Sub UsuarioAtaca(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 24/05/2011 (Amraphen)
    '13/02/2011: Amraphen - Ahora se quita la stamina en el sub UsuarioAtacaNPC.
    '24/05/2011: Amraphen - Ahora se envía la animación del pj al golpear.
    '***************************************************
    '<EhHeader>
    On Error GoTo UsuarioAtaca_Err

    '</EhHeader>

    Dim Index     As Integer

    Dim DosManos  As Boolean
    
    Dim AttackPos As WorldPos
    
    'Check bow's interval
    If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
        
    'Check Spell-Attack interval
    If Not IntervaloPermiteMagiaGolpe(UserIndex, False) Then Exit Sub
            
    'Check Attack interval
    If Not IntervaloPermiteAtacar(UserIndex) Then Exit Sub

    With UserList(UserIndex)

        'Chequeamos que tenga por lo menos 10 de stamina.
        If .Stats.MinSta < 10 Then
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(UserIndex, "Estás muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Estás muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)

            End If

            Exit Sub

        End If
        
        'Chequeamos que no esté desnudo
        'If .flags.Desnudo Then
        'If .Genero = eGenero.Hombre Then
        'Call WriteConsoleMsg(UserIndex, "No puedes atacar si estás desnudo.", FontTypeNames.FONTTYPE_INFO)
        'Else
        'Call WriteConsoleMsg(UserIndex, "No puedes atacar si estás desnuda.", FontTypeNames.FONTTYPE_INFO)
        'End If
        'Exit Sub
        'End If
        
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).QuitaEnergia > 0 Then
                If .Stats.MinSta >= ObjData(.Invent.WeaponEqpObjIndex).QuitaEnergia Then
                    Call QuitarSta(UserIndex, ObjData(.Invent.WeaponEqpObjIndex).QuitaEnergia)
                Else

                    If .Genero = eGenero.Hombre Then
                        Call WriteConsoleMsg(UserIndex, "Estás muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "Estás muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)

                    End If
                    
                    Exit Sub

                End If

            End If
            
            If ObjData(.Invent.WeaponEqpObjIndex).DosManos = 1 Then
                DosManos = True

            End If

        End If
        
        AttackPos = .Pos
        Call HeadtoPos(.Char.Heading, AttackPos)
        
        'Exit if not legal
        If AttackPos.X < XMinMapSize Or AttackPos.X > XMaxMapSize Or AttackPos.Y <= YMinMapSize Or AttackPos.Y > YMaxMapSize Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_SWING, .Pos.X, .Pos.Y, .Char.charindex))

            Exit Sub

        End If
        
        Index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).UserIndex
        
        'Look for user
        If Index > 0 Then
            
            If UsuarioAtacaUsuario(UserIndex, Index) Then
                If DosManos And UserList(Index).flags.Muerto = 0 Then
                    Call UsuarioAtacaUsuario(UserIndex, Index)

                End If

            End If
            
            Call WriteUpdateUserStats(UserIndex)
            Call WriteUpdateUserStats(Index)

            Exit Sub

        End If
        
        Index = MapData(AttackPos.Map, AttackPos.X, AttackPos.Y).NpcIndex
        
        'Look for NPC
        If Index > 0 Then
            If Npclist(Index).Attackable Then
                If Npclist(Index).MaestroUser > 0 And MapInfo(Npclist(Index).Pos.Map).Pk = False Then
                    Call WriteConsoleMsg(UserIndex, "No puedes atacar mascotas en zona segura.", FontTypeNames.FONTTYPE_FIGHT)

                    Exit Sub

                End If
                
                Dim Dead As Boolean
                
                If UsuarioAtacaNpc(UserIndex, Index, Dead) Then
                    If DosManos And Not Dead Then
                        Call UsuarioAtacaNpc(UserIndex, Index, False)

                    End If

                End If

            Else
                Call WriteConsoleMsg(UserIndex, "No puedes atacar a este NPC.", FontTypeNames.FONTTYPE_FIGHT)

            End If
            
            Call WriteUpdateUserStats(UserIndex)
            
            Exit Sub

        End If
        
        'Send animation
        Call SendCharacterSwing(UserIndex)
        
        'Send sound
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(SND_SWING, .Pos.X, .Pos.Y, .Char.charindex))
        
        Call WriteUpdateUserStats(UserIndex)
        
        If .Counters.Trabajando Then .Counters.Trabajando = .Counters.Trabajando - 1
            
        If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1

    End With

    '<EhFooter>
    Exit Sub

UsuarioAtaca_Err:
    LogError Err.description & vbCrLf & "in UsuarioAtaca " & "at line " & Erl

    '</EhFooter>
End Sub

Public Function UsuarioImpacto(ByVal AtacanteIndex As Integer, _
                               ByVal VictimaIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: 21/05/2010
    '21/05/2010: ZaMa - Evito division por cero.
    '***************************************************
    '<EhHeader>
    On Error GoTo UsuarioImpacto_Err

    '</EhHeader>

    On Error GoTo ErrHandler

    Dim ProbRechazo            As Long

    Dim Rechazo                As Boolean

    Dim ProbExito              As Long

    Dim PoderAtaque            As Long

    Dim UserPoderEvasion       As Long

    Dim UserPoderEvasionEscudo As Long

    Dim Arma                   As Integer

    Dim SkillTacticas          As Long

    Dim SkillDefensa           As Long

    Dim ProbEvadir             As Long

    Dim Skill                  As eSkill
    
    With UserList(VictimaIndex)
    
        SkillTacticas = .Stats.UserSkills(eSkill.Tacticas)
        SkillDefensa = .Stats.UserSkills(eSkill.Defensa)
        
        Arma = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
        
        'Calculamos el poder de evasion...
        UserPoderEvasion = PoderEvasion(VictimaIndex)
        
        If .Invent.EscudoEqpObjIndex > 0 Then
            If ObjData(.Invent.EscudoEqpObjIndex).NoShield = 1 Then
                UserPoderEvasionEscudo = 0
            Else
                UserPoderEvasionEscudo = PoderEvasionEscudo(VictimaIndex)
                UserPoderEvasion = UserPoderEvasion + UserPoderEvasionEscudo

            End If

        Else
            UserPoderEvasionEscudo = 0

        End If
        
        'Esta usando un arma ???
        If UserList(AtacanteIndex).Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(Arma).proyectil = 1 Then
                PoderAtaque = PoderAtaqueProyectil(AtacanteIndex)
                Skill = eSkill.Proyectiles
            Else
                PoderAtaque = PoderAtaqueArma(AtacanteIndex)
                Skill = eSkill.Armas

            End If

        End If
        
        ' Chances are rounded
        ProbExito = MaximoInt(10, MinimoInt(90, 50 + (PoderAtaque - UserPoderEvasion) * 0.4))
        
        ' Se reduce la evasion un 25%
        If .flags.Meditando Then
            ProbEvadir = (100 - ProbExito) * 0.75
            ProbExito = MinimoInt(90, 100 - ProbEvadir)

        End If
        
        UsuarioImpacto = (RandomNumber(1, 100) <= ProbExito)
        
        ' el usuario esta usando un escudo ???
        If .Invent.EscudoEqpObjIndex > 0 Then
            If ObjData(.Invent.EscudoEqpObjIndex).NoShield = 0 Then

                'Fallo ???
                If Not UsuarioImpacto Then
                    
                    Dim SumaSkills As Integer
                    
                    ' Para evitar division por 0
                    SumaSkills = MaximoInt(1, SkillDefensa + SkillTacticas)
                    
                    ' Chances are rounded
                    ProbRechazo = MaximoInt(10, MinimoInt(90, 100 * SkillDefensa / SumaSkills))
                    Rechazo = (RandomNumber(1, 100) <= ProbRechazo)
    
                    If Rechazo Then
                        'Se rechazo el ataque con el escudo
                        Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayEffect(SND_ESCUDO, .Pos.X, .Pos.Y, .Char.charindex))
                          
                        Call WriteMultiMessage(AtacanteIndex, eMessages.BlockedWithShieldother)
                        Call WriteMultiMessage(VictimaIndex, eMessages.BlockedWithShieldUser)
                        
                        Call SubirSkill(VictimaIndex, eSkill.Defensa, True)
                    Else
                        Call SubirSkill(VictimaIndex, eSkill.Defensa, False)

                    End If

                End If

            End If

        End If
        
        If Not UsuarioImpacto Then
            Call SubirSkill(AtacanteIndex, Skill, False)

        End If
        
        Call FlushBuffer(VictimaIndex)

    End With
    
    Exit Function
    
ErrHandler:

    Dim AtacanteNick As String

    Dim VictimaNick  As String
    
    If AtacanteIndex > 0 Then AtacanteNick = UserList(AtacanteIndex).Name
    If VictimaIndex > 0 Then VictimaNick = UserList(VictimaIndex).Name
    
    Call LogError("Error en UsuarioImpacto. Error " & Err.number & " : " & Err.description & " AtacanteIndex: " & AtacanteIndex & " Nick: " & AtacanteNick & " VictimaIndex: " & VictimaIndex & " Nick: " & VictimaNick)
    '<EhFooter>
    Exit Function

UsuarioImpacto_Err:
    LogError Err.description & vbCrLf & "in UsuarioImpacto " & "at line " & Erl

    '</EhFooter>
End Function

Public Function UsuarioAtacaUsuario(ByVal AtacanteIndex As Integer, _
                                    ByVal VictimaIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 24/05/2011 (Amraphen)
    '14/01/2010: ZaMa - Lo transformo en función, para que no se pierdan municiones al atacar targets
    '                    inválidos, y evitar un doble chequeo innecesario
    '24/05/2011: Amraphen - Ahora se envía la animación del user al golpear.
    '***************************************************

    On Error GoTo ErrHandler

    Dim MunicionIndex As Integer
    
    With UserList(AtacanteIndex)

        If .Invent.WeaponEqpObjIndex > 0 Then
            'If ObjData(.Invent.WeaponEqpObjIndex).StaffDamageBonus > 0 Then
            'Call WriteConsoleMsg(AtacanteIndex, "No puedes usar así estos objetos mágicos.", FontTypeNames.FONTTYPE_INFORED)
            'Exit Function
            ' End If
            
            If ObjData(.Invent.WeaponEqpObjIndex).proyectil > 0 Then
                MunicionIndex = .Invent.MunicionEqpObjIndex

            End If
          
        Else
            Call WriteConsoleMsg(AtacanteIndex, "Necesitas equipar algun tipo de arma que te ayude a luchar.", FontTypeNames.FONTTYPE_INFORED)
            Exit Function

        End If
        
        If Not PuedeAtacar(AtacanteIndex, VictimaIndex) Then Exit Function
        
        If Distancia(.Pos, UserList(VictimaIndex).Pos) > MAXDISTANCIAARCO Then
            Call WriteConsoleMsg(AtacanteIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)

            Exit Function

        End If
        
        Call UsuarioAtacadoPorUsuario(AtacanteIndex, VictimaIndex)
        
        If UsuarioImpacto(AtacanteIndex, VictimaIndex) Then
            'Send animation
            Call SendCharacterSwing(AtacanteIndex)
        
            Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayEffect(SND_IMPACTO, .Pos.X, .Pos.Y, .Char.charindex))
            
            If UserList(VictimaIndex).flags.Navegando = 0 Then
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.charindex, FXSANGRE, 0))

            End If
            
            If MunicionIndex > 0 Then
                If ObjData(MunicionIndex).VictimAnim > 0 Then
                    Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.charindex, ObjData(MunicionIndex).VictimAnim, 0))
                Else
                    Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.charindex, FXSANGRE, 0))

                End If
                
                If RandomNumber(1, 100) <= ObjData(MunicionIndex).Incineracion And UserList(VictimaIndex).flags.Incinerado = 0 Then
                    UserList(VictimaIndex).Counters.Incinerado = IntervaloFrio
                    UserList(VictimaIndex).flags.Incinerado = 1
                    Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessagePlayEffect(eSound.sFogata, UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y, UserList(VictimaIndex).Char.charindex, True))
                    Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.charindex, FXIDs.FX_INCINERADO, -1))

                End If
                
            Else
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.charindex, FXSANGRE, 0))

            End If
            
            'Pablo (ToxicWaste): Guantes de Hurto del Bandido en acción
            ' If .Clase = eClass.Bandit Then
            ' Call DoDesequipar(AtacanteIndex, VictimaIndex)
                
            'y ahora, el ladrón puede llegar a paralizar con el golpe.
            If .Clase = eClass.Thief Then
                Call DoHandInmo(AtacanteIndex, VictimaIndex)

            End If
            
            Call SubirSkill(VictimaIndex, eSkill.Tacticas, False)
            Call UserDañoUser(AtacanteIndex, VictimaIndex)
        Else

            ' Invisible admins doesn't make sound to other clients except itself
            If .flags.AdminInvisible = 1 Then
                Call SendData(ToOne, AtacanteIndex, PrepareMessagePlayEffect(SND_SWING, .Pos.X, .Pos.Y, .Char.charindex))
                
            Else
                'Send animation
                Call SendCharacterSwing(AtacanteIndex)
                
                Call SendData(SendTarget.ToPCArea, AtacanteIndex, PrepareMessagePlayEffect(SND_SWING, .Pos.X, .Pos.Y, .Char.charindex))
                Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateFX(UserList(VictimaIndex).Char.charindex, FXSWING, 0))

            End If
            
            ' Call WriteMultiMessage(AtacanteIndex, eMessages.UserSwing)
            Call WriteMultiMessage(VictimaIndex, eMessages.UserAttackedSwing, AtacanteIndex)
            Call SubirSkill(VictimaIndex, eSkill.Tacticas, True)

        End If
        
        If .Clase = eClass.Thief Then Call Desarmar(AtacanteIndex, VictimaIndex)

    End With
    
    UsuarioAtacaUsuario = True
    
    Exit Function
    
ErrHandler:
    Call LogError("Error en UsuarioAtacaUsuario. Error " & Err.number & " : " & Err.description)

End Function

Public Sub UserDañoUser(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 03/06/2011 (Amraphen)
    '12/01/2010: ZaMa - Implemento armas arrojadizas y probabilidad de acuchillar.
    '11/03/2010: ZaMa - Ahora no cuenta la muerte si estaba en estado atacable, y no se vuelve criminal.
    '18/09/2010: ZaMa - Ahora se cosidera la defensa de los barcos siempre.
    '03/06/2011: Amraphen - Agrego defensa adicional de armadura de segunda jerarquía.
    '***************************************************
    
    On Error GoTo ErrHandler

    Dim daño As Long

    Dim Lugar          As Byte

    Dim Obj            As ObjData
    
    Dim BoatDefense    As Integer
    
    Dim BodyDefense    As Integer

    Dim HeadDefense    As Integer

    Dim WeaponBoost    As Integer
    
    Dim BoatIndex      As Integer

    Dim WeaponIndex    As Integer

    Dim HelmetIndex    As Integer

    Dim ArmourIndex    As Integer

    Dim ShieldIndex    As Integer
    
    Dim BarcaIndex     As Integer

    Dim ArmaIndex      As Integer

    Dim CascoIndex     As Integer

    Dim ArmaduraIndex  As Integer
    
    Dim FactionDefense As Integer
    
    daño = CalcularDaño(AtacanteIndex)
    
    Call UserEnvenena(AtacanteIndex, VictimaIndex)
    
    With UserList(AtacanteIndex)
        
        ' ReliquiaDrag equipped
        'If .Invent.ReliquiaSlot > 0 Then
        'Daño = Effect_UpdatePorc(AtacanteIndex, Daño)
        'End If
        
        ' ReliquiaDrag equipped
        'If UserList(VictimaIndex).Invent.ReliquiaSlot > 0 Then
        'Daño = Effect_UpdatePorc(VictimaIndex, Daño)
        'End If
    
        ' Aumento de Daño por bonos
        If .flags.SelectedBono > 0 Then
            If ObjData(.flags.SelectedBono).BonoArmas > 0 Then
                daño = daño * ObjData(.flags.SelectedBono).BonoArmas
            
            End If

        End If
        
        ' Bonus faccionario
        If UserList(VictimaIndex).Faction.Status > 0 Then
            If UserList(VictimaIndex).Invent.ArmourEqpObjIndex > 0 Then
                If ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex).Caos > 0 Or ObjData(UserList(VictimaIndex).Invent.ArmourEqpObjIndex).Real > 0 Then
                    FactionDefense = RandomNumber(InfoFaction(UserList(VictimaIndex).Faction.Status).Range(UserList(VictimaIndex).Faction.Range).MinDef, InfoFaction(UserList(VictimaIndex).Faction.Status).Range(UserList(VictimaIndex).Faction.Range).MaxDef)
                    daño = daño - FactionDefense

                End If

            End If
            
        End If
        
        ' Aumento de daño por barca (atacante)
        If .flags.Navegando = 1 Then
            
            BoatIndex = .Invent.BarcoObjIndex
            
            If BoatIndex > 0 Then
                Obj = ObjData(BoatIndex)
                daño = daño + RandomNumber(Obj.MinHit, Obj.MaxHit)

            End If
            
        End If
        
        ' Aumento de defensa por barca (victima)
        If UserList(VictimaIndex).flags.Navegando = 1 Then
            
            BoatIndex = UserList(VictimaIndex).Invent.BarcoObjIndex
            
            If BoatIndex > 0 Then
                Obj = ObjData(BoatIndex)
                BoatDefense = RandomNumber(Obj.MinDef, Obj.MaxDef)

            End If
            
        End If
        
        ' Aumento de DAÑO por Gran poder (Victima)
        If Power.UserIndex = AtacanteIndex Then
            daño = daño * 1.05

        End If
        
        If .Pos.Map <> 130 And .Pos.Map <> 131 And .Pos.Map <> 132 Then

            ' Daño físico para los clanes con CASTILLO ESTE
            If Castle_CheckBonus(.GuildIndex, eCastle.CASTLE_EAST) Then
                daño = daño * 1.02

            End If

        End If
        
        ' Refuerzo arma (atacante)
        WeaponIndex = .Invent.WeaponEqpObjIndex

        If WeaponIndex > 0 Then
            WeaponBoost = ObjData(WeaponIndex).Refuerzo

        End If
        
        ' Suerte de la cabeza para los asesinos
        If .Clase = eClass.Assasin Then
            If RandomNumber(1, 10) <= 1 Then
                Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
            Else
            
                Lugar = RandomNumber(PartesCuerpo.bPiernaIzquierda, PartesCuerpo.bTorso)

            End If

        Else
            Lugar = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)

        End If
        
        Select Case Lugar
        
            Case PartesCuerpo.bCabeza
            
                'Si tiene casco absorbe el golpe
                HelmetIndex = UserList(VictimaIndex).Invent.CascoEqpObjIndex

                If HelmetIndex > 0 Then
                    Obj = ObjData(HelmetIndex)
                    HeadDefense = RandomNumber(Obj.MinDef, Obj.MaxDef)

                End If
            
            Case Else
                
                Dim MinDef As Integer

                Dim MaxDef As Integer
                
                'Si tiene armadura absorbe el golpe
                ArmourIndex = UserList(VictimaIndex).Invent.ArmourEqpObjIndex

                If ArmourIndex > 0 Then
                    Obj = ObjData(ArmourIndex)
                    MinDef = Obj.MinDef
                    MaxDef = Obj.MaxDef

                End If

                'Si tiene armadura de segunda jerarquía obtiene un porcentaje de defensa adicional.
                If UserList(VictimaIndex).Invent.FactionArmourEqpObjIndex > 0 Then
                    If UserList(VictimaIndex).Faction.Status > 0 Then
                        MinDef = MinDef + InfoFaction(UserList(VictimaIndex).Faction.Status).Range(UserList(VictimaIndex).Faction.Range).MinDef
                        MaxDef = MaxDef + InfoFaction(UserList(VictimaIndex).Faction.Status).Range(UserList(VictimaIndex).Faction.Range).MaxDef

                    End If

                End If
                
                ' Si tiene escudo, tambien absorbe el golpe
                ShieldIndex = UserList(VictimaIndex).Invent.EscudoEqpObjIndex

                If ShieldIndex > 0 Then
                    Obj = ObjData(ShieldIndex)
                    MinDef = MinDef + Obj.MinDef
                    MaxDef = MaxDef + Obj.MaxDef

                End If
                
                BodyDefense = RandomNumber(MinDef, MaxDef)
                
        End Select
        
        daño = daño + WeaponBoost - HeadDefense - BodyDefense - BoatDefense

        If daño < 0 Then daño = 1
        
        Dim Valid As Boolean: Valid = True

        If .flags.SlotEvent > 0 Then
            If Events(.flags.SlotEvent).Modality = eModalityEvent.DeathMatch Then
                Valid = False

            End If

        End If
        
        If Valid Then
            Call WriteMultiMessage(AtacanteIndex, eMessages.UserHittedUser, UserList(VictimaIndex).Char.charindex, Lugar, daño)
            Call WriteMultiMessage(VictimaIndex, eMessages.UserHittedByUser, .Char.charindex, Lugar, daño)

        End If
        
        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageUpdateControlPotas(UserList(VictimaIndex).Char.charindex, UserList(VictimaIndex).Stats.MinHp, UserList(VictimaIndex).Stats.MaxHp, UserList(VictimaIndex).Stats.MinMan, UserList(VictimaIndex).Stats.MaxMan))
        
        UserList(VictimaIndex).Stats.MinHp = UserList(VictimaIndex).Stats.MinHp - daño
        
        If UserList(AtacanteIndex).flags.SlotEvent > 0 Then
            Events_Add_Damage UserList(AtacanteIndex).flags.SlotEvent, UserList(AtacanteIndex).flags.SlotUserEvent, daño

        End If
        
        Dim exito As Boolean
        
        If .flags.Hambre = 0 And .flags.Sed = 0 Then

            'Si usa un arma quizas suba "Combate con armas"
            If WeaponIndex > 0 Then
                If ObjData(WeaponIndex).proyectil Then
                    'es un Arco. Sube Armas a Distancia
                    Call SubirSkill(AtacanteIndex, eSkill.Proyectiles, True)
                    
                    ' Si acuchilla
                    If PuedeAcuchillar(AtacanteIndex) Then
                        Call DoAcuchillar(AtacanteIndex, 0, VictimaIndex, daño)

                    End If

                Else
                    'Sube combate con armas.
                    Call SubirSkill(AtacanteIndex, eSkill.Armas, True)

                End If

            End If
                    
            'Trata de apuñalar por la espalda al enemigo
            If PuedeApuñalar(AtacanteIndex) Then
                UserList(AtacanteIndex).DañoApu = daño
                Call DoApuñalar(AtacanteIndex, 0, VictimaIndex, daño, exito)

            End If
            
            'e intenta dar un golpe crítico [Pablo (ToxicWaste)]
            Call DoGolpeCritico(AtacanteIndex, 0, VictimaIndex, daño)

        End If
        
        If Not exito Then
            Call SendData(SendTarget.ToPCArea, VictimaIndex, PrepareMessageCreateDamage(UserList(VictimaIndex).Pos.X, UserList(VictimaIndex).Pos.Y, daño, d_DañoUser))

        End If
        
        If UserList(VictimaIndex).Stats.MinHp <= 0 Then
            
            ' No cuenta la muerte si estaba en estado atacable
            If UserList(VictimaIndex).flags.AtacablePor <> AtacanteIndex Then
                'Store it!
                'Call Statistics.StoreFrag(AtacanteIndex, VictimaIndex)
                
                Call ContarMuerte(VictimaIndex, AtacanteIndex)

            End If
            
            If .MascotaIndex Then
                If Npclist(.MascotaIndex).Target = VictimaIndex Then
                    Npclist(.MascotaIndex).Target = 0
                    Call FollowAmo(.MascotaIndex)

                End If

            End If
            
            Call ActStats(VictimaIndex, AtacanteIndex)
            Call UserDie(VictimaIndex, AtacanteIndex)
        Else
            'Está vivo - Actualizamos el HP
            Call WriteUpdateHP(VictimaIndex)

        End If

    End With
    
    'Controla el nivel del usuario
    Call CheckUserLevel(AtacanteIndex)
    
    Call FlushBuffer(VictimaIndex)
    
    Exit Sub
    
ErrHandler:

    Dim AtacanteNick As String

    Dim VictimaNick  As String
    
    If AtacanteIndex > 0 Then AtacanteNick = UserList(AtacanteIndex).Name
    If VictimaIndex > 0 Then VictimaNick = UserList(VictimaIndex).Name
    
    Call LogError("Error en UserDañoUser. Error " & Err.number & " : " & Err.description & " AtacanteIndex: " & AtacanteIndex & " Nick: " & AtacanteNick & " VictimaIndex: " & VictimaIndex & " Nick: " & VictimaNick)

End Sub

Sub UsuarioAtacadoPorUsuario(ByVal AttackerIndex As Integer, ByVal VictimIndex As Integer)

    '***************************************************
    'Autor: Unknown
    'Last Modification: 05/05/2010
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    '10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
    '05/05/2010: ZaMa - Ahora no suma puntos de bandido al atacar a alguien en estado atacable.
    '***************************************************
    '<EhHeader>
    On Error GoTo UsuarioAtacadoPorUsuario_Err

    '</EhHeader>

    If UserList(VictimIndex).flags.Meditando Then
        UserList(VictimIndex).flags.Meditando = False
        UserList(VictimIndex).Char.FX = 0
        Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageMeditateToggle(UserList(VictimIndex).Char.charindex, 0))

    End If
        
    If TriggerZonaPelea(AttackerIndex, VictimIndex) = TRIGGER6_PERMITE Then Exit Sub
    
    Dim EraCriminal       As Boolean

    Dim VictimaEsAtacable As Boolean
    
    If Not Escriminal(AttackerIndex) Then
        If Not Escriminal(VictimIndex) Then
            ' Si la victima no es atacable por el agresor, entonces se hace pk
            VictimaEsAtacable = UserList(VictimIndex).flags.AtacablePor = AttackerIndex
                  
            If MapInfo(UserList(AttackerIndex).Pos.Map).FreeAttack Then
                VictimaEsAtacable = True

            End If
                  
            If Not VictimaEsAtacable Then Call VolverCriminal(AttackerIndex)

        End If

    End If
    
    With UserList(VictimIndex)

        If UserList(VictimIndex).flags.Meditando Then
            UserList(VictimIndex).flags.Meditando = False
            UserList(VictimIndex).Char.FX = 0
            Call SendData(SendTarget.ToPCArea, VictimIndex, PrepareMessageMeditateToggle(UserList(VictimIndex).Char.charindex, 0))

        End If

    End With
    
    EraCriminal = Escriminal(AttackerIndex)
    
    ' Si ataco a un atacable, no suma puntos de bandido
    If Not VictimaEsAtacable Then

        With UserList(AttackerIndex).Reputacion

            If Not Escriminal(VictimIndex) Then
                .BandidoRep = .BandidoRep + vlASALTO

                If .BandidoRep > MAXREP Then .BandidoRep = MAXREP
                
                .NobleRep = .NobleRep * 0.5

                If .NobleRep < 0 Then .NobleRep = 0
            Else
                .NobleRep = .NobleRep + vlNoble

                If .NobleRep > MAXREP Then .NobleRep = MAXREP

            End If

        End With

    End If
    
    If Escriminal(AttackerIndex) Then
    
        If UserList(AttackerIndex).Faction.Status = r_Armada Then
            Call mFacciones.Faction_RemoveUser(AttackerIndex)
        Else

            If Not EraCriminal And Escriminal(AttackerIndex) Then
                Call Guilds_CheckAlineation(AttackerIndex, a_Neutral)

            End If

        End If
        
        If Not EraCriminal Then Call RefreshCharStatus(AttackerIndex)
    ElseIf EraCriminal Then
        Call RefreshCharStatus(AttackerIndex)

    End If
    
    Call AllMascotasAtacanUser(AttackerIndex, VictimIndex)
    Call AllMascotasAtacanUser(VictimIndex, AttackerIndex)
    
    'Si la victima esta saliendo se cancela la salida
    Call CancelExit(VictimIndex)
    Call FlushBuffer(VictimIndex)
    '<EhFooter>
    Exit Sub

UsuarioAtacadoPorUsuario_Err:
    LogError Err.description & vbCrLf & "in UsuarioAtacadoPorUsuario " & "at line " & Erl

    '</EhFooter>
End Sub

Sub AllMascotasAtacanUser(ByVal victim As Integer, ByVal Maestro As Integer)

    '<EhHeader>
    On Error GoTo AllMascotasAtacanUser_Err

    '</EhHeader>

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    'Reaccion de las mascotas
    Dim iCount       As Integer

    Dim MascotaIndex As Integer
        
    MascotaIndex = UserList(Maestro).MascotaIndex
        
    If MascotaIndex Then
        If Not Npclist(MascotaIndex).Entrenable Then
            Npclist(MascotaIndex).flags.AttackedBy = UserList(victim).Name
            Npclist(MascotaIndex).Movement = TipoAI.NpcDefensa
            Npclist(MascotaIndex).Hostile = 1
            Npclist(MascotaIndex).Target = victim

        End If

    End If

    '<EhFooter>
    Exit Sub

AllMascotasAtacanUser_Err:
    LogError Err.description & vbCrLf & "in AllMascotasAtacanUser " & "at line " & Erl

    '</EhFooter>
End Sub

Public Function PuedeAtacar(ByVal AttackerIndex As Integer, _
                            ByVal VictimIndex As Integer) As Boolean

    '<EhHeader>
    On Error GoTo PuedeAtacar_Err

    '</EhHeader>

    '***************************************************
    'Autor: Unknown
    'Last Modification: 02/04/2010
    'Returns true if the AttackerIndex is allowed to attack the VictimIndex.
    '24/01/2007 Pablo (ToxicWaste) - Ordeno todo y agrego situacion de Defensa en ciudad Armada y Caos.
    '24/02/2009: ZaMa - Los usuarios pueden atacarse entre si.
    '02/04/2010: ZaMa - Los armadas no pueden atacar nunca a los ciudas, salvo que esten atacables.
    '***************************************************

    'MUY importante el orden de estos "IF"...
        
    'Estas muerto no podes atacar
    If UserList(AttackerIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(AttackerIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False

        Exit Function

    End If
    
    'No podes atacar a alguien muerto
    If UserList(VictimIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a un espíritu.", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False

        Exit Function

    End If
    
    ' No podes atacar si estas en consulta
    If UserList(AttackerIndex).flags.EnConsulta Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)

        Exit Function

    End If
    
    ' No podes atacar si esta en consulta
    If UserList(VictimIndex).flags.EnConsulta Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios mientras estan en consulta.", FontTypeNames.FONTTYPE_INFO)

        Exit Function

    End If

    ' No podes atacar si está protegido
    If UserList(AttackerIndex).Counters.Shield > 0 Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios mientras estas protegido.", FontTypeNames.FONTTYPE_INFO)

        Exit Function

    End If
    
    ' No podes atacar si esta en consulta
    If UserList(VictimIndex).Counters.Shield > 0 Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar usuarios mientras estan protegidos.", FontTypeNames.FONTTYPE_INFO)

        Exit Function

    End If
    
    ' No podes atacar a tu compañero en Retos
    If UserList(VictimIndex).flags.SlotReto > 0 Then

        With Retos(UserList(VictimIndex).flags.SlotReto)

            If .config(eRetoConfig.eFuegoAmigo) = 0 Then
                If .User(UserList(AttackerIndex).flags.SlotRetoUser).Team = .User(UserList(VictimIndex).flags.SlotRetoUser).Team Then
                    PuedeAtacar = False
    
                    Exit Function
    
                End If

            End If

        End With

    End If
    
    If UserList(VictimIndex).flags.SlotFast > 0 Then
        If UserList(AttackerIndex).flags.FightTeam = UserList(VictimIndex).flags.FightTeam Then
            Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a tu compañero.", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacar = False

            Exit Function

        End If

    End If

    ' No podes atacar si la cuenta regresiva está activo.
    If UserList(AttackerIndex).Counters.TimeFight > 0 Then
        WriteConsoleMsg AttackerIndex, "No puedes atacar hasta que no termine la cuenta regresiva.", FontTypeNames.FONTTYPE_INFO
        PuedeAtacar = False

        Exit Function

    End If
    
    ' Chequeos de no atacar en eventos.
    If UserList(AttackerIndex).flags.SlotEvent > 0 Then
        If EsGm(VictimIndex) Then
            PuedeAtacar = False

            Exit Function

        End If
        
        If UserList(VictimIndex).flags.SlotEvent <= 0 Then
            PuedeAtacar = False

            Exit Function

        End If
            
        If Events(UserList(AttackerIndex).flags.SlotEvent).Modality = eModalityEvent.DagaRusa Then
            WriteConsoleMsg AttackerIndex, "No puedes atacar en este tipo de eventos.", FontTypeNames.FONTTYPE_INFO
            PuedeAtacar = False

            Exit Function

        End If
        
        If Events(UserList(AttackerIndex).flags.SlotEvent).TimeCount > 0 Then
            WriteConsoleMsg AttackerIndex, "No puedes atacar hasta que no termine la cuenta regresiva.", FontTypeNames.FONTTYPE_INFO
            PuedeAtacar = False

            Exit Function

        End If
        
        If Events(UserList(AttackerIndex).flags.SlotEvent).Run Then
            If UserList(AttackerIndex).flags.SlotUserEvent > 0 Then
                If Events(UserList(AttackerIndex).flags.SlotEvent).Users(UserList(AttackerIndex).flags.SlotUserEvent).Team > 0 Then
                    If Not CanAttackUserEvent(AttackerIndex, VictimIndex) Then
                        WriteConsoleMsg AttackerIndex, "No puedes atacar a tu compañero", FontTypeNames.FONTTYPE_INFO
                        PuedeAtacar = False

                        Exit Function

                    End If
                        
                End If

            End If

        End If

    End If
    
    'Estamos en una Arena? o un trigger zona segura?
    Select Case TriggerZonaPelea(AttackerIndex, VictimIndex)

        Case eTrigger6.TRIGGER6_PERMITE
            PuedeAtacar = (UserList(VictimIndex).flags.AdminInvisible = 0)

            Exit Function
        
        Case eTrigger6.TRIGGER6_PROHIBE
            PuedeAtacar = False

            Exit Function
        
        Case eTrigger6.TRIGGER6_AUSENTE

            'Si no estamos en el Trigger 6 entonces es imposible atacar un gm
            If (UserList(VictimIndex).flags.Privilegios And PlayerType.User) = 0 Then
                If UserList(VictimIndex).flags.AdminInvisible = 0 Then Call WriteConsoleMsg(AttackerIndex, "El ser es demasiado poderoso.", FontTypeNames.FONTTYPE_WARNING)
                PuedeAtacar = False

                Exit Function

            End If

    End Select
        
    If UserList(AttackerIndex).flags.Privilegios And (PlayerType.SemiDios) Then
        If Not (EsGm(AttackerIndex) And EsGm(VictimIndex)) Then
            Call WriteConsoleMsg(AttackerIndex, "No tienes permitido atacar a los usuarios del juego.", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacar = False
    
            Exit Function

        End If

    End If
        
    If Not MapInfo(UserList(AttackerIndex).Pos.Map).Pk Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar en zona segura.", FontTypeNames.FONTTYPE_INFO)
        PuedeAtacar = False

        Exit Function

    End If
    
    If UserList(VictimIndex).GuildIndex > 0 Then
        If UserList(VictimIndex).GuildIndex = UserList(AttackerIndex).GuildIndex And GuildsInfo(UserList(VictimIndex).GuildIndex).Lvl >= 15 Then
            Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a tu Compañero de Clan.", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacar = False

            Exit Function

        End If
    
    End If
    
    If Not MapInfo(UserList(AttackerIndex).Pos.Map).FreeAttack Then
        If UserList(AttackerIndex).Faction.Status = 0 Then
            If Not Escriminal(AttackerIndex) And Not Escriminal(VictimIndex) Then
                If UserList(AttackerIndex).flags.Seguro Then
                    Call WriteConsoleMsg(AttackerIndex, "Debes desactivar el seguro para atacar a otro ciudadano. ¡Te convertirás en Criminal!", FontTypeNames.FONTTYPE_WARNING)
                    PuedeAtacar = False
    
                    Exit Function
    
                End If

            End If
    
        Else
    
            If Not Escriminal(AttackerIndex) And Not Escriminal(VictimIndex) Then
                If UserList(AttackerIndex).flags.Seguro Then
                    Call WriteConsoleMsg(AttackerIndex, "Debes desactivar el seguro para atacar a otro ciudadano. ¡Te convertirás en Criminal!", FontTypeNames.FONTTYPE_WARNING)
                    PuedeAtacar = False
    
                    Exit Function
    
                End If

            End If
            
            If Not mFacciones.Faction_CanAttack(AttackerIndex, VictimIndex) Then
                PuedeAtacar = False
                Call WriteConsoleMsg(AttackerIndex, "Tu facción no permite atacar a la facción de la víctima.", FontTypeNames.FONTTYPE_WARNING)
    
                Exit Function
    
            End If

        End If
        
    End If
    
    'Estas atacando desde un trigger seguro? o tu victima esta en uno asi?
    If MapData(UserList(VictimIndex).Pos.Map, UserList(VictimIndex).Pos.X, UserList(VictimIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Or MapData(UserList(AttackerIndex).Pos.Map, UserList(AttackerIndex).Pos.X, UserList(AttackerIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes pelear aquí.", FontTypeNames.FONTTYPE_WARNING)
        PuedeAtacar = False

        Exit Function

    End If
    
    PuedeAtacar = True

    '<EhFooter>
    Exit Function

PuedeAtacar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.SistemaCombate.PuedeAtacar " & "at line " & Erl
        
    '</EhFooter>
End Function

Public Function PuedeAtacarNPC(ByVal AttackerIndex As Integer, _
                               ByVal NpcIndex As Integer, _
                               Optional ByVal Paraliza As Boolean = False, _
                               Optional ByVal IsPet As Boolean = False) As Boolean
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

    On Error GoTo ErrHandler

    With Npclist(NpcIndex)

        If UserList(AttackerIndex).flags.Privilegios And (PlayerType.SemiDios) Then
            Call WriteConsoleMsg(AttackerIndex, "¡¡No puedes atacar criaturas!!", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If
        
        'Estas muerto?
        If UserList(AttackerIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(AttackerIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If
        
        'Sos SemiDios?
        If UserList(AttackerIndex).flags.Privilegios And PlayerType.SemiDios Then

            'No pueden atacar NPC los SemiDioses.
            Exit Function

        End If
        
        ' No podes atacar si estas en consulta
        If UserList(AttackerIndex).flags.EnConsulta Then
            Call WriteConsoleMsg(AttackerIndex, "No puedes atacar npcs mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If
        
        'Es una criatura atacable?
        If .Attackable = 0 Then
            Call WriteConsoleMsg(AttackerIndex, "No puedes atacar esta criatura.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If
        
        'Es valida la distancia a la cual estamos atacando?
        If Distancia(UserList(AttackerIndex).Pos, .Pos) >= MAXDISTANCIAARCO Then
            Call WriteConsoleMsg(AttackerIndex, "Estás muy lejos para disparar.", FontTypeNames.FONTTYPE_FIGHT)

            Exit Function

        End If
        
        'Es una criatura No-Hostil?
        If .Hostile = 0 Then

            'Es Guardia del Caos?
            If .NPCtype = eNPCType.GuardiasCaos Then

                'Lo quiere atacar un caos?
                If esCaos(AttackerIndex) Then
                    Call WriteConsoleMsg(AttackerIndex, "No puedes atacar Guardias del Caos siendo de la legión oscura.", FontTypeNames.FONTTYPE_INFO)

                    Exit Function

                End If

                'Es guardia Real?
            ElseIf .NPCtype = eNPCType.GuardiaReal Then

                'Lo quiere atacar un Armada?
                If esArmada(AttackerIndex) Then
                    Call WriteConsoleMsg(AttackerIndex, "No puedes atacar Guardias Reales siendo del ejército real.", FontTypeNames.FONTTYPE_INFO)

                    Exit Function

                End If

                'Tienes el seguro puesto?
                If UserList(AttackerIndex).flags.Seguro Then
                    Call WriteConsoleMsg(AttackerIndex, "Para poder atacar Guardias Reales debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)

                    Exit Function

                Else
                    Call WriteConsoleMsg(AttackerIndex, "¡Atacaste un Guardia Real! Eres un criminal.", FontTypeNames.FONTTYPE_INFO)
                    Call VolverCriminal(AttackerIndex)
                    PuedeAtacarNPC = True

                    Exit Function

                End If
        
                'No era un Guardia, asi que es una criatura No-Hostil común.
                'Para asegurarnos que no sea una Mascota:
            ElseIf .MaestroUser = 0 Then

                'Si sos ciudadano tenes que quitar el seguro para atacarla.
                If Not Escriminal(AttackerIndex) Then
                    
                    ' Si sos armada no podes atacarlo directamente
                    If esArmada(AttackerIndex) Then
                        Call WriteConsoleMsg(AttackerIndex, "Los miembros del ejército real no pueden atacar npcs no hostiles.", FontTypeNames.FONTTYPE_INFO)

                        Exit Function

                    End If
                
                    'Sos ciudadano, tenes el seguro puesto?
                    If UserList(AttackerIndex).flags.Seguro Then
                        Call WriteConsoleMsg(AttackerIndex, "Para atacar a este NPC debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)

                        Exit Function

                    Else
                        'No tiene seguro puesto. Puede atacar pero es penalizado.
                        Call WriteConsoleMsg(AttackerIndex, "Atacaste un NPC no-hostil. Continúa haciéndolo y te podrás convertir en criminal.", FontTypeNames.FONTTYPE_INFO)
                        'NicoNZ: Cambio para que al atacar npcs no hostiles no bajen puntos de nobleza
                        Call DisNobAuBan(AttackerIndex, 0, 1000)
                        PuedeAtacarNPC = True

                        Exit Function

                    End If

                End If

            End If

        End If
    
        Dim MasterIndex As Integer

        MasterIndex = .MaestroUser
        
        'Es el NPC mascota de alguien?
        If MasterIndex > 0 Then
            
            ' Dueño de la mascota ciuda?
            If Not Escriminal(MasterIndex) Then
                
                ' Atacante ciuda?
                If Not Escriminal(AttackerIndex) Then
                    
                    'Atacante armada?
                    If esArmada(AttackerIndex) Then
                        'El atacante es Armada y esta intentando atacar mascota de un Ciudadano
                        Call WriteConsoleMsg(AttackerIndex, "Los miembros del ejército real no pueden atacar mascotas de ciudadanos.", FontTypeNames.FONTTYPE_INFO)

                        Exit Function

                    End If
                    
                    'El atacante es Ciudadano y esta intentando atacar mascota de un Ciudadano.
                    If UserList(AttackerIndex).flags.Seguro Then
                        'El atacante tiene el seguro puesto. No puede atacar.
                        Call WriteConsoleMsg(AttackerIndex, "Para atacar mascotas de ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)

                        Exit Function

                    Else
                        'El atacante no tiene el seguro puesto. Recibe penalización.
                        Call WriteConsoleMsg(AttackerIndex, "Has atacado la Mascota de un ciudadano. Eres un criminal.", FontTypeNames.FONTTYPE_INFO)
                        Call VolverCriminal(AttackerIndex)
                        PuedeAtacarNPC = True

                        Exit Function

                    End If

                Else

                    'El atacante es criminal y quiere atacar un elemental ciuda, pero tiene el seguro puesto (NicoNZ)
                    If UserList(AttackerIndex).flags.Seguro Then
                        Call WriteConsoleMsg(AttackerIndex, "Para atacar mascotas de ciudadanos debes quitarte el seguro.", FontTypeNames.FONTTYPE_INFO)

                        Exit Function

                    End If

                End If
            
                'Es mascota de un caos?
            ElseIf esCaos(MasterIndex) Then

                'Es Caos el Dueño.
                If esCaos(AttackerIndex) Then
                    'Un Caos intenta atacar una criatura de un Caos. No puede atacar.
                    Call WriteConsoleMsg(AttackerIndex, "Los miembros de la legión oscura no pueden atacar mascotas de otros legionarios. ", FontTypeNames.FONTTYPE_INFO)

                    Exit Function

                End If

            End If
            
            ' No es mascota de nadie, le pertenece a alguien?
            
        ElseIf .Owner > 0 Then
        
            Dim OwnerUserIndex As Integer

            OwnerUserIndex = .Owner
            
            ' Puede atacar a su propia criatura!
            If OwnerUserIndex = AttackerIndex Then
                PuedeAtacarNPC = True
                Call IntervaloPerdioNpc(OwnerUserIndex, True) ' Renuevo el timer
                Exit Function

            End If
            
            ' Esta compartiendo el npc con el atacante? => Puede atacar!
            If UserList(OwnerUserIndex).flags.ShareNpcWith = AttackerIndex Then
                PuedeAtacarNPC = True
                Exit Function

            End If
            
            ' Si son del mismo clan o party, pueden atacar (No renueva el timer)
            If Not SameClan(OwnerUserIndex, AttackerIndex) And Not SameParty(OwnerUserIndex, AttackerIndex) Then
            
                ' Si se le agoto el tiempo
                If IntervaloPerdioNpc(OwnerUserIndex) Then ' Se lo roba :P
                    Call PerdioNpc(OwnerUserIndex)
                    Call ApropioNpc(AttackerIndex, NpcIndex)
                    PuedeAtacarNPC = True
                    Exit Function
                
                Else
                    
                    ' El npc le pertenece a un ciudadano
                    If Not Escriminal(OwnerUserIndex) Then
                        
                        'El atacante es Armada y esta intentando atacar un npc de un Ciudadano
                        If esArmada(AttackerIndex) Then
                        
                            'Intententa atacar un npc de un armada?
                            If esArmada(OwnerUserIndex) Then
                                'El atacante es Armada y esta intentando atacar el npc de un armada: No puede
                                Call WriteConsoleMsg(AttackerIndex, "Los miembros del Ejército Real no pueden atacar criaturas pertenecientes a otros miembros del Ejército Real", FontTypeNames.FONTTYPE_INFO)
                                Exit Function
                            
                            Else
                                'El atacante es Armada y esta intentando atacar un npc de un ciuda
                                Call WriteConsoleMsg(AttackerIndex, "Los miembros de la Armada Real no pueden deshonrar al Rey.", FontTypeNames.FONTTYPE_INFO)
                                Exit Function

                            End If
                            
                            ' No es aramda, puede ser criminal o ciuda
                        Else
                            
                            'El atacante es Ciudadano y esta intentando atacar un npc de un Ciudadano.
                            If Not Escriminal(AttackerIndex) Then
                                'El atacante tiene el seguro puesto. No puede atacar.
                                Call WriteConsoleMsg(AttackerIndex, "¡Ve a hacerte criminal a alguna parte y vuelve a matadme!", FontTypeNames.FONTTYPE_INFO)
                                Exit Function
                                
                                'El atacante es criminal y esta intentando atacar un npc de un Ciudadano.
                            Else
                                PuedeAtacarNPC = True

                            End If

                        End If

                    End If

                End If

            End If
            
            ' Si no tiene dueño el npc, se lo apropia
        Else

            ' Solo pueden apropiarse de npcs los caos, armadas o ciudas.
            If Not Escriminal(AttackerIndex) Or esCaos(AttackerIndex) Then

                ' No puede apropiarse de los pretos!
                If Npclist(NpcIndex).NPCtype <> eNPCType.Pretoriano Then

                    ' No puede apropiarse del dragon de dd!
                    If Npclist(NpcIndex).NPCtype <> DRAGON Then

                        ' Si es una mascota atacando, no se apropia del npc
                        If Not IsPet Then

                            ' No es dueño de ningun npc => Se lo apropia.
                            If UserList(AttackerIndex).flags.OwnedNpc = 0 Then
                                Call ApropioNpc(AttackerIndex, NpcIndex)
                                ' Es dueño de un npc, pero no puede ser de este porque no tiene propietario.
                            Else

                                ' Se va a adueñar del npc (y perder el otro) solo si no inmobiliza/paraliza
                                If Not Paraliza Then Call ApropioNpc(AttackerIndex, NpcIndex)

                            End If

                        End If

                    End If

                End If

            End If

        End If
        
        If (UserList(AttackerIndex).flags.SlotEvent) > 0 And (.flags.TeamEvent > 0) Then
            If UserList(AttackerIndex).flags.FightTeam = .flags.TeamEvent Then
                Call WriteConsoleMsg(AttackerIndex, "La criatura está de tu lado. ¡No puedes atacarla!", FontTypeNames.FONTTYPE_FIGHT)

                Exit Function

            End If

        End If

    End With
    
    ' ClanIndex = Criatura perteneciente a algun castillo
    If Npclist(NpcIndex).ClanIndex > 0 Then
        If Npclist(NpcIndex).NPCtype = eNPCType.Pretoriano Then
            If Not ClanPretoriano(Npclist(NpcIndex).ClanIndex).CanAtackMember(NpcIndex) Then
                Call WriteConsoleMsg(AttackerIndex, "Debes matar al resto del ejército antes de atacar al rey.", FontTypeNames.FONTTYPE_FIGHT)

                Exit Function

            End If

        End If

    End If

    ' Mascotas
    If Npclist(NpcIndex).NPCtype = eNPCType.Mascota Then
        If Not MapInfo(UserList(AttackerIndex).Pos.Map).Pk Then
            Call WriteConsoleMsg(AttackerIndex, "No puedes atacar estas criaturas en zona segura.", FontTypeNames.FONTTYPE_INFO)

            Exit Function

        End If

    End If
        
    ' Npcs de mascotas
    If Npclist(NpcIndex).MonturaIndex > 0 Then
        Call WriteConsoleMsg(AttackerIndex, "No puedes atacar a las monturas.", FontTypeNames.FONTTYPE_INFO)

        Exit Function

    End If
    
    ' Dragones solo se matan con objeto especial
    If Npclist(NpcIndex).NPCtype = eNPCType.DRAGON Then
        If (UserList(AttackerIndex).Invent.WeaponEqpObjIndex <> EspadaMataDragonesIndex) And (UserList(AttackerIndex).Invent.WeaponEqpObjIndex <> VaraMataDragonesIndex) Then
            
            Call WriteConsoleMsg(AttackerIndex, "Los dragones solo pueden ser atacados con armas especiales.", FontTypeNames.FONTTYPE_FIGHT)

            Exit Function
        
        End If
        
    End If
    
    ' cASTILLOS
    Dim CastleIndex As Integer

    CastleIndex = Npclist(NpcIndex).CastleIndex
    
    If CastleIndex > 0 Then
        If UserList(AttackerIndex).GuildIndex = 0 Then
            Call WriteConsoleMsg(AttackerIndex, "¡Debes pertenecer a un clan para atacar a esta criatura!", FontTypeNames.FONTTYPE_FIGHT)

            Exit Function

        End If
        
        If Castle(CastleIndex).GuildIndex = UserList(AttackerIndex).GuildIndex Then
            Call WriteConsoleMsg(AttackerIndex, "¡No puedes atacar tu Castillo!", FontTypeNames.FONTTYPE_FIGHT)

            Exit Function
            
        End If

    End If
    
    PuedeAtacarNPC = True
        
    Exit Function
        
ErrHandler:
    
    Dim AtckName  As String

    Dim OwnerName As String

    If AttackerIndex > 0 Then AtckName = UserList(AttackerIndex).Name
    'If OwnerUserIndex > 0 Then OwnerName = UserList(OwnerUserIndex).Name
    
    Call LogError("Error en PuedeAtacarNpc. Erorr: " & Err.number & " - " & Err.description & " Atacante: " & AttackerIndex & "-> " & AtckName & ". Owner: -> " & OwnerName & ". NpcIndex: " & NpcIndex & ".")

End Function

Private Function SameClan(ByVal UserIndex As Integer, _
                          ByVal OtherUserIndex As Integer) As Boolean

    '***************************************************
    'Autor: ZaMa
    'Returns True if both players belong to the same clan.
    'Last Modification: 16/11/2009
    '***************************************************
    '<EhHeader>
    On Error GoTo SameClan_Err

    '</EhHeader>
    SameClan = (UserList(UserIndex).GuildIndex = UserList(OtherUserIndex).GuildIndex) And UserList(UserIndex).GuildIndex <> 0
    '<EhFooter>
    Exit Function

SameClan_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.SistemaCombate.SameClan " & "at line " & Erl
        
    '</EhFooter>
End Function

Private Function SameParty(ByVal UserIndex As Integer, _
                           ByVal OtherUserIndex As Integer) As Boolean

    '***************************************************
    'Autor: ZaMa
    'Returns True if both players belong to the same party.
    'Last Modification: 16/11/2009
    '***************************************************
    '<EhHeader>
    On Error GoTo SameParty_Err

    '</EhHeader>
    SameParty = UserList(UserIndex).GroupIndex = UserList(OtherUserIndex).GroupIndex And UserList(UserIndex).GroupIndex <> 0
    '<EhFooter>
    Exit Function

SameParty_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.SistemaCombate.SameParty " & "at line " & Erl
        
    '</EhFooter>
End Function

Sub CalcularDarExp_Bonus(ByVal UserIndex As Integer, ByVal Exp As Long)
    
    On Error GoTo ErrHandler
    
    Dim Temp As Long
    
    With UserList(UserIndex)

        If .flags.Premium = 1 Then
            .Stats.Exp = .Stats.Exp + Int(Exp * 0.3)
            Temp = Temp + Int(Exp * 0.3)

            'WriteConsoleMsg UserIndex, "PremiumExp» Has ganado " & Int(Exp * 0.3) & " puntos de experiencia.", FontTypeNames.FONTTYPE_USERPREMIUM
        End If
            
        If .flags.Oro = 1 Then
            .Stats.Exp = .Stats.Exp + Int(Exp * 0.1)
            Temp = Temp + Int(Exp * 0.15)
            
            'Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageRenderConsole("Exp +" & CStr(Format(Int(Exp * 0.15), "###,###,###")), d_Exp, 3000, 0))
            'WriteConsoleMsg UserIndex, "OroExp» Has ganado " & Int(Exp * 0.15) & " puntos de experiencia.", FontTypeNames.FONTTYPE_DIOS
        ElseIf .flags.Plata = 1 Then
            .Stats.Exp = .Stats.Exp + Int(Exp * 0.07)
            Temp = Temp + Int(Exp * 0.1)
            'WriteConsoleMsg UserIndex, "PlataExp» Has ganado " & Int(Exp * 0.1) & " puntos de experiencia.", FontTypeNames.FONTTYPE_USERPLATA
        ElseIf .flags.Bronce = 1 Then
            .Stats.Exp = .Stats.Exp + Int(Exp * 0.03)
            Temp = Temp + Int(Exp * 0.05)

            ' WriteConsoleMsg UserIndex, "BronceExp» Has ganado " & Int(Exp * 0.05) & " puntos de experiencia.", FontTypeNames.FONTTYPE_USERBRONCE
        End If
            
        ' ## Alterar la exp que da bonificada
        If .Stats.BonusTipe = eEffectObj.e_Exp Then
            .Stats.Exp = .Stats.Exp + Int(Exp * .Stats.BonusValue)
            Temp = Temp + Int(Exp * .Stats.BonusValue)
            WriteConsoleMsg UserIndex, "BonusExp» Has ganado " & Int(Exp * .Stats.BonusValue) & " puntos de experiencia.", FontTypeNames.FONTTYPE_DIOS

        End If
            
        If .Invent.ReliquiaSlot > 0 Then
            
            If ObjData(.Invent.ReliquiaObjIndex).EffectUser.ExpNpc > 0 Then
                'ExpaDar = mEffect.Effect_UpdatePorc(UserIndex, ExpaDar)
                
                'WriteConsoleMsg UserIndex, "ReliquiaExp» Tu experiencia obtenida se ha incrementado.", FontTypeNames.FONTTYPE_DIOS
                    
            End If

        End If
        
        Dim SelectedBono As Integer

        SelectedBono = .flags.SelectedBono
            
        If SelectedBono > 0 Then
            If ObjData(SelectedBono).BonoExp > 0 Then
                .Stats.Exp = .Stats.Exp + Int(Exp * ObjData(SelectedBono).BonoExp)
                Temp = Temp + Int(Exp * ObjData(SelectedBono).BonoExp)

                'WriteConsoleMsg UserIndex, "BonoExp» Has ganado " & Int(Exp * ObjData(SelectedBono).BonoExp) & " puntos de experiencia.", FontTypeNames.FONTTYPE_INFOBOLD
            End If

        End If
        
        'Call WriteConsoleMsg(UserIndex, "Has ganado un total de " & CStr(Exp + Temp) & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
        Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageRenderConsole("Exp +" & CStr(Format(Temp + Exp, "###,###,###")), d_Exp, 3000, 0))
        
        ' Bono premio del servidor
        If NumUsers + UsersBot >= 150 Then

            Dim ExpServidor As Long

            ExpServidor = CalcularPorcentajeBonificacion(Exp)
            .Stats.Exp = .Stats.Exp + ExpServidor
            Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageRenderConsole("BonusOnlines +" & CStr(Format(ExpServidor, "###,###,###")), d_Exp, 3000, 0))

        End If
        
        ' Bono de Castillos 10%
        If .GuildIndex > 0 Then
            If CastleBonus = .GuildIndex Then
                .Stats.Exp = .Stats.Exp + Int(Exp * 0.1)
                Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageRenderConsole("BonusCastle +" & CStr(Format(Int(Exp * 0.1), "###,###,###")), d_Exp, 3000, 0))
            
            End If

        End If
        
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Ocurrio un error en Bonus de Exp")

End Sub

Sub CalcularDarExp_Bonus_Party(ByVal UserIndex As Integer, _
                               ByVal SlotUser As Byte, _
                               ByRef Exp As Long)
    
    On Error GoTo ErrHandler

    Dim TempExp As Long
    
    With UserList(UserIndex)

        If .flags.Premium = 1 Then
            TempExp = TempExp + (Int(Exp * 0.3))
            
        End If
        
        If .flags.Oro = 1 Then
            TempExp = TempExp + (Int(Exp * 0.1))
        ElseIf .flags.Plata = 1 Then
            TempExp = TempExp + (Int(Exp * 0.07))
        ElseIf .flags.Bronce = 1 Then
            TempExp = TempExp + (Int(Exp * 0.03))

        End If
        
        If .Stats.BonusTipe = eEffectObj.e_Exp Then
            TempExp = TempExp + Int(Exp * .Stats.BonusValue)
            
        End If
            
        If .Invent.ReliquiaSlot > 0 Then
            
            If ObjData(.Invent.ReliquiaObjIndex).EffectUser.ExpNpc > 0 Then
                'ExpaDar = mEffect.Effect_UpdatePorc(UserIndex, ExpaDar)
                    
            End If

        End If
        
        Dim SelectedBono As Integer

        SelectedBono = .flags.SelectedBono
            
        If SelectedBono > 0 Then
            If ObjData(SelectedBono).BonoExp > 0 Then
                TempExp = TempExp + Int(Exp * ObjData(SelectedBono).BonoExp)
                    
            End If

        End If
    
        ' Bono premio del servidor
        If NumUsers + UsersBot >= 150 Then

            Dim ExpServidor As Long

            ExpServidor = CalcularPorcentajeBonificacion(Exp)
            TempExp = TempExp + ExpServidor
            
            'Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageRenderConsole("BonusOnlines +" & CStr(Format(ExpServidor, "###,###,###")), d_Exp, 3000, 0))
        End If
        
        ' Bono de Castillos 10%
        If .GuildIndex > 0 Then
            If CastleBonus = .GuildIndex Then
                TempExp = TempExp + Int(Exp * 0.1)

            End If

        End If
        
        Exp = Exp + TempExp
        
    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Ocurrio un error en Bonus eXp Party")

End Sub

Sub CalcularDarExp(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, ByVal ElDaño As Long)

    '<EhHeader>
    On Error GoTo CalcularDarExp_Err

    '</EhHeader>

    '***************************************************
    'Autor: Nacho (Integer)
    'Last Modification: 03/09/06 Nacho
    'Reescribi gran parte del Sub
    'Ahora, da toda la experiencia del npc mientras este vivo.
    '***************************************************
    Dim ExpaDar      As Long

    Dim ResourceaDar As Long

    Dim ExpClan      As Long

    Dim A            As Long, B As Long
        
    Dim Obj          As Obj
        
    Dim Diferencia   As Long

    '[Nacho] Chekeamos que las variables sean validas para las operaciones
    If ElDaño <= 0 Then ElDaño = 0

    'If Npclist(NpcIndex).Stats.MinHp <= 0 Then Exit Sub
    If ElDaño > Npclist(NpcIndex).Stats.MinHp Then ElDaño = Npclist(NpcIndex).Stats.MinHp
        
    '[Nacho] La experiencia a dar es la porcion de vida quitada * toda la experiencia
    ExpaDar = CLng(ElDaño * (Npclist(NpcIndex).GiveEXP / Npclist(NpcIndex).Stats.MaxHp))
        
    ' Criaturas que dan experiencia de clan
    If UserList(UserIndex).GuildIndex > 0 Then
        ExpClan = CLng(ElDaño * (Npclist(NpcIndex).GiveEXPGuild / Npclist(NpcIndex).Stats.MaxHp))
            
        If ExpClan > 0 Then
            Call Guilds_AddExp(UserList(UserIndex).GuildIndex, ExpClan)
                
            If ExpClan > Npclist(NpcIndex).flags.ExpGuildCount Then
                ExpClan = Npclist(NpcIndex).flags.ExpGuildCount
                Npclist(NpcIndex).flags.ExpGuildCount = 0
            Else
                Npclist(NpcIndex).flags.ExpGuildCount = Npclist(NpcIndex).flags.ExpGuildCount - ExpClan

            End If

        End If
        
    End If
        
    ' Criaturas que dan recursos (leña,fragmentos,minerales)
    If Npclist(NpcIndex).GiveResource.ObjIndex > 0 Then
        ResourceaDar = CLng(ElDaño * (Npclist(NpcIndex).GiveResource.Amount / Npclist(NpcIndex).Stats.MaxHp))

        If ResourceaDar > Npclist(NpcIndex).flags.ResourceCount Then
            ResourceaDar = Npclist(NpcIndex).flags.ResourceCount
            Npclist(NpcIndex).flags.ResourceCount = 0
        Else
            Npclist(NpcIndex).flags.ResourceCount = Npclist(NpcIndex).flags.ResourceCount - ResourceaDar

        End If
            
        If ResourceaDar > 0 Then
            Obj.ObjIndex = Npclist(NpcIndex).GiveResource.ObjIndex
            Obj.Amount = ResourceaDar
            Call MeterItemEnInventario(UserIndex, Obj)

        End If

    End If
            
    If ExpaDar <= 0 Then Exit Sub
    
    '[Nacho] Vamos contando cuanta experiencia sacamos, porque se da toda la que no se dio al user que mata al NPC
    'Esto es porque cuando un elemental ataca, no se da exp, y tambien porque la cuenta que hicimos antes
    'Podria dar un numero fraccionario, esas fracciones se acumulan hasta formar enteros ;P
    If ExpaDar > Npclist(NpcIndex).flags.ExpCount Then
        ExpaDar = Npclist(NpcIndex).flags.ExpCount
        Npclist(NpcIndex).flags.ExpCount = 0
    Else
        Npclist(NpcIndex).flags.ExpCount = Npclist(NpcIndex).flags.ExpCount - ExpaDar

    End If

    If HappyHour And Not Npclist(NpcIndex).NPCtype = DRAGON Then
        ExpaDar = ExpaDar * 2

    End If
    
    If MapInfo(Npclist(NpcIndex).Pos.Map).Exp > 0 And Not Npclist(NpcIndex).NPCtype = DRAGON Then
        ExpaDar = ExpaDar * MapInfo(Npclist(NpcIndex).Pos.Map).Exp

    End If
        
    '[Nacho] Le damos la exp al user
    If ExpaDar > 0 Then
        If UserList(UserIndex).GroupIndex > 0 Then
            Call mGroup.AddExpGroup(UserIndex, ExpaDar)
        Else
            UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpaDar

            If UserList(UserIndex).Stats.Exp > MAXEXP Then UserList(UserIndex).Stats.Exp = MAXEXP
            
            'Call WriteConsoleMsg(UserIndex, "Has ganado " & ExpaDar & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
            
            ' // NUEVO
            Call CalcularDarExp_Bonus(UserIndex, ExpaDar)

        End If
        
        Call CheckUserLevel(UserIndex)

    End If

    '<EhFooter>
    Exit Sub

CalcularDarExp_Err:
    LogError Err.description & vbCrLf & "in CalcularDarExp " & "at line " & Erl

    '</EhFooter>
End Sub

Public Function TriggerZonaPelea(ByVal Origen As Integer, _
                                 ByVal Destino As Integer) As eTrigger6

    '<EhHeader>
    On Error GoTo TriggerZonaPelea_Err

    '</EhHeader>

    Dim tOrg As eTrigger

    Dim tDst As eTrigger
    
    tOrg = MapData(UserList(Origen).Pos.Map, UserList(Origen).Pos.X, UserList(Origen).Pos.Y).trigger
    tDst = MapData(UserList(Destino).Pos.Map, UserList(Destino).Pos.X, UserList(Destino).Pos.Y).trigger
    
    If tOrg = eTrigger.ZONAPELEA Or tDst = eTrigger.ZONAPELEA Then
        If tOrg = tDst Then
            TriggerZonaPelea = TRIGGER6_PERMITE
        Else
            TriggerZonaPelea = TRIGGER6_PROHIBE

        End If

    Else
        TriggerZonaPelea = TRIGGER6_AUSENTE

    End If

    '<EhFooter>
    Exit Function

TriggerZonaPelea_Err:
    LogError Err.description & vbCrLf & "in TriggerZonaPelea " & "at line " & Erl

    '</EhFooter>
End Function

Sub UserEnvenena(ByVal AtacanteIndex As Integer, ByVal VictimaIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo UserEnvenena_Err

    '</EhHeader>

    Dim ObjInd As Integer
    
    ObjInd = UserList(AtacanteIndex).Invent.WeaponEqpObjIndex
    
    If ObjInd > 0 Then
        If ObjData(ObjInd).proyectil = 1 Then
            ObjInd = UserList(AtacanteIndex).Invent.MunicionEqpObjIndex

        End If
        
        If ObjInd > 0 Then
            If ObjData(ObjInd).Envenena = 1 Then
                
                If RandomNumber(1, 100) < 60 Then
                    UserList(VictimaIndex).flags.Envenenado = 1
                    Call WriteConsoleMsg(VictimaIndex, "¡¡" & UserList(AtacanteIndex).Name & " te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(AtacanteIndex, "¡¡Has envenenado a " & UserList(VictimaIndex).Name & "!!", FontTypeNames.FONTTYPE_FIGHT)

                End If

            End If

        End If

    End If
    
    Call FlushBuffer(VictimaIndex)
    '<EhFooter>
    Exit Sub

UserEnvenena_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.SistemaCombate.UserEnvenena " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub LanzarProyectil(ByVal UserIndex As Integer, ByVal X As Byte, ByVal Y As Byte)

    '***************************************************
    'Autor: ZaMa
    'Last Modification: 10/07/2010
    'Throws an arrow or knive to target user/npc.
    '***************************************************
    On Error GoTo ErrHandler

    Dim MunicionSlot    As Byte

    Dim MunicionIndex   As Integer

    Dim WeaponSlot      As Byte

    Dim WeaponIndex     As Integer

    Dim targetUserIndex As Integer

    Dim TargetNpcIndex  As Integer

    Dim DummyInt        As Integer
    
    Dim Threw           As Boolean

    Threw = True
    
    'Make sure the item is valid and there is ammo equipped.
    With UserList(UserIndex)
        
        With .Invent
            MunicionSlot = .MunicionEqpSlot
            MunicionIndex = .MunicionEqpObjIndex
            WeaponSlot = .WeaponEqpSlot
            WeaponIndex = .WeaponEqpObjIndex

        End With
        
        ' Tiene arma equipada?
        If WeaponIndex = 0 Then
            DummyInt = 1
            Call WriteConsoleMsg(UserIndex, "No tienes un arco o cuchilla equipada.", FontTypeNames.FONTTYPE_INFO)
            
            ' En un slot válido?
        ElseIf WeaponSlot < 1 Or WeaponSlot > .CurrentInventorySlots Then
            DummyInt = 1
            Call WriteConsoleMsg(UserIndex, "No tienes un arco o cuchilla equipada.", FontTypeNames.FONTTYPE_INFO)
            
            ' Usa munición? (Si no la usa, puede ser un arma arrojadiza)
        ElseIf ObjData(WeaponIndex).Municion = 1 Then
        
            ' La municion esta equipada en un slot valido?
            If MunicionSlot < 1 Or MunicionSlot > .CurrentInventorySlots Then
                DummyInt = 1
                Call WriteConsoleMsg(UserIndex, "No tienes municiones equipadas.", FontTypeNames.FONTTYPE_INFO)
                
                ' Tiene munición?
            ElseIf MunicionIndex = 0 Then
                DummyInt = 1
                Call WriteConsoleMsg(UserIndex, "No tienes municiones equipadas.", FontTypeNames.FONTTYPE_INFO)
                
                ' Son flechas?
            ElseIf ObjData(MunicionIndex).OBJType <> eOBJType.otFlechas Then
                DummyInt = 1
                Call WriteConsoleMsg(UserIndex, "No tienes municiones.", FontTypeNames.FONTTYPE_INFO)
                
                ' Tiene suficientes?
            ElseIf .Invent.Object(MunicionSlot).Amount < 1 Then
                DummyInt = 1
                Call WriteConsoleMsg(UserIndex, "No tienes municiones.", FontTypeNames.FONTTYPE_INFO)

            End If
            
            ' Es un arma de proyectiles?
        ElseIf ObjData(WeaponIndex).proyectil <> 1 Then
            DummyInt = 2

        End If
        
        If DummyInt <> 0 Then
            If DummyInt = 1 Then
                Call Desequipar(UserIndex, WeaponSlot)

            End If
            
            Call Desequipar(UserIndex, MunicionSlot)

            Exit Sub

        End If
    
        'Quitamos stamina
        If .Stats.MinSta >= 10 Then
            Call QuitarSta(UserIndex, RandomNumber(1, 10))
        Else

            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(UserIndex, "Estás muy cansado para luchar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Estás muy cansada para luchar.", FontTypeNames.FONTTYPE_INFO)

            End If

            Exit Sub

        End If
        
        Call LookatTile(UserIndex, .Pos.Map, X, Y)
        
        targetUserIndex = .flags.TargetUser
        TargetNpcIndex = .flags.TargetNPC
        
        'Validate target
        If targetUserIndex > 0 Then

            'Only allow to atack if the other one can retaliate (can see us)
            If Abs(UserList(targetUserIndex).Pos.Y - .Pos.Y) > RANGO_VISION_y Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)

                Exit Sub

            End If
            
            'Prevent from hitting self
            If targetUserIndex = UserIndex Then
                Call WriteConsoleMsg(UserIndex, "¡No puedes atacarte a vos mismo!", FontTypeNames.FONTTYPE_INFO, eMessageType.Combate)

                Exit Sub

            End If
            
            'Attack!
            Threw = UsuarioAtacaUsuario(UserIndex, targetUserIndex)
            
        ElseIf TargetNpcIndex > 0 Then

            'Only allow to atack if the other one can retaliate (can see us)
            If Abs(Npclist(TargetNpcIndex).Pos.Y - .Pos.Y) > RANGO_VISION_y And Abs(Npclist(TargetNpcIndex).Pos.X - .Pos.X) > RANGO_VISION_x Then
                Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para atacar.", FontTypeNames.FONTTYPE_WARNING)

                Exit Sub

            End If
            
            'Is it attackable???
            If Npclist(TargetNpcIndex).Attackable <> 0 Then
                'Attack!
                Threw = UsuarioAtacaNpc(UserIndex, TargetNpcIndex, False)

            End If

        End If
        
        ' Algunas municiones no se pierden
        If MunicionIndex > 0 Then
            If ObjData(MunicionIndex).Ilimitado = 1 Then Threw = False

        End If
        
        ' Solo pierde la munición si pudo atacar al target, o tiro al aire
        If Threw Then
            
            Dim Slot As Byte
            
            ' Tiene equipado arco y flecha?
            If ObjData(WeaponIndex).Municion = 1 Then
                Slot = MunicionSlot
                ' Tiene equipado un arma arrojadiza
            Else
                Slot = WeaponSlot

            End If
            
            'Take 1 knife/arrow away
            Call QuitarUserInvItem(UserIndex, Slot, 1)
            Call UpdateUserInv(False, UserIndex, Slot)
            
        End If
        
    End With
    
    Exit Sub

ErrHandler:

    Dim UserName As String

    If UserIndex > 0 Then UserName = UserList(UserIndex).Name

    Call LogError("Error en LanzarProyectil " & Err.number & ": " & Err.description & ". User: " & UserName & "(" & UserIndex & ")")

End Sub

Public Sub SendCharacterSwing(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo SendCharacterSwing_Err

    '</EhHeader>

    '***************************************************
    'Autor: Amraphen
    'Last Modification: 24/05/2011
    'Sends the CharacterAttackMovement message to the PC Area
    '***************************************************
    With UserList(UserIndex)

        If Not (.flags.Navegando Or .flags.Invisible Or .flags.AdminInvisible) Then Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterAttackMovement(UserList(UserIndex).Char.charindex))

    End With

    '<EhFooter>
    Exit Sub

SendCharacterSwing_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.SistemaCombate.SendCharacterSwing " & "at line " & Erl
        
    '</EhFooter>
End Sub
