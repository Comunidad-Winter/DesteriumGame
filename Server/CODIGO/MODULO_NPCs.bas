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
    '<EhHeader>
    On Error GoTo QuitarMascota_Err

    '</EhHeader>

    Dim i As Integer
    
    With UserList(UserIndex)

        If .MascotaIndex Then
            .MascotaIndex = 0

        End If

    End With

    '<EhFooter>
    Exit Sub

QuitarMascota_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.QuitarMascota " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub QuitarMascotaNpc(ByVal Maestro As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo QuitarMascotaNpc_Err

    '</EhHeader>

    Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1
    '<EhFooter>
    Exit Sub

QuitarMascotaNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.QuitarMascotaNpc " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub MuereNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo MuereNpc_Err

    '</EhHeader>

    '********************************************************
    'Author: Unknown
    'Llamado cuando la vida de un NPC llega a cero.
    'Last Modify Date: 13/07/2010
    '22/06/06: (Nacho) Chequeamos si es pretoriano
    '24/01/2007: Pablo (ToxicWaste): Agrego para actualización de tag si cambia de status.
    '22/05/2010: ZaMa - Los caos ya no suben nobleza ni plebe al atacar npcs.
    '23/05/2010: ZaMa - El usuario pierde la pertenencia del npc.
    '13/07/2010: ZaMa - Optimizaciones de logica en la seleccion de pretoriano, y el posible cambio de alencion del usuario.
    '********************************************************

    Dim MiNPC As Npc

    Dim A     As Long
        
    MiNPC = Npclist(NpcIndex)

    Dim EraCriminal     As Boolean

    Dim PretorianoIndex As Integer
   
    ' @ Reset BOT data
    If MiNPC.BotIndex > 0 Then
        BotIntelligence(MiNPC.BotIndex).Active = False
        MiNPC.BotIndex = 0

    End If

    ' Es pretoriano?
    If MiNPC.NPCtype = eNPCType.Pretoriano Then
        Call ClanPretoriano(MiNPC.ClanIndex).MuerePretoriano(NpcIndex)

    End If
    
    If UserIndex > 0 Then
        If MiNPC.CastleIndex > 0 And UserList(UserIndex).GuildIndex > 0 Then
            Castle_Conquist MiNPC.CastleIndex, UserList(UserIndex).GuildIndex

        End If

        If UserList(UserIndex).flags.SlotEvent > 0 Then

            ' Rey vs Rey
            If MiNPC.numero = 697 Then FinishCastleMode UserList(UserIndex).flags.SlotEvent, UserList(UserIndex).flags.SlotUserEvent
            
            ' La gran Bestia
            If MiNPC.numero = 765 Then Call Events_GranBestia_MuereNpc(UserIndex)
    
        End If

    End If
    
    ' Npcs de invocacion
    If MiNPC.flags.Invocation > 0 Then
        Invocaciones(MiNPC.flags.Invocation).Activo = 0

    End If
    
    ' Npcs de invasiones
    If MiNPC.flags.Invasion > 0 Then
        If Invations(MiNPC.flags.Invasion).Time > 0 Then
            MiNPC.flags.RespawnTime = 10

        End If

    End If
    
    'Quitamos el npc
    Call QuitarNPC(NpcIndex)
    
    If UserIndex > 0 Then ' Lo mato un usuario?

        With UserList(UserIndex)
        
            If MiNPC.flags.Snd3 > 0 Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(MiNPC.flags.Snd3, MiNPC.Pos.X, MiNPC.Pos.Y, 0))

            End If
            
            .flags.TargetNPC = 0
            .flags.TargetNpcTipo = eNPCType.Comun
            
            If .MascotaIndex Then
                Call FollowAmo(.MascotaIndex)

            End If
                
            ' Experiencia de Criaturas restante
            If MiNPC.flags.ExpCount > 0 Then
                If .GroupIndex > 0 Then
                    Call mGroup.AddExpGroup(UserIndex, MiNPC.flags.ExpCount, MiNPC.GiveGLD)
                Else
                    .Stats.Exp = .Stats.Exp + MiNPC.flags.ExpCount

                    If .Stats.Exp > MAXEXP Then .Stats.Exp = MAXEXP
                    'Call WriteConsoleMsg(UserIndex, "Has ganado " & MiNPC.flags.ExpCount & " puntos de experiencia.", FontTypeNames.FONTTYPE_FIGHT, eMessageType.Combate)
                    Call SendData(SendTarget.ToOne, UserIndex, PrepareMessageRenderConsole("Exp +" & CStr(Format(MiNPC.flags.ExpCount, "###,###,###")), d_Exp, 3000, 0))

                End If

                MiNPC.flags.ExpCount = 0

            Else

                If .GroupIndex > 0 Then
                    Call mGroup.AddExpGroup(UserIndex, 0, MiNPC.GiveGLD)

                End If

            End If
                
            ' Experiencia de Clan restante
            If MiNPC.flags.ExpGuildCount > 0 Then
                If UserList(UserIndex).GuildIndex > 0 Then
                    Call Guilds_AddExp(UserIndex, MiNPC.flags.ExpGuildCount)

                End If
                    
                MiNPC.flags.ExpGuildCount = 0

            End If
                
            ' Criaturas que dan recursos (leña,fragmentos,minerales,pecesitoh)
            If MiNPC.flags.ResourceCount > 0 Then

                Dim Obj As Obj

                Obj.ObjIndex = MiNPC.GiveResource.ObjIndex
                Obj.Amount = MiNPC.flags.ResourceCount
                        
                Call MeterItemEnInventario(UserIndex, Obj)
                    
                MiNPC.flags.ResourceCount = 0

            End If
                
            If .Stats.NPCsMuertos < 32000 Then .Stats.NPCsMuertos = .Stats.NPCsMuertos + 1

            Call CheckUserLevel(UserIndex)
            
            If NpcIndex = .flags.ParalizedByNpcIndex Then
                Call RemoveParalisis(UserIndex)

            End If
            
        End With
            
        If MiNPC.MaestroUser = 0 Then
            'Tiramos el inventario
            Call NPC_TIRAR_ITEMS(UserIndex, MiNPC, MiNPC.NPCtype = eNPCType.Pretoriano)
        
            If MiNPC.flags.RespawnTime Then
                If MiNPC.flags.Respawn = 0 Then
                    If Not General.Respawn_Npc_Free(MiNPC.numero, MiNPC.Pos.Map, MiNPC.flags.RespawnTime, MiNPC.CastleIndex, MiNPC.Orig) Then
                        Call LogError("Ocurrio un error al respawnear el NPC: " & MiNPC.numero & ".")

                    End If

                End If

            Else
                'ReSpawn o no
                Call RespawnNpc(MiNPC)

            End If

        End If

        Call WriteConsoleMsg(UserIndex, "Acabaste con " & MiNPC.Name, FontTypeNames.FONTTYPE_INFORED)
    End If ' Userindex > 0
        
    '<EhFooter>
    Exit Sub

MuereNpc_Err:
    LogError Err.description & vbCrLf & "in MuereNpc " & "at line " & Erl

    '</EhFooter>
End Sub

Private Sub ResetNpcFlags(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ResetNpcFlags_Err

    '</EhHeader>

    'Clear the npc's flags
    
    With Npclist(NpcIndex).flags
        .NpcIdle = False
        .KeepHeading = 0
        .RespawnTime = 0
        .Invasion = 0
        .Invocation = 0
        .TeamEvent = 0
        .InscribedPrevio = 0
        .SlotEvent = 0
        .AfectaParalisis = 0
        .AguaValida = 0
        .AttackedBy = vbNullString
        .AttackedByInteger = 0
        .AttackedFirstBy = vbNullString
        .BackUp = 0
        .Bendicion = 0
        .Domable = 0
        .Envenenado = 0
        .Faccion = 0
        .Follow = False
        .AtacaDoble = 0
        .LanzaSpells = 0
        .Invisible = 0
        .Maldicion = 0
        .OldHostil = 0
        .OldMovement = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .Respawn = 0
        .RespawnOrigPos = 0
        .RespawnOrigPosRandom = 0
        .Snd1 = 0
        .Snd2 = 0
        .Snd3 = 0
        .TierraInvalida = 0
        .AtacaUsuarios = True
        .AtacaNPCs = True
        .AIAlineacion = e_Alineacion.ninguna

    End With

    '<EhFooter>
    Exit Sub

ResetNpcFlags_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.ResetNpcFlags " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub ResetNpcCounters(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ResetNpcCounters_Err

    '</EhHeader>

    With Npclist(NpcIndex).Contadores
        .Paralisis = 0
        .TiempoExistencia = 0
        .Attack = 0
        .Descanso = 0
        .Incinerado = 0
        .UseItem = 0
        .MovimientoConstante = 0
        .Velocity = 0
        .RuidoPocion = 0

    End With

    '<EhFooter>
    Exit Sub

ResetNpcCounters_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.ResetNpcCounters " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ResetNpcCharInfo_Err

    '</EhHeader>

    With Npclist(NpcIndex).Char
        .Body = 0
        .CascoAnim = 0
        .charindex = 0
        .FX = 0
        .Head = 0
        .Heading = 0
        .loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0

    End With

    '<EhFooter>
    Exit Sub

ResetNpcCharInfo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.ResetNpcCharInfo " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub ResetNpcCriatures(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ResetNpcCriatures_Err

    '</EhHeader>

    Dim j As Long
    
    With Npclist(NpcIndex)

        For j = 1 To .NroCriaturas
            .Criaturas(j).NpcIndex = 0
            .Criaturas(j).NpcName = vbNullString
        Next j
        
        .NroCriaturas = 0

    End With

    '<EhFooter>
    Exit Sub

ResetNpcCriatures_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.ResetNpcCriatures " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ResetExpresiones_Err

    '</EhHeader>

    Dim j As Long
    
    With Npclist(NpcIndex)

        For j = 1 To .NroExpresiones
            .Expresiones(j) = vbNullString
        Next j
        
        .NroExpresiones = 0

    End With

    '<EhFooter>
    Exit Sub

ResetExpresiones_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.ResetExpresiones " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '22/05/2010: ZaMa - Ahora se resetea el dueño del npc también.
    '***************************************************
    '<EhHeader>
    On Error GoTo ResetNpcMainInfo_Err

    '</EhHeader>

    With Npclist(NpcIndex)
        .Attackable = 0
        .Comercia = 0
        .GiveEXP = 0
        .GiveResource.ObjIndex = 0
        .GiveResource.Amount = 0
        .RequiredWeapon = 0
        .AntiMagia = 0
        .GiveGLD = 0
        .Hostile = 0
        .InvReSpawn = 0
        .QuestNumber = 0
        
        If .MaestroUser > 0 Then Call QuitarMascota(.MaestroUser, NpcIndex)
        If .MaestroNpc > 0 Then Call QuitarMascotaNpc(.MaestroNpc)
        If .Owner > 0 Then Call PerdioNpc(.Owner)
        
        .MaestroUser = 0
        .MaestroNpc = 0

        .Owner = 0
        .CaminataActual = 0
        .Mascotas = 0
        .Movement = 0
        .Name = vbNullString
        .NPCtype = 0
        .numero = 0
        .Orig.Map = 0
        .Orig.X = 0
        .Orig.Y = 0
        .PoderAtaque = 0
        .PoderEvasion = 0
        .Pos.Map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .SkillDomar = 0
        .Target = 0
        .TargetNPC = 0
        .TipoItems = 0
        .Veneno = 0
        .Desc = vbNullString
        
        .MenuIndex = 0
        
        .ClanIndex = 0
        
        Dim j As Long

        For j = 1 To .NroSpells
            .Spells(j) = 0
        Next j

    End With
    
    Call ResetNpcCharInfo(NpcIndex)
    Call ResetNpcCriatures(NpcIndex)
    Call ResetExpresiones(NpcIndex)
    '<EhFooter>
    Exit Sub

ResetNpcMainInfo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.ResetNpcMainInfo " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub QuitarNPC(ByVal NpcIndex As Integer, _
                     Optional ByVal RespawnTime As Boolean = False)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Now npcs lose their owner
    '***************************************************
    On Error GoTo ErrHandler

    With Npclist(NpcIndex)
        .flags.NPCActive = False
        
        If InMapBounds(.Pos.Map, .Pos.X, .Pos.Y) Then
            Call EraseNPCChar(NpcIndex)

        End If
        
        .Action = 0

    End With
    
    Call ResetNpcInv(NpcIndex)
    Call ResetNpcFlags(NpcIndex)
    Call ResetNpcCounters(NpcIndex)
    Call ResetNpcMainInfo(NpcIndex)
    
    If NpcIndex = LastNPC Then

        Do Until Npclist(LastNPC).flags.NPCActive
            LastNPC = LastNPC - 1

            If LastNPC < 1 Then Exit Do
        Loop

    End If
      
    If NumNpcs <> 0 Then
        NumNpcs = NumNpcs - 1

    End If

    Exit Sub

ErrHandler:
    Call LogError("Error en QuitarNPC")

End Sub

Public Sub QuitarPet(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)

    '***************************************************
    'Autor: ZaMa
    'Last Modification: 18/11/2009
    'Kills a pet
    '***************************************************
    On Error GoTo ErrHandler

    Dim i        As Integer

    Dim PetIndex As Integer

    With UserList(UserIndex)
        
        If .MascotaIndex Then
            .MascotaIndex = 0
            Call QuitarNPC(NpcIndex)

        End If

    End With
    
    Exit Sub

ErrHandler:
    Call LogError("Error en QuitarPet. Error: " & Err.number & " Desc: " & Err.description & " NpcIndex: " & NpcIndex & " UserIndex: " & UserIndex & " PetIndex: " & PetIndex)

End Sub

Private Function TestSpawnTrigger(Pos As WorldPos, _
                                  Optional PuedeAgua As Boolean = False) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    If LegalPos(Pos.Map, Pos.X, Pos.Y, PuedeAgua) Then
        TestSpawnTrigger = MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 3 And MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 2 And MapData(Pos.Map, Pos.X, Pos.Y).trigger <> 1

    End If
    
End Function

Public Function CrearNPC(NroNPC As Integer, _
                         mapa As Integer, _
                         OrigPos As WorldPos, _
                         Optional ByVal CustomHead As Integer, _
                         Optional ByVal ForcePos As Boolean = False) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    On Error GoTo ErrHandler

    'Crea un NPC del tipo NRONPC

    Dim Pos            As WorldPos

    Dim newpos         As WorldPos

    Dim altpos         As WorldPos

    Dim nIndex         As Integer

    Dim PosicionValida As Boolean

    Dim Iteraciones    As Long

    Dim PuedeAgua      As Boolean

    Dim PuedeTierra    As Boolean
    
    Dim tmpPos         As Long

    Dim nextPos        As Long

    Dim prevPos        As Long

    Dim TipoPos        As Byte
    
    Dim FirstValidPos  As Long
    
    Dim Map            As Integer

    Dim X              As Integer

    Dim Y              As Integer

    nIndex = OpenNPC(NroNPC, LeerNPCs) 'Conseguimos un indice
    
    If nIndex > MAXNPCS Then Exit Function
    
    ' Cabeza customizada
    If CustomHead <> 0 Then Npclist(nIndex).Char.Head = CustomHead
    
    PuedeAgua = Npclist(nIndex).flags.AguaValida
    PuedeTierra = IIf(Npclist(nIndex).flags.TierraInvalida = 1, False, True)
    
    'Necesita ser respawned en un lugar especifico
    'If InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
    'Necesita ser respawned en un lugar especifico
        
    If ((Npclist(nIndex).flags.RespawnOrigPos > 0 And Not Npclist(nIndex).flags.RespawnOrigPosRandom > 0) Or ForcePos = True) And InMapBounds(OrigPos.Map, OrigPos.X, OrigPos.Y) Then
        
        Npclist(nIndex).Orig.Map = OrigPos.Map
        Npclist(nIndex).Orig.X = OrigPos.X
        Npclist(nIndex).Orig.Y = OrigPos.Y
        Npclist(nIndex).Pos = Npclist(nIndex).Orig
       
    Else
        
        Pos.Map = mapa 'mapa
        altpos.Map = mapa
        
        If PuedeAgua = True Then
            If PuedeTierra = True Then
                TipoPos = RandomNumber(0, 1)
            Else
                TipoPos = 1

            End If

        Else
            TipoPos = 0

        End If
        
        If UBound(MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos) = 0 Then
            If TipoPos = 1 Then
                TipoPos = 0
            Else
                TipoPos = 1

            End If

        End If
        
        tmpPos = RandomNumber(1, UBound(MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos))
        
        nextPos = tmpPos
        prevPos = tmpPos
        
        Do While Not PosicionValida
                    
            ' Posición random
            If Npclist(nIndex).flags.RespawnOrigPosRandom > 0 Then
                Pos.X = RandomNumber(OrigPos.X - Npclist(nIndex).flags.RespawnOrigPosRandom, OrigPos.X + Npclist(nIndex).flags.RespawnOrigPosRandom)
                Pos.Y = RandomNumber(OrigPos.Y - Npclist(nIndex).flags.RespawnOrigPosRandom, OrigPos.Y + Npclist(nIndex).flags.RespawnOrigPosRandom)
            Else
                Pos.X = MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos(tmpPos).X
                Pos.Y = MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos(tmpPos).Y

            End If
            
            If FirstValidPos = 0 Then FirstValidPos = tmpPos
            
            If LegalPosNPC(Pos.Map, Pos.X, Pos.Y, PuedeAgua, ForcePos) And TestSpawnTrigger(Pos, PuedeAgua) Then
                
                If Not HayPCarea(Pos) Then

                    With Npclist(nIndex)
                        .Pos.Map = Pos.Map
                        .Pos.X = Pos.X
                        .Pos.Y = Pos.Y
                        .Orig = .Pos

                    End With
                    
                    PosicionValida = True

                End If

            End If
            
            If PosicionValida = False Then
                If tmpPos < nextPos Then
                    If nextPos < UBound(MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos) Then
                        nextPos = nextPos + 1
                        tmpPos = nextPos
                    Else

                        If prevPos > 1 Then
                            prevPos = prevPos - 1
                            tmpPos = prevPos
                        Else

                            If FirstValidPos > 0 Then

                                With Npclist(nIndex)
                                    .Pos.Map = Pos.Map
                                    .Pos.X = MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos(FirstValidPos).X
                                    .Pos.Y = MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos(FirstValidPos).Y
                                    .Orig = .Pos

                                End With
                                
                                PosicionValida = True
                            Else
                                Exit Function

                            End If

                        End If

                    End If

                Else

                    If prevPos > 1 Then
                        prevPos = prevPos - 1
                        tmpPos = prevPos
                    Else

                        If nextPos < UBound(MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos) Then
                            nextPos = nextPos + 1
                            tmpPos = nextPos
                        Else

                            If FirstValidPos > 0 Then

                                With Npclist(nIndex)
                                    .Pos.Map = Pos.Map
                                    .Pos.X = MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos(FirstValidPos).X
                                    .Pos.Y = MapInfo(Pos.Map).NpcSpawnPos(TipoPos).Pos(FirstValidPos).Y
                                    .Orig = .Pos

                                End With
                                
                                PosicionValida = True
                            Else
                                Exit Function

                            End If

                        End If

                    End If

                End If

            End If

        Loop
        
        'asignamos las nuevas coordenas
        Map = Pos.Map
        X = Npclist(nIndex).Pos.X
        Y = Npclist(nIndex).Pos.Y
        
        If Npclist(nIndex).flags.RespawnOrigPosRandom > 0 Then
            Npclist(nIndex).Orig.Map = Map
            Npclist(nIndex).Orig.X = X
            Npclist(nIndex).Orig.Y = Y
            Npclist(nIndex).Pos = Npclist(nIndex).Orig

        End If
        
        '  Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("map: " & Map & " x: " & X & " y:" & Y, FontTypeNames.FONTTYPE_INFO))
    End If
            
    'Crea el NPC
    Call MakeNPCChar(True, Map, nIndex, Map, X, Y)
    
    CrearNPC = nIndex
    Exit Function
  
ErrHandler:
    Call LogError("Error" & Err.number & "(" & Err.description & ") en Function CrearNPC de MODULO_NPCs.bas")

End Function

Public Sub MakeNPCChar(ByVal toMap As Boolean, _
                       ByVal sndIndex As Integer, _
                       ByVal NpcIndex As Integer, _
                       ByVal Map As Integer, _
                       ByVal X As Integer, _
                       ByVal Y As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo MakeNPCChar_Err

    '</EhHeader>
    
    Dim charindex As Integer

    Dim ValidInvi As Boolean

    Dim Name      As String

    Dim Color     As eNickColor
    
    If Npclist(NpcIndex).Char.charindex = 0 Then
        charindex = NextOpenCharIndex
        Npclist(NpcIndex).Char.charindex = charindex
        CharList(charindex) = NpcIndex

    End If
    
    MapData(Map, X, Y).NpcIndex = NpcIndex
    
    If isNPCResucitador(NpcIndex) Then
        Call Extra.SetAreaResuTheNpc(NpcIndex)

    End If
    
    With Npclist(NpcIndex)

        ' Castillo: Pretorianos del clan
        If .Hostile = 0 Then
            Name = Npclist(NpcIndex).Name

        End If
        
        Color = eNickColor.ieCastleGuild
  
        If Not toMap Then
            Call WriteCharacterCreate(sndIndex, .Char.Body, .Char.BodyAttack, .Char.Head, .Char.Heading, .Char.charindex, X, Y, .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim, Name, Color, 0, .Char.AuraIndex, .Char.speeding, .flags.NpcIdle, .numero)

        Else
            Call ModAreas.CreateEntity(NpcIndex, ENTITY_TYPE_NPC, .Pos, .SizeWidth, .SizeWidth)

        End If

    End With
    
    '<EhFooter>
    Exit Sub

MakeNPCChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.MakeNPCChar " & "at line " & Erl & " IN MAP: " & Map & " " & X & " " & Y & " name:" & Npclist(NpcIndex).Name & "."
        
    '</EhFooter>
End Sub

Public Sub ChangeNPCChar(ByVal NpcIndex As Integer, _
                         ByVal Body As Integer, _
                         ByVal Head As Integer, _
                         ByVal Heading As eHeading)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo ChangeNPCChar_Err

    '</EhHeader>

    If NpcIndex > 0 Then

        With Npclist(NpcIndex).Char
            .Body = Body
            .Head = Head
            .Heading = Heading
            
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterChange(Body, .BodyAttack, Head, Heading, .charindex, .WeaponAnim, .ShieldAnim, 0, 0, .CascoAnim, .AuraIndex, False, Npclist(NpcIndex).flags.NpcIdle, False))

        End With

    End If

    '<EhFooter>
    Exit Sub

ChangeNPCChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.ChangeNPCChar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Private Sub EraseNPCChar(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo EraseNPCChar_Err

    '</EhHeader>

    If Npclist(NpcIndex).Char.charindex <> 0 Then CharList(Npclist(NpcIndex).Char.charindex) = 0

    If Npclist(NpcIndex).Char.charindex = LastChar Then

        Do Until CharList(LastChar) > 0
            LastChar = LastChar - 1

            If LastChar <= 1 Then Exit Do
        Loop

    End If
    
    'Actualizamos el area
    Call ModAreas.DeleteEntity(NpcIndex, ENTITY_TYPE_NPC)

    'Quitamos del mapa
    MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0

    'Update la lista npc
    Npclist(NpcIndex).Char.charindex = 0

    'update NumChars
    NumChars = NumChars - 1

    '<EhFooter>
    Exit Sub

EraseNPCChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.EraseNPCChar " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte) As Boolean

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 06/04/2009
    '06/04/2009: ZaMa - Now npcs can force to change position with dead character
    '01/08/2009: ZaMa - Now npcs can't force to chance position with a dead character if that means to change the terrain the character is in
    '26/09/2010: ZaMa - Turn sub into function to know if npc has moved or not.
    '***************************************************
    '<EhHeader>
    On Error GoTo MoveNPCChar_Err

    '</EhHeader>

    Dim nPos               As WorldPos

    Dim UserIndex          As Integer

    Dim isZonaOscura       As Boolean

    Dim isZonaOscuraNewPos As Boolean
    
    With Npclist(NpcIndex)
        nPos = .Pos
        Call HeadtoPos(nHeading, nPos)
            
        ' es una posicion legal
        If LegalPosNPC(nPos.Map, nPos.X, nPos.Y, .flags.AguaValida = 1, .MaestroUser <> 0, .flags.TierraInvalida) Then
            
            If .flags.AguaValida = 0 And HayAgua(.Pos.Map, nPos.X, nPos.Y) Then Exit Function
            If .flags.TierraInvalida = 1 And Not HayAgua(.Pos.Map, nPos.X, nPos.Y) Then Exit Function
            
            isZonaOscura = (MapData(.Pos.Map, .Pos.X, .Pos.Y).trigger = eTrigger.zonaOscura)
            isZonaOscuraNewPos = (MapData(nPos.Map, nPos.X, nPos.Y).trigger = eTrigger.zonaOscura)
            
            UserIndex = MapData(.Pos.Map, nPos.X, nPos.Y).UserIndex

            ' Si hay un usuario a donde se mueve el npc, entonces esta muerto
            If UserIndex > 0 Then
                
                ' No se traslada caspers de agua a tierra
                If HayAgua(.Pos.Map, nPos.X, nPos.Y) And Not HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then Exit Function

                ' No se traslada caspers de tierra a agua
                If Not HayAgua(.Pos.Map, nPos.X, nPos.Y) And HayAgua(.Pos.Map, .Pos.X, .Pos.Y) Then Exit Function
                
                With UserList(UserIndex)
                    ' Actualizamos posicion y mapa
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = 0
                    .Pos.X = Npclist(NpcIndex).Pos.X
                    .Pos.Y = Npclist(NpcIndex).Pos.Y
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).UserIndex = UserIndex
                        
                    ' Si es un admin invisible, no se avisa a los demas clientes
                    If Not (.flags.AdminInvisible = 1) Then
                        Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(.Char.charindex, .Pos.X, .Pos.Y))
                    
                        'Los valores de visible o invisible están invertidos porque estos flags son del NpcIndex, por lo tanto si el npc entra, el casper sale y viceversa :P
                        If isZonaOscura Then
                            If Not isZonaOscuraNewPos Then
                                Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.charindex, True))

                            End If

                        Else

                            If isZonaOscuraNewPos Then
                                Call SendData(SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs, UserIndex, PrepareMessageSetInvisible(.Char.charindex, False))

                            End If

                        End If

                    End If
                    
                    nHeading = InvertHeading(nHeading)
                    
                    'Forzamos al usuario a moverse
                    Call WriteForceCharMove(UserIndex, nHeading)
                    
                    'Actualizamos las áreas de ser necesario
                    Call ModAreas.UpdateEntity(UserIndex, ENTITY_TYPE_PLAYER, .Pos)

                End With

            End If
                
            'Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterMove(.Char.charindex, nPos.X, nPos.Y))
                
            'Update map and user pos
            MapData(.Pos.Map, .Pos.X, .Pos.Y).NpcIndex = 0
            .Pos = nPos
            .Char.Heading = nHeading
            .LastHeading = nHeading
            MapData(.Pos.Map, nPos.X, nPos.Y).NpcIndex = NpcIndex
            
            If isZonaOscura Then
                If Not isZonaOscuraNewPos Then
                    If (.flags.Invisible = 0) Then
                        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageSetInvisible(.Char.charindex, False))

                    End If

                End If

            Else

                If isZonaOscuraNewPos Then
                    If (.flags.Invisible = 0) Then
                        Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageSetInvisible(.Char.charindex, True))

                    End If

                End If

            End If
            
            Call ModAreas.UpdateEntity(NpcIndex, ENTITY_TYPE_NPC, .Pos)
        
            ' Npc has moved
            MoveNPCChar = True

        End If

    End With
    
    '<EhFooter>
    Exit Function

MoveNPCChar_Err:
    LogError Err.description & vbCrLf & "in MoveNPCChar " & "at line " & Erl

    '</EhFooter>
End Function

Function NextOpenNPC() As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo ErrHandler

    Dim LoopC As Long
      
    For LoopC = 1 To MAXNPCS + 1

        If LoopC > MAXNPCS Then Exit For
        If Not Npclist(LoopC).flags.NPCActive Then Exit For
    Next LoopC
      
    NextOpenNPC = LoopC

    Exit Function

ErrHandler:
    Call LogError("Error en NextOpenNPC")

End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 10/07/2010
    '10/07/2010: ZaMa - Now npcs can't poison dead users.
    '***************************************************
    '<EhHeader>
    On Error GoTo NpcEnvenenarUser_Err

    '</EhHeader>

    Dim N As Integer
    
    With UserList(UserIndex)

        If .flags.Muerto = 1 Then Exit Sub
        If .flags.Envenenado = 1 Then Exit Sub
        
        N = RandomNumber(1, 100)

        If N < 30 Then
            .flags.Envenenado = 1
            Call WriteConsoleMsg(UserIndex, "¡¡La criatura te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteUpdateEffect(UserIndex)

        End If

    End With
    
    '<EhFooter>
    Exit Sub

NpcEnvenenarUser_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.NpcEnvenenarUser " & "at line " & Erl
        
    '</EhFooter>
End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, _
                  Pos As WorldPos, _
                  ByVal FX As Boolean, _
                  ByVal Respawn As Boolean) As Integer

    '<EhHeader>
    On Error GoTo SpawnNpc_Err

    '</EhHeader>

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 06/15/2008
    '23/01/2007 -> Pablo (ToxicWaste): Creates an NPC of the type Npcindex
    '06/15/2008 -> Optimizé el codigo. (NicoNZ)
    '***************************************************
    Dim newpos         As WorldPos

    Dim altpos         As WorldPos

    Dim nIndex         As Integer

    Dim PosicionValida As Boolean

    Dim PuedeAgua      As Boolean

    Dim PuedeTierra    As Boolean

    Dim Map            As Integer

    Dim X              As Integer

    Dim Y              As Integer

    nIndex = OpenNPC(NpcIndex, LeerNPCs, Respawn)   'Conseguimos un indice

    If nIndex > MAXNPCS Then
        SpawnNpc = 0
        Exit Function

    End If

    PuedeAgua = Npclist(nIndex).flags.AguaValida
    PuedeTierra = Not Npclist(nIndex).flags.TierraInvalida = 1
        
    Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra) 'Nos devuelve la posicion valida mas cercana
    Call ClosestLegalPos(Pos, altpos, PuedeAgua)
    'Si X e Y son iguales a 0 significa que no se encontro posicion valida

    If newpos.X <> 0 And newpos.Y <> 0 Then
        'Asignamos las nuevas coordenas solo si son validas
        Npclist(nIndex).Pos.Map = newpos.Map
        Npclist(nIndex).Pos.X = newpos.X
        Npclist(nIndex).Pos.Y = newpos.Y
        PosicionValida = True
    Else

        If altpos.X <> 0 And altpos.Y <> 0 Then
            Npclist(nIndex).Pos.Map = altpos.Map
            Npclist(nIndex).Pos.X = altpos.X
            Npclist(nIndex).Pos.Y = altpos.Y
            PosicionValida = True
        Else
            PosicionValida = False

        End If

    End If

    If Not PosicionValida Then
        Call QuitarNPC(nIndex)
        SpawnNpc = 0
        Exit Function

    End If
    
    Npclist(nIndex).Orig.Map = Npclist(nIndex).Pos.Map
    Npclist(nIndex).Orig.X = Npclist(nIndex).Pos.X
    Npclist(nIndex).Orig.Y = Npclist(nIndex).Pos.Y

    'asignamos las nuevas coordenas
    Map = newpos.Map
    X = Npclist(nIndex).Pos.X
    Y = Npclist(nIndex).Pos.Y

    'Crea el NPC
    Call MakeNPCChar(True, Map, nIndex, Map, X, Y)

    If FX Then
        Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessagePlayEffect(SND_WARP, X, Y))
        Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessageCreateFX(Npclist(nIndex).Char.charindex, FXIDs.FXWARP, 0))

    End If

    SpawnNpc = nIndex

    '<EhFooter>
    Exit Function

SpawnNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.SpawnNpc " & "at line " & Erl
        
    '</EhFooter>
End Function

Sub RespawnNpc(MiNPC As Npc)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo RespawnNpc_Err

    '</EhHeader>

    If (MiNPC.flags.Respawn = 0) Then Call CrearNPC(MiNPC.numero, MiNPC.Pos.Map, MiNPC.Orig)

    '<EhFooter>
    Exit Sub

RespawnNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.RespawnNpc " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Function OpenNPC(ByVal NpcNumber As Integer, _
                        ByRef ARCHIVE As clsIniManager, _
                        Optional ByVal Respawn = True) As Integer

    On Error GoTo OpenNPC_Err

    Dim NpcIndex As Integer
        
    Dim Field()  As String
        
    Dim Leer     As clsIniManager

    Dim LoopC    As Long

    Dim ln       As String
    
    Set Leer = LeerNPCs
        
    Dim Cabecera As String

    Cabecera = "NPC" & NpcNumber
        
    ' If requested index is invalid, abort
    If Not Leer.KeyExists(Cabecera) Then
        OpenNPC = MAXNPCS + 1

        Exit Function

    End If
    
    NpcIndex = NextOpenNPC
    
    If NpcIndex > MAXNPCS Then 'Limite de npcs
        OpenNPC = NpcIndex

        Exit Function

    End If
    
    With Npclist(NpcIndex)
        ' News
            
        ' Posición utilizada para:
        ' 1° Posición AFK
        ln = Leer.GetValue("NPC" & NpcNumber, "POSA")
        .PosA.Map = val(ReadField(1, ln, Asc("-")))
        .PosA.X = val(ReadField(2, ln, Asc("-")))
        .PosA.Y = val(ReadField(3, ln, Asc("-")))
            
        ' 2° Posicion Movimiento
        ln = Leer.GetValue("NPC" & NpcNumber, "POSB")
        .PosB.Map = val(ReadField(1, ln, Asc("-")))
        .PosB.X = val(ReadField(2, ln, Asc("-")))
        .PosB.Y = val(ReadField(3, ln, Asc("-")))
            
        ' 3° Posicion de Ataque
        ln = Leer.GetValue("NPC" & NpcNumber, "POSC")
        .PosC.Map = val(ReadField(1, ln, Asc("-")))
        .PosC.X = val(ReadField(2, ln, Asc("-")))
        .PosC.Y = val(ReadField(3, ln, Asc("-")))
        ' End News
        .numero = NpcNumber
        .Name = Leer.GetValue("NPC" & NpcNumber, "Name")
        .Desc = Leer.GetValue("NPC" & NpcNumber, "Desc")
        
        .Movement = val(Leer.GetValue("NPC" & NpcNumber, "Movement"))
        .flags.OldMovement = .Movement
        
        .flags.AguaValida = val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
        .flags.TierraInvalida = val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
        .flags.Faccion = val(Leer.GetValue("NPC" & NpcNumber, "Faccion"))
        .flags.AtacaDoble = val(Leer.GetValue("NPC" & NpcNumber, "AtacaDoble"))
        
        .NPCtype = val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))
        
        .Char.Body = val(Leer.GetValue("NPC" & NpcNumber, "Body"))
        .Char.BodyAttack = val(Leer.GetValue("NPC" & NpcNumber, "BodyAttack"))
        '.Char.AuraIndex(5) = val(Leer.GetValue("NPC" & NpcNumber, "AuraIndex"))
        .Char.Head = val(Leer.GetValue("NPC" & NpcNumber, "Head"))
        .Char.Heading = val(Leer.GetValue("NPC" & NpcNumber, "Heading"))
        .Char.BodyIdle = val(Leer.GetValue("NPC" & NpcNumber, "BodyIdle"))
            
        If .Char.BodyIdle = 0 Then .Char.BodyIdle = .Char.Body
              
        .Char.WeaponAnim = val(Leer.GetValue("NPC" & NpcNumber, "WeaponAnim"))
        .Char.ShieldAnim = val(Leer.GetValue("NPC" & NpcNumber, "ShieldAnim"))
        .Char.CascoAnim = val(Leer.GetValue("NPC" & NpcNumber, "CascoAnim"))
        
        .Attackable = val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
        .Comercia = val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
        .Hostile = val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
        .flags.OldHostil = .Hostile
        
        .GiveEXP = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP")) * MultExp
        .flags.ExpCount = .GiveEXP
        
        .Distancia = val(Leer.GetValue("NPC" & NpcNumber, "Distancia"))
        .GiveEXPGuild = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXPGuild"))
        .flags.ExpGuildCount = .GiveEXPGuild
        
        ' Recursos de la Criatura
        ln = Leer.GetValue("NPC" & NpcNumber, "GiveResource")
        .GiveResource.ObjIndex = val(ReadField(1, ln, 45))
        .GiveResource.Amount = val(ReadField(2, ln, 45))
              
        .flags.ResourceCount = .GiveResource.Amount
        .RequiredWeapon = val(Leer.GetValue("NPC" & NpcNumber, "RequiredWeapon"))
        .AntiMagia = val(Leer.GetValue("NPC" & NpcNumber, "AntiMagia"))
        
        ' Necesita un Arma Especifica para Atacar
         
        .Velocity = val(Leer.GetValue("NPC" & NpcNumber, "Velocity"))

        If .Velocity = 0 Then
            .Velocity = 380
            .Char.speeding = frmMain.TIMER_AI.interval / 330
        Else
                  
            .Char.speeding = 210 / .Velocity

            '
        End If
            
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< PATHFINDING >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        .pathFindingInfo.RangoVision = val(Leer.GetValue("NPC" & NpcNumber, "Distancia"))

        If .pathFindingInfo.RangoVision = 0 Then .pathFindingInfo.RangoVision = RANGO_VISION_x
            
        .pathFindingInfo.Inteligencia = val(Leer.GetValue("NPC" & NpcNumber, "Inteligencia"))

        If .pathFindingInfo.Inteligencia = 0 Then .pathFindingInfo.Inteligencia = 10
            
        ReDim .pathFindingInfo.Path(1 To .pathFindingInfo.Inteligencia + RANGO_VISION_x * 3)

        .IntervalAttack = val(Leer.GetValue("NPC" & NpcNumber, "IntervalAttack"))

        If .IntervalAttack = 0 Then .IntervalAttack = 1500
        
        .Veneno = val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))
        .Level = val(Leer.GetValue("NPC" & NpcNumber, "ELV"))
        .flags.Domable = val(Leer.GetValue("NPC" & NpcNumber, "Domable"))
        
        .GiveGLD = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLD")) * MultGld
        .QuestNumber = val(Leer.GetValue("NPC" & NpcNumber, "QuestNumber"))
        
        .PoderAtaque = val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
        .PoderEvasion = val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))
        
        .InvReSpawn = val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))
        
        .MonturaIndex = val(Leer.GetValue("NPC" & NpcNumber, "MonturaIndex"))
        .ShowName = val(Leer.GetValue("NPC" & NpcNumber, "ShowName"))
        
        With .Stats
            .MaxHp = val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
            .MinHp = val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
            .MaxHit = val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
            .MinHit = val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
            .def = val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
            .defM = val(Leer.GetValue("NPC" & NpcNumber, "DEFm"))

        End With
            
        .flags.AIAlineacion = val(Leer.GetValue("NPC" & NpcNumber, "Alineacion"))
        ' Forma de atacar de la criatura
        .PretorianAI = val(Leer.GetValue("NPC" & NpcNumber, "PretorianAI"))
            
        .CastleIndex = val(Leer.GetValue("NPC" & NpcNumber, "CastleIndex"))
            
        .Quest = val(Leer.GetValue("NPC" & NpcNumber, "Quest"))
        
        If .Quest > 0 Then
            ReDim .Quests(1 To .Quest) As Byte
            
            ln = Leer.GetValue("NPC" & NpcNumber, "Quests")
            
            For LoopC = 1 To .Quest
                .Quests(LoopC) = val(ReadField(LoopC, ln, 45))
            Next LoopC
            
        End If
        
        .Invent.NroItems = val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))

        If .Invent.NroItems > 0 Then

            For LoopC = 1 To .Invent.NroItems
                ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & LoopC)
                .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
                .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
    
            Next LoopC

        End If
        
        .NroDrops = val(Leer.GetValue("NPC" & NpcNumber, "NRODROPS"))
        
        If .NroDrops > 0 Then

            For LoopC = 1 To .NroDrops
                ln = Leer.GetValue("NPC" & NpcNumber, "Drop" & LoopC)
                .Drop(LoopC).ObjIndex = val(ReadField(1, ln, 45))
                .Drop(LoopC).Amount = val(ReadField(2, ln, 45))
                .Drop(LoopC).Probability = val(ReadField(3, ln, 45))
                .Drop(LoopC).ProbNum = val(ReadField(4, ln, 45))
            Next LoopC

        End If
            
        .flags.LanzaSpells = val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))

        If .flags.LanzaSpells > 0 Then ReDim .Spells(1 To .flags.LanzaSpells)

        For LoopC = 1 To .flags.LanzaSpells
            .Spells(LoopC) = val(Leer.GetValue("NPC" & NpcNumber, "Sp" & LoopC))
        Next LoopC
        
        If .NPCtype = eNPCType.Entrenador Then
            .NroCriaturas = val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))
            ReDim .Criaturas(1 To .NroCriaturas) As tCriaturasEntrenador

            For LoopC = 1 To .NroCriaturas
                .Criaturas(LoopC).NpcIndex = Leer.GetValue("NPC" & NpcNumber, "CI" & LoopC)
                .Criaturas(LoopC).NpcName = Leer.GetValue("NPC" & NpcNumber, "CN" & LoopC)
            Next LoopC

        End If
        
        With .flags
            .NPCActive = True
            
            If Respawn Then
                .Respawn = val(Leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
            Else
                .Respawn = 1

            End If
            
            .RespawnTime = val(Leer.GetValue("NPC" & NpcNumber, "RespawnTime"))

            .BackUp = val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
            .RespawnOrigPos = val(Leer.GetValue("NPC" & NpcNumber, "OrigPos"))
            .RespawnOrigPosRandom = val(Leer.GetValue("NPC" & NpcNumber, "OrigPosRandom"))
            .AfectaParalisis = val(Leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
                        
            .Snd1 = val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
            .Snd2 = val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
            .Snd3 = val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))

        End With
        
        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        .NroExpresiones = val(Leer.GetValue("NPC" & NpcNumber, "NROEXP"))

        If .NroExpresiones > 0 Then ReDim .Expresiones(1 To .NroExpresiones) As String

        For LoopC = 1 To .NroExpresiones
            .Expresiones(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Exp" & LoopC)
        Next LoopC

        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        
        ' Menu desplegable p/npc
        Select Case .NPCtype

            Case eNPCType.Banquero
                .MenuIndex = eMenues.ieBanquero
                
            Case eNPCType.Entrenador
                .MenuIndex = eMenues.ieEntrenador
                
            Case eNPCType.Gobernador
                .MenuIndex = eMenues.ieGobernador
                
            Case eNPCType.Noble
                .MenuIndex = eMenues.ieEnlistadorFaccion
                
            Case eNPCType.ResucitadorNewbie, eNPCType.Revividor
                .MenuIndex = eMenues.ieSacerdote
                
            Case eNPCType.Timbero
                .MenuIndex = eMenues.ieApostador
                
            Case Else

                If .flags.Domable <> 0 Then
                    .MenuIndex = eMenues.ieNpcDomable

                End If

        End Select
        
        If .Comercia = 1 Then .MenuIndex = eMenues.ieComerciante
        
        'Tipo de items con los que comercia
        .TipoItems = val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))
        
        .Ciudad = val(Leer.GetValue("NPC" & NpcNumber, "Ciudad"))
        .SizeWidth = CByte(val(Leer.GetValue("NPC" & NpcNumber, "SizeWidth")))
        .SizeHeight = CByte(val(Leer.GetValue("NPC" & NpcNumber, "SizeHeight")))
                
        If .SizeWidth = 0 Then .SizeWidth = ModAreas.DEFAULT_ENTITY_WIDTH
        If .SizeHeight = 0 Then .SizeHeight = ModAreas.DEFAULT_ENTITY_HEIGHT
        
        .EventIndex = CByte(val(Leer.GetValue("NPC" & NpcNumber, "EventIndex")))
            
        ' Por defecto la animación es idle
        If NumUsers > 0 Then
            Call AnimacionIdle(NpcIndex, True)

        End If
            
        ' Si el tipo de movimiento es Caminata
        If .Movement = Caminata Then

            ' Leemos la cantidad de indicaciones
            Dim cant As Byte

            cant = val(Leer.GetValue("NPC" & NpcNumber, "CaminataLen"))

            ' Prevengo NPCs rotos
            If cant = 0 Then
                .Movement = Estatico
            Else
                ' Redimenciono el array
                ReDim .Caminata(1 To cant)
                    
                ' Leo todas las indicaciones
                For LoopC = 1 To cant
                    Field = Split(Leer.GetValue("NPC" & NpcNumber, "Caminata" & LoopC), ":")
    
                    .Caminata(LoopC).offset.X = val(Field(0))
                    .Caminata(LoopC).offset.Y = val(Field(1))
                    .Caminata(LoopC).Espera = val(Field(2))
                Next
                    
                .CaminataActual = 1

            End If

        End If

        If .NroDrops Then
            .TempDrops = NPC_LISTAR_ITEMS(NpcIndex)

        End If

    End With
        
    'Update contadores de NPCs
    If NpcIndex > LastNPC Then LastNPC = NpcIndex
    NumNpcs = NumNpcs + 1
    
    'Devuelve el nuevo Indice
    OpenNPC = NpcIndex
    '<EhFooter>
    Exit Function

OpenNPC_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.OpenNPC " & "at line " & Erl
        
    '</EhFooter>
End Function

Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)
        
    On Error GoTo 0
        
    With Npclist(NpcIndex)
    
        If .flags.Follow Then
        
            .flags.AttackedBy = vbNullString
            .Target = 0
            .flags.Follow = False
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
   
        Else
        
            .flags.AttackedBy = UserName
            .Target = NameIndex(UserName)
            .flags.Follow = True
            .Movement = TipoAI.NpcDefensa
            .Hostile = 0

        End If
    
    End With
        
    Exit Sub

DoFollow_Err:
        
End Sub

Public Sub FollowAmo(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    '<EhHeader>
    On Error GoTo FollowAmo_Err

    '</EhHeader>

    With Npclist(NpcIndex)
        .flags.Follow = True
        .Movement = TipoAI.SigueAmo
        .Hostile = 0
        .Target = 0
        .TargetNPC = 0

    End With

    '<EhFooter>
    Exit Sub

FollowAmo_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.FollowAmo " & "at line " & Erl
        
    '</EhFooter>
End Sub

Public Sub ValidarPermanenciaNpc(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'Chequea si el npc continua perteneciendo a algún usuario
    '***************************************************
    '<EhHeader>
    On Error GoTo ValidarPermanenciaNpc_Err

    '</EhHeader>

    With Npclist(NpcIndex)

        If IntervaloPerdioNpc(.Owner) Then Call PerdioNpc(.Owner)

    End With

    '<EhFooter>
    Exit Sub

ValidarPermanenciaNpc_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.ValidarPermanenciaNpc " & "at line " & Erl
        
    '</EhFooter>
End Sub

Sub AnimacionIdle(ByVal NpcIndex As Integer, ByVal Show As Boolean)
    
    On Error GoTo Handler
    
    With Npclist(NpcIndex)
    
        If .Char.BodyIdle = 0 Then Exit Sub
        
        If .flags.NpcIdle = Show Then Exit Sub

        .flags.NpcIdle = Show
        
        Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, .Char.Heading)
        
    End With
    
    Exit Sub
Handler:

End Sub

' Autor: WyroX - 20/01/2021
' Intenta moverlo hacia un "costado" según el heading indicado. Se usa para mover NPCs del camino de otro char.
' Si no hay un lugar válido a los lados, lo mueve a la posición válida más cercana.
Sub MoveNpcToSide(ByVal NpcIndex As Integer, ByVal Heading As eHeading)

    On Error GoTo Handler

    With Npclist(NpcIndex)

        ' Elegimos un lado al azar
        Dim r As Integer

        r = RandomNumber(0, 1) * 2 - 1 ' -1 o 1

        ' Roto el heading original hacia ese lado
        Heading = Rotate_Heading(Heading, r)

        ' Intento moverlo para ese lado
        If MoveNPCChar(NpcIndex, Heading) Then Exit Sub
        
        ' Si falló, intento moverlo para el lado opuesto
        Heading = InvertHeading(Heading)

        If MoveNPCChar(NpcIndex, Heading) Then Exit Sub
        
        ' Si ambos fallan, entonces lo dejo en la posición válida más cercana
        Dim NuevaPos As WorldPos

        Call ClosestLegalPos(.Pos, NuevaPos, .flags.AguaValida = 1, .flags.TierraInvalida = 0)
        Call WarpNpcChar(NpcIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

    End With

    Exit Sub
    
Handler:

End Sub

Sub WarpNpcChar(ByVal NpcIndex As Integer, _
                ByVal Map As Byte, _
                ByVal X As Integer, _
                ByVal Y As Integer, _
                Optional ByVal FX As Boolean = False)

    '<EhHeader>
    On Error GoTo WarpNpcChar_Err

    '</EhHeader>

    Dim NuevaPos  As WorldPos

    Dim FuturePos As WorldPos

    Call EraseNPCChar(NpcIndex)

    FuturePos.Map = Map
    FuturePos.X = X
    FuturePos.Y = Y
    Call ClosestLegalPos(FuturePos, NuevaPos, True, True)

    If NuevaPos.Map = 0 Or NuevaPos.X = 0 Or NuevaPos.Y = 0 Then
        Debug.Print "Error al tepear NPC"
        Call QuitarNPC(NpcIndex)
    Else
        Npclist(NpcIndex).Pos = NuevaPos
        Call MakeNPCChar(True, 0, NpcIndex, NuevaPos.Map, NuevaPos.X, NuevaPos.Y)

        If FX Then                                    'FX
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayEffect(SND_WARP, NuevaPos.X, NuevaPos.Y))
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateFX(Npclist(NpcIndex).Char.charindex, FXIDs.FXWARP, 0))

        End If

    End If

    '<EhFooter>
    Exit Sub

WarpNpcChar_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.NPCs.WarpNpcChar " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

