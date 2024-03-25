Attribute VB_Name = "ModAI"
'********************* COPYRIGHT NOTICE*********************
' Copyright (c) 2021-22 Martin Trionfetti, Pablo Marquez
' www.argentumunited.com.ar
' All rights reserved.
' Refer to licence for conditions of use.
' This copyright notice must always be left intact.
'****************** END OF COPYRIGHT NOTICE*****************
'
Option Explicit

Public Const FUEGOFATUO                    As Integer = 964

Public Const ELEMENTAL_VIENTO              As Integer = 963

Public Const ELEMENTAL_FUEGO               As Integer = 962

Public Const DIAMETRO_VISION_GUARDIAS_NPCS As Byte = 7

Public Sub NpcAI(ByVal NpcIndex As Integer)

    On Error GoTo ErrorHandler

    'Debug.Print "NPC: " & NpcList(NpcIndex).Name
    With Npclist(NpcIndex)
                
        Select Case .Movement

            Case TipoAI.Estatico
                ' Es un NPC estatico, no hace nada.
                Exit Sub

            Case TipoAI.MueveAlAzar

                If .Hostile = 1 Then
                    Call PerseguirUsuarioCercano(NpcIndex)
                Else
                    Call AI_CaminarSinRumboCercaDeOrigen(NpcIndex)

                End If

            Case TipoAI.NpcDefensa
                Call SeguirAgresor(NpcIndex)

            Case TipoAI.eNpcAtacaNpc
                Call AI_NpcAtacaNpc(NpcIndex)

            Case TipoAI.SigueAmo
                Call SeguirAmo(NpcIndex)

            Case TipoAI.Caminata
                Call HacerCaminata(NpcIndex)

            Case TipoAI.Invasion
                Call MovimientoInvasion(NpcIndex)

            Case TipoAI.GuardiaPersigueNpc
                Call AI_GuardiaPersigueNpc(NpcIndex)

            Case TipoAI.NpcDagaRusa
                Call Events_AI_DagaRusa(NpcIndex)

        End Select

    End With

    Exit Sub

ErrorHandler:
    
    Call LogError("NPC.AI " & Npclist(NpcIndex).Name & " " & Npclist(NpcIndex).MaestroNpc & " mapa:" & Npclist(NpcIndex).Pos.Map & " x:" & Npclist(NpcIndex).Pos.X & " y:" & Npclist(NpcIndex).Pos.Y & " Mov:" & Npclist(NpcIndex).Movement & " TargU:" & Npclist(NpcIndex).Target & " TargN:" & Npclist(NpcIndex).TargetNPC)

    Dim MiNPC As Npc: MiNPC = Npclist(NpcIndex)
    
    Call QuitarNPC(NpcIndex)
    Call RespawnNpc(MiNPC)

End Sub

Private Sub PerseguirUsuarioCercano(ByVal NpcIndex As Integer)

    On Error GoTo ErrorHandler

    Dim i                         As Long

    Dim UserIndex                 As Integer

    Dim UserIndexFront            As Integer

    Dim npcEraPasivo              As Boolean

    Dim agresor                   As Integer

    Dim minDistancia              As Integer

    Dim minDistanciaAtacable      As Integer

    Dim enemigoCercano            As Integer

    Dim enemigoAtacableMasCercano As Integer

    Dim distanciaOrigen           As Long
        
    ' Numero muy grande para que siempre haya un mÃƒÂ­nimo
    minDistancia = 32000
    minDistanciaAtacable = 32000

    With Npclist(NpcIndex)
        npcEraPasivo = .flags.OldHostil = 0
        .Target = 0
        .TargetNPC = 0

        If .flags.AttackedBy <> vbNullString Then
            agresor = NameIndex(.flags.AttackedBy)

        End If
            
        distanciaOrigen = Distancia(.Pos, .Orig)
            
        If UserIndex > 0 And UserIndexFront > 0 Then
            
            If NPCHasAUserInFront(NpcIndex, UserIndexFront) And EsEnemigo(NpcIndex, UserIndexFront) Then
                enemigoAtacableMasCercano = UserIndexFront
                minDistanciaAtacable = 1
                minDistancia = 1

            End If

        Else

            ' Busco algun objetivo en el area.
            Dim query()    As Collision.UUID

            Dim TotalUsers As Integer

            ' Call ModAreas.QueryObservers(NpcIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)
                
            For i = 0 To ModAreas.QueryObservers(NpcIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)

                UserIndex = query(i).Name
                    
                If UserList(UserIndex).ConnIDValida Then

                    If EsObjetivoValido(NpcIndex, UserIndex) Then

                        ' Busco el mas cercano, sea atacable o no.
                        If Distancia(UserList(UserIndex).Pos, .Pos) < minDistancia And Not (UserList(UserIndex).flags.Invisible > 0 Or UserList(UserIndex).flags.Oculto) Then
                            enemigoCercano = UserIndex
                            minDistancia = Distancia(UserList(UserIndex).Pos, .Pos)

                        End If
                            
                        ' Busco el mas cercano que sea atacable.
                        If (UsuarioAtacableConMagia(UserIndex) Or UsuarioAtacableConMelee(NpcIndex, UserIndex)) And Distancia(UserList(UserIndex).Pos, .Pos) < minDistanciaAtacable Then
                            enemigoAtacableMasCercano = UserIndex
                            minDistanciaAtacable = Distancia(UserList(UserIndex).Pos, .Pos)

                        End If
        
                    End If

                End If
        
            Next i

        End If

        ' Al terminar el `for`, puedo tener un maximo de tres objetivos distintos.
        ' Por prioridad, vamos a decidir estas cosas en orden.
        If distanciaOrigen < 40 Then
            If npcEraPasivo Then

                ' Significa que alguien le pego, y esta en modo agresivo trantando de darle.
                ' El unico objetivo que importa aca es el atacante; los demas son ignorados.
                If EnRangoVision(NpcIndex, agresor) Then .Target = agresor
    
            Else ' El NPC es hostil siempre, le quiere pegar a alguien.
    
                If minDistanciaAtacable > 0 And enemigoAtacableMasCercano > 0 Then ' Hay alguien atacable cerca
                    .Target = enemigoAtacableMasCercano
                ElseIf enemigoCercano > 0 Then ' Hay alguien cerca, pero no es atacable
                    .Target = enemigoCercano

                End If
    
            End If

        End If

        ' Si el NPC tiene un objetivo
        If .Target > 0 And EsObjetivoValido(NpcIndex, .Target) Then
                    
            'asignamos heading nuevo al NPC seg�n el Target del nuevo usuario: .Char.Heading, si la distancia es <= 1
            If (.flags.Inmovilizado + .flags.Paralizado = 0) Then
                Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, GetHeadingFromWorldPos(.Pos, UserList(.Target).Pos))

            End If

            Call AI_AtacarUsuarioObjetivo(NpcIndex)
                    
            'Si se aleja mucho saca el target y empieza a volver a casa
            If distanciaOrigen > 60 Then .Target = 0
        Else

            If .NPCtype <> eNPCType.GuardiaReal And .NPCtype <> eNPCType.GuardiasCaos Then

                Call RestoreOldMovement(NpcIndex)
                ' No encontro a nadie cerca, camina unos pasos en cualquier direccion.
                Call AI_CaminarSinRumboCercaDeOrigen(NpcIndex)

                ' # Se fija si se puede curar ?�
                Call NpcLanzaUnSpell(NpcIndex)
                   
            Else

                If distanciaOrigen > 0 Then
                    Call AI_CaminarConRumbo(NpcIndex, .Orig)
                Else

                    If .Char.Heading <> eHeading.SOUTH Then
                        Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, eHeading.SOUTH)

                    End If

                End If

            End If

        End If

    End With

    Exit Sub

ErrorHandler:

End Sub

' Cuando un NPC no tiene target y se puede mover libremente pero cerca de su lugar de origen.
' La mayoria de los NPC deberian mantenerse cerca de su posicion de origen, algunos quedaran quietos
' en su posicion y otros se moveran libremente cerca de su posicion de origen.
Private Sub AI_CaminarSinRumboCercaDeOrigen(ByVal NpcIndex As Integer)

    On Error GoTo AI_CaminarSinRumboCercaDeOrigen_Err

    With Npclist(NpcIndex)

        If .flags.Paralizado > 0 Or .flags.Inmovilizado > 0 Then
            Call AnimacionIdle(NpcIndex, True)
        ElseIf Distancia(.Pos, .Orig) > 4 Then
            Call AI_CaminarConRumbo(NpcIndex, .Orig)
        ElseIf RandomNumber(1, 6) = 3 Then
            Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
        Else
            Call AnimacionIdle(NpcIndex, True)

        End If

    End With

    Exit Sub

AI_CaminarSinRumboCercaDeOrigen_Err:
        
End Sub

' Cuando un NPC no tiene target y se tiene que mover libremente
Private Sub AI_CaminarSinRumbo(ByVal NpcIndex As Integer)

    On Error GoTo AI_CaminarSinRumbo_Err

    With Npclist(NpcIndex)

        If RandomNumber(1, 6) = 3 And .flags.Paralizado = 0 And .flags.Inmovilizado = 0 Then
            Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
        Else
            Call AnimacionIdle(NpcIndex, True)

        End If

    End With

    Exit Sub

AI_CaminarSinRumbo_Err:

End Sub

Private Sub AI_CaminarConRumbo(ByVal NpcIndex As Integer, ByRef rumbo As WorldPos)

    On Error GoTo AI_CaminarConRumbo_Err
    
    If Npclist(NpcIndex).flags.Paralizado Or Npclist(NpcIndex).flags.Inmovilizado Then
        Call AnimacionIdle(NpcIndex, True)
        Exit Sub

    End If
    
    With Npclist(NpcIndex).pathFindingInfo

        ' Si no tiene un camino calculado o si el destino cambio
        If .PathLength = 0 Or .Destination.X <> rumbo.X Or .Destination.Y <> rumbo.Y Then
            .Destination.X = rumbo.X
            .Destination.Y = rumbo.Y

            ' Recalculamos el camino
            If SeekPath(NpcIndex, True) Then
                ' Si consiguo un camino
                Call FollowPath(NpcIndex)

            End If

        Else ' Avanzamos en el camino
            Call FollowPath(NpcIndex)

        End If

    End With

    Exit Sub

AI_CaminarConRumbo_Err:

    Dim errorDescription As String

    errorDescription = Err.description & vbNewLine & " NpcIndex: " & NpcIndex & " NPCList.size= " & UBound(Npclist)

End Sub

Private Function NpcLanzaSpellInmovilizado(ByVal NpcIndex As Integer, _
                                           ByVal tIndex As Integer) As Boolean
        
    NpcLanzaSpellInmovilizado = False
    
    With Npclist(NpcIndex)

        If .flags.Inmovilizado + .flags.Paralizado > 0 Then

            Select Case .Char.Heading

                Case eHeading.NORTH

                    If .Pos.X = UserList(tIndex).Pos.X And .Pos.Y > UserList(tIndex).Pos.Y Then
                        NpcLanzaSpellInmovilizado = True
                        Exit Function

                    End If
                    
                Case eHeading.EAST

                    If .Pos.Y = UserList(tIndex).Pos.Y And .Pos.X < UserList(tIndex).Pos.X Then
                        NpcLanzaSpellInmovilizado = True
                        Exit Function

                    End If
                
                Case eHeading.SOUTH

                    If .Pos.X = UserList(tIndex).Pos.X And .Pos.Y < UserList(tIndex).Pos.Y Then
                        NpcLanzaSpellInmovilizado = True
                        Exit Function

                    End If
                
                Case eHeading.WEST

                    If .Pos.Y = UserList(tIndex).Pos.Y And .Pos.X > UserList(tIndex).Pos.X Then
                        NpcLanzaSpellInmovilizado = True
                        Exit Function

                    End If

            End Select

        Else
            NpcLanzaSpellInmovilizado = True

        End If

    End With
    
End Function

Public Function ComputeNextHeadingPos(ByVal NpcIndex As Integer) As WorldPos

    On Error Resume Next

    With Npclist(NpcIndex)
        ComputeNextHeadingPos.Map = .Pos.Map
        ComputeNextHeadingPos.X = .Pos.X
        ComputeNextHeadingPos.Y = .Pos.Y
    
        Select Case .Char.Heading

            Case eHeading.NORTH
                ComputeNextHeadingPos.Y = ComputeNextHeadingPos.Y - 1
                Exit Function
        
            Case eHeading.SOUTH
                ComputeNextHeadingPos.Y = ComputeNextHeadingPos.Y + 1
                Exit Function
        
            Case eHeading.EAST
                ComputeNextHeadingPos.X = ComputeNextHeadingPos.X + 1
                Exit Function
        
            Case eHeading.WEST
                ComputeNextHeadingPos.X = ComputeNextHeadingPos.X - 1
                Exit Function
        
        End Select

    End With

End Function

Public Function NPCHasAUserInFront(ByVal NpcIndex As Integer, _
                                   ByRef UserIndex As Integer) As Boolean

    On Error Resume Next

    Dim NextPosNPC As WorldPos
    
    If UserList(UserIndex).flags.Muerto = 1 Then
        NPCHasAUserInFront = False
        Exit Function

    End If
    
    NextPosNPC = ComputeNextHeadingPos(NpcIndex)
    UserIndex = MapData(NextPosNPC.Map, NextPosNPC.X, NextPosNPC.Y).UserIndex
    NPCHasAUserInFront = (UserIndex > 0)

End Function

Private Sub AI_AtacarUsuarioObjetivo(ByVal AtackerNpcIndex As Integer)

    On Error GoTo ErrorHandler

    Dim AtacaConMagia       As Boolean

    Dim AtacaMelee          As Boolean

    Dim EstaPegadoAlUsuario As Boolean

    Dim tHeading            As Byte

    Dim NextPosNPC          As WorldPos

    Dim AtacaAlDelFrente    As Boolean
        
    AtacaAlDelFrente = False

    With Npclist(AtackerNpcIndex)

        If .Target = 0 Then Exit Sub
              
        EstaPegadoAlUsuario = (Distancia(.Pos, UserList(.Target).Pos) <= 1)
        AtacaConMagia = .flags.LanzaSpells And modNuevoTimer.Intervalo_CriatureAttack(AtackerNpcIndex, False) And (RandomNumber(1, 100) <= 50)
             
        AtacaMelee = EstaPegadoAlUsuario And UsuarioAtacableConMelee(AtackerNpcIndex, .Target) And .flags.Paralizado = 0
        AtacaMelee = AtacaMelee And (.flags.LanzaSpells > 0 And (UserList(.Target).flags.Invisible > 0 Or UserList(.Target).flags.Oculto > 0))
        AtacaMelee = AtacaMelee Or .flags.LanzaSpells = 0
            
        ' Se da vuelta y enfrenta al Usuario
        tHeading = GetHeadingFromWorldPos(.Pos, UserList(.Target).Pos)
            
        If AtacaConMagia Then

            ' Le lanzo un Hechizo
            If NpcLanzaSpellInmovilizado(AtackerNpcIndex, .Target) Then
                Call ChangeNPCChar(AtackerNpcIndex, .Char.Body, .Char.Head, tHeading)
                Call NpcLanzaUnSpell(AtackerNpcIndex)

            End If

        ElseIf AtacaMelee Then

            Dim ChangeHeading As Boolean

            ChangeHeading = (.flags.Inmovilizado > 0 And tHeading = .Char.Heading) Or (.flags.Inmovilizado + .flags.Paralizado = 0)
                
            Dim UserIndexFront As Integer

            NextPosNPC = ComputeNextHeadingPos(AtackerNpcIndex)
            UserIndexFront = MapData(NextPosNPC.Map, NextPosNPC.X, NextPosNPC.Y).UserIndex
            AtacaAlDelFrente = (UserIndexFront > 0)
                
            If ChangeHeading Then
                Call ChangeNPCChar(AtackerNpcIndex, .Char.Body, .Char.Head, tHeading)

            End If
                
            If AtacaAlDelFrente And Not .flags.Paralizado = 1 Then
                Call AnimacionIdle(AtackerNpcIndex, True)

                If UserIndexFront > 0 Then
                    If UserList(UserIndexFront).flags.Muerto = 0 Then
                        If UserList(UserIndexFront).Faction.Status = 1 And (.NPCtype = eNPCType.GuardiaReal) Then
                                
                        Else
                            Call NpcAtacaUser(AtackerNpcIndex, UserIndexFront, tHeading)

                        End If

                    End If

                End If

            End If

        End If

        If UsuarioAtacableConMagia(.Target) Or UsuarioAtacableConMelee(AtackerNpcIndex, .Target) Then

            ' Si no tiene un camino pero esta pegado al usuario, no queremos gastar tiempo calculando caminos.
            If .pathFindingInfo.PathLength = 0 And EstaPegadoAlUsuario Then Exit Sub
            
            Call AI_CaminarConRumbo(AtackerNpcIndex, UserList(.Target).Pos)
        Else
            Call AI_CaminarSinRumboCercaDeOrigen(AtackerNpcIndex)

        End If

    End With

    Exit Sub

ErrorHandler:

End Sub

Public Sub AI_GuardiaPersigueNpc(ByVal NpcIndex As Integer)

    On Error GoTo ErrorHandler

    Dim targetPos As WorldPos
        
    With Npclist(NpcIndex)
        
        If .TargetNPC > 0 Then
            targetPos = Npclist(.TargetNPC).Pos
                
            If Distancia(.Pos, targetPos) <= 1 Then
                Call NpcAtacaNpc(NpcIndex, .TargetNPC, False)

            End If
                
            If DistanciaRadial(.Orig, targetPos) <= (DIAMETRO_VISION_GUARDIAS_NPCS \ 2) Then
                If Npclist(.TargetNPC).Target = 0 Then
                    Call AI_CaminarConRumbo(NpcIndex, targetPos)
                ElseIf UserList(Npclist(.TargetNPC).Target).flags.NPCAtacado <> .TargetNPC Then
                    Call AI_CaminarConRumbo(NpcIndex, targetPos)
                Else
                    .TargetNPC = 0
                    Call AI_CaminarConRumbo(NpcIndex, .Orig)

                End If

            Else
                .TargetNPC = 0
                Call AI_CaminarConRumbo(NpcIndex, .Orig)

            End If
                
        Else
            .TargetNPC = BuscarNpcEnArea(NpcIndex)

            If Distancia(.Pos, .Orig) > 0 Then
                Call AI_CaminarConRumbo(NpcIndex, .Orig)
            Else
                Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, eHeading.SOUTH)

            End If

        End If
            
    End With
        
    Exit Sub
        
ErrorHandler:

End Sub

Private Function DistanciaRadial(OrigenPos As WorldPos, DestinoPos As WorldPos) As Long
    DistanciaRadial = max(Abs(OrigenPos.X - DestinoPos.X), Abs(OrigenPos.Y - DestinoPos.Y))

End Function

Private Function BuscarNpcEnArea(ByVal NpcIndex As Integer) As Integer
        
    On Error GoTo BuscarNpcEnArea
        
    Dim X As Integer, Y As Integer
       
    With Npclist(NpcIndex)
       
        For X = (.Orig.X - (DIAMETRO_VISION_GUARDIAS_NPCS \ 2)) To (.Orig.X + (DIAMETRO_VISION_GUARDIAS_NPCS \ 2))
            For Y = (.Orig.Y - (DIAMETRO_VISION_GUARDIAS_NPCS \ 2)) To (.Orig.Y + (DIAMETRO_VISION_GUARDIAS_NPCS \ 2))
                
                If MapData(.Orig.Map, X, Y).NpcIndex > 0 And NpcIndex <> MapData(.Orig.Map, X, Y).NpcIndex Then

                    Dim foundNpc As Integer

                    foundNpc = MapData(.Orig.Map, X, Y).NpcIndex
                        
                    If Npclist(foundNpc).Hostile Then
                        
                        If Npclist(foundNpc).Target = 0 Then
                            BuscarNpcEnArea = MapData(.Orig.Map, X, Y).NpcIndex
                            Exit Function
                        ElseIf UserList(Npclist(foundNpc).Target).flags.NPCAtacado <> foundNpc Then
                            BuscarNpcEnArea = MapData(.Orig.Map, X, Y).NpcIndex
                            Exit Function

                        End If
                            
                    End If
                        
                End If
                    
            Next Y
        Next X

    End With
        
    BuscarNpcEnArea = 0
        
    Exit Function

BuscarNpcEnArea:

End Function

Public Sub AI_NpcAtacaNpc(ByVal NpcIndex As Integer)

    On Error GoTo ErrorHandler
    
    Dim targetPos As WorldPos
        
    Dim Distance  As Integer
        
    With Npclist(NpcIndex)
        Distance = 3
            
        If .TargetNPC > 0 Then
            targetPos = Npclist(.TargetNPC).Pos
            
            If InRangoVisionNPC(NpcIndex, targetPos.X, targetPos.Y) Then
                    
                If .flags.Paralizado = 0 Then
                    ' Me fijo si el NPC esta al lado del Objetivo
                        
                    If .flags.LanzaSpells > 0 Then
                        Call NpcLanzaUnSpell(NpcIndex)
                    Else
                            
                        If Distancia(.Pos, targetPos) <= Distance Then
                            Call NpcAtacaNpc(NpcIndex, .TargetNPC)

                        End If

                    End If
                  
                End If
                  
                If .TargetNPC <> vbNull And .TargetNPC > 0 Then
                    Call AI_CaminarConRumbo(NpcIndex, targetPos)

                End If
               
                Exit Sub

            End If

        End If
           
        Call RestoreOldMovement(NpcIndex)
 
    End With
                
    Exit Sub
                
ErrorHandler:

End Sub

Private Sub SeguirAgresor(ByVal NpcIndex As Integer)
    ' La IA que se ejecuta cuando alguien le pega al maestro de una Mascota/Elemental
    ' o si atacas a los NPCs con Movement = e_TipoAI.NpcDefensa
    ' A diferencia de IrUsuarioCercano(), aca no buscamos objetivos cercanos en el area
    ' porque ya establecemos como objetivo a el usuario que ataco a los NPC con este tipo de IA

    On Error GoTo SeguirAgresor_Err

    If EsObjetivoValido(NpcIndex, Npclist(NpcIndex).Target) Then
        Call AI_AtacarUsuarioObjetivo(NpcIndex)
    Else
        Call RestoreOldMovement(NpcIndex)

    End If

    Exit Sub

SeguirAgresor_Err:

End Sub

Public Sub SeguirAmo(ByVal NpcIndex As Integer)

    On Error GoTo ErrorHandler
        
    With Npclist(NpcIndex)
        
        If .MaestroUser = 0 Or Not .flags.Follow Then Exit Sub
        
        ' Si la mascota no tiene objetivo establecido.
        If .Target = 0 And .TargetNPC = 0 Then
            
            If EnRangoVision(NpcIndex, .MaestroUser) Then
                If UserList(.MaestroUser).flags.Muerto = 0 And UserList(.MaestroUser).flags.Invisible = 0 And UserList(.MaestroUser).flags.Oculto = 0 And Distancia(.Pos, UserList(.MaestroUser).Pos) > 3 Then
                    
                    ' Caminamos cerca del usuario
                    Call AI_CaminarConRumbo(NpcIndex, UserList(.MaestroUser).Pos)
                    Exit Sub
                    
                End If

            End If
                
            Call AI_CaminarSinRumbo(NpcIndex)

        End If

    End With
    
    Exit Sub

ErrorHandler:

End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)

    On Error GoTo RestoreOldMovement_Err

    With Npclist(NpcIndex)
        .Target = 0
        .TargetNPC = 0
        
        ' Si el NPC no tiene maestro, reseteamos el movimiento que tenia antes.
        If .MaestroUser = 0 Then
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
            .flags.AttackedBy = vbNullString
        Else
            
            ' Si tiene maestro, hacemos que lo siga.
            Call FollowAmo(NpcIndex)
            
        End If

    End With

    Exit Sub

RestoreOldMovement_Err:

End Sub

Private Sub HacerCaminata(ByVal NpcIndex As Integer)

    On Error GoTo Handler
    
    Dim Destino   As WorldPos

    Dim Heading   As eHeading

    Dim NextTile  As WorldPos

    Dim MoveChar  As Integer

    Dim PudoMover As Boolean

    With Npclist(NpcIndex)
    
        Destino.Map = .Pos.Map
        Destino.X = .Orig.X + .Caminata(.CaminataActual).offset.X
        Destino.Y = .Orig.Y + .Caminata(.CaminataActual).offset.Y

        ' Si todaviï�½a no llego al destino
        If .Pos.X <> Destino.X Or .Pos.Y <> Destino.Y Then
        
            ' Tratamos de acercarnos (podemos pisar npcs, usuarios o triggers)
            Heading = GetHeadingFromWorldPos(.Pos, Destino)

            ' Obtengo la posicion segun el heading
            NextTile = .Pos
            Call HeadtoPos(Heading, NextTile)
            
            ' Si hay un NPC
            MoveChar = MapData(NextTile.Map, NextTile.X, NextTile.Y).NpcIndex

            If MoveChar Then
                ' Lo movemos hacia un lado
                Call MoveNpcToSide(MoveChar, Heading)

            End If
            
            ' Si hay un user
            MoveChar = MapData(NextTile.Map, NextTile.X, NextTile.Y).UserIndex

            If MoveChar Then

                ' Si no esta muerto o es admin invisible (porque a esos los atraviesa)
                If UserList(MoveChar).flags.AdminInvisible = 0 Or UserList(MoveChar).flags.Muerto = 0 Then
                    ' Lo movemos hacia un lado
                    Call MoveUserToSide(MoveChar, Heading)

                End If

            End If
            
            ' Movemos al NPC de la caminata
            PudoMover = MoveNPCChar(NpcIndex, Heading)
            
            ' Si no pudimos moverlo, hacemos como si hubiese llegado a destino... para evitar que se quede atascado
            If Not PudoMover Or Distancia(.Pos, Destino) = 0 Then
            
                ' Llegamos a destino, ahora esperamos el tiempo necesario para continuar
                .Contadores.Velocity = GetTime + .Caminata(.CaminataActual).Espera - .Velocity
                
                ' Pasamos a la siguiente caminata
                .CaminataActual = .CaminataActual + 1
                
                ' Si pasamos el ultimo, volvemos al primero
                If .CaminataActual > UBound(.Caminata) Then
                    .CaminataActual = 1

                End If
                
            End If
            
            ' Si por alguna razÃƒÂ³n estamos en el destino, seguimos con la siguiente caminata
        Else
        
            .CaminataActual = .CaminataActual + 1
            
            ' Si pasamos el ultimo, volvemos al primero
            If .CaminataActual > UBound(.Caminata) Then
                .CaminataActual = 1

            End If
            
        End If
    
    End With
    
    Exit Sub
    
Handler:

End Sub

Private Sub MovimientoInvasion(ByVal NpcIndex As Integer)

    On Error GoTo Handler
    
    With Npclist(NpcIndex)

        Dim SpawnBox         As t_SpawnBox

        'SpawnBox = Invasiones(.flags.InvasionIndex).SpawnBoxes(.flags.SpawnBox)
    
        ' Calculamos la distancia a la muralla y generamos una posicion de destino
        Dim DistanciaMuralla As Integer, Destino As WorldPos

        Destino = .Pos
        
        If SpawnBox.Heading = eHeading.EAST Or SpawnBox.Heading = eHeading.WEST Then
            DistanciaMuralla = Abs(.Pos.X - SpawnBox.CoordMuralla)
            Destino.X = SpawnBox.CoordMuralla
        Else
            DistanciaMuralla = Abs(.Pos.Y - SpawnBox.CoordMuralla)
            Destino.Y = SpawnBox.CoordMuralla

        End If

        ' Si todavia esta lejos de la muralla
        If DistanciaMuralla > 1 Then
        
            ' Tratamos de acercarnos (sin pisar)
            Dim Heading As eHeading

            Heading = GetHeadingFromWorldPos(.Pos, Destino)
            
            ' Nos aseguramos que la posicion nueva esta dentro del rectangulo valido
            Dim NextTile As WorldPos

            NextTile = .Pos
            Call HeadtoPos(Heading, NextTile)
            
            ' Si la posicion nueva queda fuera del rectangulo valido
            If Not InsideRectangle(SpawnBox.LegalBox, NextTile.X, NextTile.Y) Then
                ' Invertimos la direccion de movimiento
                Heading = InvertHeading(Heading)

            End If
            
            ' Movemos el NPC
            Call MoveNPCChar(NpcIndex, Heading)
        
            ' Si esta pegado a la muralla
        Else
        
            ' Chequeamos el intervalo de ataque
            If Not Intervalo_CriatureAttack(NpcIndex, False) Then
                Exit Sub

            End If
            
            ' Nos aseguramos que mire hacia la muralla
            If .Char.Heading <> SpawnBox.Heading Then
                Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, SpawnBox.Heading)

            End If
            
            ' Sonido de ataque (si tiene)
            If .flags.Snd1 > 0 Then
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayEffect(.flags.Snd1, .Pos.X, .Pos.Y))

            End If
            
            ' Sonido de impacto
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayEffect(SND_IMPACTO, .Pos.X, .Pos.Y))
            
            ' Da�amos la muralla
            'Call HacerDa�oMuralla(.flags.InvasionIndex, RandomNumber(.Stats.MinHit, .Stats.MaxHit))  ' TODO: Defensa de la muralla? No hace falta creo...

        End If
    
    End With

    Exit Sub
    
Handler:

    Dim errorDescription As String

    'errorDescription = Err.description & vbNewLine & "NpcId=" & Npclist(NpcIndex).Numero & " InvasionIndex:" & Npclist(NpcIndex).flags.InvasionIndex & " SpawnBox:" & Npclist(NpcIndex).flags.SpawnBox & vbNewLine
    'Call TraceError(Err.Number, errorDescription, "AI.MovimientoInvasion", Erl)
End Sub

' El NPC elige un hechizo al azar dentro de su listado, con un potencial Target.
' Depdendiendo el tipo de spell que elije, se elije un target distinto que puede ser:
' - El .Target, el NPC mismo o area.
Private Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer)

    On Error GoTo NpcLanzaUnSpell_Err
        
    If Npclist(NpcIndex).flags.LanzaSpells = 0 Then Exit Sub

    ' Elegir hechizo, dependiendo del hechizo lo tiro sobre NPC, sobre Target o Sobre area (cerca de user o NPC si no tiene)
    Dim SpellIndex          As Integer

    Dim Target              As Integer

    Dim PuedeDanarAlUsuario As Boolean

    If Not Intervalo_CriatureAttack(NpcIndex, False) Then Exit Sub

    Target = Npclist(NpcIndex).Target
    SpellIndex = Npclist(NpcIndex).Spells(RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells))
    PuedeDanarAlUsuario = Npclist(NpcIndex).flags.Paralizado = 0
        
    If SpellIndex = 0 Then Exit Sub
        
    Select Case Hechizos(SpellIndex).Target

        Case TargetType.uUsuarios

            If UsuarioAtacableConMagia(Target) And PuedeDanarAlUsuario Then
                Call NpcLanzaSpellSobreUser(NpcIndex, Target, SpellIndex)

                If UserList(Target).flags.AtacadoPorNpc = 0 Then
                    UserList(Target).flags.AtacadoPorNpc = NpcIndex

                End If

            End If

        Case TargetType.uNPC

            If Hechizos(SpellIndex).AutoLanzar = 1 Then
                Call NpcLanzaSpellSobreNpc(NpcIndex, NpcIndex, SpellIndex)

            ElseIf Npclist(NpcIndex).TargetNPC > 0 Then
                Call NpcLanzaSpellSobreNpc(NpcIndex, Npclist(NpcIndex).TargetNPC, SpellIndex)

            End If

        Case TargetType.uUsuariosYnpc

            If Hechizos(SpellIndex).AutoLanzar = 1 Then
                Call NpcLanzaSpellSobreNpc(NpcIndex, NpcIndex, SpellIndex)

            ElseIf UsuarioAtacableConMagia(Target) And PuedeDanarAlUsuario Then
                Call NpcLanzaSpellSobreUser(NpcIndex, Target, SpellIndex)

                If UserList(Target).flags.AtacadoPorNpc = 0 Then
                    UserList(Target).flags.AtacadoPorNpc = NpcIndex

                End If

            ElseIf Npclist(NpcIndex).TargetNPC > 0 Then
                Call NpcLanzaSpellSobreNpc(NpcIndex, Npclist(NpcIndex).TargetNPC, SpellIndex)

            End If

        Case TargetType.uTerreno
            'Call NpcLanzaSpellSobreArea(NpcIndex, SpellIndex)

    End Select

    Exit Sub

NpcLanzaUnSpell_Err:

End Sub

Private Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer)

    On Error GoTo NpcLanzaUnSpellSobreNpc_Err
    
    With Npclist(NpcIndex)
        
        If Not Intervalo_CriatureAttack(NpcIndex, False) Then Exit Sub
        If .Pos.Map <> Npclist(TargetNPC).Pos.Map Then Exit Sub
    
        Dim K As Integer

        K = RandomNumber(1, .flags.LanzaSpells)

        Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, .Spells(K))
    
    End With
     
    Exit Sub

NpcLanzaUnSpellSobreNpc_Err:

End Sub

' ---------------------------------------------------------------------------------------------------
'                                       HELPERS
' ---------------------------------------------------------------------------------------------------

Public Function EsObjetivoValido(ByVal NpcIndex As Integer, _
                                 ByVal UserIndex As Integer) As Boolean

    If UserIndex = 0 Then Exit Function
    If NpcIndex = 0 Then Exit Function

    ' Esta condicion debe ejecutarse independiemente de el modo de busqueda.
    If UserList(UserIndex).flags.Muerto = 1 Then Exit Function 'User muerto
    If Not EnRangoVision(NpcIndex, UserIndex) Then Exit Function 'En rango
    If Not EsEnemigo(NpcIndex, UserIndex) Then Exit Function 'Es enemigo
    If UserList(UserIndex).flags.EnConsulta = 1 Then Exit Function 'En consulta
    If EsGm(UserIndex) And Not UserList(UserIndex).flags.AdminPerseguible Then Exit Function
        
    EsObjetivoValido = True

End Function

Private Function EsEnemigo(ByVal NpcIndex As Integer, _
                           ByVal UserIndex As Integer) As Boolean

    On Error GoTo EsEnemigo_Err

    If NpcIndex = 0 Or UserIndex = 0 Then Exit Function

    With Npclist(NpcIndex)

        If .flags.AttackedBy <> vbNullString Then
            EsEnemigo = (UserIndex = NameIndex(.flags.AttackedBy))

            If EsEnemigo Then Exit Function

        End If

        Select Case .flags.AIAlineacion

            Case e_Alineacion.Real
                EsEnemigo = Escriminal(UserIndex)

            Case e_Alineacion.Caos
                EsEnemigo = Not Escriminal(UserIndex)

            Case e_Alineacion.ninguna
                EsEnemigo = True
                ' Ok. No hay nada especial para hacer, cualquiera puede ser enemigo!

        End Select

    End With

    Exit Function

EsEnemigo_Err:

End Function

Private Function EnRangoVision(ByVal NpcIndex As Integer, _
                               ByVal UserIndex As Integer) As Boolean

    On Error GoTo EnRangoVision_Err

    Dim userPos  As WorldPos

    Dim NpcPos   As WorldPos

    Dim Limite_x As Integer, Limite_y As Integer

    ' Si alguno es cero, devolve false
    If NpcIndex = 0 Or UserIndex = 0 Then Exit Function

    Limite_x = IIf(Npclist(NpcIndex).Distancia <> 0, Npclist(NpcIndex).Distancia, RANGO_VISION_x)
    Limite_y = IIf(Npclist(NpcIndex).Distancia <> 0, Npclist(NpcIndex).Distancia, RANGO_VISION_y)

    userPos = UserList(UserIndex).Pos
    NpcPos = Npclist(NpcIndex).Pos

    EnRangoVision = ((userPos.Map = NpcPos.Map) And (Abs(userPos.X - NpcPos.X) <= Limite_x) And (Abs(userPos.Y - NpcPos.Y) <= Limite_y))

    Exit Function

EnRangoVision_Err:

End Function

Private Function UsuarioAtacableConMagia(ByVal targetUserIndex As Integer) As Boolean

    On Error GoTo UsuarioAtacableConMagia_Err

    If targetUserIndex = 0 Then Exit Function

    With UserList(targetUserIndex)
        UsuarioAtacableConMagia = (.flags.Muerto = 0 And .flags.Invisible = 0 And .flags.Oculto = 0 And .flags.Mimetizado < e_EstadoMimetismo.FormaBichoSinProteccion And Not (EsGm(targetUserIndex) And Not UserList(targetUserIndex).flags.AdminPerseguible) And Not .flags.EnConsulta)

    End With

    Exit Function

UsuarioAtacableConMagia_Err:

End Function

Private Function UsuarioAtacableConMelee(ByVal NpcIndex As Integer, _
                                         ByVal targetUserIndex As Integer) As Boolean

    On Error GoTo UsuarioAtacableConMelee_Err

    If targetUserIndex = 0 Then Exit Function

    Dim EstaPegadoAlUser As Boolean
    
    With UserList(targetUserIndex)
    
        EstaPegadoAlUser = Distancia(Npclist(NpcIndex).Pos, .Pos) = 1

        UsuarioAtacableConMelee = (.flags.Muerto = 0 And (EstaPegadoAlUser Or (Not EstaPegadoAlUser And (.flags.Invisible + .flags.Oculto) = 0)) And .flags.Mimetizado < e_EstadoMimetismo.FormaBichoSinProteccion And Not (EsGm(targetUserIndex) And Not UserList(targetUserIndex).flags.AdminPerseguible) And Not .flags.EnConsulta)

    End With

    Exit Function

UsuarioAtacableConMelee_Err:

End Function

