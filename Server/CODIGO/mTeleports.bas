Attribute VB_Name = "mTeleports"
Option Explicit

Public Const TELEPORTS_DELAY_INVOKER As Long = 5000

Public Type tTeleportsCounters

    Duration As Long
    Invocation As Long
        
End Type

Public Type tTeleports

    Active As Boolean
    ObjIndex As Integer
    UserIndex As Integer
    
    PositionInvoker As WorldPos         ' Posición donde comienza a crear el teleports, se pueden posicionar npcs y usuarios, y eso hace que altere la pos del warpeo final.
    Position  As WorldPos                   ' Posición donde aparece el Teleport.
    PositionWarp As WorldPos            ' Posición donde te lleva el Teleport.
    Counters As tTeleportsCounters
    
    TeleportObj As Integer                 ' Teleport objeto que va a utilizar.
    FxInvoker As Integer                    ' Animación mientras se crea el Teleport

    CanGuild As Boolean
    CanParty As Boolean

End Type

Public Const TELEPORT_MAX_SPAWN           As Byte = 100       ' Máximo de Teleports que hay en el mundo.

Public Teleports(1 To TELEPORT_MAX_SPAWN) As tTeleports

' @ Busca un slot libre para poder crear el teleport
Private Function Teleports_FreeSlot() As Integer

    '<EhHeader>
    On Error GoTo Teleports_FreeSlot_Err

    '</EhHeader>
    Dim A As Long
    
    For A = 1 To TELEPORT_MAX_SPAWN

        With Teleports(A)

            If .Active = False Then
                Teleports_FreeSlot = A
                Exit Function

            End If

        End With
    
    Next A

    '<EhFooter>
    Exit Function

Teleports_FreeSlot_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mTeleports.Teleports_FreeSlot " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Function

Public Sub Teleports_Loop()

    '<EhHeader>
    On Error GoTo Teleports_Loop_Err

    '</EhHeader>

    Dim A As Long
    
    For A = 1 To TELEPORT_MAX_SPAWN

        With Teleports(A)

            If .Active Then

                ' @ Tiempo que tarda el Teleport en aparecer en el mapa.
                If .Counters.Invocation > 0 Then
                    .Counters.Invocation = .Counters.Invocation - 1

                    'Call SendData(SendTarget.ToPCArea, .UserIndex, PrepareMessageUpdateBar(UserList(.UserIndex).Char.CharIndex, eTypeBar.eTeleportInvoker, .Counters.Invocation, ObjData(.ObjIndex).TimeWarp))
                    Call SendData(SendTarget.ToPCArea, .UserIndex, PrepareMessageUpdateBarTerrain(.Position.X, .Position.Y, eTypeBar.eTeleportInvoker, .Counters.Invocation, ObjData(.ObjIndex).TimeWarp))
                     
                    If .Counters.Invocation = 0 Then
                        Call Teleports_Spawn(A)

                    End If
                
                Else

                    ' @ Duración del Teleport hasta que desaparece.
                    If .Counters.Duration > 0 Then
                        .Counters.Duration = .Counters.Duration - 1
                        
                        If .Counters.Duration = 0 Then
                            Call Teleports_Remove(A)

                        End If

                    End If

                End If
            
            End If
        
        End With
    
    Next A

    '<EhFooter>
    Exit Sub

Teleports_Loop_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mTeleports.Teleports_Loop " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Private Sub Teleports_DeterminateSoundWarp(ByVal UserIndex As Integer, _
                                           ByVal ObjIndex As Integer, _
                                           ByVal SourceX As Integer, _
                                           ByVal SourceY As Integer)

    '<EhHeader>
    On Error GoTo Teleports_DeterminateSoundWarp_Err

    '</EhHeader>
    
    Dim Sound As Integer
    
    With ObjData(ObjIndex)

        Select Case .TimeWarp
        
            Case 11
                Sound = eSound.sWarp10s

            Case 21
                Sound = eSound.sWarp20s

            Case 31
                Sound = eSound.sWarp30s

            Case 61
                Sound = eSound.sWarp60s

            Case Else
                Exit Sub
        
        End Select
        
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(Sound, SourceX, SourceY, 0, False, True))
        
    End With

    '<EhFooter>
    Exit Sub

Teleports_DeterminateSoundWarp_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mTeleports.Teleports_DeterminateSoundWarp " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Public Sub Teleports_AddNew(ByVal UserIndex As Integer, _
                            ByVal ObjIndex As Integer, _
                            ByVal Map As Integer, _
                            ByVal X As Byte, _
                            ByVal Y As Byte)

    '<EhHeader>
    On Error GoTo Teleports_AddNew_Err

    '</EhHeader>
                                            
    Dim Slot As Integer

    Dim Time As Double

    Dim nPos As WorldPos
    
    Time = GetTime
    Slot = Teleports_CheckWarp(UserIndex, Map, X, Y, ObjIndex, Time)
    
    If Slot > 0 Then

        With Teleports(Slot)
            
            nPos.Map = Map
            nPos.X = X
            nPos.Y = Y
            ClosestStablePos nPos, nPos

            If nPos.Map = 0 Or nPos.X = 0 Or nPos.Y = 0 Then Exit Sub
                  
            .Active = True
            .ObjIndex = ObjIndex
           
            .Counters.Invocation = ObjData(ObjIndex).TimeWarp
            .TeleportObj = ObjData(ObjIndex).TeleportObj

            .Position.Map = nPos.Map
            .Position.X = nPos.X
            .Position.Y = nPos.Y
            .PositionInvoker = .Position
            
            .PositionWarp.Map = ObjData(ObjIndex).Position.Map
            .PositionWarp.X = ObjData(ObjIndex).Position.X
            .PositionWarp.Y = ObjData(ObjIndex).Position.Y
            
            .UserIndex = UserIndex
            
            .FxInvoker = ObjData(ObjIndex).FX
                
            Call Teleports_DeterminateSoundWarp(UserIndex, ObjIndex, .Position.X, .Position.Y)
                
            With UserList(UserIndex)
                .flags.TeleportInvoker = Slot
                .flags.LastInvoker = GetTime
                '  .Char.loops = INFINITE_LOOPS
                ' .Char.FX = Teleports(Slot).FxInvoker
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(.Char.charindex, Teleports(Slot).FxInvoker, , , False))
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateBarTerrain(Teleports(Slot).Position.X, Teleports(Slot).Position.Y, eTypeBar.eTeleportInvoker, Teleports(Slot).Counters.Invocation, ObjData(ObjIndex).TimeWarp))
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFXMap(Teleports(Slot).Position.X, Teleports(Slot).Position.Y, ObjData(Teleports(Slot).TeleportObj).FX, -1))

            End With
        
        End With

    End If

    '<EhFooter>
    Exit Sub

Teleports_AddNew_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mTeleports.Teleports_AddNew " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

' @ Comprueba que el usuarip ueda teletransportarse
Private Function Teleports_CheckWarp(ByVal UserIndex As Integer, _
                                     ByVal Map As Integer, _
                                     ByVal X As Byte, _
                                     ByVal Y As Byte, _
                                     ByVal ObjIndex As Integer, _
                                     ByVal Time As Long) As Integer

    '<EhHeader>
    On Error GoTo Teleports_CheckWarp_Err

    '</EhHeader>
    
    With UserList(UserIndex)

        If Not InMapBounds(Map, X, Y) Then
            Exit Function

        End If
              
        If .flags.Meditando Then Exit Function
        If ObjData(ObjIndex).OBJType <> otTeleportInvoker Then Exit Function ' @ Seleccionó otro objeto despues del teleport.
        If .Pos.Map = ObjData(ObjIndex).Position.Map Then Exit Function
        If .flags.TeleportInvoker > 0 Then Exit Function ' @Esta invocando otro
        If MapData(Map, X, Y).ObjInfo.ObjIndex > 0 Then Exit Function ' @ Hay un objeto.
        If MapData(Map, X, Y).TileExit.Map > 0 Then Exit Function     ' @ Hay otro traslado
        If MapData(Map, X, Y).NpcIndex > 0 Then Exit Function ' @ Hay una criatura
        If MapData(Map, X, Y).UserIndex > 0 Then Exit Function ' @ Hay un usuario
        If MapData(Map, X, Y).Blocked > 0 Then Exit Function ' @ Está bloqueado
        If MapData(Map, X, Y).trigger > 0 Then Exit Function  ' @ Hay Trigger
        If MapData(Map, X, Y).TeleportIndex > 0 Then Exit Function  ' @ Hay otro Portal!
            
        ' Es un teleport inalcanzable
        If MapData(Map, X - 1, Y).Blocked > 0 And MapData(Map, X + 1, Y).Blocked > 0 And MapData(Map, X, Y + 1).Blocked > 0 And MapData(Map, X, Y - 1).Blocked > 0 Then Exit Function ' @ Está bloqueado
        
        If ObjData(ObjIndex).PuedeInsegura = 0 And MapInfo(.Pos.Map).Pk Then
            Call WriteConsoleMsg(UserIndex, "¡Este teleport no puede ser usado desde zona insegura.", FontTypeNames.FONTTYPE_INFORED)
            Exit Function

        End If
        
        If ObjData(ObjIndex).PuedeInsegura = 1 And MapInfo(.Pos.Map).Pk = False Then
            Call WriteConsoleMsg(UserIndex, "¡Este teleport solo puede ser usado desde zona insegura!", FontTypeNames.FONTTYPE_INFORED)
            Exit Function

        End If
        
        If ObjData(ObjIndex).LvlMin > .Stats.Elv Then
            Call WriteConsoleMsg(UserIndex, "Debes ser Nivel " & ObjData(ObjIndex).LvlMin & " para poder invocar el Portal.", FontTypeNames.FONTTYPE_INFORED)
            Exit Function

        End If
        
        If ObjData(ObjIndex).LvlMax < .Stats.Elv Then
            Call WriteConsoleMsg(UserIndex, "El portal puede ser invocado por personas inferiores la nivel " & ObjData(ObjIndex).LvlMax, FontTypeNames.FONTTYPE_INFORED)
            Exit Function

        End If
        
        If (GetTime - UserList(UserIndex).flags.LastInvoker) <= TELEPORTS_DELAY_INVOKER Then
            Call WriteConsoleMsg(UserIndex, "¡Debes esperar algunos segundos antes de volver a invocar un Portal!", FontTypeNames.FONTTYPE_INFORED)
            Exit Function

        End If
            
        If ObjData(ObjIndex).Dead = 1 And .flags.Muerto = 0 Then
            Call WriteConsoleMsg(UserIndex, "¡Este portal solo puede ser invocado estando muerto!", FontTypeNames.FONTTYPE_INFORED)
            Exit Function

        End If

        If .flags.SlotEvent > 0 Then Exit Function
        If .flags.SlotFast > 0 Then Exit Function
        If .flags.Desafiando > 0 Then Exit Function
        If .Counters.Pena > 0 Then Exit Function
            
        If ObjData(ObjIndex).LvlMin >= 25 Then
            If Not .Stats.UserSkills(eSkill.Navegacion) >= 35 Then
                Call WriteConsoleMsg(UserIndex, "¡Debes tener al menos una barca para poder viajar! Recuerda además tener la capacidad de usar la embarcación según tus skills.", FontTypeNames.FONTTYPE_INFORED)
                Exit Function

            End If
                
            If Not TieneObjetos(474, 1, UserIndex) And Not TieneObjetos(475, 1, UserIndex) And Not TieneObjetos(476, 1, UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "¡Debes tener al menos una barca para poder viajar! Recuerda además tener la capacidad de usar la embarcación según tus skills.", FontTypeNames.FONTTYPE_INFORED)
                Exit Function

            End If

        End If
            
        Teleports_CheckWarp = Teleports_FreeSlot

    End With

    '<EhFooter>
    Exit Function

Teleports_CheckWarp_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mTeleports.Teleports_CheckWarp " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Function

' @ El teleport aparece en el mapa
Private Sub Teleports_Spawn(ByVal Slot As Byte)

    '<EhHeader>
    On Error GoTo Teleports_Spawn_Err

    '</EhHeader>
    
    Dim Position    As WorldPos

    Dim nPos        As WorldPos

    Dim ObjTeleport As Obj
    
    With Teleports(Slot)

        If Not TieneObjetos(.ObjIndex, 1, .UserIndex) Then
            Call Teleports_Remove(Slot)
            Exit Sub

        End If

        If .PositionWarp.X = 0 And .PositionWarp.Y = 0 And .PositionWarp.Map > 0 Then
            .PositionWarp.X = RandomNumber(20, 80)
            .PositionWarp.Y = RandomNumber(20, 80)
        
        ElseIf .PositionWarp.X = 0 And .PositionWarp.Y = 0 And .PositionWarp.Map = 0 Then
            .PositionWarp.Map = UserList(.UserIndex).Hogar
            .PositionWarp.X = RandomNumber(20, 80)
            .PositionWarp.Y = RandomNumber(20, 80)

        End If
        
        ' @ Teleport que invoco en mi mapa
        ClosestStablePos .Position, nPos
              
        If nPos.Map = 0 Or nPos.X = 0 Or nPos.Y = 0 Then
            Call Teleports_Remove(Slot)
            Exit Sub

        End If
            
        .Counters.Duration = ObjData(.ObjIndex).TimeDuration
            
        MapData(nPos.Map, nPos.X, nPos.Y).TileExit = .PositionWarp
        .Position = nPos
        
        ObjTeleport.ObjIndex = .TeleportObj
        ObjTeleport.Amount = 1
        
        Call MakeObj(ObjTeleport, nPos.Map, nPos.X, nPos.Y)
        MapData(nPos.Map, nPos.X, nPos.Y).TeleportIndex = Slot
        
        'Quitamos del inv el item
        If ObjData(.ObjIndex).RemoveObj > 0 Then
            Call QuitarObjetos(.ObjIndex, ObjData(.ObjIndex).RemoveObj, .UserIndex)

        End If
        
        Call Teleports_Reset_Effect(Slot, .UserIndex)

        'Call SendData(SendTarget.ToPCArea, .UserIndex, PrepareMessageStopWaveMap(.Position.X, .Position.Y, False))
    End With

    '<EhFooter>
    Exit Sub

Teleports_Spawn_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mTeleports.Teleports_Spawn " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Private Sub Teleports_Remove(ByVal Slot As Byte)

    '<EhHeader>
    On Error GoTo Teleports_Remove_Err

    '</EhHeader>

    Dim TeleportNull As tTeleports
    
    With Teleports(Slot)
        
        With MapData(.Position.Map, .Position.X, .Position.Y)
            .TeleportIndex = 0
            
            If .ObjInfo.ObjIndex > 0 Then
                Call EraseObj(1, Teleports(Slot).Position.Map, Teleports(Slot).Position.X, Teleports(Slot).Position.Y)

            End If
            
            If .TileExit.Map > 0 Then
                .TileExit.Map = 0
                .TileExit.X = 0
                .TileExit.Y = 0

            End If

        End With
        
        UserList(.UserIndex).flags.TeleportInvoker = 0
        ' UserList(.UserIndex).flags.LastInvoker = 0
        UserList(.UserIndex).Char.FX = 0
        UserList(.UserIndex).Char.loops = 0
            
        Call Teleports_Reset_Effect(Slot, .UserIndex)
        
    End With

    Teleports(Slot) = TeleportNull
    
    '<EhFooter>
    Exit Sub

Teleports_Remove_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mTeleports.Teleports_Remove " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Public Sub Teleports_Cancel(ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo Teleports_Cancel_Err

    '</EhHeader>
    
    With UserList(UserIndex)

        If .flags.TeleportInvoker = 0 Then Exit Sub
        
        If Teleports(.flags.TeleportInvoker).Counters.Duration > 0 Then
            If Teleports(.flags.TeleportInvoker).Counters.Invocation = 0 Then Exit Sub

        End If
        
        ' Forzamos a parar el sonido
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageStopWaveMap(.Pos.X, .Pos.Y, True))
        Call Teleports_Reset_Effect(.flags.TeleportInvoker, UserIndex)
        Call Teleports_Remove(.flags.TeleportInvoker)
     
    End With

    '<EhFooter>
    Exit Sub

Teleports_Cancel_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mTeleports.Teleports_Cancel " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub

Private Sub Teleports_Reset_Effect(ByVal Slot As Byte, ByVal UserIndex As Integer)

    '<EhHeader>
    On Error GoTo Teleports_Reset_Effect_Err

    '</EhHeader>

    'Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateBar(UserList(UserIndex).Char.CharIndex, eTypeBar.eTeleportInvoker, 0, 0))
    
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageUpdateBarTerrain(Teleports(Slot).Position.X, Teleports(Slot).Position.Y, eTypeBar.eTeleportInvoker, 0, 0))
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageMeditateToggle(UserList(UserIndex).Char.charindex, 0, , , False))
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFXMap(Teleports(Slot).PositionInvoker.X, Teleports(Slot).PositionInvoker.Y, 0, 0))
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayEffect(0, Teleports(Slot).Position.X, Teleports(Slot).Position.Y, 0, False, True))
        
    '<EhFooter>
    Exit Sub

Teleports_Reset_Effect_Err:
    LogError Err.description & vbCrLf & "in ServidorArgentum.mTeleports.Teleports_Reset_Effect " & "at line " & Erl

    Resume Next

    '</EhFooter>
End Sub
