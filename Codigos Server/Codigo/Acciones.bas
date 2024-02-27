Attribute VB_Name = "Acciones"

Option Explicit

Sub Accion(ByVal Userindex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

    Dim tempIndex As Integer
    
On Error Resume Next
    '¿Rango Visión? (ToxicWaste)
    If (Abs(UserList(Userindex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(Userindex).Pos.X - X) > RANGO_VISION_X) Then
        Exit Sub
    End If
    
    '¿Posicion valida?
    If InMapBounds(map, X, Y) Then
        With UserList(Userindex)
            If MapData(map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
                tempIndex = MapData(map, X, Y).NpcIndex
                
                'Set the target NPC
                .flags.TargetNPC = tempIndex
                
                If Npclist(tempIndex).Comercia = 1 Then
                    '¿Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(Userindex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Iniciamos la rutina pa' comerciar.
                    Call IniciarComercioNPC(Userindex)
                
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Banquero Then
                    '¿Esta el user muerto? Si es asi no puede comerciar
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
                    
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(Userindex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'A depositar de una
                    Call IniciarDeposito(Userindex)
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Pirata Then
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                   
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
               
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(Userindex, "Estás demasiado lejos del pirata.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                   
                              ElseIf Npclist(tempIndex).NPCtype = eNPCType.PirataViajes Then
                    If .flags.Muerto = 1 Then
                        Call WriteConsoleMsg(Userindex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                   
                    'Is it already in commerce mode??
                    If .flags.Comerciando Then
                        Exit Sub
                    End If
               
                    If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                        Call WriteConsoleMsg(Userindex, "Estás demasiado lejos del pirata.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                   
                    Call WriteFormViajes(Userindex) 'Iniciamos el formulario
                   
                    'Call WriteFormViajes(UserIndex) 'Iniciamos el formulario
                ElseIf Npclist(tempIndex).NPCtype = eNPCType.Revividor Or Npclist(tempIndex).NPCtype = eNPCType.ResucitadorNewbie Then
                    If Distancia(.Pos, Npclist(tempIndex).Pos) > 10 Then
                        Call WriteConsoleMsg(Userindex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Revivimos si es necesario
                    If .flags.Muerto = 1 And (Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(Userindex)) Then
                        Call RevivirUsuario(Userindex)
                    End If
                    
                    If Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(Userindex) Then
                        'curamos totalmente
                        .Stats.MinHp = .Stats.MaxHp
                        Call WriteUpdateUserStats(Userindex)
                    End If
                End If
                
            '¿Es un obj?
            ElseIf MapData(map, X, Y).ObjInfo.objindex > 0 Then
                tempIndex = MapData(map, X, Y).ObjInfo.objindex
                
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, X, Y, Userindex)
                    Case eOBJType.otCarteles 'Es un cartel
                        Call AccionParaCartel(map, X, Y, Userindex)
                    Case eOBJType.otForos 'Foro
                        Call AccionParaForo(map, X, Y, Userindex)
                    Case eOBJType.otLeña    'Leña
                        If tempIndex = FOGATA_APAG And .flags.Muerto = 0 Then
                            Call AccionParaRamita(map, X, Y, Userindex)
                        End If
                End Select
            '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
            ElseIf MapData(map, X + 1, Y).ObjInfo.objindex > 0 Then
                tempIndex = MapData(map, X + 1, Y).ObjInfo.objindex
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, X + 1, Y, Userindex)
                    
                End Select
            
            ElseIf MapData(map, X + 1, Y + 1).ObjInfo.objindex > 0 Then
                tempIndex = MapData(map, X + 1, Y + 1).ObjInfo.objindex
                .flags.TargetObj = tempIndex
        
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, X + 1, Y + 1, Userindex)
                End Select
            
            ElseIf MapData(map, X, Y + 1).ObjInfo.objindex > 0 Then
                tempIndex = MapData(map, X, Y + 1).ObjInfo.objindex
                .flags.TargetObj = tempIndex
                
                Select Case ObjData(tempIndex).OBJType
                    Case eOBJType.otPuertas 'Es una puerta
                        Call AccionParaPuerta(map, X, Y + 1, Userindex)
                End Select
            End If
        End With
    End If
End Sub

Public Sub AccionParaForo(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 02/01/2010
'02/01/2010: ZaMa - Agrego foros faccionarios
'***************************************************

On Error Resume Next

    Dim Pos As WorldPos
    
    Pos.map = map
    Pos.X = X
    Pos.Y = Y
    
    If Distancia(Pos, UserList(Userindex).Pos) > 2 Then
        Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If SendPosts(Userindex, ObjData(MapData(map, X, Y).ObjInfo.objindex).ForoID) Then
        Call WriteShowForumForm(Userindex)
    End If
    
End Sub

Sub AccionParaPuerta(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

If Not (Distance(UserList(Userindex).Pos.X, UserList(Userindex).Pos.Y, X, Y) > 2) Then
    If ObjData(MapData(map, X, Y).ObjInfo.objindex).Llave = 0 Then
        If ObjData(MapData(map, X, Y).ObjInfo.objindex).Cerrada = 1 Then
                'Abre la puerta
                If ObjData(MapData(map, X, Y).ObjInfo.objindex).Llave = 0 Then
                    
                    MapData(map, X, Y).ObjInfo.objindex = ObjData(MapData(map, X, Y).ObjInfo.objindex).IndexAbierta
                    
                    Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(map, X, Y).ObjInfo.objindex).GrhIndex, X, Y))
                    
                    'Desbloquea
                    MapData(map, X, Y).Blocked = 0
                    MapData(map, X - 1, Y).Blocked = 0
                    
                    'Bloquea todos los mapas
                    Call Bloquear(True, map, X, Y, 0)
                    Call Bloquear(True, map, X - 1, Y, 0)
                    
                      
                    'Sonido
                    Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
                    
                Else
                     Call WriteConsoleMsg(Userindex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
                End If
        Else
                'Cierra puerta
                MapData(map, X, Y).ObjInfo.objindex = ObjData(MapData(map, X, Y).ObjInfo.objindex).IndexCerrada
                
                Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(map, X, Y).ObjInfo.objindex).GrhIndex, X, Y))
                                
                MapData(map, X, Y).Blocked = 1
                MapData(map, X - 1, Y).Blocked = 1
                
                
                Call Bloquear(True, map, X - 1, Y, 1)
                Call Bloquear(True, map, X, Y, 1)
                
                Call SendData(SendTarget.ToPCArea, Userindex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
        End If
        
        UserList(Userindex).flags.TargetObj = MapData(map, X, Y).ObjInfo.objindex
    Else
        Call WriteConsoleMsg(Userindex, "La puerta está cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(Userindex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
End If

End Sub

Sub AccionParaCartel(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

If ObjData(MapData(map, X, Y).ObjInfo.objindex).OBJType = 8 Then
  
  If Len(ObjData(MapData(map, X, Y).ObjInfo.objindex).texto) > 0 Then
    Call WriteShowSignal(Userindex, MapData(map, X, Y).ObjInfo.objindex)
  End If
  
End If

End Sub

Sub AccionParaRamita(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Userindex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error Resume Next

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj

Dim Pos As WorldPos
Pos.map = map
Pos.X = X
Pos.Y = Y

With UserList(Userindex)
    If Distancia(Pos, .Pos) > 2 Then
        Call WriteConsoleMsg(Userindex, "Estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If MapData(map, X, Y).trigger = eTrigger.ZONASEGURA Or MapInfo(map).Pk = False Then
        Call WriteConsoleMsg(Userindex, "No puedes hacer fogatas en zona segura.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If .Stats.UserSkills(Supervivencia) > 1 And .Stats.UserSkills(Supervivencia) < 6 Then
        Suerte = 3
    ElseIf .Stats.UserSkills(Supervivencia) >= 6 And .Stats.UserSkills(Supervivencia) <= 10 Then
        Suerte = 2
    ElseIf .Stats.UserSkills(Supervivencia) >= 10 And .Stats.UserSkills(Supervivencia) Then
        Suerte = 1
    End If
    
    exito = RandomNumber(1, Suerte)
    
    If exito = 1 Then
        If MapInfo(.Pos.map).Zona <> Ciudad Then
            Obj.objindex = FOGATA
            Obj.Amount = 1
            
            Call WriteConsoleMsg(Userindex, "Has prendido la fogata.", FontTypeNames.FONTTYPE_INFO)
            
            Call MakeObj(Obj, map, X, Y)
            
            'Las fogatas prendidas se deben eliminar
            Dim Fogatita As New cGarbage
            Fogatita.map = map
            Fogatita.X = X
            Fogatita.Y = Y
            Call TrashCollector.Add(Fogatita)
            
            Call SubirSkill(Userindex, eSkill.Supervivencia, True)
        Else
            Call WriteConsoleMsg(Userindex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        Call WriteConsoleMsg(Userindex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)
        Call SubirSkill(Userindex, eSkill.Supervivencia, False)
    End If

End With

End Sub
