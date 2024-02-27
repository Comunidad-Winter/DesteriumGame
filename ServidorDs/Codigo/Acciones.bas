Attribute VB_Name = "Acciones"
Option Explicit

''
' Modulo para manejar las acciones (doble click) de los carteles, foro, puerta, ramitas
'

''
' Ejecuta la accion del doble click
'
' @param UserIndex UserIndex
' @param Map Numero de mapa
' @param X X
' @param Y Y

Sub Accion(ByVal userIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

          Dim tempIndex As Integer
          
          '¿Rango Visión? (ToxicWaste)
   On Error GoTo Accion_Error

20        If (Abs(UserList(userIndex).Pos.Y - Y) > RANGO_VISION_Y) Or (Abs(UserList(userIndex).Pos.X - X) > RANGO_VISION_X) Then
30            Exit Sub
40        End If
          
          '¿Posicion valida?
50        If InMapBounds(map, X, Y) Then
60            With UserList(userIndex)
70                If MapData(map, X, Y).NpcIndex > 0 Then     'Acciones NPCs
80                    tempIndex = MapData(map, X, Y).NpcIndex
                      
                      'Set the target NPC
90                    .flags.TargetNPC = tempIndex
                      
100                   If Npclist(tempIndex).Comercia = 1 Then
                          '¿Esta el user muerto? Si es asi no puede comerciar
110                       If .flags.Muerto = 1 Then
                              'Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
120                           Call WriteShortMsj(userIndex, 5, FontTypeNames.FONTTYPE_INFO)
130                           Exit Sub
140                       End If
                          
                          'Is it already in commerce mode??
150                       If .flags.Comerciando Then
160                           Exit Sub
170                       End If
                          
180                       If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                              'Call WriteConsoleMsg(Userindex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
190                           Call WriteShortMsj(userIndex, 6, FontTypeNames.FONTTYPE_INFO)
200                           Exit Sub
210                       End If
                          
                          'Iniciamos la rutina pa' comerciar.
220                       Call IniciarComercioNPC(userIndex)
                      
230                   ElseIf Npclist(tempIndex).NPCtype = eNPCType.NpcCanjes Then
                          '¿Esta el user muerto? Si es asi no puede comerciar
240                       If .flags.Muerto = 1 Then
                              'Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
250                           Call WriteShortMsj(userIndex, 5, FontTypeNames.FONTTYPE_INFO)
260                           Exit Sub
270                       End If
                          
280                       If .flags.Comerciando Then
290                           Exit Sub
300                       End If
                          
310                       If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                              'Call WriteConsoleMsg(Userindex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
320                           Call WriteShortMsj(userIndex, 6, FontTypeNames.FONTTYPE_INFO)
330                           Exit Sub
340                       End If
                          
350                       Call WriteCanjeInit(userIndex, Npclist(tempIndex).Numero)
                          
360                   ElseIf Npclist(tempIndex).NPCtype = eNPCType.Banquero Then
                          '¿Esta el user muerto? Si es asi no puede comerciar
370                       If .flags.Muerto = 1 Then
                              'Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
380                           Call WriteShortMsj(userIndex, 5, FontTypeNames.FONTTYPE_INFO)
390                           Exit Sub
400                       End If
                          
                          'Is it already in commerce mode??
410                       If .flags.Comerciando Then
420                           Exit Sub
430                       End If
                          
440                       If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                              'Call WriteConsoleMsg(Userindex, "Estás demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
450                           Call WriteShortMsj(userIndex, 6, FontTypeNames.FONTTYPE_INFO)
460                           Exit Sub
470                       End If
                          
                          'A depositar de una
480                       Call IniciarDeposito(userIndex)
490                   ElseIf Npclist(tempIndex).NPCtype = eNPCType.Pirata Then
500                       If .flags.Muerto = 1 Then
                              'Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
510                           Call WriteShortMsj(userIndex, 5, FontTypeNames.FONTTYPE_INFO)
520                           Exit Sub
530                       End If
                         
                          'Is it already in commerce mode??
540                       If .flags.Comerciando Then
550                           Exit Sub
560                       End If
                     
570                       If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                              'Call WriteConsoleMsg(Userindex, "Estás demasiado lejos del pirata.", FontTypeNames.FONTTYPE_INFO)
580                           Call WriteShortMsj(userIndex, 6, FontTypeNames.FONTTYPE_INFO)
590                           Exit Sub
600                       End If
                         
610                   ElseIf Npclist(tempIndex).NPCtype = eNPCType.PirataViajes Then
620                       If .flags.Muerto = 1 Then
                              'Call WriteShortMsj(Userindex, 5, FontTypeNames.FONTTYPE_INFO)
630                           Call WriteShortMsj(userIndex, 5, FontTypeNames.FONTTYPE_INFO)
640                           Exit Sub
650                       End If
                         
                          'Is it already in commerce mode??
660                       If .flags.Comerciando Then
670                           Exit Sub
680                       End If
                     
690                       If Distancia(Npclist(tempIndex).Pos, .Pos) > 3 Then
                              'Call WriteConsoleMsg(Userindex, "Estás demasiado lejos del pirata.", FontTypeNames.FONTTYPE_INFO)
700                           Call WriteShortMsj(userIndex, 6, FontTypeNames.FONTTYPE_INFO)
710                           Exit Sub
720                       End If
                         
730                       Call WriteFormViajes(userIndex) 'Iniciamos el formulario
                         
                          'Call WriteFormViajes(UserIndex) 'Iniciamos el formulario
740                   ElseIf Npclist(tempIndex).NPCtype = eNPCType.Revividor Or Npclist(tempIndex).NPCtype = eNPCType.ResucitadorNewbie Then
750                       If Distancia(.Pos, Npclist(tempIndex).Pos) > 10 Then
                              'Call WriteConsoleMsg(Userindex, "El sacerdote no puede curarte debido a que estás demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
760                           Call WriteShortMsj(userIndex, 7, FontTypeNames.FONTTYPE_INFO)
770                           Exit Sub
780                       End If
                          
                          'Revivimos si es necesario
790                       If .flags.Muerto = 1 And (Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(userIndex)) Then
800                           Call RevivirUsuario(userIndex)
810                       End If
                          
820                       If Npclist(tempIndex).NPCtype = eNPCType.Revividor Or EsNewbie(userIndex) Then
                              'curamos totalmente
830                           .Stats.MinHp = .Stats.MaxHp
840                           Call WriteUpdateUserStats(userIndex)
850                       End If
860                   End If
                      
                  '¿Es un obj?
870               ElseIf MapData(map, X, Y).ObjInfo.ObjIndex > 0 Then
880                   tempIndex = MapData(map, X, Y).ObjInfo.ObjIndex
                      
890                   .flags.TargetObj = tempIndex
                      
900                   Select Case ObjData(tempIndex).ObjType
                          Case eOBJType.otPuertas 'Es una puerta
910                           Call AccionParaPuerta(map, X, Y, userIndex)
920                       Case eOBJType.otCarteles 'Es un cartel
930                           Call AccionParaCartel(map, X, Y, userIndex)
940                       Case eOBJType.otForos 'Foro
950                           Call AccionParaForo(map, X, Y, userIndex)
960                       Case eOBJType.otLeña    'Leña
970                           If tempIndex = FOGATA_APAG And .flags.Muerto = 0 Then
980                               Call AccionParaRamita(map, X, Y, userIndex)
990                           End If
1000                  End Select
                  '>>>>>>>>>>>OBJETOS QUE OCUPAM MAS DE UN TILE<<<<<<<<<<<<<
1010              ElseIf MapData(map, X + 1, Y).ObjInfo.ObjIndex > 0 Then
1020                  tempIndex = MapData(map, X + 1, Y).ObjInfo.ObjIndex
1030                  .flags.TargetObj = tempIndex
                      
1040                  Select Case ObjData(tempIndex).ObjType
                          
                          Case eOBJType.otPuertas 'Es una puerta
1050                          Call AccionParaPuerta(map, X + 1, Y, userIndex)
                          
1060                  End Select
                  
1070              ElseIf MapData(map, X + 1, Y + 1).ObjInfo.ObjIndex > 0 Then
1080                  tempIndex = MapData(map, X + 1, Y + 1).ObjInfo.ObjIndex
1090                  .flags.TargetObj = tempIndex
              
1100                  Select Case ObjData(tempIndex).ObjType
                          Case eOBJType.otPuertas 'Es una puerta
1110                          Call AccionParaPuerta(map, X + 1, Y + 1, userIndex)
1120                  End Select
                  
1130              ElseIf MapData(map, X, Y + 1).ObjInfo.ObjIndex > 0 Then
1140                  tempIndex = MapData(map, X, Y + 1).ObjInfo.ObjIndex
1150                  .flags.TargetObj = tempIndex
                      
1160                  Select Case ObjData(tempIndex).ObjType
                          Case eOBJType.otPuertas 'Es una puerta
1170                          Call AccionParaPuerta(map, X, Y + 1, userIndex)
1180                  End Select
1190              End If
1200          End With
1210      End If

   On Error GoTo 0
   Exit Sub

Accion_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure Accion of Módulo Acciones in line " & Erl
End Sub

Public Sub AccionParaForo(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal userIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: 02/01/2010
      '02/01/2010: ZaMa - Agrego foros faccionarios
      '***************************************************


          Dim Pos As WorldPos
          
   On Error GoTo AccionParaForo_Error

20        Pos.map = map
30        Pos.X = X
40        Pos.Y = Y
          
50        If Distancia(Pos, UserList(userIndex).Pos) > 2 Then
              'Call WriteConsoleMsg(Userindex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
60            Call WriteShortMsj(userIndex, 8, FontTypeNames.FONTTYPE_INFO)
70            Exit Sub
80        End If
          
90        If SendPosts(userIndex, ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).ForoID) Then
100           Call WriteShowForumForm(userIndex)
110       End If

   On Error GoTo 0
   Exit Sub

AccionParaForo_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure AccionParaForo of Módulo Acciones in line " & Erl
          
End Sub

Sub AccionParaPuerta(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal userIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************


   On Error GoTo AccionParaPuerta_Error

20    If Not (Distance(UserList(userIndex).Pos.X, UserList(userIndex).Pos.Y, X, Y) > 2) Then
30        If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
40            If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).Cerrada = 1 Then
                      'Abre la puerta
50                    If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).Llave = 0 Then
                          
60                        MapData(map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).IndexAbierta
                          
70                        Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).GrhIndex, X, Y))
                          
                          'Desbloquea
80                        MapData(map, X, Y).Blocked = 0
90                        MapData(map, X - 1, Y).Blocked = 0
                          
                          'Bloquea todos los mapas
100                       Call Bloquear(True, map, X, Y, 0)
110                       Call Bloquear(True, map, X - 1, Y, 0)
                          
                            
                          'Sonido
120                       Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
                          
130                   Else
                           'Call WriteConsoleMsg(Userindex, "La puerta esta cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
140                        Call WriteShortMsj(userIndex, 9, FontTypeNames.FONTTYPE_INFO)
150                   End If
160           Else
                      'Cierra puerta
170                   MapData(map, X, Y).ObjInfo.ObjIndex = ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).IndexCerrada
                      
180                   Call modSendData.SendToAreaByPos(map, X, Y, PrepareMessageObjectCreate(ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).GrhIndex, X, Y))
                                      
190                   MapData(map, X, Y).Blocked = 1
200                   MapData(map, X - 1, Y).Blocked = 1
                      
                      
210                   Call Bloquear(True, map, X - 1, Y, 1)
220                   Call Bloquear(True, map, X, Y, 1)
                      
230                   Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_PUERTA, X, Y))
240           End If
              
250           UserList(userIndex).flags.TargetObj = MapData(map, X, Y).ObjInfo.ObjIndex
260       Else
              'Call WriteConsoleMsg(Userindex, "La puerta está cerrada con llave.", FontTypeNames.FONTTYPE_INFO)
270           Call WriteShortMsj(userIndex, 9, FontTypeNames.FONTTYPE_INFO)
280       End If
290   Else
          'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
300       Call WriteShortMsj(userIndex, 8, FontTypeNames.FONTTYPE_INFO)
310   End If

   On Error GoTo 0
   Exit Sub

AccionParaPuerta_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure AccionParaPuerta of Módulo Acciones in line " & Erl

End Sub

Sub AccionParaCartel(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal userIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************

   On Error GoTo AccionParaCartel_Error

20    If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).ObjType = 8 Then
        
30      If Len(ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).Texto) > 0 Then
40        Call WriteShowSignal(userIndex, MapData(map, X, Y).ObjInfo.ObjIndex)
50      End If
        
60    End If

   On Error GoTo 0
   Exit Sub

AccionParaCartel_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure AccionParaCartel of Módulo Acciones in line " & Erl

End Sub

Sub AccionParaRamita(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal userIndex As Integer)
      '***************************************************
      'Author: Unknown
      'Last Modification: -
      '
      '***************************************************


      Dim Suerte As Byte
      Dim exito As Byte
      Dim Obj As Obj

      Dim Pos As WorldPos
   On Error GoTo AccionParaRamita_Error

20    Pos.map = map
30    Pos.X = X
40    Pos.Y = Y

50    With UserList(userIndex)
60        If Distancia(Pos, .Pos) > 2 Then
              'Call WriteShortMsj(Userindex, 8, FontTypeNames.FONTTYPE_INFO)
70            Call WriteShortMsj(userIndex, 8, FontTypeNames.FONTTYPE_INFO)
80            Exit Sub
90        End If
          
100       If MapData(map, X, Y).trigger = eTrigger.ZONASEGURA Or MapInfo(map).Pk = False Then
              'Call WriteConsoleMsg(Userindex, "No puedes hacer fogatas en zona segura.", FontTypeNames.FONTTYPE_INFO)
110           Call WriteShortMsj(userIndex, 10, FontTypeNames.FONTTYPE_INFO)
120           Exit Sub
130       End If
          
140       If .Stats.UserSkills(Supervivencia) > 1 And .Stats.UserSkills(Supervivencia) < 6 Then
150           Suerte = 3
160       ElseIf .Stats.UserSkills(Supervivencia) >= 6 And .Stats.UserSkills(Supervivencia) <= 10 Then
170           Suerte = 2
180       ElseIf .Stats.UserSkills(Supervivencia) >= 10 And .Stats.UserSkills(Supervivencia) Then
190           Suerte = 1
200       End If
          
210       exito = RandomNumber(1, Suerte)
          
220       If exito = 1 Then
230           If MapInfo(.Pos.map).Zona <> Ciudad Then
240               Obj.ObjIndex = FOGATA
250               Obj.Amount = 1
                  
                  'Call WriteConsoleMsg(Userindex, "Has prendido la fogata.", FontTypeNames.FONTTYPE_INFO)
260               Call WriteShortMsj(userIndex, 11, FontTypeNames.FONTTYPE_INFO)
                  
270               Call MakeObj(Obj, map, X, Y)
                  
                  'Las fogatas prendidas se deben eliminar
                  Dim Fogatita As New cGarbage
280               Fogatita.map = map
290               Fogatita.X = X
300               Fogatita.Y = Y
310               Call TrashCollector.Add(Fogatita)
                  
320               Call SubirSkill(userIndex, eSkill.Supervivencia, True)
330           Else
                  'Call WriteConsoleMsg(Userindex, "La ley impide realizar fogatas en las ciudades.", FontTypeNames.FONTTYPE_INFO)
340               Call WriteShortMsj(userIndex, 12, FontTypeNames.FONTTYPE_INFO)
350               Exit Sub
360           End If
370       Else
              'Call WriteConsoleMsg(Userindex, "No has podido hacer fuego.", FontTypeNames.FONTTYPE_INFO)
380           Call WriteShortMsj(userIndex, 13, FontTypeNames.FONTTYPE_INFO)
390           Call SubirSkill(userIndex, eSkill.Supervivencia, False)
400       End If

410   End With

   On Error GoTo 0
   Exit Sub

AccionParaRamita_Error:

    LogError "Error " & Err.Number & " (" & Err.Description & ") in procedure AccionParaRamita of Módulo Acciones in line " & Erl

End Sub
